#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fill the 物理学科 seminar poster (template 2, 上海大学理学院物理系) with one event.

The template is the department's own PowerPoint (20 x 30 cm, one slide).  Shape ids
are stable inside that file, so this script edits shapes by id and leaves every
other part of the deck -- theme, header lettering, 摘要 / 报告人简介 labels, footer
band -- byte-for-byte alone.  Output stays a normal editable .pptx.

Fields (JSON, flat or under "fields"):
    seminar1_no   1                         量子科技研究院Seminar第1讲
    seminar2_no   1                         物理学科Seminar第1讲
    title         Sample Talk Title ...
    speaker       某某
    speaker_title 教授                       (optional)
    affiliation   某某大学
    date          2026-01-01                (YYYY-MM-DD)
    start_time    10:00                     (HH:MM)
    end_time      11:30                     (optional)
    location      宝山校区 G601
    abstract      English or Chinese abstract
    bio           报告人简介
    inviter       某某
    inviter_title 教授                       (optional)
    inviter_email name@example.com          (optional)
    url           https://physics.shu.edu.cn/ (optional, defaults to the template)
    photo         path to the speaker photo (optional; frame is kept otherwise)

The abstract is laid out in the template's two text boxes: box 2 fills the column
next to the photo, box 4 continues underneath it at full width.  Pass
--abs-split N with the character offset where box 4 starts (measure_text.ps1
reports the exact offset PowerPoint itself would use).
"""
import argparse
import copy
import datetime as dt
import json
import os
import re
import tempfile
import zipfile
from pathlib import Path

from xml.etree import ElementTree as etree

ROOT = Path(__file__).resolve().parents[1]

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
R = '{%s}' % NS['r']

# shape ids inside the template (ppt/slides/slide1.xml)
SID = {
    'seminar': 27,
    'title': 32,
    'meta': 18,
    'abs1': 2,
    'abs2': 4,
    'bio': 6,
    'bio_label': 17,
    'url': 44,
    'inviter': 19,
    'photo': 21,
}

SLIDE = 'ppt/slides/slide1.xml'
SLIDE_RELS = 'ppt/slides/_rels/slide1.xml.rels'


# --------------------------------------------------------------------------- helpers
def qn(tag):
    prefix, local = tag.split(':')
    return '{%s}%s' % (NS[prefix], local)


def parse(data):
    """Parse a slide part, keeping the document's own namespace prefixes intact."""
    for prefix, uri in re.findall(rb'xmlns:([A-Za-z0-9_.-]+)="([^"]+)"', data):
        etree.register_namespace(prefix.decode('utf-8'), uri.decode('utf-8'))
    return etree.fromstring(data)


def dump(root):
    """Serialise with the OOXML-style declaration PowerPoint writes itself."""
    return (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            + etree.tostring(root, encoding='utf-8'))


def drop_shape(root, sid):
    """Remove a shape by id (ElementTree keeps no parent pointers)."""
    for parent in root.iter():
        for child in list(parent):
            if child.tag in (qn('p:sp'), qn('p:pic'), qn('p:graphicFrame')):
                nv = child.find('.//' + qn('p:cNvPr'))
                if nv is not None and nv.get('id') == str(sid):
                    parent.remove(child)
                    return True
    return False


def by_id(root, sid):
    for el in root.iter():
        if el.tag in (qn('p:sp'), qn('p:pic'), qn('p:graphicFrame')):
            nv = el.find('.//' + qn('p:cNvPr'))
            if nv is not None and nv.get('id') == str(sid):
                return el
    return None


def name_of(el):
    nv = el.find('.//' + qn('p:cNvPr'))
    return nv.get('name') if nv is not None else '?'


def first_rpr(sp):
    """Deep copy of the first run's rPr, used as the style stamp for new runs."""
    r = sp.find('.//' + qn('a:r'))
    if r is not None:
        rpr = r.find(qn('a:rPr'))
        if rpr is not None:
            return copy.deepcopy(rpr)
    rpr = sp.find('.//' + qn('a:endParaRPr'))
    if rpr is not None:
        return copy.deepcopy(rpr)
    return etree.Element(qn('a:rPr'))


def first_ppr(sp):
    p = sp.find(qn('p:txBody') + '/' + qn('a:p'))
    if p is not None:
        ppr = p.find(qn('a:pPr'))
        if ppr is not None:
            return copy.deepcopy(ppr)
    return etree.Element(qn('a:pPr'))


def set_sz(rpr, size):
    """sz is in 1/100 pt; PowerPoint rejects anything below 100 (1 pt)."""
    if size is None:
        return
    value = int(round(size * 100))
    if not 100 <= value <= 400000:
        raise ValueError('font size out of range: %r' % size)
    rpr.set('sz', str(value))


def set_leading(ppr, leading):
    """Exact line spacing in points, written as a:spcPts (100 = 1 pt)."""
    if leading is None:
        return
    for old in ppr.findall(qn('a:lnSpc')):
        ppr.remove(old)
    ln = etree.Element(qn('a:lnSpc'))
    etree.SubElement(ln, qn('a:spcPts')).set('val', str(int(round(leading * 100))))
    ppr.insert(0, ln)


def write_paras(sp, texts, size=None, leading=None, align=None, indent=None, rpr=None):
    """Replace every paragraph of a shape, keeping its <a:bodyPr> and pPr styling."""
    tx = sp.find(qn('p:txBody'))
    if tx is None:
        raise SystemExit('shape %s has no text body' % name_of(sp))
    body = tx.find(qn('a:bodyPr'))
    lst = tx.find(qn('a:lstStyle'))
    ppr_tpl = first_ppr(sp)
    rpr_tpl = rpr if rpr is not None else first_rpr(sp)
    for p in tx.findall(qn('a:p')):
        tx.remove(p)
    for text in texts:
        p = etree.SubElement(tx, qn('a:p'))
        ppr = copy.deepcopy(ppr_tpl)
        if align is not None:
            ppr.set('algn', align)
        ppr.set('hangingPunct', '1')  # 。，hang at the line end instead of starting a line
        set_leading(ppr, leading)
        if indent:
            ppr.set('indent', str(int(indent)))
            ppr.set('marL', '0')
        p.append(ppr)
        rp = copy.deepcopy(rpr_tpl)
        set_sz(rp, size)
        if text:
            r = etree.SubElement(p, qn('a:r'))
            r.append(copy.deepcopy(rp))
            etree.SubElement(r, qn('a:t')).text = text
        end = copy.deepcopy(rp)
        end.tag = qn('a:endParaRPr')
        p.append(end)


def image_size(path):
    data = Path(path).read_bytes()
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        return int.from_bytes(data[16:20], 'big'), int.from_bytes(data[20:24], 'big')
    i = 2
    while i < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xC0, 0xC1, 0xC2):
            h = int.from_bytes(data[i + 5:i + 7], 'big')
            w = int.from_bytes(data[i + 7:i + 9], 'big')
            return w, h
        i += 2 + int.from_bytes(data[i + 2:i + 4], 'big')
    raise ValueError('unsupported image format: %s' % path)


def has_cjk(text):
    return any('\u3400' <= ch <= '\u9fff' for ch in text)


def tidy(text):
    """One flowing paragraph: no manual line breaks, no accidental double spaces."""
    text = re.sub(r'\s*\n\s*', ' ', text)
    text = re.sub(r'  +', ' ', text).strip()
    # Chinese typography: no space before a full-width bracket, and a bracket that
    # opened full-width must close full-width too.
    text = re.sub(r'(?<=[\u4e00-\u9fff0-9])\s+（', '（', text)
    text = re.sub(r'（([A-Za-z0-9 .&/+–—-]*)\)', r'（\1）', text)
    return text


def set_box_height(sp, lines, leading, pad=7.2):
    """Grow/shrink a shape-to-fit-text box so its frame matches the measured text."""
    if not lines or not leading:
        return
    ext = sp.find('.//' + qn('a:xfrm') + '/' + qn('a:ext'))
    if ext is None:
        return
    ext.set('cy', str(int(round((lines * leading + pad) * 12700))))


def is_all_ascii(text):
    return all(ord(ch) < 128 for ch in text)


# --------------------------------------------------------------------------- data
REQUIRED = ('title', 'speaker', 'affiliation', 'date', 'start_time', 'location',
            'abstract', 'bio')


def prepare(raw):
    data = raw.get('fields', raw) if isinstance(raw.get('fields'), dict) else dict(raw)
    data = {k: (v.strip() if isinstance(v, str) else v) for k, v in data.items()}
    missing = [k for k in REQUIRED if not data.get(k)]
    if missing:
        raise SystemExit('missing required fields: ' + ', '.join(missing))
    if not (data.get('seminar1_no') and data.get('seminar2_no')) and not data.get('seminar'):
        raise SystemExit('need seminar1_no + seminar2_no (or a literal seminar)')

    when = dt.date.fromisoformat(data['date'])
    weekday = '一二三四五六日'[when.weekday()]
    clock = data['start_time'] + ('~' + data['end_time'] if data.get('end_time') else '')
    who = data['speaker'] + (' ' + data['speaker_title'] if data.get('speaker_title') else '')

    seminar = data.get('seminar') or (
        '量子科技研究院Seminar第%s讲/物理学科Seminar第%s讲'
        % (data['seminar1_no'], data['seminar2_no']))
    meta = ['报告人：%s（%s）' % (who, data['affiliation']),
            '时    间：%d年%d月%d日（周%s）%s' % (when.year, when.month, when.day, weekday, clock),
            '地    点：%s' % data['location']]
    inviter = ''
    if data.get('inviter'):
        inviter = '邀请人：' + data['inviter']
        if data.get('inviter_title'):
            inviter += ' ' + data['inviter_title']
        if data.get('inviter_email'):
            inviter += ' ' + data['inviter_email']
    fields = dict(seminar=seminar, title=data['title'], meta=meta,
                  abstract=tidy(data['abstract']), bio=tidy(data['bio']),
                  url=data.get('url', ''), inviter=inviter)
    return data, fields


# --------------------------------------------------------------------------- build
def shift(el, dy):
    off = el.find('.//' + qn('a:xfrm') + '/' + qn('a:off'))
    if off is None:
        return
    off.set('y', str(int(off.get('y')) + int(round(dy))))


def put_photo(slide, rels, payload, pic_id, path, contain=True):
    pic = by_id(slide, pic_id)
    if pic is None:
        raise SystemExit('photo frame (shape %d) missing from template' % pic_id)
    blip = pic.find('.//' + qn('a:blip'))
    rid = blip.get(R + 'embed')
    target = None
    for rel in rels:
        if rel.get('Id') == rid:
            target = rel.get('Target')
            break
    if target is None:
        raise SystemExit('photo frame has no image relationship')
    part = os.path.normpath(os.path.join('ppt/slides', target)).replace('\\', '/')
    payload[part] = Path(path).read_bytes()
    if contain:
        xfrm = pic.find('.//' + qn('a:xfrm'))
        off = xfrm.find(qn('a:off'))
        ext = xfrm.find(qn('a:ext'))
        fx, fy = int(off.get('x')), int(off.get('y'))
        fw, fh = int(ext.get('cx')), int(ext.get('cy'))
        iw, ih = image_size(path)
        scale = min(fw / iw, fh / ih)
        w, h = max(1, round(iw * scale)), max(1, round(ih * scale))
        off.set('x', str(fx + (fw - w) // 2))
        off.set('y', str(fy + (fh - h) // 2))
        ext.set('cx', str(w))
        ext.set('cy', str(h))
    return part


def build(args):
    template = args.template.resolve()
    out = args.out.resolve()
    if out.exists() and not args.force:
        raise SystemExit('%s exists; pass --force' % out)
    data, fields = prepare(json.loads(args.data.read_text(encoding='utf-8-sig')))

    with zipfile.ZipFile(template) as z:
        items = z.infolist()
        payload = {i.filename: z.read(i.filename) for i in items}

    slide = parse(payload[SLIDE])
    rels_root = parse(payload[SLIDE_RELS])

    abstract = fields['abstract']
    split = args.abs_split or 0
    if not 0 <= split <= len(abstract):
        raise SystemExit('--abs-split %d outside the abstract (0..%d)' % (split, len(abstract)))
    abs1_text = abstract[:split].rstrip() if split else abstract
    abs2_text = abstract[split:].lstrip() if split else ''

    for key in ('seminar', 'title', 'meta', 'abs1', 'abs2', 'bio', 'inviter'):
        if by_id(slide, SID[key]) is None:
            raise SystemExit('template is missing shape id %d (%s)' % (SID[key], key))

    write_paras(by_id(slide, SID['seminar']), [fields['seminar']])
    write_paras(by_id(slide, SID['title']), [fields['title']])
    write_paras(by_id(slide, SID['meta']), fields['meta'])

    write_paras(by_id(slide, SID['abs1']), [abs1_text],
                size=args.abs_size, leading=args.abs_leading, align='just')
    write_paras(by_id(slide, SID['abs2']), [abs2_text] if abs2_text else [''],
                size=args.abs_size, leading=args.abs_leading, align='just')

    bio_indent = args.bio_indent_chars * args.bio_size * 12700 if (
        args.bio_indent_chars and has_cjk(fields['bio'])) else 0
    write_paras(by_id(slide, SID['bio']), [fields['bio']],
                size=args.bio_size, leading=args.bio_leading, align='just',
                indent=bio_indent)

    if args.bio_shift:
        dy = args.bio_shift * 914400
        for key in ('bio_label', 'bio'):
            shift(by_id(slide, SID[key]), dy)

    if args.abs2_lines:
        set_box_height(by_id(slide, SID['abs2']), args.abs2_lines, args.abs_leading)
    if args.bio_lines:
        set_box_height(by_id(slide, SID['bio']), args.bio_lines, args.bio_leading)

    if fields['url']:
        write_paras(by_id(slide, SID['url']), [fields['url']])
    if fields['inviter']:
        write_paras(by_id(slide, SID['inviter']), [fields['inviter']])

    photo = data.get('photo')
    if photo:
        if not Path(photo).exists():
            raise SystemExit('photo not found: %s' % photo)
        put_photo(slide, rels_root, payload, SID['photo'], photo)
    else:
        drop_shape(slide, SID['photo'])

    payload[SLIDE] = dump(slide)

    out.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix='.pptx', dir=out.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
            for item in items:
                z.writestr(item, payload[item.filename])
        os.replace(tmp, out)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    print('wrote %s  (abs1 %d chars / abs2 %d chars)' % (out, len(abs1_text), len(abs2_text)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True, type=Path)
    ap.add_argument('--out', required=True, type=Path)
    ap.add_argument('--template', type=Path,
                    default=ROOT / 'assets' / 'poster-template-physics.pptx')
    ap.add_argument('--force', action='store_true')
    ap.add_argument('--abs-size', type=float, default=16)
    ap.add_argument('--abs-leading', type=float, default=20)
    ap.add_argument('--abs-split', type=int, default=0)
    ap.add_argument('--bio-size', type=float, default=16)
    ap.add_argument('--bio-leading', type=float, default=22)
    ap.add_argument('--bio-indent-chars', type=float, default=2)
    ap.add_argument('--bio-shift', type=float, default=0,
                    help='move 报告人简介 label+text by N inches (negative = up)')
    ap.add_argument('--abs2-lines', type=int, default=0,
                    help='set the lower abstract box height to this many lines')
    ap.add_argument('--bio-lines', type=int, default=0,
                    help='set the bio box height to this many lines')
    build(ap.parse_args())


if __name__ == '__main__':
    main()

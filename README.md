# shu-physics-poster

上海大学理学院物理系「学术报告」海报生成技能（DSH skill）：把一份学术报告的 Word 文档，做成物理系实际在用的 2 号模板海报，输出**可编辑 PPTX + PNG + PDF**。

母版就是系里自己的 PPT 原件（20×30 cm，单页）。字形、字号、颜色、标签、页脚全部保持原样，脚本只按 shape id 换文字和照片，所以产出的文件能继续在 PowerPoint 里改。

## 它解决什么

手工做一张这种海报，真正麻烦的不是排版，而是这几件事：

- **摘要要绕开右上角的照片。** 母版用了两个文本框接力（照片左侧 5.95 in、照片下方 7.33 in），PowerPoint 不会自动串接，断点得算。脚本用 PowerPoint 自己量出断行位置，再决定从第几个字符换框。
- **字数一变就得重调字号。** 摘要 300 字和 1300 字不能都用 16 pt。脚本给出每个框的行容量，先估一版，再实测校准。
- **照片不能变形。** 照片按原图宽高比等比缩放、居中，不裁切、不拉伸；没有照片就直接删掉图片框，不留灰框。
- **字体不能串。** 正文英文 Times New Roman、中文宋体；抬头是华文彩云、标签是华文行楷、信息区是微软雅黑——这些是系里其它海报的一致性要求。

## 环境要求

| 依赖 | 说明 |
|---|---|
| Python 3.8+ | `scripts/build_poster.py` **只用标准库**，不需要 lxml、也不需要 python-pptx |
| Windows + PowerPoint | `measure_text.ps1` / `render_poster.ps1` 通过 COM 实测断行、导出 PNG/PDF |
| Windows PowerShell 5.1 | 两个 `.ps1` 保持纯 ASCII，可直接跑 |

没有 PowerPoint 也能生成 pptx，只是拿不到 PNG/PDF，也无法实测断行（只能靠估算字号）。

## 用法

### 1. 写字段文件

```json
{
  "fields": {
    "seminar1_no": "1",
    "seminar2_no": "1",
    "title": "Sample Talk Title",
    "speaker": "某某",
    "speaker_title": "教授",
    "affiliation": "某某大学",
    "date": "2026-01-01",
    "start_time": "10:00",
    "end_time": "11:30",
    "location": "宝山校区 G601",
    "abstract": "...",
    "bio": "...",
    "inviter": "某某",
    "inviter_title": "教授",
    "inviter_email": "name@example.com",
    "photo": "D:\\path\\to\\photo.png"
  }
}
```

完整字段见 [`references/fields.md`](references/fields.md)。

### 2. 先建探针，量 PowerPoint 自己的断行

```powershell
python -X utf8 scripts\build_poster.py --data fields.json --out probe.pptx --force --abs-size 13.5 --abs-leading 17
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\measure_text.ps1 -InputPptx probe.pptx -OutJson probe.json -ShapeIds "2,4,6"
```

`probe.json` 里每行都带 `start`（1 起算）。要在 2 号框放 8 行，就取 `lines[8].start - 1` 作为断点。

### 3. 正式生成

```powershell
python -X utf8 scripts\build_poster.py --data fields.json --out poster.pptx --force `
    --abs-size 13.5 --abs-leading 17 --abs-split 595 `
    --bio-size 13 --bio-leading 17 --abs2-lines 8 --bio-lines 8
```

### 4. 渲染并核对

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\render_poster.ps1 -InputPptx poster.pptx -OutputDir render
```

看 `render\poster.png`，再看 `render\qa.json`：2/4/6 号框的 `overflow_pt.bottom` 必须是 0。

## 目录结构

```
├── SKILL.md                     技能正文：版式要点、shape id 对照、工作流、判定标准
├── references/fields.md          字段表、命令行参数、三张成品的实际参数
├── scripts/build_poster.py       填充母版（纯标准库）
├── scripts/measure_text.ps1      用 PowerPoint 实测断行
├── scripts/render_poster.ps1     导出 PNG/PDF + 文字边界 qa.json
└── assets/poster-template-physics.pptx   物理系在用的母版原件
```

## 母版里的 shape id

| id | 内容 | id | 内容 |
|---|---|---|---|
| 24 | 中文系名 + 英文系名 | 3 / 17 | 摘要 / 简介标签 |
| 26 | 学术报告 | 2 | 摘要上半（照片左侧） |
| 27 | 红色 Seminar 讲次行 | 4 | 摘要下半（照片下方） |
| 32 | 题目 | 6 | 报告人简介 |
| 18 | 报告人 / 时间 / 地点 | 21 | 报告人照片 |
| 44 / 19 | 页脚链接 / 邀请人 | 10 / 16 / 20 | 色块与校徽 |

## 已知限制

- **母版已脱敏**：正文、照片、邀请人邮箱都是占位内容（灰色剪影 + `name@example.com`），只有版式、字体、配色是系里真实在用的。所以这个仓库可以公开，不会带出别人的照片和邮箱。
- 摘要断点依赖 PowerPoint 实测；换机器、换字体版本后断行可能相差一行，建议重跑一遍第 2 步。
- 中文字体（华文彩云、华文行楷、微软雅黑、宋体）需本机已安装，否则会走字体替换，视觉上不一致。

## 许可

MIT，见 [LICENSE](LICENSE)。

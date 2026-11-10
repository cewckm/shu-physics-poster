---
name: shu-physics-poster
description: 把学术报告 Word 文档做成上海大学理学院物理系版式（2 号模板）的可编辑 PPT 海报，含院系抬头、Seminar 讲次、报告人照片、摘要/简介、页脚链接与邀请人。交付 PPTX + PNG + PDF。
---

# 物理系 Seminar 海报（2 号模板）

> **模板编号：2 号模板。** 用户说「按 2 号模板做 / 用第二个模板 / 物理系那种」时用这个 skill。
> 1 号模板是另一个 skill `shu-academic-poster`（量子科技研究院版式，20×32 cm，华文隶书抬头，页脚不带链接）。
> 报告挂物理学科 Seminar → 2 号；挂量子科技研究院 → 1 号。

## 两个模板的差别

| | 1 号（`shu-academic-poster`） | 2 号（本 skill） |
|---|---|---|
| 抬头 | 华文隶书「量子科技研究院」+ Calibri 英文 | 华文彩云「上海大学理学院物理系」+ Times 英文 |
| 「学术报告」 | 微软雅黑 加粗 红色 | 华文行楷 白色 |
| 画布 | 20 × 32 cm | 20 × 30 cm |
| `摘要`/`报告人简介`标签 | 方块 蓝底白字 | 圆角渐变 蓝底华文行楷 |
| 页脚 | 蓝色底条，只放邀请人 | 左链接 + 右邀请人 |
| 摘要分框 | 一个框 | 两个框（绕开右上角照片） |

`assets/poster-template-physics.pptx` 就是物理系实际在用的 PPT 原件（20×30 cm，单页）。为了能公开托管，里面的正文和照片已换成占位内容（灰底剪影 + `name@example.com`），**版式、字体、配色与原件一致**。**母版本身不要改字形、字号、颜色、位置**，`build_poster.py` 只按 shape id 换字、换照片。

## 版式要点（不要改）

- 抬头：白色校徽图片 + 「上海大学理学院物理系」**华文彩云** 20 pt；英文系名 `Department of Physics, Shanghai University` Times New Roman 16 pt；右上「学术报告」**华文行楷** 32 pt。
- 红色讲次行：**微软雅黑** 21.5 pt 加粗 `#C00000`，英文/数字走 Times New Roman。
- 题目：微软雅黑加粗（字号继承母版，约 24 pt），居中，尽量两行以内。
- 报告人 / 时间 / 地点：微软雅黑 17 pt 加粗，1.5 倍行距，左对齐在 x=1.05 in。
- 「摘 要」「报告人简介」：蓝色渐变圆角标签 + 华文行楷白字（20 pt）。
- 正文：英文 **Times New Roman**、中文 **宋体**；摘要两端对齐、无首行缩进；简介首行缩进 2 字、行距略松。
- 摘要天生分两个框：**2 号**（照片左侧，5.95 in 宽）接着 **4 号**（照片下方，7.33 in 宽），这样文字能绕开右上角的照片。PowerPoint 文本框不会自动串接，所以断点必须算出来（见工作流）。
- 照片：**21 号**图片框 1.26×1.77 in，等比缩放后居中，**不裁切、不拉伸**；没有照片就整个删掉，不留灰框。
- 页脚蓝条：左 `https://physics.shu.edu.cn/`，右「邀请人：姓名 职称 邮箱」。

## shape id 对照

| id | 内容 | 字体 |
|---|---|---|
| 10 / 16 | 页眉、页脚色块 | — |
| 20 | 校徽图片 | — |
| 24 | 中文系名 + 英文系名 | 华文彩云 20 / Times 16 |
| 26 | 学术报告 | 华文行楷 32 |
| 27 | 红色 Seminar 讲次行 | 微软雅黑 21.5 加粗 |
| 32 | 题目 | 微软雅黑（继承字号） |
| 18 | 报告人 / 时间 / 地点 | 微软雅黑 17 加粗 |
| 3 / 17 | 摘要 / 简介标签 | 华文行楷 20 |
| 2 | 摘要上半（照片左侧） | Times + 宋体 |
| 4 | 摘要下半（照片下方） | Times + 宋体 |
| 6 | 报告人简介 | Times + 宋体 |
| 21 | 报告人照片 | — |
| 44 / 19 | 页脚链接 / 邀请人 | Times 15 / 微软雅黑 15 |

## 依赖

- `build_poster.py`：**只用 Python 标准库**（3.8+ 即可，不需要 lxml、也不需要 python-pptx），PATH 上的 `python` 直接能跑。
- `render_poster.ps1` / `measure_text.ps1`：需要本机装有 PowerPoint（走 COM），Windows PowerShell 5.1 即可。
- 两个 .ps1 都是纯 ASCII，不要往里写中文，否则 5.1 会按 ANSI 解码报解析错误。

## 工作流

1. 读 DOCX 取字段（题目、报告人、单位、日期时间、地点、摘要、简介、邀请人、照片），写 `fields.json`（字段见 [references/fields.md](references/fields.md)）。
2. **先建探针**，摘要全部塞进 2 号框（字号先按下面「行容量」估，下面命令里的 13.5 / 17 只是举例）：

   ```powershell
   python -X utf8 "<skill>\scripts\build_poster.py" --data fields.json --out probe.pptx --force --abs-size 13.5 --abs-leading 17
   ```

3. **量 PowerPoint 自己的断行**，拿到第 N+1 行的起始字符位：

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<skill>\scripts\measure_text.ps1" -InputPptx probe.pptx -OutJson probe.json -ShapeIds "2,4,6"
   ```

   JSON 里 `shapes[].lines[]` 每行都带 `start`（1 起算）。要 2 号框放 8 行，就取 `lines[8].start - 1`（即第 9 行的起点）作为 `--abs-split`。

4. **正式生成**（把量出来的断点喂回去，并用 `--abs2-lines` / `--bio-lines` 让框高贴合行数）：

   ```powershell
   python -X utf8 "<skill>\scripts\build_poster.py" --data fields.json --out 海报.pptx --force `
       --abs-size 13.5 --abs-leading 17 --abs-split 595 `
       --bio-size 13 --bio-leading 17 --abs2-lines 8 --bio-lines 8
   ```

5. **渲染并核对**（需要 PowerPoint）：

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<skill>\scripts\render_poster.ps1" -InputPptx 海报.pptx -OutputDir render
   ```

   看 `poster.png`，再看 `qa.json`：2/4/6 号框的 `overflow_pt.bottom` 要都是 0（判定标准见下）。

6. 打开 PNG 检查：题目是否折行、摘要是否压到「报告人简介」标签、简介是否压到页脚、照片是否变形、抬头字形是否还是华文彩云。

## 行容量（先按这个估字号，再用第 3 步校准）

| 框 | 可用高度 | 可放行数 |
|---|---|---|
| 2 号（照片左侧） | 144 pt | `floor(144 / 行距)` |
| 4 号（照片下方，到简介标签上沿） | 145 pt | `floor(145 / 行距)` |
| 简介（到页脚上沿） | 181 pt | `floor(181 / 行距)` |

先按 `摘要总行数 ÷ 容量` 估一版字号（容量和行数都随字号反向变化，实际约按平方根收敛），再靠 `measure_text.ps1` 精确对。摘要短、下方空一大块时，用 `--bio-shift` 把「报告人简介」整块上移（负值，单位英寸），或把简介字号提到 18 补白。

## 判定标准

- 摘要最后一行在「报告人简介」标签上方，至少留 0.25 in。
- 简介最后一行在页脚蓝条上方；`qa.json` 里 2/4/6 号框的 `overflow_pt.bottom` 必须是 0。
  `right_pt` 出现 5 pt 以内的外溢是 `hangingPunct` 让标点悬挂在版心外，属正常，不用管。
- 照片宽高比 = 原图宽高比（脚本已按 contain 缩放，不要手改）。
- 抬头、红色讲次行、「摘 要」「报告人简介」标签的字形与母版一致。
- 全文只出现母版本来就有的字体组合（Times New Roman + 宋体 / 微软雅黑 / 华文彩云 / 华文行楷），不引入新字体。

## 脚本会自动做的清理

- 合并段落内的换行，折叠多余空格（`energy.   We treat` → `energy. We treat`）。
- 中文后接全角括号时去掉多余空格（`102 （Google` → `102（Google`），并把 `（AIIM)，` 这种全半角混用的括号补成全角。
- 中文简介自动加 2 字首行缩进（`--bio-indent-chars 0` 可关）。
- `fields.json` 里没有 `photo` 时，直接把照片框删掉，不留空框。

## 已知差异

- DOCX 与已发布海报在**时间/讲次**上可能不一致（例：文档写「上午 10:00–11:30」，往期海报写的是另一个时段）。**以文档为准**（用户明确要求过），交付时说明一句。
- 文档里的邀请人邮箱可能有笔误（`shu.cdu.cn`），按海报更正为 `shu.edu.cn`。
- 文档常把上标参考文献号写成普通字符（`electrocatalysis.1`），保持原样，不要自己改成上标。

# 字段

`build_poster.py --data` 接受 UTF-8 JSON，可以是扁平对象，也可以放在 `fields` 下。

| 字段 | 用途 | 备注 |
|---|---|---|
| seminar1_no / seminar2_no | 红色讲次行 | `量子科技研究院Seminar第{1}讲/物理学科Seminar第{2}讲`，数字留着可改 |
| seminar | 红色讲次行（整句） | 给了 `seminar1_no`+`seminar2_no` 就可省略 |
| title | 报告题目 | 保留原文缩写与大小写 |
| speaker / speaker_title | 报告人姓名、职称 | 职称可省略 |
| affiliation | 单位 | 显示在报告人后的全角括号里 |
| date | 日期 | `YYYY-MM-DD`，脚本自动算星期 |
| start_time / end_time | 时间 | `HH:MM`；结束时间可省略（如只写 3pm 时） |
| location | 地点 | 如 `宝山校区 G601` |
| abstract | 摘要 | 中英文均可；内部换行会被合并成一段 |
| bio | 报告人简介 | 中文简介会自动首行缩进 2 字 |
| inviter / inviter_title / inviter_email | 邀请人 | 缺省时页脚只留链接 |
| url | 页脚左链接 | 默认 `https://physics.shu.edu.cn/` |
| photo | 报告人照片路径 | PNG/JPG；按框等比缩放居中，不裁切 |

## 版面参数（命令行）

| 参数 | 默认 | 说明 |
|---|---|---|
| `--abs-size` / `--abs-leading` | 16 / 20 | 摘要字号与固定行距（行距用 `spcPts`，单位 pt） |
| `--abs-split` | 0 | 摘要断点（字符下标）：之前放 2 号框，之后放 4 号框；0 = 全放 2 号框 |
| `--bio-size` / `--bio-leading` | 16 / 22 | 简介字号与行距 |
| `--bio-shift` | 0 | 「报告人简介」标签+正文整体位移，**单位英寸**，负值向上 |
| `--abs2-lines` | 0 | 把 4 号框高度设为 N 行，让框贴合文字 |
| `--bio-lines` | 0 | 把简介框高度设为 N 行 |
| `--bio-indent-chars` | 2 | 中文简介首行缩进字数，0 关闭 |

`--abs-size` 这一类参数一律是**磅**，脚本写成 `sz=字号×100`，越界会直接报错（sz<100 会让 PowerPoint 判定文件损坏）。

## 三个跑通过的实例（参数可照抄）

| 案例 | 摘要 | 简介 |
|---|---|---|
| A：短摘要，7 行就够 | 17 pt / 20 pt，`--abs-split 0` | 18 pt / 24 pt，8 行，`--bio-shift -1.08` |
| B：长摘要，要跨两框 | 13.5 pt / 17 pt，`--abs-split 595`（2 号框 8 行 + 4 号框 8 行） | 13 pt / 17 pt，8 行，`--bio-shift 0.18` |
| C：母版原样 | 16 pt / 20 pt，7 + 6 行 | 16 pt / 22 pt，5 行 |

案例 A 摘要只占满 2 号框、4 号框空着，所以把「报告人简介」整块上移补白；案例 B 摘要太长，
必须跨两个框，反而不用动标签位置。

## 判定标准

- 摘要最后一行在「报告人简介」标签上方 ≥ 0.25 in。
- 简介最后一行在页脚蓝条上方。
- 题目一行放不下时先缩字号，不要让它长到三行。
- 正文只允许 Times New Roman + 宋体两种字体，中文块不要混进微软雅黑。

---
name: wps-docx-copyfix
description: >-
  WPS 转换 docx 复制友好化：把"WPS 把 PDF 转 Word"得到的文档修成可以放心复制粘贴的版本，
  内容零丢失。三步根因修复（浮动图转内联→删空分节符段→去定位缩进）+ 按需内容手术
  （合并折行/修顺序/线性化分数）+ 五道验收关卡。当用户说"WPS 转换的文档复制出来乱了"、
  "整理 WPS 转的 word"、"复制格式错乱"、"每行超出"、"贴过去图跑位"时使用。
user-invocable: true
metadata:
  title: WPS 转换 docx 复制友好化
  version: 1.0.0
---

# WPS 转换 docx 复制友好化（复制乱版 → 稳定可贴）

场景：用户用 WPS 把 PDF 转成 Word（或拿到别人这么转的文档），在 WPS 里打开看着正常，
**一旦复制粘贴到别处就乱**：文字带出超宽缩进、每行硬换行、图飘到文字右边、
跨页的题被拦腰拆开。本技能把这种文档修成"贴到哪里都稳定"，且正文一字不丢。

## 为什么会乱（三个根因，先懂再修）

WPS 转换的目标是**像素级复刻 PDF 版式**，为此埋了三类"定时炸弹"：

| 根因 | WPS 干了什么 | 粘贴后症状 |
|---|---|---|
| ① **浮动锚定图** `wp:anchor` | 部分图用"相对锚点段落的 X/Y 偏移 + 浮动"放置 | 锚点段落换位后图按旧坐标漂移：图压文字、堆到右侧 |
| ② **空分节符段** `pPr+sectPr` | PDF 每页生成一个空段落锁分页（为复刻每页页码） | 5~N 个强制 nextPage 换页被带进目标文档，**跨页内容被拦腰拆开** |
| ③ **定位缩进** `w:ind` | 每个视觉行一段 + 大缩进模拟 PDF 位置（可达 16cm） | 粘贴带出缩进 → "每行超出"；一行一段 → 硬换行 |

另有次级问题：折行句子被拆成多段、WPS 排序错乱（一句话的中间部分排到段尾之后）、
分数公式炸成"u2"/"R"孤行、两栏选项挤一行、成片空段。

**关键认知**：复制粘贴**带走**段落格式/浮动图属性/分节符，**不带**页面设置（纸型边距）。
所以"原文档页边距"只影响它自己的打印；乱版全部来自上面三类被带走的东西。
WPS 打开原文档没事，是因为 WPS 重新结算浮动坐标；粘贴时**不重算**，按原坐标硬放。

## 脚本

| 脚本 | 用途 |
|---|---|
| `scripts/diagnose.py` | 体检报告：浮动图明细/空分节符段/缩进分布/空段runs + 全段落 dump；`--pdf 原版.pdf` 附提取阅读顺序 |
| `scripts/auto_fix.py` | 机械修复四件套（幂等）：anchor→inline、删空分节符段、去 w:ind、折叠空段；`--center-title` 标题居中；`--a4-2cm`/`--margins` 可选页面设置 |
| `scripts/verify.py` | 五道验收关卡：XML良构/图片数/正文零丢失(忽略空白字符级diff)/zip回读/可选渲染页数 |
| `scripts/docx_lib.py` | 工具库：解包重组、段落与 run 级操作原语（内容手术用） |

依赖：python3 标准库 + `xmllint`；验收渲染需 `soffice`、`pdftotext/pdfinfo`（poppler）。

## 工作流

### Phase 0 — 备份与工作目录
- **原文档绝不改动**（尤其名字带"备份"的），产新文件：`原名(整理版).docx`
- workdir 用 `/tmp/wpsfix_<名词>/`，避免污染用户目录

### Phase 1 — 诊断
```bash
python3 scripts/diagnose.py 输入.docx --pdf 原版.pdf -o /tmp/wpsfix_x/diag
```
- 读报告确认三个根因各自的数量（浮动图 N 张 / 空分节符段 M 个 / 缩进 K 处）
- **对照 `paras_dump.txt` 与 `pdf_ground_truth.txt`**：找出被拆散的句子、顺序错乱处、
  公式孤行——这些是 Phase 3 内容手术的清单

### Phase 2 — 机械修复（安全通用，直接跑）
```bash
python3 scripts/auto_fix.py 输入.docx -o 整理版.docx --center-title   # 页面设置按用户要求加 --a4-2cm
```
顺序内置于脚本：anchor→inline → 删空分节符段 → 去缩进 → 折叠空段。
**注意**：删分节符后页数可能不变（内容自然撑满）也可能变少，都要向用户说明这是解锁分页的正常结果。

### Phase 3 — 内容手术（按 Phase 1 清单逐处修，需要判断力）
用 `docx_lib` 原语做 run 级操作。四种典型术式：

```python
import sys; sys.path.insert(0, 'scripts')
from docx_lib import runs_of, with_runs, bare_p, para_text, rebuild, para_spans

xml = open(doc_xml, encoding='utf-8').read()
paras = [xml[s:e] for s, e in para_spans(xml)]

# 术式A 合并折行段: 把段j的runs接到段i末尾
paras[i] = with_runs(paras[i], runs_of(paras[i]) + runs_of(paras[j])); paras[j] = None

# 术式B 顺序修复(WPS乱序): 段i拆head/tail，中间插入乱段j
r = runs_of(paras[i]); cut = r.find('<w:r>...流表那run的精确XML...</w:r>')
paras[i] = with_runs(paras[i], r[:cut] + runs_of(paras[j]) + r[cut:]); paras[j] = None

# 术式C 公式孤行线性化: 删碎片段, 在公式段空位run处replace成线性文本(如 u2/R)
# 术式D 两栏选项拆行: 以B.标记run为界拆两段, 拆出的新段用 bare_p(原段, 后半runs) 构造

xml = rebuild(xml, paras)  # None 被删除
open(doc_xml, 'w', encoding='utf-8').write(xml)
```
- 断言先行：每处手术 `assert 精确XML in 段`, 匹配不到就停，绝不模糊替换
- 每次手术后立即重新 dump，核对阅读顺序与 PDF ground truth 一致

### Phase 4 — 验收（无证据不交货）
```bash
python3 scripts/verify.py 原始.docx 整理版.docx --render
```
五关全过才交付：①XML良构 ②图片数前后一致 ③正文零丢失
（内容手术产生的字符差异会逐条列出，人工确认为等价改写） ④zip回读 ⑤页数报告。
必要时渲染成图目视检查（重点看做过手术的页）。

### Phase 5 — 交付
- 报告格式：改动清单表 + 验收证据 + 页数变化说明 + 剩余已知瑕疵
- 提醒用户做 30 秒真机验收：全选复制 → 贴到空白文档 → 检查原出错段落

## 踩坑备忘

- LibreOffice 渲染对浮动图宽容（自动摆正），**复现不了粘贴乱版**；判断是否修好要看
  anchor/sectPr 计数归零，真机验收靠用户
- 空分节符段删除后，footer1~N.xml 变成孤儿部件，无害不必清理
- `is_empty_para` 必须排除含 sectPr/drawing 的段，否则折叠会误删
- 内容手术优先保阅读顺序与 PDF 一致，视觉位置变化是可接受代价
- 页数变化是预期副作用：解锁分页 + 折叠空段后由内容自然撑页

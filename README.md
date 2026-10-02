# wps-docx-copyfix — WPS 转换 docx 复制友好化（Claude Code Skill）

把"WPS 把 PDF 转 Word"得到的文档修成**可以放心复制粘贴**的版本，正文一字不丢。
适合：教师拿讲义 PDF 转 Word 后要摘题、组卷、贴到别的文档，结果一贴就乱。

## 为什么 WPS 转的 Word 一贴就乱

WPS 转换的目标是像素级复刻 PDF，为此埋了三类"定时炸弹"：

| 根因 | WPS 干了什么 | 粘贴后症状 |
|---|---|---|
| ① 浮动锚定图 `wp:anchor` | 部分图用"相对锚点段落 X/Y 偏移 + 浮动"放置 | 锚点换位后图按旧坐标漂移：图压文字、堆右侧 |
| ② 空分节符段 `pPr+sectPr` | PDF 每页一个空分节符段锁分页（复刻页码） | 强制换页被带进目标文档，跨页内容拦腰拆开 |
| ③ 定位缩进 `w:ind` | 一视觉行一段 + 大缩进模拟 PDF 位置（可达 16cm） | "每行超出"、整行硬换行 |

WPS 打开原文档看着正常，是因为它重新结算浮动坐标；**粘贴时不重算**，按原坐标硬放。

## 修复流水线

```bash
# ① 诊断（含原 PDF 阅读顺序对照）
python3 scripts/diagnose.py 输入.docx --pdf 原版.pdf -o _diag

# ② 机械修复四件套（幂等）：浮动图转内联 / 删空分节符段 / 去缩进 / 折叠空段
python3 scripts/auto_fix.py 输入.docx -o 整理版.docx --center-title --a4-2cm

# ③ 内容手术（折行合并/顺序修复/公式线性化）—— Claude 按 dump 对照 PDF 逐处做

# ④ 五道验收：XML良构/图片数/正文零丢失/zip回读/渲染页数
python3 scripts/verify.py 输入.docx 整理版.docx --render
```

在 Claude Code 中说"WPS 转的文档复制出来乱了 / 整理一下"即触发本技能全流程。

## 脚本

| 脚本 | 用途 |
|---|---|
| `scripts/diagnose.py` | 体检报告：浮动图/分节符/缩进/空段 + 全段落 dump + PDF ground truth |
| `scripts/auto_fix.py` | 机械修复四件套 + 可选页面设置，幂等可重跑 |
| `scripts/verify.py` | 五道验收关卡（忽略空白的字符级零丢失 diff） |
| `scripts/docx_lib.py` | 段落/run 级手术原语库（内容手术用） |

依赖：python3 标准库 + `xmllint`；渲染验收需 `soffice` + poppler(`pdftotext/pdfinfo`)。

## 实战记录

- 初中物理 第4讲 电功率2（7 页学生版讲义）：
  5 张浮动图转内联、5 个空分节符段删除、21 个空段折叠、100+ 处定位缩进清除、
  6 处内容手术（例题4 顺序错乱修复、分数 u2/R 线性化、A/B/C/D 选项分行、折行合并）。
  验收：97 张图一张不少、字符级零丢失、粘贴跨页题（例题1/练习2/练习5）不再拆页 ✅

## 安装

```bash
cp -r SKILL.md scripts ~/.claude/skills/wps-docx-copyfix/
```

## License

MIT

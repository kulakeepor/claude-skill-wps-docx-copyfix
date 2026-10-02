# -*- coding: utf-8 -*-
"""docx XML 手术工具库：解包/重组、段落与 run 级操作、文本提取。

所有操作只动 word/document.xml 的 <w:p> 层，其余部件（媒体/样式/关系）原样保留。
"""
import os
import re
import shutil
import subprocess
import zipfile

PARA_RE = re.compile(r'<w:p\b[^>]*>.*?</w:p>|<w:p\b[^>]*/>', re.S)
EMU_PER_CM = 360000.0


# ── 解包 / 重组 ──────────────────────────────────────────────

def unpack(docx_path, workdir):
    """解包 docx 到 workdir/src/，返回 document.xml 路径。"""
    src = os.path.join(workdir, 'src')
    if os.path.exists(src):
        shutil.rmtree(src)
    os.makedirs(src)
    with zipfile.ZipFile(docx_path) as z:
        z.extractall(src)
    return os.path.join(src, 'word', 'document.xml')


def repack(docx_path, workdir):
    """把 workdir/src/ 全量重打包进 docx_path（原子替换）。"""
    src = os.path.join(workdir, 'src')
    tmp = docx_path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, _, files in os.walk(src):
            for f in sorted(files):
                full = os.path.join(root, f)
                zout.write(full, os.path.relpath(full, src))
    os.replace(tmp, docx_path)


def xmllint_ok(xml_path):
    """校验 XML 是否良构。"""
    r = subprocess.run(['xmllint', '--noout', xml_path],
                       capture_output=True, text=True)
    return r.returncode == 0, r.stderr.strip()


# ── 段落读写 ─────────────────────────────────────────────────

def para_spans(xml):
    return [(m.start(), m.end()) for m in PARA_RE.finditer(xml)]


def rebuild(xml, paras, spans=None):
    """按 spans 逆序替换段落。paras[i] 为 None 表示删除；可含多段拼接串。"""
    if spans is None:
        spans = para_spans(xml)
    assert len(spans) == len(paras), f'spans={len(spans)} paras={len(paras)}'
    for (s, e), p in reversed(list(zip(spans, paras))):
        if p is None:
            xml = xml[:s] + xml[e:]
        elif p != xml[s:e]:
            xml = xml[:s] + p + xml[e:]
    return xml


# ── 段落解析 ─────────────────────────────────────────────────

def para_text(p):
    """段落可见文本；图片位置以 ▣ 标记（用于阅读顺序比对）。"""
    parts = []
    for r in re.findall(r'<w:r\b[^>]*>(.*?)</w:r>', p, re.S):
        if '<w:drawing>' in r:
            parts.append('▣')
        else:
            t = ''.join(re.findall(r'<w:t\b[^>]*>(.*?)</w:t>', r, re.S))
            if t:
                parts.append(t)
    return ''.join(parts)


def runs_of(p):
    """段落 pPr 之后的全部 run 原文（含图片 run）。"""
    m = re.search(r'</w:pPr>', p)
    start = m.end() if m else re.search(r'<w:p\b[^>]*>', p).end()
    return p[start:p.rfind('</w:p>')]


def with_runs(p, body):
    """保留 pPr，替换段落 run 区。"""
    m = re.search(r'</w:pPr>', p)
    start = m.end() if m else re.search(r'<w:p\b[^>]*>', p).end()
    return p[:start] + body + '</w:p>'


def bare_p(ppr_src, body):
    """以 ppr_src 的 pPr 为模板构造新段落（不带 paraId，避免重复）。"""
    m = re.search(r'<w:pPr>.*?</w:pPr>', ppr_src, re.S)
    return '<w:p>' + (m.group(0) if m else '') + body + '</w:p>'


def is_empty_para(p):
    """无文字、无图片、无分节属性的空段（可安全折叠）。"""
    if '<w:drawing>' in p or '<w:sectPr' in p:
        return False
    return not para_text(p).strip()


def is_sect_break_para(p):
    """仅含 pPr+sectPr 的空分节符段（WPS 复刻 PDF 分页的产物，可整段删除）。"""
    if '<w:sectPr' not in p or '<w:drawing>' in p:
        return False
    if para_text(p).strip():
        return False
    m = re.search(r'<w:pPr>(.*?)</w:pPr>', p, re.S)
    return m is not None and '<w:sectPr' in m.group(1)

# -*- coding: utf-8 -*-
"""机械修复（安全通用，幂等）：浮动图转内联 / 删空分节符段 / 去缩进 / 折叠空段 / 页面设置。

用法:
    python3 auto_fix.py 输入.docx -o 输出.docx [--center-title]
                        [--a4-2cm | --margins 上,右,下,左(twips)]
"""
import argparse
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_lib import (is_empty_para, is_sect_break_para,  # noqa: E402
                      para_spans, para_text, rebuild, unpack, repack, xmllint_ok)


def fix_anchor_to_inline(xml, log):
    def fix(m):
        inner = m.group(1)
        inner = re.sub(r'<wp:simplePos[^/]*/>', '', inner)
        inner = re.sub(r'<wp:positionH[^>]*>.*?</wp:positionH>', '', inner, flags=re.S)
        inner = re.sub(r'<wp:positionV[^>]*>.*?</wp:positionV>', '', inner, flags=re.S)
        inner = re.sub(r'<wp:wrap\w+\s*/>', '', inner)
        dist = re.search(r'distT="[^"]*"\s+distB="[^"]*"\s+distL="[^"]*"\s+distR="[^"]*"',
                         m.group(0))
        d = dist.group(0) if dist else 'distT="0" distB="0" distL="0" distR="0"'
        return f'<wp:inline {d}>{inner}</wp:inline>'
    n = len(re.findall(r'<wp:anchor ', xml))
    xml = re.sub(r'<wp:anchor [^>]*>(.*?)</wp:anchor>', fix, xml, flags=re.S)
    if n:
        log.append(f'浮动图转内联: {n} 张 wp:anchor → wp:inline')
    return xml


def fix_strip_ind(xml, log, center_title):
    spans = para_spans(xml)
    paras = [xml[s:e] for s, e in spans]
    n_ind = 0
    centered = False
    for i, p in enumerate(paras):
        if '<w:ind ' not in p:
            continue
        if center_title and not centered and para_text(p).strip():
            # 标题段：ind → jc center
            paras[i] = re.sub(r'<w:ind [^/]*/>', '<w:jc w:val="center"/>', p)
            centered = True
            n_ind += 1
            continue
        paras[i] = re.sub(r'<w:ind [^/]*/>', '', p)
        n_ind += 1
    xml = rebuild(xml, paras, spans)
    if n_ind:
        extra = '（标题居中）' if center_title else ''
        log.append(f'去定位缩进: 删除 {n_ind} 处 w:ind{extra}')
    return xml


def fix_sect_breaks(xml, log):
    spans = para_spans(xml)
    paras = [xml[s:e] for s, e in spans]
    n = 0
    for i, p in enumerate(paras):
        if is_sect_break_para(p):
            paras[i] = None
            n += 1
    xml = rebuild(xml, paras, spans)
    if n:
        log.append(f'删空分节符段: {n} 个（解锁 PDF 分页）')
    return xml


def fix_collapse_empty(xml, log):
    spans = para_spans(xml)
    paras = [xml[s:e] for s, e in spans]
    result, prev_empty, n = [], False, 0
    for p in paras:
        e = is_empty_para(p)
        if e and prev_empty:
            result.append(None)
            n += 1
        else:
            result.append(p)
        prev_empty = e
    xml = rebuild(xml, result, spans) if n else xml
    if n:
        log.append(f'折叠连续空段: 删除 {n} 个')
    return xml


def fix_page_setup(xml, log, a4, margins):
    if a4:
        xml = re.sub(r'<w:pgSz[^/]*/>', '<w:pgSz w:w="11906" w:h="16838"/>', xml)
        xml = re.sub(r'<w:pgMar[^/]*/>',
                     '<w:pgMar w:top="1247" w:right="1247" w:bottom="1247" w:left="1247" '
                     'w:header="720" w:footer="720" w:gutter="0"/>', xml)
        log.append('页面设置: A4 + 四边 2.2cm')
    elif margins:
        t, r, b, l = margins
        xml = re.sub(r'<w:pgMar[^/]*/>',
                     f'<w:pgMar w:top="{t}" w:right="{r}" w:bottom="{b}" w:left="{l}" '
                     'w:header="720" w:footer="720" w:gutter="0"/>', xml)
        log.append(f'页面设置: 边距 {t},{r},{b},{l} twips')
    return xml


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('input')
    ap.add_argument('-o', '--output', required=True)
    ap.add_argument('--center-title', action='store_true', help='首个非空段(标题)居中')
    g = ap.add_mutually_exclusive_group()
    g.add_argument('--a4-2cm', action='store_true', help='A4 + 四边 1247 twips(2.2cm)')
    g.add_argument('--margins', help='自定义边距 上,右,下,左 (twips)')
    args = ap.parse_args()

    workdir = os.path.join(os.path.dirname(os.path.abspath(args.output)) or '.', '_autofix_work')
    if os.path.exists(workdir):
        shutil.rmtree(workdir)
    os.makedirs(workdir)

    if os.path.abspath(args.input) != os.path.abspath(args.output):
        shutil.copy(args.input, args.output)
    doc_xml = unpack(args.output, workdir)
    xml = open(doc_xml, encoding='utf-8').read()

    margins = None
    if args.margins:
        margins = [int(x) for x in args.margins.split(',')]

    log = []
    # 顺序重要：先删分节符段，再折叠空段
    xml = fix_anchor_to_inline(xml, log)
    xml = fix_sect_breaks(xml, log)
    xml = fix_strip_ind(xml, log, args.center_title)
    xml = fix_collapse_empty(xml, log)
    xml = fix_page_setup(xml, log, args.a4_2cm, margins)

    open(doc_xml, 'w', encoding='utf-8').write(xml)
    ok, err = xmllint_ok(doc_xml)
    if not ok:
        print('✗ XML 校验失败:', err)
        sys.exit(1)
    repack(args.output, workdir)
    shutil.rmtree(workdir)

    print('=== 改动清单 ===')
    for line in log:
        print(' •', line)
    if not log:
        print(' （无改动，已是干净状态——幂等）')


if __name__ == '__main__':
    main()

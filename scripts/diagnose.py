# -*- coding: utf-8 -*-
"""诊断 WPS 转换 docx：浮动图 / 分节符 / 缩进 / 空段 / 阅读顺序全量报告。

用法:
    python3 diagnose.py 文件.docx [--pdf 原版.pdf] [-o 报告目录]
"""
import argparse
import os
import re
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_lib import EMU_PER_CM, unpack, para_text  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--pdf', help='原版 PDF，提取阅读顺序 ground truth')
    ap.add_argument('-o', default='_diag', help='报告输出目录')
    args = ap.parse_args()

    if os.path.exists(args.o):
        shutil.rmtree(args.o)
    os.makedirs(args.o)
    doc_xml = unpack(args.docx, args.o)
    xml = open(doc_xml, encoding='utf-8').read()
    paras = re.findall(r'<w:p\b[^>]*>.*?</w:p>|<w:p\b[^>]*/>', xml, re.S)

    print('═' * 60)
    print('① 概览')
    print('═' * 60)
    print(f'  段落数          : {len(paras)}')
    print(f'  图片 inline     : {xml.count("<wp:inline")}')
    print(f'  图片 anchor(浮动): {xml.count("<wp:anchor")}')
    print(f'  分节符 sectPr   : {len(re.findall(chr(60) + "w:sectPr", xml))}')
    print(f'  硬分页 br page  : {len(re.findall(r"<w:br w:type=.page", xml))}')
    print(f'  文本框 txbx     : {xml.count("txbxContent")}')

    print('═' * 60)
    print('② 浮动图明细（粘贴错乱的头号元凶）')
    print('═' * 60)
    n = 0
    for i, p in enumerate(paras):
        for m in re.finditer(r'<wp:anchor[^>]*>.*?</wp:anchor>', p, re.S):
            a = m.group(0)
            n += 1
            name = re.search(r'name="([^"]+)"', a)
            ext = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"/>', a)
            ph = re.search(r'<wp:positionH relativeFrom="([^"]+)">(.*?)</wp:positionH>', a, re.S)
            pv = re.search(r'<wp:positionV relativeFrom="([^"]+)">(.*?)</wp:positionV>', a, re.S)
            wrap = 'wrapNone' if '<wp:wrapNone/>' in a else '有环绕'
            w = int(ext.group(1)) / EMU_PER_CM if ext else 0
            h = int(ext.group(2)) / EMU_PER_CM if ext else 0
            print(f'  段{i:3d} {name.group(1) if name else "?":10s} '
                  f'{w:.1f}x{h:.1f}cm {wrap}')
            print(f'        H:{ph.group(1)} {re.sub(chr(60)+"[^>]+"+chr(62), "", ph.group(2))}'
                  f'  V:{pv.group(1)} {re.sub(chr(60)+"[^>]+"+chr(62), "", pv.group(2))}'
                  f'  |{para_text(p)[:30]}|')
    if n == 0:
        print('  （无浮动图 ✓）')

    print('═' * 60)
    print('③ 空分节符段（锁分页的元凶，WPS 复刻 PDF 每页一个）')
    print('═' * 60)
    from docx_lib import is_sect_break_para
    n = 0
    for i, p in enumerate(paras):
        if is_sect_break_para(p):
            n += 1
            prev = para_text(paras[i - 1])[:26] if i else ''
            nxt = para_text(paras[i + 1])[:26] if i + 1 < len(paras) else ''
            print(f'  段{i:3d}  上文:|{prev}|  下文:|{nxt}|')
    if n == 0:
        print('  （无 ✓）')

    print('═' * 60)
    print('④ w:ind 缩进分布（复制带出"超边"的元凶）')
    print('═' * 60)
    inds = []
    for i, p in enumerate(paras):
        m = re.search(r'<w:ind [^/]*/>', p)
        if m:
            left = re.search(r'w:left="(\d+)"', m.group(0))
            inds.append((int(left.group(1)) if left else 0, i))
    inds.sort(reverse=True)
    print(f'  带缩进段落: {len(inds)} 个')
    for v, i in inds[:5]:
        print(f'    段{i:3d} left={v} twips ({v/567:.1f}cm) |{para_text(paras[i])[:30]}|')

    print('═' * 60)
    print('⑤ 连续空段 runs（可折叠）')
    print('═' * 60)
    from docx_lib import is_empty_para
    run_start, run_len = None, 0
    for i, p in enumerate(paras + ['<w:p>END</w:p>']):
        if is_empty_para(p):
            if run_start is None:
                run_start = i
            run_len += 1
        else:
            if run_len >= 2:
                print(f'  段{run_start:3d} 起连续 {run_len} 个空段')
            run_start, run_len = None, 0

    # 全段落 dump（供与 PDF 对照阅读顺序）
    dump_path = os.path.join(args.o, 'paras_dump.txt')
    with open(dump_path, 'w', encoding='utf-8') as f:
        for i, p in enumerate(paras):
            t = para_text(p)
            if t.strip():
                f.write(f'{i:4d} |{t}|\n')
    print(f'\n段落全文 dump: {dump_path}')

    if args.pdf:
        gt = os.path.join(args.o, 'pdf_ground_truth.txt')
        subprocess.run(['pdftotext', '-layout', args.pdf, gt], check=True)
        print(f'PDF 阅读顺序 : {gt}')
        print('→ 人工比对：dump 中被拆散的句子 / 顺序错乱处，做内容手术')


if __name__ == '__main__':
    main()

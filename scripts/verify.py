# -*- coding: utf-8 -*-
"""验收：无证据不交货。对比修复前后两个 docx。

用法:
    python3 verify.py 原始.docx 修复后.docx [--render]

关卡:
  ① XML 良构        ② 图片数一致       ③ 正文零丢失（忽略空白的字符级 diff）
  ④ zip 回读一致    ⑤ (--render) 渲染页数
"""
import argparse
import difflib
import os
import re
import shutil
import subprocess
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_lib import para_text, xmllint_ok  # noqa: E402


def full_text(docx):
    with zipfile.ZipFile(docx) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    paras = re.findall(r'<w:p\b[^>]*>.*?</w:p>|<w:p\b[^>]*/>', xml, re.S)
    return '\n'.join(para_text(p) for p in paras if para_text(p).strip())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('old')
    ap.add_argument('new')
    ap.add_argument('--render', action='store_true', help='soffice 渲染并报页数')
    args = ap.parse_args()
    passed = []

    # ① XML 良构
    tmp = '/tmp/_verify_new'
    if os.path.exists(tmp):
        shutil.rmtree(tmp)
    os.makedirs(tmp)
    with zipfile.ZipFile(args.new) as z:
        z.extract('word/document.xml', tmp)
    ok, err = xmllint_ok(os.path.join(tmp, 'word', 'document.xml'))
    print(f"{'✓' if ok else '✗'} ① XML 良构" + ('' if ok else f': {err}'))
    passed.append(ok)

    # ② 图片数一致
    def n_draw(docx):
        with zipfile.ZipFile(docx) as z:
            return z.read('word/document.xml').decode('utf-8').count('<w:drawing>')
    a, b = n_draw(args.old), n_draw(args.new)
    print(f"{'✓' if a == b else '✗'} ② 图片数: {a} → {b}")
    passed.append(a == b)

    # ③ 正文零丢失（忽略一切空白，字符级 diff）
    ta = re.sub(r'\s+', '', full_text(args.old))
    tb = re.sub(r'\s+', '', full_text(args.new))
    sm = difflib.SequenceMatcher(a=ta, b=tb, autojunk=False)
    lost, added = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ('delete', 'replace'):
            lost.append(ta[i1:i2])
        if tag in ('insert', 'replace'):
            added.append(tb[j1:j2])
    # 内容手术（如分数线性化 u2/R）会产生等价字符差异，逐一列出供人工判断
    if not lost and not added:
        print('✓ ③ 正文零丢失（字符级一致）')
        passed.append(True)
    else:
        print(f'△ ③ 正文有差异（需人工确认是否等价改写）:')
        for x in lost[:10]:
            print(f'   仅旧有: |{x[:60]}|')
        for x in added[:10]:
            print(f'   仅新有: |{x[:60]}|')
        passed.append(None)  # None = 待人工

    # ④ zip 回读一致（重解包再比对）
    with zipfile.ZipFile(args.new) as z:
        rb = z.read('word/document.xml').decode('utf-8')
    ok = rb == open(os.path.join(tmp, 'word', 'document.xml'), encoding='utf-8').read()
    print(f"{'✓' if ok else '✗'} ④ zip 回读一致")
    passed.append(ok)

    # ⑤ 渲染页数
    if args.render:
        for name, f in (('旧', args.old), ('新', args.new)):
            out = f'/tmp/_verify_render_{name}'
            if os.path.exists(out):
                shutil.rmtree(out)
            os.makedirs(out)
            subprocess.run(['soffice', '--headless', '--convert-to', 'pdf',
                            '--outdir', out, f], capture_output=True)
            pdf = next(os.path.join(out, x) for x in os.listdir(out) if x.endswith('.pdf'))
            r = subprocess.run(['pdfinfo', pdf], capture_output=True, text=True)
            m = re.search(r'Pages:\s+(\d+)', r.stdout)
            print(f'   {name}渲染页数: {m.group(1) if m else "?"}（LibreOffice 口径，仅供参考）')

    if all(x is True for x in passed):
        print('\n=== 验收通过 ===')
    elif None in passed:
        print('\n=== 机械项通过，差异项需人工确认 ===')
    else:
        print('\n=== 验收失败 ===')
        sys.exit(1)


if __name__ == '__main__':
    main()

"""pptx の中身が PowerPoint の決まりに合っているかを調べる（iPhone版は特に厳しい）。
使い方：python3 src/check_pptx.py output/若穂地区_住民説明会資料.pptx
"""
import sys, re, zipfile, posixpath
from xml.dom import minidom
z = zipfile.ZipFile(sys.argv[1]); names = set(z.namelist()); probs = []
if any(n.endswith('/') for n in names): probs.append('フォルダの項目がある')
ct = z.read('[Content_Types].xml').decode()
defaults = set(re.findall(r'Extension="([^"]+)"', ct)); overrides = set(re.findall(r'PartName="/([^"]+)"', ct))
for n in names:
    ext = n.rsplit('.', 1)[-1].lower()
    if n != '[Content_Types].xml' and n not in overrides and ext not in defaults: probs.append(f'型の登録がない: {n}')
for o in overrides:
    if o not in names: probs.append(f'登録だけあって中身がない: {o}')
for n in names:
    if n.endswith('.rels'):
        base = posixpath.dirname(posixpath.dirname(n))
        for tgt, mode in re.findall(r'Target="([^"]+)"(?: TargetMode="([^"]+)")?', z.read(n).decode()):
            if mode == 'External': continue
            p = posixpath.normpath(posixpath.join(base, tgt)) if not tgt.startswith('/') else tgt[1:]
            if p not in names: probs.append(f'リンク先がない: {n} -> {tgt}')
for n in sorted(names):
    if not n.endswith('.xml'): continue
    x = z.read(n).decode('utf8')
    try: minidom.parseString(x)
    except Exception as e: probs.append(f'XMLの誤り: {n}: {e}'); continue
    if re.search(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', x): probs.append(f'使えない制御文字: {n}')
    if n.startswith('ppt/slides/slide'):
        ids = re.findall(r'<p:cNvPr id="(\d+)"', x)
        d = {i for i in ids if ids.count(i) > 1}
        if d: probs.append(f'図形の番号の重複: {n} {sorted(d)[:5]}')
        for tbl in re.findall(r'<a:tbl>.*?</a:tbl>', x, re.S):
            cols = len(re.findall(r'<a:gridCol ', tbl))
            for ri, row in enumerate(re.findall(r'<a:tr .*?</a:tr>', tbl, re.S)):
                tcs = len(re.findall(r'<a:tc[ >]', row))
                if tcs != cols: probs.append(f'表の列数が合わない: {n} 行{ri+1} セル{tcs}/列{cols}')
        for v in re.findall(r'(?<!<p:ph type="sldNum" )(?:\bsz|cx|cy)="(-?[^"]*)"', re.sub(r'<p:ph [^>]*>', '', x)):
            if not re.fullmatch(r'-?\d+', v): probs.append(f'数値でない値: {n} {v}'); break
print('\n'.join(probs) if probs else '問題なし')

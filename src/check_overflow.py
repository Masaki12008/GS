"""pptx の各テキスト枠について、BIZ UDPゴシックの字幅で折り返しを計算し、
枠の高さに収まらない（PowerPoint で下の文字に重なるおそれがある）ものを一覧にする。
使い方：python3 src/check_overflow.py output/若穂地区_住民説明会資料.pptx [余裕の割合 0.9]
"""
import sys, os
from pptx import Presentation
from pptx.util import Emu
from fontTools.ttLib import TTFont
FD = os.path.expanduser('~/.fonts')
fonts = {False: TTFont(f'{FD}/BIZUDPGothic-Regular.ttf'), True: TTFont(f'{FD}/BIZUDPGothic-Bold.ttf')}
cmap = {b: f.getBestCmap() for b, f in fonts.items()}
hmtx = {b: f['hmtx'] for b, f in fonts.items()}
UPM = 2048
def width(ch, pt, bold):
    g = cmap[bold].get(ord(ch))
    return (hmtx[bold][g][0] if g else UPM) / UPM * pt
def lines_needed(text, pt, bold, wpt):
    n = 0
    for para in text.split('\n'):
        w = 0; n += 1
        for ch in para:
            cw = width(ch, pt, bold)
            if w + cw > wpt: n += 1; w = cw
            else: w += cw
    return n
slack = float(sys.argv[2]) if len(sys.argv) > 2 else 0.97
prs = Presentation(sys.argv[1])
for si, sl in enumerate(prs.slides, 1):
    for sh in sl.shapes:
        if not sh.has_text_frame or not sh.text_frame.text.strip(): continue
        tf = sh.text_frame
        bp = tf._txBody.find('{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
        ins = lambda k, d: int(bp.get(k, d))
        wpt = Emu(sh.width - ins('lIns', 91440) - ins('rIns', 91440)).pt * slack
        hpt = Emu(sh.height - ins('tIns', 45720) - ins('bIns', 45720)).pt
        need = 0
        for p in tf.paragraphs:
            runs = p.runs or []
            pt = max([r.font.size.pt for r in runs if r.font.size] or [18])
            bold = any(r.font.bold for r in runs)
            txt = ''.join(r.text for r in runs).replace('\x0b', '\n')
            need += lines_needed(txt, pt, bold, wpt) * pt * 1.2
        if need > hpt + 2:
            print(f'P{si}\t必要{need:.0f}pt / 枠{hpt:.0f}pt\t{tf.text[:40]!r}')

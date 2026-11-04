"""若穂の学校区と、0〜14歳の子どもが住む場所（250mメッシュ）の地図を描く → data/map_wakaho.png
先に src/tsugaku/tsugaku.py を実行しておく。背景は地理院タイル（淡色地図）をダウンロードする（ネット接続が必要）。
実行：pip install shapely matplotlib pillow → python3 src/tsugaku/map.py
"""
import json,math,os,csv,io,urllib.request,sys
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from tsugaku import mesh_bounds, IN, OUT
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import Polygon as MP, Rectangle
from PIL import Image
from shapely.geometry import shape
fp='/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf'; fm.fontManager.addfont(fp); plt.rcParams['font.family']=fm.FontProperties(fname=fp).get_name()
z=14; X0,X1,Y0,Y1=14482,14487,6397,6403
def px(lon,lat):
    n=2**z; x=(lon+180)/360*n; y=(1-math.log(math.tan(math.radians(lat))+1/math.cos(math.radians(lat)))/math.pi)/2*n
    return (x-X0)*256,(y-Y0)*256
W=(X1-X0+1)*256; H=(Y1-Y0+1)*256
base=Image.new('RGB',(W,H))
for x in range(X0,X1+1):
    for y in range(Y0,Y1+1):
        b=urllib.request.urlopen(urllib.request.Request(f'https://cyberjapandata.gsi.go.jp/xyz/pale/{z}/{x}/{y}.png',headers={'User-Agent':'wakaho-school-study'})).read()
        base.paste(Image.open(io.BytesIO(b)).convert('RGB'),((x-X0)*256,(y-Y0)*256))
g=json.load(open(os.path.join(IN,'国勢調査2020_小地域_若穂地区.geojson'),encoding='utf-8'))
D={f['properties']['地区']:shape(f['geometry']) for f in g['features']}
M=[dict(b=mesh_bounds(r['メッシュコード']),kids=int(r['0〜14歳2020'])) for r in csv.DictReader(open(os.path.join(OUT,'通学距離_メッシュ.csv'),encoding='utf-8'))]
NAVY='#1D4636'; RAMP=['#D3E8DC','#9CCBB0','#55A07D','#1D6B4A']; BINS=[(1,4,'1〜4人'),(5,9,'5〜9人'),(10,19,'10〜19人'),(20,999,'20人以上')]
U=None
from shapely.ops import unary_union
U=unary_union(list(D.values())); bx=U.bounds
xa,ya=px(bx[0]-0.004,bx[3]+0.003); xb,yb=px(bx[2]+0.004,36.553)
fig=plt.figure(figsize=(6.2,(yb-ya)/(xb-xa)*6.2),dpi=220); ax=fig.add_axes([0,0,1,1])
ax.imshow(base,extent=(0,W,H,0)); ax.set_xlim(xa,xb); ax.set_ylim(yb,ya); ax.axis('off')
# dim outside
for m in M:
    if m['kids']<=0: continue
    c=next(RAMP[i] for i,(a,b,_) in enumerate(BINS) if a<=m['kids']<=b)
    x0,y0=px(m['b'][0],m['b'][3]); x1,y1=px(m['b'][2],m['b'][1])
    ax.add_patch(Rectangle((x0,y0),x1-x0,y1-y0,facecolor=c,edgecolor='white',linewidth=0.6,alpha=0.92,zorder=2))
for k,p in D.items():
    for poly in (p.geoms if p.geom_type=='MultiPolygon' else [p]):
        ax.add_patch(MP([px(*c) for c in poly.exterior.coords],closed=True,fill=False,edgecolor=NAVY,linewidth=2.2,zorder=3))
SCH={'綿内小':(138.262529,36.622361,(10,-6),'left'),'川田小':(138.241565,36.604091,(-10,6),'right'),'保科小':(138.2655742,36.5874875,(10,-4),'left'),'若穂中':(138.251224,36.604752,(10,8),'left')}
for n,(lo,la,off,ha) in SCH.items():
    x,y=px(lo,la); ax.plot(x,y,'o',ms=11,mfc='white',mec=NAVY,mew=2.4,zorder=5)
    ax.annotate(n,(x,y),xytext=off,textcoords='offset points',ha=ha,va='center',fontsize=14,fontweight='bold',color=NAVY,zorder=6,
                bbox=dict(boxstyle='round,pad=0.2',fc='white',ec='none',alpha=0.85))
LAB={'綿内':(138.293,36.632),'川田':(138.2395,36.6185),'保科':(138.293,36.574)}
for k,(lo,la) in LAB.items():
    x,y=px(lo,la); ax.text(x,y,k+'地区',fontsize=16,fontweight='bold',color=NAVY,ha='center',va='center',zorder=6,bbox=dict(boxstyle='round,pad=0.3',fc='white',ec=NAVY,lw=1))
# legend
from matplotlib.patches import Patch
hs=[Patch(facecolor=c,edgecolor='white',label=l) for c,(_,_,l) in zip(RAMP,BINS)]
leg=ax.legend(handles=hs,title='0〜14歳の人数\n（250m四方、\n令和2年＝2020年）',loc='lower left',fontsize=12,title_fontsize=12,frameon=True,framealpha=0.92,borderpad=0.6)
ax.text(xb-6,yb-6,'保科地区は、この下（南）の山地まで続きます。地区の線は国勢調査の区域（大字）。地図：地理院タイル（淡色地図）',fontsize=7,color='#56635D',ha='right',va='bottom',zorder=7,bbox=dict(boxstyle='square,pad=0.25',fc='white',ec='none',alpha=0.9))
fig.savefig(os.path.join(OUT,'map_wakaho.png'),dpi=220)
print(fig.get_size_inches())
# ファイルを軽くする（256色に減色）
_p=os.path.join(OUT,'map_wakaho.png'); Image.open(_p).convert('RGB').quantize(256,dither=Image.Dither.NONE).save(_p,optimize=True)

# A3（たて）の配布用の地図 → output/地図_若穂地区_A3.pdf
import csv as _csv
fig.savefig(os.path.join(OUT,'.map_hi.png'),dpi=330)
hi=Image.open(os.path.join(OUT,'.map_hi.png'))
A3=plt.figure(figsize=(11.69,16.54))
A3.text(0.06,0.965,'若穂の地区と、子どもが住んでいる場所',fontsize=26,fontweight='bold',color=NAVY,va='top')
A3.text(0.06,0.932,'緑のマス（250m四方）は0〜14歳の子どもが住んでいる場所で、色が濃いほど人数が多いことを表します。',fontsize=13,color='#1E2A26',va='top')
axm=A3.add_axes([0.06,0.17,0.88,0.75]); axm.imshow(hi); axm.axis('off')
_rows=list(_csv.DictReader(open(os.path.join(OUT,'若穂_町別人口_R8.10.1.csv'),encoding='utf-8')))
_pop={c:sum(int(r['人口']) for r in _rows if r['地区']==c) for c in ['綿内','川田','保科']}
_kids=json.load(open(os.path.join(OUT,'通学距離_集計.json'),encoding='utf-8'))
_ji={'綿内小':273,'川田小':157,'保科小':54}  # 児童数推計（令和8年＝2026年5月1日現在、data/推計データ.csv）
lines=[f"{c}地区：人口 {_pop[c]:,}人（令和8年＝2026年10月1日）／{s}の児童 {_ji[s]}人（令和8年＝2026年5月1日）／0〜14歳 {_kids[c+'>'+s]['kids']:,}人（令和2年＝2020年、中学生も含む）" for c,s in [('綿内','綿内小'),('川田','川田小'),('保科','保科小')]]
A3.text(0.06,0.155,'\n'.join(lines),fontsize=11.5,color='#1E2A26',va='top',linespacing=1.6)
A3.text(0.06,0.075,'地区の線：国勢調査（令和2年＝2020年）の小地域（大字）。川田地区は牛島を、保科地区は若穂団地・保科温泉を含みます。\n地区の線は、通学区域（学区）と一致しない場合があります。実際の通学区域は［市に確認中］。\n出典：総務省 国勢調査 地域メッシュ統計・小地域境界（e-Stat）、長野市 町別人口及び世帯数、国土地理院 地理院タイル（淡色地図）。\n若穂地区 小中学校再編検討｜住民説明会資料（案）',fontsize=10.5,color='#56635D',va='top',linespacing=1.6)
A3.savefig(os.path.join(os.path.dirname(OUT),'output','地図_若穂地区_A3.pdf'))
os.remove(os.path.join(OUT,'.map_hi.png'))

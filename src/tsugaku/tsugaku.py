"""若穂の子どもが住む場所（250mメッシュ）から各学校までの徒歩の道のりを計算する。

入力（input/tsugaku/）
- 国勢調査2020_250mメッシュ_年齢3区分_5438.txt : e-Stat 国勢調査2020 5次メッシュ（250m）T001142（若穂周辺のみ抜粋）
- 国土数値情報_A27-23_若穂.geojson : 国土数値情報 小学校区（令和5年度）から若穂の3校を抜粋
出力（data/）
- 通学距離_メッシュ.csv : メッシュごとの0〜14歳人口と、4校までの徒歩の道のり（m）
- 通学距離_集計.json : 学区ごとの距離の分布
道のりは OpenStreetMap のデータによる経路計算（routing.openstreetmap.de の徒歩）。ネット接続が必要。
実行：pip install shapely → python3 src/tsugaku/tsugaku.py
"""
import json, csv, collections, time, urllib.request, os
from shapely.geometry import shape, Point
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
IN = os.path.join(ROOT, 'input', 'tsugaku'); OUT = os.path.join(ROOT, 'data')
SCH = {'綿内小': (138.262529, 36.622361), '川田小': (138.241565, 36.604091), '保科小': (138.2655742, 36.5874875), '若穂中': (138.251224, 36.604752)}

def mesh_bounds(code):
    c = str(code); lat = int(c[0:2]) / 1.5; lon = int(c[2:4]) + 100
    lat += int(c[4]) * (2 / 3 / 8); lon += int(c[5]) * (1 / 8)
    lat += int(c[6]) * (2 / 3 / 80); lon += int(c[7]) * (1 / 80)
    dlat, dlon = 2 / 3 / 160, 1 / 160
    for q in (int(c[8]), int(c[9])):
        lat += (q - 1) // 2 * dlat; lon += (q - 1) % 2 * dlon
        dlat /= 2; dlon /= 2
    return lon, lat, lon + dlon * 2, lat + dlat * 2

def main():
    g = json.load(open(os.path.join(IN, '国土数値情報_A27-23_若穂.geojson'), encoding='utf-8'))
    D = {f['properties']['A27_004'][:2] + '小': shape(f['geometry']) for f in g['features']}
    rows = list(csv.reader(open(os.path.join(IN, '国勢調査2020_250mメッシュ_年齢3区分_5438.txt'), encoding='cp932')))
    M = []
    for r in rows[2:]:
        b = mesh_bounds(r[0]); p = Point((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
        s = [k for k, v in D.items() if v.contains(p)]
        if not s: continue
        M.append(dict(code=r[0], school=s[0], lon=p.x, lat=p.y, b=b, pop=int(r[4]) if r[4].isdigit() else 0, kids=int(r[7]) if r[7].isdigit() else 0))
    K = [m for m in M if m['kids'] > 0]
    names = list(SCH)
    for i in range(0, len(K), 40):
        ch = K[i:i + 40]; n = len(ch)
        coords = [f"{m['lon']:.6f},{m['lat']:.6f}" for m in ch] + [f'{x},{y}' for x, y in SCH.values()]
        url = ('https://routing.openstreetmap.de/routed-foot/table/v1/foot/' + ';'.join(coords) +
               f"?sources={';'.join(map(str, range(n)))}&destinations={';'.join(map(str, range(n, n + 4)))}&annotations=distance")
        d = json.load(urllib.request.urlopen(urllib.request.Request(url, headers={'User-Agent': 'wakaho-school-study'}), timeout=60))
        for j, m in enumerate(ch): m['foot'] = {names[k]: d['distances'][j][k] for k in range(4)}
        time.sleep(1)
    with open(os.path.join(OUT, '通学距離_メッシュ.csv'), 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f); w.writerow(['メッシュコード', '学区', '経度', '緯度', '人口2020', '0〜14歳2020'] + [f'{n}まで徒歩(m)' for n in names])
        for m in M:
            w.writerow([m['code'], m['school'], f"{m['lon']:.6f}", f"{m['lat']:.6f}", m['pop'], m['kids']] + [round(m['foot'][n]) if 'foot' in m else '' for n in names])
    bins = [(0, 1000, '1km未満'), (1000, 2000, '1〜2km'), (2000, 3000, '2〜3km'), (3000, 4000, '3〜4km'), (4000, 6000, '4〜6km'), (6000, 1e9, '6km以上')]
    out = {}
    for home in ['綿内小', '川田小', '保科小', '全体']:
        sub = [m for m in K if home == '全体' or m['school'] == home]; tot = sum(m['kids'] for m in sub)
        for dest in names:
            c = collections.Counter(); wsum = 0; mx = 0
            for m in sub:
                dd = m['foot'][dest]; c[next(l for a, z, l in bins if a <= dd < z)] += m['kids']; wsum += dd * m['kids']; mx = max(mx, dd)
            out[f'{home}>{dest}'] = dict(kids=tot, bins={l: c[l] for _, _, l in bins}, avg_m=round(wsum / tot), max_m=round(mx))
    out['_pop2020'] = {k: sum(m['pop'] for m in M if m['school'] == k) for k in D}
    json.dump(out, open(os.path.join(OUT, '通学距離_集計.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

if __name__ == '__main__':
    main()

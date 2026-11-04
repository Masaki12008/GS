const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const fs = require("fs");
const path = require("path");
let applyTheme = null;
try { ({ applyTheme } = require(process.env.APPLY_THEME || "/mnt/skills/public/pptx/scripts/apply_theme.js")); } catch (e) { /* optional: theme colors only */ }

const OUT = process.argv[2] || "output/若穂地区_住民説明会資料.pptx";
const APPENDIX = process.env.APPENDIX !== "0"; // APPENDIX=0 で参考資料を除いた投影用を作る
const FONT = "BIZ UDPゴシック";
const K = { navy: "1D4636", teal: "2E7D5B", tealL: "E8F2EC", amber: "D9573E", amberL: "FCEBE6", ink: "1E2A26", sub: "56635D", line: "D7DFDA", bg: "F4F7F5", white: "FFFFFF", grey: "9AA59F", mark: "DDE7EF", hill1: "2A5C47", hill2: "3B7259", hill3: "4E8A6C" };
const THEME = { name: "Wakaho", headFontFace: FONT, bodyFontFace: FONT, colors: { dk1: K.ink, lt1: K.white, dk2: K.navy, lt2: K.bg, accent1: K.teal, accent2: K.amber, accent3: K.navy, accent4: "6FA8C0", accent5: K.grey, accent6: "B5543A", hlink: K.teal, folHlink: "6B4E9B" } };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "若穂地区 住民説明会資料（案）";
const FOOT = "若穂地区 小中学校再編検討｜住民説明会資料（案）";

pres.defineSlideMaster({ title: "COVER", background: { color: K.navy }, objects: [] });
pres.defineSlideMaster({
  title: "CONTENT", background: { color: K.white },
  objects: [
    { text: { text: FOOT, options: { x: 0.6, y: 7.05, w: 7, h: 0.25, fontSize: 10, color: K.grey, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.7, fontSize: 30, bold: true, color: K.navy, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: K.grey, align: "right" },
});
pres.defineSlideMaster({
  title: "DARK", background: { color: K.navy },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.5, w: 12.1, h: 0.8, fontSize: 32, bold: true, color: K.white, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: "9DB8AB", align: "right" },
});

const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 16, color: K.ink, valign: "top", fontFace: FONT }, o));
const box = (s, x, y, w, h, fill, line) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: line ? { color: line, width: 1 } : { type: "none" }, rectRadius: 0.12 });
const sub = (s, text) => T(s, text, { x: 0.6, y: 1.12, w: 12.1, h: 0.4, fontSize: 16, color: K.sub });
const note = (s, text, y = 6.55) => T(s, text, { x: 0.6, y, w: 12.1, h: 0.3, fontSize: 13, color: K.sub });
let SEC = "";
const content = (title, subtitle) => { const s = pres.addSlide({ masterName: "CONTENT" }); if (SEC) T(s, SEC, { x: 0.6, y: 0.2, w: 6, h: 0.3, fontSize: 14, bold: true, color: K.amber }); s.addText(title, { placeholder: "title" }); if (subtitle) sub(s, subtitle); return s; };


const hills = (s) => {
  s.addShape(pres.shapes.OVAL, { x: -2.5, y: 5.6, w: 9.5, h: 4.2, fill: { color: K.hill1 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: 5.2, y: 6.0, w: 10.5, h: 4.0, fill: { color: K.hill2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: 2.0, y: 6.6, w: 8.0, h: 3.0, fill: { color: K.hill3 }, line: { type: "none" } });
};
const divider = (n, name, desc) => {
  const s = pres.addSlide({ masterName: "COVER" });
  hills(s);
  T(s, n, { x: 0.9, y: 1.2, w: 3, h: 1.6, fontSize: 96, bold: true, color: K.amber });
  T(s, name, { x: 0.9, y: 2.95, w: 11, h: 1.0, fontSize: 48, bold: true, color: K.white });
  T(s, desc, { x: 0.9, y: 4.05, w: 11, h: 0.6, fontSize: 20, color: "CFE0D6" });
  return s;
};
const iconCache = {};
async function icon(name, color) {
  const key = name + color;
  if (!iconCache[key]) {
    const svg = RDS.renderToStaticMarkup(React.createElement(md[name], { color: "#" + color, size: 256 }));
    iconCache[key] = "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
  }
  return iconCache[key];
}
async function iconCircle(s, x, y, d, name, bg = K.teal, fg = K.white) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } });
  const p = d * 0.22;
  s.addImage({ data: await icon(name, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}

// data
const Y7 = ["R8", "R9", "R10", "R11", "R12", "R13", "R14"];
const sch = { 綿内小: [273, 265, 256, 245, 241, 230, 205], 川田小: [157, 142, 135, 130, 118, 101, 83], 保科小: [54, 56, 48, 49, 44, 42, 42] };
const tot = Y7.map((_, i) => sch.綿内小[i] + sch.川田小[i] + sch.保科小[i]);
const Y13 = ["R8", "R9", "R10", "R11", "R12", "R13", "R14", "R15", "R16", "R17", "R18", "R19", "R20"];
const jh = [284, 287, 266, 264, 248, 237, 225, 219, 206, 201, 180, 161, 129];
const city = [15757, 12561];
const dec = (a, b) => Math.round((1 - b / a) * 100);
// 令和→西暦（R8＝2026年度）
const AD = r => 2018 + Number(String(r).replace(/^R/, ""));
const YL = r => `${r}\n${AD(r)}`; // グラフ・表の年ラベル（令和と西暦の2段）
// 学年別推計（data/推計データ.csv）
const grades = {};
fs.readFileSync(path.join(__dirname, "../data/推計データ.csv"), "utf8").replace(/^﻿/, "").trim().split(/\r?\n/).slice(1).forEach(line => {
  const [k, g, ...v] = line.split(",");
  (grades[k] = grades[k] || {})[g] = v.filter(x => x !== "").map(Number);
});
const chartBase = { catAxisLabelColor: K.sub, valAxisLabelColor: K.sub, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", catGridLine: { style: "none" } };

// 通学距離（src/tsugaku/tsugaku.py の出力）と町別人口
const TG = JSON.parse(fs.readFileSync(path.join(__dirname, "../data/通学距離_集計.json"), "utf8"));
const MACHI = fs.readFileSync(path.join(__dirname, "../data/若穂_町別人口_R8.10.1.csv"), "utf8").trim().split(/\r?\n/).slice(1).map(l => { const [c, n, p, h] = l.split(","); return { c, n, p: +p, h: +h }; });
const fmt = n => n.toLocaleString("ja-JP");
const WA = y => y <= 1988 ? `昭和${y - 1925}` : y <= 2018 ? `平成${y - 1988 === 1 ? "元" : y - 1988}` : `令和${y - 2018 === 1 ? "元" : y - 2018}`;

const km = m => (m / 1000).toFixed(1);
const DBINS = ["1km未満", "1〜2km", "2〜3km", "3〜4km", "4〜6km", "6km以上"];
const share = (key, upto) => { const d = TG[key]; return Math.round(DBINS.slice(0, upto).reduce((a, b) => a + d.bins[b], 0) / d.kids * 100); };

(async () => {
  // 1 Cover
  {
    const s = pres.addSlide({ masterName: "COVER" });
    hills(s);
    T(s, "若穂地区 小中学校再編検討｜住民説明会資料（案）", { x: 0.9, y: 0.7, w: 9, h: 0.4, fontSize: 15, color: "CFE0D6" });
    T(s, "これからの若穂の学校を考える", { x: 0.9, y: 1.5, w: 11.5, h: 1.1, fontSize: 50, bold: true, color: K.white });
    T(s, "結論の前に、子どもたちにどんな学びを残したいか。", { x: 0.9, y: 2.75, w: 11.5, h: 0.7, fontSize: 28, bold: true, color: K.amber });
    T(s, "子どもたちの未来と、地域のこれからを一緒に考えるために", { x: 0.9, y: 3.55, w: 11, h: 0.5, fontSize: 18, color: "CFE0D6" });
    T(s, "令和8年（2026年）［　］月［　］日　若穂地区小中学校再編検討委員会／長野市教育委員会事務局 学校教育課", { x: 0.9, y: 4.45, w: 11.5, h: 0.4, fontSize: 14, color: "CFE0D6" });
    s.addNotes("表紙。日付は事務局の確定後に記入。委員会名・事務局名の表記は要確認。");
  }

  SEC = "はじめに";
  // 2 今日の流れ
  {
    const s = content("今日の説明会で共有したいこと");
    const steps = [["現状を知る", "児童生徒数・施設の状況", "MdInsights"], ["課題を共有する", "教育環境・通学・地域への影響", "MdForum"], ["選択肢を比べる", "今の形を続ける案も含めて", "MdCompareArrows"], ["意見を寄せる", "大切にしたいこと・懸念点", "MdRateReview"]];
    for (let i = 0; i < 4; i++) {
      const [h, d, ic] = steps[i], x = 0.6 + i * 3.08;
      box(s, x, 1.7, 2.82, 3.9, K.bg);
      T(s, `0${i + 1}`, { x: x + 0.3, y: 1.95, w: 1.2, h: 0.5, fontSize: 22, bold: true, color: K.amber });
      await iconCircle(s, x + 0.3, 2.6, 1.0, ic);
      T(s, h, { x: x + 0.3, y: 3.85, w: 2.3, h: 0.5, fontSize: 22, bold: true, color: K.navy });
      T(s, d, { x: x + 0.3, y: 4.45, w: 2.3, h: 0.9, fontSize: 16, color: K.sub });
    }
    box(s, 0.6, 5.9, 12.1, 0.6, K.amberL);
    T(s, "本資料は検討用の案です。再編方式・校地・実施時期が決定したことを示すものではありません。", { x: 0.85, y: 5.9, w: 11.7, h: 0.6, fontSize: 15, bold: true, color: "A3402C", valign: "middle" });
    note(s, "年は「令和8年度（2026年度）」のように併記しています。専門用語の説明は、巻末の参考資料にあります。", 6.62);
  }

  // 3 出発点
  {
    const s = content("検討の出発点｜誰のための検討か");
    box(s, 0.6, 1.6, 12.1, 1.7, K.teal);
    await iconCircle(s, 1.0, 1.9, 1.1, "MdChildCare", K.white, K.teal);
    T(s, "最優先は、子どもたちの学びと育ち", { x: 2.4, y: 1.6, w: 10, h: 1.7, fontSize: 32, bold: true, color: K.white, valign: "middle" });
    const pts = [["小規模校の良さを残し、課題を丁寧に考える", "MdFavorite"], ["通学の安全と、地域のつながりを守る", "MdShield"], ["結論を急がず、比較できる情報をそろえる", "MdFactCheck"]];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.1;
      box(s, x, 3.6, 3.9, 2.8, K.bg);
      await iconCircle(s, x + 0.35, 3.9, 0.8, pts[i][1], K.tealL, K.teal);
      T(s, pts[i][0], { x: x + 0.35, y: 4.9, w: 3.2, h: 1.3, fontSize: 19, bold: true, color: K.navy, lineSpacingMultiple: 1.15 });
    }
  }

  divider("01", "現状を知る", "児童生徒数と、これからの見通し"); SEC = "01  現状を知る";
  // 4 児童数の見通し
  {
    const s = content("若穂地区の3小学校｜児童数の見通し", "令和8年度（2026年度）→ 令和14年度（2032年度）　教育委員会の推計");
    s.addChart(pres.charts.BAR, [
      { name: "令和8年度（2026）", labels: Object.keys(sch), values: Object.values(sch).map(a => a[0]) },
      { name: "令和14年度（2032）", labels: Object.keys(sch), values: Object.values(sch).map(a => a[6]) },
    ], Object.assign({}, chartBase, {
      x: 0.5, y: 1.95, w: 7.6, h: 4.5, barDir: "bar", chartColors: [K.grey, K.teal], barGapWidthPct: 60, barOverlapPct: -10,
      catAxisOrientation: "maxMin", valAxisHidden: true, valGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 320,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelColor: K.ink, dataLabelFormatCode: '0"人"',
      catAxisLabelFontSize: 18, catAxisLabelColor: K.ink, showLegend: false,
    }));
    [["令和8年度（2026）", K.grey], ["令和14年度（2032）", K.teal]].forEach(([t, c], i) => {
      s.addShape(pres.shapes.RECTANGLE, { x: 1.9 + i * 3.0, y: 1.62, w: 0.3, h: 0.25, fill: { color: c }, line: { type: "none" } });
      T(s, t, { x: 2.3 + i * 3.0, y: 1.57, w: 2.6, h: 0.35, fontSize: 16, color: K.ink, valign: "middle" });
    });
    const rows = Object.entries(sch);
    rows.forEach(([n, a], i) => {
      const y = 1.75 + i * 1.55;
      box(s, 8.5, y, 4.2, 1.35, K.bg);
      T(s, n, { x: 8.8, y: y + 0.18, w: 1.6, h: 0.5, fontSize: 19, bold: true, color: K.navy, valign: "middle" });
      T(s, `−${dec(a[0], a[6])}%`, { x: 10.4, y: y + 0.12, w: 2.1, h: 0.6, fontSize: 30, bold: true, color: K.teal, align: "right", valign: "middle" });
      T(s, `${a[0]}人 → ${a[6]}人（${a[0] - a[6]}人減）`, { x: 8.8, y: y + 0.8, w: 3.7, h: 0.4, fontSize: 16, color: K.sub });
    });
    note(s, "出典：提供資料「児童数推計」（令和8年（2026年）5月1日現在）。将来推計は変動します。");
  }

  // 5 合計
  {
    const s = content("3小学校の児童数合計は約32％減少", "令和8年度（2026年度）→ 令和14年度（2032年度）");
    T(s, `${tot[0]}`, { x: 0.6, y: 1.8, w: 2.6, h: 1.2, fontSize: 66, bold: true, color: K.grey, align: "right" });
    T(s, "人", { x: 3.25, y: 2.35, w: 0.5, h: 0.6, fontSize: 22, color: K.grey });
    T(s, "令和8年度（2026）", { x: 0.6, y: 3.0, w: 3.1, h: 0.4, fontSize: 15, color: K.sub, align: "right" });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 3.95, y: 2.15, w: 1.0, h: 0.6, fill: { color: K.line }, line: { type: "none" } });
    T(s, `${tot[6]}`, { x: 5.1, y: 1.8, w: 2.6, h: 1.2, fontSize: 66, bold: true, color: K.teal });
    T(s, "人", { x: 6.95, y: 2.35, w: 0.5, h: 0.6, fontSize: 22, color: K.teal });
    T(s, "令和14年度（2032）", { x: 5.1, y: 3.0, w: 3.1, h: 0.4, fontSize: 15, color: K.sub });
    s.addChart(pres.charts.BAR, [{ name: "3小学校の合計", labels: Y7.map(YL), values: tot }], Object.assign({}, chartBase, {
      x: 8.3, y: 1.55, w: 4.5, h: 2.9, barDir: "col", chartColors: [K.teal], barGapWidthPct: 40, valAxisHidden: true, valGridLine: { style: "none" }, valAxisMinVal: 0,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: K.ink, catAxisLabelFontSize: 12, showLegend: false,
      showTitle: true, title: "3小学校の合計（人）", titleFontSize: 13, titleColor: K.ink, titleFontFace: "+mn-lt",
    }));
    box(s, 0.6, 4.75, 12.1, 1.6, K.bg);
    T(s, "学校ごとに異なる減少の姿", { x: 0.95, y: 4.95, w: 11.5, h: 0.45, fontSize: 20, bold: true, color: K.navy });
    Object.entries(sch).forEach(([n, a], i) => {
      const x = 0.95 + i * 3.9;
      T(s, n, { x, y: 5.45, w: 1.0, h: 0.45, fontSize: 18, bold: true, color: K.navy, valign: "middle" });
      T(s, `${a[0] - a[6]}人減（−${dec(a[0], a[6])}%）`, { x: x + 1.0, y: 5.45, w: 2.8, h: 0.45, fontSize: 17, valign: "middle" });
    });
    const g1 = Y7.map((_, i) => ["綿内小", "川田小", "保科小"].reduce((t, n) => t + grades[n]["1年"][i], 0));
    T(s, `3校の新1年生の合計は ${g1[0]}人（R8＝2026年度）→ ${g1[6]}人（R14＝2032年度）。`, { x: 0.95, y: 5.95, w: 11.5, h: 0.35, fontSize: 16, color: K.sub });
    note(s, "数値は提供推計（令和8年（2026年）5月1日現在）の合計から算出。");
  }

  // 6 若穂中
  {
    const s = content("若穂中学校も生徒数が減少する見通し", "令和8年度（2026年度）→ 令和20年度（2038年度）");
    s.addChart(pres.charts.BAR, [{ name: "若穂中 全校生徒数", labels: Y13.map(YL), values: jh }], Object.assign({}, chartBase, {
      x: 0.5, y: 1.55, w: 8.0, h: 4.85, barDir: "col", chartColors: [K.teal], barGapWidthPct: 35, valAxisHidden: true, valGridLine: { style: "none" }, valAxisMinVal: 0,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: K.ink, catAxisLabelFontSize: 13, showLegend: false,
    }));
    box(s, 8.9, 1.65, 3.8, 2.1, K.tealL);
    T(s, `約${dec(jh[0], jh[12])}%減`, { x: 9.2, y: 1.8, w: 3.3, h: 1.0, fontSize: 44, bold: true, color: K.teal });
    T(s, `${jh[0]}人 → ${jh[12]}人`, { x: 9.2, y: 2.85, w: 3.3, h: 0.5, fontSize: 18, color: K.navy, bold: true });
    box(s, 8.9, 4.0, 3.8, 2.35, K.bg);
    T(s, "若穂中は、3つの中学校が一つになってできた学校", { x: 9.2, y: 4.15, w: 3.3, h: 0.85, fontSize: 17, bold: true, color: K.navy });
    T(s, "昭和36年（1961年）に統合し、翌37年（1962年）に今の場所で開校（当時の全校生徒は982人）。部活動など、中学校の環境も合わせて考えます。", { x: 9.2, y: 5.05, w: 3.3, h: 1.15, fontSize: 15, color: K.sub });
    note(s, "出典：提供資料「生徒数推計」（令和8年（2026年）5月1日現在）。成り立ち：長野市立若穂中学校ウェブサイト「学校紹介」の沿革（令和8年（2026年）11月確認）。");
    s.addNotes("若穂中の統合（昭和36年統合・37年開校）を経験した世代の方に、当時のよかった点・苦労した点を聞く機会をつくることも提案できる。");
  }

  // 7 市全体と比較
  {
    const s = content("長野市全体と比べても、若穂の減少は大きい", "小学校児童数：令和8年度（2026年度）を100としたときの令和14年度（2032年度）");
    const rows = [["若穂地区 3小学校", tot[0], tot[6], K.teal], ["長野市立小学校 全体", city[0], city[1], K.grey]];
    rows.forEach(([n, a, b, c], i) => {
      const y = 1.9 + i * 1.75, W = 7.2, r = b / a;
      T(s, n, { x: 0.6, y, w: 4, h: 0.45, fontSize: 19, bold: true, color: K.navy });
      T(s, `${a.toLocaleString()}人 → ${b.toLocaleString()}人`, { x: 4.6, y, w: 3.2, h: 0.45, fontSize: 15, color: K.sub, align: "right" });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y + 0.6, w: W, h: 0.6, fill: { color: "EEF1F2" }, line: { type: "none" } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y + 0.6, w: W * r, h: 0.6, fill: { color: c }, line: { type: "none" } });
      T(s, `100 → ${Math.round(r * 100)}`, { x: 0.75, y: y + 0.6, w: 2.5, h: 0.6, fontSize: 18, bold: true, color: K.white, valign: "middle" });
      T(s, `約${dec(a, b)}%減`, { x: 8.1, y: y + 0.35, w: 2.6, h: 0.9, fontSize: 36, bold: true, color: i ? K.sub : K.teal, valign: "middle" });
    });
    box(s, 0.6, 5.55, 12.1, 0.85, K.tealL);
    T(s, "地域固有の課題として、将来を見据えた検討が必要です。", { x: 0.95, y: 5.55, w: 11.5, h: 0.85, fontSize: 20, bold: true, color: K.navy, valign: "middle" });
    note(s, "出典：提供推計データ（令和8年（2026年）5月1日現在）。減少率は四捨五入。", 6.6);
  }

  divider("02", "課題を共有する", "学校の良さと課題、施設の状況"); SEC = "02  課題を共有する";
  // 8 良さと課題
  {
    const s = content("小規模校の良さと、考えておきたい課題");
    const cols = [["大切にしたい良さ", K.teal, K.tealL, "MdThumbUp", ["一人ひとりへの目配り", "学年間の交流", "地域と学校の近い関係", "安心できる人間関係"]],
      ["将来の検討課題", K.amber, K.amberL, "MdHelpOutline", ["多様な意見に触れる機会", "集団活動や行事の選択肢", "人間関係の固定化への配慮", "教員配置や教育活動の持続性"]]];
    for (let i = 0; i < 2; i++) {
      const [h, c, cl, ic, items] = cols[i], x = 0.6 + i * 6.15;
      box(s, x, 1.55, 5.95, 4.55, cl);
      await iconCircle(s, x + 0.4, 1.85, 0.75, ic, c);
      T(s, h, { x: x + 1.35, y: 1.85, w: 4.3, h: 0.75, fontSize: 24, bold: true, color: K.navy, valign: "middle" });
      items.forEach((t, j) => {
        const y = 2.95 + j * 0.75;
        s.addShape(pres.shapes.OVAL, { x: x + 0.55, y: y + 0.17, w: 0.18, h: 0.18, fill: { color: c }, line: { type: "none" } });
        T(s, t, { x: x + 0.95, y, w: 4.8, h: 0.52, fontSize: 19, valign: "middle" });
      });
    }
    T(s, "学校規模だけで教育の良し悪しが決まるものではありません。", { x: 0.6, y: 6.2, w: 12.1, h: 0.35, fontSize: 16, bold: true, color: K.sub, align: "center" });
    note(s, "各校の具体的な良さ・取り組みは、学校と地域から伺って追加します［聞き取り中］。", 6.62);
  }

  // 追加C 望ましい学校規模
  {
    const s = content("望ましい学校規模の目安と、若穂の見通し", "市の審議のまとめ（平成30年＝2018年）で示された目安に、推計を当てはめると");
    const crit = [["国の標準", "12〜18学級", "小学校・中学校とも（学校教育法施行規則）"], ["長野県の目安", "学年に複数の学級\n学年20人程度", "クラス替えができ、一定の集団で学べる規模"]];
    crit.forEach(([h, v, d], i) => {
      const y = 1.65 + i * 2.15;
      box(s, 0.6, y, 4.9, 2.0, K.tealL);
      T(s, h, { x: 0.9, y: y + 0.18, w: 4.3, h: 0.4, fontSize: 17, bold: true, color: K.teal });
      T(s, v, { x: 0.9, y: y + 0.55, w: 4.3, h: 0.95, fontSize: 22, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: 0.9, y: y + 1.55, w: 4.3, h: 0.32, fontSize: 14, color: K.sub });
    });
    const H = t => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, align: "center" } });
    const under = (n, i) => Object.entries(grades[n]).filter(([g]) => g !== "合計").filter(([, v]) => v[i] < 20).length;
    const nG = n => Object.keys(grades[n]).length - 1;
    const rows = [[H("学校"), H("R8（2026）"), H("R14（2032）")]];
    ["綿内小", "川田小", "保科小", "若穂中"].forEach((n, i) => {
      const o = { fill: { color: i % 2 ? K.white : K.bg }, align: "center" };
      rows.push([{ text: n, options: Object.assign({}, o, { bold: true, align: "left" }) }, ...[0, 6].map(j => ({ text: `${under(n, j)} / ${nG(n)}学年`, options: o }))]);
    });
    T(s, "学年の人数が20人未満になる学年の数", { x: 5.9, y: 1.7, w: 6.8, h: 0.45, fontSize: 18, bold: true, color: K.navy });
    s.addTable(rows, { x: 5.9, y: 2.25, w: 6.8, colW: [2.0, 2.4, 2.4], rowH: [0.6, 0.62, 0.62, 0.62, 0.62], fontSize: 17, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    T(s, "学級の数は、学級の人数の基準によって変わります。各校の学級数の見通しは［事務局に確認中］。", { x: 5.9, y: 5.45, w: 6.8, h: 0.6, fontSize: 15, color: K.sub });
    box(s, 0.6, 6.0, 12.1, 0.5, K.bg);
    T(s, "目安を下回ることが、そのまま再編の理由にはなりません。施設の状況や各校の良さなども合わせて比べます。", { x: 0.8, y: 6.0, w: 11.7, h: 0.5, fontSize: 15, bold: true, color: K.navy, valign: "middle" });
    note(s, "出典：長野市「少子化に対応した子どもにとって望ましい教育環境の在り方について（審議のまとめ）」（平成30年＝2018年）、推計（令和8年（2026年）5月1日現在）", 6.62);
    s.addNotes("目安を満たさないことが、ただちに再編の理由になるわけではないことを口頭で補足する。");
  }

  // 9 なぜ今
  {
    const s = content("なぜ今、検討するのか");
    const rs = [["児童生徒数の減少は、今後も続く見通し", "小学生はR14（2032年度）までに約3割、中学生はR20（2038年度）までに5割以上減る推計", "MdTrendingDown"],
      ["施設の更新・改修を考える時期と重なる", "4校すべてが2030年度（R12）までに、計画上の更新・改修の対象に入る（中身と実施状況は市に確認中）", "MdApartment"],
      ["通学や地域活動への影響を調べるには時間が必要", "調査・比較・地区ごとの対話に、2年ほどかけて丁寧に", "MdSchedule"],
      ["今の子どもと将来の子ども、両方の視点が必要", "在校生への配慮と、これから入学する子の環境の両立", "MdFamilyRestroom"]];
    for (let i = 0; i < 4; i++) {
      const [h, d, ic] = rs[i], x = 0.6 + (i % 2) * 6.15, y = 1.55 + Math.floor(i / 2) * 2.5;
      box(s, x, y, 5.95, 2.3, K.bg);
      await iconCircle(s, x + 0.35, y + 0.35, 0.9, ic);
      T(s, h, { x: x + 1.5, y: y + 0.32, w: 4.2, h: 1.0, fontSize: 20, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: x + 1.5, y: y + 1.35, w: 4.2, h: 0.8, fontSize: 15, color: K.sub });
    }
  }

  // 10 施設 timeline
  {
    const s = content("施設の状況も、判断材料のひとつ", "建築年度と、計画上で更新・改修の対象に入る時期（長寿命化計画＝建物を計画的に改修して長く使うための市の計画）");
    const x0 = 2.6, x1 = 12.4, yr0 = 1970, yr1 = 2035, X = y => x0 + (y - yr0) / (yr1 - yr0) * (x1 - x0);
    [1970, 1980, 1990, 2000, 2010, 2020, 2030].forEach(yr => {
      s.addShape(pres.shapes.LINE, { x: X(yr), y: 1.85, w: 0, h: 4.1, line: { color: "E7ECEE", width: 1 } });
      T(s, String(yr), { x: X(yr) - 0.4, y: 1.58, w: 0.8, h: 0.25, fontSize: 13, color: K.grey, align: "center" });
    });
    s.addShape(pres.shapes.LINE, { x: X(2026), y: 1.85, w: 0, h: 3.95, line: { color: K.amber, width: 1.5, dashType: "dash" } });
    T(s, "現在", { x: X(2026) - 0.4, y: 5.8, w: 0.8, h: 0.25, fontSize: 13, color: K.amber, bold: true, align: "center" });
    // [学校, 最も古い棟, 最も新しい棟, 棟数, 計画上の年度, 内容, 計画で使われた棟の建築年度]（本文の施設一覧と添付資料。docs/提供情報まとめ.md の7）
    const rows = [["綿内小", 1972, 2000, 4, 2030, "更新", 1980], ["川田小", 1973, 1994, 5, 2023, "更新", 1973], ["保科小", 1977, 2010, 6, 2030, "築20年の改修", 2010], ["若穂中", 1987, 1997, 6, 2027, "築40年の改修", 1987]];
    rows.forEach(([n, b0, b1, cnt, p, k, bb], i) => {
      const y = 2.15 + i * 0.95;
      T(s, n, { x: 0.6, y, w: 1.9, h: 0.5, fontSize: 18, bold: true, color: K.navy, valign: "middle" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X(b0), y: y + 0.1, w: X(2026) - X(b0), h: 0.3, fill: { color: K.hill3 }, line: { type: "none" }, rectRadius: 0.15 });
      T(s, `${b0}〜${b1}年度（${WA(b0)}〜${WA(b1)}年度）に建てた${cnt}棟`, { x: X(b0) + 0.1, y: y + 0.1, w: 7.0, h: 0.3, fontSize: 13, bold: true, color: K.white, valign: "middle" });
      s.addShape(pres.shapes.DIAMOND, { x: X(p) - 0.14, y: y + 0.11, w: 0.28, h: 0.28, fill: { color: K.amber }, line: { color: K.white, width: 1 } });
      T(s, `${p}年度（R${p - 2018}）から${k}の対象（${WA(bb)}＝${bb}年度の棟で計算）`, { x: Math.min(X(p) + 0.15, X(2026) - 0.1) - 8.6, y: y + 0.45, w: 8.6, h: 0.3, fontSize: 13, color: "A3402C", bold: true, align: "right" });
    });
    note(s, "◆計画期間（R3〜R12＝2021〜2030年度）中に対象に入る時期の目安で、工事の予定ではありません（市の資料、R2＝2020年4月時点）。\n各校1つの棟で計算（旧耐震は築50年で「更新」、新耐震は築20・40年で「改修」）。「更新」が建て替えか改修かは［市に確認中］。\n4校とも全棟が長寿命化に適する判定。R8（2026）年度、綿内小・川田小の北校舎と保科小の南校舎のトイレ改修を入札公告。", 6.1);
    s.addNotes("年度は、計画の期間中に対象に入る目安を機械的に書いたもので、工事の予定ではない。旧耐震か新耐震かで、市内のどの学校も同じ書き方になっている（学校ごとに判断したものではない）。保科小は2010年度の棟（体育館の年度と同じ）で計算されているが、校舎は1977〜1999年度。どの棟で計算したかは市に確認中。出典：長野市学校施設長寿命化計画（令和3年（2021年）2月策定・令和5年（2023年）6月改訂）本文の施設一覧と添付資料。");
  }

  // 追加G 市の計画に書かれていること
  {
    const s = content("長野市の計画に書かれていること", "長野市学校施設長寿命化計画（R3＝2021年策定・R5＝2023年改訂）から。若穂について決めたものではありません");
    const cols = [["建物の使い方", "MdApartment", [["建設から40年後に長寿命化改修", "その後40年程度使えるようにする（建て替えより費用を抑える）"], ["改修に適さない建物は建て替え", "改修前に建物の強さを調べて判断する"]]],
      ["市全体の学校の費用", "MdPayments", [["R8（2026）年度以降の事業費は、年平均58億円の見込み", "R5〜R7（2023〜2025）年度の年30.4億円との差は、年約28億円"], ["「施設総量について今後20年間で20％縮減」", "学校に限らない、市の公共施設全体の目標（総合管理計画）"]]],
      ["学校の配置の考え方", "MdGroups", [["「統廃合も見据えて適正配置について検討」", "各学校区の人口の動きを見極めながら"], ["「地域に学校を残したいという地域の思いも受け止め」", "例：小学校の低・中学年は地域で、高学年以上は大きな学校で学ぶ形の検討"]]]];
    for (let i = 0; i < 3; i++) {
      const [h, ic, items] = cols[i], x = 0.6 + i * 4.1;
      box(s, x, 1.6, 3.9, 4.6, K.bg);
      await iconCircle(s, x + 0.3, 1.8, 0.7, ic);
      T(s, h, { x: x + 1.15, y: 1.8, w: 2.65, h: 0.7, fontSize: 20, bold: true, color: K.navy, valign: "middle" });
      items.forEach(([a, b], j) => {
        const y = 2.75 + j * 1.7;
        T(s, a, { x: x + 0.3, y, w: 3.4, h: 0.75, fontSize: 17, bold: true, color: K.teal });
        T(s, b, { x: x + 0.3, y: y + 0.78, w: 3.4, h: 0.8, fontSize: 15, color: K.ink });
      });
    }
    T(s, "市の方針も一つの前提として知ったうえで、若穂の子どもにとって何がよいかを、地域で考えます。", { x: 0.6, y: 6.3, w: 12.1, h: 0.35, fontSize: 16, bold: true, color: K.sub, align: "center" });
    note(s, "「 」は計画の文章の引用。出典：長野市学校施設長寿命化計画（R3＝2021年策定・R5＝2023年改訂）第4〜6章。", 6.7);
    s.addNotes("市全体の計画であり、若穂の学校をどうするかを決めたものではないことを、必ず口頭で伝える。施設総量20％縮減は学校に限らない公共施設全体の目標。");
  }

  divider("03", "選択肢を比べる", "大切にしたい視点と、考えられる選択肢"); SEC = "03  選択肢を比べる";
  // 11 5つの視点
  {
    const s = content("若穂の学校で大切にしたい5つの視点");
    const v = [["学び", "多様な学習・交流機会", "MdMenuBook"], ["安心", "一人ひとりへの支援", "MdFavorite"], ["安全", "無理のない通学と防災", "MdHealthAndSafety"], ["地域", "行事・交流・つながり", "MdGroups"], ["持続性", "施設・人材・費用", "MdAutorenew"]];
    for (let i = 0; i < 5; i++) {
      const [h, d, ic] = v[i], x = 0.6 + i * 2.46;
      box(s, x, 1.7, 2.26, 4.4, i % 2 ? K.bg : K.tealL);
      await iconCircle(s, x + 0.53, 2.1, 1.2, ic);
      T(s, h, { x: x + 0.1, y: 3.6, w: 2.06, h: 0.6, fontSize: 26, bold: true, color: K.navy, align: "center" });
      T(s, d, { x: x + 0.2, y: 4.35, w: 1.86, h: 1.2, fontSize: 16, color: K.sub, align: "center" });
    }
  }

  // 12 選択肢（学校の配置 A〜C と、学校の制度 D・E を分けて示す）
  {
    const s = content("考えられる選択肢を公平に比較する");
    const o = [["A", "現状維持", "3小学校・中学校を今の形で維持", "小規模化への対応（複式学級の可能性など）"], ["B", "一部統合", "2校の統合など。組み合わせと統合先を比較", "地域間の公平性・通学"], ["C", "小学校3校統合", "小学校を1校に集約", "校地・規模・通学"],
      ["D", "義務教育学校", "小中9年間を一つの学校で。校長は1人", "教育課程・教員配置・施設・通学"], ["E", "小中一貫型の\n小学校・中学校", "小学校と中学校は別々の学校のまま、9年間の教育課程をそろえる", "教員の連携・行事・施設"]];
    const W = 2.3, G = 0.15;
    [["学校の配置", 0, 3], ["学校の制度（区分）", 3, 2]].forEach(([h, k, n]) => {
      const x = 0.6 + k * (W + G), w = n * W + (n - 1) * G;
      T(s, h, { x, y: 1.45, w, h: 0.35, fontSize: 16, bold: true, color: K.teal, align: "center" });
      s.addShape(pres.shapes.LINE, { x, y: 1.83, w, h: 0, line: { color: K.teal, width: 2 } });
    });
    o.forEach(([L, h, d, p], i) => {
      const x = 0.6 + i * (W + G);
      box(s, x, 1.95, W, 4.3, K.bg);
      s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 2.15, w: 0.7, h: 0.7, fill: { color: K.navy }, line: { type: "none" } });
      T(s, L, { x: x + 0.2, y: 2.15, w: 0.7, h: 0.7, fontSize: 24, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, h, { x: x + 0.2, y: 2.95, w: W - 0.35, h: 0.75, fontSize: 19, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: x + 0.2, y: 3.75, w: W - 0.35, h: 1.3, fontSize: 16 });
      s.addShape(pres.shapes.LINE, { x: x + 0.2, y: 5.1, w: W - 0.4, h: 0, line: { color: K.line, width: 1 } });
      T(s, "主な論点", { x: x + 0.2, y: 5.18, w: W - 0.35, h: 0.3, fontSize: 12, color: K.teal, bold: true });
      T(s, p, { x: x + 0.2, y: 5.48, w: W - 0.35, h: 0.7, fontSize: 14, color: K.sub });
    });
    note(s, "配置（A〜C）と制度（D・E）は別の軸で、組み合わせられます（例：A＋E＝3校と若穂中はそのままで、9年間の教育課程をそろえる）。\nD（義務教育学校）とE（小中一貫型の小学校・中学校）は、法律上の区分が違う別の制度です。", 6.38);
  }

  // 13 同じ物差し matrix
  {
    const s = content("選択肢は同じ物差しで比較する", "評価欄は、R9（2027年度）の調査結果をもとに記入し、R9（2027年度）後半の中間報告でお示しします");
    const H = t => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, align: "center" } });
    const crit = [["教育", "学習環境・交流・教員体制・特別支援"], ["通学", "距離・時間・交通安全・放課後"], ["施設", "安全性・整備費・維持管理費"], ["地域", "防災・行事・跡地利用・児童センター"], ["実現性", "準備期間・財源・合意形成"]];
    const rows = [[{ text: "視点", options: { bold: true, color: K.white, fill: { color: K.navy } } }, H("比べる内容"), H("A 現状維持"), H("B 一部統合"), H("C 3校統合"), H("D 義務教育学校"), H("E 小中一貫型")]];
    crit.forEach(([a, b], i) => rows.push([{ text: a, options: { bold: true, color: K.teal, fontSize: 18 } }, { text: b, options: {} }, ...[0, 1, 2, 3, 4].map(() => ({ text: "", options: { fill: { color: i % 2 ? K.white : K.bg } } }))]));
    s.addTable(rows, { x: 0.6, y: 1.7, w: 12.1, colW: [1.25, 3.6, 1.45, 1.45, 1.45, 1.45, 1.45], rowH: 0.78, fontSize: 16, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    note(s, "評価の付け方（案）：事務局の調査結果をもとに、委員会が視点ごとに「良い・課題あり・調査中」で記入し、地区別対話で確認します。", 6.6);
  }

  // 14 義務教育学校と小中一貫型の小学校・中学校（別の区分として比べる）
  {
    const s = content("義務教育学校と小中一貫型の小学校・中学校の違い", "どちらも9年間を見通して教えますが、法律上の学校の区分が違う、別の制度です");
    const Hd = (t, c) => ({ text: t, options: { bold: true, color: K.white, fill: { color: c }, align: "center" } });
    const R = [["学校の区分", "小学校・中学校とは別の種類の学校", "小学校と中学校（それぞれ今の区分のまま）"],
      ["学校の数", "1校", "2校（小学校と中学校）"],
      ["校長・教職員", "校長1人、教職員の組織も1つ", "学校ごとに校長と教職員の組織"],
      ["学年", "1〜9年生（前期6年・後期3年）", "小学1〜6年生・中学1〜3年生"],
      ["教育課程", "9年間で1つ", "2校で9年間の教育課程をそろえる"],
      ["教員の免許", "原則、小学校と中学校の両方", "それぞれの学校の免許"]];
    const rows = [[Hd("", K.navy), Hd("D 義務教育学校", K.navy), Hd("E 小中一貫型の小学校・中学校", K.teal)]];
    R.forEach(([a, b, c], i) => rows.push([{ text: a, options: { bold: true, color: K.navy } }, { text: b, options: { fill: { color: i % 2 ? K.white : K.bg } } }, { text: c, options: { fill: { color: i % 2 ? K.white : K.bg } } }]));
    s.addTable(rows, { x: 0.6, y: 1.55, w: 12.1, colW: [2.1, 5.0, 5.0], rowH: 0.5, fontSize: 16, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    const cols = [["どちらにも期待される点", K.teal, K.tealL, "学びのつながり・中学校に進むときの不安を減らす・異学年の交流・小中の教員の連携"], ["どちらでも慎重に考える点", K.amber, K.amberL, "人間関係の固定化・年齢差の大きい子どもの生活・通学と施設"]];
    cols.forEach(([h, c, cl, d], i) => {
      const x = 0.6 + i * 6.15;
      box(s, x, 5.2, 5.95, 0.95, cl);
      T(s, h, { x: x + 0.25, y: 5.25, w: 5.5, h: 0.35, fontSize: 16, bold: true, color: c });
      T(s, d, { x: x + 0.25, y: 5.6, w: 5.5, h: 0.5, fontSize: 15, color: K.ink });
    });
    T(s, "校舎は、どちらの制度でも、一つにまとめる形も、離れたままの形もあります。", { x: 0.6, y: 6.22, w: 12.1, h: 0.35, fontSize: 16, bold: true, color: K.navy, align: "center" });
    note(s, "出典：学校教育法（平成28年＝2016年4月施行の改正）、文部科学省「小中一貫した教育課程の編成・実施に関する手引」（平成28年＝2016年12月）。", 6.62);
    s.addNotes("義務教育学校は小学校・中学校とは別の区分の学校。小中一貫型の小学校・中学校は、小学校と中学校がそれぞれの学校のまま教育課程をそろえるもの（同じ設置者の「併設型」、違う設置者の「連携型」がある）。両者を混同しないように説明する。義務教育学校の教員は原則として小中両方の免許が必要だが、当分の間はどちらか一方の免許で前期・後期を担当できる。");
  }

  // 追加A 学校間の距離
  {
    const s = content("学校と学校の距離・移動時間の目安", "各小学校の間、各小学校と若穂中の間（地図は位置関係を示す概略図）");
    box(s, 0.6, 1.6, 6.1, 5.0, K.bg);
    const P = { 綿内小: [36.622361, 138.262529], 川田小: [36.604091, 138.241565], 保科小: [36.5874875, 138.2655742], 若穂中: [36.604752, 138.251224] };
    const sc = 1.1, cx = 4.05, cy = 4.1, la0 = 36.6049, lo0 = 138.2536;
    const XY = n => [cx + (P[n][1] - lo0) * 89.4 * sc, cy - (P[n][0] - la0) * 111.0 * sc];
    const legs = [["綿内小", "川田小", 1], ["川田小", "保科小", 1], ["綿内小", "保科小", 1], ["綿内小", "若穂中", 0], ["川田小", "若穂中", 0], ["保科小", "若穂中", 0]];
    const R = 0.24;
    legs.forEach(([a, b, el], i) => {
      const [x1, y1] = XY(a), [x2, y2] = XY(b), L = Math.hypot(x2 - x1, y2 - y1), ux = (x2 - x1) / L, uy = (y2 - y1) / L;
      const sx = x1 + ux * R, sy = y1 + uy * R, ex = x2 - ux * R, ey = y2 - uy * R;
      s.addShape(pres.shapes.LINE, { x: Math.min(sx, ex), y: Math.min(sy, ey), w: Math.abs(ex - sx), h: Math.abs(ey - sy), flipH: (ex < sx) !== (ey < sy) ? true : false,
        line: { color: el ? K.teal : K.amber, width: 2.5, dashType: el ? "solid" : "dash", beginArrowType: "triangle", endArrowType: "triangle" } });
      const mx = (x1 + x2) / 2, my = (y1 + y2) / 2;
      s.addShape(pres.shapes.OVAL, { x: mx - 0.17, y: my - 0.17, w: 0.34, h: 0.34, fill: { color: K.white }, line: { color: el ? K.teal : K.amber, width: 1.5 } });
      T(s, String(i + 1), { x: mx - 0.17, y: my - 0.17, w: 0.34, h: 0.34, fontSize: 13, bold: true, color: el ? K.teal : K.amber, align: "center", valign: "middle" });
    });
    Object.keys(P).forEach(n => {
      const [x, y] = XY(n), jh = n === "若穂中";
      s.addShape(jh ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.OVAL, { x: x - 0.2, y: y - 0.2, w: 0.4, h: 0.4, fill: { color: jh ? K.navy : K.teal }, line: { color: K.white, width: 2 }, rectRadius: 0.06 });
      const left = n === "川田小";
      T(s, n, { x: left ? x - 1.45 : x + 0.28, y: y - 0.2, w: 1.15, h: 0.4, fontSize: 16, bold: true, color: K.navy, valign: "middle", align: left ? "right" : "left" });
    });
    s.addShape(pres.shapes.LINE, { x: 0.9, y: 6.3, w: 1.1, h: 0, line: { color: K.sub, width: 2 } });
    T(s, "1km", { x: 2.08, y: 6.15, w: 0.6, h: 0.3, fontSize: 13, color: K.sub });
    T(s, "↑ 北", { x: 2.9, y: 6.15, w: 0.7, h: 0.3, fontSize: 13, bold: true, color: K.sub });
    s.addShape(pres.shapes.LINE, { x: 0.9, y: 1.9, w: 0.45, h: 0, line: { color: K.teal, width: 2.5 } });
    T(s, "小学校どうし", { x: 1.42, y: 1.75, w: 1.8, h: 0.3, fontSize: 13, color: K.sub });
    s.addShape(pres.shapes.LINE, { x: 0.9, y: 2.2, w: 0.45, h: 0, line: { color: K.amber, width: 2.5, dashType: "dash" } });
    T(s, "小学校と若穂中", { x: 1.42, y: 2.05, w: 1.8, h: 0.3, fontSize: 13, color: K.sub });
    const H = t => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, align: "center" } });
    T(s, "保科小は、ほかの3校より\n約100m高い所にあります", { x: 0.9, y: 2.45, w: 3.0, h: 0.65, fontSize: 14, color: K.sub });
    // 道のり・時間：Google マップの経路検索（2026-10-27、input/maps/）。[区間, 徒歩km, 徒歩分, 車km, 車分]
    const legRows = [["綿内小 ⇔ 川田小", "3.2", "44分", "3.5", "6分"], ["川田小 ⇔ 保科小", "3.0", "46分", "3.2", "5分"], ["綿内小 ⇔ 保科小", "4.9", "72分", "5.6", "8分"],
      ["綿内小 ⇔ 若穂中", "2.5", "33分", "2.9", "5分"], ["川田小 ⇔ 若穂中", "1.1", "16分", "1.2", "3分"], ["保科小 ⇔ 若穂中", "2.5", "39分", "2.5", "4分"]];
    const rows = [[Object.assign(H("No"), { options: Object.assign(H("").options, { rowspan: 2 }) }), Object.assign(H("区間"), { options: Object.assign(H("").options, { rowspan: 2 }) }),
      Object.assign(H("徒歩"), { options: Object.assign(H("").options, { colspan: 2 }) }), Object.assign(H("車"), { options: Object.assign(H("").options, { colspan: 2 }) })],
      [H("道のり"), H("時間"), H("道のり"), H("時間")]];
    legRows.forEach((r, i) => rows.push([String(i + 1), r[0], `${r[1]}km`, r[2], `${r[3]}km`, r[4]].map((t, j) => ({ text: t, options: { align: j === 1 ? "left" : "center", bold: j < 2 || j === 3, color: j === 0 ? (i < 3 ? K.teal : K.amber) : K.ink, fill: { color: i % 2 ? K.white : K.bg } } }))));
    s.addTable(rows, { x: 6.9, y: 1.6, w: 5.8, colW: [0.5, 1.8, 0.9, 0.8, 0.9, 0.9], rowH: [0.38, 0.38, 0.46, 0.46, 0.46, 0.46, 0.46, 0.46], fontSize: 14, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    box(s, 6.9, 5.25, 5.8, 0.68, K.tealL);
    T(s, "国の目安：通学距離は小学校おおむね4km以内・中学校おおむね6km以内、通学時間はおおむね1時間以内", { x: 7.1, y: 5.25, w: 5.4, h: 0.68, fontSize: 13, bold: true, color: K.navy, valign: "middle" });
    box(s, 6.9, 6.0, 5.8, 0.62, K.amberL);
    T(s, "学校と学校の距離で、子どもの通学距離ではありません。徒歩は大人の速さの目安で、子どもはもっとかかります。", { x: 7.1, y: 6.0, w: 5.4, h: 0.62, fontSize: 12, color: "A3402C", valign: "middle" });
    note(s, "Google マップの経路検索（令和8年（2026年）10月27日夜）。2・3・6は保科小へ向かう上りの時間。国の目安：文部科学省の手引（平成27年＝2015年）。", 6.62);
    s.addNotes("学校間の距離であり、各家庭からの通学距離ではないことを必ず口頭で伝える。道のりと時間はGoogle マップの経路検索（2026年10月27日の夜）。保科小が入る区間（2・3・6）は、保科小へ向かう上りの向きの時間（川田小→保科小46分、逆の下りは36分）。徒歩は大人の速さ、車は夜の交通量での値で、朝の通学時間帯は長くなることがある。保科小へ向かう徒歩は約100mの上り（若穂中→保科小 上り102m、綿内小→保科小 上り113m、保科小→川田小 下り104m）。各家庭からの距離と子どもの歩く速さでの実測は、事務局に依頼する。");
  }

  // 追加C 地区の地図と子どもが住む場所
  {
    const s = content("若穂の地区と、子どもが住んでいる場所", "地区は市の区分（綿内・川田・保科）。地図の線は国勢調査の区域（大字）で、川田地区は牛島を含みます");
    s.addImage({ path: path.join(__dirname, "../data/map_wakaho.png"), x: 0.6, y: 1.5, h: 4.7, w: 4.7 * 1364 / 1510 });
    const pc = c => fmt(MACHI.filter(m => m.c === c).reduce((a, m) => a + m.p, 0));
    const hd = t => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, align: "center", fontSize: 14 } });
    const rows = [[hd("地区"), hd("人口"), hd("小学校の児童"), hd("0〜14歳（参考）")],
      ...[["綿内", "綿内小", 273], ["川田", "川田小", 157], ["保科", "保科小", 54]].map(([c, sc, n], i) => [
        { text: `${c}地区`, options: { bold: true, color: K.navy } }, { text: `${pc(c)}人`, options: { align: "right" } },
        { text: `${sc} ${n}人`, options: { align: "right" } }, { text: `${fmt(TG[c + ">" + sc].kids)}人`, options: { align: "right" } }]),
      [{ text: "若穂全体", options: { bold: true, color: K.navy } }, { text: `${fmt(MACHI.reduce((a, m) => a + m.p, 0))}人`, options: { align: "right" } },
        { text: "若穂中 284人", options: { align: "right" } }, { text: `${fmt(TG["全体>若穂中"].kids)}人`, options: { align: "right" } }]];
    rows.forEach((r, i) => { if (i > 0) r.forEach(c => { c.options.fill = { color: i % 2 ? K.bg : K.white }; }); });
    s.addTable(rows, { x: 5.4, y: 1.5, w: 7.3, colW: [1.8, 1.5, 1.9, 2.1], rowH: [0.5, 0.55, 0.55, 0.55, 0.55], fontSize: 16, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    T(s, "時点：人口 R8（2026）年10月1日、児童・生徒 R8（2026）年5月1日。\n0〜14歳（参考）はR2（2020）年の国勢調査の数で、中学生も含みます。", { x: 5.4, y: 4.35, w: 7.3, h: 0.55, fontSize: 13, color: K.sub });
    box(s, 5.4, 4.98, 7.3, 0.85, K.tealL);
    T(s, "緑のマス（250m四方）は0〜14歳の子どもが住む場所です（濃いほど多い）。どの地区でも、学校の近くの平地に多く住んでいます。", { x: 5.6, y: 4.98, w: 6.95, h: 0.85, fontSize: 15, color: K.navy, valign: "middle" });
    T(s, "保科地区は若穂団地・保科温泉を含みます。大きな地図は別紙（A3）。", { x: 5.4, y: 5.88, w: 7.3, h: 0.3, fontSize: 14, color: K.sub });
    note(s, "出典：人口＝長野市「町別人口及び世帯数」（令和8年（2026年）10月1日）。児童・生徒＝提供資料（令和8年（2026年）5月1日現在）。0〜14歳＝国勢調査（令和2年＝2020年）250mメッシュ。地区の線＝国勢調査の小地域（大字）。\n地区の線は、通学区域（学区）と一致しない場合があります。実際の通学区域は［市に確認中］。", 6.25);
    s.addNotes("0〜14歳の人数は2020年の国勢調査で、今の小学生の人数ではない。メッシュの中心がどの大字に入るかで地区に振り分けた目安。国土数値情報の小学校区データでは若穂団地の付近が川田小学区に入るが、市の区分では保科地区のため、資料は市の地区の区分に合わせている。小学校の児童数は市教委の推計（令和8年（2026年）5月1日現在）。地区ごとの町（区）の人口は参考資料に載せている。");
  }

  // 追加D 家から学校までの道のり（推計）
  {
    const s = content("家から学校までの道のり（推計）", "子どもが住んでいる場所から、今の学校までの徒歩の道のり");
    const R = [["綿内地区 → 綿内小", "綿内>綿内小"], ["川田地区 → 川田小", "川田>川田小"], ["保科地区 → 保科小", "保科>保科小"], ["若穂全体 → 若穂中", "全体>若穂中"]];
    const L = [["1km未満", ["1km未満"], "D3E8DC", K.ink], ["1〜2km", ["1〜2km"], "9CCBB0", K.ink], ["2〜3km", ["2〜3km"], "55A07D", K.white], ["3〜4km", ["3〜4km"], "2A6E52", K.white], ["4km以上", ["4〜6km", "6km以上"], "143D2D", K.white]];
    const x0 = 3.85, W = 8.85;
    L.forEach(([lab, , c], k) => { s.addShape(pres.shapes.RECTANGLE, { x: x0 + k * 1.55, y: 1.62, w: 0.28, h: 0.28, fill: { color: c }, line: { color: K.line, width: 0.5 } }); T(s, lab, { x: x0 + k * 1.55 + 0.36, y: 1.58, w: 1.15, h: 0.36, fontSize: 15, color: K.ink, valign: "middle" }); });
    R.forEach(([lab, key], i) => {
      const y = 2.15 + i * 0.82, d = TG[key];
      T(s, lab, { x: 0.6, y, w: 3.2, h: 0.6, fontSize: 17, bold: true, color: K.navy, valign: "middle" });
      let x = x0;
      L.forEach(([, bs, c, tc]) => {
        const v = bs.reduce((a, b) => a + d.bins[b], 0), w = W * v / d.kids;
        if (w <= 0) return;
        s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.6, fill: { color: c }, line: { color: K.white, width: 1.5 } });
        if (v / d.kids >= 0.08) T(s, `${Math.round(v / d.kids * 100)}%`, { x, y, w, h: 0.6, fontSize: 15, bold: true, color: tc, align: "center", valign: "middle" });
        x += w;
      });
    });
    const pts = [`小学校：3地区とも、地区の子どもの9割以上が今の学校まで2km未満（綿内小${share("綿内>綿内小", 2)}%・川田小${share("川田>川田小", 2)}%・保科小${share("保科>保科小", 2)}%）`,
      `いちばん遠い場所：綿内小${km(TG["綿内>綿内小"].max_m)}km・川田小${km(TG["川田>川田小"].max_m)}km・保科小${km(TG["保科>保科小"].max_m)}km／若穂中${km(TG["全体>若穂中"].max_m)}km`,
      "国の手引の目安（小学校おおむね4km以内・中学校おおむね6km以内）は、今は全員が満たしています"];
    box(s, 0.6, 5.5, 12.1, 0.95, K.bg);
    pts.forEach((p, k) => T(s, "・" + p, { x: 0.85, y: 5.53 + k * 0.3, w: 11.7, h: 0.3, fontSize: 15, color: K.ink }));
    note(s, "250m四方の中心から学校までの徒歩の道のり（OpenStreetMapの道路データによる経路計算、令和8年（2026年）11月）。子どもは令和2年（2020年）の0〜14歳。各家庭からの正確な距離ではなく、坂・歩道・冬の道は含みません。", 6.55);
    s.addNotes("中学生も含む0〜14歳の住む場所で重みを付けている。徒歩の道のりは、学校間の値ではGoogleマップと2〜6%の差。ほかの学校までの道のりは参考資料の表にある（どの選択肢でも同じ物差しで比べるため）。子どもの歩く速さ・朝の時間帯の所要時間は、事務局に依頼中。");
  }

  // 追加B 避難所（市の一覧 令和7年（2025年）7月17日時点）
  {
    const s = content("小中学校は、すべて地域の避難所", "若穂地区の指定避難所は、この4校だけです（市の一覧）");
    const hd = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true, color: K.white, fill: { color: K.navy }, align: "center", fontSize: 15 }, o) });
    const rows = [[hd("学校", { rowspan: 2 }), hd("指定避難所", { rowspan: 2 }), hd("指定緊急避難場所（災害の種類ごと）", { colspan: 4 }), hd("備考", { rowspan: 2 })],
      [hd("洪水など"), hd("土砂災害"), hd("地震"), hd("大規模な火事")]];
    const data = [["綿内小学校", "○", "×", "○", "○", "×", ""], ["川田小学校", "○", "×", "○", "○", "×", ""],
      ["保科小学校", "○", "○\n2階以上", "○\n教室棟の建物内", "○", "×", "防災備蓄倉庫あり\n児童センター併設"], ["若穂中学校", "○", "○\n2階以上", "○", "○", "×", ""]];
    data.forEach((r, i) => rows.push(r.map((v, j) => ({ text: v, options: { align: j === 0 || j === 6 ? "left" : "center", bold: j === 0 || v.startsWith("○"), fontSize: j === 0 ? 17 : (v.length > 2 ? 14 : 20), color: j === 0 ? K.navy : (v.startsWith("○") ? K.teal : K.sub), fill: { color: i % 2 ? K.white : K.bg } } }))));
    s.addTable(rows, { x: 0.6, y: 1.5, w: 12.1, colW: [2.0, 1.4, 1.7, 2.0, 1.2, 1.5, 2.3], rowH: [0.42, 0.42, 0.62, 0.62, 0.62, 0.62], fontSize: 16, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    T(s, "○：その災害のときに逃げる場所として指定　×：その災害のときの避難場所には指定されていない　「大規模な火事」は4校とも×", { x: 0.6, y: 4.95, w: 12.1, h: 0.35, fontSize: 14, color: K.sub });
    box(s, 0.6, 5.42, 12.1, 0.72, K.tealL);
    T(s, "どの選択肢でも、学校が地域の避難所として役割を果たせるかを、比べる視点に入れます。", { x: 0.95, y: 5.42, w: 11.5, h: 0.72, fontSize: 18, bold: true, color: K.navy, valign: "middle" });
    note(s, "出典：長野市「指定緊急避難場所一覧（指定避難所、広域避難場所を含む）」令和7年（2025年）7月17日時点（若穂地区は5か所。ほかの1か所は若穂中央公園で、緊急避難場所のみ）。\n防災備蓄倉庫：長野市ホームページ（令和8年（2026年）9月15日現在。若穂地区は保科小・若穂支所・若穂中央公園）。", 6.28);
    s.addNotes("「洪水など」で×の綿内小・川田小は、洪水などのときには逃げる場所に指定されていない（地震や土砂災害のときは指定）。若穂中・保科小は洪水などのとき2階以上が指定。保科小の土砂災害は教室棟の建物内。これは学校の良し悪しではなく、場所と建物の条件によるもの。どの選択肢でも、地区ごとに洪水などのときに逃げる場所がどうなるかを確認する必要がある。");
  }

  // 15 生活への影響
  {
    const s = content("どの選択肢でも、生活への影響を見える化する");
    const v = [["通学", "各地区からの所要時間・スクールバスの要否・冬の安全対策", "MdDirectionsBus"], ["送迎", "保護者の送り迎えの負担、学校周辺の駐車と交通安全", "MdDirectionsCar"],
      ["放課後", "放課後児童クラブ（※）・児童センターの場所と定員、帰り道", "MdChildFriendly"], ["部活動", "中学校の部活動の種類と人数、地域のクラブへの移行の動き", "MdSportsBasketball"],
      ["学校生活", "行事、学級の人数、友だち関係", "MdSchool"], ["支援が必要な子", "特別支援学級・通級、医療的ケア、学校に行きづらい子の居場所", "MdAccessibility"],
      ["地域", "運動会や祭りなど地区の行事、防災拠点、交流の継続", "MdHome"], ["施設・費用", "校舎・跡地の使い方、整備費・運営費・保護者の負担", "MdDomain"]];
    for (let i = 0; i < v.length; i++) {
      const [h, d, ic] = v[i], y = 1.35 + i * 0.64;
      box(s, 0.6, y, 12.1, 0.57, i % 2 ? K.white : K.bg, i % 2 ? K.line : null);
      await iconCircle(s, 0.8, y + 0.06, 0.45, ic);
      T(s, h, { x: 1.5, y, w: 2.4, h: 0.57, fontSize: 18, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: 3.9, y, w: 8.65, h: 0.57, fontSize: 17, valign: "middle" });
    }
    note(s, "※放課後児童クラブ：保護者が働いている家庭などの小学生が、放課後を過ごす場。", 6.5);
  }

  // 追加D どの案でもすぐ始められる学び
  {
    const s = content("どの選択肢でも、すぐに始められる学びの例", "結論を待たずに、3校と中学校で試せること");
    const ex = [["3校合同の授業・行事", "合同の体育・音楽、合同の遠足や宿泊学習など、大きな集団での活動", "MdGroups"], ["保育園から中学校までの連携", "保育園・幼稚園児の学校体験、中学校の先生による出前授業、小6の中学校体験", "MdSchool"],
      ["オンラインの交流授業", "タブレットで3校の教室をつなぎ、一緒に話し合う授業", "MdComputer"], ["地域と一緒の学び", "地域の方を先生に迎える授業、地区の行事への参加", "MdHandshake"]];
    for (let i = 0; i < 4; i++) {
      const [h, d, ic] = ex[i], x = 0.6 + (i % 2) * 6.15, y = 1.7 + Math.floor(i / 2) * 2.25;
      box(s, x, y, 5.95, 2.05, K.bg);
      await iconCircle(s, x + 0.35, y + 0.35, 0.9, ic);
      T(s, h, { x: x + 1.5, y: y + 0.3, w: 4.2, h: 0.6, fontSize: 21, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: x + 1.5, y: y + 0.95, w: 4.2, h: 0.9, fontSize: 16, color: K.sub });
    }
    T(s, "試した結果は、選択肢を比べるときの材料にもなります。", { x: 0.6, y: 6.2, w: 12.1, h: 0.4, fontSize: 16, bold: true, color: K.sub, align: "center" });
    note(s, "例として示すものです。実施するかどうかと内容は、各学校と教育委員会が検討します。", 6.65);
  }

  divider("04", "意見を寄せる", "これからの進め方と、皆さんに考えてほしいこと"); SEC = "04  意見を寄せる";
  // 16 ロードマップ
  {
    const s = content("今後の進め方｜検討ロードマップ（案）");
    const ph = [["R8（2026）\n秋～冬", "現状・課題の共有", "住民意見の収集"], ["R9（2027）\n前半", "保護者・子どもの意見を聞く", "未就学児の保護者も対象"], ["R9（2027）\n後半", "複数案の比較", "通学・施設の調査を反映し中間報告"], ["R10（2028）\n前半", "地区別対話", "地域の方向性の整理"], ["R10（2028）\n後半以降", "地域提言", "市による検討・判断"]];
    ph.forEach(([d, a, b], i) => {
      const x = 0.6 + i * 2.44, now = i === 0;
      s.addShape(pres.shapes.CHEVRON, { x, y: 1.5, w: 2.5, h: 0.9, fill: { color: now ? K.amber : (i < 4 ? K.teal : K.navy) }, line: { type: "none" } });
      T(s, d, { x: x + 0.4, y: 1.5, w: 1.8, h: 0.9, fontSize: 16, bold: true, color: K.white, valign: "middle", align: "center" });
      box(s, x + 0.05, 2.55, 2.25, 2.0, now ? K.amberL : K.bg);
      if (now) T(s, "今ここ", { x: x + 0.3, y: 2.65, w: 1.5, h: 0.3, fontSize: 13, bold: true, color: K.amber });
      T(s, a, { x: x + 0.3, y: 2.95, w: 1.85, h: 0.75, fontSize: 18, bold: true, color: K.navy });
      T(s, b, { x: x + 0.3, y: 3.75, w: 1.85, h: 0.75, fontSize: 16, color: K.sub });
    });
    T(s, "誰が決めるのか", { x: 0.6, y: 4.75, w: 4, h: 0.4, fontSize: 18, bold: true, color: K.navy });
    const who = [["地域（検討委員会）", "意見をまとめ、市へ提言する"], ["長野市教育委員会", "提言を受けて、方針を検討・決定する"], ["長野市議会", "学校の設置・廃止は条例で定めるため、議会で審議・議決する"]];
    who.forEach(([h, d], i) => {
      const x = 0.6 + i * 4.1;
      box(s, x, 5.2, 3.75, 1.15, K.tealL);
      T(s, h, { x: x + 0.2, y: 5.27, w: 3.4, h: 0.4, fontSize: 17, bold: true, color: K.teal });
      T(s, d, { x: x + 0.2, y: 5.67, w: 3.4, h: 0.65, fontSize: 15, color: K.ink });
      if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 3.8, y: 5.6, w: 0.25, h: 0.35, fill: { color: K.teal }, line: { type: "none" } });
    });
    note(s, "スケジュールは提案であり、正式決定ではありません。実施の時期は提言の後に市が決めます。在校生への配慮の方針・手続きの詳細は［事務局に確認中］。", 6.5);
  }

  // 追加E 意見の出し方
  {
    const s = content("ご意見の出し方と、寄せられた意見の扱い方");
    const cols = [["出し方", "MdEditNote", [["アンケート用紙", "本日配布。会場の回収箱、または［提出先］へ"], ["インターネット", "［URL・QRコード］から回答できます"], ["地区での懇談会", "［日程・会場］"]]],
      ["締切・問い合わせ先", "MdPhone", [["締切", "令和8年（2026年）［　］月［　］日"], ["提出先", "［　　　　　　　］"], ["問い合わせ", "長野市教育委員会事務局\n学校教育課\n電話［　　　　　　］"]]],
      ["意見の扱い方（案）", "MdLock", [["すべて読みます", "検討委員会の委員が、寄せられた意見をすべて読みます"], ["地区ごとに公表します", "地区ごとに集計し、委員会の考え方とあわせて［時期・方法］に公表します"], ["個人は特定しません", "お名前などの個人情報は公表しません"]]]];
    for (let i = 0; i < 3; i++) {
      const [h, ic, items] = cols[i], x = 0.6 + i * 4.1;
      box(s, x, 1.45, 3.9, 5.0, K.bg);
      await iconCircle(s, x + 0.3, 1.65, 0.7, ic);
      T(s, h, { x: x + 1.15, y: 1.65, w: 2.65, h: 0.7, fontSize: 20, bold: true, color: K.navy, valign: "middle" });
      items.forEach(([a, b], j) => {
        const y = 2.55 + j * 1.3;
        T(s, a, { x: x + 0.3, y, w: 3.4, h: 0.4, fontSize: 17, bold: true, color: K.teal });
        T(s, b, { x: x + 0.3, y: y + 0.42, w: 3.4, h: 0.85, fontSize: 15, color: K.ink });
      });
    }
    note(s, "本日の資料は［ホームページ・回覧］でもご覧いただけます。［ ］は事務局で確定後に記入します。意見の扱い方は案であり、委員会で決定します。", 6.6);
    s.addNotes("ご意見は、賛成・反対だけでなく「心配なこと」「知りたいこと」も歓迎する、と口頭で伝える。");
  }

  // 追加F 今日のまとめ
  {
    const s = content("今日のまとめ");
    const pts = [["児童生徒は減る見通し", `3小学校の合計は${tot[0]}人→${tot[6]}人（R14＝2032年度）。若穂中も減少が続きます。`, "MdTrendingDown"],
      ["何も決まっていません", "選択肢A〜Eを、同じ物差しで比べている途中です。", "MdCompareArrows"],
      ["決める前に、皆さんの声を聞きます", "保護者・子ども・地域の意見を聞き、地区ごとにまとめて公表します。", "MdForum"],
      ["ご意見をお寄せください", "締切：令和8年（2026年）［　］月［　］日　提出先：［　　　　］", "MdEditNote"]];
    for (let i = 0; i < 4; i++) {
      const [h, d, ic] = pts[i], y = 1.4 + i * 1.28;
      box(s, 0.6, y, 12.1, 1.12, i % 2 ? K.white : K.bg, i % 2 ? K.line : null);
      await iconCircle(s, 0.85, y + 0.18, 0.76, ic);
      T(s, h, { x: 1.9, y: y + 0.1, w: 10.5, h: 0.48, fontSize: 21, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: 1.9, y: y + 0.58, w: 10.5, h: 0.45, fontSize: 17, color: K.ink, valign: "middle" });
    }
  }

  // 17 一緒に考えたいこと
  {
    const s = pres.addSlide({ masterName: "DARK" });
    s.addText("皆さんと一緒に考えたいこと", { placeholder: "title" });
    const q = ["子どもたちに、どんな学びを経験してほしいですか？", "今の学校の良さで、残したいものは何ですか？", "通学・安全・地域活動で心配なことは何ですか？", "判断の前に、どんな情報が必要ですか？"];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + (i % 2) * 6.15, y = 1.65 + Math.floor(i / 2) * 1.75;
      box(s, x, y, 5.95, 1.55, "2A5A47");
      T(s, `Q${i + 1}`, { x: x + 0.35, y, w: 0.9, h: 1.55, fontSize: 28, bold: true, color: K.amber, valign: "middle" });
      T(s, q[i], { x: x + 1.3, y, w: 4.4, h: 1.55, fontSize: 19, color: K.white, valign: "middle" });
    }
    T(s, "目指すのは、子どもたちの未来の学びをつくること。", { x: 0.6, y: 5.45, w: 12.1, h: 0.8, fontSize: 26, bold: true, color: K.amber, align: "center", valign: "middle" });
  }

  if (APPENDIX) {
  divider("05", "参考資料", "学年ごとの見通し、地区ごとの道のりと人口、用語の説明"); SEC = "参考資料";
  // 参考1〜3 学年別の見通し（3校とも同じ形式）
  for (const n of ["綿内小", "川田小", "保科小"]) {
    const s = content(`${n}｜学年ごとの児童数の見通し`, "同じ子どもたちは、1年ごとに右下へ1マスずつ進みます");
    const H = t => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, align: "center" } });
    const rows = [[H("学年"), ...Y7.map(r => H(`${r}\n${AD(r)}`))]];
    Object.entries(grades[n]).forEach(([g, v], i) => {
      const sum = g === "合計";
      rows.push([{ text: g, options: { bold: true, align: "center", fill: { color: sum ? K.tealL : (i % 2 ? K.white : K.bg) } } },
        ...v.map(x => ({ text: String(x), options: { align: "center", bold: sum || x < 20, color: K.ink, fill: { color: sum ? K.tealL : (x < 20 ? K.mark : (i % 2 ? K.white : K.bg)) } } }))]);
    });
    s.addTable(rows, { x: 0.6, y: 1.6, w: 12.1, colW: [1.6, ...Y7.map(() => 1.5)], rowH: [0.7, ...Array(7).fill(0.56)], fontSize: 18, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 6.32, w: 0.3, h: 0.22, fill: { color: K.mark }, line: { color: K.line, width: 1 } });
    T(s, "20人未満の学年（長野県の目安「学年20人程度」を下回る学年）", { x: 1.0, y: 6.28, w: 11, h: 0.3, fontSize: 14, color: K.sub });
    note(s, "出典：提供資料「児童数推計」（令和8年（2026年）5月1日現在、長野市教育委員会）。将来推計は変動します。", 6.65);
    if (n === "川田小") s.addNotes("R11の4年「25」は写真からの読み取り値で、斜めに続く数値（R10の3年22人）と合わない。原本と照合すること。");
  }

  // 参考 地区ごとの子どもから各学校までの道のり
  {
    const s = content("地区ごとの、子どもから各学校までの道のり（参考）", "統合先を想定した表ではありません。どの地区からも、すべての学校までを同じ方法で計算した参考の数です");
    const hd = t => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, align: "center", fontSize: 15 } });
    const dests = ["綿内小", "川田小", "保科小", "若穂中"];
    const rows = [[hd("住んでいる地区"), ...dests.map(d => hd(`${d}まで`))]];
    [["綿内地区", "綿内", "綿内小"], ["川田地区", "川田", "川田小"], ["保科地区", "保科", "保科小"]].forEach(([lab, h, own], i) => rows.push([
      { text: `${lab}\n（0〜14歳 ${TG[h + ">" + own].kids}人）`, options: { bold: true, color: K.navy, fontSize: 15 } },
      ...dests.map(d => { const lim = d === "若穂中" ? 5 : 4; return { text: `平均 ${km(TG[h + ">" + d].avg_m)}km\n${d === "若穂中" ? "6" : "4"}km未満 ${share(h + ">" + d, lim)}%`, options: { align: "center", fontSize: 16, bold: d === own, fill: { color: d === own ? K.tealL : (i % 2 ? K.white : K.bg) } } }; })]));
    s.addTable(rows, { x: 0.6, y: 1.65, w: 12.1, colW: [3.3, 2.2, 2.2, 2.2, 2.2], rowH: [0.55, 0.95, 0.95, 0.95], fontSize: 16, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    T(s, "色の付いた欄は、今その地区の子どもが通う学校です。「4km未満」「6km未満」は、国の手引の目安（小学校おおむね4km・中学校おおむね6km）による区切りです。", { x: 0.6, y: 5.25, w: 12.1, h: 0.6, fontSize: 16, color: K.ink });
    note(s, "250m四方の中心から各学校までの徒歩の道のり（OpenStreetMapによる経路計算、令和8年（2026年）11月）。0〜14歳は国勢調査（令和2年＝2020年）。地区は国勢調査の小地域（大字）による。各家庭からの正確な距離ではありません。", 6.4);
  }

  // 参考 町ごとの人口
  {
    const s = content("地区・町ごとの人口（参考）", "令和8年（2026年）10月1日現在。町ごとの子どもの人数は公表されていません");
    const cell = (v, i, o = {}) => ({ text: v, options: Object.assign({ fill: { color: i % 2 ? K.white : K.bg } }, o) });
    const num = (m, i) => cell(m ? `${fmt(m.p)}人` : "", i, { align: "right" });
    [["綿内", 0.6, 5.7, 2], ["川田", 6.5, 2.9, 1], ["保科", 9.6, 3.1, 1]].forEach(([c, x, w, ncol]) => {
      const list = MACHI.filter(m => m.c === c), per = Math.ceil(list.length / ncol);
      const rows = [[{ text: `${c}地区　計 ${fmt(list.reduce((a, m) => a + m.p, 0))}人`, options: { colspan: ncol * 2, bold: true, color: K.white, fill: { color: K.navy }, align: "center" } }]];
      for (let r = 0; r < per; r++) { const row = []; for (let k = 0; k < ncol; k++) { const m = list[k * per + r]; row.push(cell(m ? m.n : "", r), num(m, r)); } rows.push(row); }
      const cw = w / ncol;
      s.addTable(rows, { x, y: 1.6, w, colW: Array(ncol).fill([cw * 0.55, cw * 0.45]).flat(), rowH: 0.33, fontSize: 14, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 0.5, color: K.line } });
    });
    note(s, "出典：長野市「町別人口及び世帯数」（令和8年（2026年）10月1日現在、住民基本台帳）。市の区分のため、行政連絡区や地図上の区分と一致しない場合があります。", 6.55);
  }

  // 参考4〜5 用語集
  const terms = [["推計", "今の子どもの数などをもとにした将来の見込み。転入・転出で変わります。"], ["再編", "学校の配置や仕組みを見直すこと。今の形を続けることも含めて検討します。"],
    ["統合", "複数の学校を一つにまとめること。"], ["通学区域（学区）", "住んでいる場所ごとに決められた、通う学校の範囲。"],
    ["複式学級", "子どもが少ないときに、2つの学年を1つの学級にまとめたもの。"], ["小中一貫型の\n小学校・中学校", "小学校・中学校は別の学校のまま、9年間の教育課程をそろえる形。"],
    ["義務教育学校", "小学校・中学校とは別の区分の学校。9年間を一つの学校で教える。"], ["中1ギャップ", "中学校に進んだときの環境の変化になじめず、つまずきやすくなること。"],
    ["特別支援学級", "障がいのある子どものために、小中学校の中に置かれる少人数の学級。"], ["通級", "ふだんの学級に在籍しながら、一部の時間だけ別の教室で支援を受ける仕組み。"],
    ["指定緊急避難場所", "災害が迫ったときに、命を守るために逃げる場所。災害の種類ごとに指定されます。"], ["指定避難所", "災害で家に戻れない人が、しばらく過ごすために市が指定した施設。"],
    ["長寿命化計画", "建物を計画的に改修して、長く使うための市の計画。"], ["放課後児童クラブ", "保護者が働いている家庭などの小学生が、放課後を過ごす場。"]];
  for (let k = 0; k < 2; k++) {
    const s = content(`用語の説明（${k + 1}/2）`);
    terms.slice(k * 8, k * 8 + 8).forEach(([h, d], i) => {
      const x = 0.6 + (i % 2) * 6.15, y = 1.4 + Math.floor(i / 2) * 1.22;
      box(s, x, y, 5.95, 1.1, i % 4 < 2 ? K.bg : K.white, i % 4 < 2 ? null : K.line);
      T(s, h, { x: x + 0.2, y, w: 2.0, h: 1.1, fontSize: 17, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: x + 2.25, y, w: 3.55, h: 1.1, fontSize: 16, valign: "middle" });
    });
    note(s, "年の表し方：令和の年に2018を足すと西暦になります（令和8年度＝2026年度、令和14年度＝2032年度、令和20年度＝2038年度）。", 6.4);
  }
  }

  await pres.writeFile({ fileName: OUT });
  if (applyTheme) await applyTheme(OUT, THEME);
  await require("./fix_fonts")(OUT, FONT);
  console.log("ok");
})();

const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
let applyTheme = null;
try { ({ applyTheme } = require(process.env.APPLY_THEME || "/mnt/skills/public/pptx/scripts/apply_theme.js")); } catch (e) { /* optional: theme colors only */ }

const OUT = process.argv[2] || "output/若穂地区_住民説明会資料.pptx";
const FONT = "BIZ UDPゴシック";
const K = { navy: "1D4636", teal: "2E7D5B", tealL: "E8F2EC", amber: "D9573E", amberL: "FCEBE6", ink: "1E2A26", sub: "56635D", line: "D7DFDA", bg: "F4F7F5", white: "FFFFFF", grey: "9AA59F", hill1: "2A5C47", hill2: "3B7259", hill3: "4E8A6C" };
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
const note = (s, text, y = 6.55) => T(s, text, { x: 0.6, y, w: 12.1, h: 0.3, fontSize: 11, color: K.grey });
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
const chartBase = { catAxisLabelColor: K.sub, valAxisLabelColor: K.sub, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", catGridLine: { style: "none" } };

(async () => {
  // 1 Cover
  {
    const s = pres.addSlide({ masterName: "COVER" });
    hills(s);
    T(s, "若穂地区 小中学校再編検討｜住民説明会資料（案）", { x: 0.9, y: 0.7, w: 9, h: 0.4, fontSize: 15, color: "CFE0D6" });
    T(s, "これからの若穂の学校を考える", { x: 0.9, y: 1.5, w: 11.5, h: 1.1, fontSize: 50, bold: true, color: K.white });
    T(s, "学校の数を決める前に、どんな学びを残したいか。", { x: 0.9, y: 2.75, w: 11.5, h: 0.7, fontSize: 28, bold: true, color: K.amber });
    T(s, "子どもたちの未来と、地域のこれからを一緒に考えるために", { x: 0.9, y: 3.55, w: 11, h: 0.5, fontSize: 18, color: "CFE0D6" });
    T(s, "R8.○.○　若穂地区小中学校再編検討委員会／長野市教育委員会事務局 学校教育課", { x: 0.9, y: 4.45, w: 11.5, h: 0.4, fontSize: 14, color: "CFE0D6" });
    s.addNotes("表紙。委員会名・事務局名の表記は要確認。");
  }

  SEC = "はじめに";
  // 2 今日の流れ
  {
    const s = content("今日の説明会で共有したいこと");
    const steps = [["現状を知る", "児童生徒数・施設の状況", "MdInsights"], ["課題を共有する", "教育環境・通学・地域への影響", "MdForum"], ["選択肢を比べる", "現状維持から再編まで", "MdCompareArrows"], ["意見を寄せる", "大切にしたいこと・懸念点", "MdRateReview"]];
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
  }

  // 3 出発点
  {
    const s = content("検討の出発点｜誰のための学校再編か");
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
    const s = content("若穂地区の3小学校｜児童数の見通し", "令和8年度 → 令和14年度（教育委員会推計）");
    s.addChart(pres.charts.BAR, [
      { name: "令和8年度", labels: Object.keys(sch), values: Object.values(sch).map(a => a[0]) },
      { name: "令和14年度", labels: Object.keys(sch), values: Object.values(sch).map(a => a[6]) },
    ], Object.assign({}, chartBase, {
      x: 0.5, y: 1.65, w: 7.6, h: 4.8, barDir: "bar", chartColors: [K.grey, K.teal], barGapWidthPct: 60, barOverlapPct: -10,
      catAxisOrientation: "maxMin", valAxisHidden: true, valGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 320,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelColor: K.ink, dataLabelFormatCode: '0"人"',
      catAxisLabelFontSize: 18, catAxisLabelColor: K.ink, showLegend: true, legendPos: "t", legendFontSize: 14, legendColor: K.ink,
    }));
    const rows = Object.entries(sch);
    rows.forEach(([n, a], i) => {
      const y = 1.75 + i * 1.55;
      box(s, 8.5, y, 4.2, 1.35, i === 1 ? K.amberL : K.bg);
      T(s, n, { x: 8.8, y: y + 0.2, w: 1.6, h: 0.45, fontSize: 18, bold: true, color: K.navy });
      T(s, `${a[0]} → ${a[6]}人`, { x: 8.8, y: y + 0.7, w: 2.2, h: 0.4, fontSize: 15, color: K.sub });
      T(s, `−${dec(a[0], a[6])}%`, { x: 10.8, y: y + 0.15, w: 1.7, h: 1.05, fontSize: 34, bold: true, color: i === 1 ? K.amber : K.teal, align: "right", valign: "middle" });
    });
    note(s, "出典：提供資料「児童数推計」（R8.5.1現在）。将来推計は変動します。");
  }

  // 5 合計
  {
    const s = content("3小学校の児童数合計は約32％減少", "令和8年度 → 令和14年度");
    T(s, `${tot[0]}`, { x: 0.6, y: 1.8, w: 2.6, h: 1.2, fontSize: 66, bold: true, color: K.grey, align: "right" });
    T(s, "人", { x: 3.25, y: 2.35, w: 0.5, h: 0.6, fontSize: 22, color: K.grey });
    T(s, "令和8年度", { x: 0.6, y: 3.0, w: 3.1, h: 0.4, fontSize: 15, color: K.sub, align: "right" });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 3.95, y: 2.15, w: 1.0, h: 0.6, fill: { color: K.line }, line: { type: "none" } });
    T(s, `${tot[6]}`, { x: 5.1, y: 1.8, w: 2.6, h: 1.2, fontSize: 66, bold: true, color: K.teal });
    T(s, "人", { x: 6.95, y: 2.35, w: 0.5, h: 0.6, fontSize: 22, color: K.teal });
    T(s, "令和14年度", { x: 5.1, y: 3.0, w: 3.1, h: 0.4, fontSize: 15, color: K.sub });
    s.addChart(pres.charts.BAR, [{ name: "3小学校の合計", labels: Y7, values: tot }], Object.assign({}, chartBase, {
      x: 8.3, y: 1.55, w: 4.5, h: 2.9, barDir: "col", chartColors: [K.teal], barGapWidthPct: 40, valAxisHidden: true, valGridLine: { style: "none" }, valAxisMinVal: 0,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: K.ink, catAxisLabelFontSize: 12, showLegend: false,
      showTitle: true, title: "3小学校の合計（人）", titleFontSize: 13, titleColor: K.ink, titleFontFace: "+mn-lt",
    }));
    box(s, 0.6, 4.75, 12.1, 1.6, K.bg);
    T(s, "学校ごとに異なる減少の姿", { x: 0.95, y: 4.95, w: 11.5, h: 0.45, fontSize: 20, bold: true, color: K.navy });
    T(s, "川田小は約47％減、保科小では令和8年度の1年生が2人。綿内小も減少が続きます。", { x: 0.95, y: 5.5, w: 11.5, h: 0.6, fontSize: 17 });
    note(s, "数値は提供推計（R8.5.1現在）の合計から算出。");
  }

  // 6 若穂中
  {
    const s = content("若穂中学校も生徒数が減少する見通し", "令和8年度 → 令和20年度");
    s.addChart(pres.charts.BAR, [{ name: "若穂中 全校生徒数", labels: Y13, values: jh }], Object.assign({}, chartBase, {
      x: 0.5, y: 1.55, w: 8.0, h: 4.85, barDir: "col", chartColors: [K.teal], barGapWidthPct: 35, valAxisHidden: true, valGridLine: { style: "none" }, valAxisMinVal: 0,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: K.ink, catAxisLabelFontSize: 13, showLegend: false,
    }));
    box(s, 8.9, 1.65, 3.8, 2.1, K.tealL);
    T(s, `約${dec(jh[0], jh[12])}%減`, { x: 9.2, y: 1.8, w: 3.3, h: 1.0, fontSize: 44, bold: true, color: K.teal });
    T(s, `${jh[0]}人 → ${jh[12]}人`, { x: 9.2, y: 2.85, w: 3.3, h: 0.5, fontSize: 18, color: K.navy, bold: true });
    box(s, 8.9, 4.0, 3.8, 2.35, K.bg);
    T(s, "小学校だけでなく、地区全体の教育環境を考える", { x: 9.2, y: 4.2, w: 3.3, h: 1.0, fontSize: 17, bold: true, color: K.navy });
    T(s, "小中9年間の学びの連続性も、検討する価値のある視点です。", { x: 9.2, y: 5.25, w: 3.3, h: 0.9, fontSize: 15, color: K.sub });
    note(s, "出典：提供資料「生徒数推計」（R8.5.1現在）。");
  }

  // 7 市全体と比較
  {
    const s = content("長野市全体と比べても、若穂の減少は大きい", "小学校児童数：令和8年度を100としたときの令和14年度");
    const rows = [["若穂地区 3小学校", tot[0], tot[6], K.teal], ["長野市立小学校 全体", city[0], city[1], K.grey]];
    rows.forEach(([n, a, b, c], i) => {
      const y = 1.9 + i * 1.75, W = 7.2, r = b / a;
      T(s, n, { x: 0.6, y, w: 4, h: 0.45, fontSize: 19, bold: true, color: K.navy });
      T(s, `${a.toLocaleString()}人 → ${b.toLocaleString()}人`, { x: 4.6, y, w: 3.2, h: 0.45, fontSize: 15, color: K.sub, align: "right" });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y + 0.6, w: W, h: 0.6, fill: { color: "EEF1F2" }, line: { type: "none" } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y + 0.6, w: W * r, h: 0.6, fill: { color: c }, line: { type: "none" } });
      T(s, `${Math.round(r * 100)}`, { x: 0.75, y: y + 0.6, w: 1, h: 0.6, fontSize: 18, bold: true, color: K.white, valign: "middle" });
      T(s, `約${dec(a, b)}%減`, { x: 8.1, y: y + 0.35, w: 2.6, h: 0.9, fontSize: 36, bold: true, color: i ? K.grey : K.teal, valign: "middle" });
    });
    box(s, 0.6, 5.55, 12.1, 0.85, K.tealL);
    T(s, "地域固有の課題として、将来を見据えた検討が必要です。", { x: 0.95, y: 5.55, w: 11.5, h: 0.85, fontSize: 20, bold: true, color: K.navy, valign: "middle" });
    note(s, "出典：提供推計データ（R8.5.1現在）。減少率は四捨五入。", 6.6);
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
    T(s, "学校規模だけで教育の良し悪しが決まるものではありません。", { x: 0.6, y: 6.3, w: 12.1, h: 0.4, fontSize: 16, bold: true, color: K.sub, align: "center" });
  }

  // 9 なぜ今
  {
    const s = content("なぜ今、検討するのか");
    const rs = [["児童生徒数の減少は、今後も続く見通し", "小学生はR14までに約3割、中学生はR20までに5割以上減る推計", "MdTrendingDown"],
      ["施設の更新・改修を考える時期と重なる", "4校すべてが2030年度までに計画上の更新・改修時期を迎える", "MdApartment"],
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
    const s = content("施設の状況も、判断材料のひとつ", "建築年度と、計画上の更新・改修時期（長寿命化計画に記載された情報）");
    const x0 = 2.6, x1 = 12.4, yr0 = 1970, yr1 = 2035, X = y => x0 + (y - yr0) / (yr1 - yr0) * (x1 - x0);
    [1970, 1980, 1990, 2000, 2010, 2020, 2030].forEach(yr => {
      s.addShape(pres.shapes.LINE, { x: X(yr), y: 1.85, w: 0, h: 4.1, line: { color: "E7ECEE", width: 1 } });
      T(s, String(yr), { x: X(yr) - 0.4, y: 1.6, w: 0.8, h: 0.25, fontSize: 11, color: K.grey, align: "center" });
    });
    s.addShape(pres.shapes.LINE, { x: X(2026), y: 1.85, w: 0, h: 4.1, line: { color: K.amber, width: 1.5, dashType: "dash" } });
    T(s, "現在", { x: X(2026) - 0.4, y: 5.98, w: 0.8, h: 0.25, fontSize: 11, color: K.amber, bold: true, align: "center" });
    const rows = [["川田小", 1973, 2023, "更新"], ["綿内小", 1980, 2030, "更新"], ["若穂中", 1987, 2027, "改修"], ["保科小", 2010, 2030, "改修"]];
    rows.forEach(([n, b, p, k], i) => {
      const y = 2.15 + i * 0.95;
      T(s, n, { x: 0.6, y, w: 1.0, h: 0.5, fontSize: 18, bold: true, color: K.navy, valign: "middle" });
      T(s, `${b}年度築`, { x: 1.5, y, w: 1.05, h: 0.5, fontSize: 12, color: K.sub, valign: "middle" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X(b), y: y + 0.1, w: X(2026) - X(b), h: 0.3, fill: { color: i === 0 ? K.navy : "9CC3AE" }, line: { type: "none" }, rectRadius: 0.15 });
      T(s, `築${2026 - b}年`, { x: X(b) + 0.15, y: y + 0.1, w: 1.2, h: 0.3, fontSize: 11, bold: true, color: K.white, valign: "middle" });
      s.addShape(pres.shapes.DIAMOND, { x: X(p) - 0.14, y: y + 0.11, w: 0.28, h: 0.28, fill: { color: K.amber }, line: { color: K.white, width: 1 } });
      T(s, `${p}年度 ${k}`, { x: X(p) + 0.2, y: y + 0.45, w: 1.6, h: 0.3, fontSize: 11, color: "A3402C", bold: true, align: p > 2028 ? "right" : "left", ...(p > 2028 ? { x: X(p) - 1.6 } : {}) });
    });
    note(s, "◆計画上の更新・改修時期　出典：長野市学校施設長寿命化計画（令和2年時点）。実施済み工事・最新計画の確認が必要です。", 6.4);
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

  // 12 選択肢
  {
    const s = content("考えられる選択肢を公平に比較する");
    const o = [["A", "現状維持", "3小学校・中学校を維持", "小規模化への対応"], ["B", "一部統合", "対象校を段階的に統合", "地域間の公平性・通学"], ["C", "小学校3校統合", "小学校を1校に集約", "校地・規模・通学"], ["D", "小中一貫の制度活用", "義務教育学校等を検討", "教育課程・施設・通学"]];
    o.forEach(([L, h, d, p], i) => {
      const x = 0.6 + i * 3.08;
      box(s, x, 1.6, 2.86, 4.6, K.bg);
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: 1.9, w: 0.85, h: 0.85, fill: { color: K.navy }, line: { type: "none" } });
      T(s, L, { x: x + 0.3, y: 1.9, w: 0.85, h: 0.85, fontSize: 28, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, h, { x: x + 0.3, y: 2.95, w: 2.4, h: 0.95, fontSize: 21, bold: true, color: K.navy });
      T(s, d, { x: x + 0.3, y: 3.95, w: 2.4, h: 0.8, fontSize: 16 });
      s.addShape(pres.shapes.LINE, { x: x + 0.3, y: 4.85, w: 2.26, h: 0, line: { color: K.line, width: 1 } });
      T(s, "主な論点", { x: x + 0.3, y: 5.0, w: 2.4, h: 0.3, fontSize: 12, color: K.teal, bold: true });
      T(s, p, { x: x + 0.3, y: 5.32, w: 2.4, h: 0.7, fontSize: 15, color: K.sub });
    });
    note(s, "学校配置と教育制度は別の軸です。複数の方式を組み合わせることも可能です。", 6.4);
  }

  // 13 同じ物差し matrix
  {
    const s = content("選択肢は同じ物差しで比較する", "評価欄は、今後の調査結果をもとに記入していきます");
    const H = t => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, align: "center" } });
    const crit = [["教育", "学習環境・交流・教員体制"], ["通学", "距離・時間・交通安全・費用"], ["施設", "安全性・整備費・維持管理費"], ["地域", "防災・行事・跡地利用"], ["実現性", "準備期間・財源・合意形成"]];
    const rows = [[{ text: "視点", options: { bold: true, color: K.white, fill: { color: K.navy } } }, H("比べる内容"), H("A 現状維持"), H("B 一部統合"), H("C 3校統合"), H("D 小中一貫")]];
    crit.forEach(([a, b], i) => rows.push([{ text: a, options: { bold: true, color: K.teal, fontSize: 18 } }, { text: b, options: {} }, ...[0, 1, 2, 3].map(() => ({ text: "", options: { fill: { color: i % 2 ? K.white : K.bg } } }))]));
    s.addTable(rows, { x: 0.6, y: 1.7, w: 12.1, colW: [1.4, 3.9, 1.7, 1.7, 1.7, 1.7], rowH: 0.78, fontSize: 16, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    note(s, "各選択肢の評価は、通学・施設・費用などの調査結果をもとに、委員会と地域で確認しながら記入します。", 6.6);
  }

  // 14 義務教育学校
  {
    const s = content("義務教育学校という選択肢", "9年間を見通した教育を一体的に進める学校制度");
    const cols = [["期待される点", K.teal, K.tealL, "MdTrendingUp", ["小中の学びの連続性", "教員同士の連携", "異学年交流", "中学校進学時の不安軽減"]],
      ["慎重に考える点", K.amber, K.amberL, "MdWarningAmber", ["人間関係の固定化への配慮", "年齢差の大きい子どもの生活", "通学距離と安全", "施設配置と運営体制"]]];
    for (let i = 0; i < 2; i++) {
      const [h, c, cl, ic, items] = cols[i], x = 0.6 + i * 6.15;
      box(s, x, 1.65, 5.95, 3.95, cl);
      await iconCircle(s, x + 0.4, 1.95, 0.75, ic, c);
      T(s, h, { x: x + 1.35, y: 1.95, w: 4.3, h: 0.75, fontSize: 24, bold: true, color: K.navy, valign: "middle" });
      items.forEach((t, j) => {
        const y = 2.95 + j * 0.65;
        s.addShape(pres.shapes.OVAL, { x: x + 0.55, y: y + 0.16, w: 0.18, h: 0.18, fill: { color: c }, line: { type: "none" } });
        T(s, t, { x: x + 0.95, y, w: 4.8, h: 0.5, fontSize: 19, valign: "middle" });
      });
    }
    box(s, 0.6, 5.8, 12.1, 0.7, K.navy);
    T(s, "義務教育学校＝必ず施設を一つに集約する、という意味ではありません。", { x: 0.95, y: 5.8, w: 11.5, h: 0.7, fontSize: 18, bold: true, color: K.white, valign: "middle" });
  }

  // 追加A 学校間の距離
  {
    const s = content("学校と学校の距離・移動時間の目安", "各小学校の間、各小学校と若穂中の間（地図は位置関係を示す概略図）");
    box(s, 0.6, 1.6, 6.3, 5.0, K.bg);
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
    T(s, "1km", { x: 2.08, y: 6.18, w: 0.6, h: 0.25, fontSize: 11, color: K.sub });
    T(s, "↑ 北", { x: 6.05, y: 1.75, w: 0.7, h: 0.3, fontSize: 12, bold: true, color: K.sub });
    s.addShape(pres.shapes.LINE, { x: 0.9, y: 1.9, w: 0.45, h: 0, line: { color: K.teal, width: 2.5 } });
    T(s, "小学校どうし", { x: 1.42, y: 1.77, w: 1.4, h: 0.26, fontSize: 11, color: K.sub });
    s.addShape(pres.shapes.LINE, { x: 0.9, y: 2.2, w: 0.45, h: 0, line: { color: K.amber, width: 2.5, dashType: "dash" } });
    T(s, "小学校と若穂中", { x: 1.42, y: 2.07, w: 1.6, h: 0.26, fontSize: 11, color: K.sub });
    const H = t => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, align: "center" } });
    const legRows = [["1", "綿内小 ⇔ 川田小", "2.8", "約3.6", "約9分", "約55分"], ["2", "川田小 ⇔ 保科小", "2.8", "約3.7", "約9分", "約55分"], ["3", "綿内小 ⇔ 保科小", "3.9", "約5.1", "約12分", "約75分"],
      ["4", "綿内小 ⇔ 若穂中", "2.2", "約2.9", "約7分", "約45分"], ["5", "川田小 ⇔ 若穂中", "0.9", "約1.1", "約3分", "約15分"], ["6", "保科小 ⇔ 若穂中", "2.3", "約3.0", "約7分", "約45分"]];
    const rows = [[H("No"), H("区間"), H("直線\n(km)"), H("道のり\n(km)"), H("車"), H("徒歩")]];
    legRows.forEach((r, i) => rows.push(r.map((t, j) => ({ text: t, options: { align: j === 1 ? "left" : "center", bold: j < 2, color: j === 0 ? (i < 3 ? K.teal : K.amber) : K.ink, fill: { color: i % 2 ? K.white : K.bg } } }))));
    s.addTable(rows, { x: 7.1, y: 1.6, w: 5.6, colW: [0.45, 1.85, 0.7, 0.8, 0.85, 0.95], rowH: [0.6, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48], fontSize: 13, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    box(s, 7.1, 5.15, 5.6, 1.05, K.amberL);
    T(s, "車は時速25km、徒歩は時速4km（小学生の目安）で概算。道のりは直線距離×1.3の目安です。実際の通学路での実測が必要です。", { x: 7.3, y: 5.2, w: 5.2, h: 0.95, fontSize: 12, color: "A3402C", valign: "middle" });
    note(s, "学校の位置：地図サービスの所在地情報から作図。道路・河川は省略。", 6.7);
    s.addNotes("距離は直線距離からの概算です。統合を検討する段階で、事務局が実際の通学路・スクールバスの経路で測り直す前提で示します。");
  }

  // 追加B 避難所
  {
    const s = content("小中学校は、すべて地域の避難所", "災害時の役割も、学校のあり方を考える大事な視点です");
    const H = t => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, align: "center" } });
    const data = [["綿内小学校", "若穂綿内", "◯", "要確認", "—"], ["川田小学校", "若穂川田", "◯", "要確認", "—"], ["保科小学校", "若穂保科", "◯", "要確認", "児童センター併設、防災備蓄倉庫あり"], ["若穂中学校", "若穂川田", "◯", "要確認", "—"]];
    const rows = [[H("学校"), H("所在地"), H("指定避難所"), H("災害の種類ごとの指定\n（洪水・土砂・地震など）"), H("備考")]];
    data.forEach((r, i) => rows.push(r.map((t, j) => ({ text: t, options: { align: j === 0 || j === 4 ? "left" : "center", bold: j === 0 || j === 2, fontSize: j === 2 ? 22 : 16, color: j === 2 ? K.teal : (j === 3 ? K.grey : K.ink), fill: { color: i % 2 ? K.white : K.bg } } }))));
    s.addTable(rows, { x: 0.6, y: 1.65, w: 12.1, colW: [2.2, 1.8, 1.8, 3.1, 3.2], rowH: [0.8, 0.7, 0.7, 0.7, 0.7], fontSize: 16, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    box(s, 0.6, 5.35, 12.1, 0.95, K.tealL);
    T(s, "再編で使わなくなる校舎が出る場合は、避難所の機能をどう保つかを、地域と一緒に考える必要があります。", { x: 0.95, y: 5.35, w: 11.5, h: 0.95, fontSize: 18, bold: true, color: K.navy, valign: "middle" });
    note(s, "指定避難所：長野市学校施設長寿命化計画 添付資料（令和2年4月時点）。災害種別ごとの指定は、市「指定緊急避難場所・指定避難所一覧」（令和7年7月17日）で確認してください。", 6.45);
    s.addNotes("災害の種類ごとの指定（洪水時に使えるかなど）は、市の最新一覧で事務局に確認のうえ埋めてください。");
  }

  // 15 生活への影響
  {
    const s = content("再編を考えるなら、生活への影響も見える化する");
    const v = [["通学", "各地区からの所要時間・スクールバス・冬季対策", "MdDirectionsBus"], ["学校生活", "行事、部活動、放課後の過ごし方", "MdSportsBasketball"], ["地域", "学校行事・防災拠点・交流の継続", "MdHome"], ["施設", "統合後の校舎・跡地の使い方", "MdDomain"], ["費用", "整備費・運営費・保護者負担", "MdPayments"]];
    for (let i = 0; i < 5; i++) {
      const [h, d, ic] = v[i], y = 1.5 + i * 1.02;
      box(s, 0.6, y, 12.1, 0.88, i % 2 ? K.white : K.bg, i % 2 ? K.line : null);
      await iconCircle(s, 0.85, y + 0.12, 0.64, ic);
      T(s, h, { x: 1.8, y, w: 2.2, h: 0.88, fontSize: 21, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: 4.0, y, w: 8.5, h: 0.88, fontSize: 18, valign: "middle" });
    }
  }

  divider("04", "意見を寄せる", "これからの進め方と、皆さんに考えてほしいこと"); SEC = "04  意見を寄せる";
  // 16 ロードマップ
  {
    const s = content("今後の進め方｜検討ロードマップ（案）");
    const ph = [["R8 秋～冬", "現状・課題の共有", "住民意見の収集"], ["R9 前半", "保護者調査", "通学・施設の実態把握"], ["R9 後半", "複数案の比較", "中間報告"], ["R10 前半", "地区別対話", "地域の方向性の整理"], ["R10 後半以降", "地域提言", "市による検討・判断"]];
    ph.forEach(([d, a, b], i) => {
      const x = 0.6 + i * 2.44, now = i === 0;
      s.addShape(pres.shapes.CHEVRON, { x, y: 1.9, w: 2.5, h: 0.9, fill: { color: now ? K.amber : (i < 4 ? K.teal : K.navy) }, line: { type: "none" } });
      T(s, d, { x: x + 0.4, y: 1.9, w: 1.8, h: 0.9, fontSize: 17, bold: true, color: K.white, valign: "middle", align: "center" });
      box(s, x + 0.05, 3.05, 2.25, 2.7, now ? K.amberL : K.bg);
      if (now) T(s, "今ここ", { x: x + 0.3, y: 3.2, w: 1.5, h: 0.35, fontSize: 13, bold: true, color: K.amber });
      T(s, a, { x: x + 0.3, y: 3.65, w: 1.85, h: 0.9, fontSize: 19, bold: true, color: K.navy });
      T(s, b, { x: x + 0.3, y: 4.6, w: 1.85, h: 0.9, fontSize: 16, color: K.sub });
    });
    box(s, 0.6, 6.0, 12.1, 0.55, K.bg);
    T(s, "上記は提案スケジュールであり、正式決定ではありません。", { x: 0.9, y: 6.0, w: 11.6, h: 0.55, fontSize: 15, color: K.sub, valign: "middle" });
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
    T(s, "学校の数を減らすことではなく、未来の学びをつくること。", { x: 0.6, y: 5.45, w: 12.1, h: 0.8, fontSize: 26, bold: true, color: K.amber, align: "center", valign: "middle" });
  }

  await pres.writeFile({ fileName: OUT });
  if (applyTheme) await applyTheme(OUT, THEME);
  console.log("ok");
})();

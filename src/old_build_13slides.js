const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/mnt/skills/public/pptx/scripts/apply_theme.js");

const OUT = "/tmp/claude-0/-home-claude/11587814-b1d3-575f-a42d-8369509b2074/scratchpad/wakaho_setsumeikai.pptx";
const FONT = "BIZ UDPゴシック";
const THEME = {
  name: "Wakaho",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: "222222", lt1: "FFFFFF", dk2: "2E4A3A", lt2: "F1F3F2",
    accent1: "E1582B", accent2: "2B84C6", accent3: "2E4A3A", accent4: "E2B43B",
    accent5: "8A968F", accent6: "B23A2A", hlink: "2B84C6", folHlink: "6B4E9B",
  },
};
const HEX = { orange: "E1582B", blue: "2B84C6", green: "2E4A3A", yellow: "E2B43B", grey: "8A968F", light: "F1F3F2", ink: "222222", sub: "555F5A", line: "D5DAD7" };

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "若穂地区住民説明会資料（案）";
pres.author = "若穂地区小中学校再編検討委員会";
const C = pres.SchemeColor;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.background1 },
  objects: [],
});

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.9, 0.5, 0.45, 0.5],
  objects: [
    { rect: { x: 0, y: 0, w: 10, h: 0.78, fill: { color: C.accent1 } } },
    { rect: { x: 8.35, y: 0, w: 0.55, h: 0.78, fill: { color: C.accent4 } } },
    { rect: { x: 8.9, y: 0, w: 0.55, h: 0.78, fill: { color: C.accent3 } } },
    { rect: { x: 9.45, y: 0, w: 0.55, h: 0.78, fill: { color: C.accent2 } } },
    { text: { text: "若穂地区小中学校再編検討委員会", options: { x: 0.5, y: 5.28, w: 5, h: 0.25, fontSize: 9, color: C.accent5, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.45, y: 0.09, w: 7.8, h: 0.6, fontSize: 22, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 9.1, y: 5.24, w: 0.45, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
});

pres.defineSlideMaster({
  title: "CLOSING",
  background: { color: C.accent3 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.7, w: 8.8, h: 1.3, fontSize: 34, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});

// ---------- helpers ----------
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top" }, o));
const card = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill || C.background2 }, line: { type: "none" }, rectRadius: 0.08, objectName: name });
const circleNum = (s, x, y, n, color) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: 0.42, h: 0.42, fill: { color: color || C.accent1 }, line: { type: "none" } });
  T(s, String(n), { x, y, w: 0.42, h: 0.42, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
};
const content = (title, section) => {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  return s;
};

// ---------- data (R8.5.1現在の推計) ----------
const YE = ["R8", "R9", "R10", "R11", "R12", "R13", "R14"];
const school = {
  watanuchi: [273, 265, 256, 245, 241, 230, 205],
  kawada: [157, 142, 135, 130, 118, 101, 83],
  hoshina: [54, 56, 48, 49, 44, 42, 42],
};
const first = { watanuchi: [48, 40, 43, 34, 39, 26, 23], kawada: [22, 17, 17, 19, 10, 11, 9], hoshina: [2, 15, 3, 13, 2, 7, 2] };
const wakahoElem = YE.map((_, i) => school.watanuchi[i] + school.kawada[i] + school.hoshina[i]);
const firstTotal = YE.map((_, i) => first.watanuchi[i] + first.kawada[i] + first.hoshina[i]);
const cityElem = [15757, 15177, 14686, 14251, 13676, 13166, 12561];
const YM = ["R8", "R9", "R10", "R11", "R12", "R13", "R14", "R15", "R16", "R17", "R18", "R19", "R20"];
const wakahoJH = [284, 287, 266, 264, 248, 237, 225, 219, 206, 201, 180, 161, 129];
const wakahoJH1 = [85, 94, 87, 83, 78, 76, 71, 72, 63, 66, 51, 44, 34];
const cityJH = [7904, 7777, 7577, 7428, 7141, 6850, 6592, 6285, 6070, 5850, 5555, 5206, 4791];
const idx = a => a.map(v => Math.round((v / a[0]) * 1000) / 10);
const pct = a => Math.round((1 - a[a.length - 1] / a[0]) * 100);

// ===== 1 表紙 =====
pres.addSection({ title: "はじめに" });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "はじめに" });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 3.3, h: 5.625, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "deco-blue" });
  s.addShape(pres.shapes.OVAL, { x: 1.4, y: -1.2, w: 3.2, h: 3.2, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "deco-orange" });
  s.addShape(pres.shapes.OVAL, { x: 0.6, y: 1.6, w: 2.6, h: 2.6, fill: { color: C.accent4 }, line: { type: "none" }, objectName: "deco-yellow" });
  s.addShape(pres.shapes.OVAL, { x: -1.0, y: 3.4, w: 3.4, h: 3.4, fill: { color: C.accent3 }, line: { type: "none" }, objectName: "deco-green" });
  T(s, "若穂地区住民説明会資料", { x: 4.0, y: 1.45, w: 5.6, h: 0.7, fontSize: 32, bold: true });
  T(s, "これからの若穂の小中学校を、地域で考えるために（案）", { x: 4.0, y: 2.2, w: 5.6, h: 0.45, fontSize: 16, color: C.text2 });
  T(s, "R8.○.○", { x: 4.0, y: 3.2, w: 5.6, h: 0.35, fontSize: 14, color: C.accent5 });
  T(s, "若穂地区小中学校再編検討委員会\n長野市教育委員会事務局 学校教育課", { x: 4.0, y: 3.6, w: 5.6, h: 0.7, fontSize: 14 });
  s.addNotes("本日はお集まりいただきありがとうございます。若穂の小中学校のこれからについて、検討委員会の考えと進め方をご説明し、皆さんのご意見をお聞きします。");
}

// ===== 2 本日の目的と流れ =====
{
  const s = content("本日の説明会の目的と流れ", "はじめに");
  card(s, 0.5, 1.05, 4.2, 3.95, C.accent3, "purpose-card");
  T(s, "今日は「決める場」\nではありません", { x: 0.75, y: 1.3, w: 3.7, h: 0.9, fontSize: 20, bold: true, color: C.background1 });
  T(s, "若穂の子どもたちの数の見通しと、検討委員会の考え・進め方をお伝えし、皆さんのご意見をうかがう場です。", { x: 0.75, y: 2.35, w: 3.7, h: 1.4, fontSize: 15, color: C.background1, lineSpacingMultiple: 1.3 });
  const items = ["若穂の子どもの数の見通し", "なぜ今、再編を考えるのか", "考えられる選択肢", "今後の進め方", "質疑・意見交換"];
  items.forEach((t, i) => {
    const y = 1.1 + i * 0.78;
    circleNum(s, 5.1, y + 0.07, i + 1, i === 4 ? C.accent2 : C.accent1);
    T(s, t, { x: 5.7, y, w: 3.8, h: 0.56, fontSize: 17, valign: "middle" });
  });
  s.addNotes("冒頭で『今日は決めない』ことをはっきり伝えます。会場の緊張がやわらぎ、率直な意見が出やすくなります。");
}

// ===== 3 検討委員会について =====
{
  const s = content("若穂地区小中学校再編検討委員会について", "はじめに");
  card(s, 0.5, 1.05, 4.35, 3.95, null, "members-card");
  T(s, "委員の構成", { x: 0.75, y: 1.25, w: 3.9, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
  T(s, [
    { text: "綿内・川田・保科 各地区の代表 [○名]", options: { bullet: true, breakLine: true } },
    { text: "各小学校・若穂中のPTA代表 [○名]", options: { bullet: true, breakLine: true } },
    { text: "学校長 [○名]", options: { bullet: true, breakLine: true } },
    { text: "[未就学児の保護者など]", options: { bullet: true, breakLine: true } },
    { text: "事務局：長野市教育委員会 学校教育課", options: { bullet: true } },
  ], { x: 0.75, y: 1.8, w: 3.95, h: 3.0, fontSize: 14, paraSpaceAfter: 8 });
  T(s, "これまでの経緯", { x: 5.2, y: 1.25, w: 4.3, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
  const hist = [["[R8.○]", "検討委員会を設置"], ["[R8.○]", "児童・生徒数の推計を確認"], ["[R8.○]", "再編を考える理由と進め方を協議"], ["R8.○", "住民説明会（本日）"]];
  s.addShape(pres.shapes.LINE, { x: 5.35, y: 1.95, w: 0, h: 2.75, line: { color: HEX.line, width: 2 } });
  hist.forEach(([d, t], i) => {
    const y = 1.85 + i * 0.75;
    s.addShape(pres.shapes.OVAL, { x: 5.25, y: y + 0.08, w: 0.2, h: 0.2, fill: { color: i === 3 ? C.accent1 : C.accent5 }, line: { type: "none" } });
    T(s, d, { x: 5.6, y, w: 1.0, h: 0.36, fontSize: 14, color: C.accent5 });
    T(s, t, { x: 6.6, y, w: 2.9, h: 0.36, fontSize: 15, bold: i === 3 });
  });
  s.addNotes("［ ］の部分は委員会の実際の構成・日程に差し替えてください。3地区から偏りなく委員が出ていることを強調します。");
}

// ===== 4 児童数推計 =====
pres.addSection({ title: "子どもの数の見通し" });
{
  const s = content("児童数の見通し：3校とも減っていく", "子どもの数の見通し");
  const hdr = ["", ...YE].map((t, i) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.green }, align: i ? "center" : "left" } }));
  const row = (label, arr, opt = {}) => [{ text: label, options: Object.assign({ bold: true, align: "left" }, opt) }, ...arr.map(v => ({ text: String(v), options: Object.assign({ align: "center" }, opt) }))];
  const sum = { fill: { color: HEX.light }, bold: true };
  s.addTable([
    hdr,
    row("綿内小", school.watanuchi),
    row("川田小", school.kawada),
    row("保科小", school.hoshina),
    row("3校合計", wakahoElem, sum),
    row("新1年生計", firstTotal, { color: HEX.orange, bold: true }),
  ], { x: 0.5, y: 1.05, w: 6.1, colW: [1.55, 0.65, 0.65, 0.65, 0.65, 0.65, 0.65, 0.65], rowH: 0.48, fontSize: 13, fontFace: FONT, color: HEX.ink, border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle" });
  T(s, "単位：人。R8.5.1現在の推計（長野市教育委員会）", { x: 0.5, y: 4.05, w: 6.1, h: 0.3, fontSize: 10, color: C.accent5 });
  card(s, 6.95, 1.05, 2.55, 1.85, null, "callout-total");
  T(s, `−${pct(wakahoElem)}%`, { x: 7.15, y: 1.2, w: 2.2, h: 0.75, fontSize: 40, bold: true, color: C.accent1 });
  T(s, `3校の児童数\nR8 ${wakahoElem[0]}人 → R14 ${wakahoElem[6]}人`, { x: 7.15, y: 2.0, w: 2.2, h: 0.75, fontSize: 13, color: C.text1 });
  card(s, 6.95, 3.1, 2.55, 1.85, null, "callout-hoshina");
  T(s, "2人", { x: 7.15, y: 3.25, w: 2.2, h: 0.75, fontSize: 40, bold: true, color: C.accent1 });
  T(s, "保科小で新1年生が\n2人になる年\n（R8・R12・R14）", { x: 7.15, y: 4.05, w: 2.2, h: 0.75, fontSize: 13 });
  T(s, `若穂全体の新1年生は、R8の${firstTotal[0]}人からR14には${firstTotal[6]}人と半分以下になります。`, { x: 0.5, y: 4.45, w: 6.1, h: 0.6, fontSize: 14, bold: true, color: C.text2 });
  s.addNotes("小学校の推計は、すでに生まれている子どもの数をもとにしているため、大きくは外れません。川田小の新1年生もR12以降は10人前後になります。");
}

// ===== 5 生徒数推計 =====
{
  const s = content("生徒数の見通し：若穂中はR20に半分以下へ", "子どもの数の見通し");
  s.addChart(pres.charts.BAR, [{ name: "若穂中 全校生徒数", labels: YM, values: wakahoJH }], {
    x: 0.4, y: 1.0, w: 6.3, h: 3.95, barDir: "col", chartColors: [HEX.blue], barGapWidthPct: 45,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 10, dataLabelColor: HEX.ink, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.sub, catAxisLabelFontSize: 10, catAxisLabelFontFace: "+mn-lt",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMinVal: 0,
    showLegend: false, showTitle: true, title: "若穂中 全校生徒数の推計（人）", titleFontSize: 12, titleColor: HEX.ink, titleFontFace: "+mn-lt",
  });
  card(s, 6.95, 1.05, 2.55, 1.85, null, "callout-jh-total");
  T(s, `${wakahoJH[12]}人`, { x: 7.15, y: 1.2, w: 2.2, h: 0.75, fontSize: 40, bold: true, color: C.accent1 });
  T(s, `R20の全校生徒数\n（R8は${wakahoJH[0]}人）`, { x: 7.15, y: 2.0, w: 2.2, h: 0.75, fontSize: 13 });
  card(s, 6.95, 3.1, 2.55, 1.85, null, "callout-jh-first");
  T(s, `${wakahoJH1[12]}人`, { x: 7.15, y: 3.25, w: 2.2, h: 0.75, fontSize: 40, bold: true, color: C.accent1 });
  T(s, `R20の1年生\n（R8は${wakahoJH1[0]}人）`, { x: 7.15, y: 4.05, w: 2.2, h: 0.75, fontSize: 13 });
  T(s, "R8.5.1現在の推計（長野市教育委員会）", { x: 0.5, y: 4.95, w: 6, h: 0.25, fontSize: 10, color: C.accent5 });
  s.addNotes("中学校は小学校の減少が数年遅れて表れます。1学年34人は、今の小学校1校分ほどの規模です。部活動や行事にも影響します。");
}

// ===== 6 市全体との比較 =====
{
  const s = content("若穂は、市全体より速いペースで減っていく", "子どもの数の見通し");
  const lineOpts = (title, x) => ({
    x, y: 1.0, w: 4.4, h: 3.25, chartColors: [HEX.orange, HEX.grey], lineSize: 2.5, lineDataSymbolSize: 6,
    valAxisMinVal: 40, valAxisMaxVal: 110, valAxisMajorUnit: 20, valAxisLabelColor: HEX.sub, valAxisLabelFontSize: 10, valAxisLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.sub, catAxisLabelFontSize: 10, catAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: HEX.line, size: 0.75 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "t", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: HEX.ink,
    showTitle: true, title, titleFontSize: 13, titleColor: HEX.ink, titleFontFace: "+mn-lt",
  });
  s.addChart(pres.charts.LINE, [
    { name: "若穂（3小学校）", labels: YE, values: idx(wakahoElem) },
    { name: "長野市全体", labels: YE, values: idx(cityElem) },
  ], lineOpts("小学生（R8＝100）", 0.4));
  s.addChart(pres.charts.LINE, [
    { name: "若穂中", labels: YM, values: idx(wakahoJH) },
    { name: "長野市全体", labels: YM, values: idx(cityJH) },
  ], lineOpts("中学生（R8＝100）", 5.2));
  card(s, 0.5, 4.35, 4.2, 0.7, null, "elem-compare");
  T(s, [
    { text: "R14  若穂 ", options: {} },
    { text: `${Math.round(idx(wakahoElem)[6])}`, options: { bold: true, color: C.accent1, fontSize: 20 } },
    { text: `　／　市全体 ${Math.round(idx(cityElem)[6])}`, options: {} },
  ], { x: 0.7, y: 4.4, w: 3.9, h: 0.6, fontSize: 14, valign: "middle" });
  card(s, 5.3, 4.35, 4.2, 0.7, null, "jh-compare");
  T(s, [
    { text: "R20  若穂 ", options: {} },
    { text: `${Math.round(idx(wakahoJH)[12])}`, options: { bold: true, color: C.accent1, fontSize: 20 } },
    { text: `　／　市全体 ${Math.round(idx(cityJH)[12])}`, options: {} },
  ], { x: 5.5, y: 4.4, w: 3.9, h: 0.6, fontSize: 14, valign: "middle" });
  T(s, "R8年度の人数を100とした指数。配布資料の推計（R8.5.1現在）から作成", { x: 0.5, y: 5.05, w: 8, h: 0.22, fontSize: 9, color: C.accent5 });
  s.addNotes("少子化は市全体の課題ですが、若穂はそれより速いペースで進みます。『市全体の話ではなく、若穂の話』として受け止めていただきたいページです。");
}

// ===== 7 なぜ今① =====
pres.addSection({ title: "なぜ今、考えるのか" });
{
  const s = content("なぜ今？① 子どもの数から見えてくること", "なぜ今、考えるのか");
  const cards = [
    ["学級の人数", "川田小の新1年生はR12以降10人前後。保科小は年によって2人〜15人と大きく変わります。", "話し合いや集団での活動が組みにくくなる"],
    ["クラス替え", "川田小・保科小は各学年1学級。6年間、同じ顔ぶれで過ごします。", "人間関係が固定し、こじれたときの逃げ場が少ない"],
    ["中学校の部活動", `若穂中の1年生はR8の${wakahoJH1[0]}人からR20には${wakahoJH1[12]}人に。`, "チームが組めない部活動が出てくるおそれ"],
  ];
  cards.forEach(([h, fact, imp], i) => {
    const x = 0.5 + i * 3.07;
    card(s, x, 1.05, 2.86, 3.25, null, `card-${i}`);
    T(s, h, { x: x + 0.22, y: 1.25, w: 2.42, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    T(s, fact, { x: x + 0.22, y: 1.8, w: 2.42, h: 1.45, fontSize: 14, lineSpacingMultiple: 1.2 });
    T(s, imp, { x: x + 0.22, y: 3.35, w: 2.42, h: 0.8, fontSize: 14, bold: true, color: C.accent1 });
  });
  T(s, "一方で、小規模校には「一人ひとりに目が届く」「全員に出番がある」良さもあります。良さを生かしながら考えます。", { x: 0.5, y: 4.45, w: 9, h: 0.6, fontSize: 14, color: C.text2, bold: true });
  s.addNotes("課題だけでなく小規模校の良さも必ず伝え、『統合ありき』と受け取られないようにします。");
}

// ===== 8 なぜ今② =====
{
  const s = content("なぜ今？② 少子化とは別の理由", "なぜ今、考えるのか");
  const rows = [
    ["校舎の改修・更新の時期が重なる", "川田小（1973年度築）は市の計画上の更新時期を過ぎています。若穂中はR9、綿内小・保科小はR12に改修・更新の時期を迎えます。"],
    ["中学入学時の「環境の段差」", "R8の6年生は保科小13人・川田小32人・綿内小48人。全員が若穂中の約90人の学年に入ります。"],
    ["同じ若穂でも学ぶ環境が違う", "綿内小は各学年2学級、川田小・保科小は1学級。住む地区によって学びの条件が変わります。"],
    ["9年間を見通した学び", "若穂の小学生は全員が若穂中に進みます。小中が連携した学びをつくりやすい地域です。"],
  ];
  rows.forEach(([h, d], i) => {
    const y = 1.0 + i * 0.98;
    card(s, 0.5, y, 9.0, 0.86, i % 2 ? C.background1 : C.background2, `reason-${i}`);
    circleNum(s, 0.7, y + 0.22, i + 1, C.accent1);
    T(s, h, { x: 1.3, y: y + 0.1, w: 2.75, h: 0.66, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    T(s, d, { x: 4.15, y: y + 0.08, w: 5.2, h: 0.7, fontSize: 13, valign: "middle", lineSpacingMultiple: 1.1 });
  });
  T(s, "①は長野市学校施設長寿命化計画（令和2年4月時点）による。最新の予定は事務局で確認中", { x: 0.5, y: 4.98, w: 9, h: 0.25, fontSize: 9, color: C.accent5 });
  s.addNotes("既存資料P4の『少子化とは別に再編を考える理由』への案です。どれを示すかは委員会で確定してください。①は大きな工事の前に学校の形を決めておく、という説明になります。若穂中は昭和36〜37年に3中学校が統合してできた学校で、中学校はすでに若穂でひとつです。");
}

// ===== 9 委員会の考え =====
{
  const s = content("検討委員会として大切にしたいこと（案）", "なぜ今、考えるのか");
  const items = [
    ["子どもの最善を第一に", "学びと育ちの環境を、何よりも優先して考えます。"],
    ["結論ありきで進めない", "現状維持も含め、すべての選択肢を同じ物差しで比べます。"],
    ["3地区の声を等しく聞く", "綿内・川田・保科、それぞれの地区で意見を聞く場を設けます。"],
    ["検討の過程を公開する", "資料や会議の要旨は、回覧や広報で皆さんにお知らせします。"],
  ];
  items.forEach(([h, d], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.05 + Math.floor(i / 2) * 1.95;
    card(s, x, y, 4.4, 1.75, null, `principle-${i}`);
    circleNum(s, x + 0.25, y + 0.25, i + 1, C.accent3);
    T(s, h, { x: x + 0.85, y: y + 0.25, w: 3.35, h: 0.45, fontSize: 18, bold: true, color: C.text2, valign: "middle" });
    T(s, d, { x: x + 0.85, y: y + 0.85, w: 3.35, h: 0.75, fontSize: 14, lineSpacingMultiple: 1.2 });
  });
  T(s, "※委員会で協議のうえ確定", { x: 0.5, y: 4.98, w: 9, h: 0.25, fontSize: 9, color: C.accent5 });
  s.addNotes("既存資料P4（委員会としての考え・方向性）に対応するページです。文言は委員会で確定させてください。");
}

// ===== 10 選択肢 =====
pres.addSection({ title: "選択肢" });
{
  const s = content("考えられる選択肢", "選択肢");
  const H = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.green } } });
  const B = (t) => ({ text: t, options: { bold: true, color: HEX.green } });
  s.addTable([
    [H("案"), H("内容"), H("主な良い点"), H("主な留意点")],
    [B("現状維持＋魅力化"), "3校を残し、合同授業・行事やICTで学びを補う", "通学や地域の拠点が今のまま", "少人数化の根本的な解決にはならない"],
    [B("一部統合"), "小規模化が著しい学校を、近隣校と統合", "影響を抑えつつ、課題の大きい所から対応", "組み合わせ方、将来の再検討の可能性"],
    [B("3小学校の統合"), "若穂で1つの小学校にする", "学年複数学級・クラス替えが可能に", "通学距離、スクールバス、跡地の活用"],
    [B("義務教育学校など"), "若穂中と一体の小中一貫の学校にする", "9年間を見通した教育ができる", "校舎整備の規模が大きく、時間がかかる"],
  ], { x: 0.5, y: 1.05, w: 9.0, colW: [1.9, 2.6, 2.25, 2.25], rowH: [0.45, 0.8, 0.8, 0.8, 0.8], fontSize: 13, fontFace: FONT, color: HEX.ink, border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle" });
  T(s, "いずれかに決まっているものではありません。比べるための「たたき台」です。", { x: 0.5, y: 4.85, w: 9, h: 0.35, fontSize: 14, bold: true, color: C.accent1 });
  s.addNotes("既存資料P5・P6の内容と重ならないよう整理してください。会場から新しい案が出たら歓迎する姿勢を示します。");
}

// ===== 11 義務教育学校 =====
{
  const s = content("義務教育学校について", "選択肢");
  card(s, 0.5, 1.05, 4.4, 2.3, null, "merit");
  T(s, "主なメリット（一般論）", { x: 0.7, y: 1.2, w: 4.0, h: 0.4, fontSize: 17, bold: true, color: C.accent2 });
  T(s, [
    { text: "9年間を見通した柔軟なカリキュラム", options: { bullet: true, breakLine: true } },
    { text: "「中1ギャップ」の緩和", options: { bullet: true, breakLine: true } },
    { text: "異学年交流の活性化", options: { bullet: true, breakLine: true } },
    { text: "小・中の先生の連携強化", options: { bullet: true } },
  ], { x: 0.7, y: 1.7, w: 4.0, h: 2.0, fontSize: 15, paraSpaceAfter: 8 });
  card(s, 5.1, 1.05, 4.4, 2.3, null, "caution");
  T(s, "気をつけるポイント（一般論）", { x: 5.3, y: 1.2, w: 4.0, h: 0.4, fontSize: 17, bold: true, color: C.accent1 });
  T(s, [
    { text: "人間関係が固定しやすい", options: { bullet: true, breakLine: true } },
    { text: "上級生が低学年に与える影響への配慮", options: { bullet: true, breakLine: true } },
    { text: "通学距離の負担", options: { bullet: true } },
  ], { x: 5.3, y: 1.7, w: 4.0, h: 2.0, fontSize: 15, paraSpaceAfter: 8 });
  card(s, 0.5, 3.6, 9.0, 1.3, C.accent3, "wakaho-points");
  T(s, "若穂で考えるなら", { x: 0.7, y: 3.8, w: 2.0, h: 0.9, fontSize: 16, bold: true, color: C.background1, valign: "middle" });
  T(s, "中学入学時の「段差」をなくせる一方、校地の場所によって通学距離が大きく変わります。通学方法や3小学校の跡地の活用も、あわせて検討が必要です。", { x: 2.75, y: 3.8, w: 6.6, h: 0.9, fontSize: 15, color: C.background1, valign: "middle" });
  s.addNotes("既存資料P7の要点です。詳しい説明は配布資料をご覧いただく形にします。選択肢の一覧（前のページ）の後に示すことで、『義務教育学校に決まっている』と受け取られないようにします。校地や開校年度など、委員会で決めていない具体案は示しません。");
}

// ===== 12 今後の進め方 =====
pres.addSection({ title: "今後の進め方" });
{
  const s = content("今後の進め方（目安）", "今後の進め方");
  const steps = [
    ["R8 秋〜冬", "住民説明会（本日）", "進め方を共有し、ご意見をうかがう"],
    ["R9 前半", "現状把握・アンケート", "保護者（未就学児を含む）へのアンケート、先行事例の視察"],
    ["R9 後半", "選択肢の比較・中間報告", "同じ物差しで比較し、地区ごとの懇談会で報告"],
    ["R10 前半", "地域の方向性をまとめる", "説明会と意見募集を経て、方向性案をまとめる"],
    ["R10 後半〜", "長野市へ提言・市の検討", "意見書を提出し、市が方針を決める"],
    ["その後", "移行準備（再編する場合）", "校名・通学方法・跡地活用などを協議"],
  ];
  s.addShape(pres.shapes.LINE, { x: 2.0, y: 1.2, w: 0, h: 3.6, line: { color: HEX.line, width: 2 } });
  steps.forEach(([d, h, t], i) => {
    const y = 1.0 + i * 0.66, now = i === 0;
    T(s, d, { x: 0.5, y, w: 1.3, h: 0.5, fontSize: 14, bold: true, color: now ? C.accent1 : C.accent5, valign: "middle" });
    s.addShape(pres.shapes.OVAL, { x: 1.89, y: y + 0.14, w: 0.22, h: 0.22, fill: { color: now ? C.accent1 : C.accent3 }, line: { type: "none" } });
    if (now) card(s, 2.3, y - 0.02, 7.2, 0.54, C.background2, "now-row");
    T(s, h, { x: 2.45, y, w: 2.8, h: 0.5, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    T(s, t, { x: 5.3, y, w: 4.15, h: 0.5, fontSize: 13, valign: "middle" });
  });
  T(s, "時期は目安です。市の手続きの日程によって前後します。", { x: 0.5, y: 4.98, w: 9, h: 0.25, fontSize: 9, color: C.accent5 });
  s.addNotes("交流学習など、どの案になっても役立つ取り組みは結論を待たずに始めることも検討します。");
}

// ===== 13 ご意見のお願い =====
{
  const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "今後の進め方" });
  s.addText("皆さんのご意見を\nお聞かせください", { placeholder: "title" });
  const cols = [["会場で", "お配りした意見シートにご記入ください"], ["後日", "[提出先・メールアドレス・QRコード]\n[締切：R8.○.○]"], ["お問い合わせ", "長野市教育委員会\n学校教育課 [電話番号]"]];
  cols.forEach(([h, d], i) => {
    const x = 0.6 + i * 3.0;
    s.addShape(pres.shapes.LINE, { x, y: 2.85, w: 2.6, h: 0, line: { color: HEX.yellow, width: 3 } });
    T(s, h, { x, y: 3.05, w: 2.6, h: 0.45, fontSize: 18, bold: true, color: C.background1 });
    T(s, d, { x, y: 3.6, w: 2.6, h: 1.1, fontSize: 14, color: C.background1, lineSpacingMultiple: 1.2 });
  });
  s.addNotes("質疑の時間を十分にとります。出た意見は次回までに回答をまとめ、公開します。");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

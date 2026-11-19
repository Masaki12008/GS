// 簡潔版（11ページ）の生成。前半の設定・部品は build_deck.js と同じ（build_deck.js の設定を変えたら合わせる）
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const fs = require("fs");
const path = require("path");
let applyTheme = null;
try { ({ applyTheme } = require(process.env.APPLY_THEME || "/mnt/skills/public/pptx/scripts/apply_theme.js")); } catch (e) { /* optional: theme colors only */ }

const OUT = process.argv[2] || "output/若穂地区_住民説明会資料_簡潔版.pptx";
const APPENDIX = process.env.APPENDIX !== "0"; // APPENDIX=0 で参考資料を除いた投影用を作る
const SHORT = process.env.SHORT === "1"; // SHORT=1 で配布用の短縮版（区切りのページ・Q1〜Q4のページを外し、用語の説明を1枚に）
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
// 出典などの注記。行数から高さを決め、ページ下の行（y=7.05）に重ならないように上へずらす
const noteLines = text => text.split("\n").reduce((a, l) => a + Math.max(1, Math.ceil([...l].reduce((w, c) => w + (c.charCodeAt(0) < 0x2000 ? 0.55 : 1), 0) * 13 / (12.1 * 72 * 0.93))), 0);
const note = (s, text, y = 6.55) => { const h = noteLines(text) * 13 * 1.25 / 72; T(s, text, { x: 0.6, y: Math.min(y, 6.98 - h), w: 12.1, h, fontSize: 13, color: K.sub }); };
let SEC = "";
const content = (title, subtitle) => { const s = pres.addSlide({ masterName: "CONTENT" }); if (SEC) T(s, SEC, { x: 0.6, y: 0.2, w: 6, h: 0.3, fontSize: 14, bold: true, color: K.amber }); s.addText(title, { placeholder: "title" }); if (subtitle) sub(s, subtitle); return s; };


const hills = (s) => {
  s.addShape(pres.shapes.OVAL, { x: -2.5, y: 5.6, w: 9.5, h: 4.2, fill: { color: K.hill1 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: 5.2, y: 6.0, w: 10.5, h: 4.0, fill: { color: K.hill2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: 2.0, y: 6.6, w: 8.0, h: 3.0, fill: { color: K.hill3 }, line: { type: "none" } });
};
const divider = (n, name, desc) => {
  if (SHORT) return;
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


// ===== 簡潔版（11ページ）：話の筋だけを大きな字で。本文はおおむね20pt以上 =====
const BT = (title, subtitle) => { const s = pres.addSlide({ masterName: "CONTENT" }); s.addText(title, { placeholder: "title", fontSize: 34 }); if (subtitle) T(s, subtitle, { x: 0.6, y: 1.15, w: 12.1, h: 0.45, fontSize: 20, color: K.sub }); return s; };
const src = (s, text) => T(s, text, { x: 0.6, y: text.includes("\n") ? 6.4 : 6.55, w: 12.1, h: text.includes("\n") ? 0.6 : 0.4, fontSize: 14, color: K.sub });

(async () => {
  // 1 表紙
  {
    const s = pres.addSlide({ masterName: "COVER" });
    hills(s);
    T(s, "若穂地区 小中学校再編検討｜住民説明会資料（簡潔版・案）", { x: 0.9, y: 0.7, w: 11, h: 0.4, fontSize: 18, color: "CFE0D6" });
    T(s, "これからの若穂の学校を考える", { x: 0.9, y: 1.5, w: 11.5, h: 1.1, fontSize: 52, bold: true, color: K.white });
    T(s, "結論の前に、子どもたちにどんな学びを残したいか。", { x: 0.9, y: 2.75, w: 11.5, h: 0.7, fontSize: 30, bold: true, color: K.amber });
    T(s, "令和8年（2026年）［　］月［　］日\n若穂地区小中学校再編検討委員会／長野市教育委員会事務局 学校教育課", { x: 0.9, y: 3.85, w: 11.5, h: 0.9, fontSize: 18, color: "CFE0D6" });
  }

  // 2 今日は決める場ではない
  {
    const s = BT("今日は「決める場」ではありません");
    const pts = [["1", "子どもの数の見通しを知る"], ["2", "考えられる選択肢を知る"], ["3", "皆さんの意見を聞かせてください"]];
    pts.forEach(([n, t], i) => {
      const y = 1.6 + i * 1.25;
      box(s, 0.6, y, 12.1, 1.05, K.bg);
      s.addShape(pres.shapes.OVAL, { x: 0.9, y: y + 0.17, w: 0.72, h: 0.72, fill: { color: K.teal }, line: { type: "none" } });
      T(s, n, { x: 0.9, y: y + 0.17, w: 0.72, h: 0.72, fontSize: 28, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, t, { x: 1.95, y, w: 10.5, h: 1.05, fontSize: 30, bold: true, color: K.navy, valign: "middle" });
    });
    box(s, 0.6, 5.5, 12.1, 0.85, K.amberL);
    T(s, "学校の形も、時期も、まだ何も決まっていません。", { x: 0.9, y: 5.5, w: 11.6, h: 0.85, fontSize: 26, bold: true, color: "A3402C", valign: "middle" });
  }

  // 3 子どもの数
  {
    const s = BT("子どもの数は、これからも減る見通し");
    const big = [["3つの小学校の合計", tot[0], tot[6], "R8（2026）→ R14（2032）", "約3割減"], ["若穂中学校", jh[0], jh[12], "R8（2026）→ R20（2038）", "約5割減"]];
    big.forEach(([h, a, b, yr, d], i) => {
      const x = 0.6 + i * 6.15;
      box(s, x, 1.5, 5.95, 2.75, K.tealL);
      T(s, h, { x: x + 0.35, y: 1.65, w: 5.3, h: 0.5, fontSize: 24, bold: true, color: K.teal });
      T(s, [{ text: String(a), options: { color: K.grey } }, { text: "人", options: { color: K.grey, fontSize: 24 } }, { text: " → ", options: { color: K.grey, fontSize: 28 } }, { text: String(b), options: { color: K.navy } }, { text: "人", options: { color: K.navy, fontSize: 24 } }], { x: x + 0.35, y: 2.2, w: 5.4, h: 1.0, fontSize: 44, bold: true, valign: "middle" });
      T(s, `${yr}\n${d}`, { x: x + 0.35, y: 3.25, w: 5.4, h: 0.85, fontSize: 20, bold: true, color: K.ink });
    });
    T(s, "学校ごとでは（R8 → R14）", { x: 0.6, y: 4.45, w: 12.1, h: 0.45, fontSize: 22, bold: true, color: K.navy });
    Object.entries(sch).forEach(([n, v], i) => {
      const x = 0.6 + i * 4.1;
      box(s, x, 4.95, 3.9, 1.3, K.bg);
      T(s, n, { x: x + 0.3, y: 5.0, w: 3.3, h: 0.5, fontSize: 22, bold: true, color: K.navy });
      T(s, `${v[0]}人 → ${v[6]}人`, { x: x + 0.2, y: 5.5, w: 3.6, h: 0.6, fontSize: 24, bold: true, color: K.ink });
    });
    src(s, "出典：長野市教育委員会の推計（令和8年（2026年）5月1日現在）。転入・転出で変わります。");
  }

  // 4 小規模校の良さと課題
  {
    const s = BT("小さな学校の良さと、考えておきたいこと");
    const cols = [["大切にしたい良さ", K.teal, K.tealL, ["一人ひとりに目が届く", "学年をこえた交流", "地域と学校が近い", "安心できる人間関係"]],
      ["考えておきたいこと", K.amber, K.amberL, ["多様な意見にふれる機会", "集団での活動や行事", "人間関係が固定しやすい", "先生の数や活動の続けやすさ"]]];
    cols.forEach(([h, c, cl, items], i) => {
      const x = 0.6 + i * 6.15;
      box(s, x, 1.5, 5.95, 4.6, cl);
      T(s, h, { x: x + 0.4, y: 1.7, w: 5.2, h: 0.6, fontSize: 28, bold: true, color: c });
      items.forEach((t, j) => {
        const y = 2.55 + j * 0.85;
        s.addShape(pres.shapes.OVAL, { x: x + 0.45, y: y + 0.2, w: 0.22, h: 0.22, fill: { color: c }, line: { type: "none" } });
        T(s, t, { x: x + 0.85, y, w: 4.9, h: 0.62, fontSize: 24, color: K.ink, valign: "middle" });
      });
    });
    T(s, "学校の大きさだけで、教育の良し悪しは決まりません。", { x: 0.6, y: 6.25, w: 12.1, h: 0.5, fontSize: 22, bold: true, color: K.navy, align: "center" });
  }

  // 5 なぜ今
  {
    const s = BT("なぜ今、考えるのか");
    const rs = [["子どもの数が減り続ける", "小学生はR14（2032）までに約3割、\n中学生はR20（2038）までに約5割減る見通し", "MdTrendingDown"],
      ["校舎の改修の時期と重なる", "4校とも、市の計画で\n改修などを考える時期に入ります", "MdApartment"],
      ["調べて話し合うには時間がかかる", "通学や地域への影響を調べ、\n地区ごとに話し合うため", "MdSchedule"]];
    for (let i = 0; i < 3; i++) {
      const [h, d, ic] = rs[i], y = 1.5 + i * 1.65;
      box(s, 0.6, y, 12.1, 1.5, K.bg);
      await iconCircle(s, 0.95, y + 0.3, 0.9, ic);
      T(s, h, { x: 2.2, y: y + 0.12, w: 10.2, h: 0.6, fontSize: 28, bold: true, color: K.navy });
      T(s, d, { x: 2.2, y: y + 0.72, w: 10.2, h: 0.75, fontSize: 20, color: K.ink });
    }
  }

  // 6 選択肢
  {
    const s = BT("考えられる選択肢", "どれかに決まっているわけではありません。組み合わせることもできます");
    const o = [["A", "現状維持", "3小学校と\n若穂中を\n今の形で"], ["B", "一部の学校を統合", "2校の統合など"], ["C", "小学校3校を統合", "小学校を1校に"],
      ["D", "義務教育学校", "小中9年間を\n一つの学校で"], ["E", "小中一貫型の\n小学校・中学校", "別々の学校の\nまま、9年間の\n学びをそろえる"]];
    const W = 2.3, G = 0.15;
    [["学校の配置", 0, 3], ["学校の制度", 3, 2]].forEach(([h, k, n]) => {
      const x = 0.6 + k * (W + G), w = n * W + (n - 1) * G;
      T(s, h, { x, y: 1.75, w, h: 0.45, fontSize: 20, bold: true, color: K.teal, align: "center" });
      s.addShape(pres.shapes.LINE, { x, y: 2.22, w, h: 0, line: { color: K.teal, width: 2 } });
    });
    o.forEach(([L, h, d], i) => {
      const x = 0.6 + i * (W + G);
      box(s, x, 2.35, W, 3.85, K.bg);
      s.addShape(pres.shapes.OVAL, { x: x + 0.75, y: 2.55, w: 0.8, h: 0.8, fill: { color: K.navy }, line: { type: "none" } });
      T(s, L, { x: x + 0.75, y: 2.55, w: 0.8, h: 0.8, fontSize: 30, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, h, { x: x + 0.12, y: 3.5, w: W - 0.24, h: 1.0, fontSize: 22, bold: true, color: K.navy, align: "center", valign: "middle" });
      T(s, d, { x: x + 0.12, y: 4.6, w: W - 0.24, h: 1.2, fontSize: 18, color: K.ink, align: "center" });
    });
    src(s, "D（義務教育学校）とE（小中一貫型の小学校・中学校）は、法律上の区分が違う別の制度です。");
  }

  // 7 5つの視点
  {
    const s = BT("どの選択肢も、同じ物差しで比べます");
    const v = [["学び", "いろいろな学び\nと交流", "MdMenuBook"], ["安心", "一人ひとりへの\n支え", "MdFavorite"], ["安全", "無理のない通学\nと防災", "MdHealthAndSafety"], ["地域", "行事や\nつながり", "MdGroups"], ["続けやすさ", "校舎・先生・\n費用", "MdAutorenew"]];
    for (let i = 0; i < 5; i++) {
      const [h, d, ic] = v[i], x = 0.6 + i * 2.46;
      box(s, x, 1.55, 2.26, 3.9, i % 2 ? K.bg : K.tealL);
      await iconCircle(s, x + 0.58, 1.85, 1.1, ic);
      T(s, h, { x: x + 0.1, y: 3.1, w: 2.06, h: 0.65, fontSize: 28, bold: true, color: K.navy, align: "center" });
      T(s, d, { x: x + 0.1, y: 3.85, w: 2.06, h: 1.2, fontSize: 20, color: K.ink, align: "center" });
    }
    box(s, 0.6, 5.7, 12.1, 0.8, K.bg);
    T(s, "比べた結果は、令和9年度（2027年度）に調べたうえでお示しします。", { x: 0.9, y: 5.7, w: 11.6, h: 0.8, fontSize: 22, bold: true, color: K.navy, valign: "middle" });
  }

  // 8 学校の間の距離
  {
    const s = BT("学校と学校の距離と時間", "学校と学校の間の目安です（各家庭からの距離ではありません）");
    const H = t => ({ text: t, options: { bold: true, color: K.white, fill: { color: K.navy }, align: "center", fontSize: 20 } });
    const legs = [["綿内小 ⇔ 川田小", "3.2km", "44分", "7〜8分"], ["川田小 ⇔ 保科小", "3.0km", "46分", "6分"], ["綿内小 ⇔ 保科小", "4.9km", "72分", "10〜12分"],
      ["綿内小 ⇔ 若穂中", "2.5km", "33分", "6分"], ["川田小 ⇔ 若穂中", "1.1km", "16分", "4分"], ["保科小 ⇔ 若穂中", "2.5km", "39分", "5〜6分"]];
    const rows = [[H("区間"), H("歩く道のり"), H("歩く時間"), H("車（朝8時）")]];
    legs.forEach((r, i) => rows.push(r.map((t, j) => ({ text: t, options: { align: j ? "center" : "left", bold: j === 0, fill: { color: i % 2 ? K.white : K.bg } } }))));
    s.addTable(rows, { x: 0.6, y: 1.75, w: 12.1, colW: [4.0, 2.7, 2.7, 2.7], rowH: 0.58, fontSize: 22, fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } });
    T(s, "今は、各地区の子どもの9割以上が、今の小学校まで2km未満です。", { x: 0.6, y: 5.95, w: 12.1, h: 0.5, fontSize: 20, bold: true, color: K.navy });
    src(s, "Googleマップの経路検索（徒歩は大人の速さ。令和8年（2026年）10〜11月）。保科小は、ほかの学校より約100m高い所にあります。");
  }

  // 9 生活への影響
  {
    const s = BT("どの選択肢でも、生活への影響を確かめます");
    const v = [["通学", "時間・安全・冬の道・スクールバス", "MdDirectionsBus"], ["送り迎え", "保護者の負担、学校周りの安全", "MdDirectionsCar"],
      ["放課後", "子どもプラザ・児童センターの場所", "MdChildFriendly"], ["学校生活", "行事・学級の人数・友だち関係", "MdSchool"],
      ["支えが必要な子", "特別支援学級・通級・居場所", "MdAccessibility"], ["地域", "地区の行事・防災・つながり", "MdHome"]];
    for (let i = 0; i < v.length; i++) {
      const [h, d, ic] = v[i], y = 1.45 + i * 0.85;
      box(s, 0.6, y, 12.1, 0.75, i % 2 ? K.white : K.bg, i % 2 ? K.line : null);
      await iconCircle(s, 0.8, y + 0.1, 0.55, ic);
      T(s, h, { x: 1.6, y, w: 3.4, h: 0.75, fontSize: 24, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: 5.0, y, w: 7.6, h: 0.75, fontSize: 22, valign: "middle" });
    }
    src(s, "子どもプラザ：小学校の校舎内など学校施設内にある、放課後の「子どもの居場所」。");
  }

  // 10 これからの進め方
  {
    const s = BT("これからの進め方（案）");
    const ph = [["R7（2025）\n年度", "アンケート", "済み"], ["R8（2026）\n年度後半", "地区ごとの説明会", "今ここ"], ["R9（2027）\n年度", "意見の集約・調査", ""], ["R9（2027）\n年度中", "委員会の方向性", "目標"]];
    const W = 3.0;
    ph.forEach(([d, a, tag], i) => {
      const x = 0.6 + i * (W + 0.03), now = tag === "今ここ";
      s.addShape(pres.shapes.CHEVRON, { x, y: 1.5, w: W + 0.1, h: 0.85, fill: { color: now ? K.amber : (tag === "目標" ? K.navy : K.teal) }, line: { type: "none" } });
      T(s, d, { x: x + 0.35, y: 1.5, w: W - 0.6, h: 0.85, fontSize: 17, bold: true, color: K.white, align: "center", valign: "middle" });
      box(s, x + 0.05, 2.5, W - 0.15, 1.45, now ? K.amberL : K.bg);
      if (tag) T(s, tag, { x: x + 0.25, y: 2.58, w: 2.4, h: 0.4, fontSize: 18, bold: true, color: now ? K.amber : K.teal });
      T(s, a, { x: x + 0.25, y: 3.0, w: W - 0.45, h: 0.8, fontSize: 22, bold: true, color: K.navy, valign: "middle" });
    });
    T(s, "R10（2028）年度以降は、進み具合に応じて見直します。", { x: 0.6, y: 4.05, w: 12.1, h: 0.45, fontSize: 18, color: K.sub });
    T(s, "皆さんの意見が長野市に届くまで", { x: 0.6, y: 4.6, w: 12.1, h: 0.5, fontSize: 22, bold: true, color: K.navy });
    const who = ["検討委員会", "住民自治協議会", "長野市に要望"];
    who.forEach((h, i) => {
      const x = 0.6 + i * 4.1;
      box(s, x, 5.2, 3.75, 1.0, K.tealL);
      T(s, h, { x, y: 5.2, w: 3.75, h: 1.0, fontSize: 26, bold: true, color: K.teal, align: "center", valign: "middle" });
      if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 3.8, y: 5.5, w: 0.25, h: 0.4, fill: { color: K.teal }, line: { type: "none" } });
    });
    src(s, "検討委員会＝若穂地区小中学校再編検討委員会。住民自治協議会＝若穂地区住民自治協議会。\nスケジュールは目標で、正式決定ではありません。");
  }

  // 11 ご意見を
  {
    const s = BT("ご意見をお寄せください", "賛成・反対だけでなく、心配なことや知りたいことも大歓迎です");
    const q = ["子どもたちに、どんな学びを経験してほしいですか？", "今の学校の良さで、残したいものは何ですか？", "通学・安全・地域活動で、心配なことは何ですか？", "判断の前に、どんな情報が必要ですか？"];
    q.forEach((t, i) => {
      const y = 1.75 + i * 0.85;
      box(s, 0.6, y, 12.1, 0.75, K.bg);
      T(s, `Q${i + 1}`, { x: 0.85, y, w: 0.9, h: 0.75, fontSize: 26, bold: true, color: K.amber, valign: "middle" });
      T(s, t, { x: 1.8, y, w: 10.7, h: 0.75, fontSize: 24, color: K.ink, valign: "middle" });
    });
    box(s, 0.6, 5.3, 12.1, 1.1, K.tealL);
    T(s, "締切：令和8年（2026年）［　］月［　］日\n提出先：会場の回収箱、または［　　　　］　問い合わせ：長野市教育委員会事務局 学校教育課", { x: 0.9, y: 5.3, w: 11.6, h: 1.1, fontSize: 20, bold: true, color: K.navy, valign: "middle" });
  }

  await pres.writeFile({ fileName: OUT });
  if (applyTheme) await applyTheme(OUT, THEME);
  await require("./fix_fonts")(OUT, FONT);
  require("./resave")(OUT);
  console.log("ok");
})();

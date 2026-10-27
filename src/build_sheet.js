// 説明会で配る意見用紙（A4たて・2枚：大人用、小中学生用）。P27の問い（Q1〜Q4）と同じ言葉で作る。
const pptxgen = require("pptxgenjs");

const OUT = process.argv[2] || "output/意見用紙.pptx";
const FONT = "BIZ UDPゴシック";
const K = { navy: "1D4636", teal: "2E7D5B", tealL: "E8F2EC", ink: "1E2A26", sub: "56635D", line: "B9C4BE", white: "FFFFFF" };

const pres = new pptxgen();
pres.defineLayout({ name: "A4P", width: 8.27, height: 11.69 });
pres.layout = "A4P";
pres.title = "若穂地区 住民説明会 意見用紙";

const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 12, color: K.ink, valign: "top", fontFace: FONT }, o));
const X = 0.6, W = 7.07;

// □ の選択肢を1行に並べる
const checks = (s, y, label, items) => {
  T(s, label, { x: X, y, w: 1.5, h: 0.32, fontSize: 12, bold: true, color: K.navy, valign: "middle" });
  T(s, items.map(t => `□ ${t}`).join("　"), { x: X + 1.5, y, w: W - 1.5, h: 0.32, fontSize: 12, valign: "middle" });
};
// 問いと、書き込む枠
const question = (s, y, h, q, text) => {
  s.addShape(pres.shapes.RECTANGLE, { x: X, y, w: W, h, fill: { color: K.white }, line: { color: K.line, width: 1 } });
  s.addShape(pres.shapes.RECTANGLE, { x: X, y, w: W, h: 0.42, fill: { color: K.tealL }, line: { color: K.line, width: 1 } });
  T(s, q, { x: X + 0.15, y, w: 0.6, h: 0.42, fontSize: 14, bold: true, color: K.teal, valign: "middle" });
  T(s, text, { x: X + 0.75, y, w: W - 0.9, h: 0.42, fontSize: 13, bold: true, color: K.navy, valign: "middle" });
  for (let ly = y + 0.42 + 0.42; ly < y + h - 0.05; ly += 0.42) s.addShape(pres.shapes.LINE, { x: X + 0.15, y: ly, w: W - 0.3, h: 0, line: { color: "DCE3DF", width: 0.75, dashType: "dash" } });
};
const header = (s, title, lead) => {
  T(s, "若穂地区 小中学校再編検討｜住民説明会", { x: X, y: 0.45, w: W, h: 0.3, fontSize: 11, color: K.sub });
  T(s, title, { x: X, y: 0.75, w: W, h: 0.55, fontSize: 22, bold: true, color: K.navy });
  T(s, lead, { x: X, y: 1.35, w: W, h: 0.5, fontSize: 12, color: K.sub });
};
const footer = (s, text) => {
  s.addShape(pres.shapes.RECTANGLE, { x: X, y: 10.55, w: W, h: 0.7, fill: { color: K.tealL }, line: { type: "none" } });
  T(s, text, { x: X + 0.15, y: 10.55, w: W - 0.3, h: 0.7, fontSize: 11, color: K.ink, valign: "middle" });
};

// 1枚目：大人用
{
  const s = pres.addSlide();
  header(s, "若穂の学校についての ご意見用紙", "賛成・反対だけでなく、心配なこと・知りたいことも、どうぞお書きください。当てはまるものに✓を付けてください。");
  checks(s, 1.95, "お住まいの地区", ["綿内", "川田", "保科", "その他"]);
  checks(s, 2.3, "お立場（複数可）", ["未就学児の保護者", "小学生の保護者", "中学生の保護者"]);
  checks(s, 2.62, "", ["地域の住民", "学校・保育の関係者", "その他（　　　　）"]);
  checks(s, 2.97, "年代", ["10代", "20代", "30代", "40代"]);
  checks(s, 3.29, "", ["50代", "60代", "70代以上"]);
  const qs = ["子どもたちに、どんな学びを経験してほしいですか？", "今の学校の良さで、残したいものは何ですか？", "通学・安全・地域活動で心配なことは何ですか？", "判断の前に、どんな情報が必要ですか？", "そのほか、ご意見・ご質問をお書きください。"];
  qs.forEach((q, i) => question(s, 3.75 + i * 1.34, 1.24, `Q${i + 1}`, q));
  footer(s, "締切：令和8年（2026年）［　］月［　］日　提出先：［　　　　　　］\nインターネット：［URL・QRコード］　お名前などの個人情報は公表しません。");
}

// 2枚目：小中学生用（やさしい言葉）
{
  const s = pres.addSlide();
  header(s, "わかほの学校について、みんなの考えを教えてね", "小学生・中学生のみなさんへ。思ったことを、そのまま書いてください。絵でもかまいません。");
  checks(s, 1.95, "学校", ["綿内小", "川田小", "保科小", "若穂中"]);
  checks(s, 2.3, "学年", ["1年", "2年", "3年", "4年", "5年", "6年"]);
  const qs = ["学校で、楽しいこと・好きなことは何ですか？", "今の学校で、これからも残してほしいものは何ですか？", "ほかの学校の友だちと、いっしょにやってみたいことはありますか？", "学校に行くときや、学校のことで、心配なことはありますか？", "ほかに、大人に伝えたいことがあれば書いてください。"];
  qs.forEach((q, i) => question(s, 2.8 + i * 1.5, 1.4, `${i + 1}`, q));
  footer(s, "書いた紙は、［先生・回収箱］にわたしてください。名前は書かなくてよいです。\nみなさんの考えは、これからの学校を考える大人たちが、ぜんぶ読みます。");
}

pres.writeFile({ fileName: OUT }).then(() => console.log("ok"));

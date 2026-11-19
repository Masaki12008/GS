// 説明会のあとに配るアンケート案（A4たて3枚：大人用2枚＝両面1枚、小中学生用1枚）。
// ✓を付けるだけで答えられる問いを中心にし、自由記述は最後に1つ。選択肢A〜Eの賛否や順位は聞かない（「統合ありき」に見せないため）。
const pptxgen = require("pptxgenjs");

const OUT = process.argv[2] || "output/アンケート案.pptx";
const FONT = "BIZ UDPゴシック";
const K = { navy: "1D4636", teal: "2E7D5B", tealL: "E8F2EC", ink: "1E2A26", sub: "56635D", line: "B9C4BE", white: "FFFFFF" };

const pres = new pptxgen();
pres.defineLayout({ name: "A4P", width: 8.27, height: 11.69 });
pres.layout = "A4P";
pres.title = "若穂地区 住民説明会 アンケート（案）";

const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: K.ink, valign: "top", fontFace: FONT }, o));
const X = 0.6, W = 7.07, RH = 0.4;

const header = (s, title, lead, page) => {
  T(s, "若穂地区 小中学校再編検討｜住民説明会　アンケート（案）", { x: X, y: 0.4, w: W - 1, h: 0.3, fontSize: 11, color: K.sub });
  if (page) T(s, page, { x: X + W - 1, y: 0.4, w: 1, h: 0.3, fontSize: 11, color: K.sub, align: "right" });
  T(s, title, { x: X, y: 0.72, w: W, h: 0.5, fontSize: 22, bold: true, color: K.navy });
  if (lead) T(s, lead, { x: X, y: 1.28, w: W, h: 0.55, fontSize: 12, color: K.sub });
};
// 問いの見出し（番号と問い）。補足は右に小さく
const qhead = (s, y, num, text, hint) => {
  s.addShape(pres.shapes.RECTANGLE, { x: X, y, w: W, h: 0.46, fill: { color: K.tealL }, line: { type: "none" } });
  T(s, num, { x: X + 0.12, y, w: 0.7, h: 0.46, fontSize: 15, bold: true, color: K.teal, valign: "middle" });
  T(s, text, { x: X + 0.8, y, w: W - 0.9, h: 0.46, fontSize: 15, bold: true, color: K.navy, valign: "middle" });
  if (hint) T(s, hint, { x: X + 0.8, y, w: W - 0.95, h: 0.46, fontSize: 12, color: K.sub, valign: "middle", align: "right" });
  return y + 0.52;
};
// □ の選択肢を cols 列に並べる。戻り値は次の y
const grid = (s, y, items, cols, o = {}) => {
  const x0 = X + (o.indent ?? 0.25), cw = (W - (o.indent ?? 0.25)) / cols;
  items.forEach((t, i) => T(s, `□ ${t}`, { x: x0 + (i % cols) * cw, y: y + Math.floor(i / cols) * RH, w: cw - 0.05, h: RH, valign: "middle", fontSize: o.size || 14 }));
  return y + Math.ceil(items.length / cols) * RH;
};
// ラベルつきの1行（属性）
const labeled = (s, y, label, items, cols) => {
  T(s, label, { x: X + 0.25, y, w: 1.45, h: RH, bold: true, color: K.navy, valign: "middle" });
  const cw = (W - 1.75) / cols;
  items.forEach((t, i) => T(s, `□ ${t}`, { x: X + 1.75 + (i % cols) * cw, y: y + Math.floor(i / cols) * RH, w: cw - 0.05, h: RH, valign: "middle" }));
  return y + Math.ceil(items.length / cols) * RH;
};
const writeBox = (s, y, h) => {
  s.addShape(pres.shapes.RECTANGLE, { x: X, y, w: W, h, fill: { color: K.white }, line: { color: K.line, width: 1 } });
  for (let ly = y + 0.45; ly < y + h - 0.1; ly += 0.45) s.addShape(pres.shapes.LINE, { x: X + 0.15, y: ly, w: W - 0.3, h: 0, line: { color: "DCE3DF", width: 0.75, dashType: "dash" } });
  return y + h;
};
const footer = (s, text, y = 10.6, h = 0.7) => {
  s.addShape(pres.shapes.RECTANGLE, { x: X, y, w: W, h, fill: { color: K.tealL }, line: { type: "none" } });
  T(s, text, { x: X + 0.15, y, w: W - 0.3, h, fontSize: 12, color: K.ink, valign: "middle" });
};

// 1枚目：大人用（表）
{
  const s = pres.addSlide();
  header(s, "若穂の学校についてのアンケート", "□に✓を付けてください。お名前は不要です。答えは地区ごとに集計して公表します。", "1／2");
  let y = qhead(s, 1.72, "問1", "あなたについて教えてください", "お立場はいくつでも");
  y = labeled(s, y, "お住まいの地区", ["綿内", "川田", "保科", "若穂以外"], 4);
  y = labeled(s, y, "お立場", ["未就学児の保護者", "小学生の保護者", "中学生の保護者", "学校・保育の関係者", "地域の住民", "その他"], 2);
  y = labeled(s, y, "年代", ["10代", "20代", "30代", "40代", "50代", "60代", "70代", "80代以上"], 4);
  y = qhead(s, y + 0.08, "問2", "今日の説明は分かりやすかったですか", "1つに✓");
  y = grid(s, y, ["よく分かった", "だいたい分かった", "あまり分からなかった", "分からなかった"], 2);
  y = qhead(s, y + 0.08, "問3", "学校のこれからを考えるとき、特に大切だと思うことは", "3つまで");
  y = grid(s, y, ["学び（いろいろな学びと交流）", "安心（一人ひとりへの支え）", "安全（無理のない通学と防災）", "地域（行事やつながり）", "続けやすさ（校舎・先生・費用）", "その他（　　　　　　　　）"], 2);
  y = qhead(s, y + 0.08, "問4", "気になること・心配なことは", "いくつでも");
  y = grid(s, y, ["通学の道のり・時間・安全", "家族の送迎の負担", "放課後の居場所（子どもプラザ）", "学年・学級の人数", "友だち関係・学校生活", "支援が必要な子への配慮",
    "地域の行事・つながり", "災害のときの避難所", "校舎の古さ・直す費用", "先生の数・配置", "特にない", "その他（　　　　　　　　）"], 2);
  footer(s, "裏面に続きます　→", 11.0, 0.4);
}

// 2枚目：大人用（裏）
{
  const s = pres.addSlide();
  header(s, "若穂の学校についてのアンケート（続き）", "", "2／2");
  let y = qhead(s, 1.4, "問5", "もっと詳しく知りたい選択肢は", "いくつでも");
  T(s, "どれがよいかを選ぶ問いではありません。今後の説明の参考にします。", { x: X + 0.25, y, w: W - 0.25, h: 0.32, fontSize: 12, color: K.sub, valign: "middle" });
  y = grid(s, y + 0.36, ["A　現状維持", "B　一部の学校を統合", "C　小学校3校を統合", "D　義務教育学校", "E　小中一貫型の小学校・中学校", "選択肢の組み合わせ"], 2);
  y = qhead(s, y + 0.15, "問6", "もっと知りたい情報は", "いくつでも");
  y = grid(s, y, ["家から学校までの距離・時間", "通学の方法（バスなど）", "校舎の状態と工事の見通し", "選択肢ごとの費用", "ほかの地域の例", "子どもや先生の声",
    "決まるまでの手続きと時期", "その他（　　　　　　　　）"], 2);
  y = qhead(s, y + 0.15, "問7", "これから意見を伝えやすい方法は", "いくつでも");
  y = grid(s, y, ["地区ごとの説明会", "少人数の話し合いの会", "アンケート（紙）", "アンケート（インターネット）", "学校・地区の役員を通して", "その他（　　　　　　　　）"], 2);
  y = qhead(s, y + 0.15, "問8", "そのほか、ご意見・ご質問を自由にお書きください");
  writeBox(s, y, 10.4 - y);
  footer(s, "締切：令和8年（2026年）［　］月［　］日　　提出先：［回収箱・学校・支所など］\nインターネットでも答えられます：［URL・QRコード］　　お問い合わせ：［　　　　　　］", 10.5, 0.75);
}

// 3枚目：小中学生用（やさしい言葉）
{
  const s = pres.addSlide();
  header(s, "わかほの学校について、みんなの考えを教えてね", "小学生・中学生のみなさんへ。あてはまる□に✓をつけてください。名前は書かなくてよいです。");
  let y = labeled(s, 1.95, "学校", ["綿内小", "川田小", "保科小", "若穂中"], 4);
  y = labeled(s, y, "学年", ["1年", "2年", "3年", "4年", "5年", "6年"], 6);
  y = qhead(s, y + 0.15, "1", "学校で、好きなことは？", "いくつでも");
  y = grid(s, y, ["友だちと遊ぶこと", "授業", "先生", "運動会などの行事", "給食", "自然の中の活動", "地域の人との活動", "その他（　　　　　　）"], 2);
  y = qhead(s, y + 0.15, "2", "ほかの学校の友だちと、いっしょにやってみたいことは？", "いくつでも");
  y = grid(s, y, ["授業", "運動会・音楽会", "遠足・校外学習", "オンラインで話す", "特にない", "その他（　　　　　　）"], 2);
  y = qhead(s, y + 0.15, "3", "学校のことで、心配なことや こまっていることは？", "いくつでも");
  y = grid(s, y, ["学校までの道", "友だちのこと", "クラスの人数", "勉強のこと", "特にない", "その他（　　　　　　）"], 2);
  y = qhead(s, y + 0.15, "4", "これからの学校について、自由に書いてね（絵でもいいよ）");
  writeBox(s, y, 10.4 - y);
  footer(s, "書いた紙は、［先生・回収箱］にわたしてください。\nみなさんの考えは、これからの学校を考える大人たちが、ぜんぶ読みます。", 10.5, 0.75);
}

pres.writeFile({ fileName: OUT }).then(() => require("./fix_fonts")(OUT, FONT)).then(() => require("./resave")(OUT)).then(() => console.log("ok"));

// 地域の説明会テンプレート集（販売用）を3色ぶん作る。
// 使い方：cd 副業 && NODE_PATH=../node_modules node src/build_templates.js
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const fs = require("fs");
const path = require("path");

const FONT = "BIZ UDPゴシック";
const OUTDIR = path.join(__dirname, "../テンプレート");

const mix = (a, b, t) => [0, 2, 4].map(i => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t).toString(16).padStart(2, "0")).join("").toUpperCase();
const THEMES = {
  みどり: { navy: "1D4636", main: "2E7D5B", acc: "D9573E" },
  あお: { navy: "1B3A5C", main: "2F6FA8", acc: "C9721A" },
  えんじ: { navy: "4A2330", main: "9C3D54", acc: "2E7D7A" },
};
const palette = c => Object.assign({}, c, {
  mainL: mix(c.main, "FFFFFF", 0.88), accL: mix(c.acc, "FFFFFF", 0.88), accD: mix(c.acc, "000000", 0.2),
  ink: "1E2A26", sub: "56635D", line: "D7DFDA", bg: "F4F6F5", white: "FFFFFF", grey: "8E9893",
  soft: mix(c.navy, "FFFFFF", 0.75),
  h1: mix(c.navy, c.main, 0.3), h2: mix(c.navy, c.main, 0.6), h3: mix(c.navy, c.main, 0.9),
});

const iconCache = {};
async function icon(name, color) {
  const key = name + color;
  if (!iconCache[key]) {
    const svg = RDS.renderToStaticMarkup(React.createElement(md[name], { color: "#" + color, size: 256 }));
    iconCache[key] = "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
  }
  return iconCache[key];
}

// ───────── 説明会資料（16:9） ─────────
async function buildDeck(themeName) {
  const K = palette(THEMES[themeName]);
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "地域の説明会資料テンプレート";
  pres.author = "［作成者名］";
  const FOOT = "［団体名］｜［会の名前］資料（案）";

  pres.defineSlideMaster({ title: "COVER", background: { color: K.navy }, objects: [] });
  pres.defineSlideMaster({
    title: "CONTENT", background: { color: K.white },
    objects: [
      { text: { text: FOOT, options: { x: 0.6, y: 7.05, w: 7, h: 0.25, fontSize: 10, color: K.grey, margin: 0, fontFace: FONT } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.7, fontSize: 30, bold: true, color: K.navy, valign: "middle", align: "left", margin: 0, fontFace: FONT }, text: "" } },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: K.grey, align: "right" },
  });

  const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 16, color: K.ink, valign: "top", fontFace: FONT }, o));
  const box = (s, x, y, w, h, fill, line) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: line ? { color: line, width: 1 } : { type: "none" }, rectRadius: 0.12 });
  const note = (s, text, y = 6.55) => T(s, text, { x: 0.6, y, w: 12.1, h: 0.35, fontSize: 13, color: K.sub });
  let SEC = "";
  const content = (title, subtitle, notes) => {
    const s = pres.addSlide({ masterName: "CONTENT" });
    if (SEC) T(s, SEC, { x: 0.6, y: 0.18, w: 6, h: 0.3, fontSize: 14, bold: true, color: K.acc });
    s.addText(title, { placeholder: "title" });
    if (subtitle) T(s, subtitle, { x: 0.6, y: 1.12, w: 12.1, h: 0.4, fontSize: 16, color: K.sub });
    if (notes) s.addNotes(notes);
    return s;
  };
  const hills = s => {
    s.addShape(pres.shapes.OVAL, { x: -2.5, y: 5.6, w: 9.5, h: 4.2, fill: { color: K.h1 }, line: { type: "none" } });
    s.addShape(pres.shapes.OVAL, { x: 5.2, y: 6.0, w: 10.5, h: 4.0, fill: { color: K.h2 }, line: { type: "none" } });
    s.addShape(pres.shapes.OVAL, { x: 2.0, y: 6.6, w: 8.0, h: 3.0, fill: { color: K.h3 }, line: { type: "none" } });
  };
  const divider = (n, name, desc) => {
    const s = pres.addSlide({ masterName: "COVER" });
    hills(s);
    T(s, n, { x: 0.9, y: 1.2, w: 3, h: 1.6, fontSize: 96, bold: true, color: K.acc });
    T(s, name, { x: 0.9, y: 2.95, w: 11, h: 1.0, fontSize: 48, bold: true, color: K.white });
    T(s, desc, { x: 0.9, y: 4.05, w: 11, h: 0.6, fontSize: 20, color: K.soft });
    s.addNotes("区切りのページ。章の名前と一言の説明を書き換えてください。章が少ない場合は削除してかまいません。");
    return s;
  };
  async function iconCircle(s, x, y, d, name, bg = K.main, fg = K.white) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } });
    const p = d * 0.22;
    s.addImage({ data: await icon(name, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  }
  const tableOpt = (o) => Object.assign({ fontFace: FONT, color: K.ink, valign: "middle", border: { type: "solid", pt: 1, color: K.line } }, o);
  const th = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true, color: K.white, fill: { color: K.main }, align: "center" }, o) });

  // 1 表紙
  {
    const s = pres.addSlide({ masterName: "COVER" });
    hills(s);
    T(s, "［団体名］｜［会の名前］資料（案）", { x: 0.9, y: 0.7, w: 9, h: 0.4, fontSize: 15, color: K.soft });
    T(s, "［〇〇のこれからを考える］", { x: 0.9, y: 1.5, w: 11.5, h: 1.1, fontSize: 50, bold: true, color: K.white });
    T(s, "［結論の前に、大切にしたいことを一緒に。］", { x: 0.9, y: 2.75, w: 11.5, h: 0.7, fontSize: 28, bold: true, color: K.acc });
    T(s, "［説明会の目的を一言で書きます］", { x: 0.9, y: 3.55, w: 11, h: 0.5, fontSize: 18, color: K.soft });
    T(s, "令和［ ］年（20［ ］年）［ ］月［ ］日　［主催者名］", { x: 0.9, y: 4.45, w: 11.5, h: 0.4, fontSize: 14, color: K.soft });
    s.addNotes("表紙。［ ］の部分を書き換えてください。大きな題は20字以内、キャッチコピーは25字以内が目安です。\n書き方のコツ：題に結論を書かない（例：「〇〇の廃止について」ではなく「〇〇のこれからを考える」）。");
  }

  SEC = "はじめに";
  // 2 今日の流れ
  {
    const s = content("今日の説明会で共有したいこと", null, "全体の流れを最初に見せると、参加者が安心して聞けます。4つの段階の名前は、内容に合わせて変えてください。\n下の帯は「今日は決める場ではない」ことを伝えるためのものです。決定事項を説明する会の場合は書き換えてください。");
    const steps = [["現状を知る", "数字と事実の共有", "MdInsights"], ["課題を共有する", "困っていること・良いところ", "MdForum"], ["選択肢を比べる", "今の形を続ける案も含めて", "MdCompareArrows"], ["意見を寄せる", "大切にしたいこと・心配なこと", "MdRateReview"]];
    for (let i = 0; i < 4; i++) {
      const [h, d, ic] = steps[i], x = 0.6 + i * 3.08;
      box(s, x, 1.7, 2.82, 3.9, K.bg);
      T(s, `0${i + 1}`, { x: x + 0.3, y: 1.95, w: 1.2, h: 0.5, fontSize: 22, bold: true, color: K.acc });
      await iconCircle(s, x + 0.3, 2.6, 1.0, ic);
      T(s, h, { x: x + 0.3, y: 3.85, w: 2.3, h: 0.5, fontSize: 22, bold: true, color: K.navy });
      T(s, d, { x: x + 0.3, y: 4.45, w: 2.3, h: 0.9, fontSize: 16, color: K.sub });
    }
    box(s, 0.6, 5.9, 12.1, 0.6, K.accL);
    T(s, "今日は決める場ではありません。比べるための情報を共有し、皆さんのご意見をうかがいます。", { x: 0.85, y: 5.9, w: 11.7, h: 0.6, fontSize: 16, bold: true, color: K.accD, valign: "middle" });
    note(s, "年は「令和8年度（2026年度）」のように併記しています。専門用語の説明は、巻末の参考資料にあります。", 6.62);
  }

  // 3 出発点
  {
    const s = content("検討の出発点｜誰のための検討か", null, "検討の目的と、大切にする考え方を示すページ。一番上の帯に「最優先にすること」を1つだけ書きます。");
    box(s, 0.6, 1.6, 12.1, 1.7, K.main);
    await iconCircle(s, 1.0, 1.9, 1.1, "MdFavorite", K.white, K.main);
    T(s, "［最優先にすることを一文で］", { x: 2.4, y: 1.6, w: 10, h: 1.7, fontSize: 32, bold: true, color: K.white, valign: "middle" });
    const pts = [["［大切にする考え方 1］\n例：今の良さを残す", "MdThumbUp"], ["［大切にする考え方 2］\n例：安全とつながりを守る", "MdShield"], ["［大切にする考え方 3］\n例：比べる情報をそろえる", "MdFactCheck"]];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.1;
      box(s, x, 3.6, 3.9, 2.8, K.bg);
      await iconCircle(s, x + 0.35, 3.9, 0.8, pts[i][1], K.mainL, K.main);
      T(s, pts[i][0], { x: x + 0.35, y: 4.9, w: 3.2, h: 1.3, fontSize: 19, bold: true, color: K.navy, lineSpacingMultiple: 1.15 });
    }
  }

  divider("01", "現状を知る", "［数字と事実で、今の状況を共有します］");
  SEC = "01　現状を知る";

  // 5 グラフ
  const yrs = ["R8", "R9", "R10", "R11", "R12", "R13", "R14"];
  const sample = [[120, 116, 110, 104, 99, 93, 86], [64, 60, 58, 53, 50, 47, 44], [38, 37, 35, 34, 31, 30, 28]];
  {
    const s = content("［数の見通し］", "［何の数か・単位］（［出典］、［時点］現在）", "グラフの数値はサンプルです。グラフを右クリック →「データの編集」で表の数字を差し替えてください。\n複数の地区・施設を並べるときは、同じ色の濃淡で示し、特定の1つだけを赤などで目立たせないのがコツです。\n数値には必ず出典と時点を書きます。");
    s.addChart(pres.charts.BAR, ["［地区A］", "［地区B］", "［地区C］"].map((n, i) => ({ name: n, labels: yrs.map(r => `${r}\n${2018 + +r.slice(1)}`), values: sample[i] })), {
      x: 0.6, y: 1.6, w: 8.2, h: 4.9, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
      chartColors: [K.navy, K.main, K.h3 === K.main ? K.soft : mix(K.main, "FFFFFF", 0.45)],
      showLegend: true, legendPos: "t", legendFontSize: 14, legendFontFace: FONT,
      catAxisLabelFontSize: 13, valAxisLabelFontSize: 13, catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT,
      catAxisLabelColor: K.sub, valAxisLabelColor: K.sub, valGridLine: { color: "E5E9E7", size: 0.75 }, catGridLine: { style: "none" },
      showValue: false,
    });
    box(s, 9.1, 1.6, 3.6, 4.9, K.bg);
    T(s, "読み取れること", { x: 9.35, y: 1.8, w: 3.2, h: 0.4, fontSize: 18, bold: true, color: K.navy });
    [["［地区A］", "120 → 86（−34）"], ["［地区B］", "64 → 44（−20）"], ["［地区C］", "38 → 28（−10）"]].forEach(([a, b], j) => {
      T(s, a, { x: 9.35, y: 2.35 + j * 0.85, w: 3.2, h: 0.35, fontSize: 16, bold: true, color: K.main });
      T(s, b, { x: 9.35, y: 2.7 + j * 0.85, w: 3.2, h: 0.4, fontSize: 18 });
    });
    T(s, "［一言でまとめ：例 どの地区も減る見通し］", { x: 9.35, y: 5.0, w: 3.2, h: 1.2, fontSize: 17, bold: true, color: K.navy });
    note(s, "※グラフはサンプルの数値です。年は上段が令和、下段が西暦です。");
  }

  // 6 数字のまとめ
  {
    const s = content("数字で見るポイント", "［3つの数字で要点を伝えます］", "大きな数字を3つだけ選んで示すページ。数字の下に「何の数か」と「出典・時点」を書いてください。");
    const st = [["−32%", "［合計の減り方］", "［222人 → 158人（R8→R14）］"], ["［ ］校", "［対象の施設数］", "［例：すべて築40年以上］"], ["［ ］年", "［次の節目まで］", "［例：改修の対象になる時期］"]];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.1;
      box(s, x, 1.7, 3.9, 4.4, i === 0 ? K.mainL : K.bg);
      T(s, st[i][0], { x: x + 0.3, y: 2.1, w: 3.3, h: 1.4, fontSize: 60, bold: true, color: K.main, align: "center", valign: "middle" });
      T(s, st[i][1], { x: x + 0.3, y: 3.75, w: 3.3, h: 0.5, fontSize: 21, bold: true, color: K.navy, align: "center" });
      T(s, st[i][2], { x: x + 0.3, y: 4.4, w: 3.3, h: 1.2, fontSize: 16, color: K.sub, align: "center" });
    }
    note(s, "出典：［資料名］（［発行者］、［時点］）");
  }

  divider("02", "課題を共有する", "［良いところと、困っていることを並べます］");
  SEC = "02　課題を共有する";

  // 8 良さと課題
  {
    const s = content("今の良さと、これからの課題", "どちらか一方だけでなく、両方を同じ大きさで示します", "良い点と課題を同じ数・同じ大きさで並べると、「結論ありき」に見えにくくなります。\n具体的な良さは、地域や関係者から聞き取った言葉を使うと説得力が増します。");
    const cols = [["今の良さ", K.main, K.mainL, "MdThumbUp", ["［良さ 1：例 顔の見える関係］", "［良さ 2：例 地域行事とのつながり］", "［良さ 3：例 一人ひとりに目が届く］", "［良さ 4］"]],
      ["これからの課題", K.acc, K.accL, "MdWarningAmber", ["［課題 1：例 担い手が足りない］", "［課題 2：例 建物の老朽化］", "［課題 3：例 費用の負担］", "［課題 4］"]]];
    for (let i = 0; i < 2; i++) {
      const [h, c, cl, ic, items] = cols[i], x = 0.6 + i * 6.15;
      box(s, x, 1.65, 5.95, 4.7, cl);
      await iconCircle(s, x + 0.4, 1.95, 0.75, ic, c);
      T(s, h, { x: x + 1.35, y: 1.95, w: 4.3, h: 0.75, fontSize: 24, bold: true, color: K.navy, valign: "middle" });
      items.forEach((t, j) => {
        const y = 3.0 + j * 0.75;
        s.addShape(pres.shapes.OVAL, { x: x + 0.55, y: y + 0.17, w: 0.18, h: 0.18, fill: { color: c }, line: { type: "none" } });
        T(s, t, { x: x + 0.95, y, w: 4.8, h: 0.52, fontSize: 19, valign: "middle" });
      });
    }
  }

  // 9 なぜ今
  {
    const s = content("なぜ今、考えるのか", null, "「今すぐ決めなければならない」と急がせるのではなく、「準備に時間がかかるので、今から話し始める」という書き方がおすすめです。");
    const items = [["MdSchedule", "［理由 1］", "例：準備や話し合いに数年かかるため、早めに始める必要がある"], ["MdApartment", "［理由 2］", "例：建物の改修の時期が近づいている（計画上の時期）"], ["MdGroups", "［理由 3］", "例：今の子どもたちの声を聞ける時間を確保したい"]];
    for (let i = 0; i < 3; i++) {
      const [ic, h, d] = items[i], y = 1.5 + i * 1.6;
      box(s, 0.6, y, 12.1, 1.4, i % 2 ? K.white : K.bg, i % 2 ? K.line : null);
      await iconCircle(s, 0.9, y + 0.25, 0.9, ic);
      T(s, h, { x: 2.1, y: y + 0.18, w: 10.3, h: 0.5, fontSize: 22, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: 2.1, y: y + 0.72, w: 10.3, h: 0.5, fontSize: 17, color: K.ink, valign: "middle" });
    }
    note(s, "※計画上の時期は「決まったこと」ではありません。「〇年度から改修の対象」のように書きます。");
  }

  divider("03", "選択肢を比べる", "［今の形を続ける案も含めて、同じ物差しで並べます］");
  SEC = "03　選択肢を比べる";

  // 11 選択肢
  {
    const s = content("考えられる選択肢", "どれかに決まっているわけではありません", "選択肢は、今の形を続ける案を必ず最初に入れます。案の数は2〜5つが目安です。使わない枠は削除してください。");
    const ops = [["A", "［今の形を続ける］", "［内容の説明］"], ["B", "［一部を見直す］", "［内容の説明］"], ["C", "［まとめる］", "［内容の説明］"], ["D", "［新しい形にする］", "［内容の説明］"]];
    for (let i = 0; i < 4; i++) {
      const [k, h, d] = ops[i], x = 0.6 + i * 3.08;
      box(s, x, 1.7, 2.82, 4.6, K.bg);
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: 1.95, w: 0.9, h: 0.9, fill: { color: K.main }, line: { type: "none" } });
      T(s, k, { x: x + 0.3, y: 1.95, w: 0.9, h: 0.9, fontSize: 32, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, h, { x: x + 0.25, y: 3.1, w: 2.4, h: 0.9, fontSize: 19, bold: true, color: K.navy });
      T(s, d, { x: x + 0.3, y: 4.05, w: 2.3, h: 2.0, fontSize: 16, color: K.sub });
    }
    note(s, "すべての案を同じ色・同じ大きさで示しています。");
  }

  // 12 比較表
  {
    const s = content("同じ物差しで比べる", "評価は、調査が終わってから記入します", "比較表は、評価欄を空けた状態で見せるのがおすすめです（「これから一緒に評価する」ことが伝わります）。\n評価を入れる場合は「良い・課題あり・調査中」など、記号ではなく言葉で書くと誤解が少なくなります。");
    const rows = [[th("視点"), th("確かめること"), th("A"), th("B"), th("C"), th("D")]];
    [["［視点 1］", "例：子どもや利用者にとっての良さ"], ["［視点 2］", "例：安全・移動の負担"], ["［視点 3］", "例：地域とのつながり"], ["［視点 4］", "例：建物と費用"], ["［視点 5］", "例：実現までの期間"]].forEach(([a, b], i) => {
      const f = { fill: { color: i % 2 ? K.white : K.bg } };
      rows.push([{ text: a, options: Object.assign({ bold: true, color: K.navy }, f) }, { text: b, options: f }, ...[0, 1, 2, 3].map(() => ({ text: "", options: f }))]);
    });
    s.addTable(rows, tableOpt({ x: 0.6, y: 1.7, w: 12.1, colW: [2.1, 4.4, 1.4, 1.4, 1.4, 1.4], rowH: 0.75, fontSize: 17 }));
    note(s, "評価の付け方（案）：［誰が］［何をもとに］視点ごとに記入し、［いつ］お示しします。");
  }

  // 13 生活への影響
  {
    const s = content("暮らしへの影響", "どの選択肢でも、確かめておきたいこと", "参加者が一番気にするのは「自分の生活がどう変わるか」です。項目は内容に合わせて入れ替えてください。");
    const it = [["MdDirectionsWalk", "移動・通学", "［距離・時間・安全］"], ["MdDirectionsCar", "送り迎え", "［家族の負担］"], ["MdAccessTime", "放課後・時間", "［預かり・活動の時間］"], ["MdAccessible", "支援が必要な人", "［配慮の方法］"], ["MdGroups", "地域の行事", "［続け方］"], ["MdHome", "建物と費用", "［改修・維持の費用］"]];
    for (let i = 0; i < 6; i++) {
      const x = 0.6 + (i % 3) * 4.1, y = 1.7 + Math.floor(i / 3) * 2.45;
      box(s, x, y, 3.9, 2.25, K.bg);
      await iconCircle(s, x + 0.3, y + 0.3, 0.75, it[i][0], K.mainL, K.main);
      T(s, it[i][1], { x: x + 1.25, y: y + 0.3, w: 2.5, h: 0.75, fontSize: 21, bold: true, color: K.navy, valign: "middle" });
      T(s, it[i][2], { x: x + 0.3, y: y + 1.25, w: 3.3, h: 0.8, fontSize: 17, color: K.ink });
    }
  }

  divider("04", "意見を寄せる", "［これからの進め方と、ご意見の出し方］");
  SEC = "04　意見を寄せる";

  // 15 ロードマップ
  {
    const s = content("これからの進め方（案）", null, "「今ここ」の位置を書き換えてください。時期は「案」であり、決定ではないことを下に書いておきます。\n下の段「誰が決めるのか」は、参加者からよく出る質問への答えになります。");
    const ph = [["R8（2026）\n秋～冬", "現状の共有", "意見を集める"], ["R9（2027）\n前半", "関係者の声を聞く", "［対象］"], ["R9（2027）\n後半", "案を比べる", "中間報告"], ["R10（2028）\n前半", "地区ごとの対話", "方向性の整理"], ["R10（2028）\n後半以降", "とりまとめ", "［提出先］へ"]];
    ph.forEach(([d, a, b], i) => {
      const x = 0.6 + i * 2.44, now = i === 0;
      s.addShape(pres.shapes.CHEVRON, { x, y: 1.5, w: 2.5, h: 0.9, fill: { color: now ? K.acc : (i < 4 ? K.main : K.navy) }, line: { type: "none" } });
      T(s, d, { x: x + 0.4, y: 1.5, w: 1.8, h: 0.9, fontSize: 16, bold: true, color: K.white, valign: "middle", align: "center" });
      box(s, x + 0.05, 2.55, 2.25, 2.0, now ? K.accL : K.bg);
      if (now) T(s, "今ここ", { x: x + 0.3, y: 2.65, w: 1.5, h: 0.3, fontSize: 14, bold: true, color: K.acc });
      T(s, a, { x: x + 0.3, y: 2.95, w: 1.85, h: 0.75, fontSize: 18, bold: true, color: K.navy });
      T(s, b, { x: x + 0.3, y: 3.75, w: 1.85, h: 0.75, fontSize: 16, color: K.sub });
    });
    T(s, "誰が決めるのか", { x: 0.6, y: 4.75, w: 4, h: 0.4, fontSize: 18, bold: true, color: K.navy });
    const who = [["［地域・委員会］", "［意見をまとめる］"], ["［決める機関］", "［方針を決める］"], ["［最終の手続き］", "［例：議会・総会で議決］"]];
    who.forEach(([h, d], i) => {
      const x = 0.6 + i * 4.1;
      box(s, x, 5.2, 3.75, 1.15, K.mainL);
      T(s, h, { x: x + 0.2, y: 5.27, w: 3.4, h: 0.4, fontSize: 17, bold: true, color: K.main });
      T(s, d, { x: x + 0.2, y: 5.67, w: 3.4, h: 0.65, fontSize: 16, color: K.ink });
      if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 3.8, y: 5.6, w: 0.25, h: 0.35, fill: { color: K.main }, line: { type: "none" } });
    });
    note(s, "スケジュールは案であり、決定ではありません。", 6.5);
  }

  // 16 意見の出し方
  {
    const s = content("ご意見の出し方と、寄せられた意見の扱い方", null, "「出した意見がどうなるのか」を示すと、意見が集まりやすくなります。");
    const cols = [["出し方", "MdEditNote", [["意見用紙", "本日配布。会場の回収箱、または［提出先］へ"], ["インターネット", "［URL・QRコード］から回答できます"], ["懇談会", "［日程・会場］"]]],
      ["締切・問い合わせ", "MdPhone", [["締切", "令和［ ］年（20［ ］年）［ ］月［ ］日"], ["提出先", "［　　　　　　　］"], ["問い合わせ", "［担当者・部署］\n電話［　　　　　　］"]]],
      ["意見の扱い方", "MdLock", [["すべて読みます", "［誰が］寄せられた意見をすべて読みます"], ["まとめて公表します", "［地区ごと］に集計し、［時期・方法］に公表します"], ["個人は特定しません", "お名前などの個人情報は公表しません"]]]];
    for (let i = 0; i < 3; i++) {
      const [h, ic, items] = cols[i], x = 0.6 + i * 4.1;
      box(s, x, 1.45, 3.9, 5.0, K.bg);
      await iconCircle(s, x + 0.3, 1.65, 0.7, ic);
      T(s, h, { x: x + 1.15, y: 1.65, w: 2.65, h: 0.7, fontSize: 20, bold: true, color: K.navy, valign: "middle" });
      items.forEach(([a, b], j) => {
        const y = 2.55 + j * 1.3;
        T(s, a, { x: x + 0.3, y, w: 3.4, h: 0.4, fontSize: 17, bold: true, color: K.main });
        T(s, b, { x: x + 0.3, y: y + 0.42, w: 3.4, h: 0.85, fontSize: 16, color: K.ink });
      });
    }
    note(s, "［ ］の部分は、確定してから記入します。", 6.6);
  }

  // 17 まとめ
  {
    const s = content("今日のまとめ", null, "最後に要点を4つだけ繰り返します。最後の行には、必ず締切と提出先を書きます。");
    const pts = [["［要点 1：現状］", "［一文で］", "MdTrendingDown"], ["何も決まっていません", "選択肢を、同じ物差しで比べている途中です。", "MdCompareArrows"], ["決める前に、皆さんの声を聞きます", "［誰の声を・どうまとめるか］", "MdForum"], ["ご意見をお寄せください", "締切：令和［ ］年［ ］月［ ］日　提出先：［　　　　］", "MdEditNote"]];
    for (let i = 0; i < 4; i++) {
      const [h, d, ic] = pts[i], y = 1.4 + i * 1.28;
      box(s, 0.6, y, 12.1, 1.12, i % 2 ? K.white : K.bg, i % 2 ? K.line : null);
      await iconCircle(s, 0.85, y + 0.18, 0.76, ic);
      T(s, h, { x: 1.9, y: y + 0.1, w: 10.5, h: 0.48, fontSize: 21, bold: true, color: K.navy, valign: "middle" });
      T(s, d, { x: 1.9, y: y + 0.58, w: 10.5, h: 0.45, fontSize: 17, color: K.ink, valign: "middle" });
    }
  }

  // 18 問いかけ
  {
    const s = pres.addSlide({ masterName: "COVER" });
    hills(s);
    T(s, "皆さんと一緒に考えたいこと", { x: 0.9, y: 0.7, w: 11, h: 0.8, fontSize: 34, bold: true, color: K.white });
    ["［問い 1：例 何を大切にしたいですか？］", "［問い 2：例 今の良さで残したいものは？］", "［問い 3：例 心配なことは何ですか？］", "［問い 4：例 判断の前に知りたいことは？］"].forEach((q, i) => {
      const y = 1.75 + i * 0.95;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.9, y, w: 11.5, h: 0.78, fill: { color: K.white, transparency: 88 }, line: { type: "none" }, rectRadius: 0.1 });
      T(s, `Q${i + 1}`, { x: 1.15, y, w: 0.9, h: 0.78, fontSize: 24, bold: true, color: K.acc, valign: "middle" });
      T(s, q, { x: 2.1, y, w: 10.1, h: 0.78, fontSize: 22, bold: true, color: K.white, valign: "middle" });
    });
    s.addNotes("最後のページ。意見用紙の問いと同じ言葉にしておくと、書きやすくなります。");
  }

  divider("05", "参考資料", "用語の説明・年の対応表");
  SEC = "05　参考資料";

  // 20 用語
  {
    const s = content("用語の説明", null, "資料に出てくる専門用語を、ここで一言ずつ説明します。行は増やしても減らしてもかまいません。");
    const rows = [[th("用語", { align: "left" }), th("意味", { align: "left" })]];
    [["［用語 1］", "［一言で説明］"], ["［用語 2］", "［一言で説明］"], ["［用語 3］", "［一言で説明］"], ["［用語 4］", "［一言で説明］"], ["［用語 5］", "［一言で説明］"], ["［用語 6］", "［一言で説明］"]].forEach(([a, b], i) => {
      const f = { fill: { color: i % 2 ? K.white : K.bg } };
      rows.push([{ text: a, options: Object.assign({ bold: true, color: K.navy }, f) }, { text: b, options: f }]);
    });
    s.addTable(rows, tableOpt({ x: 0.6, y: 1.5, w: 12.1, colW: [3.2, 8.9], rowH: 0.7, fontSize: 18 }));
  }

  // 21 和暦と西暦
  {
    const s = content("令和と西暦の対応", "年度は4月から翌年3月まで", "そのまま使えます。");
    for (let c = 0; c < 2; c++) {
      const rows = [[th("令和"), th("西暦")]];
      for (let r = 0; r < 8; r++) {
        const n = 4 + c * 8 + r, f = { fill: { color: r % 2 ? K.white : K.bg }, align: "center" };
        rows.push([{ text: `令和${n}年`, options: Object.assign({ bold: true, color: K.navy }, f) }, { text: `${2018 + n}年`, options: f }]);
      }
      s.addTable(rows, tableOpt({ x: 0.6 + c * 6.15, y: 1.6, w: 5.95, colW: [2.97, 2.98], rowH: 0.55, fontSize: 18 }));
    }
  }

  const out = path.join(OUTDIR, `説明会資料テンプレート_${themeName}.pptx`);
  await pres.writeFile({ fileName: out });
  return out;
}

// ───────── A4たて：意見用紙（大人用・子ども用）と、お知らせ（回覧・チラシ） ─────────
async function buildA4(themeName) {
  const K = palette(THEMES[themeName]);
  const pres = new pptxgen();
  pres.defineLayout({ name: "A4P", width: 8.27, height: 11.69 });
  pres.layout = "A4P";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "意見用紙・お知らせテンプレート";
  const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 12, color: K.ink, valign: "top", fontFace: FONT }, o));
  const X = 0.6, W = 7.07;
  const checks = (s, y, label, items) => {
    T(s, label, { x: X, y, w: 1.5, h: 0.32, fontSize: 12, bold: true, color: K.navy, valign: "middle" });
    T(s, items.map(t => `□ ${t}`).join("　"), { x: X + 1.5, y, w: W - 1.5, h: 0.32, fontSize: 12, valign: "middle" });
  };
  const question = (s, y, h, q, text) => {
    s.addShape(pres.shapes.RECTANGLE, { x: X, y, w: W, h, fill: { color: K.white }, line: { color: K.line, width: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x: X, y, w: W, h: 0.42, fill: { color: K.mainL }, line: { color: K.line, width: 1 } });
    T(s, q, { x: X + 0.15, y, w: 0.6, h: 0.42, fontSize: 14, bold: true, color: K.main, valign: "middle" });
    T(s, text, { x: X + 0.75, y, w: W - 0.9, h: 0.42, fontSize: 13, bold: true, color: K.navy, valign: "middle" });
    for (let ly = y + 0.84; ly < y + h - 0.05; ly += 0.42) s.addShape(pres.shapes.LINE, { x: X + 0.15, y: ly, w: W - 0.3, h: 0, line: { color: "DCE3DF", width: 0.75, dashType: "dash" } });
  };
  const header = (s, title, lead) => {
    T(s, "［団体名］｜［会の名前］", { x: X, y: 0.45, w: W, h: 0.3, fontSize: 11, color: K.sub });
    T(s, title, { x: X, y: 0.75, w: W, h: 0.55, fontSize: 22, bold: true, color: K.navy });
    T(s, lead, { x: X, y: 1.35, w: W, h: 0.5, fontSize: 12, color: K.sub });
  };
  const footer = (s, text) => {
    s.addShape(pres.shapes.RECTANGLE, { x: X, y: 10.55, w: W, h: 0.7, fill: { color: K.mainL }, line: { type: "none" } });
    T(s, text, { x: X + 0.15, y: 10.55, w: W - 0.3, h: 0.7, fontSize: 11, valign: "middle" });
  };

  // 意見用紙（大人用）
  {
    const s = pres.addSlide();
    header(s, "［〇〇］についての ご意見用紙", "賛成・反対だけでなく、心配なこと・知りたいことも、どうぞお書きください。当てはまるものに✓を付けてください。");
    checks(s, 1.95, "お住まいの地区", ["［地区A］", "［地区B］", "［地区C］", "その他"]);
    checks(s, 2.3, "お立場（複数可）", ["［立場 1］", "［立場 2］", "［立場 3］"]);
    checks(s, 2.62, "", ["地域の住民", "［関係者］", "その他（　　　　）"]);
    checks(s, 2.97, "年代", ["10代", "20代", "30代", "40代"]);
    checks(s, 3.29, "", ["50代", "60代", "70代以上"]);
    ["［問い 1］", "［問い 2］", "［問い 3］", "［問い 4］", "そのほか、ご意見・ご質問をお書きください。"].forEach((q, i) => question(s, 3.75 + i * 1.34, 1.24, `Q${i + 1}`, q));
    footer(s, "締切：令和［ ］年（20［ ］年）［ ］月［ ］日　提出先：［　　　　　　］\nインターネット：［URL・QRコード］　お名前などの個人情報は公表しません。");
    s.addNotes("問いは、説明会資料の最後のページと同じ言葉にしてください。");
  }
  // 意見用紙（子ども用）
  {
    const s = pres.addSlide();
    header(s, "［〇〇］について、みんなの考えを教えてね", "小学生・中学生のみなさんへ。思ったことを、そのまま書いてください。絵でもかまいません。");
    checks(s, 1.95, "学校", ["［学校A］", "［学校B］", "［学校C］"]);
    checks(s, 2.3, "学年", ["1年", "2年", "3年", "4年", "5年", "6年"]);
    ["［やさしい問い 1：例 楽しいことは何？］", "［やさしい問い 2：例 残してほしいものは？］", "［やさしい問い 3：例 やってみたいことは？］", "［やさしい問い 4：例 心配なことは？］", "ほかに、大人に伝えたいことがあれば書いてください。"].forEach((q, i) => question(s, 2.8 + i * 1.5, 1.4, `${i + 1}`, q));
    footer(s, "書いた紙は、［回収する人・場所］にわたしてください。名前は書かなくてよいです。\nみなさんの考えは、［考える大人たち］が、ぜんぶ読みます。");
  }
  // お知らせ（説明会のご案内）
  {
    const s = pres.addSlide();
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 8.27, h: 3.3, fill: { color: K.navy }, line: { type: "none" } });
    s.addShape(pres.shapes.OVAL, { x: -1.5, y: 2.5, w: 6, h: 1.8, fill: { color: K.h2 }, line: { type: "none" } });
    s.addShape(pres.shapes.OVAL, { x: 3.5, y: 2.7, w: 6.5, h: 1.6, fill: { color: K.h3 }, line: { type: "none" } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 3.3, w: 8.27, h: 1.2, fill: { color: K.white }, line: { type: "none" } });
    T(s, "［団体名］からのお知らせ", { x: X, y: 0.45, w: W, h: 0.35, fontSize: 13, color: K.soft });
    T(s, "［〇〇のこれからを考える］\n説明会のご案内", { x: X, y: 0.9, w: W, h: 1.5, fontSize: 32, bold: true, color: K.white, lineSpacingMultiple: 1.1 });
    T(s, "［どなたでも参加できます。お気軽にお越しください。］", { x: X, y: 2.45, w: W, h: 0.4, fontSize: 15, color: K.soft });
    const info = [["日時", "令和［ ］年（20［ ］年）［ ］月［ ］日（［ ］）\n午後［ ］時～［ ］時（受付［ ］時から）"], ["会場", "［会場名］（［住所］）\n駐車場：［あり・なし］"], ["内容", "［内容 1］／［内容 2］／質問・意見交換"], ["対象", "［地区名］にお住まいの方、［関係者］など"], ["申込み", "［不要・要（締切 ［ ］月［ ］日）］"]];
    info.forEach(([a, b], i) => {
      const y = 3.75 + i * 1.02;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X, y, w: 1.35, h: 0.82, fill: { color: K.main }, line: { type: "none" }, rectRadius: 0.08 });
      T(s, a, { x: X, y, w: 1.35, h: 0.82, fontSize: 17, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, b, { x: X + 1.6, y, w: W - 1.6, h: 0.82, fontSize: 16, valign: "middle" });
      s.addShape(pres.shapes.LINE, { x: X + 1.6, y: y + 0.92, w: W - 1.6, h: 0, line: { color: K.line, width: 0.75 } });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X, y: 8.95, w: W, h: 1.1, fill: { color: K.accL }, line: { type: "none" }, rectRadius: 0.1 });
    T(s, "当日は何かを決める場ではありません。\n［皆さんのご意見をうかがう場です。］", { x: X + 0.25, y: 8.95, w: W - 0.5, h: 1.1, fontSize: 16, bold: true, color: K.accD, valign: "middle" });
    T(s, "お問い合わせ：［担当者・部署］　電話［　　　　　　］", { x: X, y: 10.45, w: W, h: 0.4, fontSize: 13, color: K.sub });
    T(s, "［来られない方は、資料を［場所・URL］でご覧いただけます。］", { x: X, y: 10.85, w: W, h: 0.4, fontSize: 13, color: K.sub });
  }
  const out = path.join(OUTDIR, `意見用紙とお知らせ_${themeName}.pptx`);
  await pres.writeFile({ fileName: out });
  return out;
}

// ───────── 使い方ガイド（A4たて 2枚） ─────────
async function buildGuide() {
  const K = palette(THEMES.みどり);
  const pres = new pptxgen();
  pres.defineLayout({ name: "A4P", width: 8.27, height: 11.69 });
  pres.layout = "A4P";
  pres.title = "使い方ガイド";
  const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 12, color: K.ink, valign: "top", fontFace: FONT }, o));
  const X = 0.7, W = 6.87;
  const sec = (s, y, h, body) => {
    T(s, h, { x: X, y, w: W, h: 0.4, fontSize: 16, bold: true, color: K.main });
    s.addShape(pres.shapes.LINE, { x: X, y: y + 0.42, w: W, h: 0, line: { color: K.main, width: 1.5 } });
    T(s, body, { x: X, y: y + 0.55, w: W, h: 2, fontSize: 12.5, lineSpacingMultiple: 1.35 });
  };
  {
    const s = pres.addSlide();
    T(s, "地域の説明会 資料テンプレート", { x: X, y: 0.6, w: W, h: 0.4, fontSize: 13, color: K.sub });
    T(s, "使い方ガイド", { x: X, y: 1.0, w: W, h: 0.7, fontSize: 30, bold: true, color: K.navy });
    sec(s, 2.0, "入っているもの", "・説明会資料テンプレート（16:9、21枚）… みどり／あお／えんじ の3色\n・意見用紙（大人用・子ども用）とお知らせ（A4たて、3枚）… 3色\n・この使い方ガイド");
    sec(s, 3.85, "はじめに：フォントについて", "「BIZ UDPゴシック」を使っています。Windows 10以降には最初から入っています。Mac・その他の環境では、Google Fonts（fonts.google.com）から無料で入れられます。入っていない場合は、別のフォントで表示されます。");
    sec(s, 5.45, "使い方（3ステップ）", "1．［ ］の部分を、内容に合わせて書き換えます。\n2．使わないページは削除します（右クリック →「スライドの削除」）。\n3．グラフは右クリック →「データの編集」で数字を差し替えます。\n各ページの「ノート」欄に、そのページの書き方のコツを書いてあります。");
    sec(s, 7.6, "対応ソフト", "PowerPoint（Windows・Mac）、Keynote、Googleスライド で開けます。\nGoogleスライドでは、一部の図形の位置や文字の幅が少し変わることがあります。");
    sec(s, 9.2, "ご利用の範囲", "購入した方・団体の資料づくりに、何回でもお使いいただけます。\nテンプレートそのものの再配布・再販売はご遠慮ください。");
  }
  {
    const s = pres.addSlide();
    T(s, "伝わる説明会資料の 7つのコツ", { x: X, y: 0.6, w: W, h: 0.6, fontSize: 24, bold: true, color: K.navy });
    const tips = [["結論を先に書かない", "「今日は決める場ではない」と最初に示し、題にも結論を書きません。"], ["どの案も同じ扱いにする", "今の形を続ける案も入れ、すべて同じ色・同じ大きさで並べます。"], ["特定の1つを赤で目立たせない", "地区や施設を比べるときは、同じ色の濃淡で示します。"], ["数字には出典と時点を", "「〇〇（発行者、令和8年5月1日現在）」のように書きます。わからない数字は［ ］で空けておき、作りません。"], ["年は和暦と西暦を併記", "「令和8年度（2026年度）」。巻末の対応表も使えます。"], ["文字は大きく", "本文は16pt以上が目安。年配の方や、後ろの席の方にも読めるように。"], ["意見の行き先を示す", "出し方・締切・扱い方（誰が読み、どう公表するか）を必ず書きます。"]];
    tips.forEach(([h, d], i) => {
      const y = 1.5 + i * 1.38;
      s.addShape(pres.shapes.OVAL, { x: X, y, w: 0.6, h: 0.6, fill: { color: K.main }, line: { type: "none" } });
      T(s, String(i + 1), { x: X, y, w: 0.6, h: 0.6, fontSize: 18, bold: true, color: K.white, align: "center", valign: "middle" });
      T(s, h, { x: X + 0.85, y: y - 0.02, w: W - 0.85, h: 0.4, fontSize: 16, bold: true, color: K.navy });
      T(s, d, { x: X + 0.85, y: y + 0.42, w: W - 0.85, h: 0.85, fontSize: 12.5, lineSpacingMultiple: 1.3 });
    });
  }
  const out = path.join(OUTDIR, "使い方ガイド.pptx");
  await pres.writeFile({ fileName: out });
  return out;
}

(async () => {
  fs.mkdirSync(OUTDIR, { recursive: true });
  for (const t of Object.keys(THEMES)) { console.log(await buildDeck(t)); console.log(await buildA4(t)); }
  console.log(await buildGuide());
})();

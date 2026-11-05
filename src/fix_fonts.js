// pptx のテーマで、日本語用の書体（a:ea）が空のままだと、見る環境によって別の書体に置き換わる。
// テーマの latin / ea / cs をすべて BIZ UDPゴシック にそろえる。
// あわせて、pptxgenjs の出力の不備を直す（どちらも iPhone版 PowerPoint で開けなくなる原因になる）：
//  - [Content_Types].xml に、実際にはない部品（slideMaster2〜）が登録されている → 登録を消す
//  - 同じスライドの中で図形の番号（p:cNvPr の id）が重なる → スライドごとに付け直す
// 書き戻すときは、元のファイルの並び順を保ち（[Content_Types].xml が先頭）、
// フォルダだけの項目を入れない。これを守らないと iPhone版 PowerPoint で「ファイルを開けません」になる。
const fs = require("fs");
const JSZip = require("jszip");
module.exports = async function fixFonts(file, font = "BIZ UDPゴシック") {
  if (process.env.RAW === "1") return; // 確認用：仕上げの処理をしない
  const src = await JSZip.loadAsync(fs.readFileSync(file));
  const names = Object.keys(src.files).filter(n => !src.files[n].dir);
  names.sort((a, b) => (b === "[Content_Types].xml") - (a === "[Content_Types].xml"));
  const out = new JSZip();
  for (const name of names) {
    let data = await src.file(name).async("nodebuffer");
    if (name === "[Content_Types].xml") {
      data = Buffer.from(data.toString("utf8").replace(/<Override PartName="\/([^"]+)"[^>]*\/>/g, (m, part) => (names.includes(part) ? m : "")), "utf8");
    }
    if (/^ppt\/(slides|slideLayouts|slideMasters|notesSlides|notesMasters|charts)\/[^/]+\.xml$/.test(name)) {
      let id = 1;
      data = Buffer.from(data.toString("utf8")
        .replace(/<p:cNvPr id="\d+"/g, () => `<p:cNvPr id="${id++}"`)
        // 言語を日本語に。日本語用の書体の文字コードを中国語（-122）から日本語（-128）に。cs（アラビア語などの書体）は外す
        .replace(/lang="en-US"/g, 'lang="ja-JP"')
        .replace(/(<a:ea typeface="[^"]*"[^>]*?)charset="-122"/g, '$1charset="-128"')
        .replace(/<a:cs typeface="[^"]*"[^>]*\/>/g, ""), "utf8");
    }
    if (/^ppt\/theme\/theme\d+\.xml$/.test(name)) {
      data = Buffer.from(data.toString("utf8").replace(/<a:(latin|ea|cs) typeface="[^"]*"/g, (m, t) => `<a:${t} typeface="${font}"`), "utf8");
    }
    out.file(name, data, { createFolders: false, date: src.files[name].date });
  }
  fs.writeFileSync(file, await out.generateAsync({ type: "nodebuffer", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
};

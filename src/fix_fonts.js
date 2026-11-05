// pptx のテーマで、日本語用の書体（a:ea）が空のままだと、見る環境によって別の書体に置き換わる。
// テーマの latin / ea / cs をすべて BIZ UDPゴシック にそろえる。
// 書き戻すときは、元のファイルの並び順を保ち（[Content_Types].xml が先頭）、
// フォルダだけの項目を入れない。これを守らないと iPhone版 PowerPoint で「ファイルを開けません」になる。
const fs = require("fs");
const JSZip = require("jszip");
module.exports = async function fixFonts(file, font = "BIZ UDPゴシック") {
  const src = await JSZip.loadAsync(fs.readFileSync(file));
  const names = Object.keys(src.files).filter(n => !src.files[n].dir);
  names.sort((a, b) => (b === "[Content_Types].xml") - (a === "[Content_Types].xml"));
  const out = new JSZip();
  for (const name of names) {
    let data = await src.file(name).async("nodebuffer");
    if (/^ppt\/theme\/theme\d+\.xml$/.test(name)) {
      data = Buffer.from(data.toString("utf8").replace(/<a:(latin|ea|cs) typeface="[^"]*"/g, (m, t) => `<a:${t} typeface="${font}"`), "utf8");
    }
    out.file(name, data, { createFolders: false, date: src.files[name].date });
  }
  fs.writeFileSync(file, await out.generateAsync({ type: "nodebuffer", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
};

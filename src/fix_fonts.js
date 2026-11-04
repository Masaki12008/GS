// pptx のテーマで、日本語用の書体（a:ea）が空のままだと、見る環境によって別の書体に置き換わる。
// テーマの latin / ea / cs をすべて BIZ UDPゴシック にそろえる。
const fs = require("fs");
const JSZip = require("jszip");
module.exports = async function fixFonts(file, font = "BIZ UDPゴシック") {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/theme\/theme\d+\.xml$/.test(n))) {
    let x = await zip.file(name).async("string");
    x = x.replace(/<a:(latin|ea|cs) typeface="[^"]*"/g, (m, t) => `<a:${t} typeface="${font}"`);
    zip.file(name, x);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
};

// 仕上げ：LibreOffice で pptx を保存し直す。
// pptxgenjs が出力したファイルは、iPhone版 PowerPoint で「ファイルを開けません」になるが、
// LibreOffice で保存し直したものは開ける（2026-11-05 にユーザーの iPhone で確認）。
// soffice がない環境では何もしない（そのときは iPhone で開けない可能性があると表示する）。
const { execFileSync } = require("child_process");
const fs = require("fs"), os = require("os"), path = require("path");
module.exports = function resave(file) {
  if (process.env.RAW === "1") return;
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "resave-"));
  try {
    execFileSync("soffice", ["--headless", `-env:UserInstallation=file://${dir}/profile`, "--convert-to", "pptx", "--outdir", dir, path.resolve(file)], { stdio: "ignore", timeout: 300000 });
    const out = path.join(dir, path.basename(file));
    if (!fs.existsSync(out)) throw new Error("変換結果がない");
    fs.copyFileSync(out, file);
  } catch (e) {
    console.warn("注意：LibreOffice で保存し直せませんでした（iPhone版 PowerPoint で開けない可能性があります）:", e.message);
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
};

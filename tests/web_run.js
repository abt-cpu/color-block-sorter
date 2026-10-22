// 用瀏覽器（Playwright）跑網頁版，輸出每個檔案的分組結果 JSON。
// 用法: node web_run.js ../index.html 輸出資料夾 檔1.docx 檔2.docx ...
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
(async () => {
  const [html, outDir, ...docs] = process.argv.slice(2);
  fs.mkdirSync(outDir, { recursive: true });
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto('file://' + path.resolve(html));
  for (const f of docs) {
    const b64 = fs.readFileSync(f).toString('base64');
    const r = await p.evaluate(async ({ b64, name }) => {
      const u8 = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
      const r = await window.__ct.processDoc(name, u8, 'hl');
      return { ok: r.ok, total: r.total, groups: [...r.g].map(([k, ps]) => [k, ps.map(pp => pp.map(s => s.text).join(''))]) };
    }, { b64, name: path.basename(f) });
    fs.writeFileSync(path.join(outDir, path.basename(f, '.docx') + '.json'), JSON.stringify(r));
    console.log(path.basename(f), 'web ok =', r.ok, 'chars =', r.total);
  }
  await b.close();
})();

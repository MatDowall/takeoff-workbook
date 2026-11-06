// End-to-end in headless Chrome: keyboard editing, drill-down, live roll-up, PDF measurement → linked cell.
// Expects ../pdf-viewport and ../pdf-snap next to this folder. Run: npm run test:browser
import http from 'node:http';
import { readFile } from 'node:fs/promises';
import { extname, join, normalize } from 'node:path';
import puppeteer from 'puppeteer-core';

const parent = new URL('../..', import.meta.url).pathname;
const types = { '.html': 'text/html', '.js': 'text/javascript', '.mjs': 'text/javascript', '.pdf': 'application/pdf' };
const server = http.createServer(async (req, res) => {
  try {
    const url = decodeURIComponent(req.url.split('?')[0]);
    const p = normalize(join(parent, url));
    if (!p.startsWith(parent)) throw 0;
    let body = await readFile(p);
    if (url.endsWith('/demo/index.html')) body = Buffer.from(String(body).replaceAll('https://cdn.jsdelivr.net/npm/pdfjs-dist@6.5.438/build/', '/pdf-viewport/node_modules/pdfjs-dist/build/'));
    res.writeHead(200, { 'content-type': types[extname(p)] || 'application/octet-stream' });
    res.end(body);
  } catch { res.writeHead(404).end(); }
}).listen(0);

const browser = await puppeteer.launch({ executablePath: process.env.CHROME || '/usr/bin/google-chrome', args: ['--no-sandbox'] });
const page = await browser.newPage();
await page.setViewport({ width: 1600, height: 900 });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('dialog', (d) => d.dismiss());
await page.goto(`http://localhost:${server.address().port}/takeoff-workbook/demo/index.html`);
await page.waitForFunction(() => window.view && window.vp);

let failed = 0;
const check = (name, ok, info = '') => { console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${info ? '  — ' + info : ''}`); if (!ok) failed++; };
const text = (a1) => page.$eval(`.twb tbody td[data-a1="${a1}"]`, (td) => td.textContent.replace('▸', '').trim());
const total = () => page.$eval('.twb-total', (e) => e.textContent);
const clickCell = (a1) => page.click(`.twb tbody td[data-a1="${a1}"]`);
const pause = (ms = 60) => new Promise((r) => setTimeout(r, ms));

// 1. sample estimate maths
check('sample: row maths, subtotal, % formula, total', (await total()) === 'Total 5,494.86', await total());
check('subtotal row adds rows above', (await text('H5')) === '5,233.20', await text('H5'));

// 2. typing
await clickCell('C3');
await page.keyboard.type('120');
await page.keyboard.press('Enter');
await pause();
check('type a quantity → row and total update', (await text('H3')) === '11,520.00' && (await total()) === 'Total 17,590.86', `${await text('H3')} / ${await total()}`);
check('Enter moved selection down', await page.evaluate(() => view.activeCell().a1 === 'C4'));

// 3. drill down from the keyboard and roll up live
await clickCell('C2');
await page.keyboard.down('Control'); await page.keyboard.press('Enter'); await page.keyboard.up('Control');
await pause();
const crumbs = await page.$$eval('.twb-crumbs button', (b) => b.map((x) => x.textContent));
check('Ctrl+Enter drills into a quantity sheet (breadcrumb)', crumbs.length === 2 && /Quantity/.test(crumbs[1]), JSON.stringify(crumbs));
check('existing figure carried into the breakdown', (await text('C1')) === '42.000', await text('C1'));
await clickCell('C1'); await page.keyboard.type('2'); await page.keyboard.press('Tab');
await page.keyboard.type('10'); await page.keyboard.press('Tab');
await page.keyboard.type('3'); await page.keyboard.press('Enter');
await pause();
check('dimension line: 2 × 10 × 3 = 60', (await text('H1')) === '60.000', await text('H1'));
await page.keyboard.down('Alt'); await page.keyboard.press('ArrowUp'); await page.keyboard.up('Alt');
await pause();
check('Alt+↑ back to estimate; drilled qty = 60 and row total live', (await text('C2')) === '60.00' && (await text('H2')) === '2,310.00', `${await text('C2')} / ${await text('H2')}`);

// 4. undo the dimension edits + redo
await page.keyboard.down('Control'); await page.keyboard.press('z'); await page.keyboard.up('Control');
await pause();
const afterUndo = await text('C2');
await page.keyboard.down('Control'); await page.keyboard.press('y'); await page.keyboard.up('Control');
await pause();
check('Ctrl+Z / Ctrl+Y through a breakdown', afterUndo !== '60.00' && (await text('C2')) === '60.00', `undo → ${afterUndo}`);

// 5. PDF measurement → link → live
await (await page.$('#file')).uploadFile(join(parent, 'pdf-snap/test/fixtures/snap-test.pdf'));
await page.waitForFunction(() => window.snapper?.engine);
await page.evaluate(() => { const p = vp.pdfToPage(100, 100); vp.setView({ zoom: 3, panX: 200 - p.x * 3, panY: 300 - p.y * 3 }); });
await pause(300);
const toScreen = (x, y) => page.evaluate((x, y) => { const p = vp.pdfToPage(x, y), s = vp.pageToScreen(p.x, p.y), r = vp.container.getBoundingClientRect(); return { x: s.x + r.left, y: s.y + r.top }; }, x, y);
let a = await toScreen(100, 100), b = await toScreen(150, 100);
await page.mouse.move(a.x + 4, a.y + 3); await page.mouse.click(a.x + 4, a.y + 3);
await page.mouse.move(b.x - 3, b.y + 4); await page.mouse.click(b.x - 3, b.y + 4);
await page.evaluate(() => demo.finish());
await pause();
const mlen = await page.$eval('#meas .v', (e) => e.textContent);
check('snapped measurement 50 pt @1:100 = 1.76 m', mlen === '1.76 m', mlen);
await clickCell('C7');
await page.click('#linkBtn');
await clickCell('E7'); await page.keyboard.type('100'); await page.keyboard.press('Enter');
await pause();
check('linked cell shows measured value, unit filled', (await text('C7')) === '1.76' && (await text('D7')) === 'm' && (await text('H7')) === '176.39', `${await text('C7')} ${await text('D7')} ${await text('H7')}`);
await page.$eval('#scale', (e) => { e.value = '200'; e.dispatchEvent(new Event('input')); });
await pause();
check('change drawing scale → linked cell, row and total update live', (await text('H7')) === '352.78', await text('H7'));
// link modifier via F2
await clickCell('C7'); await page.keyboard.press('F2');
await page.keyboard.down('Control'); await page.keyboard.press('a'); await page.keyboard.up('Control');
await page.keyboard.type('=link*1.1'); await page.keyboard.press('Enter');
await pause();
check('F2 on a linked cell edits its modifier (=link*1.1), link kept', (await text('H7')) === '388.06' && (await page.evaluate(() => wb.info(wb.rootId, 'C7').type)) === 'link', await text('H7'));
await page.screenshot({ path: new URL('./demo.png', import.meta.url).pathname });

// 6. deleting the measurement is loud, not silently zero
await page.click('#meas button[title="Delete"]');
await pause();
check('deleted measurement → #REF! in cell and total', (await text('C7')) === '#REF!' && (await total()) === 'Total #REF!', `${await text('C7')} / ${await total()}`);

check('no page errors', errors.length === 0, errors.join(' | '));
await browser.close(); server.close();
console.log(failed ? `\n${failed} FAILED` : '\nall browser checks passed');
process.exit(failed ? 1 : 0);

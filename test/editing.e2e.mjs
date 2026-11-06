// Headless Chrome: formula autocomplete, argument hints, point-mode references
// (click, drag, arrows, cross-sheet), reference highlights, column/row resizing.
// Run: npm run test:browser   (needs ../pdf-viewport and ../pdf-snap beside this folder)
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
await page.waitForFunction(() => window.view);

let failed = 0;
const check = (name, ok, info = '') => { console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${info ? '  — ' + info : ''}`); if (!ok) failed++; };
const pause = (ms = 60) => new Promise((r) => setTimeout(r, ms));
const td = (a1) => `.twb tbody td[data-a1="${a1}"]`;
const text = (a1) => page.$eval(td(a1), (e) => e.textContent.replace('▸', '').trim());
const editorText = () => page.evaluate(() => view.ed?.text ?? null);
const center = async (sel) => { const b = await (await page.$(sel)).boundingBox(); return { x: b.x + b.width / 2, y: b.y + b.height / 2 }; };
const clickCell = (a1) => page.click(td(a1));
const ctrl = async (k) => { await page.keyboard.down('Control'); await page.keyboard.press(k); await page.keyboard.up('Control'); };

// ---------------------------------------------------------------- autocomplete
await clickCell('I10');
await page.keyboard.type('=RO');
await pause();
let ac = await page.$$eval('.twb-ac .it b', (b) => b.map((x) => x.textContent));
check('typing =RO lists ROUND, ROUNDUP, ROUNDDOWN at the cursor', ac.join() === 'ROUND,ROUNDUP,ROUNDDOWN', ac.join());
const acBox = await page.$eval('.twb-ac', (e) => e.getBoundingClientRect().toJSON());
const cellBox = await page.$eval(td('I10'), (e) => e.getBoundingClientRect().toJSON());
check('popup sits just under the edited cell (kept on screen)', Math.abs(acBox.top - cellBox.bottom) < 6 && acBox.right > cellBox.left && acBox.right <= 1600, `popup top ${acBox.top} vs cell bottom ${cellBox.bottom}`);
const desc = await page.$eval('.twb-ac .desc', (e) => e.textContent);
check('selected entry shows its description', /decimal places/.test(desc), desc);
await page.keyboard.press('ArrowDown');
await page.keyboard.press('Tab');
await pause();
check('↓ + Tab inserts ROUNDUP(', (await editorText()) === '=ROUNDUP(', await editorText());
let sig = await page.$eval('.twb-sig', (e) => (e.style.display !== 'none' ? e.innerHTML : ''));
check('argument hint shows signature with current arg bold', /ROUNDUP<\/span>\(<b>number<\/b>, \[digits\]\)/.test(sig), sig);

// ---------------------------------------------------------------- point mode: click + drag
let a = await center(td('H2')), b = await center(td('H4'));
await page.mouse.move(a.x, a.y);
await page.mouse.down();
await page.mouse.move(b.x, b.y, { steps: 4 });
await page.mouse.up();
await pause();
check('drag H2→H4 inserts a range while typing', (await editorText()) === '=ROUNDUP(H2:H4', await editorText());
const boxes = await page.$$eval('.twb-hl', (e) => e.map((x) => ({ pt: x.classList.contains('pt'), w: x.offsetWidth, h: x.offsetHeight, top: x.offsetTop })));
const h2 = await page.$eval(td('H2'), (e) => ({ top: e.offsetTop, h: e.offsetHeight }));
check('pointed range outlined on the grid (marching ants) and aligned with the cells', boxes.length === 1 && boxes[0].pt && boxes[0].top === h2.top && boxes[0].h === h2.h * 3, JSON.stringify({ boxes, h2 }));
await page.keyboard.type(',');
await pause();
sig = await page.$eval('.twb-sig', (e) => e.innerHTML);
check('after the comma the hint moves to [digits]', /<b>\[digits\]<\/b>/.test(sig), sig);
await page.keyboard.type('0');
await page.keyboard.press('Backspace');
await page.keyboard.press('Backspace');
await page.keyboard.press('Backspace');
await page.keyboard.press('Backspace');
await page.keyboard.press('Backspace');
await page.keyboard.press('Backspace');
await page.keyboard.press('Backspace');
await page.keyboard.press('Backspace'); // back to "=ROUNDUP(" … then build SUM instead
await page.keyboard.down('Control'); await page.keyboard.press('a'); await page.keyboard.up('Control');
await page.keyboard.type('=SUM(');
a = await center(td('H2'));
await page.mouse.click(a.x, a.y);
await page.keyboard.type('+');
b = await center(td('H4'));
await page.mouse.click(b.x, b.y);
await page.keyboard.type(')');
check('click, type +, click: =SUM(H2+H4)', (await editorText()) === '=SUM(H2+H4)', await editorText());
const colours = await page.$$eval('.twb-hl', (e) => [...new Set(e.map((x) => x.style.borderColor))].length);
check('each reference gets its own colour', colours === 2, `${colours} colours`);
await page.keyboard.press('Enter');
await pause();
check('Enter commits: value = H2 + H4', (await text('I10')) === (await page.evaluate(() => wb.format(wb.rootId, 'I10', wb.value(wb.rootId, 'H2') + wb.value(wb.rootId, 'H4')))), await text('I10'));
check('highlights cleared after commit', (await page.$$('.twb-hl')).length === 0);

// ---------------------------------------------------------------- point mode: arrows
await clickCell('E10');
await page.keyboard.type('=');
await page.keyboard.press('ArrowUp');
await page.keyboard.press('ArrowUp');
check('= then ↑↑ points at E8', (await editorText()) === '=E8', await editorText());
await page.keyboard.down('Shift'); await page.keyboard.press('ArrowUp'); await page.keyboard.press('ArrowUp'); await page.keyboard.up('Shift');
check('Shift+↑↑ extends to E6:E8', (await editorText()) === '=E6:E8', await editorText());
await page.keyboard.press('Escape');
check('Esc cancels the edit', (await editorText()) === null && (await text('E10')) === '');
// plain typed value: arrows commit and move (no pointing)
await clickCell('E10');
await page.keyboard.type('5');
await page.keyboard.press('ArrowRight');
check('typed value + → commits and moves right', (await text('E10')) === '5.00' && (await page.evaluate(() => view.activeCell().a1)) === 'F10');

// ---------------------------------------------------------------- cross-sheet pointing
await page.evaluate(() => wb.drill(wb.rootId, 'C2'));
await pause();
await clickCell('J12');
await page.keyboard.type('=2*');
await page.$eval('.twb-scroll', (e) => (e.scrollLeft = 0)); // C2 is under the sticky row numbers after editing J12
await page.click(`${td('C2')} .open`); // open the breakdown mid-formula
await pause(100);
const away = await page.$eval('.twb-fx .hint', (e) => e.textContent);
check('opening a breakdown mid-formula keeps the edit (formula bar)', /editing Estimate › J12/.test(away) && (await editorText()) === '=2*', away);
await clickCell('H1');
const xref = await editorText();
check('clicking a cell on the other sheet inserts [Sheet]!H1', /^=2\*\[.+— Quantity\]!H1$/.test(xref), xref);
await page.keyboard.press('Enter');
await pause(100);
const crumbs = await page.$$eval('.twb-crumbs button', (b) => b.length);
check('Enter commits and returns to the home sheet', crumbs === 1 && (await text('J12')) === '84', `${crumbs} crumbs, J12=${await text('J12')}`);

// ---------------------------------------------------------------- MEAS autocomplete
await page.evaluate(() => wb.setMeasurement('mx', { value: 12.5, unit: 'm2', name: 'Ground floor slab' }));
await clickCell('K12');
await page.keyboard.type('=MEAS("gro');
await pause();
ac = await page.$$eval('.twb-ac .it b', (b) => b.map((x) => x.textContent));
check('inside MEAS(" the list shows matching measurements', ac.join() === 'Ground floor slab', ac.join());
await page.keyboard.press('Enter'); // accept
await page.keyboard.press('Enter'); // commit
await pause();
check('accepting completes =MEAS("Ground floor slab") → 12.5', (await page.evaluate(() => wb.cell(wb.rootId, 'K12').raw)) === '=MEAS("Ground floor slab")' && (await text('K12')) === '12.5', await text('K12'));

// ---------------------------------------------------------------- column / row resizing
await page.$eval('.twb-scroll', (e) => (e.scrollLeft = 0));
const wB0 = await page.evaluate(() => wb.columnWidth(wb.rootId, 'B'));
let h = await center('.twb thead th[data-col="B"] .rz');
await page.mouse.move(h.x, h.y);
await page.mouse.down();
await page.mouse.move(h.x + 120, h.y, { steps: 5 });
await page.mouse.up();
await pause();
const wB1 = await page.evaluate(() => wb.columnWidth(wb.rootId, 'B'));
const domW = await page.$eval(td('B2'), (e) => e.offsetWidth);
check('drag header edge widens column B by 120px (model + DOM)', wB1 === wB0 + 120 && Math.abs(domW - wB1) <= 1, `${wB0} → ${wB1}, td ${domW}`);
const childW = await page.evaluate(() => wb.columnWidth(wb.children(wb.rootId)[0].id, 'B'));
check('cost-sheet width does not leak into quantity sheets', childW !== wB1, `${childW}`);
await ctrl('z');
await pause();
check('Ctrl+Z undoes the resize', (await page.evaluate(() => wb.columnWidth(wb.rootId, 'B'))) === wB0);
await page.click('.twb thead th[data-col="B"] .rz', { clickCount: 2 });
await pause();
const fit = await page.evaluate(() => wb.columnWidth(wb.rootId, 'B'));
const textW = await page.evaluate(() => { const c = document.createElement('canvas').getContext('2d'); c.font = '13px system-ui,sans-serif'; return c.measureText('Excavate to reduced level').width; });
check('double-click edge auto-fits to the longest description', fit >= textW + 10 && fit < 260, `fit ${fit}px for ${textW.toFixed(0)}px of text`);
h = await center('.twb tbody th[data-row="3"] .rzr');
await page.mouse.move(h.x, h.y);
await page.mouse.down();
await page.mouse.move(h.x, h.y + 30, { steps: 4 });
await page.mouse.up();
await pause();
const rowInfo = await page.evaluate(() => ({ h: wb.rowHeight(wb.rootId, 3), wrap: document.querySelector('.twb tr[data-r="3"]').classList.contains('wrap'), dom: document.querySelector('.twb tr[data-r="3"]').offsetHeight }));
check('drag row edge: row 3 taller, wraps text', rowInfo.h === 54 && rowInfo.wrap && Math.abs(rowInfo.dom - 54) <= 1, JSON.stringify(rowInfo));
const saved = await page.evaluate(() => { const j = JSON.parse(JSON.stringify(wb.toJSON())); return { layout: j.layout.cols.cost.B, row: j.sheets.find((s) => s.id === j.rootId).rowHeights['3'] }; });
check('widths and heights saved in the JSON', saved.layout === fit && saved.row === 54, JSON.stringify(saved));
await page.screenshot({ path: new URL('./editing.png', import.meta.url).pathname });

// final visual: mid-formula with autocomplete + highlights
await clickCell('I14');
await page.keyboard.type('=SUM(');
a = await center(td('H2')); b = await center(td('H3'));
await page.mouse.move(a.x, a.y); await page.mouse.down(); await page.mouse.move(b.x, b.y, { steps: 3 }); await page.mouse.up();
await page.keyboard.type(')*ROUN');
await pause();
await page.screenshot({ path: new URL('./editing-formula.png', import.meta.url).pathname });
await page.keyboard.press('Escape'); await page.keyboard.press('Escape');

check('no page errors', errors.length === 0, errors.join(' | '));
await browser.close(); server.close();
console.log(failed ? `\n${failed} FAILED` : '\nall editing checks passed');
process.exit(failed ? 1 : 0);

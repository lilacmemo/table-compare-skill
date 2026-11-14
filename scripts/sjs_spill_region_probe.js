// 探针4（用户假设验证）：导入后 #SPILL! 是否因为目标区域被占用/区域不够
// 三个假设依次验证：A 扩大工作表行列数  B 清除锚点右下方占用区  C 清区后重写公式
// 用法: node sjs_spill_region_probe.js [xlsx路径]
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const http = require('http');
const ROOT = __dirname;
const PORT = 8961;
const FILE = process.argv[2] || path.join(ROOT, '..', '公式批次1样本.xlsx');
const MIME = { '.html': 'text/html', '.js': 'application/javascript' };
const server = http.createServer((req, res) => {
    const p = path.join(ROOT, decodeURIComponent(req.url.split('?')[0]));
    fs.readFile(p, (e, d) => {
        if (e) { res.writeHead(404); res.end(); return; }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
        res.end(d);
    });
});
(async () => {
    await new Promise(r => server.listen(PORT, r));
    const browser = await chromium.launch({ channel: 'chrome', headless: true });
    const page = await browser.newPage();
    page.on('console', m => { if (m.type() === 'error') console.log('[page error]', m.text()); });
    await page.goto(`http://localhost:${PORT}/test.html`);
    await page.waitForFunction(() => typeof window.runImport === 'function');
    const b64 = fs.readFileSync(FILE).toString('base64');
    await page.evaluate(d => window.runImport(d, false), b64);
    await page.waitForTimeout(3000);

    const result = await page.evaluate(async () => {
        const sleep = ms => new Promise(r => setTimeout(r, ms));
        const spread = window.__spread;
        const isSpill = v => v && typeof v === 'object' && String(v._error || v._calcError || v.error || JSON.stringify(v)).indexOf('SPILL') >= 0;
        const read = (s, r, c) => { try { const v = s.getValue(r, c); return v && typeof v === 'object' ? (v._error || v._calcError || JSON.stringify(v)) : v; } catch (e) { return 'ERR:' + e.message; } };
        const fml = (s, r, c) => { try { return s.getFormula(r, c); } catch (e) { return null; } };

        // 1) 找出所有 #SPILL! 锚点（跳过 Evaluation Version 页）
        const anchors = [];
        for (let i = 0; i < spread.getSheetCount(); i++) {
            const s = spread.getSheet(i);
            if (/evaluation/i.test(s.name())) continue;
            const rc = s.getRowCount(), cc = s.getColumnCount();
            for (let r = 0; r < Math.min(rc, 300); r++) {
                for (let c = 0; c < Math.min(cc, 40); c++) {
                    if (isSpill(read(s, r, c))) {
                        anchors.push({ sheet: s.name(), r, c, formula: fml(s, r, c), rowCount: rc, colCount: cc });
                    }
                }
            }
        }
        const out = { fileSheets: spread.getSheetCount(), anchors: anchors.length, detail: [], sample: [] };
        out.sample = anchors.slice(0, 8).map(a => `${a.sheet}!R${a.r + 1}C${a.c + 1} = ${a.formula} (rowCount=${a.rowCount})`);

        // 只取前 6 个锚点做修复验证（不同 sheet/类型各取几个）
        const pick = [];
        const seen = {};
        for (const a of anchors) {
            const key = a.sheet;
            if ((seen[key] || 0) < 2) { pick.push(a); seen[key] = (seen[key] || 0) + 1; }
            if (pick.length >= 6) break;
        }

        for (const a of pick) {
            const s = spread.getSheetFromName(a.sheet);
            const rec = { addr: `${a.sheet}!R${a.r + 1}C${a.c + 1}`, formula: a.formula, steps: {} };
            rec.steps.before = { anchor: read(s, a.r, a.c), right: read(s, a.r, a.c + 1), below: read(s, a.r + 1, a.c) };
            rec.steps.neighborsRaw = [];
            // 记录锚点右侧 3 列、下方 3 行原始内容（哪些被占用）
            for (let dr = 0; dr <= 3; dr++) {
                for (let dc = 0; dc <= 3; dc++) {
                    if (!dr && !dc) continue;
                    const v = read(s, a.r + dr, a.c + dc);
                    if (v !== null && v !== undefined && v !== '') rec.steps.neighborsRaw.push(`R+${dr}C+${dc}=${String(v).slice(0, 20)}`);
                }
            }

            // A) 扩大行列数
            try {
                s.setRowCount(Math.max(s.getRowCount(), a.r + 60));
                s.setColumnCount(Math.max(s.getColumnCount(), a.c + 20));
                s.recalcAll(true);
                await sleep(1200);
                rec.steps.afterExpand = { anchor: read(s, a.r, a.c), right: read(s, a.r, a.c + 1), below: read(s, a.r + 1, a.c) };
            } catch (e) { rec.steps.afterExpand = 'ERR ' + e.message; }

            // B) 清除锚点右下方占用区（保留锚点本身）
            try {
                for (let dr = 0; dr <= 6; dr++) for (let dc = 0; dc <= 8; dc++) {
                    if (!dr && !dc) continue;
                    s.setValue(a.r + dr, a.c + dc, null);
                    s.setFormula(a.r + dr, a.c + dc, null);
                }
                s.recalcAll(true);
                await sleep(1200);
                rec.steps.afterClear = { anchor: read(s, a.r, a.c), right: read(s, a.r, a.c + 1), below: read(s, a.r + 1, a.c) };
            } catch (e) { rec.steps.afterClear = 'ERR ' + e.message; }

            // C) 清区后重写公式（先清锚点再写回）
            try {
                const f = a.formula;
                s.setFormula(a.r, a.c, null);
                s.setFormula(a.r, a.c, f);
                s.recalcAll(true);
                await sleep(1500);
                rec.steps.afterRewrite = { anchor: read(s, a.r, a.c), right: read(s, a.r, a.c + 1), below: read(s, a.r + 1, a.c) };
                // 判断是否恢复成数组：看锚点右侧/下方是否出现数值且非错误
                const rr = rec.steps.afterRewrite.right, bb = rec.steps.afterRewrite.below;
                rec.steps.spillRestored = !(rr && typeof rr === 'object') && rr !== null && rr !== '' && !(bb && typeof bb === 'object' && isSpill(bb));
            } catch (e) { rec.steps.afterRewrite = 'ERR ' + e.message; }

            out.detail.push(rec);
        }
        return out;
    });
    console.log(JSON.stringify(result, null, 1));
    await browser.close();
    server.close();
    process.exit(0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });

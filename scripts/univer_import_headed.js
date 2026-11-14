// batch1_univer_headed.js —— 有头模式（真实可见窗口）验证批次1
// 与 batch1_univer.js 的区别：
// 1. headless:false —— 真实渲染管线（GPU 合成、布局、可见性），非无头近似
// 2. 逐表实测 DOM 中图表容器的真实像素包围盒（getBoundingClientRect），
//    不只依赖 Univer 数据层的 size —— 直接回答"屏幕上到底画没画、画了多大"
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const FILE = process.argv[2] || path.join(__dirname, '..', '批次1样式样本.xlsx');

(async () => {
    const browser = await chromium.launch({
        channel: 'chrome',        // 用户已安装 Chrome（用户级目录，Playwright 自动识别）
        headless: false,          // 真实可见窗口
        args: ['--window-size=1560,980', '--start-maximized'],
    });
    const ctx = await browser.newContext({ viewport: null });
    const page = await ctx.newPage();
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(8000);

    const b64 = fs.readFileSync(FILE).toString('base64');
    const result = await page.evaluate(async (b64) => {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const f = new File([bytes], 'batch1.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const snap = await univerAPI.importSheetToSnapshotAsync(f);
        univerAPI.createWorkbook(snap);
        await new Promise(r => setTimeout(r, 12000));
        const wb = univerAPI.getActiveWorkbook();
        const sheets = [];
        for (const s of wb.getSheets()) {
            let charts = [];
            try { charts = s.getCharts(); } catch (e) { charts = []; }
            const list = charts.map(c => {
                const o = { id: c.getId(), type: String(c.getType()) };
                try { const info = c.getInfo(); o.size = info.size; o.anchor = info.anchor; } catch (e) { o.infoErr = String(e); }
                return o;
            });
            sheets.push({ sheet: s.getSheetName(), chartCount: charts.length, charts: list });
        }
        return { sheets };
    }, b64);

    console.log('数据层: ' + result.sheets.map(s => `${s.sheet}=${s.chartCount}图`).join(', '));

    // 逐表：DOM 实测图表容器真实包围盒 + 截图
    const domReport = [];
    for (let i = 0; i < result.sheets.length; i++) {
        const name = result.sheets[i].sheet;
        try {
            await page.evaluate((n) => {
                const wb = univerAPI.getActiveWorkbook();
                const s = wb.getSheets().find(x => x.getSheetName() === n);
                if (s) wb.setActiveSheet(s);
            }, name);
            await page.waitForTimeout(2500);
            const dom = await page.evaluate(() => {
                // Univer 图表挂在 drawing 容器下的 canvas / div 中，宽搜可能的容器
                const sels = [
                    '.univer-chart', '[class*="chart"]', '.univer-drawing',
                    '[class*="drawing"]', 'canvas',
                ];
                const seen = new Set();
                const boxes = [];
                for (const sel of sels) {
                    document.querySelectorAll(sel).forEach(el => {
                        const r = el.getBoundingClientRect();
                        // 只要有实际面积（>8px）且可见的元素才记录
                        if (r.width > 8 && r.height > 8 && !seen.has(el)) {
                            seen.add(el);
                            const cls = (el.className && el.className.toString().slice(0, 60)) || el.tagName;
                            boxes.push({ sel, cls, x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height) });
                        }
                    });
                }
                return boxes;
            });
            // 过滤出图表相关的（非表格主 canvas）
            const chartish = dom.filter(b => /chart|drawing/i.test(b.cls) || (b.sel === 'canvas' && b.h < 800));
            domReport.push({ sheet: name, chartishCount: chartish.length, sample: chartish.slice(0, 6) });
            await page.screenshot({ path: path.join(__dirname, `batch1_headed_${String(i).padStart(2, '0')}_${name.replace(/[^\w]/g, '')}.png`) });
        } catch (e) { console.error('err', name, e.message); }
    }
    fs.writeFileSync(path.join(__dirname, 'batch1_headed_dom.json'), JSON.stringify(domReport, null, 1));
    for (const d of domReport) {
        console.log(`[DOM] ${d.sheet}: 疑似图表容器 ${d.chartishCount} 个`);
        for (const b of d.sample) console.log(`    ${b.cls} @(${b.x},${b.y}) ${b.w}x${b.h}`);
    }
    await browser.close();
    console.log('DONE headed verification');
    process.exit(0);
})().catch(e => { console.error('FATAL:', e); process.exit(1); });

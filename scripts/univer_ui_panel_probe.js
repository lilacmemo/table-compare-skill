// batch1_ui_probe7.js —— 右键 Edit chart 打开编辑面板，枚举全部可改配置项
const { chromium } = require('playwright');
const path = require('path');

(async () => {
    const browser = await chromium.launch({ channel: 'chrome', headless: false, args: ['--window-size=1660,1000'] });
    const ctx = await browser.newContext({ viewport: null });
    const page = await ctx.newPage();
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(8000);

    await page.evaluate(async () => {
        const wb = univerAPI.getActiveWorkbook();
        const sheet = wb.insertSheet('UITest', { sheet: { rowCount: 40, columnCount: 15 } });
        sheet.getRange('A1:B6').setValues([['月份','销量'],['一月',120],['二月',180],['三月',220],['四月',260],['五月',310]]);
        const chart = sheet.newChart(univerAPI.Enum.ChartTypeString.Column)
            .setSource({ range: 'A1:B6', orientation: univerAPI.Enum.ChartSourceOrientation.Columns })
            .setPosition(1, 3, 0, 0).setSize(560, 300).setTitle('UI探测图').build();
        await sheet.insertChart(chart);
    });
    await page.waitForTimeout(3500);

    // 右键 → Edit chart
    await page.mouse.click(550, 340, { button: 'right' });
    await page.waitForTimeout(1200);
    const editItem = await page.evaluate(() => {
        for (const el of document.querySelectorAll('div,span,li')) {
            const t = (el.textContent || '').trim();
            if (t === 'Edit chart') {
                const r = el.getBoundingClientRect();
                if (r.width > 0) return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
            }
        }
        return null;
    });
    console.log('Edit chart 菜单项:', JSON.stringify(editItem));
    if (!editItem) { await browser.close(); process.exit(1); }
    await page.mouse.click(editItem.x, editItem.y);
    await page.waitForTimeout(2500);
    await page.screenshot({ path: path.join(__dirname, 'ui_probe7_01_panel.png') });

    // 枚举面板：找屏幕右侧新出现的容器
    const dump = await page.evaluate(() => {
        // 找最右侧的大容器
        let panel = null, best = 0;
        document.querySelectorAll('div').forEach(el => {
            const r = el.getBoundingClientRect();
            if (r.x >= 1150 && r.width >= 200 && r.width <= 420 && r.height > 200 && r.height <= 1000) {
                const area = r.width * r.height;
                if (area > best) { best = area; panel = el; }
            }
        });
        if (!panel) return null;
        const items = [];
        panel.querySelectorAll('button,label,select,input,[role="tab"],[role="switch"],h1,h2,h3,h4,h5,span').forEach(el => {
            const t = (el.getAttribute('title') || el.textContent || '').trim().replace(/\s+/g, ' ');
            const type = el.tagName === 'INPUT' ? ('input:' + (el.type || 'text')) : el.tagName;
            if (t && t.length < 60) items.push({ type, text: t.slice(0, 50) });
        });
        return [...new Map(items.map(x => [x.text + x.type, x])).values()];
    });
    console.log('编辑面板内容:', JSON.stringify(dump, null, 0).slice(0, 5000));

    // 列出面板的 tab 并逐个点击截图
    const tabs = await page.evaluate(() => {
        const out = [];
        document.querySelectorAll('[role="tab"], button').forEach(el => {
            const r = el.getBoundingClientRect();
            if (r.x >= 1150 && r.width > 0 && r.y > 80 && r.y < 400) {
                const t = (el.textContent || '').trim();
                if (t && t.length < 30) out.push({ x: r.x + r.width / 2, y: r.y + r.height / 2, text: t });
            }
        });
        return out;
    });
    console.log('面板 tabs:', JSON.stringify(tabs));
    let i = 0;
    for (const tb of tabs) {
        await page.mouse.click(tb.x, tb.y);
        await page.waitForTimeout(1500);
        await page.screenshot({ path: path.join(__dirname, `ui_probe7_tab${i++}_${tb.text.replace(/[^\w]/g, '')}.png`) });
        const dump2 = await page.evaluate(() => {
            let panel = null, best = 0;
            document.querySelectorAll('div').forEach(el => {
                const r = el.getBoundingClientRect();
                if (r.x >= 1150 && r.width >= 200 && r.width <= 420 && r.height > 200 && r.height <= 1000) {
                    const area = r.width * r.height;
                    if (area > best) { best = area; panel = el; }
                }
            });
            if (!panel) return null;
            const items = [];
            panel.querySelectorAll('button,label,select,input,span,h3,h4,h5').forEach(el => {
                const t = (el.getAttribute('title') || el.textContent || '').trim().replace(/\s+/g, ' ');
                if (t && t.length < 60) items.push(t.slice(0, 50));
            });
            return [...new Set(items)];
        });
        console.log(`tab[${tb.text}] 内容:`, JSON.stringify(dump2).slice(0, 1500));
    }
    await browser.close();
    console.log('DONE ui probe7');
    process.exit(0);
})().catch(e => { console.error('FATAL:', e); process.exit(1); });

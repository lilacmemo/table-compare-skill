/**
 * measure_chart_size.js —— 量化「导入后图表尺寸丢失」
 *
 * 四层证据里的几何层：直接读图表对象 getInfo().size / position / anchor
 * 用途：把"图表看不见"从主观描述变成可复现的数字（size=0×0 占比）
 *
 * 用法: node measure_chart_size.js <xlsx> [--stretch]
 *   --stretch 会选中若干图表执行 setSize(640,360) 后截图，验证渲染链路是否正常
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { PNG } = require('pngjs');

const FILE = process.argv[2];
const DO_STRETCH = process.argv.includes('--stretch');
if (!FILE) { console.error('用法: node measure_chart_size.js <xlsx> [--stretch]'); process.exit(1); }

function inkRatio(buf) {
    const png = PNG.sync.read(buf);
    const bg = [png.data[0], png.data[1], png.data[2]];
    let ink = 0;
    for (let i = 0; i < png.data.length; i += 4) {
        const d = Math.abs(png.data[i] - bg[0]) + Math.abs(png.data[i + 1] - bg[1]) + Math.abs(png.data[i + 2] - bg[2]);
        if (d > 12) ink++;
    }
    return +(ink / (png.width * png.height)).toFixed(4);
}

(async () => {
    const browser = await chromium.launch({ channel: 'msedge', headless: true });
    const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 });
    const page = await ctx.newPage();
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(10000);

    const b64 = fs.readFileSync(FILE).toString('base64');
    const data = await page.evaluate(async (b64) => {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const f = new File([bytes], 't.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const snap = await univerAPI.importSheetToSnapshotAsync(f);
        if (!snap) return { ok: false };
        univerAPI.createWorkbook(snap);
        await new Promise(r => setTimeout(r, 9000));
        const wb = univerAPI.getActiveWorkbook();
        const rows = [];
        for (const s of wb.getSheets()) {
            const charts = (typeof s.getCharts === 'function') ? s.getCharts() : [];
            for (const c of charts) {
                let info = null;
                try { info = c.getInfo(); } catch (e) { }
                rows.push({
                    sheet: s.getSheetName(),
                    id: c.getId(),
                    type: String(c.getType()),
                    size: info?.size || null,
                    position: info?.position || null,
                    anchor: info?.anchor || null,
                    title: info?.config?.title?.text || null,
                });
            }
        }
        return { ok: true, sheetCount: wb.getSheets().length, rows };
    }, b64);

    if (!data.ok) { console.error('导入失败'); process.exit(1); }

    const rows = data.rows;
    const zero = rows.filter(r => !r.size || (r.size.width === 0 && r.size.height === 0)).length;
    const tinyH = rows.filter(r => r.size && r.size.height > 0 && r.size.height < 24).length;
    const normal = rows.filter(r => r.size && r.size.height >= 24).length;

    console.log(`文件: ${path.basename(FILE)}`);
    console.log(`工作表 ${data.sheetCount}，图表对象 ${rows.length}`);
    console.log(`\n=== ② 几何层：导入后图表尺寸分布 ===`);
    console.log(`  size = 0×0（默认尺寸完全丢失）: ${zero} / ${rows.length}  (${(zero / rows.length * 100).toFixed(0)}%)`);
    console.log(`  0 < height < 24px（被压成一条缝）: ${tinyH}`);
    console.log(`  height >= 24px（可正常显示）: ${normal}`);

    console.log(`\n分表统计：`);
    const bySheet = {};
    for (const r of rows) {
        bySheet[r.sheet] = bySheet[r.sheet] || { total: 0, zero: 0, heights: [] };
        bySheet[r.sheet].total++;
        if (!r.size || r.size.height === 0) bySheet[r.sheet].zero++;
        bySheet[r.sheet].heights.push(r.size ? r.size.height : 0);
    }
    for (const [k, v] of Object.entries(bySheet)) {
        console.log(`  ${k}: ${v.total} 图，size=0 的 ${v.zero} 个，高度样本 [${v.heights.slice(0, 6).join(', ')}]`);
    }

    fs.writeFileSync(path.join(__dirname, 'chart_size_report.json'), JSON.stringify({ file: FILE, total: rows.length, zero, tinyH, normal, rows }, null, 1));

    // ③ 渲染链路验证：手动 setSize 后是否画出来
    if (DO_STRETCH) {
        const target = rows.find(r => /柱形图/.test(r.sheet));
        if (target) {
            await page.evaluate(async (sheetName) => {
                const wb = univerAPI.getActiveWorkbook();
                const s = wb.getSheets().find(x => x.getSheetName().includes(sheetName));
                wb.setActiveSheet(s);
                await new Promise(r => setTimeout(r, 1500));
                const charts = s.getCharts();
                charts.forEach((c, i) => {
                    try {
                        c.setSize(640, 360);
                        if (i > 0) c.setPosition(0, i % 2 === 0 ? 500 : 0);
                    } catch (e) { }
                });
            }, '柱形图');
            await page.waitForTimeout(5000);
            const shot = await page.screenshot({ scale: 'css' });
            fs.writeFileSync(path.join(__dirname, 'vc_stretch_result.png'), shot);
            console.log(`\n=== ③ 渲染层：setSize(640×360) 后截图 ===`);
            console.log(`  截图: vc_stretch_result.png，整屏着墨率 ${inkRatio(shot)}（对照：空白表约 0.02~0.05）`);
        }
    }

    console.log(`\n明细已写入 chart_size_report.json`);
    await browser.close();
})().catch(e => { console.error('FAIL:', e.message); process.exit(1); });

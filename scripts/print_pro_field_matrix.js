// 线上 Pro 打印 · 字段级矩阵（与 SJS printInfo 逐字段对位）
// 读回通道：Facade 无 getter，但 BeforeSheetPrintConfirm 事件载荷带
//          { layoutConfig, renderConfig } 全量有效配置 → 设一个字段、print() 一次、读回该字段、cancel 掉不真打。
// 判定：PASS = 设置被接受且读回一致；FAIL = 抛错 / 读回不一致；N/A = 该版本无此字段。
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const URL = 'https://office.univer.ai/zh-CN/playground/sheets/charts';
const OUT = path.join(__dirname, 'out', 'pro-print');
fs.mkdirSync(OUT, { recursive: true });

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await (await browser.newContext({ viewport: { width: 1680, height: 1050 } })).newPage();
  const errs = [];
  page.on('pageerror', (e) => errs.push(String(e).slice(0, 160)));

  const go = async () => {
    await page.goto(URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 45; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 60; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
    return false;
  };
  if (!await go()) { console.log('PRO ENV NOT READY'); await browser.close(); process.exit(1); }
  await page.waitForTimeout(2500);

  const result = await page.evaluate(async () => {
    const u = window.univerAPI;
    const E = u.Enum;
    const wb = u.getActiveWorkbook();
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    const J = (v) => { try { return v === undefined ? undefined : JSON.parse(JSON.stringify(v)); } catch (e) { return String(v); } };

    // ---- 读回通道：挂一次事件，每次 print() 后取 window.__cap ----
    let capCount = 0;
    u.addEvent(u.Event.BeforeSheetPrintConfirm, (e) => {
      window.__cap = { layout: J(e.layoutConfig), render: J(e.renderConfig) };
      capCount++;
      e.cancel = true; // 取消，不进实际渲染
    });

    const readback = async (kind) => {
      window.__cap = null;
      try { wb.print(); } catch (e) { return { err: 'print: ' + String(e).slice(0, 120) }; }
      for (let i = 0; i < 30; i++) {
        if (window.__cap) break;
        await wait(100);
      }
      if (!window.__cap) return { err: 'no event payload' };
      return kind === 'render' ? window.__cap.render : window.__cap.layout;
    };

    // 单字段：设置 → print → 读回该字段
    const one = async (kind, label, cfg, field, expect, note) => {
      const setter = kind === 'render' ? 'updatePrintRenderConfig' : 'updatePrintConfig';
      let status = 'PASS', err = '', got;
      try {
        wb[setter](cfg);
        await wait(120);
        const rb = await readback(kind);
        if (rb && rb.err) { status = 'FAIL'; err = rb.err; }
        else {
          got = rb ? rb[field] : undefined;
          // 对象按 key 排序后序列化，避免 {top,right,bottom,left} 与 {top,bottom,left,right} 误判
          const norm = (x) => {
            if (x === undefined || x === null) return String(x);
            if (typeof x !== 'object') return String(x);
            const sort = (v) => Array.isArray(v) ? v.map(sort)
              : v && typeof v === 'object' ? Object.keys(v).sort().reduce((a, k) => (a[k] = sort(v[k]), a), {})
              : v;
            return JSON.stringify(sort(J(x)));
          };
          if (norm(got) !== norm(expect)) { status = 'FAIL'; err = '读回不一致'; }
        }
      } catch (e) { status = 'FAIL'; err = String(e).slice(0, 150); }
      return { kind, label, set: J(cfg), field, expect: norm_(expect), got: norm_(got), status, err, note };
    };
    const norm_ = (x) => {
      if (x === undefined || x === null) return String(x);
      if (typeof x !== 'object') return String(x);
      const sort = (v) => Array.isArray(v) ? v.map(sort)
        : v && typeof v === 'object' ? Object.keys(v).sort().reduce((a, k) => (a[k] = sort(v[k]), a), {})
        : v;
      return JSON.stringify(sort(J(x)));
    };

    const PAPER = Object.keys(E.PrintPaperSize || {});
    const pick = (cands) => cands.find((c) => PAPER.includes(c)) || PAPER[0];

    const rows = [];

    // ================= LAYOUT =================
    rows.push(await one('layout', '打印范围 area = 当前工作表', { area: E.PrintArea.CurrentSheet }, 'area', 'CurrentSheet'));
    rows.push(await one('layout', '打印范围 area = 整个工作簿', { area: E.PrintArea.workbook }, 'area', 'Workbook'));
    rows.push(await one('layout', '打印范围 area = 当前选区', { area: E.PrintArea.CurrentSelection }, 'area', 'CurrentSelection'));
    rows.push(await one('layout', '指定打印区域 subUnitIds.range', {
      area: E.PrintArea.AllSelection,
      subUnitIds: [{ id: wb.getActiveSheet().getSheetId(), range: { startRow: 2, endRow: 20, startColumn: 1, endColumn: 5 } }],
    }, 'subUnitIds', J([{ id: wb.getActiveSheet().getSheetId(), range: { startRow: 2, endRow: 20, startColumn: 1, endColumn: 5 } }]), 'SJS P1/P2 的打印区域对应项'));

    const A3 = pick(['A3']);
    const A4 = pick(['A4']);
    rows.push(await one('layout', '纸张 paperSize = ' + A3, { paperSize: E.PrintPaperSize[A3] }, 'paperSize', A3));
    rows.push(await one('layout', '纸张 paperSize = ' + A4, { paperSize: E.PrintPaperSize[A4] }, 'paperSize', A4));
    rows.push(await one('layout', '自定义纸张 pageSizeCustom 500x700', { pageSizeCustom: { w: 500, h: 700 } }, 'pageSizeCustom', { w: 500, h: 700 }, '实测读回 394x588（固定减 106/112），且会就地改写传入对象'));

    rows.push(await one('layout', '方向 direction = 横向', { direction: E.PrintDirection.Landscape }, 'direction', 'Landscape'));
    rows.push(await one('layout', '方向 direction = 纵向', { direction: E.PrintDirection.Portrait }, 'direction', 'Portrait'));

    rows.push(await one('layout', '缩放 scale = 适应页宽', { scale: E.PrintScale.FitWidth }, 'scale', 'FitWidth', '对位 SJS P13'));
    rows.push(await one('layout', '缩放 scale = 适应页高', { scale: E.PrintScale.FitHeight }, 'scale', 'FitHeight', '对位 SJS P14'));
    rows.push(await one('layout', '缩放 scale = 整页', { scale: E.PrintScale.FitPage }, 'scale', 'FitPage'));
    rows.push(await one('layout', '缩放 scale = 原始', { scale: E.PrintScale.Origin }, 'scale', 'Origin'));
    rows.push(await one('layout', '自定义缩放 scale=Custom + customScale=80', { scale: E.PrintScale.Custom, customScale: 80 }, 'customScale', 80, '对位 SJS P15'));
    rows.push(await one('layout', '自定义缩放越界 customScale=800（看是否钳制）', { scale: E.PrintScale.Custom, customScale: 800 }, 'customScale', 800, '无上限校验：800 原样接受；SJS 传 80 会被钳到 4，两边都不报错'));

    rows.push(await one('layout', '重复行 freeze = [Row]', { freeze: [E.PrintFreeze.Row] }, 'freeze', ['Row'], '对位 SJS P3，但无行号粒度'));
    rows.push(await one('layout', '重复列 freeze = [Column]', { freeze: [E.PrintFreeze.Column] }, 'freeze', ['Column'], '对位 SJS P4'));
    rows.push(await one('layout', '重复行列 freeze = [Row,Column]', { freeze: [E.PrintFreeze.Row, E.PrintFreeze.Column] }, 'freeze', ['Row', 'Column']));

    rows.push(await one('layout', '页边距 margin = 窄', { margin: E.PrintPaperMargin.Narrow }, 'margin', 'Narrow'));
    rows.push(await one('layout', '页边距 margin = 宽', { margin: E.PrintPaperMargin.Wide }, 'margin', 'Wide'));
    rows.push(await one('layout', '页边距 margin = 无', { margin: E.PrintPaperMargin.None }, 'margin', 'None'));
    rows.push(await one('layout', '自定义页边距 margin=Custom + marginCustom', { margin: E.PrintPaperMargin.Custom, marginCustom: { top: 10, right: 12, bottom: 14, left: 16 } }, 'marginCustom', { top: 10, bottom: 14, left: 16, right: 12 }, '对位 SJS P22'));

    rows.push(await one('layout', '每页最大行 maxRowsEachPage = 10', { maxRowsEachPage: 10 }, 'maxRowsEachPage', 10, 'SJS 用 rowEnd/分页符，无同名字段'));
    rows.push(await one('layout', '每页最大列 maxColumnsEachPage = 6', { maxColumnsEachPage: 6 }, 'maxColumnsEachPage', 6));
    rows.push(await one('layout', '每页最大行 = 0（清空）', { maxRowsEachPage: 0 }, 'maxRowsEachPage', 0));

    // ================= RENDER =================
    rows.push(await one('render', '网格线 gridlines = false', { gridlines: false }, 'gridlines', false, '对位 SJS P5'));
    rows.push(await one('render', '网格线 gridlines = true', { gridlines: true }, 'gridlines', true));
    rows.push(await one('render', '水平居中 hAlign = End', { hAlign: E.PrintAlign.End }, 'hAlign', 'End', '对位 SJS P23'));
    rows.push(await one('render', '水平居中 hAlign = Start', { hAlign: E.PrintAlign.Start }, 'hAlign', 'Start'));
    rows.push(await one('render', '垂直居中 vAlign = Middle', { vAlign: E.PrintAlign.Middle }, 'vAlign', 'Middle'));
    rows.push(await one('render', '页眉页脚元素 headerFooter 数组', { headerFooter: [E.PrintHeaderFooter.PageSize, E.PrintHeaderFooter.WorksheetTitle, E.PrintHeaderFooter.Date] }, 'headerFooter', ['PageSize', 'WorksheetTitle', 'Date'], '对位 SJS P18/P19 的枚举式写法'));
    rows.push(await one('render', '自定义页眉页脚 headerFooterSetting 六槽', {
      isCustomHeaderFooter: true,
      headerFooterSetting: { topLeft: 'HL', topCenter: 'HC', topRight: 'HR', bottomLeft: 'FL', bottomCenter: 'FC', bottomRight: 'FR' },
    }, 'headerFooterSetting', { topLeft: 'HL', topCenter: 'HC', topRight: 'HR', bottomLeft: 'FL', bottomCenter: 'FC', bottomRight: 'FR' }, '对位 SJS P18/P19/P21'));
    rows.push(await one('render', '关闭自定义 isCustomHeaderFooter = false 后槽位是否清空', { isCustomHeaderFooter: false }, 'isCustomHeaderFooter', false));
    rows.push(await one('render', '水印 watermark = 字符串', { watermark: 'CONFIDENTIAL' }, 'watermark', 'CONFIDENTIAL', '对位 SJS P27'));
    rows.push(await one('render', '页码占位符 @Page / @TotalPage', {
      isCustomHeaderFooter: true,
      headerFooterSetting: { topLeft: '', topCenter: '', topRight: '', bottomLeft: '', bottomCenter: '@Page / @TotalPage', bottomRight: '' },
    }, 'headerFooterSetting', { topLeft: '', topCenter: '', topRight: '', bottomLeft: '', bottomCenter: '@Page / @TotalPage', bottomRight: '' }, '对位 SJS &P/&N'));

    // ---- 未知字段是否被静默吸收（判定陷阱验证）----
    let bogus = {};
    try {
      wb.updatePrintConfig({ __bogusField: 'X' });
      await wait(100);
      const rb = await readback('layout');
      bogus.layoutAbsorbed = rb && Object.prototype.hasOwnProperty.call(rb, '__bogusField');
      bogus.layoutKeys = rb ? Object.keys(rb) : null;
    } catch (e) { bogus.err = String(e).slice(0, 120); }

    // ---- Facade getter 是否存在（复核旧结论）----
    const getters = ['getPrintConfig', 'getPrintRenderConfig', 'getLayoutConfig', 'getRenderConfig']
      .reduce((a, k) => (a[k] = typeof wb[k], a), {});

    return { rows, bogus, getters, capCount, enums: { paper: Object.keys(E.PrintPaperSize || {}), area: Object.keys(E.PrintArea || {}), scale: Object.keys(E.PrintScale || {}), margin: Object.keys(E.PrintPaperMargin || {}), freeze: Object.keys(E.PrintFreeze || {}), align: Object.keys(E.PrintAlign || {}), hf: Object.keys(E.PrintHeaderFooter || {}) } };
  });

  fs.writeFileSync(path.join(OUT, 'print-matrix-pro.json'), JSON.stringify(result, null, 2));

  const pad = (s, n) => String(s === undefined ? '' : s).padEnd(n);
  console.log('=== UNIVER PRO 打印字段矩阵 ===');
  result.rows.forEach((r) => {
    console.log([r.status.padEnd(4), pad(r.kind, 6), pad(r.label, 42), 'expect=' + pad(r.expect, 26), 'got=' + pad(r.got, 26), r.err].join(' '));
  });
  const cnt = (k) => result.rows.filter((r) => r.kind === k && r.status === 'PASS').length + '/' + result.rows.filter((r) => r.kind === k).length;
  console.log('LAYOUT 通过', cnt('layout'), '| RENDER 通过', cnt('render'), '| 总计', result.rows.filter((r) => r.status === 'PASS').length + '/' + result.rows.length);
  console.log('未知字段被吸收?', JSON.stringify(result.bogus));
  console.log('getter:', JSON.stringify(result.getters));
  console.log('枚举:', JSON.stringify(result.enums));
  console.log('pageerrors:', errs.length, errs.slice(0, 3));

  await browser.close();
})().catch((e) => { console.error('FAILED', e); process.exit(1); });

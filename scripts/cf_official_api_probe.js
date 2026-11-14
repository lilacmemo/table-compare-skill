// 在【官网 playground】里建 dataBar / colorScale / iconSet，判断 Univer 到底支持哪几类条件格式
//   官网示例只展示了 formula 类（错误高亮），没有数据条/色阶/图标集，
//   所以必须在官网同一环境里现建一次，才能区分：
//     a) 官网三类都画 → 我们的构建缺能力（bundle/插件问题）
//     b) 官网只画 formula → Univer 本身不支持这三类（用户观察需修正）
//   用法：NODE_PATH=<poc-compare/node_modules> node tools/probe_cf_official_api.js
const { chromium } = require('playwright');
const path = require('path');

const URL = 'https://office.univer.ai/playground/sheets/conditional-format-error-types';
const ROOT = path.join(__dirname, '..');

(async () => {
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1500, height: 950 }, deviceScaleFactor: 2 });
  const p = await ctx.newPage();
  const errs = [];
  p.on('pageerror', (e) => errs.push(e.message.slice(0, 140)));
  await p.goto(URL, { waitUntil: 'networkidle', timeout: 60000 });
  await p.waitForFunction(() => window.univerAPI && window.univerAPI.getActiveWorkbook && window.univerAPI.getActiveWorkbook(), null, { timeout: 60000 });
  await p.waitForTimeout(3000);

  const r = await p.evaluate(() => {
    const out = {};
    const U = window.univerAPI;
    const sh = U.getActiveWorkbook().getActiveSheet();
    out.sheet = sh.getSheetName();

    // 现有规则（官网示例自带 formula 类型，用来确认 API 结构一致）
    try {
      const cur = sh.getConditionalFormattingRules();
      out.existing = JSON.stringify(cur, (k, v) => (k === 'children' || k === 'parent' ? undefined : v)).slice(0, 800);
      out.existingCount = cur.length;
    } catch (e) { out.eExist = e.message.slice(0, 120); }

    // E 列 5~22 行造数据，分三段挂三类规则
    const vals = Array.from({ length: 18 }, (_, i) => [100 + i * 45]);
    sh.getRange(4, 4, 18, 1).setValues(vals);

    const mk = (r0, r1, type, config, isShowValue) => ({
      ranges: [{ startRow: r0, startColumn: 4, endRow: r1, endColumn: 4, startAbsoluteRefType: 0, endAbsoluteRefType: 0, rangeType: 0 }],
      stopIfTrue: false,
      rule: Object.assign({ type }, isShowValue == null ? {} : { isShowValue }, { config }),
    });
    const cards = [
      ['dataBar', mk(4, 8, 'dataBar', { min: { type: 'min' }, max: { type: 'max' }, isGradient: true, minLength: 10, maxLength: 90, positiveColor: '#638EC6', nativeColor: '#FF0000', negativeSameAsPositive: false }, true)],
      ['colorScale', mk(9, 12, 'colorScale', [{ index: 0, color: '#F8696B', value: { type: 'min' } }, { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } }, { index: 2, color: '#63BE7B', value: { type: 'max' } }], null)],
      ['iconSet', mk(14, 18, 'iconSet', [{ operator: 'greaterThanOrEqual', value: { type: 'percent', value: 67 }, iconType: '3TrafficLights1', iconId: '0' }, { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 33 }, iconType: '3TrafficLights1', iconId: '1' }, { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 0 }, iconType: '3TrafficLights1', iconId: '2' }], true)],
    ];
    out.added = [];
    cards.forEach(([name, rule]) => {
      try { sh.addConditionalFormattingRule(rule); out.added.push({ name, ok: true }); }
      catch (e) { out.added.push({ name, ok: false, err: e.message.slice(0, 120) }); }
    });
    try {
      out.afterCount = sh.getConditionalFormattingRules().length;
    } catch (e) {}
    return out;
  });
  console.log('OFFICIAL', JSON.stringify(r, null, 1));
  console.log('PAGEERRORS', JSON.stringify(errs.slice(0, 5)));

  await p.waitForTimeout(3000);
  await p.screenshot({ path: path.join(ROOT, 'out/shots', '_official-cf-newrule.png') });
  console.log('shot -> out/shots/_official-cf-newrule.png');
  await b.close();
})();

// 条件格式渲染能力的最小验证：新建空簿，A 列造 18 行数字，分三段挂三种规则
//   目的：剥离"导入"这个变量，只看渲染器本身画不画条件格式。
//   - 画得出来 → 我们工程渲染正常，问题在导入还原（模型缺渲染必需字段）
//   - 画不出来 → 这个构建（Univer Pro 1.3.1 + 我们的 preset）根本不渲染条件格式
// 用法：NODE_PATH=<poc-compare/node_modules> node tools/probe_cf_blank.js
const { chromium } = require('playwright');
const path = require('path');

const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html?autoload=0';
const ROOT = path.join(__dirname, '..');

(async () => {
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1680, height: 1000 } });
  const p = await ctx.newPage();
  const errs = [];
  p.on('pageerror', (e) => errs.push(e.message.slice(0, 140)));
  await p.goto(BASE, { waitUntil: 'networkidle' });
  await p.waitForFunction(() => window.APP && APP.booted === true, null, { timeout: 90000 });
  await p.waitForTimeout(2500);

  const r = await p.evaluate(() => {
    const out = { step: [] };
    const U = window.univerAPI;
    const wb = U.createWorkbook();
    const sh = wb.getActiveSheet();

    // A 列 3~8 数据条 / 10~15 色阶 / 17~22 图标集，值域 100~1000
    const vals = Array.from({ length: 20 }, (_, i) => [100 + i * 45]);
    try { sh.getRange(2, 0, 20, 1).setValues(vals); out.step.push('A 列造数 ok'); } catch (e) { out.step.push('造数失败 ' + e.message); }
    // 顺便在 B 列放个对照文字，确认这列没被什么样式盖住
    try { sh.getRange(2, 1, 20, 1).setValues(vals.map((v) => ['行 ' + v[0]])); } catch (e) {}

    const mk = (r0, r1, type, config, isShowValue) => ({
      ranges: [{ startRow: r0, startColumn: 0, endRow: r1, endColumn: 0, startAbsoluteRefType: 0, endAbsoluteRefType: 0, rangeType: 0 }],
      stopIfTrue: false,
      rule: Object.assign({ type }, isShowValue == null ? {} : { isShowValue }, { config }),
    });
    const cards = [
      ['dataBar', mk(2, 7, 'dataBar', { min: { type: 'min' }, max: { type: 'max' }, isGradient: true, minLength: 10, maxLength: 90, positiveColor: '#638EC6', nativeColor: '#FF0000', negativeSameAsPositive: false }, true)],
      ['colorScale', mk(9, 14, 'colorScale', [{ index: 0, color: '#F8696B', value: { type: 'min' } }, { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } }, { index: 2, color: '#63BE7B', value: { type: 'max' } }], null)],
      ['iconSet', mk(16, 21, 'iconSet', [{ operator: 'greaterThanOrEqual', value: { type: 'percent', value: 67 }, iconType: '3TrafficLights1', iconId: '0' }, { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 33 }, iconType: '3TrafficLights1', iconId: '1' }, { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 0 }, iconType: '3TrafficLights1', iconId: '2' }], true)],
    ];
    out.added = [];
    cards.forEach(([name, rule]) => {
      try { sh.addConditionalFormattingRule(rule); out.added.push({ name, ok: true }); }
      catch (e) { out.added.push({ name, ok: false, err: e.message.slice(0, 120) }); }
    });
    out.rulesBack = (sh.getConditionalFormattingRules() || []).map((x) => ({ id: x.cfId, type: x.rule && x.rule.type, ranges: (x.ranges || []).length }));
    out.sheet = sh.getSheetName();
    return out;
  });
  console.log('BLANK', JSON.stringify(r, null, 1));
  console.log('PAGEERRORS', JSON.stringify(errs.slice(0, 6)));

  await p.waitForTimeout(2500);
  await p.screenshot({ path: path.join(ROOT, 'out/shots', '_cf-uni-blank.png') });
  console.log('shot -> out/shots/_cf-uni-blank.png');
  await b.close();
})();

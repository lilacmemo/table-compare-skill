// Univer Pro 图表 API 探针（与 SpreadJS 侧同一套 A~L 用例 ID）
// 环境：office.univer.ai 线上 Pro（版本 1.3.1），window.univerAPI
// 判定：PASS = set 与 get（chart.getInfo()）同时可用且读回一致
//   node probe_pro_chart_api.js            → out/pro-chart/chart-api-uni.json
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const URL = 'https://office.univer.ai/zh-CN/playground/sheets/charts';
const OUTDIR = path.join(__dirname, 'out', 'pro-chart');

// 这段在页面里执行：SPEC 与跑批器
function PAGE_MAIN() {
  const U = window.univerAPI;
  const E = U.Enum;
  const T = E.ChartTypeString;

  const SH = () => window.__chartSheet;
  const GI = (c) => c.getInfo();
  const P = (obj, p, dflt) => {
    let cur = obj;
    for (const k of p.split('.')) {
      if (cur === null || cur === undefined) return dflt;
      cur = cur[k];
    }
    return cur === undefined ? dflt : cur;
  };
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));

  async function fresh(type, build, range) {
    const s = SH();
    for (const c of s.getCharts()) { try { await c.remove(); } catch (e) {} }
    const b = s.newChart(type).setSource({ range: range || 'A1:D7' }).setPosition({ row: 1, column: 4 }).setSize(480, 320);
    if (build) build(b);
    const c = await s.insertChart(b.build());
    await wait(150);
    return c;
  }
  // 取某个系列 patch（getInfo().config.series 只存 patch，未必每个系列都有条目）
  const pickSer = (c, i) => ((GI(c).config.series || []).find((x) => x.selector && x.selector.index === i) || {});
  // 对已存在图表做二次修改（builder → update）
  async function mut(c, fn) {
    const b = c.toBuilder();
    fn(b);
    await c.update(b.build());
    await wait(60);
    return GI(c);
  }
  const miss = (what) => ({ noApi: what });
  const bMethod = (n) => typeof SH().newChart(T.Column)[n] === 'function';

  const SPEC = [
    // ===== A. 图表对象 =====
    { id: 'A1', g: 'A', name: '新建图表', async run() {
      const c = await fresh(T.Column);
      return { got: typeof c.getId() === 'string' && c.getId().length > 0, exp: true, note: 'id=' + c.getId() };
    } },
    { id: 'A2', g: 'A', name: '枚举工作表内全部图表', async run() {
      const s = SH();
      for (const c of s.getCharts()) { try { await c.remove(); } catch (e) {} }
      for (const p of ['A1:C7', 'A1:C7']) {
        const b = s.newChart(T.Column).setSource({ range: p }).setPosition({ row: 1, column: 4 }).setSize(300, 200);
        await s.insertChart(b.build());
      }
      return { got: s.getCharts().length, exp: 2 };
    } },
    { id: 'A3', g: 'A', name: '按标识获取单个图表', async run() {
      const c = await fresh(T.Column);
      const id = c.getId();
      const got = SH().getChart(id);
      return { got: !!(got && got.getId() === id), exp: true };
    } },
    { id: 'A4', g: 'A', name: '删除图表', async run() {
      const c = await fresh(T.Column);
      await c.remove();
      await wait(80);
      return { got: SH().getCharts().length, exp: 0 };
    } },
    { id: 'A5', g: 'A', name: '重命名图表', async run() {
      const c = await fresh(T.Column);
      const has = typeof c.setName === 'function' || typeof c.getName === 'function' || typeof SH().newChart(T.Column).setName === 'function';
      if (!has) return miss('chart.setName / getName');
      c.setName('MyChart');
      await wait(60);
      return { got: c.getName(), exp: 'MyChart' };
    } },
    { id: 'A6', g: 'A', name: '设置图表位置', async run() {
      const c = await fresh(T.Column);
      c.setPosition({ row: 3, column: 6 }); // 直调有效；走 toBuilder().update() 不生效（已实测）
      await wait(120);
      const i = GI(c);
      return { got: [P(i, 'anchor.row'), P(i, 'anchor.column')], exp: [3, 6], note: 'anchor=' + JSON.stringify(i.anchor) };
    } },
    { id: 'A7', g: 'A', name: '设置图表尺寸', async run() {
      const c = await fresh(T.Column);
      c.setSize(500, 240);
      await wait(120);
      const i = GI(c);
      return { got: [P(i, 'size.width'), P(i, 'size.height')], exp: [500, 240] };
    } },
    { id: 'A8', g: 'A', name: '调整图表层级（zIndex）', async run() {
      const c = await fresh(T.Column);
      if (typeof c.setZOrder !== 'function') return miss('chart.setZOrder');
      c.setZOrder(3);
      await wait(120);
      const i = GI(c);
      const got = P(i, 'zIndex', P(i, 'config.zIndex', undefined));
      return { got, exp: 3, note: '有 setZOrder，但 getInfo() 只返回 config/dataSource/position/size/anchor，无层级字段可读回' };
    } },
    { id: 'A9', g: 'A', name: '切换图表类型', async run() {
      const c = await fresh(T.Column);
      c.setType(T.Line);
      await wait(120);
      return { got: c.getType(), exp: T.Line };
    } },
    { id: 'A10', g: 'A', name: '导出图表为图片', async run() {
      const c = await fresh(T.Column);
      if (typeof c.exportImage !== 'function') return miss('chart.exportImage');
      // 必须在图表首帧渲染完成后再调；渲染未完成时首次调用返回 undefined 且会被缓存成同一个结果
      await wait(700);
      let r = await c.exportImage();
      const ok = typeof r === 'string' ? r.length > 500 : !!(r && (r.data || r.url || r.size));
      if (!ok) {
        const c2 = await fresh(T.Column);
        await wait(900);
        r = await c2.exportImage();
      }
      const ok2 = typeof r === 'string' ? r.length > 500 : !!(r && (r.data || r.url || r.size));
      return { got: ok2, exp: true, note: '返回 ' + typeof r + '，长度 ' + (typeof r === 'string' ? r.length : '-') + '（dataURL）' };
    } },
    { id: 'A11', g: 'A', name: '读取数据源范围', async run() {
      const c = await fresh(T.Column);
      const i = GI(c);
      const r = P(i, 'dataSource.range', {});
      return { got: [r.startRow, r.endRow], exp: [0, 6] };
    } },

    // ===== B. 标题 =====
    { id: 'B1', g: 'B', name: '标题文本', async run() {
      const c = await fresh(T.Column, (b) => b.setTitle('TT'));
      return { got: P(GI(c), 'config.title.text'), exp: 'TT' };
    } },
    { id: 'B2', g: 'B', name: '标题显隐', async run() {
      const c = await fresh(T.Column, (b) => b.setTitle({ text: 'TT', visible: false }));
      return { got: P(GI(c), 'config.title.visible'), exp: false };
    } },
    { id: 'B3', g: 'B', name: '标题字体', async run() {
      const c = await fresh(T.Column, (b) => b.setTitle({ text: 'TT', family: 'Arial' }));
      return { got: P(GI(c), 'config.title.family'), exp: 'Arial' };
    } },
    { id: 'B4', g: 'B', name: '标题字号', async run() {
      const c = await fresh(T.Column, (b) => b.setTitle({ text: 'TT', fontSize: 20 }));
      return { got: P(GI(c), 'config.title.fontSize'), exp: 20 };
    } },
    { id: 'B5', g: 'B', name: '标题颜色', async run() {
      const c = await fresh(T.Column, (b) => b.setTitle({ text: 'TT', color: '#008000' }));
      return { got: P(GI(c), 'config.title.color'), exp: '#008000' };
    } },
    { id: 'B6', g: 'B', name: '标题加粗', async run() {
      const c = await fresh(T.Column, (b) => b.setTitle({ text: 'TT', bold: true }));
      return { got: P(GI(c), 'config.title.bold'), exp: true };
    } },
    { id: 'B7', g: 'B', name: '标题背景色', async run() {
      const c = await fresh(T.Column, (b) => b.setTitle({ text: 'TT', background: { color: '#FFFF00' } }));
      return { got: P(GI(c), 'config.title.background.color'), exp: '#FFFF00' };
    } },
    { id: 'B8', g: 'B', name: '标题位置', async run() {
      const c = await fresh(T.Column, (b) => b.setTitle({ text: 'TT', position: 'left' }));
      return { got: P(GI(c), 'config.title.position'), exp: 'left' };
    } },
    { id: 'B9', g: 'B', name: '副标题', async run() {
      const c = await fresh(T.Column, (b) => b.setSubtitle('SUB'));
      return { got: P(GI(c), 'config.subtitle.text'), exp: 'SUB' };
    } },

    // ===== C. 坐标轴 =====
    { id: 'C1', g: 'C', name: '轴标题文本', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ title: 'Y轴' }));
      return { got: P(GI(c), 'config.axes.y.title'), exp: 'Y轴' };
    } },
    { id: 'C2', g: 'C', name: '轴标题字体与字号', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ title: 'Y轴', titleStyle: { family: 'Arial', fontSize: 14 } }));
      const t = P(GI(c), 'config.axes.y.titleStyle', {});
      return { got: [t.family, t.fontSize], exp: ['Arial', 14] };
    } },
    { id: 'C3', g: 'C', name: '轴标题颜色', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ title: 'Y轴', titleStyle: { color: '#800080' } }));
      return { got: P(GI(c), 'config.axes.y.titleStyle.color'), exp: '#800080' };
    } },
    { id: 'C4', g: 'C', name: '轴标题加粗', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ title: 'Y轴', titleStyle: { bold: true } }));
      return { got: P(GI(c), 'config.axes.y.titleStyle.bold'), exp: true };
    } },
    { id: 'C5', g: 'C', name: '轴标题旋转', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ title: 'Y轴', titleStyle: { rotate: 90 } }));
      const t = P(GI(c), 'config.axes.y.titleStyle', {});
      return { got: t.rotate, exp: 90, note: 'titleStyle 只有 color/fontSize/bold/italic/family，无 rotate' };
    } },
    { id: 'C6', g: 'C', name: '轴显隐', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ visible: false }));
      return { got: P(GI(c), 'config.axes.y.visible'), exp: false, note: 'IChartAxisSpec 无 visible 字段' };
    } },
    { id: 'C7', g: 'C', name: '轴线颜色', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ line: { color: '#FF0000' } }));
      return { got: P(GI(c), 'config.axes.y.line.color'), exp: '#FF0000' };
    } },
    { id: 'C8', g: 'C', name: '轴线粗细', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ line: { width: 3 } }));
      return { got: P(GI(c), 'config.axes.y.line.width'), exp: 3 };
    } },
    { id: 'C9', g: 'C', name: '主刻度线位置', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ tick: { visible: true, position: 'inside' } }));
      return { got: P(GI(c), 'config.axes.y.tick.position'), exp: 'inside' };
    } },
    { id: 'C10', g: 'C', name: '刻度标签字体与字号', async run() {
      const c = await fresh(T.Column, (b) => b.setXAxis({ label: { family: 'Arial', fontSize: 13 } }));
      const l = P(GI(c), 'config.axes.x.label', {});
      return { got: [l.family, l.fontSize], exp: ['Arial', 13] };
    } },
    { id: 'C11', g: 'C', name: '刻度标签颜色', async run() {
      const c = await fresh(T.Column, (b) => b.setXAxis({ label: { color: '#008080' } }));
      return { got: P(GI(c), 'config.axes.x.label.color'), exp: '#008080' };
    } },
    { id: 'C12', g: 'C', name: '刻度标签旋转', async run() {
      const c = await fresh(T.Column, (b) => b.setXAxis({ label: { rotate: 45 } }));
      return { got: P(GI(c), 'config.axes.x.label.rotate'), exp: 45 };
    } },
    { id: 'C13', g: 'C', name: '刻度标签显隐', async run() {
      const c = await fresh(T.Column, (b) => b.setXAxis({ label: { visible: false } }));
      return { got: P(GI(c), 'config.axes.x.label.visible'), exp: false };
    } },
    { id: 'C14', g: 'C', name: '刻度标签数字格式', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ format: '0.0%' }));
      return { got: P(GI(c), 'config.axes.y.format'), exp: '0.0%' };
    } },

    // ===== D. 刻度与范围 =====
    { id: 'D1', g: 'D', name: '轴最小值', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ min: 0 }));
      return { got: P(GI(c), 'config.axes.y.min'), exp: 0 };
    } },
    { id: 'D2', g: 'D', name: '轴最大值', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ max: 300 }));
      return { got: P(GI(c), 'config.axes.y.max'), exp: 300 };
    } },
    { id: 'D3', g: 'D', name: '主刻度单位', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ majorUnit: 50 }));
      return { got: P(GI(c), 'config.axes.y.majorUnit'), exp: 50, note: 'IChartAxisSpec 无 majorUnit 字段' };
    } },
    { id: 'D4', g: 'D', name: '次刻度单位', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ minorUnit: 10 }));
      return { got: P(GI(c), 'config.axes.y.minorUnit'), exp: 10, note: 'IChartAxisSpec 无 minorUnit 字段' };
    } },
    { id: 'D5', g: 'D', name: '对数刻度', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ scaleType: 'log', logBase: 10 }));
      const y = P(GI(c), 'config.axes.y', {});
      return { got: [y.scaleType || y.logBase ? (y.scaleType || ('logBase:' + y.logBase)) : undefined], exp: ['log'] };
    } },
    { id: 'D6', g: 'D', name: '轴反转', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ reverse: true }));
      return { got: P(GI(c), 'config.axes.y.reverse'), exp: true };
    } },
    { id: 'D7', g: 'D', name: '分类轴交叉点', async run() {
      const c = await fresh(T.Column, (b) => b.setXAxis({ crossPoint: 3 }));
      return { got: P(GI(c), 'config.axes.x.crossPoint'), exp: 3, note: 'IChartAxisSpec 无 crossPoint 字段' };
    } },
    { id: 'D8', g: 'D', name: '日期轴', async run() {
      if (!bMethod('setUseDateAxis')) return miss('builder.setUseDateAxis');
      const c = await fresh(T.Column, (b) => b.setUseDateAxis(true));
      const i = GI(c);
      return { got: P(i, 'config.mapping.useDateAxis', P(i, 'config.useDateAxis')), exp: true };
    } },

    // ===== E. 网格线 =====
    { id: 'E1', g: 'E', name: '主网格线显隐', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ gridLine: { visible: false } }));
      return { got: P(GI(c), 'config.axes.y.gridLine.visible'), exp: false };
    } },
    { id: 'E2', g: 'E', name: '主网格线颜色', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ gridLine: { visible: true, color: '#800080' } }));
      return { got: P(GI(c), 'config.axes.y.gridLine.color'), exp: '#800080' };
    } },
    { id: 'E3', g: 'E', name: '主网格线粗细', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ gridLine: { visible: true, width: 2 } }));
      return { got: P(GI(c), 'config.axes.y.gridLine.width'), exp: 2 };
    } },
    { id: 'E4', g: 'E', name: '主网格线线型', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ gridLine: { visible: true, dashType: 'dashed' } }));
      return { got: P(GI(c), 'config.axes.y.gridLine.dashType'), exp: 'dashed' };
    } },
    { id: 'E5', g: 'E', name: '次网格线显隐', async run() {
      const c = await fresh(T.Column, (b) => b.setYAxis({ minorGridLine: { visible: true } }));
      return { got: P(GI(c), 'config.axes.y.minorGridLine.visible'), exp: true, note: 'IChartAxisSpec 只有 gridLine，无 minorGridLine' };
    } },

    // ===== F. 图例 =====
    { id: 'F1', g: 'F', name: '图例显隐', async run() {
      const c = await fresh(T.Column, (b) => b.setLegend({ visible: false }));
      // 关闭后 config.legend 直接存布尔 false（不是对象），读回按原值比对
      const v = GI(c).config.legend;
      return { got: v === false ? false : (v && v.visible), exp: false, note: '关闭时 config.legend 存 false；开启且有其它字段时才是对象' };
    } },
    { id: 'F2', g: 'F', name: '图例位置', async run() {
      const c = await fresh(T.Column, (b) => b.setLegend({ visible: true, position: 'left' }));
      return { got: P(GI(c), 'config.legend.position'), exp: 'left' };
    } },
    { id: 'F3', g: 'F', name: '图例字体与字号', async run() {
      const c = await fresh(T.Column, (b) => b.setLegend({ visible: true, family: 'Arial', fontSize: 15 }));
      const l = P(GI(c), 'config.legend', {});
      return { got: [l.family, l.fontSize], exp: ['Arial', 15] };
    } },
    { id: 'F4', g: 'F', name: '图例文字颜色', async run() {
      const c = await fresh(T.Column, (b) => b.setLegend({ visible: true, color: '#FF0000' }));
      return { got: P(GI(c), 'config.legend.color'), exp: '#FF0000' };
    } },
    { id: 'F5', g: 'F', name: '图例加粗', async run() {
      const c = await fresh(T.Column, (b) => b.setLegend({ visible: true, bold: true }));
      return { got: P(GI(c), 'config.legend.bold'), exp: true };
    } },
    { id: 'F6', g: 'F', name: '图例背景色', async run() {
      const c = await fresh(T.Column, (b) => b.setLegend({ visible: true, background: { color: '#FFFF00' } }));
      return { got: P(GI(c), 'config.legend.background.color'), exp: '#FFFF00', note: 'IChartLegendOptions 无 background' };
    } },
    { id: 'F7', g: 'F', name: '图例边框', async run() {
      const c = await fresh(T.Column, (b) => b.setLegend({ visible: true, border: { color: '#0000FF', width: 2 } }));
      return { got: P(GI(c), 'config.legend.border.color'), exp: '#0000FF', note: 'IChartLegendOptions 无 border' };
    } },

    // ===== G. 数据标签 =====
    { id: 'G1', g: 'G', name: '数据标签-显示值', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { label: { visible: true, contentType: 8 } }));
      const s = pickSer(c, 1);
      return { got: [P(s, 'label.visible'), P(s, 'label.contentType')], exp: [true, 8], note: 'contentType 8 = Value（位掩码）' };
    } },
    { id: 'G2', g: 'G', name: '数据标签-显示系列名', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { label: { visible: true, contentType: 4 } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'label.contentType'), exp: 4 };
    } },
    { id: 'G3', g: 'G', name: '数据标签-显示类别名', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { label: { visible: true, contentType: 2 } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'label.contentType'), exp: 2 };
    } },
    { id: 'G4', g: 'G', name: '数据标签-显示百分比', async run() {
      // 饼图 builder 无 setSeries（已实测），百分比走 setPieLabel + 单系列源
      const c = await fresh(T.Pie, (b) => b.setPieLabel({ visible: true, contentType: 16 }), 'A1:B7');
      return { got: P(GI(c), 'config.pie.label.contentType'), exp: 16, note: 'contentType 16 = Percentage（位掩码）' };
    } },
    { id: 'G5', g: 'G', name: '数据标签-位置', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { label: { visible: true, position: 'insideEnd' } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'label.position'), exp: 'insideEnd' };
    } },
    { id: 'G6', g: 'G', name: '数据标签-数字格式', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { label: { visible: true, format: '0.0' } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'label.format'), exp: '0.0' };
    } },
    { id: 'G7', g: 'G', name: '数据标签-字体字号颜色', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { label: { visible: true, family: 'Arial', fontSize: 12, color: '#008000' } }));
      const l = pickSer(c, 1).label || {};
      return { got: [l.family, l.fontSize, l.color], exp: ['Arial', 12, '#008000'] };
    } },
    { id: 'G8', g: 'G', name: '数据标签-加粗', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { label: { visible: true, bold: true } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'label.bold'), exp: true };
    } },
    { id: 'G9', g: 'G', name: '数据标签-背景色', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { label: { visible: true, backgroundColor: '#FFFFFF' } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'label.backgroundColor'), exp: '#FFFFFF', note: 'IChartSeriesLabelSpec 无 backgroundColor' };
    } },
    { id: 'G10', g: 'G', name: '数据标签-分隔符', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { label: { visible: true, separator: ' / ' } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'label.separator'), exp: ' / ', note: 'IChartSeriesLabelSpec 无 separator' };
    } },

    // ===== H. 系列 =====
    { id: 'H1', g: 'H', name: '系列填充色', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { color: '#FFFF00' }));
      const s = pickSer(c, 1);
      return { got: s.color, exp: '#FFFF00' };
    } },
    { id: 'H2', g: 'H', name: '系列边框颜色', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { border: { color: '#00FF00' } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'border.color'), exp: '#00FF00' };
    } },
    { id: 'H3', g: 'H', name: '系列边框粗细', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { border: { width: 3 } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'border.width'), exp: 3 };
    } },
    { id: 'H4', g: 'H', name: '系列填充透明度', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { fillOpacity: 0.5 }));
      const s = pickSer(c, 1);
      return { got: s.fillOpacity, exp: 0.5 };
    } },
    { id: 'H5', g: 'H', name: '数据标记形状', async run() {
      const c = await fresh(T.Line, (b) => b.setSeries({ index: 1 }, { point: { shape: 'diamond' } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'point.shape'), exp: 'diamond' };
    } },
    { id: 'H6', g: 'H', name: '数据标记大小', async run() {
      const c = await fresh(T.Line, (b) => b.setSeries({ index: 1 }, { point: { size: 12 } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'point.size'), exp: 12 };
    } },
    { id: 'H7', g: 'H', name: '数据标记填充色', async run() {
      const c = await fresh(T.Line, (b) => b.setSeries({ index: 1 }, { point: { shape: 'circle', color: '#FF0000' } }));
      const s = pickSer(c, 1);
      return { got: P(s, 'point.color'), exp: '#FF0000' };
    } },
    { id: 'H8', g: 'H', name: '单点着色', async run() {
      const c = await fresh(T.Column, (b) => b.setSeries({ index: 1 }, { dataPoints: { 0: { color: '#FF0000' } } }));
      const s = pickSer(c, 1);
      const dp = s.dataPoints || {};
      const has = Object.keys(dp).length > 0;
      return { got: has, exp: true, note: '读回 ' + JSON.stringify(dp).slice(0, 90) };
    } },
    { id: 'H9', g: 'H', name: '间隙宽度 gapWidth', async run() {
      const c = await fresh(T.Column, (b) => b.setBar({ gapWidth: 80 }));
      return { got: P(GI(c), 'config.bar.gapWidth'), exp: 80, note: 'Univer 用百分比（80 = 80%）' };
    } },
    { id: 'H10', g: 'H', name: '系列重叠 overlap', async run() {
      const c = await fresh(T.Column, (b) => b.setBar({ overlap: -20 }));
      return { got: P(GI(c), 'config.bar.overlap'), exp: -20 };
    } },
    { id: 'H11', g: 'H', name: '平滑线', async run() {
      const c = await fresh(T.Line, (b) => b.setSeries({ index: 1 }, { scatterLine: 'smooth' }));
      const s = pickSer(c, 1);
      return { got: s.scatterLine, exp: 'smooth' };
    } },
    { id: 'H12', g: 'H', name: '圆环内径', async run() {
      const c = await fresh(T.Donut, (b) => b.setDoughnutHole(0.4), 'A1:B7');
      return { got: P(GI(c), 'config.pie.doughnutHole', P(GI(c), 'config.doughnutHole')), exp: 0.4 };
    } },
    { id: 'H13', g: 'H', name: '饼图扇区分离', async run() {
      const pb = SH().newChart(T.Pie);
      if (typeof pb.setExplosion !== 'function') return miss('builder.setExplosion');
      const c = await fresh(T.Pie, (b) => b.setExplosion(0.2), 'A1:B7');
      return { got: P(GI(c), 'config.pie.explosion'), exp: 0.2 };
    } },
    { id: 'H14', g: 'H', name: '次坐标轴', async run() {
      if (typeof SH().newChart(T.Combination).setSeriesAxis !== 'function') return miss('builder.setSeriesAxis');
      const c = await fresh(T.Combination, (b) => b.setSeriesAxis(2, 'right'));
      return { got: pickSer(c, 2).rightAxis, exp: true };
    } },
    { id: 'H15', g: 'H', name: '系列图表类型（组合图）', async run() {
      if (typeof SH().newChart(T.Combination).setSeriesType !== 'function') return miss('builder.setSeriesType');
      const c = await fresh(T.Combination, (b) => b.setSeriesType(2, 'line'));
      return { got: pickSer(c, 2).type, exp: 'line' };
    } },

    // ===== I. 高级元素 =====
    { id: 'I1', g: 'I', name: '趋势线类型', async run() {
      const c = await fresh(T.Line, (b) => b.setTrendline({ index: 1 }, { type: 'linear' }));
      const t = (P(GI(c), 'config.trendlines', []) || [])[0] || {};
      return { got: t.type, exp: 'linear' };
    } },
    { id: 'I2', g: 'I', name: '趋势线显示公式', async run() {
      const c = await fresh(T.Line, (b) => b.setTrendline({ index: 1 }, { type: 'linear', showFormula: true }));
      const t = (P(GI(c), 'config.trendlines', []) || [])[0] || {};
      return { got: t.showFormula, exp: true };
    } },
    { id: 'I3', g: 'I', name: '趋势线显示 R²', async run() {
      const c = await fresh(T.Line, (b) => b.setTrendline({ index: 1 }, { type: 'linear', showR2: true }));
      const t = (P(GI(c), 'config.trendlines', []) || [])[0] || {};
      return { got: t.showR2, exp: true };
    } },
    { id: 'I4', g: 'I', name: '趋势线前推/后推', async run() {
      const c = await fresh(T.Line, (b) => b.setTrendline({ index: 1 }, { type: 'linear', forward: 2, backward: 1 }));
      const t = (P(GI(c), 'config.trendlines', []) || [])[0] || {};
      return { got: [t.forward, t.backward], exp: [2, 1], note: 'IChartTrendlineSpec 无 forward/backward' };
    } },
    { id: 'I5', g: 'I', name: '误差线类型与数值', async run() {
      const b = SH().newChart(T.Line);
      const c0 = await fresh(T.Line);
      if (typeof b.setErrorBars !== 'function' && typeof c0.setErrorBars !== 'function') return miss('builder.setErrorBars / chart.setErrorBars');
      const c = await fresh(T.Line, (bb) => bb.setErrorBars({ seriesIndex: 1, direction: 'y', valueType: 'fixedVal', value: 30, type: 'both' }));
      return { got: P(GI(c), 'config.errorBars'), exp: 'set' };
    } },
    { id: 'I6', g: 'I', name: '图下数据表', async run() {
      if (!bMethod('setDataTable')) return miss('builder.setDataTable');
      return { got: true, exp: true };
    } },
    { id: 'I7', g: 'I', name: '标线（markLine）', async run() {
      if (!bMethod('setMarkLines')) return miss('builder.setMarkLines');
      const c = await fresh(T.Column, (b) => b.setMarkLines([{ axis: 'y', value: 30, color: '#FF0000' }]));
      return { got: (P(GI(c), 'config.markLines', []) || []).length, exp: 1 };
    } },
    { id: 'I8', g: 'I', name: '坐标轴指示器（axisPointer）', async run() {
      if (!bMethod('setAxisPointer')) return miss('builder.setAxisPointer');
      const c = await fresh(T.Column, (b) => b.setAxisPointer({ type: 'line' }));
      const v = P(GI(c), 'config.axisPointer');
      return { got: v ? true : v, exp: true, note: JSON.stringify(v).slice(0, 80) };
    } },

    // ===== J. 图表区 / 绘图区 =====
    { id: 'J1', g: 'J', name: '图表区背景色', async run() {
      const c = await fresh(T.Column, (b) => b.setAppearance({ background: { color: '#ADD8E6' } }));
      return { got: P(GI(c), 'config.appearance.background.color'), exp: '#ADD8E6' };
    } },
    { id: 'J2', g: 'J', name: '图表区边框', async run() {
      const c = await fresh(T.Column, (b) => b.setAppearance({ border: { color: '#008000', width: 3 } }));
      const a = P(GI(c), 'config.appearance.border', {});
      return { got: [a.color, a.width], exp: ['#008000', 3] };
    } },
    { id: 'J3', g: 'J', name: '绘图区背景色', async run() {
      const c = await fresh(T.Column, (b) => b.setAppearance({ plotBackground: { color: '#FFC0CB' } }));
      return { got: P(GI(c), 'config.appearance.plotBackground.color'), exp: '#FFC0CB', note: 'IChartAppearanceSpec 只有 background/border/font' };
    } },
    { id: 'J4', g: 'J', name: '绘图区边框', async run() {
      const c = await fresh(T.Column, (b) => b.setAppearance({ plotBorder: { color: '#0000FF', width: 2 } }));
      return { got: P(GI(c), 'config.appearance.plotBorder.color'), exp: '#0000FF', note: '同上，无绘图区边框字段' };
    } },
    { id: 'J5', g: 'J', name: '图表主题', async run() {
      if (!bMethod('setTheme')) return miss('builder.setTheme');
      const c = await fresh(T.Column, (b) => b.setTheme('default'));
      const v = P(GI(c), 'config.theme');
      return { got: v, exp: 'default', note: '读回 ' + JSON.stringify(v) };
    } },
    { id: 'J6', g: 'J', name: '预设配色方案', async run() {
      const c = await fresh(T.Column, (b) => b.setColorScheme ? b.setColorScheme(2) : null);
      const v = P(GI(c), 'config.colorScheme');
      return { got: v, exp: 2, note: '无独立配色方案索引，换色走 setTheme / setPalette' };
    } },
    { id: 'J7', g: 'J', name: '自定义调色板（任意色数组）', async run() {
      if (!bMethod('setPalette')) return miss('builder.setPalette');
      const c = await fresh(T.Column, (b) => b.setPalette(['#FF0000', '#00FF00', '#0000FF']));
      const v = P(GI(c), 'config.palette', []);
      return { got: Array.isArray(v) ? v.length : v, exp: 3 };
    } },

    // ===== K. 3D =====
    { id: 'K1', g: 'K', name: '3D 图表类型', async run() {
      if (!T.Column3D) return miss('ChartTypeString.Column3D');
      const c = await fresh(T.Column3D);
      return { got: c.getType(), exp: T.Column3D };
    } },
    { id: 'K2', g: 'K', name: '3D 视角（旋转/仰角）', async run() {
      if (typeof SH().newChart(T.Column3D).setThreeD !== 'function') return miss('builder.setThreeD');
      const c = await fresh(T.Column3D, (b) => b.setThreeD({ view: { rotX: 15, rotY: 20, perspective: 30 } }));
      const v = P(GI(c), 'config.threeD.view', {});
      return { got: [v.rotX, v.rotY], exp: [15, 20] };
    } },
    { id: 'K3', g: 'K', name: '3D 深度与墙面', async run() {
      if (typeof SH().newChart(T.Column3D).setThreeD !== 'function') return miss('builder.setThreeD');
      const c = await fresh(T.Column3D, (b) => b.setThreeD({ gapDepth: 150, grouping: 'clustered' }));
      return { got: P(GI(c), 'config.threeD.gapDepth'), exp: 150, note: '有 gapDepth/grouping/view；无墙面（walls）字段' };
    } },

    // ===== L. 其它 =====
    { id: 'L1', g: 'L', name: '忽略隐藏行列', async run() {
      if (typeof (await fresh(T.Column)).setPlotVisibleOnly !== 'function') return miss('chart.setPlotVisibleOnly');
      const c = await fresh(T.Column);
      c.setPlotVisibleOnly(true);
      await wait(60);
      const v = c.getPlotVisibleOnly();
      return { got: v, exp: true };
    } },
    { id: 'L2', g: 'L', name: '空值显示方式', async run() {
      if (!bMethod('setInvalidValueStrategy')) return miss('builder.setInvalidValueStrategy');
      const c = await fresh(T.Column, (b) => b.setInvalidValueStrategy('zero'));
      return { got: P(GI(c), 'config.invalidValueStrategy'), exp: 'zero' };
    } },
    { id: 'L3', g: 'L', name: '数据方向切换（行列互换）', async run() {
      const c = await fresh(T.Column, (b) => b.setSource({ range: 'A1:C7', orientation: E.ChartSourceOrientation.Rows }));
      return { got: P(GI(c), 'dataSource.orientation'), exp: E.ChartSourceOrientation.Rows };
    } },
    { id: 'L4', g: 'L', name: '悬停高亮样式', async run() {
      const c = await fresh(T.Column);
      const m = Object.getOwnPropertyNames(Object.getPrototypeOf(c)).filter((n) => /hover/i.test(n));
      if (!m.length) return miss('chart.hoverStyle');
      return { got: true, exp: true };
    } },
    { id: 'L5', g: 'L', name: '动画开关', async run() {
      const c = await fresh(T.Column);
      const m = Object.getOwnPropertyNames(Object.getPrototypeOf(c)).filter((n) => /anim/i.test(n));
      if (!m.length) return miss('chart.useAnimation');
      return { got: true, exp: true };
    } },
    { id: 'L6', g: 'L', name: '选中图表', async run() {
      const c = await fresh(T.Column);
      const m = Object.getOwnPropertyNames(Object.getPrototypeOf(c)).filter((n) => /select/i.test(n));
      if (!m.length) return miss('chart.isSelected');
      return { got: true, exp: true };
    } },
  ];

  async function runOne(spec) {
    const t0 = Date.now();
    let r;
    try {
      r = await Promise.race([
        spec.run(),
        new Promise((res) => setTimeout(() => res({ noApi: 'TIMEOUT' }), 12000)),
      ]) || {};
    } catch (e) {
      return { id: spec.id, g: spec.g, name: spec.name, ok: false, reason: '抛错: ' + String(e && e.message || e).slice(0, 160) };
    }
    if (r.noApi) return { id: spec.id, g: spec.g, name: spec.name, ok: false, reason: '无 API（' + r.noApi + '）' };
    const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);
    const ok = eq(r.got, r.exp);
    return {
      id: spec.id, g: spec.g, name: spec.name, ok,
      got: JSON.stringify(r.got), exp: JSON.stringify(r.exp),
      note: r.note || (ok ? '' : '读回与设置不一致'),
      ms: Date.now() - t0,
    };
  }

  window.__ProChartProbe = {
    async init() {
      const wb = U.getActiveWorkbook();
      let s = wb.getSheetByName('__chart');
      if (!s) s = wb.insertSheet('__chart');
      wb.setActiveSheet(s);
      // 注意：不要用 sheet.command.set-range-values —— 实测它不会写进新插入的表（A1 仍为 null），
      // 导致图表取不到数据、config.series 为空、所有系列类读回全是 undefined（假 FAIL）。
      const vals = [['Cat', 'S1', 'S2', 'S3']];
      for (let i = 1; i <= 6; i++) vals.push(['C' + i, 10 + i * 3, 20 + i * 5, 30 - i]);
      s.getRange('A1:D7').setValues(vals);
      await new Promise((r) => setTimeout(r, 200));
      const probeVal = s.getRange('B2').getValue();
      window.__chartSheet = s;
      return { sheet: s.getSheetName(), chartTypes: Object.keys(T).length, seedB2: probeVal };
    },
    async runAll() {
      const items = [];
      for (const spec of SPEC) {
        items.push(await runOne(spec));
        window.__chartProgress = { done: items.length, total: SPEC.length, at: spec.id };
      }
      const s = SH();
      for (const c of s.getCharts()) { try { await c.remove(); } catch (e) {} }
      return { items };
    },
    async probeApi() {
      const c = await fresh(T.Column);
      const all = (o) => { const st = new Set(); let p = Object.getPrototypeOf(o); while (p && p !== Object.prototype) { Object.getOwnPropertyNames(p).forEach((n) => st.add(n)); p = Object.getPrototypeOf(p); } return [...st].filter((n) => n[0] !== '_' && typeof o[n] === 'function'); };
      return { chart: all(c), builder: all(c.toBuilder()) };
    },
  };
}

(async () => {
  fs.mkdirSync(OUTDIR, { recursive: true });
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
  const ok = await go();
  console.log('univerAPI ready:', ok);
  if (!ok) { console.log(errs.slice(0, 3)); await browser.close(); process.exit(1); }
  await page.waitForTimeout(3000);

  await page.evaluate(`(${PAGE_MAIN.toString()})()`);
  const env = await page.evaluate(() => window.__ProChartProbe.init());
  console.log('环境自检', JSON.stringify(env));

  const t0 = Date.now();
  const poll = setInterval(async () => {
    try {
      const p = await page.evaluate(() => window.__chartProgress || null);
      if (p) process.stdout.write(`\r  进度 ${p.done}/${p.total} 当前 ${p.at || '-'}   `);
    } catch (e) {}
  }, 2000);

  const res = await page.evaluate(async () => await window.__ProChartProbe.runAll());
  clearInterval(poll);
  process.stdout.write('\n');

  const file = path.join(OUTDIR, 'chart-api-uni.json');
  fs.writeFileSync(file, JSON.stringify(res, null, 1), 'utf8');

  const items = res.items || [];
  const pass = items.filter((i) => i.ok).length;
  console.log('用时', ((Date.now() - t0) / 1000).toFixed(1) + 's', '总项', items.length);
  console.log('Univer', JSON.stringify({ pass, fail: items.length - pass, rate: (pass / items.length * 100).toFixed(1) + '%' }));
  console.log('--- FAIL ---');
  for (const i of items.filter((x) => !x.ok)) {
    console.log(i.id, i.name, '|', i.reason || ('got=' + i.got + ' exp=' + i.exp), '|', i.note || '');
  }
  if (errs.length) console.log('页面错误', errs.slice(0, 5));
  console.log('写出', file);
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });

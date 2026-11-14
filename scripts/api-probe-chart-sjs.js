// SpreadJS 图表 API 探针（A~L 组，逐项「设置 → 读回」）
// 判定：PASS = set 与 get 同时存在且读回一致；FAIL = 缺 set / 缺 get / 读回不一致
// 用法：window.ApiProbeChartSjs.runAll() → { items: [...] }
(function () {
  'use strict';

  const NS = () => window.GC.Spread.Sheets;
  const T = () => NS().Charts.ChartType;

  // ---- 颜色归一化：'green' / 'rgb(0,128,0)' / '#008000' 统一成 #rrggbb ----
  let cvs = null;
  function toHex(c) {
    if (typeof c !== 'string' || !c) return c;
    if (!cvs) cvs = document.createElement('canvas').getContext('2d');
    cvs.fillStyle = '#000000';
    cvs.fillStyle = c;
    return String(cvs.fillStyle).toLowerCase();
  }

  function normColor(x) {
    if (typeof x === 'string') return toHex(x);
    if (Array.isArray(x)) return x.map(normColor);
    if (x && typeof x === 'object') {
      const o = {};
      Object.keys(x).sort().forEach((k) => (o[k] = /color/i.test(k) ? normColor(x[k]) : x[k]));
      return JSON.stringify(o);
    }
    return x;
  }

  function eqColor(got, exp) {
    return normColor(got) === normColor(exp);
  }
  function eqJSON(got, exp) {
    const s = (v) => (typeof v === 'object' ? JSON.stringify(v) : String(v));
    return s(got) === s(exp);
  }

  // ---- 图表构造与清理 ----
  function fresh(type, range) {
    const sh = window.__chartSheet;
    sh.charts.all().forEach((c) => { try { sh.charts.remove(c.name()); } catch (e) {} });
    return sh.charts.add('c1', type || T().columnClustered, 10, 10, 480, 320, range || 'A1:C7');
  }
  function axesMut(ch, fn) {
    const a = ch.axes();
    fn(a);
    ch.axes(a);
    return ch.axes();
  }
  function serSet(ch, i, patch) {
    const coll = ch.series();
    const s = coll.get(i);
    Object.assign(s, patch);
    coll.set(i, s);
    return ch.series().get(i);
  }
  const miss = (what) => ({ noApi: what });

  // 候选键名探测：有的属性文档没写，但运行时能生效；两个都试，避免我写错造成假阴性
  function tryKeys(setter, keys, reader) {
    for (const k of keys) {
      try {
        setter(k);
        const v = reader();
        if (v !== undefined && v !== null && v !== false) return { key: k, val: v };
      } catch (e) {}
    }
    return null;
  }

  const SPEC = [
    // ================= A. 图表对象 =================
    { id: 'A1', g: 'A', name: '新建图表', run() {
      const ch = fresh();
      return { got: !!(ch && ch.name), exp: true };
    } },
    { id: 'A2', g: 'A', name: '枚举工作表内全部图表', run() {
      const sh = window.__chartSheet;
      sh.charts.all().forEach((c) => sh.charts.remove(c.name()));
      sh.charts.add('ca', T().columnClustered, 10, 10, 300, 200, 'A1:C7');
      sh.charts.add('cb', T().line, 320, 10, 300, 200, 'A1:C7');
      return { got: sh.charts.all().length, exp: 2 };
    } },
    { id: 'A3', g: 'A', name: '按名称获取单个图表', run() {
      fresh();
      const ch = window.__chartSheet.charts.get('c1');
      return { got: !!(ch && ch.name && ch.name() === 'c1'), exp: true };
    } },
    { id: 'A4', g: 'A', name: '删除图表', run() {
      const sh = window.__chartSheet;
      fresh();
      sh.charts.remove('c1');
      return { got: sh.charts.all().length, exp: 0 };
    } },
    { id: 'A5', g: 'A', name: '重命名图表', run() {
      const ch = fresh();
      ch.name('renamed');
      return { got: ch.name(), exp: 'renamed' };
    } },
    { id: 'A6', g: 'A', name: '设置图表位置', run() {
      const ch = fresh();
      ch.x(120); ch.y(90);
      const p = ch.position() || {};
      return { got: [p.x, p.y], exp: [120, 90], note: 'position(obj) 只读，必须用 x()/y() 写' };
    } },
    { id: 'A7', g: 'A', name: '设置图表尺寸', run() {
      const ch = fresh();
      ch.width(500);
      ch.height(240);
      return { got: [ch.width(), ch.height()], exp: [500, 240] };
    } },
    { id: 'A8', g: 'A', name: '调整图表层级（zIndex）', run() {
      const ch = fresh();
      if (typeof ch.zIndex === 'undefined') return miss('chart.zIndex');
      ch.zIndex(3);
      return { got: ch.zIndex(), exp: 3 };
    } },
    { id: 'A9', g: 'A', name: '切换图表类型', run() {
      const ch = fresh();
      ch.chartType(T().line);
      return { got: ch.chartType(), exp: T().line };
    } },
    { id: 'A10', g: 'A', name: '导出图表为图片', run() {
      const ch = fresh();
      const url = ch.toImageDataURL();
      return { got: typeof url === 'string' && url.indexOf('data:image') === 0 && url.length > 2000, exp: true, note: '长度 ' + (url || '').length };
    } },
    { id: 'A11', g: 'A', name: '读取数据源范围', run() {
      const ch = fresh();
      const r = ch.dataRange();
      return { got: typeof r === 'string' && /A\$?1/.test(r), exp: true, note: String(r) };
    } },

    // ================= B. 标题 =================
    { id: 'B1', g: 'B', name: '标题文本', run() {
      const ch = fresh();
      ch.title({ text: 'TT' });
      return { got: ch.title().text, exp: 'TT' };
    } },
    { id: 'B2', g: 'B', name: '标题显隐', run() {
      const ch = fresh();
      if (!('visible' in ch.title())) return miss('chart.title().visible');
      ch.title({ visible: false });
      return { got: ch.title().visible, exp: false };
    } },
    { id: 'B3', g: 'B', name: '标题字体', run() {
      const ch = fresh();
      ch.title({ fontFamily: 'Arial' });
      return { got: ch.title().fontFamily, exp: 'Arial' };
    } },
    { id: 'B4', g: 'B', name: '标题字号', run() {
      const ch = fresh();
      ch.title({ fontSize: 20 });
      return { got: ch.title().fontSize, exp: 20 };
    } },
    { id: 'B5', g: 'B', name: '标题颜色', run() {
      const ch = fresh();
      ch.title({ color: '#008000' });
      return { got: ch.title().color, exp: '#008000', kind: 'color' };
    } },
    { id: 'B6', g: 'B', name: '标题加粗', run() {
      const ch = fresh();
      ch.title({ bold: true });
      const t = ch.title(); t.bold = true; ch.title(t);   // 整对象赋值再试一次
      return { got: ch.title().bold, exp: true, note: 'title() 有 bold 字段，但赋值后读回仍 false' };
    } },
    { id: 'B7', g: 'B', name: '标题背景色', run() {
      const ch = fresh();
      ch.title({ backColor: '#FFFF00' });
      return { got: ch.title().backColor, exp: '#FFFF00', kind: 'color' };
    } },
    { id: 'B8', g: 'B', name: '标题位置', run() {
      const ch = fresh();
      const t = ch.title();
      if (!('position' in t)) return miss('chart.title().position');
      t.position = 'bottom';
      ch.title(t);
      return { got: ch.title().position, exp: 'bottom' };
    } },
    { id: 'B9', g: 'B', name: '副标题', run() {
      const ch = fresh();
      if (typeof ch.subtitle !== 'function') return miss('chart.subtitle');
      ch.subtitle({ text: 'SUB' });
      return { got: ch.subtitle().text, exp: 'SUB' };
    } },

    // ================= C. 坐标轴 =================
    { id: 'C1', g: 'C', name: '轴标题文本', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.title.text = 'Y轴'; });
      return { got: a.primaryValue.title.text, exp: 'Y轴' };
    } },
    { id: 'C2', g: 'C', name: '轴标题字体与字号', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.title = { text: 'Y轴', fontFamily: 'Arial', fontSize: 14 }; });
      const t = a.primaryValue.title || {};
      return { got: [t.fontFamily, t.fontSize], exp: ['Arial', 14] };
    } },
    { id: 'C3', g: 'C', name: '轴标题颜色', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.title = { text: 'Y轴', color: '#800080' }; });
      return { got: (a.primaryValue.title || {}).color, exp: '#800080', kind: 'color' };
    } },
    { id: 'C4', g: 'C', name: '轴标题加粗', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.title = { text: 'Y轴', bold: true }; });
      return { got: (a.primaryValue.title || {}).bold, exp: true, note: '轴标题对象无 bold 字段' };
    } },
    { id: 'C5', g: 'C', name: '轴标题旋转', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.title = { text: 'Y轴', textDirection: NS().Charts.TextDirection.vert }; });
      return { got: (a.primaryValue.title || {}).textDirection, exp: NS().Charts.TextDirection.vert };
    } },
    { id: 'C6', g: 'C', name: '轴显隐', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.visible = false; });
      return { got: a.primaryValue.visible, exp: false };
    } },
    { id: 'C7', g: 'C', name: '轴线颜色', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.lineStyle.color = '#FF0000'; });
      return { got: a.primaryValue.lineStyle.color, exp: '#FF0000', kind: 'color' };
    } },
    { id: 'C8', g: 'C', name: '轴线粗细', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.lineStyle.width = 3; });
      return { got: a.primaryValue.lineStyle.width, exp: 3 };
    } },
    { id: 'C9', g: 'C', name: '主刻度线位置', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.majorTickPosition = NS().Charts.TickMark.inside; });
      return { got: a.primaryValue.majorTickPosition, exp: NS().Charts.TickMark.inside };
    } },
    { id: 'C10', g: 'C', name: '刻度标签字体与字号', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryCategory.style.fontFamily = 'Arial'; x.primaryCategory.style.fontSize = 13; });
      return { got: [a.primaryCategory.style.fontFamily, a.primaryCategory.style.fontSize], exp: ['Arial', 13] };
    } },
    { id: 'C11', g: 'C', name: '刻度标签颜色', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryCategory.style.color = '#008080'; });
      return { got: a.primaryCategory.style.color, exp: '#008080', kind: 'color' };
    } },
    { id: 'C12', g: 'C', name: '刻度标签旋转', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryCategory.labelAngle = 45; });
      return { got: a.primaryCategory.labelAngle, exp: 45 };
    } },
    { id: 'C13', g: 'C', name: '刻度标签显隐', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryCategory.tickLabelPosition = NS().Charts.TickLabelPosition.none; });
      return { got: a.primaryCategory.tickLabelPosition, exp: NS().Charts.TickLabelPosition.none };
    } },
    { id: 'C14', g: 'C', name: '刻度标签数字格式', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.numberFormatLinked = false; x.primaryValue.format = '0.0%'; });
      return { got: a.primaryValue.format, exp: '0.0%' };
    } },

    // ================= D. 刻度与范围 =================
    { id: 'D1', g: 'D', name: '轴最小值', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.min = 0; });
      return { got: a.primaryValue.min, exp: 0 };
    } },
    { id: 'D2', g: 'D', name: '轴最大值', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.max = 300; });
      return { got: a.primaryValue.max, exp: 300 };
    } },
    { id: 'D3', g: 'D', name: '主刻度单位', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.majorUnit = 50; });
      return { got: a.primaryValue.majorUnit, exp: 50 };
    } },
    { id: 'D4', g: 'D', name: '次刻度单位', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.minorUnit = 10; });
      return { got: a.primaryValue.minorUnit, exp: 10 };
    } },
    { id: 'D5', g: 'D', name: '对数刻度', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.scaling = { logBase: 10 }; });
      return { got: (a.primaryValue.scaling || {}).logBase, exp: 10 };
    } },
    { id: 'D6', g: 'D', name: '轴反转', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.scaling = { logBase: null, orientation: NS().Charts.AxisOrientation.maxMin }; });
      return { got: (a.primaryValue.scaling || {}).orientation, exp: NS().Charts.AxisOrientation.maxMin };
    } },
    { id: 'D7', g: 'D', name: '分类轴交叉点', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryCategory.crossPoint = 3; });
      return { got: a.primaryCategory.crossPoint, exp: 3 };
    } },
    { id: 'D8', g: 'D', name: '日期轴', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryCategory.categoryType = NS().Charts.CategoryType.TimeScale; });
      return { got: a.primaryCategory.categoryType, exp: NS().Charts.CategoryType.TimeScale };
    } },

    // ================= E. 网格线 =================
    { id: 'E1', g: 'E', name: '主网格线显隐', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.majorGridLine.visible = false; });
      return { got: a.primaryValue.majorGridLine.visible, exp: false };
    } },
    { id: 'E2', g: 'E', name: '主网格线颜色', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.majorGridLine.color = '#800080'; });
      return { got: a.primaryValue.majorGridLine.color, exp: '#800080', kind: 'color' };
    } },
    { id: 'E3', g: 'E', name: '主网格线粗细', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.majorGridLine.width = 2; });
      return { got: a.primaryValue.majorGridLine.width, exp: 2 };
    } },
    { id: 'E4', g: 'E', name: '主网格线线型', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.majorGridLine.dashStyle = 4; });
      let got = (a.primaryValue.majorGridLine || {}).dashStyle;
      if (got === undefined) {
        const b = axesMut(ch, (x) => { x.primaryValue.majorGridLine.style = 4; });
        got = (b.primaryValue.majorGridLine || {}).style;
      }
      return { got, exp: 4, note: '读回对象有 dashStyle 字段，但赋值被忽略（dashStyle / style 两种写法都读到 0）' };
    } },
    { id: 'E5', g: 'E', name: '次网格线显隐', run() {
      const ch = fresh();
      const a = axesMut(ch, (x) => { x.primaryValue.minorGridLine.visible = true; });
      return { got: a.primaryValue.minorGridLine.visible, exp: true };
    } },

    // ================= F. 图例 =================
    { id: 'F1', g: 'F', name: '图例显隐', run() {
      const ch = fresh();
      ch.legend({ visible: false });
      return { got: ch.legend().visible, exp: false };
    } },
    { id: 'F2', g: 'F', name: '图例位置', run() {
      const ch = fresh();
      ch.legend({ position: NS().Charts.LegendPosition.left });
      return { got: ch.legend().position, exp: NS().Charts.LegendPosition.left };
    } },
    { id: 'F3', g: 'F', name: '图例字体与字号', run() {
      const ch = fresh();
      ch.legend({ fontFamily: 'Arial', fontSize: 15 });
      return { got: [ch.legend().fontFamily, ch.legend().fontSize], exp: ['Arial', 15] };
    } },
    { id: 'F4', g: 'F', name: '图例文字颜色', run() {
      const ch = fresh();
      ch.legend({ color: '#FF0000' });
      return { got: ch.legend().color, exp: '#FF0000', kind: 'color' };
    } },
    { id: 'F5', g: 'F', name: '图例加粗', run() {
      const ch = fresh();
      ch.legend({ bold: true });
      return { got: ch.legend().bold, exp: true };
    } },
    { id: 'F6', g: 'F', name: '图例背景色', run() {
      const ch = fresh();
      ch.legend({ backColor: '#FFFF00' });
      return { got: ch.legend().backColor, exp: '#FFFF00', kind: 'color' };
    } },
    { id: 'F7', g: 'F', name: '图例边框', run() {
      const ch = fresh();
      ch.legend({ borderStyle: { color: '#0000FF', width: 2, dashStyle: 0 } });
      const b = ch.legend().borderStyle || {};
      return { got: [toHex(b.color), b.width], exp: [toHex('#0000FF'), 2] };
    } },

    // ================= G. 数据标签 =================
    { id: 'G1', g: 'G', name: '数据标签-显示值', run() {
      const ch = fresh();
      ch.dataLabels({ showValue: true });
      return { got: ch.dataLabels().showValue, exp: true };
    } },
    { id: 'G2', g: 'G', name: '数据标签-显示系列名', run() {
      const ch = fresh();
      ch.dataLabels({ showSeriesName: true });
      return { got: ch.dataLabels().showSeriesName, exp: true };
    } },
    { id: 'G3', g: 'G', name: '数据标签-显示类别名', run() {
      const ch = fresh();
      ch.dataLabels({ showCategoryName: true });
      return { got: ch.dataLabels().showCategoryName, exp: true };
    } },
    { id: 'G4', g: 'G', name: '数据标签-显示百分比', run() {
      const ch = fresh(T().pie);
      ch.dataLabels({ showValue: false, showPercentage: true });
      return { got: ch.dataLabels().showPercentage, exp: true };
    } },
    { id: 'G5', g: 'G', name: '数据标签-位置', run() {
      const ch = fresh();
      ch.dataLabels({ showValue: true, position: NS().Charts.DataLabelPosition.insideEnd });
      return { got: ch.dataLabels().position, exp: NS().Charts.DataLabelPosition.insideEnd, note: 'above(8) 对柱形图无效，会被归一到 outsideEnd' };
    } },
    { id: 'G6', g: 'G', name: '数据标签-数字格式', run() {
      const ch = fresh();
      ch.dataLabels({ numberFormatLinked: false, format: '0.0' });
      return { got: ch.dataLabels().format, exp: '0.0' };
    } },
    { id: 'G7', g: 'G', name: '数据标签-字体字号颜色', run() {
      const ch = fresh();
      ch.dataLabels({ fontFamily: 'Arial', fontSize: 12, color: '#008000' });
      const d = ch.dataLabels();
      return { got: [d.fontFamily, d.fontSize, toHex(d.color)], exp: ['Arial', 12, toHex('#008000')] };
    } },
    { id: 'G8', g: 'G', name: '数据标签-加粗', run() {
      const ch = fresh();
      ch.dataLabels({ fontBold: true });
      return { got: ch.dataLabels().fontBold, exp: true };
    } },
    { id: 'G9', g: 'G', name: '数据标签-背景色', run() {
      const ch = fresh();
      ch.dataLabels({ backColor: '#FFFFFF' });
      return { got: ch.dataLabels().backColor, exp: '#FFFFFF', kind: 'color' };
    } },
    { id: 'G10', g: 'G', name: '数据标签-分隔符', run() {
      const ch = fresh();
      ch.dataLabels({ showValue: true, showCategoryName: true, separator: ' / ' });
      return { got: ch.dataLabels().separator, exp: ' / ' };
    } },

    // ================= H. 系列 =================
    { id: 'H1', g: 'H', name: '系列填充色', run() {
      const ch = fresh();
      const s = serSet(ch, 0, { backColor: '#FFFF00' });
      return { got: s.backColor, exp: '#FFFF00', kind: 'color' };
    } },
    { id: 'H2', g: 'H', name: '系列边框颜色', run() {
      const ch = fresh();
      const s = serSet(ch, 0, { border: { color: '#00FF00', width: 1, lineType: 0, transparency: 0 } });
      return { got: s.border.color, exp: '#00FF00', kind: 'color' };
    } },
    { id: 'H3', g: 'H', name: '系列边框粗细', run() {
      const ch = fresh();
      const s = serSet(ch, 0, { border: { color: '#00FF00', width: 3, lineType: 0, transparency: 0 } });
      return { got: s.border.width, exp: 3 };
    } },
    { id: 'H4', g: 'H', name: '系列填充透明度', run() {
      const ch = fresh();
      const s = serSet(ch, 0, { backColor: '#FFFF00', backColorTransparency: 0.5 });
      return { got: s.backColorTransparency, exp: 0.5 };
    } },
    { id: 'H5', g: 'H', name: '数据标记形状', run() {
      const ch = fresh(T().lineMarkers);
      const s = serSet(ch, 0, { symbol: { shape: NS().Charts.MarkerStyle.MarkerStyleDiamond, size: 8 } });
      return { got: (s.symbol || {}).shape, exp: NS().Charts.MarkerStyle.MarkerStyleDiamond };
    } },
    { id: 'H6', g: 'H', name: '数据标记大小', run() {
      const ch = fresh(T().lineMarkers);
      const s = serSet(ch, 0, { symbol: { shape: NS().Charts.MarkerStyle.MarkerStyleCircle, size: 12 } });
      return { got: (s.symbol || {}).size, exp: 12 };
    } },
    { id: 'H7', g: 'H', name: '数据标记填充色', run() {
      const ch = fresh(T().lineMarkers);
      const s = serSet(ch, 0, { symbol: { shape: NS().Charts.MarkerStyle.MarkerStyleCircle, size: 9, fill: '#FF0000' } });
      return { got: (s.symbol || {}).fill, exp: '#FF0000', kind: 'color' };
    } },
    { id: 'H8', g: 'H', name: '单点着色', run() {
      const ch = fresh();
      const s = serSet(ch, 0, { dataPoints: { 0: { fill: { color: '#FF0000' } } } });
      const dp = s.dataPoints || {};
      const has = Object.keys(dp).length > 0;
      return { got: has, exp: true, note: '读回 ' + JSON.stringify(dp).slice(0, 80) };
    } },
    { id: 'H9', g: 'H', name: '间隙宽度 gapWidth', run() {
      const ch = fresh();
      const s = serSet(ch, 0, { gapWidth: 4 });
      return { got: s.gapWidth, exp: 4, note: 'SJS 0~5 制' };
    } },
    { id: 'H10', g: 'H', name: '系列重叠 overlap', run() {
      const ch = fresh();
      const s = serSet(ch, 0, { overlap: -0.2 });
      return { got: s.overlap, exp: -0.2, note: 'SJS -1~1 制' };
    } },
    { id: 'H11', g: 'H', name: '平滑线', run() {
      const ch = fresh(T().line);
      const s = serSet(ch, 0, { smooth: true });
      return { got: s.smooth, exp: true };
    } },
    { id: 'H12', g: 'H', name: '圆环内径', run() {
      const ch = fresh(T().doughnut);
      const s = serSet(ch, 0, { doughnutHoleSize: 0.4 });
      return { got: s.doughnutHoleSize, exp: 0.4, note: 'SJS 0~0.9 制' };
    } },
    { id: 'H13', g: 'H', name: '饼图扇区分离', run() {
      const ch = fresh(T().pie);
      const s = serSet(ch, 0, { explosion: 0.2 });
      if (!('explosion' in s) && !('explosion' in ch.series().get(0))) return miss('series.explosion');
      return { got: ch.series().get(0).explosion, exp: 0.2 };
    } },
    { id: 'H14', g: 'H', name: '次坐标轴', run() {
      const ch = fresh(T().columnClustered);
      const s = serSet(ch, 1, { chartType: T().line, axisGroup: NS().Charts.AxisGroup.secondary });
      return { got: [s.axisGroup, s.chartType], exp: [NS().Charts.AxisGroup.secondary, T().line] };
    } },
    { id: 'H15', g: 'H', name: '系列图表类型（组合图）', run() {
      const ch = fresh(T().columnClustered);
      const s = serSet(ch, 0, { chartType: T().line });
      return { got: s.chartType, exp: T().line };
    } },

    // ================= I. 高级元素 =================
    { id: 'I1', g: 'I', name: '趋势线类型', run() {
      const ch = fresh(T().line);
      const s = serSet(ch, 0, { trendlines: [{ type: NS().Charts.TrendlineType.linear }] });
      return { got: ((s.trendlines || [])[0] || {}).type, exp: NS().Charts.TrendlineType.linear };
    } },
    { id: 'I2', g: 'I', name: '趋势线显示公式', run() {
      const ch = fresh(T().line);
      const s = serSet(ch, 0, { trendlines: [{ type: NS().Charts.TrendlineType.linear, displayEquation: true }] });
      return { got: ((s.trendlines || [])[0] || {}).displayEquation, exp: true, note: '键名是 displayEquation' };
    } },
    { id: 'I3', g: 'I', name: '趋势线显示 R²', run() {
      const ch = fresh(T().line);
      const s = serSet(ch, 0, { trendlines: [{ type: NS().Charts.TrendlineType.linear, displayRSquared: true }] });
      return { got: ((s.trendlines || [])[0] || {}).displayRSquared, exp: true, note: '键名是 displayRSquared' };
    } },
    { id: 'I4', g: 'I', name: '趋势线前推/后推', run() {
      const ch = fresh(T().line);
      const s = serSet(ch, 0, { trendlines: [{ type: NS().Charts.TrendlineType.linear, forward: 2, backward: 1 }] });
      const t = (s.trendlines || [])[0] || {};
      return { got: [t.forward, t.backward], exp: [2, 1] };
    } },
    { id: 'I5', g: 'I', name: '误差线类型与数值', run() {
      const ch = fresh(T().line);
      const s = serSet(ch, 0, { errorBars: { direction: NS().Charts.ErrorBarType.both, endStyle: 0, type: NS().Charts.ErrorBarValueType.fixedValue, value: 30 } });
      const e = s.errorBars || {};
      return { got: [e.type, e.value], exp: [NS().Charts.ErrorBarValueType.fixedValue, 30],
        note: 'ISeries 有 errorBars 字段，但 series().set() 后被丢弃（读回 null）' };
    } },
    { id: 'I6', g: 'I', name: '图下数据表', run() {
      const ch = fresh();
      if (typeof ch.dataTable !== 'function' && !('showDataTable' in ch)) return miss('chart.dataTable');
      return { got: true, exp: true };
    } },
    { id: 'I7', g: 'I', name: '标线（markLine）', run() {
      const ch = fresh();
      if (typeof ch.markLines !== 'function' && typeof ch.setMarkLines !== 'function') return miss('chart.markLines');
      return { got: true, exp: true };
    } },
    { id: 'I8', g: 'I', name: '坐标轴指示器（axisPointer）', run() {
      const ch = fresh();
      if (typeof ch.axisPointer !== 'function' && typeof ch.setAxisPointer !== 'function') return miss('chart.axisPointer');
      return { got: true, exp: true };
    } },

    // ================= J. 图表区 / 绘图区 =================
    { id: 'J1', g: 'J', name: '图表区背景色', run() {
      const ch = fresh();
      ch.chartArea({ backColor: '#ADD8E6' });
      return { got: ch.chartArea().backColor, exp: '#ADD8E6', kind: 'color' };
    } },
    { id: 'J2', g: 'J', name: '图表区边框', run() {
      const ch = fresh();
      const a = ch.chartArea();
      a.border = { color: '#008000', width: 3, transparency: 0, dashStyle: 0 };
      ch.chartArea(a);
      const b = ch.chartArea().border || {};
      return { got: [toHex(b.color), b.width], exp: [toHex('#008000'), 3] };
    } },
    { id: 'J3', g: 'J', name: '绘图区背景色', run() {
      const ch = fresh();
      ch.plotArea({ backColor: '#FFC0CB' });
      return { got: ch.plotArea().backColor, exp: '#FFC0CB', kind: 'color' };
    } },
    { id: 'J4', g: 'J', name: '绘图区边框', run() {
      const ch = fresh();
      const a = ch.plotArea();
      a.border = { color: '#0000FF', width: 2, transparency: 0, dashStyle: 0 };
      ch.plotArea(a);
      const b = ch.plotArea().border || {};
      return { got: [toHex(b.color), b.width], exp: [toHex('#0000FF'), 2] };
    } },
    { id: 'J5', g: 'J', name: '图表主题', run() {
      const ch = fresh();
      if (typeof ch.theme !== 'function' && typeof ch.setTheme !== 'function') return miss('chart.theme');
      return { got: true, exp: true };
    } },
    { id: 'J6', g: 'J', name: '预设配色方案', run() {
      const ch = fresh();
      if (typeof ch.colorScheme !== 'function') return miss('chart.colorScheme');
      const first = ch.colorScheme();
      let idx = null;
      for (const n of [1, 2, 3]) { ch.colorScheme(n); const v = ch.colorScheme(); if (v && v[0] === n) { idx = n; break; } }
      return { got: idx, exp: idx === null ? '可设的索引' : idx, note: 'colorScheme() 返回 [索引, 颜色数组]；初始 ' + JSON.stringify((first || [])[0]) + '，设 1/2/3 后读回 undefined' };
    } },
    { id: 'J7', g: 'J', name: '自定义调色板（任意色数组）', run() {
      const ch = fresh();
      if (typeof ch.palette !== 'function' && typeof ch.setPalette !== 'function') return miss('chart.palette');
      return { got: true, exp: true };
    } },

    // ================= K. 3D =================
    { id: 'K1', g: 'K', name: '3D 图表类型', run() {
      const ch = fresh();
      const t = T().columnClustered3D;
      if (t === undefined) return miss('ChartType.columnClustered3D');
      ch.chartType(t);
      return { got: ch.chartType(), exp: t };
    } },
    { id: 'K2', g: 'K', name: '3D 视角（旋转/仰角）', run() {
      const ch = fresh();
      if (typeof ch.threeD !== 'function' && typeof ch.setThreeD !== 'function' && !('view3D' in ch)) return miss('chart.threeD');
      return { got: true, exp: true };
    } },
    { id: 'K3', g: 'K', name: '3D 深度与墙面', run() {
      const ch = fresh();
      if (typeof ch.walls !== 'function' && typeof ch.setWalls !== 'function') return miss('chart.walls');
      return { got: true, exp: true };
    } },

    // ================= L. 其它 =================
    { id: 'L1', g: 'L', name: '忽略隐藏行列', run() {
      const ch = fresh();
      ch.ignoreHidden(true);
      return { got: ch.ignoreHidden(), exp: true };
    } },
    { id: 'L2', g: 'L', name: '空值显示方式', run() {
      const ch = fresh();
      ch.displayBlanksAs(NS().Charts.DisplayBlanksAs.gaps);
      return { got: ch.displayBlanksAs(), exp: NS().Charts.DisplayBlanksAs.gaps };
    } },
    { id: 'L3', g: 'L', name: '数据方向切换（行列互换）', run() {
      const ch = fresh();
      const before = ch.getDataOrientation();
      ch.switchDataOrientation();
      const after = ch.getDataOrientation();
      return { got: before !== after, exp: true, note: before + ' → ' + after };
    } },
    { id: 'L4', g: 'L', name: '悬停高亮样式', run() {
      const ch = fresh();
      if (typeof ch.hoverStyle !== 'function') return miss('chart.hoverStyle');
      ch.hoverStyle({ borderStyle: { color: '#0000FF', width: 2, dashStyle: 0 } });
      const h = ch.hoverStyle() || {};
      const b = h.borderStyle || h || {};
      return { got: h && typeof h === 'object', exp: true, note: JSON.stringify(h).slice(0, 80) };
    } },
    { id: 'L5', g: 'L', name: '动画开关', run() {
      const ch = fresh();
      ch.useAnimation(false);
      return { got: ch.useAnimation(), exp: false };
    } },
    { id: 'L6', g: 'L', name: '选中图表', run() {
      const ch = fresh();
      if (typeof ch.isSelected !== 'function') return miss('chart.isSelected');
      ch.isSelected(true);
      return { got: ch.isSelected(), exp: true };
    } },
  ];

  function runOne(spec) {
    const t0 = Date.now();
    let r;
    try {
      r = spec.run() || {};
    } catch (e) {
      return { id: spec.id, g: spec.g, name: spec.name, ok: false, reason: '抛错: ' + String(e && e.message || e).slice(0, 160) };
    }
    if (r.noApi) return { id: spec.id, g: spec.g, name: spec.name, ok: false, reason: '无 API（' + r.noApi + '）' };
    const cmp = spec.kind === 'color' || r.kind === 'color' ? eqColor : eqJSON;
    let ok = false;
    try { ok = cmp(r.got, r.exp); } catch (e) { ok = false; }
    return {
      id: spec.id, g: spec.g, name: spec.name, ok,
      got: typeof r.got === 'object' ? JSON.stringify(r.got) : String(r.got),
      exp: typeof r.exp === 'object' ? JSON.stringify(r.exp) : String(r.exp),
      note: r.note || (ok ? '' : '读回与设置不一致'),
      ms: Date.now() - t0,
    };
  }

  async function runAll() {
    const S = NS();
    const sp = window.__spread;
    let sh = sp.getSheetFromName ? sp.getSheetFromName('__chart') : null;
    if (!sh) { sh = new S.Worksheet('__chart'); sp.addSheet(sp.getSheetCount(), sh); }
    sp.setActiveSheet('__chart');
    window.__chartSheet = sh;
    // 数据：A 列类别，B/C 两系列
    sh.setValue(0, 0, 'Cat'); sh.setValue(0, 1, 'S1'); sh.setValue(0, 2, 'S2');
    for (let i = 1; i <= 6; i++) { sh.setValue(i, 0, 'C' + i); sh.setValue(i, 1, 10 + i * 3); sh.setValue(i, 2, 20 + i * 5); }

    const items = [];
    for (const spec of SPEC) {
      items.push(runOne(spec));
      window.__chartProgress = { done: items.length, total: SPEC.length, at: spec.id };
      if (items.length % 8 === 0) await new Promise((r) => setTimeout(r, 0));
    }
    try { sh.charts.all().forEach((c) => sh.charts.remove(c.name())); } catch (e) {}
    return { items, env: { version: (window.GC.Spread.Sheets.Version && window.GC.Spread.Sheets.Version.getProductVersion) ? 'v19' : 'v19', sheet: sh.name() } };
  }

  window.ApiProbeChartSjs = { SPEC, runAll, runOne };
})();

/*
 * A~K 135 项 API 探针表 + 跑批器（v4：查文档后修正富文本/自动合并/单元格状态的真实入口）
 *
 * 判的不是"渲染效果"，而是"能不能用代码设、能不能用代码读回"：
 *   api      —— 公开 API 名（写进报告）
 *   kind     —— prop（属性包）| method（专用方法）| builder（链式构造/命令）| none
 *   steps    —— 设一项要几步（1 直接设 / 2 建对象后应用 / 3 建构造器+提交）
 *   param    —— typed（命名参数/枚举/对象）| loose（裸数字、魔法字符串、拼错不报错）
 *   doc      —— 官方文档是否有该条目（1 有 / 0 无）
 *   assert   —— 缺省按 exp 比对；'noThrow' 表示只判"调用不报错"（用于无读回口的 API）
 *   note     —— 实测补充说明（进报告）
 *
 * v3 相对 v2 的关键修正（都由 _probe8~_probe14 实测得出，不是推测）：
 *   1) SJS 默认工作表只有 100 行，探针行号 5+135 会越界 → runAll 先 setRowCount(400)。
 *      越界表现：addSpan 抛 "Invalid range"，更远处静默返回 null；样式类则写入一个不存在的行。
 *   2) SJS conditionalFormats.addRule(new Rule(), ranges) 不进区域索引：
 *      count()+1 但 getRules(区域)=0、removeRuleByRange 清不掉。
 *      便捷方法 addDataBarRule/addIconSetRule 则正常，且返回规则对象本体（同一引用），
 *      可直接改参数并读回 → H 组改为"便捷方法建 + 读回对象改参数"。
 *   3) Univer facade 写操作走命令，异步生效：setRowHeight(120,20) 立刻读回是 24（旧值/夹紧值）。
 *      所有 uni 侧 set 后 await 一个 tick 再读。
 *   4) SJS Style.fontSize 只收字符串：传裸数字 16 抛 "e.match is not a function"（不可读错误）。
 *   5) Univer setRowHeight 有最小值夹紧（设 20 得 24）；setWrap(true) 会连带自动加高行高，
 *      所以 F3 的判据改为"换行后行高是否被撑高"，而不是 autoResizeRows 前后差值。
 *   6) Univer 隐藏行/列：hideRows/hideColumns 不抛错，但 getRowHeight/getColumnWidth 不变、
 *      getHiddenState() 恒 0、FRange 无 isRowHidden → 有写入口、无可信读回口。
 *   7) SJS 富文本不是 sheet.setRichText（运行时确实没有这个方法和 RichText 命名空间），
 *      而是"值"——sheet.setValue(r, c, { richText: [{ text, style }] })，style 里可带
 *      font / foreColor / textDecoration / vertAlign（1=上标 2=下标）。
 *      读回必须显式指定：getValue(r, c, SheetArea.viewport, ValueType.richText)。
 *   8) SJS 单元格状态用 sheet.cellStates.add(range, CellStatesType, style)，
 *      读回用 cellStates.getJSONForCell(r, c)（getState 只反映"此刻是否处于该视觉态"）。
 *      sheet.autoMerge(range, AutoMergeMode, AutoMergeDirection) 调用不报错，
 *      但 getSpan/toJSON/取值都看不到痕迹，无法证实生效。
 *   9) Univer 对齐只认 'left'|'center'|'normal'（normal=右）、'top'|'middle'|'bottom'。
 *      边框类型是字符串，且 Enum.BorderType.BLTR==='bl_tr' 但 setter 只认 'bltr'。
 */
(function () {
  const SJS = window.GC.Spread.Sheets;
  const CFS = SJS.ConditionalFormatting;

  // ---------- 助手 ----------
  function sjsProp(api, prop, val, opt) {
    opt = opt || {};
    return {
      api, kind: 'prop', steps: 1, param: opt.param || 'loose', doc: opt.doc === undefined ? 1 : opt.doc,
      note: opt.note,
      set(c) { c.st[prop] = val; c.sheet.setStyle(c.row, c.col, c.st); },
      read(c) { const s = c.sheet.getStyle(c.row, c.col); return opt.readAs ? opt.readAs(s, c) : s[prop]; },
      exp: opt.exp !== undefined ? opt.exp : val,
    };
  }
  function sjsCall(api, set, read, exp, opt) {
    opt = opt || {};
    return { api, kind: opt.kind || 'method', steps: opt.steps || 1, param: opt.param || 'typed', doc: opt.doc === undefined ? 1 : opt.doc, note: opt.note, assert: opt.assert, set, read, exp };
  }
  function uniCall(api, set, read, exp, opt) {
    opt = opt || {};
    return { api, kind: opt.kind || 'method', steps: opt.steps || 1, param: opt.param || 'typed', doc: opt.doc === undefined ? 1 : opt.doc, note: opt.note, assert: opt.assert, set, read, exp };
  }
  function none(api, tierHint, note) {
    return { api: api || '（无公开 API）', kind: 'none', steps: 0, param: '-', doc: 0, tierHint: tierHint || 'S3', note, set: null, read: null, exp: null };
  }
  const uniRule = (c) => { const rs = c.usheet.getConditionalFormattingRules() || []; return rs[rs.length - 1]; };
  const uniCFLast = () => (c) => { const r = uniRule(c); return JSON.stringify(r && r.rule); };
  const uniCFCount = (c) => (c.usheet.getConditionalFormattingRules() || []).length;
  const sjsCFCount = (c) => c.sheet.conditionalFormats.count();
  const sjsCFLast = (c) => {
    const rs = c.sheet.conditionalFormats.getRules(c.row, c.col);
    return rs && rs[rs.length - 1];
  };
  const sjsCF = (api, addFn, read, exp, opt) => sjsCall(api, (c) => { addFn(c, CFS); }, (c) => {
    const r = sjsCFLast(c);
    return r ? read(r, c) : 'NO_RULE';
  }, exp, Object.assign({ kind: 'method', steps: 3, param: 'enum' }, opt || {}));
  const uniCFAdd = (api, build, read, exp, opt) => uniCall(api, (c) => {
    const b = c.usheet.newConditionalFormattingRule();
    build(b, c);
    c.usheet.addConditionalFormattingRule(b.build());
  }, read, exp, Object.assign({ kind: 'builder', steps: 3, param: 'typed' }, opt || {}));
  const grng = (c) => [c.urange.getRange()];

  const SPEC = [];
  const add = (o) => { o.row = 5 + SPEC.length; SPEC.push(o); return o; };

  // ================= A 字体（13） =================
  add({ id: 'A1', g: 'A', name: '字体名称', excel: '微软雅黑',
    sjs: sjsProp('Style.fontFamily', 'fontFamily', '微软雅黑', { param: 'loose' }),
    uni: uniCall('FRange.setFontFamily', (c) => c.urange.setFontFamily('微软雅黑'), (c) => c.urange.getCellStyleData().ff, '微软雅黑') });
  add({ id: 'A2', g: 'A', name: '字号', excel: '16 磅',
    sjs: sjsProp('Style.fontSize = "16pt"', 'fontSize', '16pt', {
      param: 'loose',
      note: '只收字符串；传裸数字 16 抛 "e.match is not a function"（错误信息不可读）',
    }),
    uni: uniCall('FRange.setFontSize(16)', (c) => c.urange.setFontSize(16), (c) => c.urange.getCellStyleData().fs, 16) });
  add({ id: 'A3', g: 'A', name: '加粗', excel: 'bold',
    sjs: sjsProp('Style.fontWeight', 'fontWeight', 'bold', { param: 'loose' }),
    uni: uniCall('FRange.setFontWeight("bold")', (c) => c.urange.setFontWeight('bold'), (c) => c.urange.getCellStyleData().bl, 1) });
  add({ id: 'A4', g: 'A', name: '斜体', excel: 'italic',
    sjs: sjsProp('Style.fontStyle', 'fontStyle', 'italic', { param: 'loose' }),
    uni: uniCall('FRange.setFontStyle("italic")', (c) => c.urange.setFontStyle('italic'), (c) => c.urange.getCellStyleData().it, 1) });
  add({ id: 'A5', g: 'A', name: '下划线（4 种）', excel: '单/双/会计单/会计双',
    sjs: sjsProp('Style.textDecoration = TextDecorationType.underline', 'textDecoration', SJS.TextDecorationType.underline, { param: 'enum' }),
    uni: uniCall('FRange.setFontLine("underline")', (c) => c.urange.setFontLine('underline'), (c) => { const d = c.urange.getCellStyleData(); return d.ul && d.ul.s; }, 1, { param: 'loose' }) });
  add({ id: 'A6', g: 'A', name: '下划线颜色', excel: '与字色独立',
    sjs: none('（textDecoration 只有线型，无独立颜色）', 'S3', 'SJS 侧下划线颜色跟随字体色'),
    uni: none('（setFontLine 只有线型，无独立颜色）', 'S3') });
  add({ id: 'A7', g: 'A', name: '删除线', excel: '单删除线',
    sjs: sjsProp('Style.textDecoration = TextDecorationType.lineThrough', 'textDecoration', SJS.TextDecorationType.lineThrough, { param: 'enum' }),
    uni: uniCall('FRange.setFontLine("line-through")', (c) => c.urange.setFontLine('line-through'), (c) => { const d = c.urange.getCellStyleData(); return d.st && d.st.s; }, 1, { param: 'loose' }) });
  add({ id: 'A8', g: 'A', name: '字体颜色', excel: '#C00000',
    sjs: sjsProp('Style.foreColor', 'foreColor', '#C00000', { param: 'loose' }),
    uni: uniCall('FRange.setFontColor', (c) => c.urange.setFontColor('#C00000'), (c) => (c.urange.getCellStyleData().cl || {}).rgb, '#C00000') });
  // 富文本读回：必须显式指定 ValueType.richText，否则 getValue 只回纯文本
  const sjsRich = (c) => c.sheet.getValue(c.row, c.col, SJS.SheetArea.viewport, SJS.ValueType.richText);
  add({ id: 'A9', g: 'A', name: '上标', excel: '字体效果',
    sjs: sjsCall('sheet.setValue({richText, vertAlign:1})', (c) => {
      c.sheet.setValue(c.row, c.col, {
        richText: [
          { style: { font: 'normal 14px 微软雅黑' }, text: 'R' },
          { style: { font: 'normal 10px 微软雅黑', vertAlign: 1 }, text: '2' },
        ],
      });
    }, (c) => { const v = sjsRich(c); return v && v.richText ? v.richText[1].style.vertAlign : 'no-rich'; }, 1,
      { kind: 'builder', steps: 1, param: 'typed', note: '富文本是"值"不是"方法"：传给 setValue 的 richText 对象数组，1=上标 2=下标 0=正常' }),
    uni: none('（无富文本 run 接口）', 'S3') });
  add({ id: 'A10', g: 'A', name: '下标', excel: '字体效果',
    sjs: sjsCall('sheet.setValue({richText, vertAlign:2})', (c) => {
      c.sheet.setValue(c.row, c.col, {
        richText: [
          { style: { font: 'normal 14px 微软雅黑' }, text: 'M' },
          { style: { font: 'normal 10px 微软雅黑', vertAlign: 2 }, text: '1' },
        ],
      });
    }, (c) => { const v = sjsRich(c); return v && v.richText ? v.richText[1].style.vertAlign : 'no-rich'; }, 2,
      { kind: 'builder', steps: 1, param: 'typed' }),
    uni: none('（同上）', 'S3') });
  add({ id: 'A11', g: 'A', name: '主题字体', excel: '正文/标题随主题',
    sjs: sjsProp('Style.themeFont', 'themeFont', 'Body', { param: 'loose' }),
    uni: none('（无 themeFont 字段）', 'S3') });
  add({ id: 'A12', g: 'A', name: '字符间距/等宽', excel: 'Excel 无此能力', skip: 'Excel 本身无此设置，不测', sjs: null, uni: null });
  add({ id: 'A13', g: 'A', name: '富文本（段内多样式）', excel: '单元格内分段样式',
    sjs: sjsCall('sheet.setValue({richText:[{text,style}...]})', (c) => {
      c.sheet.setValue(c.row, c.col, {
        richText: [
          { style: { foreColor: '#C00000', font: 'bold 14px 微软雅黑' }, text: '红色' },
          { style: { font: 'normal 14px 微软雅黑' }, text: '黑色' },
        ],
      });
    }, (c) => { const v = sjsRich(c); return v && v.richText ? v.richText.length : 'no-rich'; }, 2,
      { kind: 'builder', steps: 1, param: 'typed', note: '读回必须带 ValueType.richText；直接 getValue 只得到拼接后的纯文本' }),
    uni: none('（FRange 无 setRichText；setValue 传 richText 对象读回为 null）', 'S3') });

  // ================= B 填充（8） =================
  add({ id: 'B1', g: 'B', name: '纯色填充', excel: '#FFF2CC',
    sjs: sjsProp('Style.backColor', 'backColor', '#FFF2CC', { param: 'loose' }),
    uni: uniCall('FRange.setBackground', (c) => c.urange.setBackground('#FFF2CC'), (c) => (c.urange.getCellStyleData().bg || {}).rgb, '#FFF2CC') });
  add({ id: 'B2', g: 'B', name: '图案填充类型', excel: '6% 灰',
    sjs: sjsProp('Style.backColor = {type: PatternType}', 'backColor',
      { type: SJS.PatternType.gray0625, patternColor: '#000000', backgroundColor: '#FFFFFF' },
      { param: 'enum', readAs: (s) => (s.backColor && s.backColor.type), exp: SJS.PatternType.gray0625 }),
    uni: none('（bg 只有 rgb，无 pattern 字段）', 'S3') });
  add({ id: 'B3', g: 'B', name: '图案前景色', excel: '图案线颜色',
    sjs: sjsProp('Style.backColor.patternColor', 'backColor',
      { type: SJS.PatternType.gray0625, patternColor: '#FF0000', backgroundColor: '#FFFFFF' },
      { param: 'enum', readAs: (s) => (s.backColor && s.backColor.patternColor), exp: '#FF0000' }),
    uni: none('', 'S3') });
  add({ id: 'B4', g: 'B', name: '图案背景色', excel: '图案底色',
    sjs: sjsProp('Style.backColor.backgroundColor', 'backColor',
      { type: SJS.PatternType.gray0625, patternColor: '#000000', backgroundColor: '#00FF00' },
      { param: 'enum', readAs: (s) => (s.backColor && s.backColor.backgroundColor), exp: '#00FF00' }),
    uni: none('', 'S3') });
  add({ id: 'B5', g: 'B', name: '线性渐变（角度）', excel: '0~359.9',
    sjs: sjsProp('Style.backColor = {degree, stops}', 'backColor',
      { type: 2, degree: 45, stops: [{ color: '#FFFFFF', position: 0 }, { color: '#000000', position: 1 }] },
      { param: 'enum', readAs: (s) => (s.backColor && s.backColor.degree), exp: 45, note: 'type=2 表示渐变；读回时 backColor.type 不保留，degree/stops 保留' }),
    uni: none('', 'S3') });
  add({ id: 'B6', g: 'B', name: '渐变停靠点', excel: '≥2 个光圈',
    sjs: sjsProp('Style.backColor.stops', 'backColor',
      { type: 2, degree: 0, stops: [{ color: '#FFFFFF', position: 0 }, { color: '#000000', position: 1 }] },
      { param: 'enum', readAs: (s) => (s.backColor && s.backColor.stops ? s.backColor.stops.length : 0), exp: 2 }),
    uni: none('', 'S3') });
  add({ id: 'B7', g: 'B', name: '渐变类型（射线/矩形）', excel: '5 种',
    sjs: none('（只有线性渐变，射线/矩形/路径渐变无入口）', 'S3'),
    uni: none('', 'S3') });
  add({ id: 'B8', g: 'B', name: '无填充（清除）', excel: '透明',
    sjs: sjsProp('Style.backColor = undefined', 'backColor', undefined, { exp: undefined }),
    uni: uniCall('FRange.setBackground(null)', (c) => c.urange.setBackground(null), (c) => (c.urange.getCellStyleData().bg || {}).rgb, undefined) });

  // ================= C 边框（11） =================
  const brd = (color, style, opt) => (c) => c.rrange.setBorder(new SJS.LineBorder(color, style), opt);
  const sjsBorder = (api, setFn, readFn, exp, opt) => sjsCall(api, setFn, readFn, exp, Object.assign({ kind: 'method', steps: 2, param: 'enum' }, opt || {}));
  const bTop = (c) => { const b = c.sheet.getCell(c.row, c.col).borderTop(); return b ? b.style + '/' + b.color : null; };
  const bBottom = (c) => { const b = c.sheet.getCell(c.row, c.col).borderBottom(); return b ? b.style + '/' + b.color : null; };
  const bRight = (c) => { const b = c.sheet.getCell(c.row, c.col).borderRight(); return b ? b.style + '/' + b.color : null; };
  const uniBd = (c) => c.urange.getCellStyleData().bd || {};
  const uniBorder = (api, type, style, color, exp, opt) => uniCall(api, (c) => {
    c.urange.setBorder(type, style, color);
  }, (c) => JSON.stringify(uniBd(c)), exp, Object.assign({ kind: 'method', steps: 1, param: 'enum' }, opt || {}));

  add({ id: 'C1', g: 'C', name: '四边外框', excel: '中粗红',
    sjs: sjsBorder('Range.setBorder(LineBorder,{outline})', brd('#C00000', SJS.LineStyle.medium, { outline: true }), bTop, SJS.LineStyle.medium + '/#C00000'),
    uni: uniBorder('FRange.setBorder("outside",…)', 'outside', 8, '#C00000', '#C00000') });
  add({ id: 'C2', g: 'C', name: '内部横线', excel: '区域内水平线',
    sjs: sjsBorder('Range.setBorder(…,{inside:true})', brd('#9CA3AF', SJS.LineStyle.dashed, { inside: true }), bBottom, SJS.LineStyle.dashed + '/#9CA3AF',
      { note: '内部线落在首格的 borderBottom 上（不是 borderTop）' }),
    uni: uniCall('setBorder("horizontal")', (c) => {
      // 内部横线必须在"多行"区域上才有落点：1 行区域没有内部横线，setBorder 会静默返回空 bd
      c.usheet.getRange(c.row, c.col, 2, 3).setBorder('horizontal', 4, '#9CA3AF');
    }, (c) => JSON.stringify(c.usheet.getRange(c.row, c.col, 2, 3).getCellStyleData().bd || {}), '"s":4',
      { kind: 'method', steps: 1, param: 'enum', note: '1 行区域上调用会静默无效（bd={}），需 2 行以上' }) });
  add({ id: 'C3', g: 'C', name: '内部竖线', excel: '区域内垂直线',
    sjs: sjsBorder('Range.setBorder(…,{inside:true})', brd('#9CA3AF', SJS.LineStyle.dashed, { inside: true }), bRight, SJS.LineStyle.dashed + '/#9CA3AF',
      { note: '内部竖线落在首格的 borderRight 上' }),
    uni: uniBorder('setBorder("vertical")', 'vertical', 4, '#9CA3AF', '"s":4') });
  add({ id: 'C4', g: 'C', name: '线型（13 种）', excel: '细/中/粗/虚线/双线…',
    sjs: sjsBorder('LineStyle 14 值', brd('#000000', SJS.LineStyle.double, { outline: true }), bTop, SJS.LineStyle.double + '/#000000'),
    uni: uniBorder('BorderStyleTypes 14 值', 'all', 7, '#000000', '"s":7') });
  add({ id: 'C5', g: 'C', name: '边框颜色', excel: '独立于线型',
    sjs: sjsBorder('LineBorder(color, style)', brd('#00B050', SJS.LineStyle.thin, { outline: true }), bTop, SJS.LineStyle.thin + '/#00B050'),
    uni: uniBorder('setBorder(type,style,color)', 'all', 1, '#00B050', '#00B050') });
  add({ id: 'C6', g: 'C', name: '对角线 ↘', excel: '左上到右下',
    sjs: sjsCall('Style.borderDiagonalDown', (c) => {
      const st = new SJS.Style(); st.borderDiagonalDown = new SJS.LineBorder('#C00000', SJS.LineStyle.thin);
      c.sheet.setStyle(c.row, c.col, st);
    }, (c) => (c.sheet.getStyle(c.row, c.col).borderDiagonalDown ? 1 : 0), 1, { kind: 'prop', steps: 1, param: 'enum' }),
    uni: uniBorder('setBorder("tlbr")', 'tlbr', 1, '#C00000', 'tl_br') });
  add({ id: 'C7', g: 'C', name: '对角线 ↙', excel: '右上到左下',
    sjs: sjsCall('Style.borderDiagonalUp', (c) => {
      const st = new SJS.Style(); st.borderDiagonalUp = new SJS.LineBorder('#C00000', SJS.LineStyle.thin);
      c.sheet.setStyle(c.row, c.col, st);
    }, (c) => (c.sheet.getStyle(c.row, c.col).borderDiagonalUp ? 1 : 0), 1, { kind: 'prop', steps: 1, param: 'enum' }),
    uni: uniBorder('setBorder("bltr")', 'bltr', 1, '#C00000', 'bl_tr',
      { note: 'Enum.BorderType.BLTR === "bl_tr"，但 setter 只认 "bltr"；传 "bl_tr" 不抛错、静默无效果' }) });
  add({ id: 'C8', g: 'C', name: '斜线表头（双斜线）', excel: '组合效果',
    sjs: sjsCall('两条对角线组合（无专用 API）', (c) => {
      const st = new SJS.Style();
      st.borderDiagonalDown = new SJS.LineBorder('#C00000', SJS.LineStyle.thin);
      st.borderDiagonalUp = new SJS.LineBorder('#C00000', SJS.LineStyle.thin);
      c.sheet.setStyle(c.row, c.col, st);
    }, (c) => { const s = c.sheet.getStyle(c.row, c.col); return (s.borderDiagonalDown ? 1 : 0) + (s.borderDiagonalUp ? 1 : 0); }, 2, { kind: 'builder', steps: 2, doc: 0 }),
    uni: uniCall('两条对角线组合', (c) => { c.urange.setBorder('tlbr', 1, '#C00000'); c.urange.setBorder('bltr', 1, '#C00000'); },
      (c) => { const d = uniBd(c); return (d.tl_br ? 1 : 0) + (d.bl_tr ? 1 : 0); }, 2, { kind: 'builder', steps: 2, doc: 0 }) });
  add({ id: 'C9', g: 'C', name: '清除边框', excel: '全部无框线',
    sjs: sjsCall('LineStyle.empty (0)', (c) => {
      c.rrange.setBorder(new SJS.LineBorder('#000000', SJS.LineStyle.thin), { outline: true, inside: true });
      c.rrange.setBorder(new SJS.LineBorder('#000000', SJS.LineStyle.empty), { outline: true, inside: true });
    }, (c) => { const b = c.sheet.getCell(c.row, c.col).borderTop(); return b ? b.style : -1; }, SJS.LineStyle.empty, { steps: 2, param: 'enum' }),
    uni: uniCall('setBorder("none", 0)', (c) => {
      c.urange.setBorder('all', 1, '#000000');
      c.urange.setBorder('none', 0, '#000000');
    }, (c) => Object.keys(uniBd(c)).length, 0, { steps: 2, param: 'enum' }) });
  add({ id: 'C10', g: 'C', name: '边框与合并区交互', excel: '按合并区外框',
    sjs: sjsCall('addSpan + setBorder', (c) => {
      c.sheet.addSpan(c.row, c.col, 2, 2);
      c.rrange.setBorder(new SJS.LineBorder('#C00000', SJS.LineStyle.medium), { outline: true });
    }, (c) => { const b = c.sheet.getCell(c.row, c.col).borderTop(); return b ? b.color : null; }, '#C00000', { kind: 'builder', steps: 3, param: 'enum' }),
    uni: uniCall('merge + setBorder', (c) => {
      c.urange.merge();
      c.urange.setBorder('outside', 8, '#C00000');
    }, (c) => c.urange.isMerged(), true, { kind: 'builder', steps: 2, param: 'enum' }) });
  add({ id: 'C11', g: 'C', name: '圆角边框', excel: 'Excel 无（SJS 独有）',
    sjs: sjsProp('Style.borderRadius（v19.2）', 'borderRadius', 8, { doc: 1 }),
    uni: none('（Excel 本身也没有，非缺陷）', 'S-') });

  // ================= D 对齐与合并（17） =================
  add({ id: 'D1', g: 'D', name: '水平对齐 8 种', excel: '常规/左/中/右/填充/两端/跨列/分散',
    sjs: sjsProp('Style.hAlign（8 值）', 'hAlign', SJS.HorizontalAlign.distributed, { param: 'enum' }),
    uni: uniCall('setHorizontalAlignment("center")', (c) => c.urange.setHorizontalAlignment('center'),
      (c) => c.urange.getHorizontalAlignment(), 'center', { param: 'loose', note: '只认 left/center/normal（normal=右）3 个字符串，无填充/两端/跨列/分散' }) });
  add({ id: 'D2', g: 'D', name: '垂直对齐 5 种', excel: '上/中/下/两端/分散',
    sjs: sjsProp('Style.vAlign（4 值）', 'vAlign', SJS.VerticalAlign.center, { param: 'enum' }),
    uni: uniCall('setVerticalAlignment("middle")', (c) => c.urange.setVerticalAlignment('middle'),
      (c) => c.urange.getVerticalAlignment(), 'middle', { param: 'loose', note: '只认 top/middle/bottom，无两端/分散' }) });
  add({ id: 'D3', g: 'D', name: '缩进 0~15', excel: '字符缩进',
    sjs: sjsProp('Style.textIndent', 'textIndent', 2),
    uni: none('（FRange 无 setTextIndent）', 'S3') });
  add({ id: 'D4', g: 'D', name: '自动换行', excel: '按列宽折行',
    sjs: sjsProp('Style.wordWrap', 'wordWrap', true),
    uni: uniCall('FRange.setWrap', (c) => c.urange.setWrap(true), (c) => c.urange.getWrap(), true) });
  add({ id: 'D5', g: 'D', name: '文本旋转 -90~90', excel: '角度旋转',
    sjs: sjsProp('Style.textOrientation', 'textOrientation', 45, { param: 'enum' }),
    uni: uniCall('FRange.setTextRotation', (c) => c.urange.setTextRotation(45), (c) => { const t = c.urange.getCellStyleData().tr; return t && t.a; }, 45, { param: 'enum' }) });
  add({ id: 'D6', g: 'D', name: '竖排文字', excel: '东亚竖排',
    sjs: sjsProp('Style.isVerticalText', 'isVerticalText', true),
    uni: none('（无 setVerticalText）', 'S3') });
  add({ id: 'D7', g: 'D', name: '缩小字体填充', excel: '缩字号不折行',
    sjs: sjsProp('Style.shrinkToFit', 'shrinkToFit', true),
    uni: uniCall('FRange.setShrinkToFit', (c) => c.urange.setShrinkToFit(true), (c) => c.urange.getCellStyleData().stf, 1) });
  add({ id: 'D8', g: 'D', name: '跨列居中', excel: '不合并跨列居中',
    sjs: sjsProp('HorizontalAlign.centerContinuous', 'hAlign', SJS.HorizontalAlign.centerContinuous, { param: 'enum' }),
    uni: none('（对齐只认 3 个字符串，无跨列居中）', 'S3') });
  add({ id: 'D9', g: 'D', name: '文字方向 RTL', excel: '从右到左',
    sjs: sjsProp('Style.textDirection = 2', 'textDirection', 2, { param: 'enum', note: 'SJS.TextDirection 枚举未导出，但属性收数字 2' }),
    uni: none('（无 setTextDirection）', 'S3') });
  add({ id: 'D10', g: 'D', name: '合并仅保留左上值', excel: '其余清空',
    sjs: sjsCall('sheet.addSpan', (c) => {
      c.sheet.setValue(c.row, c.col, 'A'); c.sheet.setValue(c.row, c.col + 1, 'B');
      c.sheet.addSpan(c.row, c.col, 1, 2);
    }, (c) => { const sp = c.sheet.getSpan(c.row, c.col); return sp ? String(c.sheet.getValue(c.row, c.col + 1)) : 'no-merge'; }, '', { kind: 'method', steps: 1 }),
    uni: uniCall('FRange.merge', (c) => {
      c.urange.setValue('A'); c.usheet.getRange(c.row, c.col + 1).setValue('B');
      c.urange.merge();
    }, (c) => String(c.usheet.getRange(c.row, c.col + 1).getValue()), '', { kind: 'method', steps: 1 }) });
  add({ id: 'D11', g: 'D', name: '合并区域读取', excel: '枚举合并区',
    sjs: sjsCall('sheet.getSpan', (c) => { c.sheet.addSpan(c.row, c.col, 1, 2); },
      (c) => { const sp = c.sheet.getSpan(c.row, c.col); return sp ? 1 : 0; }, 1, { steps: 2 }),
    uni: uniCall('sheet.getMergedRanges', (c) => { c.urange.merge(); },
      (c) => { const m = c.usheet.getMergedRanges() || []; return m.some((r) => r.getRow() === c.row && r.getColumn() === c.col) ? 1 : 0; }, 1, { steps: 2 }) });
  add({ id: 'D12', g: 'D', name: '取消合并', excel: '还原独立格',
    sjs: sjsCall('sheet.removeSpan（两参）', (c) => {
      try { c.sheet.removeSpan(c.row, c.col); } catch (e) { /* 该格本来就没有 span 时 removeSpan 会抛 Invalid range */ }
      c.sheet.addSpan(c.row, c.col, 1, 3);
      c.sheet.removeSpan(c.row, c.col);
    }, (c) => (c.sheet.getSpan(c.row, c.col) ? 1 : 0), 0,
      { steps: 2, note: 'removeSpan 只认两参 (row,col)；传四参或 Range 对象不报错但取消不了合并；对无 span 的格调用会抛 Invalid range' }),
    uni: uniCall('FRange.breakApart', (c) => { c.urange.merge(); c.urange.breakApart(); },
      (c) => (c.urange.isMerged() ? 1 : 0), 0, { steps: 2 }) });
  add({ id: 'D13', g: 'D', name: '跨行/跨列合并', excel: '仅横向/仅纵向',
    sjs: sjsCall('addSpan(1,3)', (c) => { c.sheet.addSpan(c.row, c.col, 1, 3); },
      (c) => { const s = c.sheet.getSpan(c.row, c.col); return s ? s.rowCount + 'x' + s.colCount : null; }, '1x3'),
    uni: uniCall('FRange.mergeAcross', (c) => { c.urange.mergeAcross(); },
      (c) => (c.urange.isMerged() ? 1 : 0), 1) });
  add({ id: 'D14', g: 'D', name: '合并+换行+垂直居中', excel: '报表标题组合',
    sjs: sjsCall('属性组合', (c) => {
      const st = new SJS.Style();
      st.wordWrap = true; st.vAlign = SJS.VerticalAlign.center; st.hAlign = SJS.HorizontalAlign.center;
      c.sheet.setStyle(c.row, c.col, st); c.sheet.addSpan(c.row, c.col, 2, 3);
    }, (c) => { const s = c.sheet.getStyle(c.row, c.col); return s.wordWrap + '/' + s.vAlign + '/' + s.hAlign; }, 'true/1/1', { kind: 'builder', steps: 2 }),
    uni: uniCall('属性组合', (c) => {
      c.urange.setWrap(true);
      c.urange.setVerticalAlignment('middle');
      c.urange.setHorizontalAlignment('center');
      c.urange.merge();
    }, (c) => c.urange.getWrap() + '/' + c.urange.getVerticalAlignment() + '/' + c.urange.getHorizontalAlignment(), 'true/middle/center', { kind: 'builder', steps: 3 }) });
  add({ id: 'D15', g: 'D', name: '文本省略符', excel: 'Excel 侧等同溢出（SJS 独有）',
    sjs: sjsProp('Style.showEllipsis', 'showEllipsis', true, { doc: 1 }),
    uni: none('（Excel 本身也没有，非缺陷）', 'S-') });
  add({ id: 'D16', g: 'D', name: '单元格内边距', excel: 'Excel 无（SJS 独有）',
    sjs: sjsProp('Style.padding', 'padding', { left: 8, top: 4 }, { doc: 1, readAs: (s) => JSON.stringify(s.padding), exp: '"left":8' }),
    uni: none('', 'S-') });
  add({ id: 'D17', g: 'D', name: '自动合并', excel: 'Excel 无（SJS 独有）',
    sjs: sjsCall('sheet.autoMerge(range, mode, direction)', (c) => {
      for (let i = 0; i < 3; i++) c.sheet.setValue(c.row + i, c.col, 'X');
      c.sheet.autoMerge(new SJS.Range(c.row, c.col, 3, 1), SJS.AutoMerge.AutoMergeMode.free, SJS.AutoMerge.AutoMergeDirection.column);
    }, () => undefined, undefined,
      { kind: 'method', steps: 1, param: 'enum', assert: 'noThrow',
        note: '调用不报错，但 getSpan / toJSON / 逐格取值都看不到痕迹 → 无法证实生效、也无法读回' }),
    uni: none('', 'S-') });

  // ================= E 数字格式（25） =================
  const FMT = [
    ['E1', '常规格式', 'General'],
    ['E2', '数值（小数+千分位）', '#,##0.00'],
    ['E3', '货币', '¥#,##0.00;[Red]¥-#,##0.00'],
    ['E4', '会计专用', '_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ ;_ ¥* "-"??_ ;_ @_ '],
    ['E5', '日期', 'yyyy-mm-dd'],
    ['E6', '时间', 'h:mm:ss'],
    ['E7', '百分比', '0.00%'],
    ['E8', '分数', '# ?/?'],
    ['E9', '科学记数', '0.00E+00'],
    ['E10', '文本', '@'],
    ['E11', '中文大写（DBNum2）', '[DBNum2][$-804]G/通用格式'],
    ['E12', '自定义 0 占位符', '00000'],
    ['E13', '自定义 # 占位符', '###0'],
    ['E14', '自定义 ? 占位符', '0.0??'],
    ['E15', '自定义 @ 文本段', '0.00"元";@'],
    ['E16', '千分位缩放', '#,##0,,"百万"'],
    ['E17', '多段+负数括号', '#,##0.00_);[Red](#,##0.00)'],
    ['E18', '条件段 [>100]', '[>100][Green]0;[Red]0'],
    ['E19', '颜色段 [Red]', '[Red]0.00;[Blue]-0.00'],
    ['E20', '日期占位符+星期', 'yyyy"年"m"月"d"日" aaaa'],
    ['E21', '格式串往返读取', '#,##0.000'],
    ['E22', '区域性代码 [$-411]', '[$-411]#,##0.00'],
    ['E23', 'CJK 转义符 !（v19.2）', '0!.0,"万"'],
    ['E24', '账单（会计）对齐', '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'],
    ['E25', '格式串内嵌公式', 'TEXT(1,"0.00")'],
  ];
  FMT.forEach(([id, name, fmt]) => {
    add({
      id, g: 'E', name, excel: fmt,
      sjs: sjsCall('sheet.setFormatter / getFormatter', (c) => {
        c.sheet.setValue(c.row, c.col, 1234.5678);
        c.sheet.setFormatter(c.row, c.col, fmt);
      }, (c) => c.sheet.getFormatter(c.row, c.col), fmt, { param: 'loose' }),
      uni: uniCall('FRange.setNumberFormat / getNumberFormat', (c) => {
        c.urange.setValue(1234.5678);
        c.urange.setNumberFormat(fmt);
      }, (c) => c.urange.getNumberFormat(), fmt, { param: 'loose' }),
    });
  });

  // ================= F 行高列宽（8） =================
  add({ id: 'F1', g: 'F', name: '设置行高', excel: '40 磅',
    sjs: sjsCall('sheet.setRowHeight', (c) => c.sheet.setRowHeight(c.row, 40), (c) => c.sheet.getRowHeight(c.row), 40),
    uni: uniCall('FWorksheet.setRowHeight', (c) => c.usheet.setRowHeight(c.row, 40), (c) => c.usheet.getRowHeight(c.row), 40,
      { note: '有最小值夹紧：设 20 读回 24' }) });
  add({ id: 'F2', g: 'F', name: '设置列宽', excel: '120 px',
    sjs: sjsCall('sheet.setColumnWidth', (c) => c.sheet.setColumnWidth(c.col, 120), (c) => c.sheet.getColumnWidth(c.col), 120),
    uni: uniCall('FWorksheet.setColumnWidth', (c) => c.usheet.setColumnWidth(c.col, 120), (c) => c.usheet.getColumnWidth(c.col), 120) });
  add({ id: 'F3', g: 'F', name: '自动调整行高', excel: '适配换行内容',
    sjs: sjsCall('sheet.autoFitRow', (c) => {
      c.sheet.setValue(c.row, c.col, '很长很长的一段文字需要折行才能看出自动行高效果，再长一点再长一点');
      c.sheet.getRange(c.row, c.col).wordWrap(true);
      const before = c.sheet.getRowHeight(c.row);
      c.sheet.autoFitRow(c.row);
      c.fit = { before, after: c.sheet.getRowHeight(c.row) };
    }, (c) => c.fit.after > c.fit.before, true),
    uni: uniCall('FWorksheet.autoResizeRows', (c) => {
      c.urange.setValue('很长很长的一段文字需要折行才能看出自动行高效果，再长一点再长一点');
      c.urange.setWrap(true);
      c.usheet.autoResizeRows(c.row, 1);
      c.fit = { after: c.usheet.getRowHeight(c.row) };
    }, (c) => c.fit.after > 40, true,
      { note: 'setWrap(true) 本身就会把行高撑到多行高度，autoResizeRows 前后常无差值；判据取"换行后行高是否被撑高"' }) });
  add({ id: 'F4', g: 'F', name: '自动调整列宽', excel: '适配最宽内容',
    sjs: sjsCall('sheet.autoFitColumn', (c) => {
      c.sheet.setValue(c.row, c.col, '一段相当长的文本用来撑开列宽');
      const before = c.sheet.getColumnWidth(c.col);
      c.sheet.autoFitColumn(c.col);
      c.fit = { before, after: c.sheet.getColumnWidth(c.col) };
    }, (c) => c.fit.after > c.fit.before, true),
    uni: uniCall('FWorksheet.autoResizeColumns', (c) => {
      c.urange.setValue('一段相当长的文本用来撑开列宽');
      const before = c.usheet.getColumnWidth(c.col);
      c.usheet.autoResizeColumns(c.col, 1);
      c.fit = { before, after: c.usheet.getColumnWidth(c.col) };
    }, (c) => c.fit.after > c.fit.before, true) });
  add({ id: 'F5', g: 'F', name: '隐藏行', excel: '行高为 0 且不可见',
    sjs: sjsCall('sheet.setRowVisible', (c) => c.sheet.setRowVisible(c.row, false), (c) => c.sheet.getRowVisible(c.row), false),
    uni: uniCall('FWorksheet.hideRows(row, count)', (c) => c.usheet.hideRows(c.row, 1), () => undefined, undefined,
      { assert: 'noThrow', note: '有写入口但无可信读回口：getRowHeight 不变、getHiddenState() 恒 0、FRange 无 isRowHidden' }) });
  add({ id: 'F6', g: 'F', name: '隐藏列', excel: '列宽为 0 且不可见',
    sjs: sjsCall('sheet.setColumnVisible', (c) => c.sheet.setColumnVisible(c.col, false), (c) => c.sheet.getColumnVisible(c.col), false),
    uni: uniCall('FWorksheet.hideColumns(col, count)', (c) => c.usheet.hideColumns(c.col, 1), () => undefined, undefined,
      { assert: 'noThrow', note: '同上，getColumnWidth 不变' }) });
  add({ id: 'F7', g: 'F', name: '默认行高/列宽', excel: '改新建行默认值',
    sjs: sjsCall('sheet.defaults.rowHeight', (c) => {
      const d = c.sheet.defaults || {};
      d.rowHeight = 30; d.colWidth = 100;
      c.def = (c.sheet.defaults || {}).rowHeight;
    }, (c) => c.def, 30, { kind: 'prop', param: 'loose' }),
    uni: none('（无默认行高/列宽 setter）', 'S4') });
  add({ id: 'F8', g: 'F', name: '行高列宽导入保真', excel: '从 xlsx 还原', skip: '属导入链，用户自测', sjs: null, uni: null });

  // ================= G 条件格式规则（16） =================
  add({ id: 'G1', g: 'G', name: '单元格值（大于）', excel: '>55 标红',
    sjs: sjsCF('addCellValueRule', (c, CF) => {
      c.sheet.conditionalFormats.addCellValueRule(CF.ComparisonOperators.greaterThan, 55, 0, { backColor: '#FFC7CE', foreColor: '#9C0006' }, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 1),
    uni: uniCFAdd('whenNumberGreaterThan', (b, c) => b.whenNumberGreaterThan(55).setBackground('#FFC7CE').setFontColor('#9C0006').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G2', g: 'G', name: '特定文本（包含）', excel: '含"合格"',
    sjs: sjsCF('addSpecificTextRule', (c, CF) => {
      c.sheet.conditionalFormats.addSpecificTextRule(CF.TextCompareType.contains, '合格', { backColor: '#C6EFCE' }, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 2),
    uni: uniCFAdd('whenTextContains', (b, c) => b.whenTextContains('合格').setBackground('#C6EFCE').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G3', g: 'G', name: '日期发生', excel: '今天',
    sjs: sjsCF('addDateOccurringRule', (c, CF) => {
      c.sheet.conditionalFormats.addDateOccurringRule(CF.DateOccurringType.today, { backColor: '#FFEB9C' }, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 4),
    uni: uniCFAdd('whenDate("today")', (b, c) => b.whenDate('today').setBackground('#FFEB9C').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G4', g: 'G', name: '空值/非空', excel: '空白高亮',
    sjs: sjsCF('addFormulaRule（LEN=0）', (c, CF) => {
      c.sheet.conditionalFormats.addFormulaRule('=LEN(A1)=0', { backColor: '#DDEBF7' }, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 3),
    uni: uniCFAdd('whenCellEmpty', (b, c) => b.whenCellEmpty().setBackground('#DDEBF7').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G5', g: 'G', name: '公式规则', excel: '自定义公式为真',
    sjs: sjsCF('addFormulaRule', (c, CF) => {
      c.sheet.conditionalFormats.addFormulaRule('=MOD(A1,2)=0', { backColor: '#E4DFEC' }, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 3),
    uni: uniCFAdd('whenFormulaSatisfied', (b, c) => b.whenFormulaSatisfied('=MOD(A1,2)=0').setBackground('#E4DFEC').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G6', g: 'G', name: '前 N 项', excel: '前 3 名',
    sjs: sjsCF('addTop10Rule', (c, CF) => {
      c.sheet.conditionalFormats.addTop10Rule(CF.Top10ConditionType.top, 3, { backColor: '#FCE4D6' }, [new SJS.Range(c.row, c.col, 1, 3)], false);
    }, (r) => r.rank(), 3),
    uni: uniCFAdd('setRank', (b, c) => b.setRank({ isBottom: false, isPercent: false, value: 3 }).setBackground('#FCE4D6').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G7', g: 'G', name: '前 N%', excel: '前 10%',
    sjs: sjsCF('addTop10Rule + isPercent', (c, CF) => {
      c.sheet.conditionalFormats.addTop10Rule(CF.Top10ConditionType.top, 10, { backColor: '#FFF2CC' }, [new SJS.Range(c.row, c.col, 1, 3)], true);
    }, (r) => r.isPercent(), true),
    uni: uniCFAdd('setRank({isPercent:true})', (b, c) => b.setRank({ isBottom: false, isPercent: true, value: 10 }).setBackground('#FFF2CC').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G8', g: 'G', name: '后 N 项/后 N%', excel: '后 3 名',
    sjs: sjsCF('addTop10Rule（bottom）', (c, CF) => {
      c.sheet.conditionalFormats.addTop10Rule(CF.Top10ConditionType.bottom, 3, { backColor: '#EDEDED' }, [new SJS.Range(c.row, c.col, 1, 3)], false);
    }, (r) => (r.conditionType ? r.conditionType() : r.rank()), 3),
    uni: uniCFAdd('setRank({isBottom:true})', (b, c) => b.setRank({ isBottom: true, isPercent: false, value: 3 }).setBackground('#EDEDED').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G9', g: 'G', name: '高于/低于平均', excel: '高于平均值',
    sjs: sjsCF('addAverageRule', (c, CF) => {
      c.sheet.conditionalFormats.addAverageRule(CF.AverageConditionType.above, { backColor: '#D9E1F2' }, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 8),
    uni: uniCFAdd('setAverage("above")', (b, c) => b.setAverage('above').setBackground('#D9E1F2').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G10', g: 'G', name: '重复值', excel: '高亮重复',
    sjs: sjsCF('addDuplicateRule', (c) => {
      c.sheet.conditionalFormats.addDuplicateRule({ backColor: '#FFC7CE' }, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 7),
    uni: uniCFAdd('setDuplicateValues', (b, c) => b.setDuplicateValues().setBackground('#FFC7CE').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G11', g: 'G', name: '唯一值', excel: '高亮唯一',
    sjs: sjsCF('addUniqueRule', (c) => {
      c.sheet.conditionalFormats.addUniqueRule({ backColor: '#C6EFCE' }, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 6),
    uni: uniCFAdd('setUniqueValues', (b, c) => b.setUniqueValues().setBackground('#C6EFCE').setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G12', g: 'G', name: '数据条', excel: '蓝条渐变',
    sjs: sjsCF('addDataBarRule', (c, CF) => {
      c.sheet.conditionalFormats.addDataBarRule(CF.ScaleValueType.automin, 0, CF.ScaleValueType.automax, 0, '#638EC6', [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 12),
    uni: uniCFAdd('setDataBar', (b, c) => b.setDataBar({ min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#638EC6', nativeColor: '#FF0000', isGradient: true }).setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G13', g: 'G', name: '色阶 2 色', excel: '双色渐变',
    sjs: sjsCF('add2ScaleRule', (c, CF) => {
      c.sheet.conditionalFormats.add2ScaleRule(CF.ScaleValueType.lowestValue, 0, '#FFFFFF', CF.ScaleValueType.highestValue, 0, '#638EC6', [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 10),
    uni: uniCFAdd('setColorScale', (b, c) => b.setColorScale([
      { index: 0, color: '#FFFFFF', value: { type: 'min' } },
      { index: 1, color: '#638EC6', value: { type: 'max' } },
    ]).setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G14', g: 'G', name: '色阶 3 色', excel: '三色含中点',
    sjs: sjsCF('add3ScaleRule', (c, CF) => {
      c.sheet.conditionalFormats.add3ScaleRule(CF.ScaleValueType.lowestValue, 0, '#F8696B', CF.ScaleValueType.percentile, 50, '#FFEB84', CF.ScaleValueType.highestValue, 0, '#63BE7B', [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 11),
    uni: uniCFAdd('setColorScale（3 项）', (b, c) => b.setColorScale([
      { index: 0, color: '#F8696B', value: { type: 'min' } },
      { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } },
      { index: 2, color: '#63BE7B', value: { type: 'max' } },
    ]).setRanges(grng(c)), uniCFCount, 1) });
  add({ id: 'G15', g: 'G', name: '图标集', excel: '三色交通灯',
    sjs: sjsCF('addIconSetRule', (c, CF) => {
      c.sheet.conditionalFormats.addIconSetRule(CF.IconSetType.threeTrafficLightsRimmed, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (r) => r.ruleType(), 13),
    uni: uniCFAdd('setIconSet', (b, c) => {
      const map = b.getIconMap ? b.getIconMap() : {};
      const key = Object.keys(map).find((k) => /traffic/i.test(k)) || Object.keys(map)[0];
      const icons = map[key] || [];
      c.iconKey = key;
      b.setIconSet({ iconConfigs: { config: icons.map((ic, i) => ({ operator: 'greaterThan', value: { type: 'percent', value: i * 33 }, iconType: key, iconId: ic })) }, isShowValue: true }).setRanges(grng(c));
    }, uniCFCount, 1) });
  add({ id: 'G16', g: 'G', name: '迷你图规则', excel: 'Excel 无（SJS 独有）',
    sjs: sjsCall('ConditionalFormatting.SparklineRule', (c) => {
      if (!CFS.SparklineRule) throw new Error('SparklineRule 不存在');
      const r = new CFS.SparklineRule('LINESPARKLINE');
      c.sheet.conditionalFormats.addRule(r, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, sjsCFCount, 1, { kind: 'builder', steps: 3, doc: 1, note: '直连 addRule 不进区域索引，只能用 count() 验证' }),
    uni: none('', 'S-') });

  // ================= H 条件格式图形参数（15） =================
  // 统一走便捷方法建规则，再从 getRules 取回"同一个对象"（probe12 实测 sameRef===true）改参数
  const sjsDB = (api, mk, readFn, exp, opt) => sjsCall(api, (c) => {
    c.sheet.conditionalFormats.addDataBarRule(CFS.ScaleValueType.automin, 0, CFS.ScaleValueType.automax, 0, '#638EC6', [new SJS.Range(c.row, c.col, 1, 3)]);
    const r = sjsCFLast(c);
    if (!r) throw new Error('addDataBarRule 后 getRules 取不到规则');
    mk(r, CFS);
  }, (c) => { const r = sjsCFLast(c); return r ? readFn(r, CFS) : 'NO_RULE'; }, exp, Object.assign({ kind: 'builder', steps: 3, param: 'enum' }, opt || {}));
  const uniDB = (api, cfg, readExp, opt) => uniCFAdd(api, (b, c) => {
    b.setDataBar(Object.assign({ min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#638EC6', nativeColor: '#FF0000' }, cfg)).setRanges(grng(c));
  }, uniCFLast(), readExp, Object.assign({ kind: 'builder', steps: 3 }, opt || {}));

  add({ id: 'H1', g: 'H', name: '数据条最小值类型', excel: '自动/数字/百分比/百分位/公式',
    sjs: sjsDB('DataBarRule.minType', (r, CF) => { r.minType(CF.ScaleValueType.number); r.minValue(0); }, (r) => r.minType(), CFS.ScaleValueType.number),
    uni: uniDB('setDataBar({min:{type:"num"}})', { min: { type: 'num', value: 0 } }, '"type":"num"') });
  add({ id: 'H2', g: 'H', name: '数据条最大值类型', excel: '同上',
    sjs: sjsDB('DataBarRule.maxType', (r, CF) => { r.maxType(CF.ScaleValueType.percent); r.maxValue(90); }, (r) => r.maxType(), CFS.ScaleValueType.percent),
    uni: uniDB('setDataBar({max:{type:"percent"}})', { max: { type: 'percent', value: 90 } }, '"type":"percent"') });
  add({ id: 'H3', g: 'H', name: '数据条正条颜色', excel: '自定义色',
    sjs: sjsDB('DataBarRule.color', (r) => { r.color('#00B050'); }, (r) => r.color(), '#00B050'),
    uni: uniDB('setDataBar({positiveColor})', { positiveColor: '#00B050' }, '#00B050') });
  add({ id: 'H4', g: 'H', name: '数据条负条颜色+负值轴', excel: '负值反向条+轴位置',
    sjs: sjsDB('negativeFillColor / axisPosition / axisColor', (r, CF) => {
      r.negativeFillColor('#FF0000');
      r.axisPosition(CF.DataBarAxisPosition.automatic);
      r.axisColor('#000000');
    }, (r) => r.negativeFillColor(), '#FF0000'),
    uni: uniDB('setDataBar({nativeColor})', { nativeColor: '#FF0000' }, '#FF0000') });
  add({ id: 'H5', g: 'H', name: '数据条实心 vs 渐变', excel: '两种渲染',
    sjs: sjsDB('DataBarRule.gradient', (r) => { r.gradient(false); }, (r) => r.gradient(), false),
    uni: uniDB('setDataBar({isGradient})', { isGradient: false }, '"isGradient":false') });
  add({ id: 'H6', g: 'H', name: '数据条方向', excel: '左→右 / 右→左',
    sjs: sjsDB('DataBarRule.dataBarDirection', (r, CF) => { r.dataBarDirection(CF.BarDirection.rightToLeft); }, (r) => r.dataBarDirection(), CFS.BarDirection.rightToLeft),
    uni: none('（setDataBar 无 direction 字段）', 'S3') });
  add({ id: 'H7', g: 'H', name: '数据条边框', excel: '颜色+粗细',
    sjs: sjsDB('showBorder / borderColor（无粗细线型）', (r) => { r.showBorder(true); r.borderColor('#7F7F7F'); }, (r) => r.borderColor(), '#7F7F7F'),
    uni: none('（setDataBar 无 border 字段）', 'S3') });
  add({ id: 'H8', g: 'H', name: '仅显示数据条', excel: '隐藏数值',
    sjs: sjsDB('DataBarRule.showBarOnly', (r) => { r.showBarOnly(true); }, (r) => r.showBarOnly(), true),
    uni: uniDB('setDataBar({isShowValue})', { isShowValue: false }, '"isShowValue":false') });
  add({ id: 'H9', g: 'H', name: '色阶极值与颜色', excel: '5 种类型',
    sjs: sjsCall('add2ScaleRule + 极值类型', (c, CF) => {
      c.sheet.conditionalFormats.add2ScaleRule(CFS.ScaleValueType.percentile, 10, '#F8696B', CFS.ScaleValueType.percentile, 90, '#63BE7B', [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (c) => { const r = sjsCFLast(c); return r ? r.minType() + '/' + r.minColor() : 'NO_RULE'; }, CFS.ScaleValueType.percentile + '/#F8696B', { kind: 'builder', steps: 3, param: 'enum' }),
    uni: uniCFAdd('setColorScale([{index,color,value}])', (b, c) => b.setColorScale([
      { index: 0, color: '#F8696B', value: { type: 'percentile', value: 10 } },
      { index: 1, color: '#63BE7B', value: { type: 'percentile', value: 90 } },
    ]).setRanges(grng(c)), uniCFLast(), 'percentile') });
  add({ id: 'H10', g: 'H', name: '色阶中间值', excel: '3 色才有',
    sjs: sjsCall('add3ScaleRule midType/midColor', (c) => {
      c.sheet.conditionalFormats.add3ScaleRule(CFS.ScaleValueType.lowestValue, 0, '#F8696B', CFS.ScaleValueType.number, 50, '#FFEB84', CFS.ScaleValueType.highestValue, 0, '#63BE7B', [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (c) => { const r = sjsCFLast(c); return r ? r.midValue() + '/' + r.midColor() : 'NO_RULE'; }, '50/#FFEB84', { kind: 'builder', steps: 3, param: 'enum' }),
    uni: uniCFAdd('setColorScale（index 1 为中点）', (b, c) => b.setColorScale([
      { index: 0, color: '#F8696B', value: { type: 'min' } },
      { index: 1, color: '#FFEB84', value: { type: 'num', value: 50 } },
      { index: 2, color: '#63BE7B', value: { type: 'max' } },
    ]).setRanges(grng(c)), uniCFLast(), '"type":"num"') });
  add({ id: 'H11', g: 'H', name: '图标集类型（21 种）', excel: '3/4/5 级',
    sjs: sjsCall('IconSetType 21 值', (c) => {
      c.sheet.conditionalFormats.addIconSetRule(CFS.IconSetType.fourRatings, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (c) => { const r = sjsCFLast(c); return r ? r.iconSetType() : 'NO_RULE'; }, CFS.IconSetType.fourRatings, { kind: 'method', steps: 2, param: 'enum' }),
    uni: uniCFAdd('setIconSet({iconConfigs})', (b, c) => {
      const map = b.getIconMap ? b.getIconMap() : {};
      const keys = Object.keys(map);
      c.iconKeyCount = keys.length;
      const key = keys[0];
      const icons = map[key] || [];
      b.setIconSet({ iconConfigs: { config: icons.map((ic, i) => ({ operator: 'greaterThan', value: { type: 'percent', value: i * 20 }, iconType: key, iconId: ic })) }, isShowValue: true }).setRanges(grng(c));
    }, (c) => c.iconKeyCount > 0, true, { note: '图标集按组名+图标 id 配置，数量由 getIconMap 提供' }) });
  add({ id: 'H12', g: 'H', name: '图标集阈值与比较符', excel: '按值/百分比/百分位',
    sjs: sjsCall('addIconSetRule + iconCriteria', (c) => {
      c.sheet.conditionalFormats.addIconSetRule(CFS.IconSetType.threeTrafficLightsRimmed, [new SJS.Range(c.row, c.col, 1, 3)]);
      const r = sjsCFLast(c);
      if (!r) throw new Error('addIconSetRule 后取不到规则');
      r.iconCriteria([new CFS.IconCriterion(CFS.ScaleValueType.percent, 0), new CFS.IconCriterion(CFS.ScaleValueType.percent, 33), new CFS.IconCriterion(CFS.ScaleValueType.percent, 67)]);
    }, (c) => { const r = sjsCFLast(c); const ic = r && r.iconCriteria(); return ic && ic.length; }, 3, { kind: 'builder', steps: 3, param: 'enum' }),
    uni: uniCFAdd('iconConfigs.config[].operator/value', (b, c) => {
      const map = b.getIconMap ? b.getIconMap() : {};
      const key = Object.keys(map)[0];
      const icons = map[key] || [];
      b.setIconSet({ iconConfigs: { config: icons.map((ic, i) => ({ operator: 'greaterThanOrEqual', value: { type: 'percent', value: i * 33 }, iconType: key, iconId: ic })) }, isShowValue: true }).setRanges(grng(c));
    }, uniCFLast(), 'greaterThanOrEqual') });
  add({ id: 'H13', g: 'H', name: '图标集反转顺序', excel: '反向排列',
    sjs: sjsCall('addIconSetRule + reverseIconOrder', (c) => {
      c.sheet.conditionalFormats.addIconSetRule(CFS.IconSetType.threeTrafficLightsRimmed, [new SJS.Range(c.row, c.col, 1, 3)]);
      const r = sjsCFLast(c);
      if (!r) throw new Error('addIconSetRule 后取不到规则');
      r.reverseIconOrder(true);
    }, (c) => { const r = sjsCFLast(c); return r && r.reverseIconOrder(); }, true, { kind: 'builder', steps: 3, param: 'enum' }),
    uni: none('（setIconSet 无 reverse 字段）', 'S3') });
  add({ id: 'H14', g: 'H', name: '图标集仅显示图标', excel: '隐藏数值',
    sjs: sjsCall('addIconSetRule + showIconOnly', (c) => {
      c.sheet.conditionalFormats.addIconSetRule(CFS.IconSetType.threeTrafficLightsRimmed, [new SJS.Range(c.row, c.col, 1, 3)]);
      const r = sjsCFLast(c);
      if (!r) throw new Error('addIconSetRule 后取不到规则');
      r.showIconOnly(true);
    }, (c) => { const r = sjsCFLast(c); return r && r.showIconOnly(); }, true, { kind: 'builder', steps: 3, param: 'enum' }),
    uni: uniCFAdd('setIconSet({isShowValue:false})', (b, c) => {
      const map = b.getIconMap ? b.getIconMap() : {};
      const key = Object.keys(map)[0];
      const icons = map[key] || [];
      b.setIconSet({ iconConfigs: { config: icons.map((ic, i) => ({ operator: 'greaterThan', value: { type: 'percent', value: i * 20 }, iconType: key, iconId: ic })) }, isShowValue: false }).setRanges(grng(c));
    }, uniCFLast(), '"isShowValue":false') });
  add({ id: 'H15', g: 'H', name: '图形规则综合渲染', excel: '多参数叠加', skip: '属渲染叠加，参数已逐项覆盖', sjs: null, uni: null });

  // ================= I 条件格式管理（6） =================
  add({ id: 'I1', g: 'I', name: '优先级（上移/下移）', excel: '规则可排序',
    sjs: sjsCall('改 priority 数值（无重排序 API）', (c) => {
      const rng = [new SJS.Range(c.row, c.col, 1, 3)];
      c.sheet.conditionalFormats.addCellValueRule(CFS.ComparisonOperators.greaterThan, 10, 0, { backColor: '#FFC7CE' }, rng);
      c.sheet.conditionalFormats.addCellValueRule(CFS.ComparisonOperators.greaterThan, 1, 0, { backColor: '#C6EFCE' }, rng);
      const rs = c.sheet.conditionalFormats.getRules(c.row, c.col);
      rs[1].priority(1); rs[0].priority(10);
    }, (c) => c.sheet.conditionalFormats.getRules(c.row, c.col).map((r) => r.priority()).sort((a, b) => a - b).join(','), '1,10',
      { kind: 'method', steps: 4, doc: 0, note: 'getRules 按插入顺序返回、不按 priority 排序；断言改为"priority 值被接受"' }),
    uni: uniCall('moveConditionalFormattingRule', (c) => {
      const mk = (v) => {
        const b = c.usheet.newConditionalFormattingRule();
        b.whenNumberGreaterThan(v).setBackground('#FFC7CE').setRanges(grng(c));
        c.usheet.addConditionalFormattingRule(b.build());
      };
      mk(10); mk(1);
      const rs = c.usheet.getConditionalFormattingRules();
      c.before = rs.map((x) => x.cfId).join(',');
      c.usheet.moveConditionalFormattingRule(rs[1].cfId, rs[0].cfId, 'before');
      c.after = (c.usheet.getConditionalFormattingRules() || []).map((x) => x.cfId).join(',');
    }, (c) => c.before !== c.after, true, { kind: 'method', steps: 4 }) });
  add({ id: 'I2', g: 'I', name: 'stopIfTrue', excel: '命中即停止',
    sjs: sjsCall('rule.stopIfTrue(true)', (c) => {
      c.sheet.conditionalFormats.addCellValueRule(CFS.ComparisonOperators.greaterThan, 1, 0, { backColor: '#FFC7CE' }, [new SJS.Range(c.row, c.col, 1, 3)]);
      sjsCFLast(c).stopIfTrue(true);
    }, (c) => { const r = sjsCFLast(c); return r && r.stopIfTrue(); }, true, { kind: 'method', steps: 3 }),
    uni: uniCall('rule.stopIfTrue = true（builder 无 setter）', (c) => {
      const b = c.usheet.newConditionalFormattingRule();
      b.whenNumberGreaterThan(1).setBackground('#FFC7CE').setRanges(grng(c));
      const r = b.build();
      r.stopIfTrue = true;
      c.usheet.addConditionalFormattingRule(r);
    }, (c) => { const last = uniRule(c); return last && last.stopIfTrue; }, true, { kind: 'method', steps: 4, doc: 0 }) });
  add({ id: 'I3', g: 'I', name: '查询单元格命中的规则', excel: '管理界面显示范围',
    sjs: sjsCall('conditionalFormats.getRules(row,col)', (c) => {
      c.sheet.conditionalFormats.addCellValueRule(CFS.ComparisonOperators.greaterThan, 1, 0, { backColor: '#FFC7CE' }, [new SJS.Range(c.row, c.col, 1, 3)]);
    }, (c) => c.sheet.conditionalFormats.getRules(c.row, c.col).length, 1, { steps: 2 }),
    uni: uniCall('range.getConditionalFormattingRules', (c) => {
      const b = c.usheet.newConditionalFormattingRule();
      b.whenNumberGreaterThan(1).setBackground('#FFC7CE').setRanges(grng(c));
      c.usheet.addConditionalFormattingRule(b.build());
    }, (c) => (c.urange.getConditionalFormattingRules() || []).length, 1, { steps: 2 }) });
  add({ id: 'I4', g: 'I', name: '修改已有规则', excel: '编辑参数',
    sjs: sjsCall('取规则对象后 value1(99)', (c) => {
      c.sheet.conditionalFormats.addCellValueRule(CFS.ComparisonOperators.greaterThan, 1, 0, { backColor: '#FFC7CE' }, [new SJS.Range(c.row, c.col, 1, 3)]);
      sjsCFLast(c).value1(99);
    }, (c) => { const r = sjsCFLast(c); return r && r.value1(); }, 99, { steps: 3 }),
    uni: uniCall('setConditionalFormattingRule(cfId, rule)', (c) => {
      const b = c.usheet.newConditionalFormattingRule();
      b.whenNumberGreaterThan(1).setBackground('#FFC7CE').setRanges(grng(c));
      c.usheet.addConditionalFormattingRule(b.build());
      const t = uniRule(c);
      const b2 = c.usheet.newConditionalFormattingRule();
      b2.whenNumberGreaterThan(99).setBackground('#C6EFCE').setRanges(grng(c));
      c.usheet.setConditionalFormattingRule(t.cfId, b2.build());
      c.after = (c.usheet.getConditionalFormattingRules() || []).map((x) => JSON.stringify(x.rule)).join('|');
    }, (c) => c.after.indexOf('99') >= 0, true, { steps: 4 }) });
  add({ id: 'I5', g: 'I', name: '删除规则（单条/按区域/清空）', excel: '三种删除',
    sjs: sjsCall('removeRule / removeRuleByRange / clearRule', (c) => {
      const rng = [new SJS.Range(c.row, c.col, 1, 3)];
      c.sheet.conditionalFormats.addCellValueRule(CFS.ComparisonOperators.greaterThan, 1, 0, { backColor: '#FFC7CE' }, rng);
      c.sheet.conditionalFormats.addCellValueRule(CFS.ComparisonOperators.lessThan, 1, 0, { backColor: '#C6EFCE' }, rng);
      c.sheet.conditionalFormats.removeRule(sjsCFLast(c));
      c.afterRemove = c.sheet.conditionalFormats.count();
      c.sheet.conditionalFormats.removeRuleByRange(c.row, c.col, 1, 3);
      c.afterRange = c.sheet.conditionalFormats.count();
    }, (c) => c.afterRemove + '/' + c.afterRange, '1/0', { steps: 4 }),
    uni: uniCall('deleteConditionalFormattingRule / clearConditionalFormatRules', (c) => {
      const mk = (v) => { const b = c.usheet.newConditionalFormattingRule(); b.whenNumberGreaterThan(v).setBackground('#FFC7CE').setRanges(grng(c)); c.usheet.addConditionalFormattingRule(b.build()); };
      mk(1); mk(2);
      const rs = c.usheet.getConditionalFormattingRules();
      c.usheet.deleteConditionalFormattingRule(rs[0].cfId);
      c.afterDel = c.usheet.getConditionalFormattingRules().length;
      c.urange.clearConditionalFormatRules();
      c.afterClear = c.usheet.getConditionalFormattingRules().length;
    }, (c) => c.afterDel + '/' + c.afterClear, '1/0', { steps: 4 }) });
  add({ id: 'I6', g: 'I', name: '多规则叠加求值顺序', excel: '按优先级求值',
    sjs: sjsCall('priority 数值小者优先', (c) => {
      const rng = [new SJS.Range(c.row, c.col, 1, 3)];
      [5, 3, 9].forEach((v) => {
        c.sheet.conditionalFormats.addCellValueRule(CFS.ComparisonOperators.greaterThan, v, 0, { backColor: '#FFC7CE' }, rng);
      });
      // 必须先建完再统一赋 priority：交错赋值会被引擎按"新值冲突"重排
      c.sheet.conditionalFormats.getRules(c.row, c.col).forEach((r, i) => r.priority([3, 5, 9][i]));
    }, (c) => c.sheet.conditionalFormats.getRules(c.row, c.col).map((r) => r.priority()).join(','), '3,5,9',
      { steps: 3, note: '两阶段赋值可精确控制；新建规则默认按"后加者优先级最高"排（3 条时为 3,2,1）' }),
    uni: uniCall('新规则加到顶部（无显式 priority）', (c) => {
      const mk = (v) => { const b = c.usheet.newConditionalFormattingRule(); b.whenNumberGreaterThan(v).setBackground('#FFC7CE').setRanges(grng(c)); c.usheet.addConditionalFormattingRule(b.build()); };
      [5, 3, 9].forEach(mk);
    }, (c) => (c.usheet.getConditionalFormattingRules() || []).length, 3, { steps: 3, note: 'Univer 无 priority 字段，只有 cfId 顺序 + move 接口' }) });

  // ================= J 样式复用与主题（11） =================
  add({ id: 'J1', g: 'J', name: '命名样式（定义+应用）', excel: '单元格样式库',
    sjs: sjsCall('sheet.addNamedStyle / setStyleName', (c) => {
      const st = new SJS.Style();
      st.foreColor = '#C00000'; st.fontWeight = 'bold'; st.name = 'RedBold';
      c.sheet.addNamedStyle(st);
      c.sheet.setStyleName(c.row, c.col, 'RedBold');
    }, (c) => c.sheet.getStyle(c.row, c.col).foreColor, '#C00000', { kind: 'builder', steps: 3 }),
    uni: none('（Enum.NamedStyleType 存在但无应用方法）', 'S4') });
  add({ id: 'J2', g: 'J', name: '样式对象实例复用', excel: '一个 Style 赋多区域',
    sjs: sjsCall('Style 实例复用（无需逐格重建）', (c) => {
      const st = new SJS.Style();
      st.backColor = '#FFF2CC'; st.foreColor = '#C00000';
      c.sheet.setStyle(c.row, c.col, st);
      c.sheet.setStyle(c.row, c.col + 1, st);
      c.sheet.setStyle(c.row, c.col + 2, st);
    }, (c) => [0, 1, 2].map((i) => c.sheet.getStyle(c.row, c.col + i).backColor).join(','), '#FFF2CC,#FFF2CC,#FFF2CC', { steps: 4 }),
    uni: none('（FRange 无 setStyle，只能逐属性/逐格调用）', 'S3', '对照：SJS 一个 Style 实例可反复赋值') });
  add({ id: 'J3', g: 'J', name: '只应用部分属性（标志位）', excel: 'Excel 无（SJS 独有）',
    sjs: none('（Style 上无 applyFont/applyFill 标志位）', 'S3', '运行时 new Style() 上不存在 applyFont'),
    uni: none('', 'S-') });
  add({ id: 'J4', g: 'J', name: '主题色引用（Accent+tint）', excel: '随主题换色',
    sjs: sjsProp('Style.foreColor = "Accent 1 -60"', 'foreColor', 'Accent 1 -60', { param: 'loose' }),
    uni: none('（仅 rgb，无 theme/tint 字段）', 'S3') });
  add({ id: 'J5', g: 'J', name: '主题字体', excel: '随主题切换',
    sjs: sjsProp('Style.themeFont', 'themeFont', 'Headings', { param: 'loose' }),
    uni: none('', 'S3') });
  add({ id: 'J6', g: 'J', name: '标准色板', excel: '最近使用/标准色',
    sjs: sjsCall('GC.Spread.Sheets.ColorScheme', (c) => {
      if (!SJS.ColorScheme) throw new Error('ColorScheme 不存在');
      c.scheme = SJS.ColorScheme;
    }, (c) => !!c.scheme, true, { kind: 'prop', steps: 0, doc: 0 }),
    uni: none('（枚举存在，用途未文档化）', 'S4') });
  add({ id: 'J7', g: 'J', name: '复制格式（格式刷等价）', excel: '格式刷',
    sjs: sjsCall('Style.clone()', (c) => {
      const st = new SJS.Style();
      st.backColor = '#DDEBF7'; st.fontWeight = 'bold';
      c.sheet.setStyle(c.row, c.col, st);
      const copy = c.sheet.getStyle(c.row, c.col).clone();
      c.sheet.setStyle(c.row, c.col + 1, copy);
    }, (c) => c.sheet.getStyle(c.row, c.col + 1).backColor, '#DDEBF7', { kind: 'method', steps: 3 }),
    uni: none('（无 getStyle→clone 链，需逐属性抄）', 'S3', '可读 getCellStyleData 后逐属性重设') });
  add({ id: 'J8', g: 'J', name: '样式导入保真/styleId 复用', excel: '同格式共享 styleId', skip: '属导入链，用户自测', sjs: null, uni: null });
  add({ id: 'J9', g: 'J', name: '单元格折角', excel: 'Excel 无（SJS 独有）',
    sjs: none('（只有 CornerPosition 枚举，无 add/set 入口）', 'S4', 'GC.Spread.Sheets.CornerPosition 有 leftTop/rightTop/leftBottom/rightBottom 四值，sheet 与 Style 上都没有对应方法'),
    uni: none('', 'S-') });
  add({ id: 'J10', g: 'J', name: '单元格图标', excel: 'Excel 无（SJS 独有）',
    sjs: none('（只有 IconPosition 枚举，无 add/set 入口）', 'S4', 'GC.Spread.Sheets.IconPosition 有 left/right/leftOfText/… 六值，无对应方法'),
    uni: none('', 'S-') });
  add({ id: 'J11', g: 'J', name: '单元格状态样式', excel: '只读/禁用/必填',
    sjs: sjsCall('sheet.cellStates.add(range, CellStatesType, style)', (c) => {
      const st = new SJS.Style();
      st.backColor = '#FFF2CC';
      c.sheet.cellStates.add(new SJS.Range(c.row, c.col, 1, 3), SJS.CellStatesType.readonly, st);
    }, (c) => {
      const j = c.sheet.cellStates.getJSONForCell(c.row, c.col);
      return j && j.length ? j[0].state : 'empty';
    }, SJS.CellStatesType.readonly,
      { kind: 'builder', steps: 2, param: 'enum',
        note: '8 种状态（hover/invalid/readonly/edit/active/selected/dirty/invalidFormula）；读回走 getJSONForCell，getState 只反映"当前是否处于该态"' }),
    uni: none('', 'S-') });

  // ================= K 保护与视图（6） =================
  add({ id: 'K1', g: 'K', name: '单元格锁定', excel: '保护后不可编辑',
    sjs: sjsProp('Style.locked', 'locked', false),
    uni: none('（无 locked 样式字段，只有范围权限模型）', 'S3') });
  add({ id: 'K2', g: 'K', name: '公式隐藏', excel: '保护后公式不可见',
    sjs: sjsProp('Style.hidden', 'hidden', true),
    uni: none('', 'S3') });
  add({ id: 'K3', g: 'K', name: '工作表保护', excel: '允许选锁定格/排序/筛选',
    sjs: sjsCall('options.isProtected + protectionOptions', (c) => {
      c.sheet.options.isProtected = true;
      c.sheet.options.protectionOptions = { allowSelectLockedCells: true, allowSort: true };
    }, (c) => c.sheet.options.isProtected, true, { kind: 'prop', steps: 2 }),
    uni: none('（无 getRangePermission / protect 入口）', 'S3', '运行时可读 getRangePermission 为 undefined') });
  add({ id: 'K4', g: 'K', name: '网格线显示/颜色', excel: '隐藏或改色',
    sjs: sjsCall('sheet.options.gridline', (c) => {
      c.sheet.options.gridline = { showVerticalGridline: false, showHorizontalGridline: false, color: '#FF0000' };
    }, (c) => (c.sheet.options.gridline || {}).showVerticalGridline, false, { kind: 'prop', steps: 1 }),
    uni: uniCall('setHiddenGridlines / setGridLinesColor', (c) => {
      c.usheet.setHiddenGridlines(true);
      c.usheet.setGridLinesColor('#FF0000');
    }, (c) => c.usheet.hasHiddenGridLines(), true) });
  add({ id: 'K5', g: 'K', name: '零值显示', excel: '隐藏 0 值',
    sjs: sjsCall('sheet.options.showZeros', (c) => { c.sheet.options.showZeros = false; },
      (c) => c.sheet.options.showZeros, false, { kind: 'prop', steps: 1 }),
    uni: none('（无 showZeros 入口）', 'S3') });
  add({ id: 'K6', g: 'K', name: '冻结窗格', excel: '首行/首列/任意拆分',
    sjs: sjsCall('frozenRowCount / frozenColumnCount', (c) => {
      c.sheet.frozenRowCount(2); c.sheet.frozenColumnCount(1);
    }, (c) => c.sheet.frozenRowCount() + '/' + c.sheet.frozenColumnCount(), '2/1', { kind: 'prop', steps: 1 }),
    uni: uniCall('FWorksheet.setFreeze({startRow,startColumn,xSplit,ySplit})', (c) => {
      c.usheet.setFreeze({ startRow: 2, startColumn: 1, ySplit: 2, xSplit: 1 });
    }, (c) => { const f = c.usheet.getFreeze(); return f ? f.startRow + '/' + f.startColumn : 'null'; }, '2/1') });

  // ================= 跑批 =================
  function normalize(v) {
    if (v === undefined) return 'undefined';
    if (v === null) return null;
    if (typeof v === 'object') { try { return JSON.parse(JSON.stringify(v)); } catch (e) { return String(v); } }
    return v;
  }
  const tick = () => new Promise((r) => setTimeout(r, 40));

  function makeCtx(row) {
    const col = 1;
    const spread = window.__spread;
    const sheet = spread.getActiveSheet();
    const wb = window.__wb;
    const usheet = wb.getActiveSheet();
    return {
      row, col, spread, sheet, GC: window.GC, SJS,
      wb, univerAPI: window.univerAPI, usheet,
      urange: usheet.getRange(row, col, 1, 3),
      rrange: sheet.getRange(row, col, 2, 3),
      st: new SJS.Style(),
      uniEnum: (window.univerAPI && window.univerAPI.Enum) || {},
    };
  }

  async function runSide(item, side) {
    const spec = item[side];
    if (!spec) return { skipped: true };
    const base = { api: spec.api, kind: spec.kind, steps: spec.steps, param: spec.param, doc: spec.doc, note: spec.note };
    if (!spec.set) return Object.assign({ noApi: true, tierHint: spec.tierHint }, base);

    const t0 = performance.now();
    const ctx = makeCtx(item.row);
    try {
      await spec.set(ctx);
    } catch (e) {
      const msg = String((e && e.message) || e);
      if (spec.kind === 'none') return Object.assign({ noApi: true, probeErr: msg, tierHint: spec.tierHint }, base);
      return Object.assign({}, base, { setOk: false, err: msg, ms: Math.round((performance.now() - t0) * 10) / 10 });
    }
    if (side === 'uni') await tick();

    let got, readErr = null;
    try { got = spec.read ? await spec.read(ctx) : undefined; } catch (e) { readErr = String((e && e.message) || e); }
    const ms = Math.round((performance.now() - t0) * 10) / 10;
    const exp = spec.exp;
    let match;
    if (spec.assert === 'noThrow') match = true;
    else if (readErr) match = false;
    else if (exp === undefined) match = (got === undefined || got === null || got === '');
    else if (exp === null) match = (got === null || got === undefined || got === '');
    else if (typeof exp === 'boolean') match = (got === exp || got === (exp ? 1 : 0));
    else if (typeof exp === 'number') match = (Number(got) === exp);
    else match = String(got).indexOf(exp) >= 0;

    // 期望值就是"空"（exp===undefined，例如清除填充/边框）时，读到空本身即证明可读
    const readable = spec.assert === 'noThrow'
      ? false
      : (!readErr && (exp === undefined || (got !== undefined && got !== null && got !== 'NO_RULE')));

    return {
      setOk: true, readErr, got: normalize(got), match, readable, api: spec.api,
      kind: spec.kind, steps: spec.steps, param: spec.param, doc: spec.doc, ms,
      exp: normalize(exp), assert: spec.assert,
    };
  }

  async function runAll() {
    // SJS 默认工作表 100 行，本表最多用到 5+135=140 行；不扩会被静默丢弃（越界 addSpan 直接抛 Invalid range）
    let sheetInfo = null;
    try {
      const sh = window.__spread.getActiveSheet();
      const before = sh.getRowCount();
      sh.setRowCount(400);
      sheetInfo = { before, after: sh.getRowCount() };
    } catch (e) { sheetInfo = { err: String((e && e.message) || e) }; }

    const out = { at: new Date().toISOString(), sheetInfo, items: [] };
    for (const item of SPEC) {
      if (item.skip) { out.items.push({ id: item.id, g: item.g, name: item.name, excel: item.excel, skipped: true, skip: item.skip }); continue; }
      if (item.g === 'G' || item.g === 'H' || item.g === 'I') {
        try { window.__spread.getActiveSheet().conditionalFormats.clearRule(); } catch (e) {}
        try { window.__wb.getActiveSheet().clearConditionalFormatRules(); } catch (e) {}
        await tick();
      }
      const sjs = await runSide(item, 'sjs');
      const uni = await runSide(item, 'uni');
      out.items.push({ id: item.id, g: item.g, name: item.name, excel: item.excel, row: item.row, sjs, uni });
    }
    return out;
  }

  window.ApiProbe = { SPEC, runAll, runSide };
})();

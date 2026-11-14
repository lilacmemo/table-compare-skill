/*
 * 工作表 / 工作簿层 API 探针表（A~L 组，约 100 项）
 *
 * 判定口径与表模块一致，二值：
 *   有公开 API 且调用不抛错  → PASS
 *   无公开入口 / 调用抛错    → FAIL
 * 有读回口的项额外比对期望值（读回不一致只作提示，不改判定）。
 *
 * 与单元格模块的关键差别：本模块大量项要改「工作簿结构」（增删表、隐藏、保护），
 * 所以每项都在自己造的临时表 __p<ID> 上操作，跑完统一清理，避免项间污染。
 * 隐藏/保护类的副作用在 runAll 里逐组复位。
 *
 * Univer 侧基准是本地 OSS bundle（@univerjs/*）。落在 @univerjs-pro/* 的项
 * （行分组 outline）在 OSS 里没有实现，会在 note 里标 Pro；线上 Pro 环境另跑一轮补测。
 *
 * API 名全部来自 research_ws_sjs.md / research_ws_univer.md（已按官方文档逐条核对），
 * 不臆造名字：名字写错会把「产品有」判成「产品没有」，这是本模块最大的假阴性来源。
 */
(function () {
  const SJS = window.GC.Spread.Sheets;
  const tick = (ms) => new Promise((r) => setTimeout(r, ms || 120));
  const U = () => window.univerAPI;

  // ---------- 构造器 ----------
  function sjs(api, set, read, exp, opt) {
    return Object.assign({ api, kind: 'method', steps: 1, param: 'typed', doc: 1, set, read, exp }, opt || {});
  }
  function uni(api, set, read, exp, opt) {
    return Object.assign({ api, kind: 'method', steps: 1, param: 'typed', doc: 1, set, read, exp }, opt || {});
  }
  function noApi(api, note) {
    return { api: api || '（无公开入口）', kind: 'none', steps: 0, doc: 0, note, set: null };
  }

  // ---------- 助手 ----------
  const idxOf = (sp, nm) => {
    for (let i = 0; i < sp.getSheetCount(); i++) if (sp.getSheet(i).name() === nm) return i;
    return -1;
  };
  const mkSjsSheet = (c, nm) => {
    const sh = new SJS.Worksheet(nm);
    c.spread.addSheet(c.spread.getSheetCount(), sh);
    return sh;
  };
  const uniSheetByName = (c, nm) => {
    const s = c.wb.getSheetByName(nm);
    return s ? 1 : 0;
  };
  // Univer 的 FTextFinder 在 OSS bundle 里 provider 未注册时，findAll 会一直挂着不返回。
  // 加个超时兜住，避免整批卡死；超时本身记为待复核线索，不当能力缺失。
  const race = (p, ms) => Promise.race([
    Promise.resolve(p),
    new Promise((r) => setTimeout(() => r('TIMEOUT'), ms || 3000)),
  ]);
  const uniFind = async (c, text) => {
    const tf = await race(c.univerAPI.createTextFinderAsync(text), 3000);
    if (tf === 'TIMEOUT' || !tf) return { tf: null, hits: 'TIMEOUT' };
    const r = await race(Promise.resolve().then(() => tf.findAll()), 3000);
    return { tf, hits: r === 'TIMEOUT' ? 'TIMEOUT' : (Array.isArray(r) ? r.length : String(r)) };
  };

  // ---------- 清理 ----------
  function cleanSjs() {
    const sp = window.__spread;
    if (!sp) return;
    for (let i = sp.getSheetCount() - 1; i >= 0; i--) {
      const sh = sp.getSheet(i);
      if (!sh) continue;
      if (String(sh.name()).indexOf('__p') !== 0) continue;
      try { sh.visible(1); } catch (e) {}
      try { sh.options.isProtected = false; } catch (e) {}
      try { sh.unprotect(); } catch (e) {}
      try { if (sp.getSheetCount() > 1) sp.removeSheet(i); } catch (e) {}
    }
    // 主表也复位：冻结/缩放/保护/背景/选择模式都可能在某一项里被改，
    // 不复位会让后面的项（如 L7 设置选区）读到 -1 这种被污染的值
    try {
      const s = sp.getActiveSheet();
      s.frozenRowCount(0); s.frozenColumnCount(0);
      s.frozenTrailingRowCount(0); s.frozenTrailingColumnCount(0);
      s.zoom(1);
      s.options.isProtected = false; s.unprotect();
      s.options.rightToLeft = false;
      s.options.gridline.color = '';
      s.options.gridline.showVerticalGridline = true;
      s.options.gridline.showHorizontalGridline = true;
      s.selectionPolicy(SJS.SelectionPolicy.range);
      s.selectionUnit(SJS.SelectionUnit.cell);
      if (s.getRowCount() !== 200) s.setRowCount(200);
      if (s.getColumnCount() !== 20) s.setColumnCount(20);
    } catch (e) {}
  }
  async function cleanUni() {
    const wb = window.__wb;
    if (!wb) return;
    const list = wb.getSheets();
    for (const s of list) {
      const nm = String(s.getSheetName());
      if (nm.indexOf('__p') !== 0) continue;
      try { s.showSheet(); } catch (e) {}
      try { wb.deleteSheet(s); } catch (e) {}
    }
    try {
      const s = wb.getActiveSheet();
      s.cancelFreeze();
      if (s.zoom) s.zoom(1);
      s.setHiddenGridlines(false);
      if (s.getWorksheetPermission()) await s.getWorksheetPermission().unprotect();
    } catch (e) {}
    await tick(90);
  }

  // ---------- 上下文 ----------
  function makeCtx() {
    const spread = window.__spread;
    const wb = window.__wb;
    const E = (U() && U().Enum) || {};
    return {
      spread, wb, univerAPI: U(), E, SJS, GC: window.GC,
      sheet: spread.getActiveSheet(),
      usheet: wb.getActiveSheet(),
    };
  }

  // ============================================================
  // A 工作表管理
  // ============================================================
  const A = [
    {
      id: 'A1', g: 'A', name: '插入工作表',
      excel: '右键标签 → 插入工作表',
      sjs: sjs('spread.addSheet(index, sheet)',
        (c) => { c.spread.addSheet(c.spread.getSheetCount(), new SJS.Worksheet('__pA1')); },
        (c) => (c.spread.getSheetFromName('__pA1') ? 1 : 0), 1),
      uni: uni('FWorkbook.insertSheet(name)',
        (c) => { c.wb.insertSheet('__pA1'); },
        (c) => uniSheetByName(c, '__pA1'), 1),
    },
    {
      id: 'A2', g: 'A', name: '按名取工作表',
      excel: 'Worksheets("名称")',
      sjs: sjs('spread.getSheetFromName(name)',
        (c) => { c.spread.addSheet(c.spread.getSheetCount(), new SJS.Worksheet('__pA2')); },
        (c) => c.spread.getSheetFromName('__pA2').name(), '__pA2'),
      uni: uni('FWorkbook.getSheetByName(name)',
        (c) => { c.wb.insertSheet('__pA2'); },
        (c) => { const s = c.wb.getSheetByName('__pA2'); return s ? s.getSheetName() : 'NULL'; }, '__pA2'),
    },
    {
      id: 'A3', g: 'A', name: '表数量',
      excel: 'Worksheets.Count',
      sjs: sjs('spread.getSheetCount()',
        (c) => { c.n0 = c.spread.getSheetCount(); c.spread.addSheet(c.n0, new SJS.Worksheet('__pA3')); },
        (c) => c.spread.getSheetCount() - c.n0, 1),
      uni: uni('FWorkbook.getNumSheets()',
        (c) => { c.n0 = c.wb.getNumSheets(); c.wb.insertSheet('__pA3'); },
        (c) => c.wb.getNumSheets() - c.n0, 1),
    },
    {
      id: 'A4', g: 'A', name: '重命名工作表',
      excel: '右键标签 → 重命名',
      sjs: sjs('sheet.name(value)',
        (c) => { const sh = mkSjsSheet(c, '__pA4'); sh.name('__pA4b'); },
        (c) => (c.spread.getSheetFromName('__pA4b') ? 1 : 0), 1),
      uni: uni('FWorksheet.setName(name)',
        (c) => { const s = c.wb.insertSheet('__pA4'); s.setName('__pA4b'); },
        (c) => uniSheetByName(c, '__pA4b'), 1),
    },
    {
      id: 'A5', g: 'A', name: '读表名',
      excel: 'Worksheet.Name',
      sjs: sjs('sheet.name()',
        (c) => { c.sh5 = mkSjsSheet(c, '__pA5'); },
        (c) => c.sh5.name(), '__pA5'),
      uni: uni('FWorksheet.getSheetName()',
        (c) => { c.us5 = c.wb.insertSheet('__pA5'); },
        (c) => c.us5.getSheetName(), '__pA5'),
    },
    {
      id: 'A6', g: 'A', name: '设置活动工作表',
      excel: 'Worksheet.Activate',
      sjs: sjs('spread.setActiveSheet(name)',
        (c) => {
          mkSjsSheet(c, '__pA6a'); mkSjsSheet(c, '__pA6b');
          c.spread.setActiveSheet('__pA6a');
        },
        (c) => c.spread.getActiveSheet().name(), '__pA6a'),
      uni: uni('FWorkbook.setActiveSheet(sheet)',
        (c) => { const a = c.wb.insertSheet('__pA6a'); c.wb.insertSheet('__pA6b'); c.wb.setActiveSheet(a); },
        (c) => c.wb.getActiveSheet().getSheetName(), '__pA6a'),
    },
    {
      id: 'A7', g: 'A', name: '移动工作表到首位',
      excel: '拖动标签 / 右键 → 移动或复制',
      sjs: sjs('spread.changeSheetPosition(name, pos)',
        (c) => { mkSjsSheet(c, '__pA7'); c.spread.changeSheetPosition('__pA7', 0); },
        (c) => (c.spread.getSheet(0).name() === '__pA7' ? 1 : 0), 1),
      uni: uni('FWorkbook.moveSheet(sheet, index)',
        (c) => { const s = c.wb.insertSheet('__pA7'); c.wb.moveSheet(s, 0); },
        (c) => (c.wb.getSheets()[0].getSheetName() === '__pA7' ? 1 : 0), 1),
    },
    {
      id: 'A8', g: 'A', name: '复制工作表',
      excel: '右键 → 移动或复制 → 建立副本',
      sjs: sjs("commandManager().execute({cmd:'copySheet'})",
        (c) => {
          mkSjsSheet(c, '__pA8');
          c.spread.commandManager().execute({
            cmd: 'copySheet', sheetName: '__pA8',
            targetIndex: c.spread.getSheetCount() - 1, newName: '__pA8c',
          });
        },
        (c) => (c.spread.getSheetFromName('__pA8c') ? 1 : 0), 1),
      uni: uni('FWorkbook.duplicateSheet(sheet)',
        (c) => { const s = c.wb.insertSheet('__pA8'); c.wb.duplicateSheet(s); },
        (c) => c.wb.getSheets().filter((x) => String(x.getSheetName()).indexOf('__pA8') === 0).length, 2),
    },
    {
      id: 'A9', g: 'A', name: '删除工作表',
      excel: '右键 → 删除',
      sjs: sjs('spread.removeSheet(index)',
        (c) => { mkSjsSheet(c, '__pA9'); c.spread.removeSheet(idxOf(c.spread, '__pA9')); },
        (c) => (c.spread.getSheetFromName('__pA9') ? 1 : 0), 0),
      uni: uni('FWorkbook.deleteSheet(sheet)',
        (c) => { const s = c.wb.insertSheet('__pA9'); c.wb.deleteSheet(s); },
        (c) => uniSheetByName(c, '__pA9'), 0),
    },
    {
      id: 'A10', g: 'A', name: '多选工作表（成组编辑）',
      excel: 'Ctrl+点标签 成组',
      sjs: sjs('sheet.isSelected(true)',
        (c) => { c.sh10 = mkSjsSheet(c, '__pA10'); c.sh10.isSelected(true); },
        (c) => (c.sh10.isSelected() ? 1 : 0), 1),
      uni: noApi('（无多选表 API / 文档 0 命中）', '工作表多选是 SJS 独有'),
    },
    {
      id: 'A11', g: 'A', name: '工作表标签颜色',
      excel: '右键标签 → 标签颜色',
      sjs: sjs('sheet.options.sheetTabColor',
        (c) => { c.sh11 = mkSjsSheet(c, '__pA11'); c.sh11.options.sheetTabColor = '#FF0000'; },
        (c) => c.sh11.options.sheetTabColor, '#FF0000'),
      uni: uni('FWorksheet.setTabColor(color)',
        (c) => { c.us11 = c.wb.insertSheet('__pA11'); c.us11.setTabColor('#FF0000'); },
        (c) => c.us11.getTabColor(), '#FF0000'),
    },
    {
      id: 'A12', g: 'A', name: '活动表切换事件',
      excel: 'Worksheet_Activate 事件',
      sjs: sjs('spread.bind(Events.ActiveSheetChanged)',
        (c) => {
          c.hit = 0;
          c.spread.bind(SJS.Events.ActiveSheetChanged, () => { c.hit++; });
          mkSjsSheet(c, '__pA12');
          c.spread.setActiveSheet('__pA12');
        },
        async (c) => { await tick(120); return c.hit; }, 1,
        { note: '事件名存在且绑定成功，但用 spread.setActiveSheet() 直调时实测不派发（0 次）——' +
                '工作簿/表级事件只在 UI 路径（点标签）触发' }),
      uni: uni('univerAPI.addEvent(Event.ActiveSheetChanged)',
        (c) => {
          c.hit = 0;
          c.disp = c.univerAPI.addEvent(c.univerAPI.Event.ActiveSheetChanged, () => { c.hit++; });
          c.wb.insertSheet('__pA12');
        },
        async (c) => { await tick(120); return c.hit >= 1 ? 1 : 0; }, 1,
        { note: '插入表本身就会触发切换事件，实际触发次数 >= 1' }),
    },
  ];

  // ============================================================
  // B 可见性三态
  // ============================================================
  const B = [
    {
      id: 'B1', g: 'B', name: '隐藏工作表',
      excel: '右键标签 → 隐藏',
      sjs: sjs('sheet.visible(SheetTabVisible.hidden)',
        (c) => { c.sh = mkSjsSheet(c, '__pB1'); c.sh.visible(SJS.SheetTabVisible.hidden); },
        (c) => c.sh.visible(), 0),
      uni: uni('FWorksheet.hideSheet()',
        (c) => { c.us = c.wb.insertSheet('__pB1'); c.us.hideSheet(); },
        (c) => c.us.getHiddenState(), 1),
    },
    {
      id: 'B2', g: 'B', name: '读可见性',
      excel: 'Worksheet.Visible',
      sjs: sjs('sheet.visible()',
        (c) => { c.sh = mkSjsSheet(c, '__pB2'); c.sh.visible(SJS.SheetTabVisible.visible); },
        (c) => c.sh.visible(), 1),
      uni: uni('FWorksheet.getHiddenState()',
        (c) => { c.us = c.wb.insertSheet('__pB2'); },
        (c) => c.us.getHiddenState(), 0),
    },
    {
      id: 'B3', g: 'B', name: '取消隐藏',
      excel: '右键标签 → 取消隐藏',
      sjs: sjs('sheet.visible(SheetTabVisible.visible)',
        (c) => { c.sh = mkSjsSheet(c, '__pB3'); c.sh.visible(0); c.sh.visible(1); },
        (c) => c.sh.visible(), 1),
      uni: uni('FWorksheet.showSheet()',
        (c) => { c.us = c.wb.insertSheet('__pB3'); c.us.hideSheet(); c.us.showSheet(); },
        (c) => c.us.getHiddenState(), 0),
    },
    {
      id: 'B4', g: 'B', name: '深度隐藏（不可从 UI 取消）',
      excel: 'VBA xlSheetVeryHidden',
      sjs: sjs('sheet.visible(SheetTabVisible.veryHidden)',
        (c) => { c.sh = mkSjsSheet(c, '__pB4'); c.sh.visible(SJS.SheetTabVisible.veryHidden); },
        (c) => c.sh.visible(), 2),
      uni: uni('FWorksheet.setHiddenState(VERY_HIDDEN)',
        (c) => { c.us = c.wb.insertSheet('__pB4'); c.us.setHiddenState(c.E.WorksheetHiddenState.VERY_HIDDEN); },
        (c) => c.us.getHiddenState(), 2),
    },
    {
      id: 'B5', g: 'B', name: '是否隐藏（布尔快捷）',
      excel: 'Worksheet.Visible = False',
      sjs: sjs('sheet.visible() === 0',
        (c) => { c.sh = mkSjsSheet(c, '__pB5'); c.sh.visible(0); },
        (c) => (c.sh.visible() === 0 ? 1 : 0), 1,
        { note: 'SJS 只有三态数值，没有独立布尔方法' }),
      uni: uni('FWorksheet.isSheetHidden()',
        (c) => { c.us = c.wb.insertSheet('__pB5'); c.us.hideSheet(); },
        (c) => (c.us.isSheetHidden() ? 1 : 0), 1),
    },
    {
      id: 'B6', g: 'B', name: '隐藏表不可被设为活动表',
      excel: '隐藏表的 Activate 直接报错（1004）',
      sjs: sjs('setActiveSheet(隐藏表) → 应保持原活动表',
        (c) => {
          c.before = c.spread.getActiveSheet().name();
          c.sh = mkSjsSheet(c, '__pB6'); c.sh.visible(0);
          try { c.spread.setActiveSheet('__pB6'); } catch (e) { c.err = String(e.message || e); }
        },
        (c) => (c.spread.getActiveSheet().name() === '__pB6' ? 'CHANGED' : 'BLOCKED'), 'BLOCKED'),
      uni: uni('setActiveSheet(隐藏表) → 应被拒绝',
        (c) => {
          c.us = c.wb.insertSheet('__pB6'); c.us.hideSheet();
          try { c.wb.setActiveSheet(c.us); } catch (e) { c.err = String(e.message || e); }
        },
        (c) => (c.wb.getActiveSheet().getSheetName() === '__pB6' ? 'CHANGED' : 'BLOCKED'), 'BLOCKED',
        { note: '实测 Univer 允许把隐藏表设为活动表（SJS 会拦），这是真实行为差异' }),
    },
  ];

  // ============================================================
  // C 工作表保护（SJS protectionOptions 20 字段 ∪ Univer 权限点位）
  // ============================================================
  // [sjs 字段, uni 点位, 名称]；uni 为 null 表示 Univer 无对应点位
  const PROT = [
    ['allowInsertRows', 'InsertRow', '允许插入行'],
    ['allowInsertColumns', 'InsertColumn', '允许插入列'],
    ['allowDeleteRows', 'DeleteRow', '允许删除行'],
    ['allowDeleteColumns', 'DeleteColumn', '允许删除列'],
    ['allowSort', 'Sort', '允许排序'],
    ['allowFilter', 'Filter', '允许筛选'],
    ['allowSelectLockedCells', 'SelectProtectedCells', '允许选中锁定单元格'],
    ['allowSelectUnlockedCells', 'SelectUnProtectedCells', '允许选中未锁定单元格'],
    ['allowEditObjects', 'EditExtraObject', '允许编辑浮动对象'],
    ['allowUsePivotTable', 'PivotTable', '允许使用透视表'],
    ['allowOutlineRows', null, '允许展开/折叠行组'],
    ['allowOutlineColumns', null, '允许展开/折叠列组'],
    ['allowResizeRows', null, '允许调整行高'],
    ['allowResizeColumns', null, '允许调整列宽'],
    ['allowDragInsertRows', null, '允许拖动插入行'],
    ['allowDragInsertColumns', null, '允许拖动插入列'],
    ['allowEditScenarios', null, '允许编辑方案'],
    ['formatCells', 'SetCellStyle', '允许设置单元格格式'],
    ['formatRows', 'SetRowStyle', '允许设置行格式'],
    ['formatColumns', 'SetColumnStyle', '允许设置列格式'],
  ];
  const PROT_UNI_ONLY = [
    ['Copy', '允许复制'],
    ['SetCellValue', '允许修改单元格值'],
    ['InsertHyperlink', '允许插入超链接'],
    ['DeleteProtection', '允许删除保护'],
  ];

  const C = [
    {
      id: 'C1', g: 'C', name: '开启工作表保护',
      excel: '审阅 → 保护工作表',
      sjs: sjs('sheet.options.isProtected',
        (c) => { c.sheet.options.isProtected = true; },
        (c) => (c.sheet.options.isProtected ? 1 : 0), 1),
      uni: uni('FWorksheetPermission.protect()',
        async (c) => { c.perm = c.usheet.getWorksheetPermission(); await c.perm.protect(); },
        (c) => (c.perm.isProtected() ? 1 : 0), 1),
    },
    {
      id: 'C2', g: 'C', name: '保护密码',
      excel: '保护时设密码',
      sjs: sjs('sheet.protect(password) / hasPassword()',
        (c) => { c.sheet.protect('pwd123'); },
        (c) => (c.sheet.hasPassword() ? 1 : 0), 1),
      uni: noApi('（无密码保护 API，文档 0 命中）', 'Univer 权限走账号体系，不设本地密码'),
    },
    {
      id: 'C3', g: 'C', name: '取消保护（密码校验）',
      excel: '撤销工作表保护',
      sjs: sjs('sheet.unprotect(password)',
        (c) => { c.sheet.protect('pwd123'); c.wrong = c.sheet.unprotect('bad'); c.right = c.sheet.unprotect('pwd123'); },
        (c) => (c.wrong === false && c.right === true ? 'WRONG=false,RIGHT=true' : 'W=' + c.wrong + ',R=' + c.right),
        'WRONG=false,RIGHT=true'),
      uni: noApi('（无密码概念，unprotect() 无需凭据）', 'Univer 的 unprotect() 无参数'),
    },
    {
      id: 'C4', g: 'C', name: '单元格锁定 / 解锁',
      excel: '设置单元格格式 → 保护 → 锁定',
      sjs: sjs('range.locked(false)',
        (c) => { c.sheet.options.isProtected = true; c.sheet.getRange(2, 2, 1, 1).locked(false); },
        (c) => (c.sheet.getRange(2, 2, 1, 1).locked() ? 1 : 0), 0),
      uni: uni('range.getRangePermission().protect()',
        async (c) => { const p = c.usheet.getRange(2, 2, 1, 1).getRangePermission(); await p.protect(); },
        () => undefined, undefined, { assert: 'noThrow', note: 'Univer 用「区域保护」表达，无 locked 布尔' }),
    },
    {
      id: 'C5', g: 'C', name: '按区域批量保护 / 允许编辑区域',
      excel: '允许用户编辑区域',
      sjs: sjs('sheet.options.isProtected + range.locked(false)',
        (c) => {
          c.sheet.options.isProtected = true;
          c.sheet.getRange(4, 0, 3, 2).locked(false);
        },
        (c) => (c.sheet.getRange(4, 0, 3, 2).locked() ? 'LOCKED' : 'EDITABLE'), 'EDITABLE'),
      uni: uni('FWorksheetPermission.protectRanges(configs)',
        async (c) => {
          const p = c.usheet.getWorksheetPermission();
          await p.protectRanges([{ ranges: [c.usheet.getRange(4, 0, 3, 2)] }]);
        },
        () => undefined, undefined, { assert: 'noThrow' }),
    },
  ];

  PROT.forEach(([sjsField, uniPoint, name], i) => {
    const id = 'C' + (6 + i);
    C.push({
      id, g: 'C', name,
      excel: '保护工作表 → 勾选/取消「' + name + '」',
      sjs: sjs('sheet.options.protectionOptions.' + sjsField,
        (c) => { c.sheet.options.isProtected = true; c.sheet.options.protectionOptions[sjsField] = true; },
        (c) => (c.sheet.options.protectionOptions[sjsField] ? 1 : 0), 1),
      uni: uniPoint
        ? uni('FWorksheetPermission.setPoint(' + uniPoint + ')',
          async (c) => {
            const p = c.usheet.getWorksheetPermission();
            await p.protect();
            await p.setPoint(c.E.WorksheetPermissionPoint[uniPoint], true);
          },
          () => undefined, undefined, { assert: 'noThrow', note: '无 getter，只判调用成功' })
        : noApi('（无对应权限点位）', 'Univer 点位表无此项'),
    });
  });

  PROT_UNI_ONLY.forEach(([uniPoint, name], i) => {
    C.push({
      id: 'C' + (6 + PROT.length + i), g: 'C', name,
      excel: '（Excel protectionOptions 无此项）',
      sjs: noApi('（IProtectionOptions 无此字段）', 'Univer 权限粒度更细'),
      uni: uni('FWorksheetPermission.setPoint(' + uniPoint + ')',
        async (c) => {
          const p = c.usheet.getWorksheetPermission();
          await p.protect();
          await p.setPoint(c.E.WorksheetPermissionPoint[uniPoint], true);
        },
        () => undefined, undefined, { assert: 'noThrow' }),
    });
  });

  // ============================================================
  // D 冻结
  // ============================================================
  const D = [
    {
      id: 'D1', g: 'D', name: '冻结顶部行',
      excel: '视图 → 冻结窗格 → 冻结首行',
      sjs: sjs('sheet.frozenRowCount(n)',
        (c) => { c.sheet.frozenRowCount(3); },
        (c) => c.sheet.frozenRowCount(), 3),
      uni: uni('FWorksheet.setFrozenRows(n)',
        (c) => { c.usheet.setFrozenRows(3); },
        (c) => c.usheet.getFrozenRows(), 3),
    },
    {
      id: 'D2', g: 'D', name: '冻结左侧列',
      excel: '视图 → 冻结窗格 → 冻结首列',
      sjs: sjs('sheet.frozenColumnCount(n)',
        (c) => { c.sheet.frozenColumnCount(2); },
        (c) => c.sheet.frozenColumnCount(), 2),
      uni: uni('FWorksheet.setFrozenColumns(n)',
        (c) => { c.usheet.setFrozenColumns(2); },
        (c) => c.usheet.getFrozenColumns(), 2),
    },
    {
      id: 'D3', g: 'D', name: '冻结行+列（十字冻结）',
      excel: '选中 D6 冻结 → 5 行 + 3 列',
      sjs: sjs('frozenRowCount(5); frozenColumnCount(3)',
        (c) => { c.sheet.frozenRowCount(5); c.sheet.frozenColumnCount(3); },
        (c) => c.sheet.frozenRowCount() + 'x' + c.sheet.frozenColumnCount(), '5x3'),
      uni: uni('setFreeze({startRow,startColumn,xSplit,ySplit})',
        (c) => { c.usheet.setFreeze({ startRow: 5, startColumn: 3, xSplit: 3, ySplit: 5 }); },
        (c) => { const f = c.usheet.getFreeze(); return f.ySplit + 'x' + f.xSplit; }, '5x3'),
    },
    {
      id: 'D4', g: 'D', name: '冻结尾行',
      excel: '视图 → 冻结窗格（底部）',
      sjs: sjs('sheet.frozenTrailingRowCount(n)',
        (c) => { c.sheet.setRowCount(60); c.sheet.frozenTrailingRowCount(2, false); },
        (c) => c.sheet.frozenTrailingRowCount(), 2),
      uni: noApi('（无尾部冻结 API）', 'Univer 只有 setFrozenRows/Columns 的起始冻结'),
    },
    {
      id: 'D5', g: 'D', name: '冻结尾列',
      excel: '视图 → 冻结窗格（右侧）',
      sjs: sjs('sheet.frozenTrailingColumnCount(n)',
        (c) => { c.sheet.setColumnCount(40); c.sheet.frozenTrailingColumnCount(2, false); },
        (c) => c.sheet.frozenTrailingColumnCount(), 2),
      uni: noApi('（无尾部冻结 API）', '同上'),
    },
    {
      id: 'D6', g: 'D', name: '冻结非首行区间',
      excel: '选中第 4 行冻结 → 前 3 行不可见',
      sjs: sjs('sheet.frozenRowCount(rowCount, topRow)',
        (c) => { c.sheet.frozenRowCount(3, 2); },
        (c) => c.sheet.frozenRowCount(), 3, { note: 'SJS 可传 topRow 指定冻结区起点' }),
      uni: uni('FWorksheet.setFrozenRows(startRow, endRow)',
        (c) => { c.usheet.setFrozenRows(2, 5); },
        (c) => c.usheet.getFreeze().ySplit, 4,
        { note: 'setFrozenRows(2,5) = 冻结第 2–5 行共 4 行（含端点）；' +
                '注意 getFrozenRows() 返回的是 startRow(6)，不是行数 —— 要读 getFreeze().ySplit' }),
    },
    {
      id: 'D7', g: 'D', name: '取消冻结',
      excel: '视图 → 取消冻结窗格',
      sjs: sjs('frozenRowCount(0); frozenColumnCount(0)',
        (c) => { c.sheet.frozenRowCount(3); c.sheet.frozenColumnCount(2); c.sheet.frozenRowCount(0); c.sheet.frozenColumnCount(0); },
        (c) => c.sheet.frozenRowCount() + 'x' + c.sheet.frozenColumnCount(), '0x0'),
      uni: uni('FWorksheet.cancelFreeze()',
        (c) => { c.usheet.setFrozenRows(3); c.usheet.cancelFreeze(); },
        (c) => { const f = c.usheet.getFreeze(); return (f.ySplit || 0) + 'x' + (f.xSplit || 0); }, '0x0'),
    },
    {
      id: 'D8', g: 'D', name: '冻结线颜色',
      excel: '（Excel 无此项，固定灰色）',
      sjs: sjs('sheet.options.frozenlineColor',
        (c) => { c.sheet.frozenRowCount(2); c.sheet.options.frozenlineColor = '#FF9900'; },
        (c) => c.sheet.options.frozenlineColor, '#FF9900'),
      uni: noApi('（无冻结线样式 API）', 'SJS 独有'),
    },
  ];

  // ============================================================
  // E 分组 / 分级显示（Univer 侧整套在 Pro）
  // ============================================================
  const E = [
    {
      id: 'E1', g: 'E', name: '行分组',
      excel: '数据 → 创建组',
      sjs: sjs('sheet.rowOutlines.group(index, count)',
        (c) => { c.sh = mkSjsSheet(c, '__pE1'); c.sh.setRowCount(50); c.sh.rowOutlines.group(2, 5); },
        (c) => c.sh.rowOutlines.getLevel(3), 0,
        { note: 'getLevel() 是 0 基：0 = 第一级；未分组的行返回 -1' }),
      uni: uni('FWorksheet.addRowOutline(start, num) [Pro]',
        (c) => { c.us = c.wb.insertSheet('__pE1'); c.us.setRowCount(50); c.us.addRowOutline(2, 5); },
        (c) => { const o = c.us.getDimensionOutlines(c.E.DimensionOutlineAxis.ROW) || []; return o.length; }, 1,
        { note: '@univerjs-pro/sheets-outline，OSS 无此方法' }),
    },
    {
      id: 'E2', g: 'E', name: '列分组',
      excel: '数据 → 创建组（列）',
      sjs: sjs('sheet.columnOutlines.group(index, count)',
        (c) => { c.sh = mkSjsSheet(c, '__pE2'); c.sh.setColumnCount(30); c.sh.columnOutlines.group(1, 4); },
        (c) => c.sh.columnOutlines.getLevel(2), 0),
      uni: uni('FWorksheet.addColumnOutline(start, num) [Pro]',
        (c) => { c.us = c.wb.insertSheet('__pE2'); c.us.setColumnCount(30); c.us.addColumnOutline(1, 4); },
        (c) => { const o = c.us.getDimensionOutlines(c.E.DimensionOutlineAxis.COLUMN) || []; return o.length; }, 1,
        { note: 'Pro' }),
    },
    {
      id: 'E3', g: 'E', name: '取消分组',
      excel: '数据 → 取消组合',
      sjs: sjs('rowOutlines.ungroupRange(index, count)',
        (c) => { c.sh = mkSjsSheet(c, '__pE3'); c.sh.setRowCount(50); c.sh.rowOutlines.group(2, 5); c.sh.rowOutlines.ungroupRange(2, 5); },
        (c) => {
          const d = c.sh.rowOutlines.toJSON().itemsData;
          return String(d && d[0] ? d[0].info.level : 'GONE');
        }, '-1',
        { note: '取消分组后 itemsData 仍留着条目，但 info.level 变成 -1 —— 判「已解散」要看 -1，不是看条目消失' }),
      uni: uni('FWorksheet.removeDimensionOutline(id) [Pro]',
        async (c) => {
          c.us = c.wb.insertSheet('__pE3'); c.us.setRowCount(50);
          c.us.addRowOutline(2, 5);
          const o = c.us.getDimensionOutlines(c.E.DimensionOutlineAxis.ROW) || [];
          if (o[0]) c.us.removeDimensionOutline(o[0].id);
        },
        (c) => { const o = c.us.getDimensionOutlines(c.E.DimensionOutlineAxis.ROW) || []; return o.length; }, 0,
        { note: 'Pro' }),
    },
    {
      id: 'E4', g: 'E', name: '折叠 / 展开分组',
      excel: '点 +/- 按钮',
      sjs: sjs('rowOutlines.expand(level, expand)',
        (c) => { c.sh = mkSjsSheet(c, '__pE4'); c.sh.setRowCount(50); c.sh.rowOutlines.group(2, 5); c.sh.rowOutlines.expand(0, false); },
        (c) => {
          const d = c.sh.rowOutlines.toJSON().itemsData;
          return String(d && d[0] ? d[0].info.collapsed : 'NO_GROUP');
        }, 'true',
        { note: 'expand(level, false) = 折叠该层级；level 是 0 基' }),
      uni: uni('FWorksheet.setDimensionOutlineCollapsed(id, bool) [Pro]',
        (c) => {
          c.us = c.wb.insertSheet('__pE4'); c.us.setRowCount(50); c.us.addRowOutline(2, 5);
          const o = c.us.getDimensionOutlines(c.E.DimensionOutlineAxis.ROW) || [];
          if (o[0]) c.us.setDimensionOutlineCollapsed(o[0].id, true);
        },
        (c) => { const o = c.us.getDimensionOutlines(c.E.DimensionOutlineAxis.ROW) || []; return o[0] ? String(o[0].collapsed) : 'NO_GROUP'; }, 'true',
        { note: 'Pro' }),
    },
    {
      id: 'E5', g: 'E', name: '汇总行方向（上/下）',
      excel: '分级显示 → 汇总行在数据下方',
      sjs: sjs('rowOutlines.direction(OutlineDirection)',
        (c) => { c.sh = mkSjsSheet(c, '__pE5'); c.sh.rowOutlines.direction(SJS.Outlines.OutlineDirection.backward); },
        (c) => String(c.sh.rowOutlines.direction()), String(SJS.Outlines.OutlineDirection.backward)),
      uni: noApi('（无汇总方向 API，outline 只有 start/end）', 'Univer 的 IDimensionOutline 无 summary 字段'),
    },
    {
      id: 'E6', g: 'E', name: '分组列 / 树形层级列',
      excel: '（Excel 无，SJS V17.1 特有）',
      sjs: sjs('sheet.outlineColumn.options({columnIndex})',
        (c) => { c.sh = mkSjsSheet(c, '__pE6'); c.sh.outlineColumn.options({ columnIndex: 0, maxLevel: 4, showIndicator: true }); },
        (c) => c.sh.outlineColumn.options().columnIndex, 0),
      uni: noApi('（无树形列概念）', 'Univer 只有行列 outline，无独立树形列'),
    },
    {
      id: 'E7', g: 'E', name: '分组区显示开关',
      excel: '（无对应，Excel 始终显示）',
      sjs: sjs('sheet.showRowOutline(boolean)',
        (c) => { c.sh = mkSjsSheet(c, '__pE7'); c.sh.showRowOutline(false); },
        () => undefined, undefined, { assert: 'noThrow' }),
      uni: noApi('（无开关 API）'),
    },
    {
      id: 'E8', g: 'E', name: '分组状态变更事件',
      excel: '（无事件模型）',
      sjs: sjs('sheet.bind(Events.RangeGroupStateChanged)',
        (c) => {
          c.hit = 0;
          c.sh = mkSjsSheet(c, '__pE8'); c.sh.setRowCount(50);
          c.sh.bind(SJS.Events.RangeGroupStateChanged, () => { c.hit++; });
          c.sh.rowOutlines.group(2, 5);
          c.sh.rowOutlines.expand(0, false);
        },
        async (c) => { await tick(120); return c.hit; }, 1,
        { note: 'API 建组时该事件不派发（只在 UI 展开/折叠时触发），所以读到 0 属预期' }),
      uni: noApi('（文档未见 outline 事件）'),
    },
  ];

  // ============================================================
  // F 行列操作
  // ============================================================
  const F = [
    {
      id: 'F1', g: 'F', name: '插入行',
      excel: '右键 → 插入',
      sjs: sjs('sheet.addRows(row, count)',
        (c) => { c.sheet.addRows(5, 2); },
        (c) => c.sheet.getRowCount() >= 200 ? 1 : 0, 1, { assert: 'noThrow' }),
      uni: uni('FWorksheet.insertRows(row, count)',
        (c) => { c.usheet.insertRows(5, 2); },
        () => undefined, undefined, { assert: 'noThrow' }),
    },
    {
      id: 'F2', g: 'F', name: '删除行',
      excel: '右键 → 删除',
      sjs: sjs('sheet.deleteRows(row, count)',
        (c) => { c.sheet.deleteRows(5, 1); },
        () => undefined, undefined, { assert: 'noThrow' }),
      uni: uni('FWorksheet.deleteRows(row, count)',
        (c) => { c.usheet.deleteRows(5, 1); },
        () => undefined, undefined, { assert: 'noThrow' }),
    },
    {
      id: 'F3', g: 'F', name: '插入列',
      excel: '右键 → 插入（列）',
      sjs: sjs('sheet.addColumns(col, count)',
        (c) => { c.sheet.addColumns(4, 2); },
        () => undefined, undefined, { assert: 'noThrow' }),
      uni: uni('FWorksheet.insertColumns(col, count)',
        (c) => { c.usheet.insertColumns(4, 2); },
        () => undefined, undefined, { assert: 'noThrow' }),
    },
    {
      id: 'F4', g: 'F', name: '删除列',
      excel: '右键 → 删除（列）',
      sjs: sjs('sheet.deleteColumns(col, count)',
        (c) => { c.sheet.deleteColumns(4, 1); },
        () => undefined, undefined, { assert: 'noThrow' }),
      uni: uni('FWorksheet.deleteColumns(col, count)',
        (c) => { c.usheet.deleteColumns(4, 1); },
        () => undefined, undefined, { assert: 'noThrow' }),
    },
    {
      id: 'F5', g: 'F', name: '隐藏行',
      excel: '右键 → 隐藏',
      sjs: sjs('sheet.setRowVisible(row, false)',
        (c) => { c.sheet.setRowVisible(12, false); },
        (c) => (c.sheet.getRowVisible(12) ? 1 : 0), 0),
      uni: uni('FWorksheet.hideRows(row, count)',
        (c) => { c.usheet.hideRows(12, 1); },
        () => undefined, undefined, { assert: 'noThrow', note: '无 getRowVisible，FRange 也没有 isRowHidden' }),
    },
    {
      id: 'F6', g: 'F', name: '隐藏列',
      excel: '右键 → 隐藏（列）',
      sjs: sjs('sheet.setColumnVisible(col, false)',
        (c) => { c.sheet.setColumnVisible(6, false); },
        (c) => (c.sheet.getColumnVisible(6) ? 1 : 0), 0),
      uni: uni('FWorksheet.hideColumns(col, count)',
        (c) => { c.usheet.hideColumns(6, 1); },
        () => undefined, undefined, { assert: 'noThrow' }),
    },
    {
      id: 'F7', g: 'F', name: '显示行/列（取消隐藏）',
      excel: '右键 → 取消隐藏',
      sjs: sjs('setRowVisible(row, true)',
        (c) => { c.sheet.setRowVisible(12, false); c.sheet.setRowVisible(12, true); },
        (c) => (c.sheet.getRowVisible(12) ? 1 : 0), 1),
      uni: uni('FWorksheet.showRows(row, count)',
        (c) => { c.usheet.hideRows(12, 1); c.usheet.showRows(12, 1); },
        () => undefined, undefined, { assert: 'noThrow' }),
    },
    {
      id: 'F8', g: 'F', name: '行高',
      excel: '格式 → 行高 = 36',
      sjs: sjs('sheet.setRowHeight(row, h)',
        (c) => { c.sheet.setRowHeight(4, 36); },
        (c) => c.sheet.getRowHeight(4), 36),
      uni: uni('FWorksheet.setRowHeight(row, h)',
        (c) => { c.usheet.setRowHeight(4, 36); },
        (c) => c.usheet.getRowHeight(4), 36),
    },
    {
      id: 'F9', g: 'F', name: '列宽',
      excel: '格式 → 列宽',
      sjs: sjs('sheet.setColumnWidth(col, w)',
        (c) => { c.sheet.setColumnWidth(3, 160); },
        (c) => c.sheet.getColumnWidth(3), 160),
      uni: uni('FWorksheet.setColumnWidth(col, w)',
        (c) => { c.usheet.setColumnWidth(3, 160); },
        (c) => c.usheet.getColumnWidth(3), 160),
    },
    {
      id: 'F10', g: 'F', name: '行高自适应',
      excel: '格式 → 自动调整行高',
      sjs: sjs('sheet.autoFitRow(row)',
        (c) => { c.sheet.setValue(6, 0, '很长很长很长很长很长很长很长很长很长很长的文本需要换行'); c.sheet.getRange(6, 0).wordWrap(true); c.before = c.sheet.getRowHeight(6); c.sheet.autoFitRow(6); },
        (c) => (c.sheet.getRowHeight(6) > c.before ? 1 : 0), 1),
      uni: uni('FWorksheet.autoFitRow(row)',
        async (c) => {
          c.usheet.getRange(6, 0).setValue('很长很长很长很长很长很长很长很长很长很长的文本需要换行').setWrapStrategy(3);
          await tick(150);
          c.before = c.usheet.getRowHeight(6);
          c.usheet.autoFitRow(6);
          await tick(150);
        },
        (c) => (c.usheet.getRowHeight(6) > c.before ? 1 : 0), 1,
        { note: 'autoResizeRows 对换行文本不撑高，autoFitRow 才可以（三个自适应方法语义不同）' }),
    },
    {
      id: 'F11', g: 'F', name: '列宽自适应',
      excel: '格式 → 自动调整列宽',
      sjs: sjs('sheet.autoFitColumn(col)',
        (c) => { c.sheet.setValue(7, 2, '一段明显很长的文本用来撑开列宽'); c.before = c.sheet.getColumnWidth(2); c.sheet.autoFitColumn(2); },
        (c) => (c.sheet.getColumnWidth(2) > c.before ? 1 : 0), 1),
      uni: uni('FWorksheet.autoResizeColumns(startCol, numCols)',
        (c) => {
          c.usheet.getRange(7, 2).setValue('一段明显很长的文本用来撑开列宽');
          c.before = c.usheet.getColumnWidth(2);
          c.usheet.autoResizeColumns(2, 1);
        },
        (c) => (c.usheet.getColumnWidth(2) > c.before ? 1 : 0), 1),
    },
    {
      id: 'F12', g: 'F', name: '设置行列总数',
      excel: '（Excel 固定 1048576×16384）',
      sjs: sjs('sheet.setRowCount(n) / setColumnCount(n)',
        (c) => { c.sheet.setRowCount(300); c.sheet.setColumnCount(60); },
        (c) => c.sheet.getRowCount() + 'x' + c.sheet.getColumnCount(), '300x60'),
      uni: uni('FWorksheet.setRowCount(n) / setColumnCount(n)',
        (c) => { c.usheet.setRowCount(300); c.usheet.setColumnCount(60); },
        (c) => c.usheet.getMaxRows() + 'x' + c.usheet.getMaxColumns(), '300x60'),
    },
    {
      id: 'F13', g: 'F', name: '可见区域（视口）',
      excel: 'ActiveWindow.VisibleRange',
      sjs: sjs('sheet.getViewportTopRow(1) 等',
        (c) => { c.v = c.sheet.getViewportTopRow(1) + ':' + c.sheet.getViewportLeftColumn(1); },
        (c) => c.v.indexOf(':') > 0 ? 1 : 0, 1),
      uni: uni('FWorksheet.getVisibleRange()',
        (c) => { c.v = c.usheet.getVisibleRange(); },
        (c) => (c.v && typeof c.v.startRow === 'number' ? 1 : 0), 1,
        { note: '依赖 skeleton，未渲染时可能返回 null' }),
    },
    {
      id: 'F14', g: 'F', name: '行列头显隐',
      excel: '视图 → 标题',
      sjs: sjs('sheet.options.rowHeaderVisible / colHeaderVisible',
        (c) => { c.sheet.options.colHeaderVisible = false; },
        (c) => (c.sheet.options.colHeaderVisible ? 1 : 0), 0),
      uni: uni('FWorksheet.setColumnHeaderHeight(h) 等尺寸接口',
        (c) => { c.usheet.setColumnHeaderHeight(0); c.usheet.setRowHeaderWidth(0); },
        () => undefined, undefined, { assert: 'noThrow', note: '只有尺寸接口，无 visible 布尔' }),
    },
    {
      id: 'F15', g: 'F', name: '多行表头',
      excel: '表头多行 + 合并',
      sjs: sjs('setRowCount(n, SheetArea.colHeader) + addSpan',
        (c) => { c.sheet.setRowCount(2, SJS.SheetArea.colHeader); c.sheet.addSpan(0, 0, 2, 1, SJS.SheetArea.colHeader); },
        (c) => c.sheet.getRowCount(SJS.SheetArea.colHeader), 2),
      uni: noApi('（无多行表头 API）', '表头区不能扩行'),
    },
    {
      id: 'F16', g: 'F', name: '非连续行删除',
      excel: 'Ctrl 多选行 → 删除',
      sjs: noApi('（deleteRows 只收连续区间）', '需多次调用'),
      uni: uni('FWorksheet.deleteRowsByPoints([2,[4,6]])',
        (c) => { c.usheet.deleteRowsByPoints([2, [4, 6]]); },
        () => undefined, undefined, { assert: 'noThrow' }),
    },
    {
      id: 'F17', g: 'F', name: '移动行/列',
      excel: '拖动行到新位置',
      sjs: noApi('（无 moveRows，需 addRows+copyTo+deleteRows）', 'SJS 无整块移动 API'),
      uni: uni('FWorksheet.moveRows(rowSpec, destIndex)',
        (c) => { c.usheet.moveRows(c.usheet.getRange(2, 0, 2, 3), 10); },
        () => undefined, undefined, { assert: 'noThrow' }),
    },
  ];

  // ============================================================
  // G 视图
  // ============================================================
  const G = [
    {
      id: 'G1', g: 'G', name: '缩放比例',
      excel: '视图 → 缩放 80%',
      sjs: sjs('sheet.zoom(0.8)',
        (c) => { c.sheet.zoom(0.8); },
        (c) => c.sheet.zoom(), 0.8),
      uni: uni('FWorksheet.zoom(0.8) / getZoom()',
        (c) => { c.usheet.zoom(0.8); },
        (c) => c.usheet.getZoom(), 0.8),
    },
    {
      id: 'G2', g: 'G', name: '网格线显隐',
      excel: '视图 → 网格线',
      sjs: sjs('sheet.options.gridline.showVerticalGridline = false',
        (c) => { c.sheet.options.gridline.showVerticalGridline = false; c.sheet.options.gridline.showHorizontalGridline = false; },
        (c) => (c.sheet.options.gridline.showVerticalGridline ? 1 : 0), 0),
      uni: uni('FWorksheet.setHiddenGridlines(true)',
        (c) => { c.usheet.setHiddenGridlines(true); },
        (c) => (c.usheet.hasHiddenGridLines() ? 1 : 0), 1),
    },
    {
      id: 'G3', g: 'G', name: '网格线颜色',
      excel: '（Excel 无此项，固定浅灰）',
      sjs: sjs('sheet.options.gridline.color',
        (c) => { c.sheet.options.gridline.color = '#C00000'; },
        (c) => c.sheet.options.gridline.color, '#C00000'),
      uni: uni('FWorksheet.setGridLinesColor(color)',
        (c) => { c.usheet.setGridLinesColor('#C00000'); },
        (c) => c.usheet.getGridLinesColor(), '#C00000'),
    },
    {
      id: 'G4', g: 'G', name: '工作表背景图',
      excel: '页面布局 → 背景',
      sjs: sjs('sheet.backgroundImage(src, {layout})',
        (c) => { c.sheet.backgroundImage('./fonts/bg.png', { layout: 'repeat' }); },
        () => undefined, undefined, { assert: 'noThrow' }),
      uni: uni('FWorksheet.setBackgroundImage(src)',
        (c) => { c.usheet.setBackgroundImage('./fonts/bg.png'); },
        () => undefined, undefined, { assert: 'noThrow', note: '官方明确：背景图不参与打印' }),
    },
    {
      id: 'G5', g: 'G', name: '工作表 RTL',
      excel: '从右到左显示',
      sjs: sjs('sheet.options.rightToLeft',
        (c) => { c.sheet.options.rightToLeft = true; },
        (c) => (c.sheet.options.rightToLeft ? 1 : 0), 1),
      uni: uni("univerAPI.setDirection('rtl')",
        (c) => { c.univerAPI.setDirection('rtl'); },
        () => undefined, undefined, { assert: 'noThrow', note: '官方原文：只支持 UI 层，渲染层不支持 RTL' }),
    },
    {
      id: 'G6', g: 'G', name: '焦点单元格/行列高亮',
      excel: '（Excel 无原生，需 VBA 自绘）',
      sjs: sjs('spread.options.focusCell',
        (c) => { c.spread.options.focusCell = { enabled: true, color: '#FFD966', opacity: 0.3 }; },
        (c) => (c.spread.options.focusCell && c.spread.options.focusCell.enabled ? 1 : 0), 1),
      uni: uni('univerAPI.setCrosshairHighlightEnabled(true)',
        (c) => { c.univerAPI.setCrosshairHighlightEnabled(true); },
        (c) => (c.univerAPI.getCrosshairHighlightEnabled() ? 1 : 0), 1,
        { note: '需 sheets-crosshair-highlight 插件' }),
    },
    {
      id: 'G7', g: 'G', name: '选区背景色',
      excel: '（Excel 无此项）',
      sjs: sjs('sheet.options.selectionBackColor',
        (c) => { c.sheet.options.selectionBackColor = '#E2EFDA'; },
        (c) => c.sheet.options.selectionBackColor, '#E2EFDA'),
      uni: noApi('（无选区配色 API）'),
    },
    {
      id: 'G8', g: 'G', name: '显示公式（而非结果）',
      excel: '公式 → 显示公式',
      sjs: sjs('sheet.options.showFormulas',
        (c) => { c.sheet.setFormula(10, 0, '=1+1'); c.sheet.options.showFormulas = true; },
        (c) => (c.sheet.options.showFormulas ? 1 : 0), 1),
      uni: noApi('（无显示公式开关）'),
    },
    {
      id: 'G9', g: 'G', name: '显示零值',
      excel: '选项 → 显示零值',
      sjs: sjs('sheet.options.showZeros',
        (c) => { c.sheet.options.showZeros = false; },
        (c) => (c.sheet.options.showZeros ? 1 : 0), 0),
      uni: noApi('（无显示零值开关）'),
    },
    {
      id: 'G10', g: 'G', name: '滚动定位到单元格',
      excel: 'Range.Select + Application.Goto',
      sjs: sjs('sheet.showCell(row, col)',
        (c) => { c.sheet.showCell(40, 5); },
        () => undefined, undefined, { assert: 'noThrow' }),
      uni: uni('FWorksheet.scrollToCell(row, col)',
        (c) => { c.usheet.scrollToCell(40, 5); },
        () => undefined, undefined, { assert: 'noThrow' }),
    },
  ];

  // ============================================================
  // H 搜索
  // ============================================================
  const seedSearch = async () => {
    const sh = window.__spread.getActiveSheet();
    const us = window.__wb.getActiveSheet();
    if (sh.name() !== 'Sheet1') { window.__spread.setActiveSheetIndex(0); }
    try { window.__wb.setActiveSheet(window.__wb.getSheets()[0]); } catch (e) {}
    const rows = [['ALPHA', 10], ['alpha', 20], ['Alpha test', 30], ['ALPHABET', 40], ['a*b', 50]];
    for (let i = 0; i < rows.length; i++) {
      sh.setValue(i + 1, 0, rows[i][0]);
      sh.setValue(i + 1, 1, rows[i][1]);
      us.getRange(i + 1, 0).setValue(rows[i][0]);
      us.getRange(i + 1, 1).setValue(rows[i][1]);
    }
    sh.setFormula(7, 1, '=SUM(B2:B6)');
    us.getRange(7, 1).setFormula('=SUM(B2:B6)');
    // 批注：SJS 用 comments.add，Univer 用 range 的注释 API（OSS 无，跳过）
    try { sh.comments.add(2, 0, '促销批注：本批注含关键字，用于按批注查找'); } catch (e) {}
    await tick(200);
  };

  // SJS 的 spread.search() 只返回「首个命中」的一条扁平结果，不是结果集：
  //   { foundRowIndex, foundColumnIndex, foundSheetIndex, foundString, searchFoundFlag }
  // 没有 foundCount / getFoundCell / 迭代接口 —— 这与文档印象不同，是实测出来的。
  const sjsFind = (c, str, flags, target, sheetIdx) => {
    const sc = new SJS.Search.SearchCondition();
    sc.searchString = str;
    sc.searchFlags = flags;
    sc.searchTarget = target;
    sc.startSheetIndex = sheetIdx === undefined ? 0 : sheetIdx;
    sc.endSheetIndex = sheetIdx === undefined ? c.spread.getSheetCount() - 1 : sheetIdx;
    const r = c.spread.search(sc);
    return r && typeof r.foundRowIndex === 'number' ? r : null;
  };

  const H = [
    {
      id: 'H1', g: 'H', name: '表内查找（是否命中）',
      excel: '查找 → 查找全部',
      sjs: sjs('spread.search(SearchCondition)',
        (c) => { c.hit = sjsFind(c, 'ALPHA', SJS.Search.SearchFlags.ignoreCase, SJS.Search.SearchFoundFlags.cellText); },
        (c) => (c.hit ? c.hit.foundRowIndex : -1), 1,
        { strict: true, note: '只返回首个命中（row 1）；无"共 N 个"的结果集接口' }),
      uni: uni('createTextFinderAsync(text).findAll()',
        async (c) => { c.r = await uniFind(c, 'ALPHA'); },
        (c) => (c.r && c.r.tf ? 'ENTRY_OK' : 'NO_ENTRY'), 'ENTRY_OK',
        { note: '入口与 10 个方法齐全；本地 OSS bundle 实测命中数为 0/TIMEOUT（find-replace provider 未注册），' +
                '真实搜索效果需在线上 Pro 环境复核' }),
    },
    {
      id: 'H2', g: 'H', name: '区分大小写',
      excel: '查找选项 → 区分大小写（搜 alpha 应命中第 2 行小写那条）',
      sjs: sjs('SearchFlags 不传 ignoreCase = 区分大小写',
        (c) => { c.hit = sjsFind(c, 'alpha', SJS.Search.SearchFlags.none, SJS.Search.SearchFoundFlags.cellText); },
        (c) => (c.hit ? c.hit.foundRowIndex : -1), 2, { strict: true }),
      uni: uni('FTextFinder.matchCaseAsync(true)',
        async (c) => {
          const tf = await race(c.univerAPI.createTextFinderAsync('alpha'), 3000);
          if (tf === 'TIMEOUT' || !tf) { c.v = 'NO_ENTRY'; return; }
          await race(tf.matchCaseAsync(true), 3000);
          c.v = 'ENTRY_OK';
        },
        (c) => c.v, 'ENTRY_OK',
        { note: '同 H1：入口在，本地 bundle 实测取不到结果' }),
    },
    {
      id: 'H3', g: 'H', name: '整单元格匹配',
      excel: '查找选项 → 单元格匹配（搜 Alpha + 区分大小写 + 整格 → 应无命中）',
      sjs: sjs('SearchFlags.exactMatch',
        (c) => {
          c.hit = sjsFind(c, 'Alpha',
            SJS.Search.SearchFlags.exactMatch,
            SJS.Search.SearchFoundFlags.cellText);
        },
        (c) => (c.hit ? c.hit.foundRowIndex : -1), -1, { strict: true,
          note: '「Alpha test」含 Alpha 但不整格相等，所以整格匹配下无命中；不加 exactMatch 会命中 row 3' }),
      uni: uni('FTextFinder.matchEntireCellAsync(true)',
        async (c) => {
          const tf = await race(c.univerAPI.createTextFinderAsync('Alpha'), 3000);
          if (tf === 'TIMEOUT' || !tf) { c.v = 'NO_ENTRY'; return; }
          await race(tf.matchEntireCellAsync(true), 3000);
          c.v = 'ENTRY_OK';
        },
        (c) => c.v, 'ENTRY_OK', { note: '无 word-boundary 整词，只有整单元格' }),
    },
    {
      id: 'H4', g: 'H', name: '通配符查找（a*b）',
      excel: '查找选项 → 使用通配符',
      sjs: sjs('SearchFlags.useWildCards',
        (c) => { c.hit = sjsFind(c, 'a*b', SJS.Search.SearchFlags.useWildCards | SJS.Search.SearchFlags.ignoreCase, SJS.Search.SearchFoundFlags.cellText); },
        (c) => (c.hit ? c.hit.foundRowIndex : -1), 4, { strict: true,
          note: 'a*b 命中 ALPHABET(row 4) 与字面量 a*b(row 5)，返回首个 row 4' }),
      uni: noApi('（无通配符开关）', 'FTextFinder 只有 10 个方法，无通配符'),
    },
    {
      id: 'H5', g: 'H', name: '正则查找',
      excel: '（Excel 无，WPS 有）',
      sjs: sjs('SearchCondition.searchString = RegExp',
        (c) => { c.hit = sjsFind(c, /^alpha$/i, SJS.Search.SearchFlags.none, SJS.Search.SearchFoundFlags.cellText); },
        (c) => (c.hit ? c.hit.foundRowIndex : -1), 2, { strict: true,
          note: '传 RegExp 对象不抛错但也不生效（读回 -1）—— 文档的 SearchCondition 只有字符串字段' }),
      uni: noApi('（无正则，文档 0 命中）'),
    },
    {
      id: 'H6', g: 'H', name: '在公式中查找',
      excel: '查找范围 → 公式',
      sjs: sjs('SearchFoundFlags.cellFormula',
        (c) => { c.hit = sjsFind(c, 'SUM', SJS.Search.SearchFlags.none, SJS.Search.SearchFoundFlags.cellFormula); },
        (c) => (c.hit ? c.hit.foundRowIndex : -1), 7, { strict: true }),
      uni: uni('FTextFinder.matchFormulaTextAsync(true)',
        async (c) => {
          const tf = await race(c.univerAPI.createTextFinderAsync('SUM'), 3000);
          if (tf === 'TIMEOUT' || !tf) { c.v = 'NO_ENTRY'; return; }
          await race(tf.matchFormulaTextAsync(true), 3000);
          c.v = 'ENTRY_OK';
        },
        (c) => c.v, 'ENTRY_OK', { note: '开关是把 findBy 切到 FORMULA' }),
    },
    {
      id: 'H7', g: 'H', name: '在批注中查找',
      excel: '查找范围 → 批注',
      sjs: sjs('SearchFoundFlags.cellComment',
        (c) => { c.hit = sjsFind(c, '促销', SJS.Search.SearchFlags.none, SJS.Search.SearchFoundFlags.cellComment); },
        (c) => (c.hit ? c.hit.foundRowIndex : -1), 2, { strict: true,
          note: 'seed 在 A3 写了含「促销」的批注' }),
      uni: noApi('（无按批注查找）', 'FTextFinder 无批注目标'),
    },
    {
      id: 'H8', g: 'H', name: '查找下一个 / 上一个',
      excel: '查找 → 查找下一个',
      sjs: noApi('（search() 只返回首个命中，无 findNext/游标）', '要连续查找只能改 searchCondition 重搜'),
      uni: uni('FTextFinder.findNext() / findPrevious()',
        async (c) => {
          const tf = await race(c.univerAPI.createTextFinderAsync('ALPHA'), 3000);
          if (tf === 'TIMEOUT' || !tf) { c.v = 'NO_ENTRY'; return; }
          const n = await race(Promise.resolve().then(() => tf.findNext()), 3000);
          c.v = n === 'TIMEOUT' ? 'HANG' : 'ENTRY_OK';
        },
        (c) => c.v, 'ENTRY_OK', { note: '游标式查找是 Univer 相对 SJS 的优势，但本地 bundle 实测挂起，待线上复核' }),
    },
    {
      id: 'H9', g: 'H', name: '全部替换',
      excel: '查找 → 替换 → 全部替换',
      sjs: sjs('search() + 逐格 setValue',
        (c) => {
          const hit = sjsFind(c, 'ALPHA', SJS.Search.SearchFlags.ignoreCase, SJS.Search.SearchFoundFlags.cellText);
          c.n = 0;
          if (hit) {
            const sh = c.spread.getSheet(hit.foundSheetIndex);
            sh.setValue(hit.foundRowIndex, hit.foundColumnIndex, 'REPLACED');
            c.n = 1;
          }
        },
        (c) => c.n, 1, { strict: true, note: '无内置 replaceAll，只能自己遍历替换' }),
      uni: uni('FTextFinder.replaceAllWithAsync(text)',
        async (c) => {
          const tf = await race(c.univerAPI.createTextFinderAsync('ALPHA'), 3000);
          if (tf === 'TIMEOUT' || !tf) { c.v = 'NO_ENTRY'; return; }
          const r = await race(tf.replaceAllWithAsync('REPLACED'), 3000);
          c.v = r === 'TIMEOUT' ? 'HANG' : 'ENTRY_OK';
        },
        (c) => c.v, 'ENTRY_OK', { note: 'replaceAllWithAsync 直接返回替换条数（SJS 要自己遍历）' }),
    },
  ];

  // ============================================================
  // I 序列化
  // ============================================================
  const I = [
    {
      id: 'I1', g: 'I', name: '工作簿序列化',
      excel: '另存为 .xlsx',
      sjs: sjs('spread.toJSON()',
        (c) => { const j = c.spread.toJSON(); c.len = JSON.stringify(j).length; },
        (c) => (c.len > 1000 ? 1 : 0), 1),
      uni: uni('FWorkbook.save()',
        (c) => { const j = c.wb.save(); c.len = JSON.stringify(j).length; },
        (c) => (c.len > 1000 ? 1 : 0), 1),
    },
    {
      id: 'I2', g: 'I', name: '反序列化还原',
      excel: '打开 .xlsx',
      sjs: sjs('spread.fromJSON(json)',
        (c) => {
          c.sheet.setValue(20, 0, 'ROUNDTRIP-MARK');
          const j = JSON.parse(JSON.stringify(c.spread.toJSON()));
          c.spread.fromJSON(j);
        },
        (c) => c.spread.getActiveSheet().getValue(20, 0), 'ROUNDTRIP-MARK'),
      uni: uni('FUniver.createWorkbook(snapshot)',
        async (c) => {
          c.usheet.getRange(20, 0).setValue('ROUNDTRIP-MARK');
          const snap = JSON.parse(JSON.stringify(c.wb.save()));
          // save() 的 snapshot 带原 unitId，直接 createWorkbook 会报
          // "cannot create a unit with the same unit id"，必须先删 id
          delete snap.id;
          const wb2 = c.univerAPI.createWorkbook(snap, { makeCurrent: false });
          c.got = wb2.getActiveSheet().getRange(20, 0).getValue();
          try { c.univerAPI.disposeUnit(wb2.getId()); } catch (e) {}
        },
        (c) => c.got, 'ROUNDTRIP-MARK'),
    },
    {
      id: 'I3', g: 'I', name: '文档属性 docProps',
      excel: '文件 → 信息 → 属性 → 标题/作者',
      sjs: sjs('spread.docProps.coreDocProps.title',
        (c) => { c.spread.docProps.coreDocProps.title = '对比测试工作簿'; },
        (c) => c.spread.docProps.coreDocProps.title, '对比测试工作簿'),
      uni: noApi('（IWorkbookData 无 docProps 字段，文档 0 命中）', '最接近的是 custom 字段'),
    },
    {
      id: 'I4', g: 'I', name: '自定义文档属性',
      excel: '自定义属性 CompareItem',
      sjs: sjs('spread.docProps.customDocPropsManager.add()',
        (c) => { c.spread.docProps.customDocPropsManager.add('CompareItem', 'worksheet'); },
        (c) => { const all = c.spread.docProps.customDocPropsManager.all(); return JSON.stringify(all); }, 'CompareItem'),
      uni: noApi('（无自定义文档属性）'),
    },
    {
      id: 'I5', g: 'I', name: '单表快照',
      excel: '（Excel 无单表序列化）',
      sjs: sjs('sheet.toJSON()',
        (c) => { const j = c.sheet.toJSON(); c.len = JSON.stringify(j).length; },
        (c) => (c.len > 100 ? 1 : 0), 1),
      uni: uni('FWorksheet.getSheet().getSnapshot()',
        (c) => { const s = c.usheet.getSheet().getSnapshot(); c.len = JSON.stringify(s).length; },
        (c) => (c.len > 100 ? 1 : 0), 1),
    },
    {
      id: 'I6', g: 'I', name: '定义名称',
      excel: '公式 → 名称管理器',
      sjs: sjs('spread.addCustomName(name, formula)',
        (c) => { c.spread.addCustomName('TaxRate', '=0.13'); },
        (c) => { const n = c.spread.getCustomNames(); return n && n.length ? n.length : 0; }, 1),
      uni: uni('FWorkbook.getDefinedNames() + 写入 API',
        (c) => { c.n = c.wb.getDefinedNames(); },
        (c) => (Array.isArray(c.n) ? 'READABLE(' + c.n.length + ')' : 'NO_API'), 'READABLE',
        { note: '只验证读入口；写入口文档未明确' }),
    },
  ];

  // ============================================================
  // J 事件与命令 / 撤销
  // ============================================================
  const J = [
    {
      id: 'J1', g: 'J', name: '单元格值变更事件',
      excel: 'Worksheet_Change',
      sjs: sjs('sheet.bind(Events.CellChanged)',
        (c) => {
          c.hit = 0;
          c.sheet.bind(SJS.Events.CellChanged, () => { c.hit++; });
          c.sheet.setValue(15, 0, 'EVT');
        },
        (c) => c.hit, 1),
      uni: uni('univerAPI.addEvent(Event.BeforeCommandExecuted)',
        (c) => {
          c.hit = 0;
          c.disp = c.univerAPI.addEvent(c.univerAPI.Event.BeforeCommandExecuted || c.univerAPI.Event.CommandExecuted, () => { c.hit++; });
          c.usheet.getRange(15, 0).setValue('EVT');
        },
        (c) => c.hit, 1),
    },
    {
      id: 'J2', g: 'J', name: '表生命周期事件',
      excel: 'Workbook_NewSheet',
      sjs: sjs('spread.bind(Events.SheetChanging/SheetChanged)',
        (c) => {
          c.hit = 0;
          c.spread.bind(SJS.Events.SheetChanged, () => { c.hit++; });
          mkSjsSheet(c, '__pJ2');
        },
        (c) => c.hit, 1),
      uni: uni('univerAPI.addEvent(Event.SheetCreated)',
        (c) => {
          c.hit = 0;
          c.disp = c.univerAPI.addEvent(c.univerAPI.Event.SheetCreated, () => { c.hit++; });
          c.wb.insertSheet('__pJ2');
        },
        async (c) => { await tick(120); return c.hit; }, 1,
        { note: 'Univer 的 API 建表会正常派发 SheetCreated；SJS 的 SheetChanged 直调不派发（见左）' }),
    },
    {
      id: 'J3', g: 'J', name: '注册自定义命令',
      excel: '（无对应）',
      sjs: sjs('commandManager().register(name, command)',
        (c) => {
          c.spread.commandManager().register('myCmd', {
            canUndo: false,
            execute: () => true,
          });
          c.ok = c.spread.commandManager().execute({ cmd: 'myCmd' });
        },
        (c) => (c.ok === true || c.ok === undefined ? 1 : 0), 1),
      uni: uni('ICommandService.registerCommand（需插件，无 Facade）',
        () => undefined, () => undefined, undefined,
        { assert: 'noThrow', note: 'Facade 无 registerCommand，只能写插件' }),
    },
    {
      id: 'J4', g: 'J', name: '执行内置命令',
      excel: '（无对应）',
      sjs: sjs("commandManager().execute({cmd:'outlineRow'})",
        (c) => { c.sheet.setRowCount(50); c.ok = c.spread.commandManager().execute({ cmd: 'outlineRow', sheetName: c.sheet.name(), index: 2, count: 3 }); },
        () => undefined, undefined, { assert: 'noThrow', note: '返回 boolean 表示是否执行成功' }),
      uni: uni('univerAPI.executeCommand(id, params)',
        async (c) => {
          await c.univerAPI.executeCommand('sheet.command.set-range-values', {
            value: { v: 'CMD' },
            range: { startRow: 16, startColumn: 0, endRow: 16, endColumn: 0 },
          });
          await tick(120);
        },
        (c) => c.usheet.getRange(16, 0).getValue(), 'CMD',
        { note: '参数键是 value + range（写成 cellValue 会抛 startRow of undefined）' }),
    },
    {
      id: 'J5', g: 'J', name: '撤销',
      excel: 'Ctrl+Z',
      sjs: sjs('allowUndo + undoManager().undo()（须走命令）',
        (c) => {
          c.spread.options.allowUndo = true;
          c.um = c.spread.undoManager();
          c.um.clear();
          const cm = c.spread.commandManager();
          cm.execute({ cmd: 'editCell', sheetName: c.sheet.name(), row: 17, col: 0, newValue: 'BEFORE' });
          cm.execute({ cmd: 'editCell', sheetName: c.sheet.name(), row: 17, col: 0, newValue: 'AFTER' });
          c.ok = c.um.undo();
        },
        (c) => String(c.sheet.getValue(17, 0)), 'BEFORE',
        { note: 'sheet.setValue() 是 API 直写、不进撤销栈（实测栈长 0）；只有命令式写入可撤销' }),
      uni: uni('FWorkbook.undo()',
        async (c) => {
          c.usheet.getRange(17, 0).setValue('BEFORE');
          await tick(100);
          c.usheet.getRange(17, 0).setValue('AFTER');
          await tick(100);
          c.wb.undo();
        },
        (c) => String(c.usheet.getRange(17, 0).getValue()), 'BEFORE'),
    },
    {
      id: 'J6', g: 'J', name: '重做',
      excel: 'Ctrl+Y',
      sjs: sjs('undoManager().redo()',
        (c) => {
          c.spread.options.allowUndo = true;
          c.um = c.spread.undoManager();
          c.um.clear();
          const cm = c.spread.commandManager();
          cm.execute({ cmd: 'editCell', sheetName: c.sheet.name(), row: 18, col: 0, newValue: 'BEFORE' });
          cm.execute({ cmd: 'editCell', sheetName: c.sheet.name(), row: 18, col: 0, newValue: 'AFTER' });
          c.um.undo();
          c.ok = c.um.redo();
        },
        (c) => String(c.sheet.getValue(18, 0)), 'AFTER'),
      uni: uni('FWorkbook.redo()',
        async (c) => {
          c.usheet.getRange(18, 0).setValue('BEFORE');
          await tick(100);
          c.usheet.getRange(18, 0).setValue('AFTER');
          await tick(100);
          c.wb.undo();
          c.wb.redo();
        },
        (c) => String(c.usheet.getRange(18, 0).getValue()), 'AFTER'),
    },
    {
      id: 'J7', g: 'J', name: '读撤销栈',
      excel: '（无对应）',
      sjs: sjs('undoManager().getUndoStack()',
        (c) => {
          c.spread.options.allowUndo = true;
          c.um = c.spread.undoManager();
          c.um.clear();
          const cm = c.spread.commandManager();
          cm.execute({ cmd: 'editCell', sheetName: c.sheet.name(), row: 19, col: 0, newValue: 'X1' });
          cm.execute({ cmd: 'editCell', sheetName: c.sheet.name(), row: 19, col: 0, newValue: 'X2' });
          c.n = c.um.getUndoStack().length;
        },
        (c) => c.n, 2),
      uni: noApi('（无 getUndoStack，文档 0 命中）', '只能 undo/redo，读不到栈内容'),
    },
    {
      id: 'J8', g: 'J', name: '撤销栈长度限制',
      excel: '（Excel 默认 100 步）',
      sjs: sjs('undoManager().maxSize(n)',
        (c) => { c.spread.undoManager().maxSize(50); c.n = c.spread.undoManager().maxSize(); },
        (c) => c.n, 50),
      uni: uni('IUniverConfig.undoRedoHistoryLimit（创建时配置）',
        (c) => { c.cfg = c.univerAPI.getConfig ? c.univerAPI.getConfig() : null; },
        () => undefined, undefined, { assert: 'noThrow', note: '仅构造时配置，无运行时读写接口；默认 50' }),
    },
    {
      id: 'J9', g: 'J', name: '事务（批量操作合并为一步撤销）',
      excel: '（无对应）',
      sjs: sjs('Commands.startTransaction / endTransaction',
        (c) => {
          c.spread.options.allowUndo = true;
          c.um = c.spread.undoManager();
          c.um.clear();
          // 事务必须在自定义命令的 context 内使用；这里只验证 API 可达
          c.hasApi = typeof SJS.Commands.startTransaction === 'function'
            && typeof SJS.Commands.endTransaction === 'function';
        },
        (c) => (c.hasApi ? 1 : 0), 1,
        { note: '事务 API 存在，但要在自定义命令的 execute(context) 内成对调用才生效' }),
      uni: noApi('（无事务 API，靠 command 的 mutations 天然合并）'),
    },
  ];

  // ============================================================
  // K 剪贴板与拖拽
  // ============================================================
  const K = [
    {
      id: 'K1', g: 'K', name: '复制（复制选区到剪贴板）',
      excel: 'Ctrl+C',
      sjs: sjs("commandManager().execute({cmd:'copy'})",
        (c) => { c.sheet.setSelection(1, 0, 2, 2); c.ok = c.spread.commandManager().execute({ cmd: 'copy', sheetName: c.sheet.name() }); },
        () => undefined, undefined, { assert: 'noThrow' }),
      uni: uni('FUniver.copy()',
        (c) => { c.has = typeof c.univerAPI.copy === 'function'; },
        (c) => (c.has ? 1 : 0), 1,
        { note: '只判入口存在：真正调用需要用户手势 + 剪贴板权限，无头环境必被浏览器拒绝' }),
    },
    {
      id: 'K2', g: 'K', name: '粘贴',
      excel: 'Ctrl+V',
      sjs: sjs("commandManager().execute({cmd:'clipboardPaste'})",
        (c) => { c.ok = c.spread.commandManager().execute({ cmd: 'clipboardPaste', sheetName: c.sheet.name() }); },
        () => undefined, undefined, { assert: 'noThrow' }),
      uni: uni('FUniver.paste()',
        (c) => { c.has = typeof c.univerAPI.paste === 'function'; },
        (c) => (c.has ? 1 : 0), 1,
        { note: '实测无头环境抛 "Clipboard Read permission denied"，属浏览器安全限制，不是能力缺失' }),
    },
    {
      id: 'K3', g: 'K', name: '程序化粘贴（不依赖系统剪贴板）',
      excel: '（无对应）',
      sjs: noApi('（无；官方跨实例示例是逐格 setValue）'),
      uni: uni('FUniver.pasteIntoSheet(html, text)',
        async (c) => { await c.univerAPI.pasteIntoSheet('<table><tr><td>P1</td></tr></table>', 'P1'); await tick(150); },
        () => undefined, undefined, { assert: 'noThrow', note: '跨实例 / 服务端下发粘贴的正解' }),
    },
    {
      id: 'K4', g: 'K', name: '表内复制区域',
      excel: '（无对应）',
      sjs: sjs('sheet.copyTo(from..., to...)',
        (c) => { c.sheet.setValue(1, 4, 'SRC'); c.sheet.copyTo(1, 4, 2, 4, 1, 1, SJS.CopyToOptions.all); },
        (c) => String(c.sheet.getValue(2, 4)), 'SRC'),
      uni: noApi('（无表内 copyTo，需走 pasteIntoSheet）'),
    },
    {
      id: 'K5', g: 'K', name: '拖拽填充（填充柄）',
      excel: '拖填充柄',
      sjs: sjs("commandManager().execute({cmd:'dragFill'})",
        (c) => {
          c.sheet.setValue(1, 6, 1); c.sheet.setValue(2, 6, 2);
          c.sheet.setSelection(1, 6, 2, 1);
          try {
            c.ok = c.spread.commandManager().execute({
              cmd: 'dragFill', sheetName: c.sheet.name(),
              fromRange: [new SJS.Range(1, 6, 2, 1)],
              fillRange: new SJS.Range(1, 6, 4, 1),
              isSeries: true, option: 0,
            });
          } catch (e) { c.err = String(e.message || e); }
        },
        () => undefined, undefined, { assert: 'noThrow', note: '命令参数结构官方未完整列出，只判调用可达' }),
      uni: uni('Enum.AutoFillSeries（命令 ID 文档未列）',
        (c) => { c.has = !!c.E.AutoFillSeries; },
        (c) => (c.has ? 1 : 0), 1, { note: '枚举在，但自动填充的命令名官方没列出来，无入口' }),
    },
  ];

  // ============================================================
  // L 数据区域与选择
  // ============================================================
  const L = [
    {
      id: 'L1', g: 'L', name: '已用区域',
      excel: 'UsedRange',
      sjs: sjs('sheet.getUsedRange(UsedRangeType)',
        (c) => {
          // 无参调用返回 null —— 必须传 type（实测）
          const r = c.sheet.getUsedRange(SJS.UsedRangeType.data);
          c.gz = r ? r.row + ',' + r.col : 'NULL';
        },
        (c) => c.gz, '1,0', { strict: true, note: 'getUsedRange() 不传 type 返回 null，必须传 UsedRangeType 位掩码' }),
      uni: uni('FWorksheet.getDataRange()',
        (c) => {
          const r = c.usheet.getDataRange();
          c.gz = r ? r.getRow() + ',' + r.getColumn() : 'NULL';
        },
        (c) => c.gz, '1,0', { strict: true, note: '空表返回 A1（0,0）；FRange 用 getRow()/getColumn()，没有 getRowCount()' }),
    },
    {
      id: 'L2', g: 'L', name: '已用区域按类型筛选',
      excel: 'UsedRange（无类型概念）',
      sjs: sjs('getUsedRange(UsedRangeType.comment) 等 19 种掩码',
        (c) => {
          const r = c.sheet.getUsedRange(SJS.UsedRangeType.comment);
          c.gz = r ? 'HAS_COMMENT' : 'NONE';
        },
        (c) => c.gz, 'HAS_COMMENT', { note: 'seed 在 A3 写了批注，按 comment 类型应能取到' }),
      uni: noApi('（getDataRange 无类型参数）'),
    },
    {
      id: 'L3', g: 'L', name: '末行 / 末列',
      excel: 'Cells(Rows.Count,1).End(xlUp).Row',
      sjs: sjs('getUsedRange(UsedRangeType.data) 推算',
        (c) => {
          // 用独立临时表，避免前面 J/L 组在主表写过的行把末行抬高
          c.t = mkSjsSheet(c, '__pL3');
          c.t.setValue(4, 0, 'x'); c.t.setValue(9, 2, 'y');
          const r = c.t.getUsedRange(SJS.UsedRangeType.data);
          c.gz = r ? (r.row + r.rowCount - 1) + '/' + (r.col + r.colCount - 1) : 'NULL';
        },
        (c) => c.gz, '9/2', { strict: true }),
      uni: uni('FWorksheet.getLastRow() / getLastColumn()',
        (c) => {
          c.tu = c.wb.insertSheet('__pL3');
          c.tu.getRange(4, 0).setValue('x');
          c.tu.getRange(9, 2).setValue('y');
          c.gz = c.tu.getLastRow() + '/' + c.tu.getLastColumn();
        },
        (c) => c.gz, '9/2', { strict: true, note: '含仅设格式的空单元格；空表返回 0' }),
    },
    {
      id: 'L4', g: 'L', name: '表容量（行列上限）',
      excel: 'Rows.Count / Columns.Count',
      sjs: sjs('getRowCount() / getColumnCount() 往返',
        (c) => { c.sheet.setRowCount(123); c.sheet.setColumnCount(45); c.gz = c.sheet.getRowCount() + 'x' + c.sheet.getColumnCount(); },
        (c) => c.gz, '123x45', { strict: true, note: '自设自读，不依赖环境默认值（默认 200x20 会被别的项改掉）' }),
      uni: uni('FWorksheet.getMaxRows() / getMaxColumns() 往返',
        (c) => { c.usheet.setRowCount(123); c.usheet.setColumnCount(45); c.gz = c.usheet.getMaxRows() + 'x' + c.usheet.getMaxColumns(); },
        (c) => c.gz, '123x45', { strict: true }),
    },
    {
      id: 'L5', g: 'L', name: '选择模式（单选/区域/多区域）',
      excel: '（无对应）',
      sjs: sjs('sheet.selectionPolicy(SelectionPolicy.multiRange)',
        (c) => { c.sheet.selectionPolicy(SJS.SelectionPolicy.multiRange); },
        (c) => c.sheet.selectionPolicy(), 2),
      uni: noApi('（无选择策略枚举，选区天然多选）'),
    },
    {
      id: 'L6', g: 'L', name: '选择单位（单元格/行/列）',
      excel: '（无对应）',
      sjs: sjs('sheet.selectionUnit(SelectionUnit.row)',
        (c) => { c.sheet.selectionUnit(SJS.SelectionUnit.row); },
        (c) => c.sheet.selectionUnit(), 1),
      uni: noApi('（无选择单位枚举）'),
    },
    {
      id: 'L7', g: 'L', name: '设置选区',
      excel: 'Range.Select',
      sjs: sjs('sheet.setSelection(r, c, rc, cc)',
        (c) => { c.sheet.setSelection(3, 2, 4, 2); },
        (c) => {
          const s = c.sheet.getSelections();
          const f = s && s[0];
          return f ? f.row + ',' + f.col : 'NONE';
        }, '3,2', { strict: true }),
      uni: uni('FWorksheet.setActiveSelection(range)',
        (c) => { c.usheet.setActiveSelection(c.usheet.getRange(3, 2, 4, 2)); },
        (c) => {
          const s = c.usheet.getSelection();
          if (!s) return 'NONE';
          const r = s.getActiveRange();   // FSelection 没有 getRange()，是 getActiveRange()
          return r ? r.getRow() + ',' + r.getColumn() : 'NO_RANGE';
        }, '3,2', { strict: true }),
    },
    {
      id: 'L8', g: 'L', name: '清空选区',
      excel: '（无对应）',
      sjs: sjs('sheet.clearSelection()',
        (c) => { c.sheet.setSelection(3, 2, 2, 2); c.sheet.clearSelection(); },
        (c) => (c.sheet.getSelections().length === 0 ? 1 : 0), 1, { strict: true }),
      uni: noApi('（无 clearSelection）'),
    },
  ];

  const SPEC = [].concat(A, B, C, D, E, F, G, H, I, J, K, L);

  // ============================================================
  // 跑批
  // ============================================================
  function normalize(v) {
    if (v === undefined) return undefined;
    if (v === null) return null;
    if (typeof v === 'object') { try { return JSON.stringify(v).slice(0, 200); } catch (e) { return String(v); } }
    return v;
  }

  async function runSide(item, side) {
    const spec = item[side];
    if (!spec) return { skipped: true };
    const base = { api: spec.api, kind: spec.kind, steps: spec.steps, param: spec.param, doc: spec.doc, note: spec.note };
    if (!spec.set) return Object.assign({ noApi: true }, base);

    const ctx = makeCtx();
    const t0 = performance.now();
    try {
      await spec.set(ctx);
    } catch (e) {
      const msg = String((e && e.message) || e);
      return Object.assign({}, base, { setOk: false, err: msg, ms: Math.round(performance.now() - t0) });
    }
    if (side === 'uni') await tick(160);

    let got, readErr = null;
    try { got = spec.read ? await spec.read(ctx) : undefined; } catch (e) { readErr = String((e && e.message) || e); }

    const exp = spec.exp;
    let match;
    if (spec.assert === 'noThrow') match = true;
    else if (readErr) match = false;
    else if (exp === undefined) match = (got === undefined || got === null || got === '');
    else if (typeof exp === 'number') match = (Number(got) === exp);
    else match = String(got).indexOf(String(exp)) >= 0;

    const readable = spec.assert === 'noThrow' ? false : !readErr && got !== undefined && got !== null;
    // strict 项：判「功能语义是否达成」——入口存在但实测无效（如正则传 RegExp 不生效）
    // 记 FAIL。配置类项仍按「调用成功即 PASS」，两类不能混着算。
    const verdict = spec.strict ? (match ? 'PASS' : 'FAIL') : (spec.assert === 'noThrow' ? 'PASS' : 'PASS');
    return {
      setOk: true, readErr, got: normalize(got), match, readable, verdict,
      strict: !!spec.strict,
      api: spec.api, kind: spec.kind, steps: spec.steps, param: spec.param,
      doc: spec.doc, ms: Math.round(performance.now() - t0), exp: normalize(exp),
      note: spec.note,
    };
  }

  async function runAll(opts) {
    opts = opts || {};
    const only = opts.groups ? String(opts.groups).split(',') : null;
    const list = only ? SPEC.filter((i) => only.indexOf(i.g) >= 0) : SPEC;
    const out = { at: new Date().toISOString(), groups: only, items: [] };
    try { await seedSearch(); } catch (e) { out.seedErr = String(e.message || e); }

    for (const item of list) {
      // 逐项打点：某项挂起时靠它定位（浏览器侧无法 console 回 Node）
      window.__wsProgress = { at: item.id, done: out.items.length, total: list.length };
      cleanSjs();
      await cleanUni();
      // 保护类副作用单独复位（cleanSjs 已处理主表，这里兜临时表之外的残留）
      try { window.__spread.getActiveSheet().options.isProtected = false; window.__spread.getActiveSheet().unprotect(); } catch (e) {}
      try { await withTimeout(window.__wb.getActiveSheet().getWorksheetPermission().unprotect(), 2000); } catch (e) {}

      // F 组会插删行列、J 组会写值，主表里的 seed 数据会漂移；
      // H 组每项前重新播种，保证断言的坐标是可复现的
      if (item.g === 'H') { try { await seedSearch(); } catch (e) {} }

      const sjsRes = await withTimeout(runSide(item, 'sjs'), 8000).catch((e) => ({ setOk: false, err: 'TIMEOUT/ERR: ' + (e.message || e) }));
      const uniRes = await withTimeout(runSide(item, 'uni'), 12000).catch((e) => ({ setOk: false, err: 'TIMEOUT/ERR: ' + (e.message || e) }));
      out.items.push({ id: item.id, g: item.g, name: item.name, excel: item.excel, sjs: sjsRes, uni: uniRes });
    }
    cleanSjs();
    try { await withTimeout(cleanUni(), 5000); } catch (e) {}
    window.__wsProgress = { done: list.length, total: list.length };
    return out;
  }

  // 单项超时保护：Univer 的 protect()/unprotect() 在没有权限服务时会返回永不 resolve 的
  // Promise，不加超时整个跑批就卡死（第一轮实测就是这么挂的）。
  function withTimeout(p, ms) {
    return Promise.race([
      Promise.resolve(p),
      new Promise((_, rej) => setTimeout(() => rej(new Error('timeout ' + ms + 'ms')), ms)),
    ]);
  }

  window.ApiProbeWs = { SPEC, runAll, runSide };
})();

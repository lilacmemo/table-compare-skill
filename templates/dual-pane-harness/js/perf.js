/**
 * 性能用例适配层：把同一份负载分别灌进 SpreadJS / Univer，并按统一口径计时。
 *
 * 口径（与 poc-compare/perf_run.js 一致）：
 *   1. setup      建表 + 灌入原始值      —— 不计时
 *   2. write      写入全部公式           —— 计时（SJS 在 suspendCalcService 状态下写）
 *   3. fullCalc   强制全量重算 N 轮       —— 每轮单独计时
 *   4. editRecalc 改一个输入后重算        —— 计时
 *   5. samples    读取 9 个核对样本       —— 用于校验双端结果一致
 */
window.PerfEngine = (function () {
  'use strict';

  const SHEET_DEFS = [
    ['Data', 10001, 9],
    ['Params', 4, 2],
    ['Products', 21, 2],
    ['Report', 20, 21],
    ['Chain', 1001, 4],
    ['Dynamic', 10001, 30]
  ];

  function colIdx(cell) {
    const m = /^([A-Z]+)(\d+)$/.exec(cell);
    let col = 0;
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    return { r: parseInt(m[2], 10) - 1, c: col - 1 };
  }

  function median(arr) {
    const a = arr.slice().sort((x, y) => x - y);
    const n = a.length;
    return n % 2 ? a[(n - 1) / 2] : (a[n / 2 - 1] + a[n / 2]) / 2;
  }

  // ---------------- SpreadJS ----------------
  const sjs = {
    label: 'SpreadJS',
    reset(hostId) {
      let sp = window.__spread;
      try { if (sp && sp.destroy) sp.destroy(); } catch (e) {}
      const host = document.getElementById(hostId);
      if (host) host.innerHTML = '';
      sp = new GC.Spread.Sheets.Workbook(host, { sheetCount: SHEET_DEFS.length });
      window.__spread = sp;
      if (window.__onSjsReplaced) window.__onSjsReplaced(sp);
      sp.options.allowDynamicArray = true;
      this._sheets = {};
      SHEET_DEFS.forEach(([name, rows, cols], i) => {
        const s = sp.getSheet(i);
        s.name(name);
        s.setRowCount(rows);
        s.setColumnCount(cols);
        this._sheets[name] = s;
      });
      return sp;
    },
    async setup(data) {
      const sp = this.reset('sjsHost');
      sp.suspendPaint();
      const setVals = (s, vals, r0 = 0, c0 = 0) => {
        for (let r = 0; r < vals.length; r++) {
          const row = vals[r];
          for (let c = 0; c < row.length; c++) {
            if (row[c] !== null && row[c] !== undefined) s.setValue(r0 + r, c0 + c, row[c]);
          }
        }
      };
      const S = this._sheets;
      setVals(S.Data, data.dataValues);
      setVals(S.Params, data.paramsValues);
      setVals(S.Products, data.productsValues);
      setVals(S.Report, data.reportValues);
      setVals(S.Report, data.reportPart2, 8, 0);
      setVals(S.Chain, data.chainValues);
      sp.resumePaint();
      return true;
    },
    write(data) {
      const sp = window.__spread;
      const S = this._sheets;
      sp.suspendCalcService(false);
      sp.suspendPaint();
      const t0 = performance.now();
      const ds = S.Data;
      for (let r = 1; r < data.dataFormulas.length; r++) {
        const row = data.dataFormulas[r];
        if (!row) continue;
        for (let c = 0; c < 4; c++) {
          const f = row[5 + c];
          if (f) ds.setFormula(r, 5 + c, f);
        }
      }
      const put = (sheet, cell, formula) => {
        const { r, c } = colIdx(cell);
        sheet.setFormula(r, c, formula);
      };
      for (const f of data.reportFormulas) put(S.Report, f.cell, f.formula);
      for (const f of data.chainFormulas) put(S.Chain, f.cell, f.formula);
      for (const f of data.dynFormulas) put(S.Dynamic, f.cell, f.formula);
      const writeMs = performance.now() - t0;
      sp.resumePaint();
      const t1 = performance.now();
      sp.resumeCalcService(false);
      const resumeMs = performance.now() - t1;
      return { writeMs, resumeMs };
    },
    fullCalc(runs) {
      const sp = window.__spread;
      const times = [];
      for (let i = 0; i < runs; i++) {
        const t0 = performance.now();
        sp.calculate(GC.Spread.Sheets.CalculationType.all);
        times.push(performance.now() - t0);
      }
      return times;
    },
    editRecalc() {
      const sp = window.__spread;
      const t0 = performance.now();
      this._sheets.Params.setValue(1, 1, 0.17);
      sp.calculate(GC.Spread.Sheets.CalculationType.all);
      return performance.now() - t0;
    },
    samples(list) {
      const out = [];
      for (const s of list) {
        const sh = this._sheets[s.sheet];
        const { r, c } = colIdx(s.cell);
        let v = sh.getValue(r, c);
        if (v instanceof GC.Spread.CalcEngine.CalcError) v = String(v);
        out.push({ sheet: s.sheet, cell: s.cell, value: v === undefined ? null : v });
      }
      return out;
    }
  };

  // ---------------- Univer ----------------
  const uni = {
    label: 'Univer',
    async setup(data) {
      const wb = window.univerAPI.createWorkbook({ name: 'perf' });
      window.__perfWb = wb;
      const S = {};
      for (const [name, rows, cols] of SHEET_DEFS) {
        let s = wb.getSheetByName(name);
        if (!s) s = wb.create(name, rows, cols);
        S[name] = s;
      }
      this._sheets = S;
      const setVals = (ws, vals, r0 = 0, c0 = 0) => {
        if (!vals.length) return;
        const nr = vals.length;
        const nc = Math.max(...vals.map(v => v.length));
        const grid = [];
        for (let r = 0; r < nr; r++) {
          const row = [];
          for (let c = 0; c < nc; c++) {
            const v = vals[r][c];
            row.push(v === undefined ? null : v);
          }
          grid.push(row);
        }
        ws.getRange(r0, c0, nr, nc).setValues(grid);
      };
      setVals(S.Data, data.dataValues);
      setVals(S.Params, data.paramsValues);
      setVals(S.Products, data.productsValues);
      setVals(S.Report, data.reportValues);
      setVals(S.Report, data.reportPart2, 8, 0);
      setVals(S.Chain, data.chainValues);
      return true;
    },
    write(data) {
      const S = this._sheets;
      const t0 = performance.now();
      const grid = [];
      for (let r = 1; r < data.dataFormulas.length; r++) {
        const row = data.dataFormulas[r];
        grid.push([row[5], row[6], row[7], row[8]]);
      }
      S.Data.getRange(1, 5, grid.length, 4).setFormulas(grid);
      const put = (sheet, cell, formula) => {
        const { r, c } = colIdx(cell);
        sheet.getRange(r, c).setFormula(formula);
      };
      for (const f of data.reportFormulas) put(S.Report, f.cell, f.formula);
      for (const f of data.chainFormulas) put(S.Chain, f.cell, f.formula);
      for (const f of data.dynFormulas) put(S.Dynamic, f.cell, f.formula);
      return { writeMs: performance.now() - t0, resumeMs: null };
    },
    _calcOnce(timeout) {
      const formula = window.univerAPI.getFormula();
      return new Promise((resolve) => {
        let done = false;
        const d = formula.calculationEnd(() => {
          if (!done) { done = true; try { d.dispose(); } catch (e) {} resolve(); }
        });
        formula.executeCalculation();
        setTimeout(() => {
          if (!done) { done = true; try { d.dispose(); } catch (e) {} resolve(); }
        }, timeout || 180000);
      });
    },
    async fullCalc(runs) {
      const times = [];
      for (let i = 0; i < runs; i++) {
        const t0 = performance.now();
        await this._calcOnce();
        times.push(performance.now() - t0);
      }
      return times;
    },
    async editRecalc() {
      const t0 = performance.now();
      this._sheets.Params.getRange(1, 1).setValue(0.17);
      await this._calcOnce();
      return performance.now() - t0;
    },
    samples(list) {
      const out = [];
      for (const s of list) {
        const sh = this._sheets[s.sheet];
        const { r, c } = colIdx(s.cell);
        const v = sh.getRange(r, c).getValue();
        out.push({ sheet: s.sheet, cell: s.cell, value: v === undefined ? null : v });
      }
      return out;
    }
  };

  /**
   * 跑一个完整性能用例
   * @param {'sjs'|'uni'} which
   * @param {object} data PerfWorkload.generate() 的结果
   * @param {number} rounds 全量重算轮次
   * @param {function} onStep 进度回调 (msg) => void
   */
  async function run(which, data, rounds, onStep) {
    const eng = which === 'sjs' ? sjs : uni;
    const say = onStep || function () {};
    say(`${eng.label}：建立 6 个工作表并灌入原始值…`);
    await eng.setup(data);
    say(`${eng.label}：写入 ${data.formulaCount.toLocaleString()} 条公式…`);
    const w = eng.write(data);
    say(`${eng.label}：写公式完成 ${w.writeMs.toFixed(0)} ms，开始全量重算 ${rounds} 轮…`);
    const times = await eng.fullCalc(rounds);
    say(`${eng.label}：改一个输入值后重算…`);
    const editMs = await eng.editRecalc();
    let samples = null;
    try { samples = eng.samples(data.samples); } catch (e) { samples = null; }
    // 跑完统一回到 Data 表左上角，便于左右对照
    try {
      if (which === 'sjs') {
        const sp = window.__spread;
        sp.setActiveSheetIndex(0);
        sp.getActiveSheet().setActiveCell(0, 0);
      } else {
        const wb = window.__perfWb;
        const sh = wb.getSheetByName('Data');
        if (sh && sh.activate) sh.activate();
      }
    } catch (e) {}
    return {
      engine: which,
      label: eng.label,
      writeMs: w.writeMs,
      resumeMs: w.resumeMs,
      fullCalcMs: times,
      fullCalcMedian: median(times),
      editRecalcMs: editMs,
      samples
    };
  }

  return { run, sjs, uni, median, SHEET_DEFS };
})();

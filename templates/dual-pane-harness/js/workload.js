/**
 * 公式性能基准 —— 共享工作负载生成器（确定性随机，双端完全一致）
 * 与 poc-compare/perf_workload.js 的负载定义保持一致（种子 42）。
 * 用法：const data = PerfWorkload.generate(9999);
 */
window.PerfWorkload = (function () {
  'use strict';

  function mulberry32(seed) {
    let a = seed >>> 0;
    return function () {
      a |= 0; a = (a + 0x6D2B79F5) | 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  const REGIONS = ['华东', '华南', '华北', '西南', '东北', '西北'];
  const CHAIN_ROWS = 1000;
  const PRODUCT_COUNT = 20;

  /**
   * @param {number} dataRows Data 表数据行数（默认 9999，即行 2..10000）
   */
  function generate(dataRows) {
    const DATA_ROWS = dataRows || 9999;
    const rnd = mulberry32(42);

    // Data：DATA_ROWS 行 × 9 列，A 区域 / B 产品 / C 数量 / D 单价
    const dataValues = [];
    for (let r = 0; r < DATA_ROWS; r++) {
      const qty = Math.floor(rnd() * 200) + 1;
      const price = Math.round((rnd() * 90 + 10) * 100) / 100;
      const region = REGIONS[Math.floor(rnd() * 6)];
      const prod = Math.floor(rnd() * PRODUCT_COUNT) + 1;
      dataValues.push([r + 1, region, prod, qty, price, null, null, null, null]);
    }
    // Data 公式：E=金额 F=税后 G=VLOOKUP H=评级
    const dataFormulas = [];
    for (let r = 2; r <= DATA_ROWS + 1; r++) {
      dataFormulas.push([
        null, null, null, null, null,
        `=D${r}*E${r}`,
        `=ROUND(F${r}*(1+Params!$B$2),2)`,
        `=VLOOKUP(C${r},Products!$A$2:$B$21,2,FALSE)`,
        `=IF(G${r}>Params!$B$3,"A",IF(G${r}>1000,"B","C"))`
      ]);
    }

    const paramsValues = [['参数', '值'], ['税率', 0.13], ['高额阈值', 5000], ['链式系数', 0.02]];

    const productsValues = [['产品ID', '系数']];
    for (let i = 1; i <= PRODUCT_COUNT; i++) {
      productsValues.push([i, Math.round((0.8 + rnd() * 0.6) * 1000) / 1000]);
    }

    // Report：120 SUMIFS + 120 SUMPRODUCT
    const reportValues = [['区域', ...Array.from({ length: PRODUCT_COUNT }, (_, i) => i + 1)]];
    for (const reg of REGIONS) reportValues.push([reg, null]);
    const reportPart2 = [['区域', ...Array.from({ length: PRODUCT_COUNT }, (_, i) => i + 1)]];
    for (const reg of REGIONS) reportPart2.push([reg, null]);

    const reportFormulas = [];
    for (let ri = 0; ri < 6; ri++) {
      for (let ci = 0; ci < PRODUCT_COUNT; ci++) {
        const colL = String.fromCharCode(66 + ci);
        const r = ri + 2;
        reportFormulas.push({
          cell: `${colL}${r}`,
          formula: `=SUMIFS(Data!$D$2:$D$${DATA_ROWS + 1},Data!$B$2:$B$${DATA_ROWS + 1},$A${r},Data!$C$2:$C$${DATA_ROWS + 1},${colL}$1)`
        });
      }
    }
    for (let ri = 0; ri < 6; ri++) {
      for (let ci = 0; ci < PRODUCT_COUNT; ci++) {
        const colL = String.fromCharCode(66 + ci);
        const r = ri + 10;
        reportFormulas.push({
          cell: `${colL}${r}`,
          formula: `=SUMPRODUCT((Data!$B$2:$B$${DATA_ROWS + 1}=$A${r})*(Data!$C$2:$C$${DATA_ROWS + 1}=${colL}$1)*Data!$D$2:$D$${DATA_ROWS + 1})`
        });
      }
    }

    // Chain：1,000 行三级链式依赖
    const chainValues = [['输入', '变换', '累计', '取整']];
    for (let r = 2; r <= CHAIN_ROWS + 1; r++) chainValues.push([Math.round(rnd() * 1000) + 1, null, null, null]);
    const chainFormulas = [];
    for (let r = 2; r <= CHAIN_ROWS + 1; r++) {
      chainFormulas.push({ cell: `B${r}`, formula: `=A${r}*Params!$B$4+1` });
      chainFormulas.push({ cell: `C${r}`, formula: r === 2 ? `=B${r}` : `=C${r - 1}+B${r}` });
      chainFormulas.push({ cell: `D${r}`, formula: `=ROUND(C${r},2)` });
    }

    // Dynamic：动态数组
    const dynFormulas = [
      { cell: 'A1', formula: `=UNIQUE(Data!B2:B${DATA_ROWS + 1})` },
      { cell: 'C1', formula: `=SORT(Data!F2:F2001,1,-1)` },
      { cell: 'E1', formula: `=SUMIFS(Data!$D$2:$D$${DATA_ROWS + 1},Data!$B$2:$B$${DATA_ROWS + 1},"华东")` }
    ];
    for (let i = 1; i <= PRODUCT_COUNT; i++) {
      const colL = String.fromCharCode(68 + i + 1);
      dynFormulas.push({ cell: `${colL}1`, formula: `=FILTER(Data!F2:F${DATA_ROWS + 1},Data!C2:C${DATA_ROWS + 1}=${i})` });
    }

    const samples = [
      { sheet: 'Data', cell: 'F2' }, { sheet: 'Data', cell: 'G2' }, { sheet: 'Data', cell: 'H2' },
      { sheet: 'Data', cell: 'I2' }, { sheet: 'Report', cell: 'B2' }, { sheet: 'Report', cell: 'B10' },
      { sheet: 'Chain', cell: 'C501' }, { sheet: 'Chain', cell: 'D501' }, { sheet: 'Dynamic', cell: 'E1' }
    ];

    // Data 每行写 4 条公式（F/G/H/I）
    const formulaCount = DATA_ROWS * 4 + reportFormulas.length + chainFormulas.length + dynFormulas.length;

    return {
      dataRows: DATA_ROWS,
      dataValues, dataFormulas, paramsValues, productsValues,
      reportValues, reportPart2, reportFormulas,
      chainValues, chainFormulas, dynFormulas, samples, formulaCount
    };
  }

  return { generate, CHAIN_ROWS, PRODUCT_COUNT };
})();

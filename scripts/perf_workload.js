// 公式性能基准 —— 共享工作负载生成器（确定性随机，双端一致）
// 用法：在 Node 侧 require，生成 JSON 后注入两个产品各自的宿主页
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
const DATA_ROWS = 10000;   // Data 表数据行数
const CHAIN_ROWS = 1000;   // Chain 表行数
const PRODUCT_COUNT = 20;

function generate() {
    const rnd = mulberry32(42);
    // ---- Data 表：10,000 行 × 9 列（A..I）----
    const dataValues = []; // [row][col] 起点 A1
    for (let r = 0; r < DATA_ROWS; r++) {
        const qty = Math.floor(rnd() * 200) + 1;
        const price = Math.round((rnd() * 90 + 10) * 100) / 100;
        const region = REGIONS[Math.floor(rnd() * 6)];
        const prod = Math.floor(rnd() * PRODUCT_COUNT) + 1;
        dataValues.push([r + 1, region, prod, qty, price, null, null, null, null]);
    }
    // Data 公式：F=金额 G=税后 H=VLOOKUP I=评级（行 2..10001，列 F..I）
    const dataFormulas = [];
    for (let r = 2; r <= DATA_ROWS + 1; r++) {
        dataFormulas.push([null, null, null, null, null,
            `=D${r}*E${r}`,
            `=ROUND(F${r}*(1+Params!$B$2),2)`,
            `=VLOOKUP(C${r},Products!$A$2:$B$21,2,FALSE)`,
            `=IF(G${r}>Params!$B$3,"A",IF(G${r}>1000,"B","C"))`]);
    }

    // ---- Params 表 ----
    const paramsValues = [['参数', '值'], ['税率', 0.13], ['高额阈值', 5000], ['链式系数', 0.02]];

    // ---- Products 表：20 行（含表头）----
    const productsValues = [['产品ID', '系数']];
    for (let i = 1; i <= PRODUCT_COUNT; i++) productsValues.push([i, Math.round((0.8 + rnd() * 0.6) * 1000) / 1000]);

    // ---- Report 表：120 SUMIFS + 120 SUMPRODUCT ----
    const reportValues = [['区域', ...Array.from({ length: PRODUCT_COUNT }, (_, i) => i + 1)]];
    for (const reg of REGIONS) reportValues.push([reg, null]);
    // 行 9 起第二块表头，行 10..15 SUMPRODUCT
    const reportFormulas = []; // {cell, formula}
    for (let ri = 0; ri < 6; ri++) {
        for (let ci = 0; ci < PRODUCT_COUNT; ci++) {
            const colL = String.fromCharCode(66 + ci); // B..U
            const r = ri + 2;
            reportFormulas.push({
                cell: `${colL}${r}`,
                formula: `=SUMIFS(Data!$D$2:$D$${DATA_ROWS + 1},Data!$B$2:$B$${DATA_ROWS + 1},$A${r},Data!$C$2:$C$${DATA_ROWS + 1},${colL}$1)`
            });
        }
    }
    const reportPart2 = [['区域', ...Array.from({ length: PRODUCT_COUNT }, (_, i) => i + 1)]];
    for (const reg of REGIONS) reportPart2.push([reg, null]);
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

    // ---- Chain 表：1,000 行链式依赖 ----
    const chainValues = [['输入', '变换', '累计', '取整']];
    for (let r = 2; r <= CHAIN_ROWS + 1; r++) chainValues.push([Math.round(rnd() * 1000) + 1, null, null, null]);
    const chainFormulas = [];
    for (let r = 2; r <= CHAIN_ROWS + 1; r++) {
        chainFormulas.push({ cell: `B${r}`, formula: `=A${r}*Params!$B$4+1` });
        chainFormulas.push({ cell: `C${r}`, formula: r === 2 ? `=B${r}` : `=C${r - 1}+B${r}` });
        chainFormulas.push({ cell: `D${r}`, formula: `=ROUND(C${r},2)` });
    }

    // ---- Dynamic 表：动态数组 ----
    const dynFormulas = [
        { cell: 'A1', formula: `=UNIQUE(Data!B2:B${DATA_ROWS + 1})` },
        { cell: 'C1', formula: `=SORT(Data!F2:F2001,1,-1)` },
        { cell: 'E1', formula: `=SUMIFS(Data!$D$2:$D$${DATA_ROWS + 1},Data!$B$2:$B$${DATA_ROWS + 1},"华东")` },
    ];
    for (let i = 1; i <= PRODUCT_COUNT; i++) {
        const colL = String.fromCharCode(68 + i + 1); // G.. 起每列一个 FILTER
        dynFormulas.push({ cell: `${colL}1`, formula: `=FILTER(Data!F2:F${DATA_ROWS + 1},Data!C2:C${DATA_ROWS + 1}=${i})` });
    }

    // 核对样本（双端读取后比对）
    const samples = [
        { sheet: 'Data', cell: 'F2' }, { sheet: 'Data', cell: 'G2' }, { sheet: 'Data', cell: 'H2' },
        { sheet: 'Data', cell: 'I2' }, { sheet: 'Report', cell: 'B2' }, { sheet: 'Report', cell: 'B10' },
        { sheet: 'Chain', cell: 'C501' }, { sheet: 'Chain', cell: 'D501' }, { sheet: 'Dynamic', cell: 'E1' },
    ];

    const formulaCount = DATA_ROWS * 5 + reportFormulas.length + chainFormulas.length + dynFormulas.length;
    return { dataValues, dataFormulas, paramsValues, productsValues, reportValues, reportPart2, reportFormulas, chainValues, chainFormulas, dynFormulas, samples, formulaCount };
}

module.exports = { generate, DATA_ROWS, CHAIN_ROWS };

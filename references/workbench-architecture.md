# 单工程双引擎对比台（对比工程）架构说明

> v1 · 2026-11-06 · 由 SpreadJS vs Univer 六模块实战沉淀
> 适用场景：用户要求「**一个工程里同时集成两家的全部功能**」，用同一套用例、同一份样本、同一个跑批器出结论。

## 0. 为什么要有它（取代旧做法的理由）

旧做法是「文档查证 + 线上 playground 探针」：Univer 侧在 `office.univer.ai` 逐条敲代码、SpreadJS 侧在另一个零散 POC 里跑。三个硬伤，实战全部踩到：

| 硬伤 | 具体表现 | 后果 |
| :-- | :-- | :-- |
| **口径不可比** | SJS 在本地 v19.2.5、Univer 在线上站点某个构建，两边版本会漂 | 同一项结论今天 PASS 明天 FAIL，被客户一问就崩 |
| **样本对不上** | 用例里写第 140 行，样本只有 20 行 | 整批"假失败"，还以为是产品缺陷 |
| **没法复跑** | 结论只活在一次性的临时脚本里 | 客户/同事质疑时无法当场重算，报告可信度靠"你信我" |

对比台把这三件事一次性解决：**同一页面、同一份样本、同一套用例表、一条命令出 JSON**。

## 1. 目录结构与文件职责

```
对比工程/
├─ index.html                 单页：左右双栏（#paneSjs / #paneUni）+ 模块下拉 + 用例结果表
├─ 交互手感-复核台.html        交互手感 40 项人工复核页
├─ 打印元素矩阵.html           元素层 21 项三档判定复核页（读 js/cases/print-elements.js）
├─ js/
│  ├─ app.js                  装配层：initSjs / initUni / loadSample / exportSjs / exportUni /
│  │                          statsSjs / statsUni / runCases  ← 全局 APP 对象
│  ├─ runner.js               代码串执行器：__cmpClear(side) / __cmpRun(side, code, timeout)
│  ├─ cases/                  用例注册表，一个模块一个文件，全部挂到 window.CASES
│  │  ├─ charts.js  print.js  worksheet.js  format.js  table.js  formula.js
│  │  └─ print-elements.js    元素层前端清单（window.PRINT_ELEMENTS）——**生成物，勿手改**，
│  │                          由 tools/build_print_matrix_verdict.py 从 print-matrix-items.json 出
│  ├─ probe-ws.js             工作表 127 项探针表（window.ApiProbeWs）
│  ├─ probe-format.js         格式 137 项探针表（window.ApiProbe）
│  ├─ probe-table.js          表/筛选/验证探针（window.CASES_TABLE 的代码串来源）
│  └─ data/formula-oracle.json  500 条 Excel 缓存值 oracle（由 tools/build_formula_oracle.py 生成）
├─ samples/                   六份模块样本 xlsx（Excel 本体产出，见 SKILL.md 第 1 步）
├─ vendor/
│  ├─ spreadjs/               @grapecity-software 分模块包（all/shapes/charts/io/pdf）
│  └─ univer-pro.bundle.js    Univer Pro 本地整包（25 MB，含 license）
├─ tools/
│  ├─ serve.js                HTTPS 静态站 + /universer-api/* 反向代理
│  ├─ build_pro_bundle.js     打 Pro bundle（esbuild）
│  ├─ run_cases.js            通用跑批：开页 → 跑模块 → 落 out/data/<模块>-cases.json
│  ├─ capture_io.js           导入导出抓取：截图 + 导出 xlsx + 计数
│  ├─ diff_io.py              导出文件双层比对（zip 部件层 + openpyxl 内容层）
│  ├─ compose_io_figure.py    拼对比图：默认并排版；--doc 出竖排文档版 → docs/assets/
│  ├─ rebuild_print_matrix.ps1  打印元素矩阵一键重建（样本→三家出页→判定→出图→前端清单）
│  ├─ build_print_matrix_verdict.py  读 print-matrix-items.json → verdict.json + print-elements.js
│  ├─ compose_print_matrix.py  元素矩阵对照图 → docs/assets/print-mx-*.png
│  └─ build_formula_oracle.py 从样本抽 Excel 缓存值成 JSON oracle
└─ out/                       全部产物，可整目录删除重跑
   ├─ data/    <模块>-cases.json / <模块>-io.json / <模块>-diff.md / _io-all.json
   ├─ shots/   <模块>-sjs.png / <模块>-uni.png / <模块>-对比图.png
   └─ export/  <模块>-sjs.xlsx / <模块>-uni.xlsx
```

**分层原则**：`app.js` 只管引擎装配和全局工具，**不认识任何一条用例**；`cases/*.js` 只管产出 `{id, name, sjs, uni, note}` 四元组，**不认识 DOM**。换模块 = 换一个 `cases/*.js`，装配层一行不改。

## 2. Univer Pro 本地集成（最容易卡住的一节）

### 2.1 三件套缺一不可

```js
// 1) advanced preset 必须叠在 core preset 之上
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import { UniverSheetsAdvancedPreset } from '@univerjs/preset-sheets-advanced'

const { univerAPI } = createUniver({
  locale: LocaleType.ZH_CN,
  presets: [
    UniverSheetsCorePreset({ container: 'uniHost' }),
    UniverSheetsAdvancedPreset({
      license: '<license 文本>',
      universerEndpoint: 'https://office.univer.ai',   // 走本地代理时必须同源
    }),
  ],
})
```

### 2.2 **插件与 Facade 必须成对注册**（本轮最大的假阴性来源）

Univer 每个能力 = npm 包 + `registerPlugin` + `/facade` 子路径 **三件套**，漏任一件：
- 只 import facade、不 `registerPlugin` → 命令没注册，调用抛 `redi: Expect 1 dependency item(s) for id "xm"` 或 `command "sheet.command.addDataValidation" is not registered`
- 只 `registerPlugin`、不 import facade → `newConditionalFormattingRule` / `createTextFinderAsync` / `setCrosshairHighlightEnabled` / `setBackgroundImage` 直接 `undefined`

实测影响：格式模块 Univer 从 95 掉到 **63**（`newConditionalFormattingRule is not a function`），表模块从 28 掉到 **13**（筛选/校验全抛错）。

```js
// entry 里统一成对写；注册必须逐个 try，一个 already registered 不能中断后续
const extras = {
  UniverFindReplacePlugin, UniverSheetsFindReplacePlugin,
  UniverSheetsCrosshairHighlightPlugin, UniverSheetsDrawingPlugin,
  UniverSheetsFilterPlugin, UniverSheetsFilterUIPlugin,
  UniverSheetsDataValidationPlugin, UniverSheetsTablePlugin,
  UniverSheetsConditionalFormattingPlugin,
  // 批注 = 4 个包 + 2 个 facade；漏装的表现是 sheet.getComments 为 undefined，
  // 会被误读成「Univer 不支持批注」。
  UniverThreadCommentPlugin, UniverThreadCommentUIPlugin,
  UniverSheetsThreadCommentPlugin, UniverSheetsThreadCommentUIPlugin,
}
window.__proExtraPlugins = {}
for (const [name, P] of Object.entries(extras)) {
  try { univer.registerPlugin(P); window.__proExtraPlugins[name] = 'ok' }
  catch (e) { window.__proExtraPlugins[name] = 'skip: ' + String(e && e.message).slice(0, 80) }
}
```

`window.__proExtraPlugins` 是**自证清单**：报告里写"已注册 9 个插件"时，依据就是这个对象，不是记忆。

**自检方法**：任何"Univer 这个方法不存在"的结论，先跑一遍
`Object.keys(window.__proExtraPlugins).filter(k => window.__proExtraPlugins[k] !== 'ok')`，
非空就说明是工程缺件，不是产品缺陷。

**这条已升级为跑批前置门禁**（2026-11-14 用户定版，见 SKILL.md 第 0 步）：不是等出问题再回溯，而是**跑任何用例之前先断言**。可执行版 `scripts/preflight_univer.js`，三组检查任一组失败即 `exit 1`：

1. `deps` —— 该模块所需的包是否都在 `node_modules` 里（功能包 + UI 包 + `/facade` 子路径三件套），并用**依赖闭包**区分"合法间接依赖"与"孤立安装"（只有后者会被 `npm i` 清掉，如历史上的 react / react-dom）。
2. `plugins` —— `__proExtraPlugins` 全为 `ok`（页面内调用时检查，Node 侧跳过）。
3. `facade` —— 传入的方法名数组逐个 `typeof !== 'undefined'`。

用法：`node scripts/preflight_univer.js --module comment --facade "sheet.getComments,univerAPI.getActiveWorkbook"`；页面内在完成 init 后执行 `await preflight(...)` 拿同一份报告。

### 2.3 advanced preset **不自动建簿**

`UniverSheetsAdvancedPreset` 会接管建簿流程，页面上出现空壳但 `getActiveWorkbook()` 返回 null。必须在 init 后显式：

```js
if (!api.getActiveWorkbook()) api.createWorkbook()
window.__wb = api.getActiveWorkbook()          // 探针要用，必须重挂
```

**导入后 `__wb` 会失效**（旧簿被 dispose），必须重挂一次，否则探针报 `Range is out of bounds` 或表名不符：

```js
APP.uni.createWorkbook(snap)                    // 导入
await sleep(400)
window.__wb = APP.uni.getActiveWorkbook()       // 重挂，这一步漏了整批跑飞
```

### 2.4 exchange 导入导出必须走服务端 + HTTPS + 同源代理

链路：`/universer-api/stream/file/upload` → `/exchange/{type}/import` → `/exchange/task/{id}` → `/file/{id}/sign-url`。

- **必须 HTTPS**：`POST /universer-api/anonymous` 种下的 `Serve` cookie 带 `Secure` 标记，HTTP 下浏览器不保存 → 后续 upload 直接 401。
- **必须代理**：官方端点不带 CORS 头，从本机直连会被浏览器拦。
- **代理不能删 `content-encoding`**：上游返回 gzip，删了这个头会把压缩字节当明文喂给 `JSON.parse`，表现为 `importSheetToSnapshotAsync` **静默返回 undefined**（不报错，最难查）。
- 代理要把 `set-cookie` 里的 `Domain=office.univer.ai` 和 `Secure` 摘掉，否则本机存不下。

`scripts/io_serve_https.js` 已实现全部四点，自签证书放 `tools/cert.pem|key.pem`，Playwright 用 `ignoreHTTPSErrors: true` 忽略告警。

## 3. 用例注册表契约（新增模块只做这一件事）

`js/cases/<module>.js` 二选一：

```js
// 形态 A：批量式（探针表已有全量 SPEC，只做映射）
window.CASES = window.CASES || {}
window.CASES.worksheet = {
  title: '工作表与工作簿',
  total: 127,
  async batch() {
    const r = await window.ApiProbeWs.runAll()
    const results = (r.items || []).map((i) => ({
      id: i.id, name: i.g + ' · ' + i.name,
      sjs: cell(i.sjs), uni: cell(i.uni), note: noteOf(i),
    }))
    return { results, meta: { sheetInfo: r.sheetInfo } }
  },
}

// 形态 B：逐条式（每条自带两侧代码串，适合"同一语义两种写法"的对比）
window.CASES.table = {
  title: '表 / 排序 / 筛选 / 验证', total: 46,
  cases: [{ id: 'T-A1', name: '…', async run() { /* 返回 {sjs,uni,note} */ } }],
}
```

结果四元组的硬约定：

| 字段 | 取值 | 说明 |
| :-- | :-- | :-- |
| `sjs` / `uni` | `PASS` / `FAIL` / `NA` | **只有这三档**，不要中间态 |
| `note` | 字符串 | `FAIL` 必写原因；`NA` 也要写理由（详见下节） |

`APP.runCases(moduleKey)` 会算 `{pass, fail, na, total}` 并渲染表格，`tools/run_cases.js` 落盘同结构 JSON。**报告里的数字只能来自这个 JSON**，不许手抄。

### 3.0 形态 C：画面判定型（不走 cases 数组）

有一类用例判不了自动化——`一致/降级/丢失` 三档里的"降级"必须人看图（打印元素矩阵 21 项、交互手感 40 项）。它们**不进 `CASES.<module>.cases`**：

- 判定写在 `out/data/<模块>-items.json`（唯一数据源），脚本据此生成 `-verdict.json` 与前端清单 `js/cases/<模块>-elements.js`；
- 界面用**独立复核页**（`打印元素矩阵.html` / `交互手感-复核台.html`），顶栏按钮只在对应模块显示；
- 跑批文件里只挂 summary（`print-cases.json` 的 `elements` 字段），**不并入 PASS/FAIL 分母**。

塞进 cases 数组的代价：`verdict()` 只认 `PASS/FAIL/NA`，三档值全部渲染成「—」；而且两档判定与三档判定混在一张表，读者会默认是一个分母。

### 3.1 NA 的三种理由要分开写

NA 不是一句话能盖住的。`format.js` 里做了分级，照抄：

```js
const NO_EXCEL = 'Excel 本身无此设置（不计入分母）'          // Excel 真没有
const NA_REASON = {
  F8: '归导入链路（见 §3 数据层比对），不在 API 表重复计分',   // 别处已覆盖
  H15: '归渲染层（见 §3.2 人工复核），不进 API 通过率',        // 效果层
}
```

教训：第一版统一写"Excel 无此能力"，把"导入保真""渲染综合"也塞进同一句话，文档里就解释不通了。

## 4. 三条命令（全部可复跑）

```bash
# 0) 起服务（HTTPS + 代理），保持后台
node tools/serve.js 8124

# 1) 跑某个模块的 API 用例 → out/data/<模块>-cases.json
NODE_PATH=../poc-compare/node_modules node tools/run_cases.js format

# 2) 抓导入导出（截图 + 导出文件 + 计数）→ out/shots / out/export / out/data
#    这一步会顺带产出「导出后重新打开」的截图 <模块>-<side>-rt.png
NODE_PATH=../poc-compare/node_modules node tools/capture_io.js format

# 3) 比导出差异 → out/data/format-diff.md
python tools/diff_io.py format

# 4) 拼对比图：默认并排版 → out/shots/<模块>-对比图.png
python tools/compose_io_figure.py format

# 5) 出文档版竖排图（嵌 md 用）→ docs/assets/<模块>-io-compare.png
python tools/compose_io_figure.py --doc format

# 6) 出「导出后重新打开」的图 → docs/assets/<模块>-io-roundtrip.png
python tools/compose_io_figure.py --rt --doc format

# 7) 逐条用例表 → 追加到文档末尾的「附录 A」
node tools/gen_case_tables.js format
```

`NODE_PATH` 指向任意已装 playwright 的目录（本机是 `poc-compare/node_modules`）。浏览器用 `channel:'msedge'`（本机无 Chrome，同 Chromium 内核）；装了 Chrome 就换 `channel:'chrome'`。

## 4.1 报告里的证据图怎么嵌（2026-11-06 增补）

`.md` 报告可以直接嵌图，相对路径即可：`![说明](assets/<模块>-io-compare.png)`。但**不能直接嵌并排版原图**——并排版 1696px 宽，md 正文按 800px 显示时每侧只剩 400px，表里的字全糊了。所以 `compose_io_figure.py` 出两个尺寸：

| 版本 | 命令 | 尺寸 | 用途 |
| :-- | :-- | :-- | :-- |
| 并排版 | `compose_io_figure.py` | 1696×851（两侧各 820） | 汇报 / PPT / 微信发送 |
| **文档版** | `compose_io_figure.py --doc` | 1140×2084（上下竖排，单侧 1100） | **嵌进 md 正文**，缩到 800px 仍可读 |

三条硬约束：

1. **文件名必须 ASCII**（`format-io-compare.png`，不是 `格式-对比图.png`）——转 Word / 飞书 / GitHub 时中文路径会丢图。
2. **图放 `docs/assets/`，md 里写 `assets/xxx.png`**，交付时整个 `docs/` 一起给；只发单个 md 文件图会全挂。要单文件自包含就把图转 base64 data URI 内联（体积 +33%）。
3. 图下加一行 `<sub>图 3-1　上：SpreadJS x.y；下：Univer Pro a.b（已授权）。同一份 `<样本名>`，未做任何手动调整。</sub>`——**"未做任何手动调整"必须写**，否则读者会以为差异是调出来的。

## 4.2 往返截图：怎么证明「导出也没坏」（2026-11-06 增补）

导入侧截图只能证明"读得进"，销售最常被追问的是"那**导出**呢？"——尤其是"导入时丢的东西，会不会是导出才丢的"。取证方式是把各家**刚导出的 xlsx 再喂回同一个引擎**，截第二张图：

| 环节 | 通道 | 产物 |
| :-- | :-- | :-- |
| 导出 | `APP.exportXlsx(side)` | `out/export/<模块>-<side>.xlsx` |
| **重开** | `APP.importOne(side, file, name)`（新增的单侧通道，只喂一侧，另一侧不动） | — |
| 截图 | 重开后再跑一次 `alignViews()`，等 1.6s 渲染 | `out/shots/<模块>-<side>-rt.png` |
| 拼图 | `compose_io_figure.py --rt [--doc]` | `docs/assets/<模块>-io-roundtrip.png` |

三个必须注意的点：

1. **`APP.importOne` 必须存在**。原来的 `APP.import` 是双侧同导，往返验证时会把另一侧也冲掉，截图就不可比了。
2. **重开后必须重新对齐视图**（同表名、A1、视口归零、zoom 0.8），否则两张图像素对不上。视口复位的坑见 `pitfalls.md`「单工程双引擎对比台」第 11 条。
3. **SJS 重开会多一张 `Evaluation Version` 授权提示表**（试用版导出即带）。`alignViews` 要按首次导入的表名去定位、找不到才取第一张非授权表；否则截到空白授权页，看起来像"导出全丢了"。

往返图最有说服力的一处用法（图表模块实例）：

> 导入后 SJS 63 图存活、Univer 59 图存活（丢 4 张）→ 往返后仍是 63 / 59。**图数不随导出再掉**，就证明丢失发生在**导入**环节，导出只是把现状固化。没有这张图，对手一句"你们导出也丢吧"就无法回答。

## 4.3 附录 A：逐条用例表（2026-11-06 增补）

§4 只给分组通过率，客户/销售追问"具体哪个 case 没过"时当场答不出来。`tools/gen_case_tables.js` 把 `out/data/<模块>-cases.json` 渲成 md 表格，**追加**到对应文档末尾：

```bash
node tools/gen_case_tables.js          # 五个模块全刷
node tools/gen_case_tables.js format   # 只刷一个
```

设计要点（都是踩过的坑）：

1. **幂等**：先截到 `## 附录 A` 之前再追加，重复跑不会叠两份。
2. **同源**：表里的 PASS/FAIL 直接读 cases.json，不手抄，杜绝"正文和附录两个数"。
3. **单元格净化**：`|` 转义成 `\|`、换行压成空格，否则整张表错位（一行错会带歪后面全部）。
4. **截断要收边界**：探针把整段 JSON 塞进说明（`sample={"id":…}`），长说明截断必须①把长 JSON 收成 `{…}`；②不能停在半个括号/半个英文词上。否则读者看到 `（颜色须 rgb 形` 这种断头句，比不写还糟。
5. **图表模块不生成附录**：`01-图表.md` 的 §4 本身就是 A～L 十二组 103 项的逐条表，再追加一份会打架。生成器里 `MAP` 显式跳过 charts，并在文档 §4 开头写明"本节即逐条表"。
6. **模块注（`MAP[mod].note`）不能省**：
   - 打印：PASS = 有 API 且读回语义一致，Univer 读回走事件载荷；
   - 公式：PASS = 形状与每格结果**与 Excel 一致**（数值按 15 位有效数字归一后逐字比、文本原样、不设相对容差）；FAIL 含"不一致"与"`#NAME?` 不支持"两类，**必须拆成支持度 / 准确度两个分母**（485 / 483），别混成一个数；
   - 表模块：说明列连 PASS 行也记了读回值，那是给人复核的，不是 FAIL 的证据。



## 4.4 UI 降级台账（2026-11-06 增补，配合 verdict-framework「判定红线 ①」）

「API 通但效果不对」的项不能靠人记，要落成数据，否则下一轮重跑又变回 PASS。

**两份东西**：台账 `js/data/ui-downgrade.js`（数据）+ 判定链 `applyUiDowngrade()`（`js/app.js`，在 `renderResults` 之前跑）。

```js
// js/data/ui-downgrade.js —— 只登记"真看过渲染结果"的项
window.UI_DOWNGRADE = {
  charts: {
    K1: { sjs: 'API 可设 3D 视角，图表实际按 2D 平面渲染/导出，用户看不到 3D' },
  },
  print: {
    // 例：P27 页脚-右本身是纯文本测试，仍判 PASS；
    // &F 令牌的问题记在文档 §4.3，不要挂到 P27 上（会把一个真 PASS 打成假 FAIL）
  },
};
```

```js
// js/app.js —— 只对**当前判 PASS 的那一侧**生效，原始 API 结论另存
function applyUiDowngrade(results, moduleKey) {
  const reg = (window.UI_DOWNGRADE || {})[moduleKey] || {};
  results.forEach((r) => {
    const d = reg[r.id];
    if (!d) return;
    ['sjs', 'uni'].forEach((k) => {
      if (!d[k] || r[k] !== 'PASS') return;
      r[k + 'Api'] = r[k];                  // 原始结论：报告里写"API 通、UI 降级"
      r[k] = 'FAIL';
      r.note = (r.note && r.note !== '—' ? r.note + '；' : '') + 'UI 降级：' + d[k];
    });
  });
}
```

三条纪律：
1. **用例 ID 要选对**。别把输出层缺陷挂到字段层用例上（例：页脚 `&F` 不输出 ≠ 「页脚-右」这条纯文本用例失败）。
2. **只写"看过"的**。没复核的项写进文档「待复核」，不进台账——否则台账变成猜测库。
3. **`index.html` 里台账必须早于 `js/cases/*` 加载**，否则判定时读不到 `UI_DOWNGRADE`。

## 4.5 Excel oracle：需要"Excel 说什么"时的取法（2026-11-06 增补）

| 要什么 | 怎么取 | 脚本 |
| :-- | :-- | :-- |
| 页面设置（方向/纸张/缩放/边距/页眉页脚/居中/网格线/打印区域） | COM 逐表 `Activate()` 后读 `PageSetup` | `excel-print-oracle.ps1` → `out/data/print-excel-oracle.json` |
| 打印输出（页数/页尺寸/文本层/每页内容） | `ExportAsFixedFormat(0, pdf)` | 同上 → `out/pdf/excel-print.pdf` |
| 公式实算值 | 活公式 sheet + Excel 计算 + 读缓存值 | 公式模块管线 |

注意事项全在 `pitfalls.md` 第 22 条（BOM / 逐表激活 / pt 与 1/100 inch 换算）。

按顺序做，缺一项就跑不起来：

1. `tools/serve.js`（`io_serve_https.js`）+ 自签证书 → 起 8124
2. `vendor/spreadjs/`：`@grapecity-software/spread-sheets` + `-shapes` + `-charts` + `-io` + `-pdf`（**加载顺序不可换**，charts 依赖 Shapes）
3. `vendor/univer-pro.bundle.js`：`tools/build_pro_bundle.js` + esbuild，license 从客户授权文件读
4. `js/app.js`：装配 + `APP.*` 工具（照现有文件改 SAMPLES 表即可）
5. `samples/`：六份模块样本（**先造样本，见 SKILL.md 第 1 步**）
6. `js/cases/<模块>.js` + 对应 `probe-*.js`
7. `tools/run_cases.js` / `capture_io.js` / `diff_io.py` / `compose_io_figure.py`

## 6. 已知边界（写进报告的"局限"节）

- **图表/形状这类 canvas 内容无法 DOM 量测**，效果层只能截图 + 人工复核；API 层不依赖截图。
- **Univer 的 `updatePrintConfig` 会静默吸收未知字段并原样回显**，所以"读回有这个键"毫无证明力，字段白名单必须来自 `.d.ts`（详见 SKILL.md 打印模块一节）。
- **`setActiveSheet(sheet)` 不刷新画布**，截图对齐要用 `setActiveSheetIndex(0)`；且样本两家表顺序可能不同，必须**按表名**对齐。
- **SpreadJS 试用版会插 "Evaluation Version" 授权 sheet**，统计工作表数时要排除，否则差值恒为 1。

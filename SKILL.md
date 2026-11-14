---
name: table-sdk-compare
agent_created: true
description: 表格类 SDK（SpreadJS/Univer/Handsontable/ag-Grid 等）功能覆盖度对比方法论。以 Excel 对象模型为基准线，通过"基准清单→官方功能清单反推→文档查证→自动化实测→内部模型深挖→判定分级（API 类模块 PASS/FAIL 二值、图表保真类 S1～S5、公式支持度/准确度双分母）→Chrome 有头渲染闭环"产出带证据的竞争分析报告，内置"保真层/可编程层"两分法、PASS/FAIL 公开 API 判定、判定尺子标定法与销售话术规则。当用户需要对比两个表格/电子表格 SDK 的功能差异、生成对比矩阵、做选型或竞争分析、验证 xlsx 导入导出保真度、对比公式函数准确度或计算性能时使用。
---

# 表格 SDK 功能对比（Excel 基准法 v3）

以 Excel 桌面版为"满分答案"，逐项核对参评 SDK 的覆盖度。官网对比只够表层，真实差距藏在配置颗粒度、可编程深度和 xlsx 兼容性里——这些只能实测。本方法论经 SpreadJS vs Univer 图表模块完整实战校验（63 图基线 + 6 轮实测修正 + 批次 1 全量配置项实测），公式模块已按"支持度/准确度"双分母定稿（500 条用例 / 15 位口径 / 本地工程测性能），每轮被证据推翻的结论均已转化为流程修正。

**范围约定**：主战场为原生 Excel 图表（Excel 63 种类型 + 配置项）。Data Charts（数据管理器驱动图表）与迷你图 Sparkline 属可选扩展维度（见 dimensions-template M/N 组），默认不测，仅在用户点名时启用。

**对比基线（2026-11-05 用户最终定版，优先级最高）**：所有验证固定为 **SpreadJS 商业版 vs Univer Pro**，**不对比开源版 Univer**——开源版行为一律不进入结论矩阵，也不作为销售口径（背景理解可用）。操作含义：
- Univer 侧实测环境 = **office.univer.ai 官方 playground**（已实证为 Pro 构建：bundle 内含 `@univerjs-pro` 标识，探测方法见下）或客户正式授权环境；本地用 npm 开源 preset（`@univerjs/preset-sheets-core` 等）搭的实例只可做机制探查，**结论不得引用**。
- 文档查证以 `docs.univer.ai` 为主；涉及开源/Pro 分层的能力（如跨工作簿稳定 ID 绑定 `FFormula.buildReference`、图表）必须落到 **Pro 官方文档页面**并附 Pro 环境实测。
- **证据合法性（2026-11-05 用户定版，最高优先级）**：**"支持"只认官方文档 / API Reference 记载的公开 API**；`.d.ts`、npm 包导出、源码实现、未文档化的命令参数**一律不得作为支持依据**，只能在备注写"存在但未公开"且不计入统计。文档里明确写"不支持"的，直接引文档原文（例：Univer 官方功能对比页写明"不提供内置单变量目标求解"，比 grep 源码 0 命中更有力）。实测只用于验证"文档写的能力是否真生效"，不用于发现新能力。
- Pro 构建判定法：①拉取 playground 首屏 bundle（`document.querySelector('script[src*="assets/index"]').src` → fetch 文本）grep `@univerjs-pro|univerjs-pro|UniverPro`；②结合站点 `Last-Modified` 与 `npm view @univerjs-pro/* time` 判版本号。参考脚本 `scripts/univer_pro_dynarray_check2.js`（含 Pro 判定 + Facade 动态数组复测模板）。

## 总体交付方案（v5，2026-11-06 用户定版：单工程双引擎对比台）

> 用户原话：「先创建一个工程，里边分别集成 SpreadJS 和 Univer Pro 的全部功能……设计测试用例，基于 Excel 标准能力构造，并且要有具体的落地文件……做导入导出兼容性对比……API 可靠性对比（未使用公开 API、设置成功 PASS，否则 FAIL）……**测试工程每个功能的测试要和你设置的 case 吻合**。」

**这一条取代"零散 POC + 线上 playground 探针"的旧组织方式**。八个交付要件，缺一不可：

| # | 要件 | 落地物 | 不合格的样子 |
| :-: | :-- | :-- | :-- |
| 1 | **一个工程集成双引擎** | `对比工程/index.html` 左右双栏，Univer 侧必须是**本地授权 bundle**，不是线上 playground | 两侧在不同页面、不同版本，结论不可比 |
| 2 | **用例绑定落地样本** | `samples/<模块>-<要素>.xlsx`，用例里每一项都能在样本中指认到载体 | 用例写第 140 行、样本只有 20 行 |
| 3 | **导入导出对比** | `tools/capture_io.js` 截图 + `tools/diff_io.py` 双层比对 + `tools/compose_io_figure.py` 对比图 | 只写"导入正常"，没有截图也没有差异清单 |
| 4 | **API 可靠性** | PASS = 公开 API set 成功（strict 项加读回一致）；FAIL 逐条写原因 | 拿未文档化字段/内部 injector 判"支持" |
| 5 | **用例与工程吻合** | 报告里每个数字都能用 `node tools/run_cases.js <模块>` 复现 | 报告数字来自手抄或旧轮次 |
| 6 | **导出侧单独取证** | 各家刚导出的 xlsx **再喂回同一引擎**截图（`<模块>-<side>-rt.png` → `assets/<模块>-io-roundtrip.png`），配合 `diff_io.py` 的 XML 级比对 | 只有"导入后"一张图，被问"导出也丢吧"答不上来 |
| 7 | **逐条用例表** | 每份文档末尾「附录 A」：ID / 测试项 / 两侧判定 / 原因，由 `tools/gen_case_tables.js` 从 `cases.json` 生成（图表模块例外，其 §4 即逐条表） | 只有分组通过率，客户问"哪个 case 没过"当场翻不出来 |
| 8 | **UI 降级台账** | `js/data/ui-downgrade.js` + `applyUiDowngrade()`：API 通但效果降级的项改判 FAIL，原始 API 结论留 `sjsApi/uniApi`，说明列自动写「UI 降级：…」 | 拿"API 调用成功"当通过率上报，客户一试就露 |

**架构、契约、命令、最小复刻清单** → `references/workbench-architecture.md`。新增模块只写一个 `js/cases/<模块>.js`（产出 `{id,name,sjs,uni,note}` 四元组），装配层不动。

**六模块实测结果（同一工程内复跑，报告在 `对比工程/docs/`）**：

| 模块 | 用例 | 可判 | SpreadJS | Univer Pro |
| :-- | :-: | :-: | :-: | :-: |
| 图表 | 103 | 103 | 87（84.5%） | 81（78.6%） |
| 打印与 PDF | 35 | 35 | 35（100%） | 19（54.3%） |
| 工作表 / 工作簿 | 127 | 127 | 118（92.9%） | 94（74.0%） |
| 格式与条件格式 | 137 | 124 | 122（98.4%） | 95（76.6%） |
| 表 / 排序 / 筛选 / 验证 | 46 | 46 | 46（100%） | 28（60.9%） |
| 公式 | 519 函数 / 504 进分母 | 504 | 支持 494（98.02%）· 准确 448（91.80%） | 支持 491（97.42%）· 准确 324（66.80%） |

> 图表 88→87 是 2026-11-06 用户定 UI 降级口径后改的：`K1` 3D 图表类型 API 通过但渲染成平面 2D，判 FAIL。
> 打印模块这一行的 35 项只覆盖**配置字段可读写**。另有两组独立口径，**都不并入 35 项分母**：①页面设置**导入保真**（Excel oracle 为准）SpreadJS **8/8**、Univer **2/8**；②**打印元素支持矩阵 21 项**（一致/降级/丢失）SpreadJS **19/0/2＝90.5%**、Univer **11/3/7＝52.4%**。详见 `02-打印与PDF.md` §2.5 / §2.8。作业细则见 `references/print-pdf-module.md`。
> 公式行是**两个分母**：支持度 ÷ 500（含 5 条易失函数、`IMAGE`），准确度 ÷（支持数 − 该家支持的易失数）= 485 / 483。作业细则见 `references/formula-module.md`。

## 主流程（v4，2026-11-06 用户定版：样本先行）

> **第一条铁律：任何模块开跑的第一件事都是构造样本文件。**
> 用户原话（2026-11-06）："所有的模块测试，第一步都需要先构造样本文件。"
>
> 含义：**不许先写矩阵、先查文档、先跑探针再回头补样本**。样本是唯一真源——测试项从样本里长出来，结论拿样本回验，交付时客户能直接拿这份文件到对方官方站自己核对。
> 反例（都真实踩过）：①先按文档列 136 项探针表再补样本 → 探针行号越界、样本里没有载体，整批"假失败"；②打印模块先写矩阵后搭样本 → 矩阵一半项只能标"未测"，分母失真；③图表模块早期用库造样本 → DrawingML 细节缺失，被用户否掉重造。
>
> 与第 0 步的关系：**第 0 步是环境前置（开跑资格），不是业务步骤**——它不产出任何测试结论，只负责证明"跑出来的 FAIL 是真 FAIL"。业务意义上的第一步仍是样本。

### 第 0 步：环境门禁（Univer 侧跑任何用例之前，未过不许开跑）

> 2026-11-14 用户定版：**univer 相关问题验证，必须保证需要安装的包已经安装**。这条从"事后回溯"前移成"事前断言"——跑批前不过门禁，结果一律不可信，也不许记成 FAIL。

三条断言全绿才允许跑批：

1. **依赖核对**：本模块所需的 npm 包必须**可解析**，且是**功能包 + UI 包 + `/facade` 子路径**三件套（缺 UI 包不报错，只是列头不画漏斗/面板空白，截图会伪装成产品缺陷）。三种状态要分清，别一刀切：
   - **已声明**（写在 `package.json`）→ 安全。
   - **合法间接依赖**（没声明，但某个已声明包在 `dependencies`/`peerDependencies` 里依赖它）→ 安全，`npm i` 会跟着父包重装。**不要因为"没写进 package.json"就报警**——Univer 的 UI 包与 pro 包大量走这条路（`sheets-filter` 是 `sheets-filter-ui` 的依赖、`@univerjs-pro/sheets-outline` 是 `preset-sheets-advanced` 的依赖），误报会让人把门禁当噪音关掉。
   - **孤立安装**（既没声明、也没有任何已声明包依赖它）→ 危险，`npm i` 会当垃圾清掉。本项目被这样删过 `react` / `react-dom` / `@univerjs/sheets-hyper-link-ui`，随后 esbuild 报 `Could not resolve "react"`。动依赖前先备份 `package.json`。
   判定方式：从声明的包出发遍历 `node_modules` 依赖闭包，闭包外的才叫孤立——`scripts/preflight_univer.js` 已内置。
2. **插件自证**：`Object.keys(window.__proExtraPlugins).filter(k => window.__proExtraPlugins[k] !== 'ok')` 必须为空数组。**但要看值的内容**：`skip: ...duplicated plugin name...` / `already registered` 属于 preset 已注册、能力其实在，只记警告不阻塞（实测 22 个插件里 `UniverSheetsPrintPlugin` 就是这种情况，误判成缺件会白卡门禁）。其余 `skip` 才是真缺件。注册统一逐个 `try`，一个 already registered 不许中断后续（完整清单见 `references/workbench-architecture.md` §2.2）。
3. **facade 抽样断言**：本模块要用的 facade 方法必须 `typeof !== 'undefined'`。例：批注要 `sheet.getComments`、条件格式要 `newConditionalFormattingRule`、查找替换要 `createTextFinderAsync`、行分组要 `setRowGroupCollapsed` 类入口。

一键自检：`scripts/preflight_univer.js`（Node CLI，退出码非 0 即失败；`--json` 给跑批链读，`--url` 额外做页面内插件自证 + facade 断言，自签 HTTPS 已放行）：

```bash
node scripts/preflight_univer.js --root <工程根> --module comment,filter \
  --url https://127.0.0.1:8124/对比工程/index.html \
  --facade "univerAPI.getActiveWorkbook().getActiveSheet().getComments"
```

**为什么是硬门禁**：方法变 `undefined` 就是装配问题，不是能力缺失。历史上因此产生的误判至少 6 次——背景图、焦点高亮、查找替换、行分组、批注、图表，全部是自家工程缺件，却被写成对方缺陷。格式模块 Univer 从 95 掉到 63、表模块从 28 掉到 13，都是漏装造成的假 FAIL。

### 第 1 步：构造样本文件（所有模块共同的第一步）

**1.1 五条硬要求**

1. **必须由 Excel 本体产出**（COM 自动化）。库（openpyxl/python-pptx）只能做 COM 写不了的补充（条件格式规则细节、dxfs、批注 XML、x14 扩展），且最终必须用 Excel 打开核对。理由：用户只信任 Excel 本身的呈现效果，库造文件缺 DrawingML 细节。
2. **一项一载体**：测试计划里每个待测项，必须能在样本里指认到具体的单元格 / 图表 / 对象。指不到 → 要么删项，要么补样本。**禁止出现"表里有这项、样本里没这个东西"**。
3. **元素刻意铺满**：一次造数把该模块所有相关元素都铺上（图表多类型 / 形状 / 图片 / 批注 / 迷你图 / 条件格式 / 表对象 / 打印设置 / 分组 / 冻结 / 保护）。打印模块的教训：样本只有单元格时，25 项保真矩阵有 17 项没东西可核，只能标"未测"。
4. **造完必须三重核验**：①解包源 xlsx 数 XML 节点（chart/drawing/comment/conditionalFormatting 各多少个）②Excel 打开逐项读回（COM 或人工）③把核验数字写进样本说明。COM 静默失败率极高——参数一错位就全默认值，不核验的样本会污染后续全部结论。
5. **样本即交付物**：命名 `<模块>-<要素摘要>.xlsx`，放 `测试工程/samples/`，并在报告里写明"可直接拖进官方站核对"。一个模块 = 一份主样本 + 按批次的增补样本。
6. **装饰元素不许混进被测对象的容器**（2026-11-07 血亏教训）：注释用的文本框 / 说明条 / 标注箭头本身没问题，但它们会和被测对象共享同一个 XML 容器（图表的 `drawingN.xml`、批注区、表对象区）。实测：Excel `Shapes.AddTextbox` 写出来的 `xdr:sp` 一进 Univer，**同表 91 张图的富文本标题被整条丢掉**，我据此写下"Univer 丢 82/82 标题"的错误结论，直到做单变量对照实验（只删这 91 个文本框）才发现是自家样本的锅。
   - 因此：**交付样本可以带注释，测量必须用"去掉装饰元素"的派生样本**（`samples/<模块>-无标注.xlsx`），两者都留档；
   - 一旦得出"全丢 / 零例外 / 100%"这类过于干净的规律，先做单变量对照实验再写结论。

**1.2 各模块样本形态（同一套逻辑，不同载体）**

| 模块 | 样本形态 | 构造脚本 |
| :-- | :-- | :-- |
| 图表 | 63 种类型各一张 + 逐项显式样式簿（batch1/batch2） | `make_baseline_charts.ps1` / `make_batch1_styled.ps1` / `make_batch2_styled.ps1` |
| 公式 | 共享数据区 + 519 条用例表（三方同源，日期存序列号） | `data_block_v2.py` → `semantic_cases_v2.py` |
| 格式 / 条件格式 | 单簿 12 sheet，A～K 135 项，D 组带"同文本同列宽只改属性"对照组 | `make_style_cf_samples.ps1` + `patch_cf.py` |
| 表 / 排序 / 筛选 / 验证 | 单簿 6 sheet 对应 A～F 六组 | `build_samples_table.py` |
| 打印 / PDF | 单簿 3 sheet 多元素（图表 3 类 / 形状 3 类 / 图片 / 批注 / 迷你图 20 行 / CF 3 类 / 表对象 / 页眉页脚页码 / 重复标题行 / 横纵混排） | `build_print_sample.py` + `annotate_print_sample.ps1` |
| 工作表 / 工作簿 | 单簿 15 sheet，每 sheet 一个能力域；"要看见"的项配可见载体（分组、隐藏列、冻结线、标签色、填充柄…） | `build_samples_worksheet.py` + `annotate_worksheet_sample.ps1` + `verify_sample_worksheet.py` |

**1.3 图表模块造数细则**（其余模块的 COM 细节同理）

- 用本机 Excel COM 生成 63 种各一张的基线簿；配置项批次另造**逐项显式样式样本簿**（每图对应一个配置项组合，含 `object[,]` 数据写入范式）。
- 造数后**必须解包源 xlsx 核验**（数 chart XML + 抽查引用/标志位）。
- COM 细节见 pitfalls：曲面/股价图需先绑数据再改类型；新式图表需先选数据区再 AddChart2；`SetSourceData` 必须指定 PlotBy；PS 二维数组字面量传 COM 失败；PS 5.1 `Value2` 对 int/DateTime 抛 `InvalidCastException`（批量写值优先走 openpyxl）。

### 第 2 步：建立基准清单（从样本反推）
- 选定功能模块，以 Excel 对象模型为基准，把配置项**拆到颗粒度**（不是"支持饼图"，而是"饼图扇区分离""刻度标签旋转角度"这种粒度），分 10～12 组约 100 项。
- 图表模块直接复用 `references/dimensions-template.md`（A～L 必测 + M/N 可选扩展）。
- 每项记录：Excel 基线行为 + 参评方结论（**API 类模块一律 PASS/FAIL，只按公开 API 判**；图表类模块的"导入保真/仅界面可改"仍用 S1～S5，见 verdict-framework）+ 证据来源（文档查证/对象层实测/渲染层截图，**报告里用文字列标注，不用 emoji 小图标**——用户实战反馈：图标太多看不懂含义）+ API 名称 + 文档 URL。

### 第 3 步：收集官方功能清单反推验证项，产出全量测试计划
- **SpreadJS 官方功能清单已随 skill 分发**：`references/SpreadJS 功能清单（v19.2）.xlsx`（605 条 / 28 模块，含版本标记与场景权重），列结构与反推规则见 `references/spreadjs-feature-inventory.md`；提取用 `scripts/extract_feature_inventory.py`（`--module 05 --grep 格式,填充` 直接出 md 表格）。不要再向用户索要此文件。
- 其他参评方（Univer/Handsontable/ag-Grid）仍按需向用户索要或从官网功能页收集，用 sheet-agent 全量提取目标模块条目。
- **测试项来源必须双轨合并**：①官方功能清单（对方"有的"）+ ②Excel 对象模型基准（对方没提但 Excel 支持的）——只看功能清单会漏掉对方未宣传的能力，只看 Excel 基准会漏掉版本差异素材。
- 产出《<模块>配置项测试计划.md》交用户审阅后才开跑：每项标注**来源**（功能清单行号或版本标记 / Excel 基准）、**测试方法**（见下）、**已实测/待测**（之前轮次实测过的项标结论，增量测试不重复跑）、**批次归属**。
- 测试方法三分类：**①导入保真实测**（Excel COM 造显式样式样本 → 双端导入 → 内部模型+渲染截图+再导出 XML 三向验证）/ **②API 可编程对照**（两端代码设置同一配置验证生效）/ **③文档查证**（reference 层下探 + setter 探测，适用于 Excel 无法造数的交互/动态项）。
- 按"用户点名项 + 销售价值"排批次（如：批次 1 样式主战场 → 批次 2 系列与高级元素 → 批次 3 类型核对+交互查证）。清单中标注 V 新版本的能力项往往是版本差异的好素材。

### 第 4 步：文档查证填矩阵（并行双线）
- 并行派两个调研子代理各查一家官方文档/demo，逐项输出"项目→结论→API 名→证据 URL"；查不到标"文档未提及"，**禁止臆断为不支持**。
- **通道固定，不许混用**：SpreadJS 侧一律 `spreadjs-mcp`（`search`→`fetch`，功能命名搜法见关键原则 4）；Univer 侧一律 docs.univer.ai **reference 层**。网页搜索只作补充；MCP/文档查不到时转实测，不转"不支持"。
- **查证必须下探到 reference/类型定义层**（如 docs.univer.ai/reference/facade/axis-chart-builder），不能停在指南页——Univer 大量可编程能力深埋在 builder 层且零文档。探测脚本：`scripts/probe_setters.js`（运行时枚举 toBuilder() 原型方法并逐个试调用）。
- 若参评方有开源/商业分层（如 Univer），**固定只对比商用版本**（2026-11-04 用户定版，2026-11-05 再次确认为最高优先级基线）：SpreadJS 商业版 vs Univer Pro，开源版行为不进入结论矩阵（背景理解可用）。Pro 构建判定与环境约定见顶部"对比基线"节。

### 第 5 步：自动化实测 + 三向验证（最硬的证据）
- **SpreadJS 侧**：`scripts/runner.js` + `test.html`（npm 分模块装 **@grapecity-software/spread-sheets** / `-charts` / `-io` / `-shapes`——当前官方指向的 scope；旧 `@mescius/spreadjs` 整包已 404，`@mescius` 分模块包为海外渠道同源 dist，仅在 @grapecity-software 不可用时替代。**注意 @grapecity-software 版本可能滞后于 @mescius**（实战：前者最新 19.2.2、后者 19.2.5），以实际装到的版本号记录）。Playwright 无头导入，逐表统计图表存活 + 截图。API 校准细节见 pitfalls（v19 已移除 getCharts，用 sheet.charts.all()；import 需 FileType.excel；排除试用版授权 sheet）。配置还原读取：`chart.toJSON().shapeData.graphicFrame.graphic.graphicData.chart` 保留 OOXML 原样 chartSpace（提取脚本 `scripts/extract_sjs_models.py`）。
- **Univer 侧**：官方 playground 暴露 `window.univerAPI`。`scripts/univer_roundtrip.js`：API 建图 → exportSheetBySnapshotAsync 导出 → importSheetToSnapshotAsync 再导入 → getCharts() 统计 → 截图。匿名 demo 导出 unitId 路线会被拒，用 snapshot 路线。
- **Excel 对照**：Excel COM 打开同一文件统计（`$ws.ChartObjects().Count`），作为三方对照基准臂。
- **双栏对照台（代码级差异的最快呈现方式，2026-11-05 新建）**：`测试工程/index.html`（工程模板随工作空间保存，结构见 `references/dual-pane-harness.md`）。左 SpreadJS、右 Univer 两个真实实例，同一用例的"两侧代码"一键执行，异常即 FAIL——**天然满足"公开 API 调用成功才算支持"的口径**，且能肉眼看到设置是否符合预期。要点：①Univer 侧本地 OSS 无 xlsx 导入通道（Pro exchange 依赖转换后端），导入按钮只对 SpreadJS 生效，样本要另拖官方 playground；②SpreadJS 试用版会插 "Evaluation Version" 授权页且删不掉，需填 license；③两侧排版引擎不同，亚像素差异不算缺陷；④**宿主页样式必须与引擎 DOM 隔离**——Univer 会在容器内自己渲染 `<header>`/`<footer>`，页面写裸元素选择器 `header{}`/`footer{}` 会把它一起改掉（症状：面板顶部莫名黑条 + 画布只有 257px 高、下方大片空白，像"产品没渲染"），布局规则一律用 id 选择器，排查脚本 `scripts/layout_diag.js`。
- **三向验证**（判定"配置项支不支持"的标准动作）：Excel 造显式样式文件 → 对方导入读内部模型（Univer 用 chart.getInfo()）→ 截图验证渲染 → 再导出解包 XML 对比。脚本：`scripts/verify_config_preserve.js`。文档"未提及"≠不支持，此法曾把 Univer 一批"❌"翻案为全链路支持。
- **内部模型深挖（三向验证的第四臂，必做）**：Facade/API 层"查不到字段"≠"没存"——必须抓取导入后的内部模型再判。Univer 的图表真身在 `wb.save()` 快照 `resources` 通道 `SHEET_CHART_PLUGIN` 资源里（自有 schema：`style.xAxis/yAxis/titles/legend/seriesStyleMap`），Facade config 只是它的一个投影。实战教训：majorUnit 在 Facade config 无键（差点判 ❌），resources 模型里 `majorUnit:100` 完好保留。抓取脚本：`scripts/univer_resources_dig.js`。⚠️ 快照 resources 可能被截断，抓取时不要切片。
- **量化"看不见"**：用 size 分布（0×0 占比、几像素占比）代替主观描述。脚本：`scripts/measure_chart_size.js`。
- **类型支持必须"代码新建"实测**（用户要求，口径=代码添加，UI 面板插入不作为判定依据）：遍历参评方类型枚举逐个建图（SJS `charts.add` 61 数字枚举 / Univer `newChart(ChartTypeString)` + `insertChart` 41 字符串枚举），以"创建返回对象 + 类型读回一致"判可创建，**再配渲染抽查截图**——"创建成功"≠"渲染正确"（SJS 曲面 4 种渲染降级柱形、Univer 3D 系列大量空白）。脚本：`scripts/sjs_create_types*.js`、`scripts/univer_create_types.js`；API 签名与数据布局坑见 pitfalls（charts.add 参数顺序、combo 逐系列组合、股价图布局校验）。
- **导出方向实测**（与导入方向同等地位，**三层核对缺一不可**）：①SDK 导入样本 → 导出 xlsx → Excel COM 打开核对图表数；②解包导出文件 grep chart XML 抽查关键节点（prstDash/趋势线/误差线/pattFill/explosion 等）；③**Excel 实际展示逐图核对**——Excel COM 打开导出文件逐图 `Chart.Export`（3D 图 0 字节用 `CopyPicture→临时图 Paste→Export` 兜底，基线文件同样兜底），与源基线同索引拼三列对照图。实战教训：Univer 计数 19/19、XML 键齐全，但 7 张图导出后实际高度仅 5px——**计数存活 ≠ 展示可用**，尺寸塌缩会经内部模型固化进导出文件。**Facade 层无键 ≠ 导出 XML 不含**（Univer explosion 实战），"丢失"结论必须查导出 XML。脚本：`scripts/univer_export_batch1.js`、`scripts/export_excelview.ps1`、`scripts/export_baseline_fallback.ps1`。
- **闭环测试**（导出再导入）：数据层（对象+配置）与渲染层（画布截图）**分开下结论**，两者可能严重背离。

### 第 6 步：渲染闭环 + 评分报告
- **凡"看得见/看不见"的结论，用 Chrome 有头模式自动闭环**：playwright `chromium.launch({channel:'chrome', headless:false})` 弹出真实窗口导入并逐表截图——截图即人眼等价证据，**不需要再请用户人工复核**（历史教训是无头模式造成的，有头模式已根治；脚本 `scripts/univer_import_headed.js`）。注意：canvas 渲染的表格 SDK 无独立图表 DOM 节点，DOM 量测无效，只能靠截图判定。
- **报告结构固定走顶部「产出物清单」的四段式**（2026-11-14 用户定版）：**结论先行 → 导入导出对比 → 三方/功能截图 → API 完整度验证**。不再按"①图表类型支持 / ②配置项支持"两大主线分卷，也不再累积时间线文档。图表模块在这套骨架下的落位：类型支持（63 种基线还原率 + 授权分层）进第 1 段结论，B～L 配置项矩阵进第 4 段 API 验证。**报告开头必须设两个说明节**（用户实战要求）：①"支持度分级"节——先给 S1～S5 标尺定义再进结论（S1 代码可改有文档 / S2 代码可改无文档 / S3 仅界面可改 / S4 值在但无任何读写入口 / S5 丢失；无代码入口的项**必须单独标记，禁止直接给对号**——用户实战反馈："对开发人员来说没有代码设置是很大风险"）；②"名词解释"节——保真层/可编程层/对象层/渲染层/模型层保留/UI 面板等行话用大白话逐条解释，不解释读者（尤其销售/管理层）看不懂。
- **矩阵呈现规范（用户实战要求）**：全文禁用 emoji 小图标（📄🧪🖼✅⚠️❌ 等）；证据来源放表格最右"证据"列（文字：文档/实测/文档+实测）；每个配置项格子以 S 级别开头加一句事实描述。
- **证据截图直接嵌入报告正文**（用户实战要求，不要单独索引节）：截图嵌在对应结论章节内；Excel 基线列用 COM `Chart.Export` 导出（屏幕截图会黑块）；三列对照拼图每行加与矩阵分组对应的行标注；**图挂进报告前必须逐张读图核对"图与说明一致"**（两次翻车：对照图挂错章节、基线列黑块标注对不上）；保持文件名不变重制内容则引用自动生效。双端实测项数量不同时必须在汇总处写明计数口径（分母相同、差异=可编程覆盖面差距）。**"代码修改实测"的两侧证据图必须同口径镜像复刻**（用户第四次抓出图不对）：同一张图、同一语义修改、同一修改方向、同一构图（逐图导出 2×2 + 相同标注），两侧脚本一一对应，严禁一边整表截图一边逐图拼图；详见 pitfalls"两侧同口径"节。
- 结论按"我方强项 / 对方短板 / 双方都弱 / 需避开战场"四类整理，每条配销售话术，话术措辞遵循 `references/verdict-framework.md`。**不记版本演进、不留修正轨迹**（见关键原则 16 ④）。

## 公式函数模块：支持度 / 准确度两个分母（v3，2026-11-07 定稿）

**作业细则见 `references/formula-module.md`**（尺子标定 / 六步管线 / 特殊函数 / 性能口径 / 话术 / 自检六条）。这里只留口径与定稿结论。

**口径**（用户 2026-11-07 定版：只判"算出来的结果和 Excel 一样不一样"，不做精度误差分析、不做位数归因）：
- **两个分母不混算**：支持度 ÷ **504**（519 个函数 − 15 条无 oracle；`#NAME?`、写入被拒 = 不支持）；准确度 ÷（支持数 − 5 易失 − `IMAGE`）= **488 / 485**。`#NAME?` 只扣支持度，不进准确度分母。
- **能测的必须进分母**：只要有 Excel 标量真值，哪怕要另造样本（透视表、多工作表）也算进 504——客户会问"这几个为什么排除"，只有"本机给不出 oracle"才够格排除。
- **数值按 Excel 自身的 15 位有效数字归一后逐字比，文本（含复数）原样逐字比，不设相对容差**。这是"位数归一"不是放水——15 位以内差一位就判不一致。
- **跑之前先把 SpreadJS 拉到同一精度**：`GC.Spread.CalcEngine.Functions.CALC_PRECISION = 15`，静态属性，必须在 `new Workbook` **之前**设；对比台与跑批默认都开（`?prec=default` 可退回复现旧数据）。Univer 无此开关，比较时同归一到 15 位。
- **只算支持度**：易失 5 条（`NOW`/`TODAY`/`RAND`/`RANDBETWEEN`/`RANDARRAY`）、`IMAGE`（Excel 本机也返回错误码 `0x800A0803`，没有标量真值）。**环境依赖 15 条**（`CALL`/`RTD`/`WEBSERVICE`/CUBE 全家等）Excel 给不出 oracle，不进主表。
- **单独搭场景 4 条（7 条用例）计入 504，不要另立分母**：一个函数一票——SJS 支持 4/4、准确 3/3；Univer 支持 3/4、准确 2/2。既不能只报支持度（比过的结果扔掉），也不能把它们排除在总分母外（客户会问凭什么）。
- 动态数组一律开着测（SJS `allowDynamicArray=true`，Univer 无此开关），关掉不是能力差异。

**尺子必须先标定**（本轮最大的方法修正）：按"double 严格逐位相同"比，SJS 281/484 = 58.06%、Univer 289/482 = 59.96%——**精度差的那家反而更高**，因为两家都存不下 Excel 的 17 位，`=PI()` 在两边都被判"不一样"。标定三步：①`=PI()` / `=1+1` 两家都要过；②人为差 1e-3 的对照两家都要不过；③排名反直觉就先怀疑尺子。同理，1e-9 相对容差会把 `BESSEL*`（差 1e-8）和复数 26 条全放过，必须逐字比、再按误差幅度分档。

**六步管线**：①共享数据区 + 每函数一条常规用法用例（数组函数不套 `INDEX`，让它自然溢出）②Excel oracle 先行（500/500 必须全出值，Excel 自己报错的用例当场修或剔除）③双端同数据区跑批 ④Python 判定与 xlsx 活公式判定逐行对账 0 不一致 ⑤特殊函数单独搭场景 ⑥性能与导出（导出只比 `<f>` 节点，不用各家 `getFormula()`）。

**定稿数字**（SpreadJS 19.2 / Univer Pro 1.3.1 / 500 条 / Excel 365 基线）：

| 指标 | SpreadJS | Univer Pro |
| :-- | :--: | :--: |
| 支持度 | **494/504 = 98.02%** | **491/504 = 97.42%** |
| 准确度 | **448/488 = 91.80%** | **324/485 = 66.80%** |
| 一次跑通 | 448/504 = 88.89% | 324/504 = 64.29% |
| 多值溢出 | 34/36 = 94.44% | 29/36 = 80.56% |

差异幅度要一起说（差 **25.0 pp**）：**两家的失分没有一条超过 1e-3**，SJS 37 条数值里 27 条在 1e-12 以内，Univer 157 条里 104 条在 1e-12 以内、52 条在 1e-9 以内——没有量级错误。Univer 低分集中在统计（43.93%）、财务（46.30%）、兼容性（32.50%）三个数值族，根子是它写回单元格的位数比 Excel 少；SJS 现在最低的是财务与兼容性（80.00%），工程族设 15 位精度后 66.67% → 88.89%（`BESSEL*` 仍差 1e-8）。复数那一族最严：SJS 24/26 = 92.31%、Univer 23/26 = 88.46%。

**特殊函数**（4 条单独搭场景，7 条用例，**已计入 504**）：5 条 Excel 有标量真值（`SHEET`/`SHEETS` 3 条 + `GETPIVOTDATA` 2 条），两家全对；差别只在支不支持（SJS 4/4、Univer 3/4）。
- `SHEET` / `SHEETS`：两家都是 2 / 2 / 1，全对。前提是工作簿只有两张表，多一张序号整体后移。
- `GETPIVOTDATA`：**SJS 2/2（500 / 1520）**；**Univer 0/2，两条都 `#NAME?`**。SJS 侧要装 `@grapecity-software/spread-sheets-pivot-addon` 且**版本与主包完全一致**（本轮 19.2.2）、加载顺序 **core → pivot → designer**，不装就是 `#REF!`——构建缺件，不是函数不支持。
- `IMAGE`：本组唯一测不了结果的，只计支持度（两家 2/2）；SJS 接住为图片对象（文本 `Picture` / 替代文本），Univer 返回空，Excel 本机给错误码 `0x800A0803`。
- **报告正文只写函数口径**（能不能算、结果对不对），透视表插件、透视表导入渲染这些工程细节留在工程文档和本 skill 的 pitfalls，不写进客户报告章节。

**性能（本地对比工程实测）**：43,263 条公式 / 10,000 行。写入 1,038 / 2,434 ms（2.34×，SJS 更快，前提是 `suspendEvent` + `suspendPaint` + `suspendCalcService` 三样一起挂起）；**全量重算 2,704 / 13,195 ms = 4.88×**；改一个输入后重算 3.51×（多轮波动 4.88–5.11×、3.51–4.61×，对外讲"约 5 倍"）。五条纪律：**每轮新开页面**（同页连跑 Univer 11 s→22 s）、各跑 2 轮**取中位**（6 次极差占中位 SJS 45% 含首轮冷启动、Univer 19%）、**批量写公式要挂事件**（不挂事件每条 `setFormula` 都通知 Ribbon，写入 1.3 s 会变 3.2 s，那是写法成本不是引擎能力）、**条数按实际写公式的列算**（`Data` 5 列数据只写 4 列公式）、**9 个抽样格双端逐字相同**才认。

**计算控制机制对照（回答"Univer 有没有挂起计算"）**：Univer **有初始化期挂起、无运行时任意挂起**——
- `FFormula.setInitialFormulaComputing(CalculationMode.NO_CALCULATION=2)`（**建簿前**设置）：建簿/导入数据时不计算，之后 `executeCalculation()` 一次全量算出；是 SJS `suspendCalcService + resumeCalcService` 的部分等价物。`FORCED=0` 强制全算、`WHEN_EMPTY=1` 默认。
- `FFormula.stopCalculation()`：中断当前计算，不是暂停后恢复。
- 边界：NO_CALCULATION 对建簿后**新写**的公式无效（写入立即算出），"运行时批量写公式再统一算"没有对等开关，只能 `executeCalculation()` 收口。对外说"初始化期可挂起 + 全量重算"，不说"Univer 没有挂起计算 API"。脚本：`scripts/univer_suspend_check2.js` / `scripts/perf_univer_bulk.js`。

## 格式/条件格式/单元格层模块：API 支持度与友好性探针（v1，2026-11-05 用户定版）

**口径**（用户原话："你不需要对比浏览器导入导出效果，这部分改为我自对比。但是你需要检查 API 支持情况和友好性"）：**导入导出效果客户自测，不进本管线**。本管线只回答四问——**有没有代码入口？能不能设？能不能读回？要绕几道路？** 外加"用起来顺不顺手"。完整方法论见 `references/api-probe-method.md`，脚本三件套 `scripts/api_probe_table.js` / `api_run_probe.js` / `api_gen_report.py`。

五步管线：
1. **建探针表**（`api_probe_table.js`）：以 Excel 对象模型拆项（实战 A～K 12 组 136 项），每项一条 spec：`{id, name, sjs:{api,kind,steps,param,doc,set,read,exp}, uni:{...}}`。判定字段统一为 **set 抛错？读回值 == exp？doc 有无？**——这三条就能机械化出 S 级，不靠主观。
2. **双端跑批**（`api_run_probe.js`）：Playwright 起 localhost（端口固定），宿主页注入 `window.ApiProbe.runAll('sjs'|'uni')`，结果落 `out/api-items.json`。**跑批前必须先扩行数**（`setRowCount(400)`，SJS 新版默认仅 200 行，探针行号越界会造出整批"假失败"）。
3. **洗假失败（最关键的一步）**：任何 FAIL 先当作"我的写法错"排查，顺序是 **①查官方文档示例 ②枚举运行时命名空间 ③用最小探针单独验证**，确认引擎缺陷后才留档。实战 25+ 处假失败全部定位归零，最大一条：
   > **运行时没有某个方法 ≠ 没有这个能力，能力可能挂在数据结构的形状上。**
   > `typeof sheet.setRichText === 'undefined'` → 一度判"富文本 S3"；查文档后发现 SpreadJS 富文本是**值**——`setValue(r,c,{richText:[{text,style:{vertAlign:1}}]})`，读回要 `getValue(r,c,sheetArea,ValueType.richText)`。三项（上标/下标/富文本）全部改判 S1。
4. **机械定级**（`api_gen_report.py` 的 `grade()`）：`skipped→—` / `noApi→S3 或 tierHint` / `setOk===false→S3` / 否则 `S1`(有文档) 或 `S2`(无文档)，读不回再加 `⁻`。**覆盖率分母必须剔除 S- 项**（Excel 本身没有的能力，实测 9 项），否则两侧同比分母不一致。
5. **五维友好性评分**（各 0～5，合计 25）：覆盖广度（分母剔 S-）/ 一步到位（1 步 5｜2 步 3.8｜3 步 2.4｜≥4 步 1.2）/ 可读回 / 参数类型化（typed 5｜enum 4｜loose 2）/ 错误反馈（**5 − 0.5×静默失败条数**）。静默失败单列成表，这是"友好性"里最有销售杀伤力的一节——调用不报错但不生效，比报错更难排查。

本轮实测结果（SpreadJS 19.2.5 vs Univer Pro 1.3.1；本管线口径 136 项 / 123 项有效分母）：**SJS 121 项可编程（98.4%）、Univer 94 项（76.4%）**；五维 20.5 vs 19.2；SJS 独有 9 项中 Univer 覆盖 0 项。Univer 侧 38 项"无公开入口"，集中在条件格式规则细节、单元格状态、自动合并等。
**口径更新（2026-11-06，对比台复跑）**：把 9 项"Excel 本身没有的能力"从"记 Univer FAIL"改为 NA 移出分母后，同一套 SPEC 的结论是 **137 项 / 124 可判 / SJS 122（98.4%）/ Univer 95（76.6%）**，SJS 只剩 `A6` 下划线颜色、`B7` 射线渐变两条 FAIL（这两条 Univer 也没有）。**对外用 122/124 这一组**，五维评分仍可沿用（它是另一套评价轴，不影响覆盖率）。

## 表 / 排序 / 筛选 / 数据验证模块：46 条用例双端跑批（v1，2026-11-05 晚用户定版）

**口径**：继续用 PASS / FAIL 二值，**但加两条硬判据**——"接口返回成功但数据没动"判 FAIL；"读回值 != Excel 基线"即使 PASS 也要单列行为差异。

五步管线（脚本全部入 `scripts/`）：
1. **工程前置**：Univer 侧按"功能包 + UI 包"两层装齐并重建 bundle（`build_bundle.js` + `univer_bundle_entry.js`）。缺 UI 包不会报错，只会让列头不画筛选漏斗，截图会伪装成产品缺陷。
2. **能力面探针**（`probe_table_gaps.js`）：对每个"疑似不支持"的项，先查**类型定义/枚举/命令表**，再真实调用。本轮靠它翻案 5 个假 FAIL（颜色筛选×2、前N项、排序命令、汇总行字段）。
3. **造样本**（`build_samples_table.py`）：6 sheet 对应 A～F 六组，用真 Excel 打开核对（表对象属性、结构化引用实算、Validation.Type）。
4. **用例库 + 跑批**（`cases_table.js` 46 条 + `run_cases_table.js`）：用例 id 形如 `T-A1`，批量器按组过滤；写法铁律"先算先读、最后才 throw，观测值拼进错误消息"。
5. **证据图**（`shoot_table_showcase.js` → `compose_table_evidence.py`）：同一套演示代码两侧各跑一遍，截面板图拼左右，附三行差异说明。

本轮结果（SJS 19.2.2 vs **Univer Pro 1.3.1**，46 项）：**SJS 46/46（100%）、Univer Pro 28/46（60.9%）**；缺口集中在排序（0/8，含"命令空转"）、表样式四开关、汇总行、前N项筛选、文本长度与输入提示。Univer 赢的一项：表内计算列有真正的列级 API。

## 打印 / PDF 导出模块：字段层 + 导入层 + 输出层 + 元素层（v4，2026-11-10 定版：模块已完成）

**作业细则见 `references/print-pdf-module.md`**（四层判据 / 样本必备元素 / 35 项字段分组 / 21 项元素矩阵与三档判定 / 五个易被反问的结论写法 / 七条自检 / 复现命令）。这里只留口径与定稿结论。

**口径（2026-11-06 用户修订）**：
- **一张表 `# | 测试项 | SpreadJS | Univer | 说明`**，两侧对位在同一行，**不要拆成"SJS 一张表 + Univer 一张表"**。
- **PASS = set 与 get 同时存在且读回一致**；**FAIL = 缺 set / 缺 get / 读回≠设置**，三种都记 FAIL，且**每条 FAIL 必须在说明列写原因**。
- **对方没有的一律写 FAIL**，不写"无此元素""暂不支持"等中间态（它会把分母做小、分数抬高）。
- **只用公开 API**：不用 injector、不碰内部字段；读回走 `printInfo`（SJS）/ 事件载荷（Univer）。
- **四层分开写，不许混**：字段层（能不能设）、导入层（模板导进来还在不在）、输出层（纸面一样吗）、元素层（纸面出现哪些元素、像不像 Excel）。**字段层 100% 不代表输出一致**；**35 项与 21 项的分母不可混算**。

四步管线（脚本全部入 `scripts/`）：
1. **多元素样本**（`build_print_sample.py` + `annotate_print_sample.ps1`）：openpyxl 造数据/样式/条件格式/批注/超链接/行分组/表格/打印设置，COM 补形状与迷你图并导出 Excel 基线 PDF。**元素必须铺满**——图表 3 类、形状 3 类、图片、批注、迷你图 20 行，否则保真矩阵没东西可核。样本的打印区域必须**明显小于数据区**，列宽必须**窄到会跨页**。
2. **SJS 导出**（`print_sjs.js`）：导入样本 → **注册中文字体** → 逐表 `savePDF` → 再逐页转 PNG 供比对。缺 charts/shapes/print/pdf 任一插件都会静默丢元素，先断言再统计。
3. **Univer 打印**（`univer_print_capture.js`）：`updatePrintConfig` + `updatePrintRenderConfig` → `print()` → 轮询抓 `.printing-canvas canvas` 存"打印预览"。
4. **字段矩阵**（`print_api_matrix.js` / Pro 侧 `probe_pro_print_matrix.js`）：35 项统一用例，逐项 `set → readback → compare`；Univer 的读回走 `BeforeSheetPrintConfirm` 事件载荷（getter 全 undefined，见 pitfalls）。最后 `compose_print_evidence.py` 拼三栏证据图。

**定稿结果（SJS 19.2.2 vs Univer Pro 1.3.1，同一 35 项分母）**：字段层 **SJS 35/35（100%）**——`printInfo().watermark()` 有 getter，能读回完整对象；**Univer Pro 19/35（54.3%）**，16 条 FAIL 分三档 = **11 条字段不存在**（外边框、黑白、页序、页码范围、起始页码、首页不同、奇偶页不同、页眉图片、最佳列宽、最佳行高、打印质量）+ **1 条字段在但功能坏**（`pageSizeCustom` 传 500×700 读回 394×588，固定减 106/112，连续打印累减到 −3556）+ **4 条字段在、读回也对但打不出来**（打印区域四角 P1～P4：`area=AllSelection` 时打印预览 0 页，只有 `CurrentSheet` / `Workbook` 能出页）。Univer 反向独有 5 项：`area:workbook`、`subUnitIds[]`、`maxRowsEachPage`、`headerFooter[]`、`PrintHeaderFooterSymbol` 的 15 个占位符。

三组同判 PASS 但**语义不等价**，讲的时候必须带边界：

| 组 | 差异 |
| :-- | :-- |
| 重复标题 P5～P8 | Univer 走**工作表冻结**：`setFrozenRows(n)` / `setFrozenColumns(m)`。冻结区只能从第 1 行/列起；打印重复与编辑器冻结窗格是同一状态（Excel `Print_Titles` 可独立设） |
| 适配页宽/页高 | 只有 `FitWidth/FitHeight` 枚举，不能指定"占几页宽" |
| 缩放比 | SJS 收小数 0.8，Univer 收百分比 80 |

P1～P4 判 FAIL 不是"设不进去"，而是**设进去也打不出来**（预览 0 页）；两家的模型也不同：SJS 是**具名固定区域**（存在 `Print_Area` 自定义名称里，插删行列跟着变、公式可引用），Univer 是 `area` 枚举 + `subUnitIds[].range` 的**选区列表**，没有"与当前选中无关的具名区域"这一档。

**导入层与输出层**（Excel 为 oracle）：

- **页面设置导入**：SpreadJS **8/8**；Univer **2/8**（只对纸张、水平居中），方向/缩放/页边距/页眉页脚/打印网格线全部回落默认 → 同一张表 Excel 1 页、Univer 2 页。
- **Univer 的页面设置是纯运行时状态**：改方向后导出，XML 仍是 `landscape`（SJS 会变 `portrait`）→ 既不读文件也不写文件。话术只能落在"导进来要重设、改完存回去也不带走"，不能说"会把设置弄丢"（三家导出的文件都完整保留 `pageSetup`/`headerFooter`/`_xlnm.Print_Titles`）。
- **PDF 输出**：页数（4）、方向、顺序一致，尺寸差 1pt；三处真差异：页脚 `&F` 不输出、`&D` 日期格式不同、长表分页断点差 3 行。
- **两个必做配置**：导出中文 PDF 前必须 `PDFFontsManager.registerFont`；PDF 结论必须用 Chromium 二次复核（pymupdf 对子集字体会假性乱码），两边不一致时先查字体注册。
- **元素级保真必须逐元素一行**，且不能只判"有/无"——见下节 21 项矩阵。

### 元素层：21 项支持矩阵（2026-11-10 新增，样本 `打印元素矩阵-样本.xlsx`）

三档判据：`一致`（元素在、同页同区、相对几何成立）/ `降级`（元素在但有可指名差异）/ `丢失`（没有该元素，或只剩空框）。**只给单一"支持率"会把空框和"箭头丢了"算成同一档。**

| 分组 | 项 | SpreadJS 一致/降级/丢失 | Univer Pro 一致/降级/丢失 |
| :-- | :-: | :-- | :-- |
| 图表（柱/折/饼/散） | 4 | 4 / 0 / 0 | **0 / 0 / 4** |
| 图片与形状 | 9 | 8 / 0 / 1 | 6 / 2 / 1 |
| 表格结构 | 4 | 4 / 0 / 0 | 2 / 1 / 1 |
| 条件格式与错误值 | 4 | 3 / 0 / 1 | 3 / 0 / 1 |
| **合计** | **21** | **19 / 0 / 2＝90.5%** | **11 / 3 / 7＝52.4%**（可用 14/21＝66.7%） |

必须带上的五条：

1. **图表是最大断层**：柱/折/饼/散四类图的打印输出只剩圆角空框。同文件在编辑器里正常渲染，等 15s 再截仍空 → 缺口在**打印渲染**段（导入、编辑器、文件、配置四段都正常）。
2. **"降级"这一档不能省**：Univer 连接线丢箭头、圆柱正文被压到边缘；这些只判"有/无"会全部漏掉。
3. **单元格内图片三家都打得出**（7 张，含透明通道印章）。与"Univer 只能看不能建"不矛盾：导入能显示，打印就能出。
4. **批注清单与错误值打印方式两家都丢，但归因不同**：SJS 能导入批注但 `PrintInfo` 无 `printComments`（能力缺失）；Univer 传统 Note 导入即丢且无字段（导入通道 + 能力）。不合并写。
5. **手动分页符只有 Univer 丢**：它的打印配置只有 `maxRowsEachPage` / `maxColumnsEachPage`，没有分页符对象。

样本设计硬约束：**一份文件七张表、每表一类元素、每表控制在一页内**（唯一例外是故意留 2 页测分页符）。Univer 不读 `printArea`/`fitToPages`，跨页会把对象切断，切断会被误读成"元素丢失"。

**工程落点（元素层与字段层的最大区别）**：元素层判不了自动化——"降级"必须人看图。所以它不做 `run()` 用例，走「判定表 + 复核页」两件套：

- 判定表 `out/data/print-matrix-items.json` 是**唯一数据源**，改判定只改它；
- `build_print_matrix_verdict.py` 出 `print-matrix-verdict.json` 与前端清单 `js/cases/print-elements.js`（生成物，勿手改）；
- 复核页 `打印元素矩阵.html`，入口在对比台顶栏「元素矩阵 21 项 ⇗」，只在打印模块出现；
- 一条命令重建：`tools/rebuild_print_matrix.ps1`（`-SkipSample` 只重算判定与出图）；
- 跑批文件 `print-cases.json` 里 `elements` 字段**只挂不并**——35 与 21 两个分母各算各的。

**别把它塞进字段层那张用例表**：那是 PASS/FAIL 两档 + 耗时列，三档值进去会被 `verdict()` 渲染成「—」，混排还会让人以为是一个分母。细则见 `references/print-pdf-module.md` §5.6。

## 工作表 / 工作簿模块：127 项 API 层双端跑批（v1，2026-11-06）

**口径**：纯 API 层，只判「设置 → 读回」，**不判像素、不追截图**。有公开 API 且调用成功 = 支持；标了 `strict` 的项还要求读回值达成语义（"入口在但实测无效"判未达成）。视觉层交给 15 sheet 样本走人工复核。

与前几个模块的三处结构性不同（换模块时最容易翻车的点）：

1. **每一项独占一张临时表**（`__p<ID>`），跑完即删。因为这个模块要增删表、隐藏、保护、改行列数——共用一张表会互相污染：实测出现过前面某项把 `selectionPolicy` 改成行模式，导致后面 `getSelections()[0]` 读成 `3,-1`。
2. **必须有单项超时保护**。Univer 的 `protect()` / `createTextFinderAsync()` 在缺少对应服务时返回**永不 resolve 的 Promise**，首轮整批直接 SIGTERM、无任何输出。解法是 `withTimeout(runSide(item,'sjs'), 8000)` + 页面内 `__wsProgress` 逐项打点，挂住时立刻能看出卡在哪个 ID。
3. **`strict` 与非 `strict` 两轨分开统计**。配置类项（保护 20 个点位、冻结、缩放）只判"调用成功"；语义类项（查找、选区、末行末列、已用区域）判"读回是否达成"。混算会把分母搞乱。

五步管线（脚本全部入 `scripts/`）：

1. **样本先行**：`build_samples_worksheet.py`（15 sheet）+ `annotate_worksheet_sample.ps1`（COM 补背景图/批注 + 23 项回读断言）+ `verify_sample_worksheet.py`（解包 43 项包级断言）。样本是效果层人工复核的载体，API 层跑批不读它。
2. **探针表** `api_probe_worksheet.js`：A 表管理 12 / B 可见性 6 / C 保护 29 / D 冻结 8 / E 分组 8 / F 行列 17 / G 视图 10 / H 查找 9 / I 序列化 6 / J 事件命令撤销 9 / K 剪贴板 5 / L 数据区域 8 = **127 项**。每项 `{id,g,name,excel,sjs,uni}`，`sjs|uni` 内是 `{api,set,read,exp,strict,assert}`。
3. **双端跑批** `api_run_probe_worksheet.js`：起 localhost:8462 + Playwright(msedge)，出 `out/api-ws.json`，四档统计 **pass / noApi / err / strictFail**。跑单组：`node tools/run_probe_ws.js C`。
4. **三轮自检洗假失败** `api_ws_diag1/2/3.js`：首轮 47 秒报出 19+10 处 mismatch，**大半是我把 API 写错了**——不洗就写报告，会把 7 项"产品有"判成"产品没有"。
5. **线上 Pro 补测**（本地 OSS 缺的插件）：`univer_pro_outline_probe.js`（E1～E6）、`univer_pro_caps_probe.js`（10 个示例页逐页探入口）、`univer_pro_surface_dump.js`（抓 Facade 完整方法面）、`univer_pro_permission_probe.js`（保护模型方法与 42 个权限点位）、`univer_pro_outline_evt.js`（事件枚举 + 嵌套深度）。

本轮结果（SJS 19.2.2 vs Univer OSS 1.3.1 + Pro 1.3.1 补测）：**SJS 118/127（92.9%）、8 项无 API、1 项未达成**；**Univer 94/127（74.0%）、33 项无 API**（其中 4 项行分组在 Pro 有，实缺 29）。交集 **86 项双方都能干**，仅 SJS 32 项，仅 Univer 8 项，**双方都不会 1 项（正则查找）**。SJS 的 32 项独有集中在三处：保护策略完整性（密码 + 7 个 Excel 原生开关）、大纲完整形态（汇总方向/树形列/显示开关/事件）、显示层开关（显示公式/零值/选区配色/行列头 visible）。

## 图表模块：103 项 API 层双端跑批（v1，2026-11-06）

**口径**：与打印模块同一套——**PASS = set + get 同时存在且读回一致**；缺 set / 缺 get / 读回被改写一律 FAIL，说明列写原因；**对方没有的一律写 FAIL**；只用公开 API。两条新增：①**判"没有"必须能落到类型定义**（`.d.ts` 里确实没有该字段），否则只能算"我没找到"；②**判 PASS 看能力是否可达成，不看方法名是否和对方一样**——`colorScheme()` 能做主题切换与自定义调色板，就不因为不叫 `theme`/`palette` 判 FAIL（这是用户看图质疑"标了 ✕ 却写有等价能力"后定下的规则）。

四步管线（脚本全部入 `scripts/`）：

1. **探针表**：SJS 侧 `api-probe-chart-sjs.js`，Univer 侧 `probe_pro_chart_api.js`，**两侧共用同一套 A～L 103 条用例 ID**（A 对象 11 / B 标题 9 / C 轴 14 / D 刻度 8 / E 网格线 5 / F 图例 7 / G 数据标签 10 / H 系列 15 / I 高级元素 8 / J 图表区绘图区 7 / K 3D 3 / L 其它 6）。数据样本统一 `A1:D7`（1 类目 + 3 数值），**每项独立新建一张图表**。
2. **SJS 宿主页单独开**（`chart-probe.html`）：只加载 all + shapes + charts 三件，**顺序不可换**（charts 引用 `Shapes.ShapeBaseData`）。不要塞进主对照台。
3. **双端跑批**：`run_probe_chart_sjs.js`（本地 8466 + headless，出 `out/chart-api-sjs.json`）／`probe_pro_chart_api.js`（线上 playground，`getInfo()` 读回，出 `out/pro-chart/chart-api-uni.json`）。
4. **洗假失败（本模块最要紧的一步）**：首轮 Univer 报 **47 条 FAIL，其中 25 条是我写错 / 环境没就绪**，收敛到 **22 条**；SJS 侧另有 3 条同类错误。六类自伤：①新表灌数据用 `sheet.command.set-range-values` 静默失败 → 20 条集体假 FAIL；②饼图 builder 没有 `setSeries`（应 `setPieLabel`）；③拿 2D builder 判断 3D 的 `setThreeD`；④位置变更走 `toBuilder().update()`；⑤`exportImage()` 需「图表所在表为活动表 + 首帧已渲染」，否则永远 undefined；⑥**入参类型错**（`colorScheme(1/2/3)` 传数字、`setColorScheme()` 调不存在的方法）——不抛错、不生效、读回 undefined，最隐蔽。诊断脚本 `_diag_pro_chart2～9.js`、`_diag_theme_palette.js`、`diag_color_scheme.js`。

本轮结果（SJS 19.2.2 vs Univer Pro 1.3.1，103 项）：**SJS 88/103（85.4%）**、**Univer 81/103（78.6%）**；双方都过 69、仅 SJS 19、仅 Univer 12、双方都不过 3（zIndex 层级 / 误差线 / 图下数据表）。Univer 的 22 条 FAIL 结构：14 条字段不存在 + 6 条方法不存在 + 1 条有 set 读不回（`setZOrder`）+ 1 条字段被丢弃（`label.family`）。**缺口方向完全不同**：SJS 丢在坐标轴刻度密度与绘图区样式（majorUnit/minorUnit/crossPoint/次网格线/轴显隐/绘图区填充），Univer 丢在标题背景、图例背景与边框、数据标签背景与分隔符、轴标题旋转。Univer 反向独有 12 项（标题显隐/加粗/位置、副标题、轴标题加粗、网格线线型、图例加粗、扇区分离、标线、轴指示器、3D 视角与深度）。**主题 / 调色板两侧都有**：SJS `colorScheme()`（17 套内置 + 自定义数组，但改方案不刷新已有系列）、Univer `setTheme()`（13 套命名主题 + registerTheme）/ `setPalette()`（即时生效）。

## 数据透视表模块：71 条双端跑批（v1，2026-11-09 定稿）

**口径**：PASS = **能读回**（字段布局 / 汇总方式 / 值显示方式 / 筛选 / 折叠状态可从对象模型读出且与 Excel 落盘一致）**+ 能对上**（按 Excel 网格锚点 `TableRange2` 逐格比对显示值）。结构读回对、渲染不对，记 FAIL。导入 / 导出统一**双口径**：**对象数 + 逐格值**（本样本 1067 格）——只报一个方向必错（值全对但对象少了，或对象在但值错）。

样本与期望值（`对比工程/samples/数据透视表-全量用例.xlsx` + `数据透视表-全量用例-layout.json`）：

- 21 个工作表 = 3 张源数据 + 17 张透视用例表 + 1 张 `GETPIVOTDATA` 表；共 **22 张透视表、2 个切片器、1 张透视图**。
- **全部由 Excel 本体建**（`PivotTableWizard` + `PivotFields` 摆字段、菜单设值显示方式、界面插切片器与透视图），不手写 XML——否则 `pivotShowAs`、切片器关系、`pivotSource` 都不是 Excel 原生写法，测出来的"保真"没有意义。
- 期望值 `js/pivot-expect.js` 用 COM 读 `TableRange2` 整块（含页字段行）逐格固化，**任何引擎自算值都不作基准**。

五步管线（脚本在对比工程 `tools/`）：

1. **造样本 + 抽期望值**（COM）。字段布局要覆盖：两级嵌套行、两级嵌套列、多值横排、值竖排、页字段、紧凑 / 表格型。
2. **双端跑批** `run_cases.js pivot`：10 组 71 条（P-A 对象与数据源 8 / P-B 字段布局 8 / P-C 聚合与值显示方式 17 / P-D 分组 6 / P-E 排序筛选折叠 7 / P-F 计算字段 5 / P-G 刷新改源 4 / P-H 透视图 4 / P-I 切片器 5 / P-J 函数·文案·文件级 7），出 `out/data/pivot-cases.json`（报告数字唯一来源）。
3. **「新建表 vs 导入表」对照**：API 在新建表可用、在导入表静默不落地（`setLabelManualFilter` / `setCellCollapse` / `updateSourceRange` / `setNumberGroup` / `setDateGroupingMode` 五项），结论只能写「**导入表上不落地**」，不能写「无此能力」——所以每条都要在新建表上补一次对照（7-9 ms 返回 vs 导入表 6-25 s 不返回）。
4. **语言基线**：`CultureManager.culture('zh-cn')` **+** `gc.spread.sheets.resources.zh.min.js`（排在 pivot 插件之后），两者缺一面板与总计标签全英文；Univer 侧走中文语言包。三家中文化后再比对文案（见 pitfalls 59/60）。
5. **导入导出三方取证**：引擎侧按表截 `shoot_pivot_evidence.js`（SJS 画布 / Univer 画布）、Excel 侧 `shoot_excel_evidence.py`（COM `PrintWindow` 后台抓窗口）、`compose_pivot_evidence.py` 拼三联图。**每个不支持项都要有「三家界面上的样子」**，不能只给文字。

本轮结果：**SJS 62/71（87.3%）、Univer 52/71（73.2%）**。两家缺口方向几乎相反：

- **SJS 9 条 FAIL**：**没有透视图对象类型**（P-H 0/4，导入的透视图只剩普通图 `chartType=12`）、值显示方式两项有条件（`percentOfParent` 单层行字段整列空白、导入未还原基本字段）、数值分组只能读不能建、值标签语序与 Excel 相反（`汇总 金额合计` vs `金额合计汇总`）。强项：22 张表全部读回、导出**结构级保真**（切片器 2/2、图表对象 1/1，计算字段表不退化）。
- **Univer 19 条 FAIL**：**计算字段 / 计算项整块不进模型**（P-F 0/5，导入后退化成静态值）、切片器能渲染但不稳定（对象通道可用、UI 通道取选项抛 `field-missing`）、标签筛选与自定义序列排序未还原、`GETPIVOTDATA` 无实现。反向强项：值显示方式 17/17 全过、日期分组三级读回、透视图 4/4——其中透视图与值显示方式恰是 SJS 的失分处。导出值保真但**结构塌陷**（切片器 2→0、透视图 1→0、计算字段表丢对象）。
- 导入侧读数：对象 22 / 22 / 21；逐格值 1067 / 1067 / 1067；切片器 2 / 2 / 2（Univer 渲染不稳定）；透视图 1 / 1 / 1。

本模块的九条坑单独记在 pitfalls 59-67：中文化资源包（59）、语言包导致文案翻案（60）、偶发接口采样（61）、create/render 分离（62）、配色必须比像素（63）、`baseFieldName`（64）、「空白」的口径（65）、Excel 截图（66）、生成器吃尾章（67）。

## 剪贴板 / 粘贴模块：76 项三方跑批（v1，2026-11-12）

**口径分两层，别混着判**：①**通道层**——「复制/粘贴到底走哪条通道」（系统剪贴板 / 引擎内部缓存 / 私有格式），只有**有头浏览器 + 真 Windows 剪贴板**测得准，无头环境的 `navigator.clipboard` 是进程内缓冲，只能证明 Chromium 自己没被写（pitfalls 71）；②**参数层**——`pasteOption` / 粘贴类型的覆盖度，走 API 枚举即可。

- **判定只认两件事**：目标区读回的值/格式，以及**剪贴板里到底有什么**（逐格式字节数）。不比错误文案与返回值。
- **NA 只给两种理由**：「本引擎读不到」（如程序化点击不触发原生 `paste` 事件，SJS 右键「粘贴」即此）与「基准建不起来」（如 Excel `CopyPicture` 图表的剪贴板里没有位图）。NA 不进分母。
- **哨兵法（必备）**：每轮先把系统剪贴板写成哨兵（如 `WINSENTINEL-CTX-0713`）→ 引擎里贴出**源区数据**即走内部缓存、贴出**哨兵**即走系统剪贴板；再补一步「复制完把剪贴板换回哨兵」，两条通道完全分离。
- **本样本结论**：SpreadJS 是**两套互不相通的缓存**（`Ctrl+C` 写 63 B 文本 + 2803 B HTML、HTML 里带 `gc-sjs-clipboard="true"`；右键「复制」一个字节都不写，走内部缓存），Univer 是**一套系统剪贴板**（右键复制与 `Ctrl+C` 逐字节相同）。后果：SJS 右键复制的内容 `Ctrl+V` 贴不回、粘到 Excel 只落哨兵；Univer 两条路等价。
- **界面层**：右键菜单都带「粘贴选项：」6 图标行；SJS 的「选择性粘贴...」是**拆分菜单**，点文字区直接弹对话框（与 Ribbon 同一个 `gc-sjs-paste-special-dialog-container`），点箭头才是子菜单——**不要凭「带 ▸」判「只能走子菜单」**（pitfalls 69）。
- 76 项结果：**SJS 62/10/4（可判 72，86.1%）、Univer 38/32/6（可判 70，54.3%）**，主差项是 J2 选择性粘贴（24 项里 SJS 16/5/3、Univer 5/16/3）。
- 本模块六条坑记在 pitfalls 69-74：拆分菜单（69）、SJS 右键菜单定位三坑（70）、无头剪贴板口径与 Excel `IDataObject`/Esc 清空（71）、**NA 的理由要能被反证**（72，本轮靠它把 J4-08 从 NA 改回 PASS）。

## 关键原则

1. **结论两分法**：区分"**保真层**（xlsx 导入还原能力）"与"**可编程层**（API 配置深度）"，分开下结论。对方可能在保真层表现出色（导入全链路保真）而在可编程层文档缺失、能力深埋——此时**不可说"不支持"**（会被保真层证据反驳），话术应转向"全部文档化、开箱可查、开发效率可预期"。
2. **可修改性通道四分法 → 支持度五级呈现**：接口文档没写 ≠ 改不了。判定按四通道：**①文档化 API ②未文档化但实测可用 API ③内置 UI 面板可改 ④仅模型层保留（无任何读写入口）**，判定顺序 **UI 面板 → API 层 → 内部模型层**（"API 层无键"只判到模型层确认无字段才可判 S5）。呈现时换算为 S1～S5 支持度分级（S1=通道①，S2=通道②，S3=仅通道③，S4=仅通道④，S5=三层全无）。判定前必须探测**双方**的内置 UI 面板——SpreadJS 有设计器组件（在线表格编辑器：功能区/公式栏/侧面板/上下文菜单，可嵌入 React/Vue/Angular，需单独部署许可；数据图表另有 DataChartConfigPanel 配置面板），Univer 有右键 Edit chart 面板（脚本 `scripts/univer_ui_panel_probe.js`）。实战发现 Univer Customize 面板能力远超其 Facade 文档（趋势线/刻度线/标签旋转等全有 UI 控件），漏测 UI 会严重低估对方；但 UI 能力要标 S3（仅界面可改）——对开发者是风险不是优势。话术口径="配置能力分散三处、程序化试错成本高、部分无任何修改入口"而非"不能配置"。
   **（2026-11-04 用户定版基准，见 verdict-framework"判定基准原则"）**：最终"支持"判定只认**官方公开 API 且调用成功**；有公开 API 但设置不生效=特殊标记"API 缺陷"单列统计；**无公开 API 一律判不支持**（原 S2 类项降级，运行时实测仅作附注）；导入/导出判定口径=图表是否保留、各配置项是否保留且正常显示。
3. **环境声明必写，版本号精确到两位**：每份测试记录和报告必须分别注明参评双方的**具体测试版本号**（如 SpreadJS v19.2.2 + 所用 npm scope、Univer Pro v1.3.1 + 测试页面），开源版与 Pro 版分开记；匿名 demo 与正式授权环境行为可能不同（Univer 匿名导出 permission denied），注明测试环境。**线上 demo 不内嵌版本常量时的判定法**（实战用于 office.univer.ai）：取站点 HTTP 响应头 `Last-Modified`（构建部署时间）与 npm `@univerjs-pro/*` 各版本发布时间比对，构建晚于发布即判定为该版本；再用 `npm view <pkg> time --json` 取时间线。
4. **SpreadJS 侧所有 API 判定，第一动作走 spreadjs-mcp**（2026-11-14 用户定版）。SDK 大版本 API 变动大，写脚本前必须先校准调用方式；校准通道固定：`spreadjs-mcp` 的 `search` → `fetch`（@grapecity-software 中文文档库，与 demo.grapecity.com.cn 同源）。三件事都用它：①查 API 名与签名；②查是否真有接口页（**有接口页 = S1 的文档证据**）；③查功能族清单。用法要点：按**功能命名**搜（"动态数组函数""文本操作函数"），不要只搜函数名——官网函数列表页本身不完整，11 个数组函数列表页缺失但实际全支持。网页搜索、凭记忆、读 npm 导出只能作补充，**不得单独作为"SJS 有/无此能力"的依据**；MCP 查不到 ≠ 不支持（改为实测定案）。Univer 侧对应通道固定为 docs.univer.ai 的 reference 层（不是指南页）。统计逻辑同样先做 API 探测，防止误判"0 存活"。
5. **分母要真实**：存活率分母用实测文件的图表数（解包 xlsx 数 chart XML），不硬编码。
6. **同一产品不同官方页面行为可能完全不同**（转换 demo vs 编辑器 demo vs playground），先确认用户在哪个页面测的，报告里分开记录。
7. **统计口径统一用通过率**（2026-11-05 用户定版）：一句话结论、分组表、分节标题一律写「通过 X/Y（Z%）」，**不写"X 通 Y 不通""X/Y 不通"**；分组表列头写「通过 / 通过率」，跨产品差额用 pp。分母写清楚是"可判项"还是"基线项"（两者常不等，因为有不测项）。生成统计的脚本要同步改，否则重新生成会把旧口径带回来。
8. **API 同名族要整族查**：`setColumnFormula`（页脚汇总）与 `setColumnDataFormula`（表内计算列）只差一个词，漏查一个会把我方能力判成"没有"——对外材料里比技术错误更贵。查法：把 `setColumn*` / `getColumn*` 全列一遍，再看官方示例落在数据区还是页脚行。
9. **样本先行（2026-11-06 用户定版，最高优先级）**：任何模块的第一步都是造样本文件，**先样本、后计划、再查证、最后跑批**。样本里指不到的测试项一律删掉，不许出现"矩阵有这项、样本没这载体"。样本同时是交付物——客户拿它到对方官方站自己就能核对，这比任何截图证据都硬。

10. **插件化 SDK：方法为 `undefined` 先怀疑装配，不是能力**（2026-11-06 工作表模块教训）。Univer 每个能力 = npm 包 + Plugin + `/facade` 子路径三件套，漏一件只让方法变 `undefined`、不报错。本轮四次误判：背景图（`@univerjs/sheets-drawing`）、焦点高亮（`@univerjs/sheets-crosshair-highlight`，不在 Core Preset）、查找替换（`@univerjs/find-replace` + `@univerjs/sheets-find-replace` **两个包**）、行分组（`@univerjs-pro/sheets-outline`，Pro 专属）。第五次是**批注**（`@univerjs/thread-comment` + `thread-comment-ui` + `sheets-thread-comment` + `sheets-thread-comment-ui` 四个包，另加两个 `/facade`），漏装的直接后果是 `sheet.getComments` 为 `undefined`，报告里就写成了「Univer 不支持批注」——补齐后 API 全通，真正的缺口是**只认新式讨论、不认 Excel 传统批注**（见 pitfalls 43）。**OSS 与 Pro 必须分开出结论**，对外只报 Pro 口径会失真，只报 OSS 口径会被"我们 Pro 有"当场打回。
11. **线上 demo 地址会变，旧 URL 打不开 ≠ 能力消失**（2026-11-06 实测）。`office.univer.ai/zh-CN/playground/sheets` 已下线（401 / Showcase not found），新结构是 Showcase 站点：示例页 `/zh-CN/showcase/sheets/<slug>`，编辑器实体在 iframe `/zh-CN/playground/sheets/<slug>`，**直接打开 playground 深链才能拿到 `window.univerAPI`**。写死的旧地址要在每轮复核时重验，并在报告里注明采集日期与页面路径。
12. **API 通 ≠ 效果对，UI 降级一律判 FAIL**（2026-11-06 用户定版，细则见 verdict-framework「判定红线 ①」）。代码能调、读回也对、但用户界面上看到的是降级效果 → 判 FAIL，说明列写「UI 降级：…」，原始 API 结论另存一列。已确认两例：SpreadJS 3D 图表（设得进去、画出来是平面 2D）、SpreadJS 页脚 `&F` 令牌（存得住、预览与 PDF 都不输出）。**进台账的门槛是真看过渲染结果**，凭印象不许进。
13. **基准线一律取 Excel 本体**（2026-11-06 用户定版，细则见 verdict-framework「判定红线 ②」）。凡是"一致/不一致"的结论，表里必须有 **Excel 那一列**；没有 Excel 列的对比只能写成"两家行为差异"（描述），不能写成"谁不对"（判定）。页面设置走 COM 读 `PageSetup` 做 oracle、打印输出走 `ExportAsFixedFormat` 出 PDF、公式走 Excel 365 实算值。

14. **判定层的错比产品差异更贵——数字进报告前先人工抽样核对**（2026-11-06 晚，一次抓出三类假差异）。
    - **读回是字符串先怀疑格式**：Univer 会给公式结果自动套数字格式，`getValues()` 返回 `(¥188.71)`/`8%`。判数值前两侧复算列先设「常规」。这一条曾把 11 个**完全正确**的财务函数判成"算不准"。
    - **日期序列号用本地墙钟字段重建 UTC**，别用 `getTime()` 相减（东八区差 0.667 天）；精度留到 1e-9，留 1e-6 会把 `TIME()` 截断。
    - **oracle 本身是错误值的用例单独判**（`NA()` 返回 `#N/A` 就是对的，记哨兵 `ERR`）。
    - 抽样核对法：挑 3～5 个典型 case，把「引擎读回原值 / Excel 原值 / 判定结论」三者对齐看一遍，对不上就是判定写错了。
15. **"对象在"不等于"还原了"**（2026-11-06 晚）。图表类必须三看：对象数、`chartType()` 读回、**画布实际画什么**。SpreadJS 曲面图四型（`surface`/`surfaceWireframe`/`surfaceTopView`/`surfaceTopViewWireframe`）对象在、类型读回也对，但导入时被转成簇状柱形（官方不支持清单）。同清单里的 `bubble3DEffect`/`boxWhisker`/`funnel`/`paretoLine`/`regionMap`/`waterfall`/`histogram` 同样要按"看画布"验。
16. **报告五条硬口径**（①②③ 2026-11-06 定版；④⑤ 2026-11-14 用户定版）：
    - ①导入导出不一致处**必须配对比截图**（两侧同表、同缩放、同取景）。
    - ②**不写"与上一版报告的区别"**，只写当前结果。
    - ③公式等无可见样式变化的模块**不加图**，但必须写清 Excel 样本路径 + 三方交叉表（都对/仅谁错/都错）。
    - ④**只陈述本次测试结果，不记录历史过程**：全文不得出现"上一版错了 / 本轮纠错 / 已作废 / 踩过的坑 / 改判轨迹"，一律改写为「口径 / 前置条件 / 复现要求」。**删除任何"保留修正轨迹"的要求**——读者只看得到这一版，版本对比是内部过程，进报告只稀释结论。
    - ⑤**不写测试方法叙事**：采集步骤、判定条件、复现命令、边界表全部删掉；能进代码注释就进代码注释。**唯一保留的是一张「测了什么」小表**（测试项 / 对象 / 量什么 / 结果在哪节），让数字有出处可查。删章后**下一章编号不动**，避免几十处交叉引用连带返工。
17. **判「无 API / 无读回口」前必须走三步，否则会把自家遗漏写成对方缺陷**（2026-11-07，打印区域事故）。
    - 事故：报告写「SJS 打印区域无读回口」，用户凭印象反驳；实际三条通道全通、读回值与 Excel 逐字一致（`rowStart/rowEnd/columnStart/columnEnd` + `getCustomName('Print_Area').getExpression()`）。真因是**自己的探针脚本漏采这四个字段**。
    - 三步：①先从 `.d.ts` 把类的**全部方法名**抓出来（`PrintInfo` 实测 32 个），不靠探针字段表反推能力边界；②字段表与方法表做差集，差集项逐个实测；③**同模块内对账**——元素级表格判「无/无法比对」的项，必须去 API 用例表确认也是 FAIL/NA，**两张表打架必有一张错**。
    - 措辞分三档，禁止混写：**无此方法**（读写都没有）/ **有 setter 没 getter** / **读得到但导入不回填**。
    - 连带教训：截不出图 ≠ 没法判，也可能是**样本不合格**（打印区域恰好等于整表，打了等于没打）→ 去补样本，不要降级成"无法比对"。
18. **判「没有这个能力」前必须先找替代入口 + 做效果实验**（2026-11-07，打印重复标题行事故）。
    - 事故：Univer 打印配置的 12 个 layout 字段里没有 `repeatRow/Column`，于是判 P5～P8 FAIL、写"只能整行整列冻结，无法指定第几行"。用户反问"冻结行列不就能实现吗"——实测一改就对：`setFrozenRows(3)` 后打印第 2 页就重复前 3 行，`setFrozenColumns(1)` 后横向每页重复第 1 列。
    - 三层错误：①把「配置对象里没有这个字段」当成「产品没有这个能力」；②只看字段名，没查该能力在产品里是否**另有入口**（打印配置 vs 工作表冻结）；③没做**效果实验**就下结论。
    - 正确顺序：①列出该能力的所有可能入口（配置对象、sheet 级属性、事件、命令）；②对每个入口在真实引擎里跑「改状态 → 导出/打印 → 看输出差异」，**读回一致不算证明，输出变化才算**；③把边界一起写清（本例：冻结区只能从第 1 行/列开始，`$3:$5` 做不到；打印重复与编辑器冻结窗格共用同一状态，Excel 的 `Print_Titles` 可以独立设置）。
    - 与第 17 条的区别：17 是「找不到 API」，18 是「找到了 API、但字段名对不上就判无能力」。**能力按"能不能做出效果"判，字段按"有没有这个键"判，两者要分开落表。**
    - 连带规矩：**API 用例的 FAIL 说明必须写清是哪一档**——字段不存在 / 有字段但功能坏 / 能力在但走别的入口。只有前两档能算 FAIL。
19. **元素级「有 / 无」结论：整页取证 + 逐元素落表**（2026-11-07，迷你图事故）。
    - 事故：打印报告写「条件格式、迷你图、合计行三家都打得出来」，实际 Univer 打印时迷你图整列空白。根因是拼图只抓了第 1 页、按墨迹裁到「日期」列就收边，**把「没裁到」当成了「在那儿」**。
    - 硬规矩三条：①判「有/无」必须先看**整页、全部页**，跨页元素按页遍历，「总页数 × 每页内容」要能拼回完整表；②**一张图只证一个元素**，不许一张图塞多个结论（塞两个必然有一个是顺推）；③元素名不许打包成「元素都没问题」，条件格式 / 迷你图 / 图表 / 合计行各占一行结论。
    - 排除布局因素的标准动作：把目标列**放宽 / 压窄邻列**拽进同一页再测一次。仍空白才叫"不画"，否则只是"太窄"。
    - 反查数据层三步：编辑器视图有没有 → 导出文件 XML 里在不在（迷你图查 `x14:sparklineGroup`） → 打印配置里有没有对应字段。三者都说"有"而打印为空白，才能把责任定到**打印渲染**这一段。
20. **打印 / PDF 模块：三层分开判，且"设得进去"永远不等于"打得出来"**（2026-11-07 定版，细则见 `references/print-pdf-module.md`）。
    - 三层 = **字段层**（set/get/读回）/ **导入层**（客户模板导进来还在不在）/ **输出层**（纸面是否一致）。三层的结论各自成段，混着写必被反问。实测反差：SJS 字段层 35/35，输出仍有 `&F` 不输出。
    - **改完设置要看导出文件有没有跟着变**。Univer 改页方向后导出，XML 仍是 `landscape` → 打印设置是**纯运行时状态**，既不读文件也不写文件。据此话术只能是"导进来要重设、改完存回去也不带走"，不能说"会把设置弄丢"。
    - **元素级结论逐元素一行**（图表 / 图片 / 形状 / 批注 / 条件格式 / 迷你图 / 表对象 / 页眉页脚 / 合计行），同一页里两种元素可以一个画得出、一个画不出。
    - **"顺序不同"这类结论要看模型**：SJS 单一有序集合 `shapes.all()`，Univer 拆成 `getShapes()/getCharts()/getImages()` 三个集合、跨类型前后关系不由文件顺序表达；对方元素压根没画出来时，只能写"本样本无从比对"。
21. **判"一致"的尺子必须先标定，标定不过就不许出数**（2026-11-07，公式模块）。
    - 事故：按"double 严格逐位相同"重算，SpreadJS 58.06%、Univer 59.96%——**被普遍认为精度差的那家反而更高**。因为两家都存不下 Excel 的 17 位，`=PI()` 在两边都被判"不一样"；这个尺子比的是"谁输出位数接近 17 位"，不是"谁算得准"。
    - 标定三步：①拿必然正确的用例（`=1+1`、`=PI()`）看两家是否都过；②拿人为差 1e-3 的对照看是否都不过；③看排名是否符合已确认的能力事实——反直觉先怀疑尺子。
    - 正式口径定在 **Excel 自身精度（15 位有效数字）归一后逐字比**，并明确"这是位数归一、不是容差"；**不许再加相对容差**（1e-9 会放过 `BESSEL*` 的 1e-8 和复数 26 条，报告变成"两家都对"）。
    - 同理适用于任何"逐值比对"模块：尺子的松紧直接决定排名，先证明尺子，再谈产品。
22. **SpreadJS 的能力也分包：addon 版本必须与主包完全一致**（2026-11-07，透视表）。
    - `GETPIVOTDATA` 不装 `spread-sheets-pivot-addon` 就返回 `#REF!`；**跨小版本注册失败，表现和"没装"一模一样**（本轮主包 19.2.2、插件也必须 19.2.2）。判"不支持"前先查官方 addon 清单，和 Univer 的"功能包 + UI 包 + facade"是同一类问题。
    - 加载顺序 **core → pivot addon → resources-zh → designer**（Designer 最后接管 Ribbon/UI）。**装了 addon 不等于界面只能英文**：面板/表头/弹框全英文的真因是缺核心中文资源包 `@grapecity-software/spread-sheets-resources-zh`（Designer 的 `resource.cn` 只管 Ribbon 与普通对话框）；装上后面板与单元格文案同时中文化。只切 `culture('zh-cn')` 无效——插件的 `SR.langs` 只有 `en`，zh-cn 槽位被英文兜底填满 358 键，永远走不到缺词条分支。资源包必须排在 pivot 插件**之后**（包内注册有 `GC.Spread.Pivot.SR` 守卫，提前加载静默跳过且不自愈）。详见 pitfalls 59。
    - 对侧同理：Univer 的透视表插件在、`addPivotTable` 可用、透视表结果区数值全对，但 `GETPIVOTDATA` 两条用例（无参 / 单条件）全 `#NAME?`，探针另测的双条件也一样。结论只能写"透视表有、这个取值函数没有"，不能写"不支持透视表"。
23. **"单独搭场景"不等于"只报支持度"**（2026-11-07，用户当场纠正）。
    - 事故：把 `SHEET`/`SHEETS`/`GETPIVOTDATA`/`IMAGE` 这 4 个函数判成"只算支持度、不进准确度"，理由写成"要专门样本、参照系不同"。用户反问：**明明能测结果，为什么不比？**
    - 事实：7 条用例里 5 条 Excel 有标量真值，两家全对（SJS 5/5、Univer 3/3）。写成"只算支持度"等于把已经比过的结果扔掉，还让读者以为"这组没法评价"。
    - 正确做法：**单列一组，两个分母都报**——支持度 ÷ 全部 7 条，准确度 ÷ 该家"支持且有真值"的条数。只有 `IMAGE` 两条（Excel 本机返回错误码 `0x800A0803`）确实没有真值，才只判支持。
    - 一般化：任何被排除出主分母的用例，先问一句"它的结果到底能不能比"；能比就必须给准确度，不能比才只给支持度。参照系不同只影响"放不放进主表"，不影响"要不要比结果"。
24. **性能数字的三条前置：每轮新页、取中位、只看重算**（2026-11-07，公式性能）。
    - **每轮新开页面**：同页连跑，Univer 第 2 轮重算从 11 s 涨到 22 s（`dispose()` 后公式 worker 残留），会把倍数凭空翻倍。
    - **取中位不取均值**：6 次极差占中位 45%（SJS，含首轮冷启动）/ 19%（Univer），均值会被一次抖动带偏。
    - **写入耗时不能当引擎指标**：SJS 3.2 s 里有 Designer 每次 `setFormula` 通知 Ribbon 的成本（裸 Workbook 1.1 s）；比引擎只看重算那一栏，写入单列为"带真实外壳的工程成本"。
    - 附带：负载条数按**实际写公式的列**算（`Data` 5 列数据只写 4 列公式，按 5 列会多算 10,000 条）；跑完抽 9 个格双端逐字相同，才能排除"一边少算了一批"。

25. **UI 入口判定：自动化点不开 ≠ 按钮不可用**（2026-11-14，用户用一张截图推翻整节结论）。
    - 事故：探针点 Univer Ribbon「排序」主按钮，无菜单、无报错、数据不变，据此写成"普通表格 UI 上无排序入口"。用户手动一点，弹出 5 项菜单（当前区域升序 / 拓展区域升序 / 当前区域降序 / 拓展区域降序 / **自定义排序**）。整节作废，且已写进报告。
    - 四个叠加原因（缺一不可）：**① split button 主按钮与箭头是两个热区**，点 rect 中心落在主按钮；**② 标签在 `aria-label` 而非 `title`**，`[title="排序"]` 命中 0；**③ 菜单 portal 到 `body`**，`#uniHost *` 里查不到；**④ 菜单在 pointerdown 展开、pointerup 关闭**，`mouse.click()` 太快，需 `move → down → wait → 查询 → up`。
    - 纪律：判"UI 上没有 X"前，**先 dump 当前 tab 下全部 `[title]`/`[aria-label]` 用清单说话**；`disabled=false` + `cursor:pointer` 一律先按"入口存在"处理；**自动化点不开必须补人工点击或查产品文档**，不能把"脚本没点开"写成"产品没做"。
    - 这是"用 A 的交互方式否定 B 的能力"的第二例（第一例：拿 SJS 的 `sortRange` 去 Univer 找排序）。同类误判的共同特征：**探针返回的是"我的调用失败"，不是"对方的能力缺失"**。截图要覆盖菜单展开态，只截主界面会给错误结论背书。
    - 详见 pitfalls「UI 入口判定：自动化点不开 ≠ 按钮不可用」。

26. **判 FAIL 前先自证探针：参数类型对不对、默认规则能不能改**（2026-11-14，同日两条 FAIL 自我翻案）。
    - 两起事故：B7「自定义表样式」判 SpreadJS 不可用，真因是给 `TableTheme` 的样式槽位传了普通对象 `{backColor:'#FFD966'}`，官方要求 `new TableStyle(...)`；换成实例后四条路径全通（A1 底色 `#FFD966`、`toJSON` 保留主题）。D14「中文排序」判只按码点序，真因是没传 `compareFunction`；补 `localeCompare(zh)` 后与 Excel 逐值相同。
    - 两条通用规则：**① 样式/主题/配置类参数一律用 SDK 构造器实例**，报错里出现 `v.h5`、`e.h0` 这种压缩短名时，先怀疑自己传错类型；**② 默认行为不符 Excel 先找补齐参数**（`compareFunction` / `locale` / `ignoreHidden` / 选项对话框），补齐后能对齐就判 PASS，把"默认规则差异"写进差异说明而不是缺口。
    - 纪律：一条 FAIL 的说明列必须写清「试过哪些参数、补哪个参数能对齐」；判 FAIL 前跑一次最小对照（同路径改官方写法再跑一遍），两次结果不同即是探针问题。附带：`table.styleName` 在当前 SpreadJS build 是**字符串属性**不是方法，文档写法不等于实现。
    - 详见 pitfalls 76 / 77。


## 产出物清单

**v5 交付形态（2026-11-06 定版）：一个工程 + 一套文档，不散落多份时间线报告。**

- `对比工程/`（可整包复制）：`index.html` + `js/app.js|runner.js|cases/*` + `samples/*.xlsx` + `vendor/` + `tools/` + `out/`
- `对比工程/docs/` 模块报告，**统一四段式结构**（2026-11-14 用户定版，顺序固定，不许倒装）：
  1. **结论（先行）** —— 一句话结论 + 可判/通过/通过率数字表 + 强项 / 短板 / 双方都弱 / 需避开战场 + 销售话术。禁止把结论放文末。
  2. **导入导出对比** —— 3.1 导入后视图截图 / 3.2 导出文件 XML 级比对 / 3.3 导出后重开视图截图 + 差异清单。**只留本模块相关内容**，不铺全局。
  3. **三方截图 / 功能截图** —— Excel 基准 / SJS / Univer 拼一张（`assets/<module>-io-compare.png`），必要时加往返版（`-io-roundtrip.png`）。
  4. **API 完整度验证** —— 分组通过率表 + FAIL 逐条原因 + 附录 A 逐条用例表（`tools/gen_case_tables.js` 生成，禁止手写）。

  前置固定块：标题下先放「测了什么」小表（测试项 / 对象 / 量什么 / 结果在哪节）+ 环境版本行（双方版本号精确到两位）。**不设「用例设计」章**——分组与 NA 判定边界压进「测了什么」小表；方法叙事一律不进正文。已有报告保持原状不回改。
  - `00-方法论与工程.md`（总览 + 三处必知前提 + 口径与前置条件，**不写修订对账**）
  - `01-图表` `02-打印与PDF` `03-工作表与工作簿` `04-格式与条件格式` `05-表-排序-筛选-验证` `06-公式` `07-绘图对象` `08-数据透视表` `09-表单控件` `10-设计器与Ribbon` `11-交互手感与录入效率` `12-剪贴板与粘贴`
- `对比工程/out/`：`data/<模块>-cases.json`（**报告数字的唯一来源**）、`data/<模块>-diff.md`、`data/_io-all.json`（各模块往返统计，主图页脚用）、`shots/<模块>-对比图.png|对比图-往返.png`、`shots/<模块>-<side>-rt.png`、`export/<模块>-{sjs,uni}.xlsx`
- `对比工程/docs/` 每份报告固定含 **两张证据图 + 一张逐条表**：
  - `assets/<module>-io-compare.png`（导入后，竖排文档版 1140px，单侧 1100 才可读）
  - `assets/<module>-io-roundtrip.png`（**导出后重新打开**，回答"导出会不会更糟"）
  - 末尾「附录 A　逐条用例结果」全量 ID/测试项/两侧判定/原因，由 `tools/gen_case_tables.js` 生成；**图表模块不设附录**（其 API 判定章的「逐项判定」表即逐条表，项数以该表表头为准）
  - 文件名必须 ASCII，正文写相对路径 `assets/xxx.png`，否则转 Word/飞书丢图
- 旧根目录报告**必须加"已被取代"横幅**（指向 `对比工程/docs/`），否则两套口径并存、客户先看到旧的那份就白做了

历史形态（v4 及以前，仅作参考）：样本 + 测试计划 + 各模块 API 报告单份 + 证据图 + 《测试结论-<模块>.md》。

## 资源索引
- `references/workbench-architecture.md` — **单工程双引擎对比台架构**（目录职责 / Univer Pro 本地集成三要点 / 用例注册表契约 / 三条命令 / 从零复刻最小清单 / 已知边界）
- `references/dimensions-template.md` — 图表模块 A～O 基准清单模板（M/N 可选扩展 + 测试计划模式 + 证据强度标注）
- `references/print-pdf-module.md` — **打印 / PDF 导出模块作业指导**（三层判据 / 样本必备元素 / 35 项分组与 PASS 口径 / 元素级保真清单 / 五个易被反问的结论写法 / 判定前六条自检 / 复现命令）
- `references/formula-module.md` — **公式模块作业指导（v3 定稿）**（支持度/准确度两个分母 / 判定尺子三步标定 / 六步管线 / 4 条特殊函数含透视表插件版本锁死 / 性能五条纪律 / 话术边界 / 自检六条）
- `references/pitfalls.md` — SpreadJS/Univer 实测踩坑全记录（API 差异、COM 造数坑、内部模型路径、UI 探测方法、浏览器局限）
- `references/verdict-framework.md` — 保真层/可编程层两分法 + 可修改性四通道 + 竞争话术规则
- `references/api-probe-method.md` — **API 支持度与友好性探针法**（五步管线 + PASS/FAIL 二值判定 + 五维友好性评分口径 + 27 条陷阱清单；核心原则"运行时没有某个方法 ≠ 没有这个能力""A 环境看不到 ≠ B 环境不支持"）
- `scripts/` — 可复用脚本：
  - **单工程对比台（2026-11-06 增补，六模块统一复跑）**：`io_serve_https.js`（HTTPS 静态站 + `/universer-api/*` 反代，**不能删 `content-encoding`**，`set-cookie` 要摘 `Domain`/`Secure`）→ `univer_pro_bundle_build.js`（esbuild 打 25MB Pro bundle，license 从客户文件读）+ `univer_bundle_entry.js`（**插件与 facade 成对注册**，逐项 try 并写入 `window.__proExtraPlugins` 自证）→ `io_run_cases.js`（通用跑批，落 `out/data/<模块>-cases.json`）→ `io_capture.js`（导入截图 + 导出 xlsx + **导出后重开截图** `<模块>-<side>-rt.png`；内含 `setActiveSheetIndex`、按表名对齐、跳过 `Evaluation Version` 表、`showRow(frozenRowCount())` 视口复位四处修复）→ `io_diff.py`（导出文件双层比对：zip 部件层 + openpyxl 内容层，chartEx 单列）→ `io_compose_figure.py`（拼对比图：默认并排 / `--doc` 竖排文档版 / `--rt` 往返版）→ `io_gen_case_tables.js`（cases.json → 文档末尾「附录 A 逐条表」，幂等、单元格净化、边界安全截断）；`formula_oracle_build.py`（`data_only=True` 抽 Excel 缓存值成 oracle JSON）
  - 造数：`make_baseline_charts.ps1`（63 图）、`make_batch1_styled.ps1`（批次 1 逐项显式样式 21 图）、`make_batch2_styled.ps1`（批次 2 系列与高级元素显式样式）、`make_styled_chart.ps1`（单图三向验证样本）
  - 格式/条件格式全量样本：`make_style_cf_samples.ps1`（COM 一次产出单工作簿 12 sheet、A～K 135 项用例，D 组内置同文本同列宽只改属性的对照组）+ `patch_cf.py`（COM 造不出的规则直接补 sheet XML，含 dxfs 与 x14 层）
  - SpreadJS：`runner.js`+`test.html`（导入实测）、`extract_sjs_models.py`（toJSON→chartSpace 配置提取）、`sjs_create_types*.js`（逐类型代码新建实测：枚举遍历+读回核对+渲染截图）
  - Univer：`univer_roundtrip.js`（闭环）、`univer_import_headed.js`（Chrome 有头导入+逐表截图）、`univer_resources_dig.js`（resources 内部模型抓取）、`univer_ui_panel_probe.js`（右键 Edit chart 面板枚举）、`univer_create_types.js`（逐类型代码新建实测）、`univer_export_batch1.js`（导出方向链路）
  - 通用：`verify_config_preserve.js`（三向验证）、`measure_chart_size.js`（尺寸量化）、`probe_setters.js`（builder 层 setter 探测）、`rebuild_compare_shots.py`（三列对照拼图：COM Chart.Export 基线 + 行标注对应矩阵分组）、`export_excelview.ps1`（导出文件 Excel 实际展示逐图导出，含 3D 图 CopyPicture 兜底）、`export_baseline_fallback.ps1`（源基线同法重导补齐）、`sjs_api_batch2_fixed.js`（SJS 侧修改实测镜像脚本，与 `univer_api_batch2_fixed.js` 一一同口径）、`sjs_spill_region_probe.js`（#SPILL! 区域假设探针：扩行列 vs 清占用区 vs 清区重写公式）、`build_formula_compare_xlsx.py`（公式三方对照 xlsx：公式｜Excel 基线｜SJS｜Univer｜双列判定）
  - 公式 v2 管线（历史脚本，判定口径已被 v3 取代；造数 / oracle / 活公式对账仍可复用）：`data_block_v2.py`（共享数据区）→ `semantic_cases_v2.py`（语义用例+EXCLUDE）→ `excel_case_eval.ps1`（Excel oracle）→ `judge_v2.py`（Python 判定）→ `build_compare_v2.py`（七页对照 xlsx）→ `fill_live_v2.ps1`（活公式写入）→ `headless_readback2.ps1`（无头对账）；`formula_probe.html`+`sjs_v2.js`/`univer_v2.js`（双端同数据区跑批）
  - **公式 v3（2026-11-07 定稿，随对比工程分发，位于工程内 `tools/`）**：`run_cases.js formula`（双端跑批 500 条，出 `formula-cases.json`）｜`probe_special.js`（透视表 / `SHEET` / `SHEETS` / `IMAGE`）｜`probe_uni_pivot3.js`（Univer 透视表插件状态 + 透视表对象 + `GETPIVOTDATA` 三层取证）｜`run_perf.js` + `js/cases/perf.js`（本地工程性能：每轮新开页面、各 2 轮取中位、内置 9 格等价性校验）｜`recount_strict.py`（按"与 Excel 一致"重算支持度 / 准确度两个分母）｜`build_formula_xlsx_v3.py`（主报告 xlsx 构建）
  - 性能（旧宿主页版，保留作对照）：`perf_workload.js`（种子 42 确定性负载）+ `perf_run.js`（无头跑批，五步计时口径）+ `perf_sjs.html`/`perf_univer.html`（双端宿主页）+ `perf_univer_bulk.js`（NO_CALCULATION 口径复测）
  - **导入/渲染专项（2026-11-05 增补）**：`probe_numfmt.js`（不经过导入，直接 API 写值+套格式，验"格式化会不会把值变成文本"）→ `uni_roundtrip.js`（样本导入官方线上环境再导出 xlsx，做模型层复核）→ `render_sjs_samples.js`（同一份样本用 SpreadJS 渲染做基线截图）→ `export_excel_cf.ps1`（Excel COM 导出 PDF 做真基线，脚本内只写 ASCII 路径）
  - **API 支持度探针（三件套，配套 `references/api-probe-method.md`）**：`api_probe_table.js`（136 项探针表 + 跑批器，`window.ApiProbe={SPEC,runAll,runSide}`，含 9 条实测校准事实注释，换模块时只改 SPEC）→ `api_run_probe.js`（Playwright 双端跑批，`surface` 摸接口面 / `items` 出逐项结果）→ `api_gen_report.py`（`grade()` 机械定级 + 覆盖率 + 五维友好性评分，出 `api-group-scores.md` / `api-matrix.md`）
  - **宿主页布局诊断**：`layout_diag.js`（逐层打印 `#uniHost` 子树高度，找第一个"比父元素矮"的层。面板里出现莫名横条、或画布明显比容器矮时先跑它——实战一根因是 Univer 自己渲染的 `<header>/<footer>` 被宿主页裸元素选择器改了高度，详见 pitfalls"宿主页 CSS 污染"节）
  - **表/排序/筛选/验证模块（2026-11-05 增补）**：`cases_table.js`（46 条用例库，A～F 六组）→ `run_cases_table.js`（双端跑批，可传组别）→ `probe_table_gaps.js`（缺口复核：颜色筛选/前N项/排序命令/汇总行字段/文本长度）→ `build_samples_table.py`（6 sheet 基线样本）→ `shoot_table_showcase.js` + `compose_table_evidence.py`（两侧对照证据图）；环境侧 `univer_bundle_entry.js` + `build_bundle.js`（插件与 facade 成对引入、功能包与 UI 包两层）
  - **打印 / PDF 导出模块（2026-11-05 深夜增补）**：`build_print_sample.py`（openpyxl 造 3 sheet 多元素样本）→ `annotate_print_sample.ps1`（COM 补形状+迷你图并导 Excel 基线 PDF，纯 ASCII）→ `print_sjs.js`（导入+注册中文字体+逐表 savePDF）→ `univer_print_capture.js`（打印配置+print() 抓离屏 canvas）→ `print_api_matrix.js`（SJS 35 项 + Univer 16 项字段级 PASS/FAIL）→ `compose_print_evidence.py`（三栏/两栏证据图）；探针 `probe_print.js` / `probe_univer_print.js`（入口存在性、有无 getter、print() 的 DOM 变化）
  - **工作表 / 工作簿层样本（2026-11-06 增补，样本先行的首个落地）**：`build_samples_worksheet.py`（openpyxl 造 15 sheet：表管理/三态可见性/保护矩阵/冻结/三级分组/隐藏行列/网格线缩放/搜索靶子/docProps/UsedRange 四档/RTL/背景图，并包级注入 `docProps/custom.xml`）→ `annotate_worksheet_sample.ps1`（COM 补背景图与批注、23 项回读断言、`SaveAs` 重存成 Excel 本体产物）→ `verify_sample_worksheet.py`（解包 43 项包级断言）
  - **工作表 / 工作簿 API 层（2026-11-06 增补，127 项）**：`api_probe_worksheet.js`（A～L 12 组 127 项探针表，每项独占临时表 `__p<ID>`，换模块只改 SPEC）→ `api_run_probe_worksheet.js`（双端跑批，内置单项超时 + 逐项打点，出 `out/api-ws.json`）→ `api_gen_report_worksheet.js`（判定表生成 + Pro 改判表 `PRO`，改判来源在文件内可见）；自检三件套 `api_ws_diag1/2/3.js`；线上 Pro 补测 `univer_pro_outline_probe.js` / `univer_pro_caps_probe.js` / `univer_pro_surface_dump.js` / `univer_pro_permission_probe.js` / `univer_pro_outline_evt.js`
  - **打印 / PDF 字段矩阵（2026-11-06 增补，SJS 35 项 vs Univer Pro 19 字段）**：`print_pro_field_matrix.js`（Pro 线上逐字段「设置→print→拦 BeforeSheetPrintConfirm 事件→读回→cancel」，35 个用例，出 `out/pro-print/print-matrix-pro.json`）、`print_pro_effect_check.js`（改纸张/方向看输出画布尺寸是否跟着变）；本地侧 `测试工程/tools/print_api_matrix.js`（SJS 35 项 + OSS 16 项）
  - **读数口径与局部取证（2026-11-06 晚增补）**：`formula_readback_probe.js`（同一格同时读 `getValues()` / `getDisplayValue()` / 快照 `cellData.v` / 数字格式，区分"引擎算错"与"套了格式的显示值"）｜`cf_blank_probe.js`（空白簿 + API 建三类条件格式规则，剥离导入变量）｜`cf_official_api_probe.js`（同一段代码在官网 playground 跑一遍，分离"我们的构建"与"产品能力"）｜`io_shot_sheet.js`（按表名切表、两侧同缩放截图，第 4 个参数是缩放比，**缩放优先于滚动**）｜`compose_surface_diff.py`（曲面图缺失对照图：上下两段 + 计数徽标）
  - **展示台**：`build_showcase.py`（多份 md 合并成单页成果工程：`DOCS` 列表追加即入册，图片 base64 内嵌，自动重建 nav）+ `verify_showcase.js`（逐 tab 切换冒烟：h2/table/img 计数、图片解码、pageerror）
- `templates/dual-pane-harness/` — 双栏对照台模板（`index.html` 功能对照+性能基准+API 检查三视图、`js/perf.js` 双引擎适配层、`js/workload.js` 负载生成器、`js/api-probe.js` 136 项探针表、`README.md` 口径说明）。复制后需自备 `vendor/`（SpreadJS 模块包 + Univer 本地 bundle），详见 `references/dual-pane-harness.md`。

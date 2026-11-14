# 打印 / PDF 导出模块作业指导（v4，2026-11-10 定版）

适用：`打印`、`PDF 导出`、`页面设置`这一整块对比。**核心口径：四层分开判，不许混着写。**

| 层 | 回答的问题 | 判据来源 | 常见混淆 |
| :-- | :-- | :-- | :-- |
| **字段层** | 有没有 set + get，读回一致吗 | 35 条用例逐项 `set → readback → compare` | 把"能设"当成"打印出来就对" |
| **导入层** | 客户模板里的页面设置，导入后还在不在 | Excel `PageSetup` 做 oracle，两侧读回 | 把"导入丢设置"写成"没有这个能力" |
| **输出层** | 打出来的纸面一样吗 | Excel `ExportAsFixedFormat` 的 PDF 逐页比 | 用编辑器视图代替输出 |
| **元素层** | 纸面上出现哪些元素、像不像 Excel | 21 类元素 × 三家 × 三档（一致/降级/丢失），§5 | 只判"有/无"，漏掉"有但不一样" |

四层的结论**必须各自成段**，**百分数不可混算**（35 项是字段、21 项是元素，分母含义不同）。字段层 100% 不代表输出一致（本模块实测：SJS 字段层 35/35，输出仍有 `&F` 不输出这类 UI 降级）。

---

## 1. 样本（先样本，后一切）

主样本 `samples/打印-PDF-多元素样本.xlsx`，3 sheet，元素必须铺满——少一样，对应的保真结论就只能标"未测"：

| 类别 | 最低要求 | 载体 |
| :-- | :-- | :-- |
| 图表 | ≥ 3 类（柱/折线/饼） | Visuals 表 |
| 图片 | 1 张 | Visuals 表 |
| 形状 | ≥ 3 类（矩形 / 圆柱 / 文本框） | Visuals 表 |
| 批注 | 传统 Note ≥ 1（另造新式讨论样本，见 §6） | Data 表 |
| 条件格式 | ≥ 3 类（数据条 / 色阶 / 图标集） | Mixed 表 |
| 表对象 | 1 个 | Mixed 表 |
| 迷你图 | 一整列（≥ 20 行），列宽窄到会跨页 | Mixed 表 |
| 页眉页脚 | 六槽全填 + 页码 + 日期 + 文件名 | 三表都设 |
| 分页 | 至少一表纵向跨页、一表横向跨页 | Mixed 表设横向 |
| 打印区域 | **必须明显小于数据区** | 见 §7 样本缺陷 |

造数分工：openpyxl 写数据/样式/CF/批注/打印设置；COM 补形状、迷你图，并导出 Excel 基线 PDF（`build_print_sample.py` + `annotate_print_sample.ps1`）。

三个样本缺陷会让整块结论失真，造完必须检查：

1. **打印区域 ≈ 数据区** → 打出来看不出裁切效果，"打印区域生效"无从证明。
2. **列宽宽到一页放得下** → 横向分页、重复标题列、跨页元素全部测不出来。
3. **装饰用文本框与被测对象同容器** → 会污染对方导入结果（图表标题整条丢失的经典翻车点，见 pitfalls 33）。

---

## 2. 字段层：35 项，一家一张表、两侧对位在同一行

口径（用户定版，不要再改）：

- **PASS = set 与 get 同时存在且读回一致**；缺 set / 缺 get / 读回被改写，三种都记 FAIL，说明列写原因。
- **对方没有的一律写 FAIL**，不写"无此元素""暂不支持"——中间态会把分母做小、分数抬高。
- **只用公开 API**；读回走 `printInfo`（SJS）/ 事件载荷（Univer）。
- **分组**：打印范围 4 / 重复标题 4 / 输出形态 11 / 页眉页脚 8 / 版式 7 / 其它 1。

**Univer 的读回口不是 getter**：`getPrintConfig` 等四个方法全 `undefined`，真实读回口是事件载荷——

```js
univerAPI.addEvent(univerAPI.Event.BeforeSheetPrintConfirm, (e) => {
  const { layoutConfig, renderConfig } = e;   // 完整配置
  e.cancel = true;
});
wb.updatePrintConfig({ paperSize: univerAPI.Enum.PrintPaperSize.A3 });
wb.print();                                   // 不触发一次 print()，事件不会来
```

两个判定陷阱：

- `updatePrintConfig` **静默吸收未知字段**：随便传一个不存在的键，读回里也有它。字段白名单只能来自 `.d.ts`（`ISheetPrintLayoutConfig` 12 + `ISheetPrintRenderConfig` 7 = 19），不能来自"读回里有这个键"。
- **对象键序会变**，比较前按 key 排序再 `JSON.stringify`。

SJS 侧相反：`printInfo` 是 getter/setter 二合一（`pi.x(v)` 设、`pi.x()` 读），**未知方法直接抛错**，所以"设进去没报错"在 SJS 侧才算数。

### 定稿结果（SJS 19.2.2 vs Univer Pro 1.3.1）

**SpreadJS 35/35（100%）｜Univer Pro 19/35（54.3%）**

Univer 的 16 条 FAIL 分三档：**11 条字段不存在**（外边框、黑白、页序、页码范围、起始页码、首页不同、奇偶页不同、页眉图片、最佳列宽、最佳行高、打印质量）+ **1 条字段在但功能坏**（`pageSizeCustom` 传 500×700 读回 394×588，固定减 106/112，连续打印累减到 −3556）+ **4 条字段在、读回也对但打不出来**（打印区域四角 P1～P4：`area=AllSelection` 时打印预览 0 页，只有 `CurrentSheet` / `Workbook` 能出页）。

打印区域这一条最容易误判：`updatePrintConfig` 能写、能读回、读回值也对，**但区域不参与分页**。判它必须真开预览数页数，不能只判读回。

三组同判 PASS 但语义不等价，**讲的时候必须带边界**：

| 组 | 差异 |
| :-- | :-- |
| 重复标题 P5～P8 | Univer 走**工作表冻结**：`setFrozenRows(n)` / `setFrozenColumns(m)`。边界：只能从第 1 行/列起；打印重复与编辑器冻结窗格是同一状态 |
| 适配页宽/页高 | 只有 `scale: FitWidth/FitHeight` 枚举，不能指定"占几页宽" |
| 缩放比 | SJS 收小数 0.8，Univer 收百分比 80 |

---

## 3. 导入层：8 组页面设置，Excel 做 oracle

脚本：`excel-print-oracle.ps1`（COM 读 `PageSetup`）→ `probe_print_io.js`（两侧读回）。

实测：**SpreadJS 8/8 与 Excel 一致；Univer 2/8**（只有纸张、水平居中对得上），方向/缩放/页边距/页眉页脚/打印网格线全部回落默认 → 同一张表 Excel 1 页、Univer 2 页。

**并测定性：Univer 的页面设置是纯运行时状态，既不读文件也不写文件。**

| 动作 | SJS 导出后 XML | Univer 导出后 XML |
| :-- | :-- | :-- |
| 导入样本，原样导出 | `landscape` / 9 | `landscape` / 9 |
| 改成纵向再导出 | **`portrait`**，改动落盘 | **仍是 `landscape`**，改动不落盘 |

这条决定话术：不能说"会把你的设置弄丢"（三家导出的文件都完整保留 `pageSetup`/`headerFooter`/`_xlnm.Print_Titles`），只能说**"导进来要重设一遍，改完存回去也不会带走"**。

---

## 4. 输出层：PDF 逐页比

1. Excel：`ExportAsFixedFormat` 出基线 PDF；SJS：`savePDF`；Univer：抓打印 canvas 拼 PDF 或直接看离屏 canvas。
2. **导出前必须 `PDFFontsManager.registerFont`**（SJS 只内置 14 种西文标准字体，不注册中文乱码）。
3. 渲染比对用 Chromium（`render_pdf_pages.js`），文本层与页尺寸用 pymupdf。**两边结论不一致时，先查字体注册，再怀疑渲染器**（pymupdf 对子集字体会假性乱码）。
4. 硬指标：页数、每页方向、顺序、页尺寸；内容差异逐页列，不合并。

实测（同一份样本）：页数 4、方向、顺序完全一致，尺寸差 1pt；三处真差异——页脚 `&F` 不输出、`&D` 日期格式不同（`2026/11/6` vs `11/6/2026`）、长表分页断点差 3 行。

---

## 5. 元素级支持矩阵：21 项、三档判定

字段层（§2 的 35 项）回答"配置字段有没有"；元素层回答"纸面上出现什么"。**两层数字不可混算**——分母含义不同（35 项是字段、21 项是元素），混在一起两个百分数都失真。

### 5.1 样本：一份文件七张表，每表一类元素、每表一页

- 表名即分类：`Obj1`(图表 4) / `Obj2`(图片+形状+文本框) / `Obj3`(组合+连接线+表格对象) / `Struct`(合并/隐藏/轮廓/分页) / `Fmt`(条件格式+错误值) / `Note`(批注打印) / `CellImg`(单元格内图片)。
- 每项一个 ID（`E01`…`E21`），**这一套号贯穿三处**：判定表 `print-matrix-items.json`、报告 §2.8.1 的表格、复核页 `打印元素矩阵.html`。任一处改了号，另两处必须一起改（脚本只保证前两处同源，报告是手写的）。
- **每表必须控制在一页内**。Univer 不读 `printArea` / `fitToPages`，跨页会把对象切断，切断会被误读成"元素丢失"（旧样本的 LOGO 就吃过这个亏）。唯一例外：`Struct` 故意留 2 页，用来测手动分页符。
- 元素尺寸按 pt 显式给，不靠自动布局——同一元素在三家的相对位置才可比。

### 5.2 三档判据

| 档 | 判据 | 不算差异的例子 | 算差异的例子 |
| :-- | :-- | :-- | :-- |
| 一致 | 元素在输出里，同页同区、相对几何成立、内容可读 | 轴刻度上限 800↔700、图例位置、数据条最大长度 | — |
| 降级 | 元素在，但有**可指名**的可见差异 | — | 连接线箭头没了、圆柱正文被压到边缘、合并块位置偏移 |
| 丢失 | 输出里没有该元素，或**只剩空框** | — | Univer 四类图表、批注清单、手动分页符 |

"降级"必须写清差在哪，只写"有差异"等于没判。

### 5.3 两段归因（决定责任落哪一段）

| 段 | 做法 | 能排除什么 |
| :-- | :-- | :-- |
| A 导入侧 | Excel 造样本 → 三家直接打印 | 客户真实路径 |
| B 自建侧 | SDK API 建同一元素 → 打印 | 排除导入通道，单看打印渲染 |

判定矩阵：A✗B✓ → **导入通道缺口**；A✗B✗ → **打印渲染缺口**；A✓B✗ → 极少见（渲染对但导入丢）。

**B 段不必 21 项全做**：只对 A 段判"丢失/降级"的项补测，把归因补上即可。省下来的时间用在 A 段把元素铺满。

### 5.4 定稿结果（2026-11-10，Excel 365 基准）

| 分组 | 项 | SpreadJS 一致/降级/丢失 | Univer Pro 一致/降级/丢失 |
| :-- | :-: | :-- | :-- |
| 图表（柱/折/饼/散） | 4 | 4 / 0 / 0 | **0 / 0 / 4** |
| 图片与形状（浮动图/内嵌图/矩形/圆柱/文本框/组合/连接线/批注/表格对象） | 9 | 8 / 0 / 1 | 6 / 2 / 1 |
| 表格结构（合并/隐藏行/轮廓符号/手动分页符） | 4 | 4 / 0 / 0 | 2 / 1 / 1 |
| 条件格式与错误值（数据条/图标集/公式 CF/错误值打印） | 4 | 3 / 0 / 1 | 3 / 0 / 1 |
| **合计** | **21** | **19 / 0 / 2＝90.5%** | **11 / 3 / 7＝52.4%**；可用 14/21＝66.7% |

要点：

- **图表是最大断层**：Univer 四类图打印出来只有圆角外框，框内全空。同文件在编辑器里四张图正常渲染，等 15s 再截仍是空框 → 缺口在打印渲染这一段。
- **连接线箭头、圆柱正文**是"降级"的典型：形状几何对，细节排版不对。只判"有/无"会漏掉这一整层。
- **单元格内图片三家都打得出**（7 张，含透明通道印章）。与"Univer 只能看不能建"不矛盾——导入能显示，打印就能出。
- **批注清单、错误值打印方式两家都丢**：SJS 有批注导入但 `PrintInfo` 无 `printComments`；Univer 传统 Note 导入即丢且无字段。归因不同，结论不能合并写。
- **手动分页符只有 Univer 丢**：它的打印配置只有 `maxRowsEachPage` / `maxColumnsEachPage`，没有分页符对象。

### 5.5 取证纪律

沿用旧三条，新增四条：

1. 判"有/无"先看**整页、全部页**；跨页元素按页遍历。
2. **一张图只证一个元素**——塞两个结论必有一个是顺推。
3. 排除布局因素的标准动作：放宽/压窄邻列把目标拽进同一页再测一次。
4. **判"丢失"前先排除渲染延迟**：同一页隔 15s 再截一次。打印预览的图表渲染可能晚于首帧，但不会晚 15s。
5. **判"丢失"前先扫完全部页**：Univer 的图表宽度会超出它自己的一页，右半部分在第 2 页——只看第 1 页会得出"半个图表"。
6. **位图 vs 矢量要在方法节写死**：Excel/SJS 出 PDF（矢量），Univer 只能取 canvas。判"降级"时已排除分辨率因素（同一张图上其他元素清晰）。
7. **行为类项要标 `kind=behavior`**：隐藏行、分组轮廓符号的基准是"Excel 本身也不打"，结论是"三家行为一致"，不能写成"三家都支持打印轮廓符号"。

反查数据层三步（判"打印渲染缺陷"前必做）：编辑器视图有没有 → 导出文件 XML 里在不在（迷你图查 `x14:sparklineGroup`）→ 打印配置里有没有对应字段。三者都说"有"而纸面空白，责任才落在打印渲染这一段。

### 5.6 工程落点：判定表是唯一数据源

元素层**判不了自动化**（"降级"必须人看图），所以它不做 `run()` 用例，而是"判定表 + 复核页"两件套。链路单向，不许反向改：

| 环节 | 文件 | 谁生成 | 能否手改 |
| :-- | :-- | :-- | :-- |
| 判定表（源） | `out/data/print-matrix-items.json` | 人 | **只能改这里** |
| 判定产物 | `out/data/print-matrix-verdict.json` | `build_print_matrix_verdict.py` | 否（重跑覆盖） |
| 复核页清单 | `js/cases/print-elements.js` | 同上 | 否（文件头标"生成物"） |
| 复核页 | `打印元素矩阵.html` | 人 | 是（读上面那份清单） |
| 报告表格 | `docs/02-打印与PDF.md` §2.8.1 | 人 | 是，但数字只能抄 `print-matrix-verdict.json` |
| 跑批挂载 | `print-cases.json` 的 `elements` 字段 | `run_cases.js print` | 否 |

三条硬约束：

1. **不许塞进用例表**。字段层那张表是 PASS/FAIL 两档 + 耗时列，三档值进去会被 `verdict()` 渲染成「—」；35 行与 21 行混排还会让人以为是一个分母。
2. **页数与页映射是同一条链**。`render_matrix_pages.py` 按判定表的 `pages` 校验 PDF 页数，页数一变就报错——改样本布局必须同时改 `pages`，否则后面所有"某页有没有这个元素"的结论都作废。
3. **一条命令重建**：`tools/rebuild_print_matrix.ps1`；只改了判定、页图没动时加 `-SkipSample`（5 秒跑完）。

入口：对比台顶栏「元素矩阵 21 项 ⇗」（只在打印模块出现）→ 复核页，可按组筛选 / 只看有差异 / 点行展开三家证据图并放大。

---

## 6. 五个最容易被反问的结论，各有专门写法

| 结论 | 不能写 | 正确写法 | 依据 |
| :-- | :-- | :-- | :-- |
| 打印区域 | "SJS 无读回口" | 三条通道全通，读回值与 Excel 逐字一致 | pitfalls 36/37 |
| 重复标题行 | "Univer 不能指定重复行" | 走冻结：`setFrozenRows(3)` 即每页重复前 3 行；边界=只能从第 1 行起、与冻结窗格同状态 | pitfalls 38/39 |
| 浮动元素层叠顺序 | "两边层叠顺序不同" | SJS 单一有序集合 `shapes.all()`（顺序与 `drawing1.xml` 一致）；Univer 按类型拆 `getShapes()/getCharts()/getImages()` 三个集合，跨类型前后关系不由文件顺序表达；且 Univer 图表对象在、画布无像素 → **本样本无从比对**，不能说"有差异" | 实测 `out/data/shape-zorder.json` |
| 迷你图 | "三家打印都在" | 条件格式、图表、合计行三家都打得出来；**迷你图 Univer 打印整列空白**（数据在、编辑器能画、列宽放宽到 200px 仍空白） | pitfalls 41/42 |
| 批注 | "Univer 不支持批注" | Univer 只有新式讨论一条通道，**Excel 传统批注（Note）导入即丢**；API 侧 `addCommentAsync/getComment/getComments` 全通 | pitfalls 43 |
| 图表打印 | "Univer 打印不了图表" | 柱/折/饼/散四类图的**打印输出只剩圆角空框**；同一份文件在编辑器里四张图正常渲染，等 15s 再截仍空 → 缺口在打印渲染段。这是抽测 4 类，63 类型的编辑器层结论另见图表模块 | `print-matrix-verdict.json` |

**批注专项**：另造 `samples/批注-新式与传统对照.xlsx`（Excel 365 COM：A1 `AddCommentThreaded`、A2 `AddComment`；`AddCommentThreaded` 只有 365 有，2021 抛错）。判定顺序：先测 API 能力 → 再测导入还原 → 解包分开数 `xl/comments*.xml` 与 `xl/threadedComments/*.xml`。Univer 导出会把讨论批注同时写成 `comments1.xml`（正文是 `[Threaded comment]` 占位串），**不能计成传统批注**。

---

## 7. 判定前自检（六条，缺一条就别发报告）

1. 字段表里的每一项，是否都能在 `.d.ts` / 官方 Reference 找到对应方法或字段？（找不到就写"无此方法"，不能写"没有能力"）
2. 元素级表格判「无 / 无法比对」的项，API 用例表里对应项是否也是 FAIL/NA？**两张表打架必有一张错。**
3. 每条 FAIL 是否写清了属于哪一档：**无此字段 / 有字段但功能坏 / 能力在但走别的入口**？只有前两档能算 FAIL。
4. 每个"看不见"的结论，是否看过整页、全部页，并做过放宽列宽的复测？
5. 每张证据图是否写明**哪张表（含序号）、第几页、什么状态（导入后/手工设置后/导出重开）**？
6. 结论里是否混了"字段层通过"与"输出正确"？混了就拆开重写。
7. 元素矩阵的百分数是否与字段层分开？是否写了档位分布（一致/降级/丢失）而不只给一个"支持率"？只给单一百分数会把"空框"和"箭头丢了"算成同一档。
8. 元素层三个地方的 ID 与数字是否同源？报告 §2.8.1 的 `E01~E21`、`print-matrix-items.json`、复核页必须是同一套号；报告里的 19/0/2 与 11/3/7 必须能在 `print-matrix-verdict.json` 里逐字找到。**报告手写、脚本生成——这是唯一一处可能两处不一致的地方，发报告前逐项核。**

---

## 8. 复现产物

| 用途 | 命令 / 文件 |
| :-- | :-- |
| 用例跑批 | `node tools/run_cases.js print` → `out/data/print-cases.json`（**35 项字段层报告数字唯一来源**；`elements` 字段另挂 21 项元素层 summary，只挂不并——两个分母各算各的） |
| 打印区域三通道 | `node tools/probe_sjs_print_area.js` → `out/data/print-area-sjs.json` |
| 冻结=重复标题 | `node tools/probe_uni_print_freeze.js` / `..._cols.js` → 图 `print-el-7-freeze-repeat.png` |
| 迷你图专项 | `node tools/probe_uni_print_sparkline.js` → `print-sparkline.json` + `print-el-8-sparkline.png` |
| 批注专项 | `node tools/probe_uni_comment.js` / `probe_uni_comment_types.js` |
| 浮动对象集合 | `node tools/probe_shape_zorder2.js` → `shape-zorder.json` |
| 页面设置 oracle | `excel-print-oracle.ps1` → `probe_print_io.js` |
| PDF 逐页渲染 | `render_pdf_pages.js` + `diff_io.py` |
| 证据图 | `compose_print_elements.py` / `compose_print_freeze.py` / `compose_print_sparkline.py` / `compose_print_preview.py` |
| **元素矩阵·一张命令全链** | `tools/rebuild_print_matrix.ps1`（造样本 → Excel/SJS 出 PDF → Univer 逐页采集 → 判定 → 出图 → 前端清单）；只重算判定与出图加 `-SkipSample` |
| **元素矩阵·造样本** | `build_print_matrix_sample.ps1 -Root .. -Out ..`（Excel COM，7 张表）+ `append_cellimg_sheet.ps1`（把富值内嵌图那张表并进来） |
| **元素矩阵·三家取图** | 由 rebuild 脚本内部依次调：`excel-print-oracle.ps1`、`export_print_pdf.js`、`render_matrix_pages.py`、`capture_print_pages.js`（分步参数见脚本头注释） |
| **元素矩阵·判定与图** | `build_print_matrix_verdict.py`（读判定表 → 档位 + 证据校验 → verdict.json + `js/cases/print-elements.js`）；`compose_print_matrix.py` → `print-mx-1..7-*.png` |
| **元素矩阵·复核页** | `打印元素矩阵.html`（对比台顶栏「元素矩阵 21 项 ⇗」进入；按组筛选 / 只看有差异 / 点行看三家证据图） |
| **排除渲染延迟** | `probe_uni_print_chart_wait.js`（4.5s / 9s / 15s 三档截图 + 编辑器视图对照） |

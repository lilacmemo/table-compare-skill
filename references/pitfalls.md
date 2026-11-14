# 实测踩坑记录（SpreadJS / Univer，2026-11）

## Excel COM 造数（基线样本）
1. **曲面图/股价图**：AddChart2 直接造会失败或空图——先 `SetSourceData` 绑好数据再改 `ChartType`。
2. **新式图表**（瀑布/树状/旭日/直方/排列/箱形/漏斗）：先 Range.Select() 选中数据区，再 `AddChart2`，否则得到空图表。
3. 63 图样本分 ～10 个工作表存放；PowerShell 脚本存 UTF-8 BOM 防中文乱码。
4. **PS 二维数组字面量传 COM 失败（批次 1 教训）**：`$ws.Range('A1:B6').Value2 = @('a','b',...),(...)` 这类嵌套字面量**只写入第一行**且不报错——必须显式 `object[,]`：`$data = New-Object 'object[,]' $rows,$cols` 逐格赋值后整体赋给 Value2。
5. **SetSourceData 必须指定 PlotBy**：`$ch.SetSourceData($range, 2)`（xlColumns）。缺省时 COM 按行/列推断错乱，产生 `$C$6` 式碎片引用、图表无柱形——且 Excel 端基线 PNG 也会坏，容易误判成"SDK 导入丢数据"。
6. **AddChart2 参数位置对齐**：AddChart2(Gallery, Type, Left, Top, Width, Height)——位置参数错位会把 Left/Top 当成 Type，图表类型全部静默变成默认柱状图。造带类型参数的逐项样式样本时逐个核对。
7. **数据标签**：series 级用 `$s.HasDataLabels=$true` 后配 `$s.DataLabels` 各属性；`ApplyDataLabels()` 可选；单点标签用 `$s.Points(i).DataLabel.Text=...`。**类别名标签（ShowCategoryName）经 COM 设置后可能不持久化进 xlsx**——造数后解包源文件核验，无法持久化的项在测试计划中改判"Excel 无法造数，降级文档查证"。
8. **图表位置防撞**：同一 sheet 多张图 Left 参数必须错开（图宽 520 + 间距），否则重叠渲染、截图互相遮挡。
9. **造数后必须核验源 XML**：解包 xlsx 数 chart XML 数量 + 抽查 `<c:f>` 引用完整性、showPercent/showVal 等标志位、chart 类型标签。COM 静默失败率高，未核验的样本会污染全部后续结论。

## SpreadJS（v19.2.5）
1. **npm 包用 @grapecity-software scope**：当前官方指向 `@grapecity-software/spread-sheets` + `-charts` + `-io` + `-shapes`（分模块安装）。旧整包 `@mescius/spreadjs` 已 404；`@mescius` 分模块包为海外渠道同源 dist，仅在 @grapecity-software 不可用时替代。**报告与测试脚本注释中必须写明所用包名 + 具体版本号**。
2. **xlsx 导入入口**：`spread.import(blob, success, error, { fileType: GC.Spread.Sheets.FileType.excel })`。不要传 ArrayBuffer（报格式错误）；不要用旧式 ExcelIO。
3. **读取图表**：v19 **已移除 `sheet.getCharts()`**，用 `sheet.charts.all()`。统计前先做 API 探测，否则会误判"图表 0 存活"。
4. **preserveUnsupportedChart** 是 **sheet 级 API**（`sheet.charts.preserveUnsupportedChart(true)`），不是 workbook 选项，且需在导入后逐表设置。
5. 试用版导入 xlsx 后会多一个 "Evaluation Version" 授权提示 sheet，统计时按表名排除。
6. 缺 shapes 插件时导入含 DrawingML 的文件会解析失败。
7. **配置还原读取路径**：`chart.toJSON().shapeData.graphicFrame.graphic.graphicData.chart`——SpreadJS 将 OOXML chartSpace **原样**保留在对象模型中，可直接按 OOXML 语义读（`chart.chartSpace.chart.plotArea.axes[i].scaling.min/max`、`chartGroups[].ser[].dLbls`、`majorGrid.dash` 等）。注意系列在 `plotArea.chartGroups[].ser`，不是 `plotArea.charts[].ser`；value axis 的 min/max 在 `scaling` 键下。批量提取脚本：`scripts/extract_sjs_models.py`。

## Univer（官方在线 demo / playground，2026-11 时点）
1. **不同官方页面行为完全不同**：
   - `office.univer.ai/exchange`：纯转换 demo（Upload/Download），Excel 图表在服务端转换即丢失；
   - `office.univer.ai/zh-CN/playground/sheets/charts`：编辑器 + 图表插件，`window.univerAPI` 暴露，数据层可还原；
   - `univer.ai/demo/spreadsheets`：成品展示页，匿名导出被拒。
   **先确认用户在哪个页面测的，再下结论。**
2. **Facade API**：图表用 `sheet.newChart(univerAPI.Enum.ChartTypeString.X).setSource({range,orientation}).setPosition(row,col).setSize(w,h).setTitle(text).build()` + `await sheet.insertChart(info)`；读取 `sheet.getCharts()`（返回 FChart：getId/getType/getInfo/toBuilder）。旧式 `newChart()` 无参 + `setChartType` 已废弃。
3. **导出**：`exportSheetByUnitIdAsync` 依赖协同服务（匿名 demo 返回 permission denied）；用 `exportSheetBySnapshotAsync(wb.save(), 'xlsx')`。ExchangeFormat 枚举值即小写字符串 `'xlsx'`，demo 未暴露枚举时直接传字符串。
4. **导入**：`importSheetToSnapshotAsync(File)` 返回 IWorkbookData JSON（可直接检查 resources）；`importSheetToUnitIdAsync(File)` 返回 unitId（需协同客户端 load）。`downloadFile(file, name, ext)` 触发浏览器下载（Playwright 可捕获）。
5. **数据层与渲染层背离（⚠️ 已被人工复核修正）**：导入后 `getCharts()` 对象和配置（标题/图例/坐标轴 JSON）完整（ID 带 `rustDrawing_` 前缀，Rust 引擎转换），自动化无头截图只见空占位框。**人工复核（真实浏览器）证实：图表其实是被导入后高度压缩到极低（绘图区不可见），手动拉高后正常渲染——即渲染链路是通的，缺陷在尺寸/布局还原**。教训：①不能只统计对象数就下"图表保留"结论；②也不能只凭无头截图就下"渲染失效"结论——"看不见"≠"没渲染"，截图证据必须有人工复核闭环。
6. **导入裁剪**：导入时工作表裁剪到实际使用区域（30 行 20 列 → 7 行 14 列），锚定在被裁剪区域的浮动对象（图表）位置越界、不可见。测试用图表锚点应放在数据区内。
7. `wb.save()` 快照顶层无 `drawing` 键，图表数据在 `resources` 通道，且 `chartType` 关键字可能不在快照中出现（服务端转换路径）——快照层检查不可靠，以导入后 `getCharts()` 为准。
8. **内部模型深挖（resources 层）**：`importSheetToSnapshotAsync` 返回的快照 `resources` 中 `SHEET_CHART_PLUGIN` 资源存图表真身——自有 schema（`data` 内按 unitId 分组，每图 `style.xAxis/yAxis/titles/legend/seriesStyleMap`），字段比 Facade config 丰富得多（如 `tick.position`、`tickLabelPosition`、`lineStyle`、`titles.xAxisTitle.rotate`、`dataPoints[].label`、`majorUnit` 全在模型层）。**Facade config 无键 ≠ 没存**（majorUnit 实战差点误判 ❌）。判定顺序：UI 面板 → Facade config → resources 模型。抓取脚本：`scripts/univer_resources_dig.js`。⚠️ 简单 `JSON.stringify(snap.resources).slice(0,N)` 会截断资源 JSON，必须整体取回。
9. **UI 编辑面板探测**：右键图表 → 菜单项文本定位 "Edit chart"（约 1.2s 延迟后出现）→ 面板含 Setup/Customize 两个 tab。**双击图表无效**；选中框是 canvas 画的（DOM 里搜不到小控件）；demo 页有干扰图表，先 `wb.insertSheet` 建干净 sheet 再操作。Customize 面板可枚举 label 元素（刻度线/标签旋转/反转/min-max/对数/趋势线/系列样式全有控件）。脚本：`scripts/univer_ui_panel_probe.js`。

## 通用
- **配置项"支不支持"必须三向验证，不能只看文档**：文档未提及 ≠ 不支持。方法：用 Excel COM 造一张显式设置样式的图表（标题/图例/轴/标签/网格线逐项设字体颜色格式）→ 对方 SDK 导入后读内部模型（Univer 用 `chart.getInfo()`，数据落 `resources`）→ 截图验证渲染 → 再导出解包 XML 对比。本项目因此把 Univer 一批"— 未提及"改判为"内部模型+渲染+导入导出全链路支持"（唯一丢失项：网格线虚线 dash）。脚本：`scripts/verify_config_preserve.js`。
- **"无公开 setter"也要二次排查**：Univer 的可编程能力很多藏在 **builder 层**（`FAxisChartBuilder` 参考页 docs.univer.ai/reference/facade/axis-chart-builder），且正确用法是 `chart.toBuilder().setXxx().build()` → `await chart.update(info)`（直接 update 裸对象无效）。`setLegend` 甚至接受文档未写的 fontSize/color/bold。排查顺序：① reference 站全量搜（不只 Facade 主页）② 运行时枚举 `toBuilder()` 原型方法 ③ 逐个试调用后读 getInfo 验证。教训：查文档要查到"参考页/类型定义"层，不能停在指南页。
- **量化"看不见"**：Univer 图表几何信息在 `getInfo().size/position/anchor`，导入后 size 常为 0×0 或几像素——用 size 分布（0×0 占比）代替"看不见"的主观描述。脚本：`scripts/measure_chart_size.js`。
- Playwright 无头浏览器截图不能验证画布内部滚动（wheel 事件对 canvas 可能无效），用超高视口（如 1440×2800）代替滚动。
- **无头自动化结论必须真实浏览器闭环（已根治）**：历史教训——无头截图曾把 Univer"图表高度被压缩"误判为"渲染失效"。现在用 Chrome 有头模式（见下节）自动闭环，不再依赖用户人工复核；无头模式仅用于对象层统计与批量截图初筛。
- 统计图表存活率时分母用实测文件的图表数（先解包 xlsx 数 chart XML），不要硬编码。
- Excel COM 对照：`$ws.ChartObjects().Count` 逐表统计最可靠；Quit 可能抛 0x800AC472（可忽略）。
- 证据呈现：用 PIL 拼接四宫格对比图（暗底 + 中文标签 msyh.ttc），一张图讲完整个故事；标签文字若与后续复核结论冲突，需重新生成或加注说明。

## 渲染验证方式（2026-11-02 增补）
- **用 Chrome 有头模式代替"用户人工复核"**：playwright `chromium.launch({channel:'chrome', headless:false, args:['--window-size=1560,980']})`，屏幕弹出真实窗口，截图即人眼等价证据。Chrome 用户级安装位于 `%LOCALAPPDATA%/Google/Chrome/Application`，channel:'chrome' 自动识别。context 用 `viewport:null` 时不能再传 deviceScaleFactor（报错）。
- **DOM 量测对 canvas 渲染的表格 SDK 无效**：Univer 整表渲染在少数几个大 canvas 上，单个图表没有独立 DOM 节点，getBoundingClientRect 只能拿到整片画布。判定"实际渲染了什么"只能靠截图（+像素着墨率分析）。
- 数据层 size=0×0 的两种真实表现要区分：63 图样本部分图表非零尺寸被压成 4.5px 极扁（可见可拉伸）；全 0×0 时（批次1样本 21/21）真实浏览器中**完全不可见**。不能拿一个样本的渲染表现外推另一个。
- 数据层 0×0 ≠ 无头浏览器伪影：Chrome 有头复核确认批次1样本 21/21 全部不可见，结论闭环。

## 可修改性通道判定（2026-11-02 增补，用户方法论贡献）
- 对比维度升级：接口文档没写 ≠ 改不了。每个配置项按四通道判定：①文档化 API ②未文档化但实测可用 API ③内置 UI 面板 ④仅模型层保留（无读写入口）。
- Univer 图表 UI 面板探测方法：右键图表 → "Edit chart"（menu item 文本精确定位）→ Setup/Customize 两个 tab。**双击无效、右键菜单要等 1.2s、选中框是 canvas 画的（DOM 里搜不到小控件）**。Customize 面板能力远超 Facade 文档：刻度线位置/长度/粗细/颜色、标签旋转、反转、min/max、对数、数字格式、趋势线（Add trendline）、系列透明度/边框全有 UI 控件。
- **SpreadJS 也有 UI 面板（用户指正，2026-11-03）**：设计器组件（在线表格编辑器）= 功能区 + 公式栏 + 状态栏 + 侧面板 + 上下文菜单，可嵌入 React/Vue/Angular（npm 包 @grapecity-software/spread-sheets-designer + designer-resources-cn/en），**需单独部署许可**；数据图表另有独立组件 DataChartConfigPanel（选中图表即在侧面板显示全部可配置项）。"SpreadJS 组件形态无内置面板"是错误表述——正确说法是"组件形态可选配设计器，纯代码场景也可不引 UI"。
- 结论纪律：评价**任何一方**"配置能力"必须先探测 UI 面板，否则会低估（会话中两次栽在这：先漏判 Univer 的 Customize 面板，又误判 SpreadJS 无设计器）。
- demo 页有干扰图表：用 wb.insertSheet 新建干净 sheet 再建图，避免点击命中 demo 元素。

## 报告可读性（用户要求，2026-11-03）
- 报告必须设"名词解释"节：保真层/可编程层/对象层/渲染层/**模型层保留**/UI 面板/四通道/0×0/✅⚠❌ 全部用大白话解释。"模型层保留"的通俗讲法："值还在内部数据库，但没给任何读写入口——钱存进保险柜没发钥匙"。行话不解释，销售/管理层读者看不懂，报告等于白写。

## API 修改实测轮踩坑（2026-11-03，代码修改生效性验证）

### SpreadJS 侧
- **读回颜色格式是 `RGB(r,g,b)` 或 `#RGBA(r,g,b,a)` 字符串**，与设置的 `#RRGGBB` 不等——比对必须归一化，否则全误判 FAIL。
- **枚举读回是数字**（LegendPosition.left=3、TextDirection.vert=1、AxisGroup.secondary=1、LineDashStyle.dashDot=4），比对用枚举常量而非字符串。
- **dataLabels.format 不生效的坑**：必须先 `numberFormatLinked=false`，否则 format 被"链接源格式"覆盖。
- **单位口径与 Excel 不同**：`series.gapWidth` 0～5（Excel 0～500%）、`series.overlap` -1～1（Excel -100～100）、`series.doughnutHoleSize` 0～0.9（Excel 百分比）——三者都在 **ISeries 上**（不是 chart 方法），代码设置必须换算。
- **网格线 dashStyle 枚举归一化**：设 LineDashStyle.dashDot(4) 读回 2(=dash)——LineDashStyle 与 LineType 两套枚举并存，以读回值核对。
- **API 面比想象大**：legend/dataLabels 的字体是平铺键（color/fontFamily/fontSize/fontBold 直接可设）；plotArea()/chartArea()/hoverStyle()/useAnimation() 都存在且生效；系列透明度用 backColorTransparency。
- **无 API 确认**：series 无 explosion 键（扇区分离）、无 pattern 专用 setter（仅 backColor 类型签名支持 IPatternFillBackColor）、chart 原型无 dataTable/gapWidth/holeSize 方法（后两者在 series 上）。

### Univer 侧
- **update() 局部 config 是主要可编程入口**：`update({axes:{x:{label...},y:{gridLine/line/tick/min/max/reverse}}})` 大多生效；但 **不支持的键会抛 `Error: Chart source is required.`**（信息与原因无关），且**失败调用也会把图表锚点重置**（row→0、负 offset）——锚点漂移是独立缺陷，判定"生效"必须只认 config.* 路径变化。
- **builder 路径与 update 路径能力不同**：scaleType/logBase/format/titleStyle 只有 builder（`toBuilder().setYAxis(...)`）能写；title 样式只有 builder.setTitle(全量对象) 能写。两条路都要试。
- **config 真实形状要先读再写**：selector 是 `{index:N}`（从 1 起）；series[i].label 是全量对象（visible/contentType/color/fontSize/bold/italic/format）；趋势线在**顶层 config.trendlines[]**（带 selector，键含 showFormula/showR2/dashType/period）；组合图次轴是 series[i].rightAxis。照真实形状全量写回，partial merge 大多被拒。
- **builder.setSeries 对导入图表直接崩**：`TypeError: Cannot read properties of undefined (reading 'dataPoints')`——导入系列无 dataPoints 字段。
- **setTrendline 需要系列名解析**：导入图系列名丢失 → `must resolve to exactly one series named "undefined"`——导入保真缺陷联动可编程层。
- 判定脚本骨架：`univer_api_verify*.js`（before/after getInfo JSON diff + 仅认 config.* + 候选逐试）。

## 类型新建实测 + 导出方向 + 证据图（2026-11-03 增补，v5.1/v5.2 轮）

### 图表类型"代码新建"实测
- **SpreadJS charts.add 参数顺序**：`(name, chartType, x, y, width, height, dataRange?)`——**name 在前、chartType 在第二位**。写反报 `Cannot read properties of undefined (reading 'indexOf')`（报错信息与真实原因无关）。dataRange 传公式字符串如 `"'表名'!A1:C5"`。
- **combo 不能直接 add**：报 `Can not set combo chart type!`。组合路径可用：`charts.add(name, columnClustered, ...)` 后 `chart.series().add({chartType: line, name, xValues, yValues})` 逐系列组合（实测柱+线成功）。
- **股价图有数据布局校验**（报错文案抄自 Excel）：stockHLC 要求标签列紧邻价格列——连续 4 列（标签+H/L/C）或纯 3 列均可；"1 列宽标签块 + 3 列价格块"的不连续区域会失败。stockOHLC（1+4）、stockVHLC（2+3 不连续）反而成功。**stockVOHLC 连续 6 列标准布局也失败，10 种布局穷举后如实记"无法创建"**——穷举过再下结论，避免被质疑"你姿势不对"。
- **"创建成功 + 类型读回一致" ≠ "渲染正确"**：SJS surface 4 种对象级创建成功、chartType() 读回一致，但画布渲染降级为普通柱形图；regionMap 以普通数据源渲染回退柱形。判定"可创建"必须配**渲染抽查截图**，报告分开写"对象层/渲染层"两个结论。
- **SJS 全 2D 渲染引擎**：3D 枚举类型（columnClustered3D 等）可代码创建，但画布是 2D 平面效果。报告必须设"3D 创建 vs 导入降级"的口径澄清节——用户/客户高发混淆点："能代码创建 3D 柱形图，为什么导入时 3D 会降级？"答案：①SJS 自己的 61 枚举含 3D 类型（渲染为 2D 平面）②Excel 独有 3D 变体（圆柱/圆锥/棱锥/直方/排列）无对应枚举、导入时映射降级，两回事。
- **SJS 截图前保证图表所在 sheet 是活动 sheet**：新建专用 sheet（`new GC.Spread.Sheets.Worksheet(name)` + `spread.addSheet`）+ `spread.setActiveSheet` + `spread.refresh()` 后再截；否则截到的是另一个 sheet 的画面（空白），白跑一轮。
- **Univer 新建图表正确范式**：`sheet.newChart(Enum.ChartTypeString.X).setSource('A1:C5').setPosition('M2').setSize(400,260).build()` → `await sheet.insertChart(info)`；给 newChart 传 config 对象报 `Chart builder is not registered`。枚举名是 **ChartTypeString**（值为小写字符串 'line'/'column3D'…），不是 ChartType；`sheet.newChart` 与 `insertChart` 是两个方法（newChart 返回 builder，insertChart 落库）。41 个枚举 41/41 可创建且 getType() 读回一致。
- **Univer 渲染抽查要点**：代码创建后立即截图可能全空白（渲染未完成），sleep 数秒 + PageDown 翻页二次截图；"编程面宽（41/41）"与"渲染质量（3D 系列大量空白/残缺）"分开下结论。
- **类型清单口径**：SJS 遍历 `GC.Spread.Sheets.Charts.ChartType`（61 个数字枚举）；Univer 遍历 `univerAPI.Enum.ChartTypeString`（41 个字符串）。以"遍历枚举逐个建 + 读回核对"为实测口径，UI 面板插入路径不作为判定依据（明确写进口径说明）。

### 导出方向实测（导出 xlsx → Excel 核对 + XML 抽查）
- 标准链路：SDK 导入样本 → `spread.export`（SJS，需 `{fileType: GC.Spread.Sheets.FileType.excel}`）/ `exportSheetBySnapshotAsync(wb.save())`（Univer）→ Excel COM 打开核对 `$ws.ChartObjects().Count` → 解包导出文件 grep `xl/charts/chart*.xml` 抽查。
- **三层分离警惕**：Facade 层无某键 ≠ 导出 XML 不含——Univer Facade config 无 explosion 键，但导出 XML 里 `explosion val="30"` 完好保留。"丢失"结论必须查**导出 XML**，不能只凭 Facade/model 层判。
- 导出 XML 抽查项直接对应矩阵：prstDash（虚线）、dispEq/dispRSqr/period/forward/backward（趋势线）、errBarType/errValType/stdDev（误差线）、pattFill prst（图案填充，注意 Univer 会把映射偏差固化进导出文件）、gapWidth/overlap/holeSize/symbol/explosion。
- **SJS 导出 XML 的颜色/枚举可能是小写写法**（如 `c80000` 小写）——grep 用 `-i` 或大小写都搜。
- Excel COM 打开导出文件核对图表数时，**排除试用版多出的 Evaluation Version sheet**（SJS 导出多一页，正式授权无）；此事是商业行为，报告正文按用户要求**不写**（付费即消除，与产品能力无关），仅测试脚本内部处理。

## 公式函数模块实测注意事项（2026-11-04 增补）

### 文档查证
- **SJS 官方"函数列表页"不完整，"页面没列"≠"不支持"**：VSTACK/HSTACK/TOCOL/TOROW/TAKE/DROP/CHOOSECOLS/CHOOSEROWS/EXPAND/WRAPROWS/WRAPCOLS（11 个数组操作函数）、TEXTBEFORE/TEXTAFTER/TEXTSPLIT、GROUPBY/PERCENTOF 在函数总列表页缺失但实际全支持。交叉核对渠道：① official demos `demo.grapecity.com.cn/SpreadJSTutorial/features/calculation/*`（按功能分组的 demo 最全）；② 函数参考页 `/help/docs/formulas/FormulaFunctions/<族>/<函数名小写>`；③ spreadjs-mcp 搜索（搜"动态数组函数/文本操作函数"等按功能命名，不要只搜函数名）。本轮"513 函数清单"口径已因此低估，实测才是最终判定。

### 跑批用例设计
- **易失函数**（NOW/TODAY/RAND/RANDBETWEEN/OFFSET/INDIRECT/CELL/RANDARRAY）：Excel COM 基线的缓存值是"造数时刻"，双端导入重算后必变。判法：只判"函数被识别 + 返回类型正确"（NOW() 返回日期序列而非 #NAME?），值本身不比对；RAND 类判"数值落在合理区间"。
- **SpreadJS 试用版多出的 "Evaluation Version" 首页 sheet 会偏移 sheet 索引**：SHEET()/SHEETS() 返回值 +1；SHEET("表名")、'Sheet1'!A1 引用、3D 引用不受影响。判定时按"偏移恒定"修正，并把 Univer（无偏移）作为对照，偏移差异本身写进报告。
- **SJS `spread.options.allowDynamicArray` 默认 false**：导入含动态数组公式的 xlsx 后必须显式开启 + recalcAll，否则 VSTACK/FILTER 等全部失效。测试双态：①默认态（记录与 Excel 的差异，即客户"忘开"真实场景）；②开启态（真实能力）。Univer 无此开关（恒开）。
- **"导入不重算读缓存值"与"强制重算"分开判**：SJS `doNotRecalculateAfterLoad` / Univer `CalculationMode.NO_CALCULATION` 都能保缓存值；Excel COM 保存的 xlsx 自带缓存值是基线。重算路径的准确性才反映引擎真实差异。
- **Excel COM 造数注意**：含动态数组公式的单元格用 `.Formula2` 属性写入（`.Formula` 会把 @ 隐式插入符带进来）；保存后用 xl.CalculationState 确认算完再存。

### 证据图制作与嵌入
- **导入侧图表"看不见"≠"没渲染"：先修复可见性再验配置**（用户方法论，实战验证有效）：Univer 导入图 0×0 不可见时，用官方 `FChart.setSize(w, h)` 修复尺寸（Excel pt×4/3 换算原始宽高）→ `exportImage()` 逐图导出渲染图 → 再核对配置项还原。1.3.1 实测：修复尺寸后标题/系列色/网格线/次轴/图例/数据标签全部正常渲染——**渲染链路可用，缺陷集中于导入尺寸**；同时暴露真实还原差异（刻度标签旋转未还原、虚线渲染成实线、图例位置偏移）。`exportImage()` 与 Excel COM `Chart.Export`、SJS `toImageDataURL()` 三端同构，是三列对照图的标准素材来源。脚本：`scripts/univer_restore_shots.js`。
- **Excel COM `Chart.Export` 对部分图表挂死或产出 0 字节**（实测：误差线图、折线+趋势线图稳定失败；柱形+趋势线正常）。解法：用图表的 `TopLeftCell.Address + ":" + BottomRightCell.Address` 设 `PageSetup.PrintArea` → `Chart.ExportAsFixedFormat(0, path)` 导出**单图 PDF** → `pymupdf` 转 PNG（`python -m pip install pymupdf`）→ PIL `ImageChops.difference` 裁白边。注意：不设 PrintArea 时导出的是整个工作簿（8 页）。CopyPicture+剪贴板、Location 到 chart sheet 两条路均失败勿试。
- **对照图行标注必须与所选图内容一致**：行标注写"趋势线/误差线"时必须选中实际包含该元素的图（解包 XML：`grep -l c:trendline / c:errBars` 定位 chartN.xml → drawing rels 反查第几个 anchor = 第几张图），不能默认用 c1。
- **Univer 修改后截图同构化**：修改完成后对关键图逐图 `exportImage()` 拼网格（2×2，每格标注"XX 项修改后"），比整表截图更有说服力——修改生效与否（如趋势线未变红）与矩阵 FAIL 结论互相印证。脚本：`scripts/univer_api_batch2_fixed.js`。
- **Excel 窗口屏幕截图会黑块**（COM 无头打开 + 截屏方式失败率高）：Excel 基线列**必须用 `$chartObj.Chart.Export($path, "PNG")` 导出图表对象**，不依赖屏幕渲染，稳定清晰。
- **对照图三列必须是"同一张图 + 同一状态"**（第三次翻车教训，用户连续三轮抓出图不对）：
  - 中列（被测产品导入列）必须用**导入后未做任何代码修改的原始状态**素材——误用 API 实测"修改后"截图，网格线颜色/标签样式自然和 Excel 基线对不上（实战踩中）。
  - **同表多图叠放是常态**（造数省空间常把多图放同位置）：按坐标裁切只能截到最上层图；且 v19.2 Chart 原型**没有 visible()/remove() 方法**（混淆名方法，隐藏法全失败）。**正解 = `chart.toImageDataURL()` 逐图导出**（与 COM `Chart.Export` 同构，叠放图/视口外图都能独立导出）。脚本：`scripts/sjs_restore_shots.js`。
  - **图表对应关系用 XML 验证，别猜**：drawing XML anchor 顺序 = COM ChartObjects 顺序 = charts.all() 顺序（正序 cN↔kN，实战已验证 chart1=图表1=月度销售统计=c1）。位置可做校验：Excel pt × 4/3 = SpreadJS px（96dpi）。
  - 逐图核对后仍要**整图读图验收**：逐行确认"左中列同类型同数据同样式"才交付。
- 拼图规范：PIL + `msyh.ttc` 中文字体；**每行加行标注并与报告矩阵分组一一对应**（行标注写明该行测什么：如"04 网格线（红色虚线/颜色/粗细）"）；列头色块白字（Excel 基线 / SpreadJS 导入 / Univer 导入）；行高固定、等比缩放、白底补齐。脚本：`scripts/rebuild_compare_shots.py`（改行定义即可复用到新模块）。
- **嵌入报告的截图必须逐张"读图"核对与文字说明一致**——本会话三次翻车：①把导入还原三列对照图挂在"代码修改实测"章节下（张冠李戴）；②基线列黑块导致"标题样式/标签旋转"标注对不上图；③中列用错"修改后"素材 + 叠图裁切错位。图挂进报告前用视觉模型读一遍：图里能否看到说明所声称的内容。
- 技术细节：PowerShell `Out-File -Encoding utf8` 带 BOM，Python 读用 `utf-8-sig`；Playwright `screenshot({clip})` 坐标是 CSS px（deviceScaleFactor 只影响输出分辨率）；SJS 整表截图前 `setActiveSheet + showCell(0,0) + refresh`，否则截到别的 sheet。
- 图片放报告同目录或子目录、用相对路径引用；**保持文件名不变重制内容**，报告引用自动生效无需改动；旧图先备份 `.bak`。

### 报告结构口径
- **测试数据/证据文件随结果章节标注**（"实测数据（本节）：..."行），不设独立证据索引节——用户反馈"数据和结果分开写有点割裂"。
- **双端实测项数量不同必须写明计数口径**（实战 61 vs 31 被用户质疑）：分母相同（同一份配置项基准清单），差异来自"对方能否发起修改调用"——无代码入口的项（S4/S5）不产生用例，差值本身就是可编程覆盖面差距，且未实测项在矩阵中均有 S4/S5 结论，未遗漏。不写清就像测试工作量不一致。

## 导出→Excel 实际展示核对 + 同口径修改对照（2026-11-04 增补，v5.2 轮）

### 对照证据图必须"两侧同口径"（用户抓出的第四次图不对）
- **代码修改实测的两侧截图必须镜像复刻**：一边用整表截图、另一边用逐图导出拼图，构图口径不一致就无法逐图对照（用户实战质疑"spreadjs和univer的还是没有完全对应起来"）。正解：两侧脚本一一镜像——相同的图（按谓词定位 k1～kN）、相同语义的修改调用、相同的修改方向（如 Univer `rightAxis:false` ↔ SJS `axisGroup=primary`）、相同的逐图导出方式（Univer `exportImage()` / SJS `toImageDataURL()`）、相同的 2×2 标注，最后 PIL 上下拼成"上 SJS / 下 Univer 逐位对应"的对照图。脚本：`scripts/sjs_api_batch2_fixed.js` ↔ `scripts/univer_api_batch2_fixed.js`。
- SJS series 集合没有 `.count()`（用 `for i<6 + try/catch` 遍历）；`series().get(i)` 改属性后必须 `series().set(i, s)` 写回才生效。

### 导出核对必须做到第三层："Excel 实际展示逐图核对"
- **计数 → XML → 实际展示，三层缺一不可**：Excel COM 计数 19/19 只说明对象在；XML 抽查只说明键在；**对象在 + 键在 ≠ 用户打开 Excel 能看见**。Univer 批次2 导出文件实战：计数 19/19、XML 键齐全，但 7 张图（08系列样式 c7～c10、09高级元素 c7～c9）在 Excel 中实际高度仅 5px（868×5 细条）——导入侧 3.4pt 尺寸塌缩经内部模型固化进导出文件，"导入塌缩→导出仍塌缩"闭环（导入 0×0 的图导出后反而正常，3.4pt 的仍塌缩）。脚本：`scripts/export_excelview.ps1`。
- **量化展示可用性**：对逐图导出 PNG 统计尺寸分布（height<120px 视为塌缩），一行代码抓出"哪些图不可见"，比肉眼翻图可靠。
- **Chart.Export 0 字节兜底更新**：3D 图 `Chart.Export` 稳定产出 0 字节文件（基线与导出文件同批索引复现）。`CopyPicture → ChartObjects().Add(临时图) → Chart.Paste → Export` 本场景实测有效（此前"勿试"结论仅适用于单图直出场景），比 PrintArea+PDF+pymupdf 轻量。**基线文件也要同样兜底重导**，否则三列对照图基线列出现粉格缺图。脚本：`scripts/export_baseline_fallback.ps1`。
- Excel COM 打开导出文件按 sheet 名定位图表（试用版多出的 Evaluation sheet 不影响按名取图）；ChartObjects 集合顺序 = 导出文件内 anchor 顺序，cN 编号跨文件可比。

### 渲染层计算正确性也要抽查
- **Univer 渲染层趋势线 R² 显示值与真实拟合显著偏差**（实测：同一趋势线 Univer 显示 R²=0.0291，Excel/SJS 计算 R²=0.9978，公式本身一致 y≈2.042x+1.18）——渲染层不仅画样式还算统计量，计算正确性单独作观察项记录，勿与样式保真混为一谈。

### Windows 脚本编码
- **工具直写的 .ps1 无 BOM**，Windows PowerShell 5.1 按 ANSI 读 → 中文字符串乱码报语法错。修复：`[IO.File]::ReadAllText($p, UTF8)` + `[IO.File]::WriteAllText($p, $c, (New-Object Text.UTF8Encoding $true))` 转 UTF-8 BOM 后重跑；含中文的 .ps1 一律先转 BOM 再执行。


### 公式批次1 导入实测结果与比对脚本归一化（2026-11-04 补，批次1 已跑）
- **Univer 动态数组导入即降级（@ 隐式交叉）**：Univer Pro 1.3.1 导入含动态数组的 xlsx 后，20/64 公式被自动加 `@` 前缀（`=@SEQUENCE(3,3)`），只算锚点首值、溢出区在模型中不存在，**全程无报错**——"无提示的数据静默截断"，比报错更危险，竞争口径强调这一点。
- **【已证伪，勿再引用】SJS"多列溢出导入链路缺陷"是伪结论**（2026-11-04 夜修正）：先前把 15 项 #SPILL! 判为"导入 xlsx 后重算链路缺陷（疑缓存溢出区占用）"，实际根因是**导入后工作表列数=数据实际宽度**（本例工作表只用到 A:B 两列，`getColumnCount()===2`），多列溢出（SEQUENCE(n,m)/FILTER/VSTACK/HSTACK/TAKE/DROP/CHOOSECOLS/EXPAND/WRAPROWS/TEXTSPLIT/MAKEARRAY/GROUPBY）越出工作表边界即报 #SPILL!。修复：`sheet.setColumnCount(30); sheet.setRowCount(300); sheet.recalcAll(true)` → **15/15 全部算出正确值，有值单元格 182→315，#SPILL! 归零，错误值归零**。
  - 教训 1：**#SPILL! 先查工作表行列数，再谈"引擎缺陷"**——`getRowCount()/getColumnCount()` 与公式所需溢出区域比大小，是零成本第一步。Excel 会自动扩展工作表、SJS 不会，这是**语义差异不是缺陷**；客户集成时若导入窄表就写多列溢出公式，默认视图确实会看到 #SPILL!，需在代码里扩区（或给出说明）。
  - 教训 2：**动态数组结论必须三态测试**——默认态 / 开 allowDynamicArray / **扩大工作表区域**。只测两态会把"工作表太小"错判成"引擎不支持"。
  - 教训 3：导入后单元格里可能是 **Excel 的缓存值**（如 #NAME?），必须重算后再读，否则会把 Excel 2021 的 #NAME? 当成产品计算结果（本例 VSTACK 初始读回 #NAME?，扩区重算后得到正确值）。
  - 更正后对照话术：SJS **全部算对**（含多列溢出，只需设置足够行列数）；Univer 的 `@` 隐式交叉截断**无论怎么设置都无法恢复**。
- **Excel 2021 基线的天然局限**：本机 Excel 2021 对 365 独占函数（VSTACK/GROUPBY/PERCENTOF/TOCOL/TEXTSPLIT 等）判 #NAME?，产品算对这些项时入 prod_beyond_excel 桶——这是"产品超越本机基线"，不是异常；竞争口径仍以文档查证层为准。
- **比对脚本归一化四件套**（漏一个就出假 mismatch/假通过）：
  1. 日期：产品给 ISO 串、Excel 基线给序列号 → `to_serial`（fromisoformat → astimezone() 转本机时区再减 1899-12-30；直接用 UTC 会跨日错一天）；转出的 float 序列号 vs 基线字符串数字必须**双转 float 再比**（'46330' != '46330.0' 是假 mismatch）。
  2. 错误对象：SJS 值可能是 dict（{'zYr':True,'_error':'#SPILL!'}）——**传原始对象给 classify 内的 norm_pv**，先 str() 再 json.loads 会因单引号 repr 解析失败，#SPILL! 被误判成普通值（prod_beyond 虚高）。
  3. 布尔：true/True 大小写不敏感比较。
  4. 公式判定：仅 `=` 开头才算公式单元格，否则数值单元格会被标"公式丢失"。
- **易失函数/溢出类用例读值注意**：wb.save() 快照里溢出区"没有单元格"≠"没算出来"，也可能是快照机制不落溢出区——下结论前用渲染层截图或导出→Excel 打开复核。
- 脚本：`scripts/univer_formula_batch1.js`（playground 导入+快照 dump）、`scripts/compare_formula_batch1.py`（三方比对，含上述归一化）。


### "运行时枚举存在"≠"文档支持"：SJS 3D 代码创建的证据口径修正（2026-11-04 补）
- **API 文档枚举页成员数必须与运行时枚举数对账**：SpreadJS v19.2.5 运行时 `Object.keys(ChartType)` 实测 61 个成员，而官方 API 文档 ChartType 枚举页只列 39 个——缺的 22 个值（4-7、15-17、21-23、29-32、39-41、44-48）恰好是 22 种 3D/变体类型（pie3D/line3D/column3D/area3D/柱条面积 3D 9 种/复合饼/分离型/曲面 4 种/bubble3DEffect）。结论：这 22 种代码创建是 **S2（运行时可创建+读回一致，但无文档且文档反向声明不支持）**，不是 S1。
- **官方《保留不支持的图表》（preserve-unsupported-charts）页整体过时**：它把 22 种 3D/变体连同 funnel/waterfall/boxWhisker/paretoLine 都列为"当前不支持"，而 v19 运行时实测后 4 种均支持（导入存活+枚举存在）——"文档说不支持"不能直接采信，必须运行时复测；反之"运行时能用"也不能写成"文档支持"。
- 判定方法沉淀：代码新建类结论的证据列分两段写——2D 文档化类型="文档+实测（S1）"；运行时枚举有但文档没有的类型="实测（S2，存在版本演进风险，文档反向声明不支持）"。销售话术引用时必须带风险说明。
- 交叉验证技巧：把文档枚举成员的数值收集成集合，与 0～max 值域求差集，差集常正好对上"未文档化类型清单"（本例 22=22 完美对上），可作为文档完整性的快速体检法。


## 判定口径"落格"纪律 + 公式三方对照表（2026-11-04 夜增补）

### 结论改了必须逐格改矩阵，不能只改声明（用户第五次抓图）
- 上一轮已在报告开头写入"v5.3 新基准：S2 一律判不支持"，但 **2.2 矩阵里 10 个单元格仍写着 S2**，用户一眼看出"等级根本没改"。规则：**任何全局口径变更，都必须把矩阵中每一格改写到位**，并在版本注记里写清"实改 N 格"；只在标尺/声明处写"按新基准重分级"视为未完成。
- 呈现层同时清理：等级表增加"判定为支持？"列，让 S2 明确显示 ❌；`S2` 不应作为任何"支持度等级"出现在商业口径矩阵中（要么写"不支持（附注…）"，要么写"API 缺陷"）。
- 对双方一视同仁：口径收紧后必须回头抽查**己方**（SJS）的 S1 是否都靠公开 API。抽查方式=spreadjs-mcp 搜每个可疑 API 的接口页/demo；本轮抽查 hoverStyle/useAnimation、IPatternFillBackColor/PatternType、chartArea/plotArea 全部有 API 接口页 + 官方 demo，S1 成立。

### 公式/配置项对比的交付物：一张 xlsx 胜过十段文字
- 用户要求"一列公式、一列 SJS 结果、一列 Univer 结果"的直观对照表 → 脚本 `scripts/build_formula_compare_xlsx.py`（读 def/baseline/SJS dump/Univer dump 四个 JSON 直接生成）：列 = 工作表｜单元格｜公式（含标签）｜Excel 基线｜SJS｜Univer｜两列判定；冻结前两列 + 自动筛选 + 判定列配色；第二个 sheet 为判定桶汇总 + 口径说明。
- 判定桶设计（避免假差异）：`一致` / `超越基线（产品算出值）` / `基线不支持，产品也报错` / `偏移 +1（试用版首页 sheet）` / `易失函数（类型正确｜异常）` / `不一致` / `产品报错`。易失函数按**返回类型类别**（num/text/date/bool）与基线比对，不做值比较；INDIRECT/OFFSET 也属易失但返回值可比对，归入类型判定即可。
- 生成后必须回读校验：openpyxl 重开、打印行数与前 5 行、grep 判定为"不一致/异常"的行数应为 0；数字字符串比较（'4' vs 5）要转 float，否则 SHEET 偏移判定失效。

## 公式用例四原则 + v2 管线（2026-11-04 深夜，用户定版，最高优先级）

### 用例设计四原则（用户对本轮批次的四点反馈，全部定版）
1. **用例必须先在 Excel 中返回真实值（Excel oracle），禁止"错误类型对比"**：上一版按函数族配额抽的用例有相当一部分在 Excel 本身就报错（如 FORECAST 缺参数返回 #N/A），拿错误码做对比是伪命题。正解：为每个函数构造"Excel 能出值"的语义用例，先用 Excel COM 逐条评估（`excel_case_eval.ps1`），Excel 报错的用例当场修（参数顺序/数据区）或剔除（环境依赖 EXCLUDE），最终 500/500 条 Excel 全部出值（唯一例外 NA 本身就该返回 #N/A）。
2. **先判"是否支持"，再判"是否一致"（两级判定链）**：产品返回 `#NAME?` = 函数未实现，直接判"不支持"，**不参与值比较**（旧版把 #NAME? 算"不一致"虚增了差异量）；写入被拒 = 无法写入；其他错误 = 产品报错；双方都有值才比较。**（2026-11-07 v3 修订：值比较已改为"数值按 Excel 15 位有效数字归一后逐字比、文本原样逐字比、不设相对容差"，取代本条原来的 1e-9 相对容差；见 §47～49）** xlsx 里体现为 G/J 列（是否支持）+ I/L 列（判定）四列活公式。
3. **差异函数单独成 sheet**：「SpreadJS差异明细」「Univer差异明细」两张独立页（#｜函数名｜函数族｜用例公式｜Excel 结果｜产品结果｜差异类型｜类别｜可能原因），便于人工逐条确认；类别分"精度差异/随机·时间/区域设置/格式差异/类型语义/功能差异/报错"。
4. **Python 判定与 Excel 活公式判定必须对账 0 不一致**：xlsx 里的判定列用活公式（支持列 `LEFT(...,10)="ERR:#NAME?"`、判定列先 #NAME?/SET 再数值容差/布尔/文本分支），Excel 重算后逐行读回与 judge_v2.py 的 Python 判定比对，0 行不一致才算交付。本轮曾发现 9 行口径漂移（取消"15 位有效数字归一"改为显示文本 EXACT 后修复）。

### v2 管线（脚本已入 skill/scripts/）
`data_block_v2.py`（三方共享数据区唯一真源，108 数值格 + 1 公式格，日期用 Excel 序列号存避免时区差异）→ `semantic_cases_v2.py`（503 语义用例 + 19 条环境依赖 EXCLUDE；输出 cases_v2.json）→ `excel_case_eval.ps1`（Excel COM oracle 逐条评估，产出 excel_eval_v2.json）→ `sjs_v2.js`/`univer_v2.js`（双端同数据区跑同用例）→ `judge_v2.py`（Python 判定，verdicts_v2.json）→ `build_compare_v2.py`（7 页 xlsx：对比表/结论/两张差异明细/分族统计/口径说明/Excel实测）→ `fill_live_v2.ps1`（Excel COM 把 500 条活公式写入「Excel实测」Z 列，行间隔 12 容纳动态数组溢出）→ `headless_readback2.ps1`（无头读回对账）。

### v2 管线技术坑
- **易失函数判定口径（2026-11-05 用户定版）**：NOW/TODAY/RAND/RANDBETWEEN/RANDARRAY 每次重算结果都不同，做值对比无意义。口径：**支持且返回非错误值 → 一律记「支持（易失函数，结果不对比）」**（黄色条件格式），不计入一致也不计入不一致、不进差异明细。落地三处必须同步：judge_v2.py（Python 判定豁免）、build_compare_v2.py judge_formula（活公式按 $B 函数名分支豁免）、结论页新增「其中：易失函数」行 + 提示语。
- **文件被占用的"静默丢写"陷阱（2026-11-05 实战）**：目标 xlsx 被别的进程以拒绝写入共享打开（python `open('r+b')` 全共享模式探测不出来！），Excel COM 会**降级只读打开**（`$wb.ReadOnly=True`），写单元格/公式在内存中全部"成功"、setFail=0，但 `Save()` **静默不落盘**（mtime 不变）。排障信号：Save 后 openpyxl 读回看不到刚写的公式。检测方法：用 .NET `[IO.File]::Open($p,'Open','ReadWrite','None')` 试独占——失败即有进程占用（WorkBuddy 预览面板、WPS、Excel 僵尸进程都可能持有）；必须等用户真正关闭持有方（关预览标签页）再跑写入链。**每次 Save 后必须用 openpyxl 回读 mtime+内容验证**，不能只看脚本输出。
- **Excel COM 错误值是负数 int**：`Value2` 读错误单元格得 -2146826273 之类，必须 `-band 0xFFFF` 后映射：2000=#NULL! 2007=#DIV/0! 2015=#VALUE! 2023=#REF! 2029=#NAME? 2036=#NUM! 2042=#N/A 2047=#SPILL! 2050=#CALC!（只 -band 0xFF 会截成 #ERR244）。
- **PowerShell 十六进制字面量陷阱**：`0x80010001` 被解析为**正数** 2147417857，与 COMException.HResult（-2147418111）永远不等 → 重试逻辑永不命中。比较 HResult 一律用十进制 -2147418111。
- **锁屏会话 Excel 自动化**：会话锁定时 `Visible=$true` 的 Excel 无法渲染界面、首启/激活对话框无法关闭，COM 全部 RPC_E_CALL_REJECTED，屏幕截图只能截到锁屏。**无头方案**：`Visible=$false` 全程可行；先开一个无关小工作簿 → `$xl.Calculation = -4135`（手动计算）→ 再开目标簿 → 直接读 `.Text`（缓存值即 Excel 上次真实计算结果）→ 关簿前可对单 sheet `$sh.Calculate()` 验证重算链路。**`CalculateFullRebuild` 在该簿上会永久忙**（60s 重试全被拒），用单 sheet Calculate 替代；"打不开文件"多数是锁屏+可见窗口的假象，先用无头只读模式排除环境问题再怀疑文件本身。
- **openpyxl 写"以 = 开头的纯文本"必须 `data_type='s'`**（快照值如 `=A1+1`、FORMULATEXT 结果），否则变成活公式，Excel 打开时整文件拒开/重算污染。
- **PS1 param 名避开内置变量**（$Host/$Data/$Error 等，管道 ConvertFrom-Json 会被 param 同名遮蔽）；含中文的 .ps1 必须 UTF-8 BOM，纯 ASCII 的不加。
- **用例修正实例**（构造"Excel 能出值"时的参数坑）：FORECAST.ETS* 需 ≥12 点完整时间序列（AK1:AL12）；ISOMITTED 判第二参需尾逗号 `LAMBDA(x,y,ISOMITTED(y))(1,)`；ODDFYIELD/ODDLYIELD **price 在前 yield 在后**；PROB 的概率列需单独数据区且和为 1；NA 是唯一"就该返回 #N/A"的函数。
- **三方语义不同导致假差异的函数直接 EXCLUDE**：SHEET/SHEETS（三方容器结构不同）、CUBE*（需 OLAP）、CALL/REGISTER.ID（宏表）、RTD/WEBSERVICE/STOCKHISTORY（外部服务）、GETPIVOTDATA（透视表）、JIS（需日文语言包）等，Excluded 清单连同原因写进对比表尾部，分母统计必须截到有效行（否则 COUNTIF 合计对不上）。
- **SJS Date→Excel 序列号按本地日历字段推算**（Y/M/D/H/M/S 逐项算），不要用 epoch 毫秒+当前时区偏移——1899 年上海时区是 GMT+08:05:43，偏移法差 5 分 43 秒（TIME 用例实测踩中）。
- **结论页百分号格式套对行**：比率行（支持率/一致率）才设 `0.0%`，计数行误套会显示 "46600.0%"（本轮翻车已修：`for row in (10, 12)` 不是 (9,11)）。

## SpreadJS 侧文档证据优先走 spreadjs-mcp（2026-11-04 提出，2026-11-14 提级为第一通道）
- SpreadJS 相关查证优先用 spreadjs-mcp（search/fetch，@grapecity-software 中文文档库，与 demo.grapecity.com.cn 同源），少走网页搜索。**提级后口径**：任何 SJS 的 API 名 / 签名 / 有无能力判定，第一个动作必须是 spreadjs-mcp `search` → `fetch`；网页搜索、凭记忆、读 npm 导出只能作补充，不得单独作为依据。查法按**功能命名**搜（"动态数组函数""文本操作函数"），不要只搜函数名——官网函数总列表页缺失 11 个数组函数但实际全支持。MCP 查不到 ≠ 不支持，转实测定案。
- **中英文档同源同错位**：中文枚举页把 clusteredColumn(=60) 描述为"直方图"，英文页同样；实测渲染为簇状柱形。不能只引一个页面当证据，遇可疑描述必须跑渲染实测（sjs_histogram_check.js 模式：50 个数值建图看是否分箱）。
- SJS 的 ChartType 数值不是 Excel XlChartType 映射（51=stockVHLC，Excel 51=columnClustered），跨产品对值会全错，只按成员名对。
- 官方"当前不支持的图表类型"清单页可能是过时内容（仍列 boxWhisker/funnel/waterfall 等 5 个实际可创建的类型），引用时须与枚举页+实测三方交叉验证。
- SJS 配置 API 文档化程度高：IChartLegend（fontSize/fontFamily/color/position/layout）、IAxis（format/max/min/scaling.logBase 2-1000 专文）均有完整接口页，判 S1 的文档证据可直接引用接口页 URL。

## 竞品文档时效性（2026-11-04 教训）
- "无公开文档"的结论有效期极短：Univer 官方指南两天内从"30 个类型（无 3D）"扩到"30 2D + 11 3D + setThreeD API"。任何"文档未收录"的判定，落报告当天必须重抓一次官方指南/reference 页原文。
- 给厂商引用链接时用对方中文官方域（如 SpreadJS 用 demo.grapecity.com.cn，不用 developer.mescius.com）——同一内容不同域，客户认定口径不同。
- 报告结构以"用户四问"组织（支持几种/3D 支持吗/配置有无公开 API/导入导出兼容吗），速览放结论、证据放后文；避免多口径长文混排。

## 导入保真必须分层记录（2026-11-04）
- "63/63 可见"是三层里最粗的一层。导入判定拆四层记录：①对象存活 ②直接可见（不手工修复）③类型读回 vs 基线 ④标题/图例文本 vs 基线。客户一定会在第③④层发现问题（如 SJS 新式图表标题变"Chart Title"/系列名、股价图图例丢失），漏写即失信。
- Excel 基线用 COM 逐图导出 ChartType/HasTitle/ChartTitle.Text/HasLegend/Legend.Position/宽高（excel_baseline_meta.json 模式），与产品导入后 API 读回逐图配对，才能出"X 张一致 / Y 张降级"的硬数字。
- SJS preserveUnsupportedChart 开关改变失败模式：true=画布"Unsupported Chart Type"占位框；false=静默降级为柱形。测导入必须两种模式各跑一遍。
- 对比标题差异时区分：自定义标题丢失（真缺陷）vs Excel 自动占位文案变化（"图表标题"→"Chart Title"，占位差异）vs 自动标题取系列名（单系列图）。

## 精度/布尔专项（2026-11-05）
- **（v3 口径提醒）公式模块已定稿为"只判与 Excel 一致"，不在报告里单列 `CALC_PRECISION=15` 精度档，也不做位数归因页**——下面的机制知识仅用于排障，不要据此再生成"设精度后 91.75%"这类数字。
- **SpreadJS 精度差异一行修复**：`GC.Spread.CalcEngine.Functions.CALC_PRECISION = 15`（静态属性，创建 Workbook 前设置，v19.2.5 实测存在可写）。实测 19 个精度差异函数修复 12 个（IM* 复数函数除 IMCOT/IMTAN 全修复）；BESSEL×3/T.INV.2T/TINV 是算法级 ～1e-8 数值差，参数不改变值；BASE/DOLLAR 格式/区域差异不受影响（对照组）。探针 sjs_prec.js 复用 formula_probe.html，注入静态属性后 runFinal 即生效。
- **Univer"布尔 1/0"真相（重要改口径）**：快照 cell={v:1,t:3}，t=3=BOOLEAN 类型标记；getDisplayValue() 返回 'TRUE'/'FALSE'（显示层正确）；下游语义全对（TYPE(ISBLANK)=4、ISLOGICAL=TRUE、ISNUMBER=FALSE、IF(bool=TRUE())）。差异仅在 Facade getValue()/getRawValue() 返回 JS number 与快照 v 字段——是**数据读取约定**不是计算错误，无官方配置可改。竞争话术勿说"Univer 布尔计算错误"，只说"API 取值需按 t===3 判布尔 / 用 getDisplayValue()"。探针 univer_bool.js/univer_bool2.js。
- **Excel COM 无头追加 sheet 三坑**：① `$xl.Calculation` 必须在至少一个工作簿打开后才能设置（先 Workbooks.Add 空簿再切 manual，防止目标簿打开时自动重算 500 公式导致后续调用 RPC_E_CALL_REJECTED）；② 无头模式 `$xl.ActiveWindow` 为 null，FreezePanes/Select 会抛异常且被 finally 的 Quit 异常吞掉——无头脚本禁用窗口类操作、catch 里打印真实错误；③ COM 集合用 for+Item(i) 遍历，不要 foreach。
- **PS1 脚本反复转换的坑**：utf-8-sig 重写已含 BOM 的文件会出双 BOM（\xef\xbb\xbf×2）导致解析器报"赋值表达式无效"；newline='\r\n' 配合已 CRLF 内容出 CRCRLF。安全流程：decode('utf-8-sig').lstrip('\ufeff') → 全文 replace 成 \n → 以 newline='\r\n' 写出单 BOM。
- **JSON 读入 PS 必须 -Encoding UTF8**：中文 JSON 不指定编码会按 GBK 读成乱码，ConvertFrom-Json 报"传入的对象无效"。

## v0.7 修订两坑（2026-11-05）
- **Univer 快照布尔判定必须看 t 标记**：cellData 里布尔单元格是 {v:1,t:3}——v 字段序列化为 1/0，t 才是真类型（经验映射：t=1 字符串、t=2 数字、t=3 布尔，错误值是对象）。只读 v 会把 IS 族 18 个函数误判成"类型不同（布尔/数值）"；正确做法：t===3 → (bool(v),'bool')，与 Excel 布尔比对记一致。全量探针务必同时捕获 getDisplayValue() 与 t 标记（univer_v3.js）。教训：**比较口径里"显示文本"必须真的取显示层，不能拿原始值字段代替**。
- **环境依赖函数也要做支持性实测**：无法构造"Excel 能出值"用例 ≠ 不用测。把公式直接写进产品看是否 #NAME?：#NAME?=不支持；其他任何结果（#REF!、空串、富对象、数值）=引擎认识该函数。SJS 的 EUROCONVERT 能真算出 196、GETPIVOTDATA 返回 #REF!（认识）、IMAGE 返回 rich 图片对象；Univer 仅 IMAGE/SHEET/SHEETS。探针 env_probe.js 双端同跑。

## 宿主页 CSS 污染 Univer 内部 header/footer（2026-11-05，症状极具误导性）
- **现象**：对照台性能视图里 Univer 面板"看不到产品"——顶部一条莫名黑条，网格只画出 9 行就断了，下面一大片空白。第一反应会是"canvas 没跟着 resize"，但实测 canvas **确实跟着容器变了**（容器 −32px，画布也 −32px），只是永远比容器矮 321px。
- **根因**：Univer 在容器内**自己渲染 `<header>` 和 `<footer>`**（即使 preset 传 `header:false, footer:false`——传 false 只是不渲染内容，标签还在）。宿主页面的裸元素选择器 `header{height:52px;background:#1f2937}`、`footer{height:268px}` 精确命中了这两个内部元素：header 被染成黑底 52px，footer 被撑成 268px，sheet 区被挤成 `578−52−268−1=257`px。
- **定量校验式**：`canvas 高 = 容器高 − 页面 header 高 − 页面 footer 高 − 1`。两个视图分别验算（578→257 / 546→225）都能对上，就能确认是"宿主 CSS 泄漏"而非渲染引擎问题。
- **排查动作**：遍历 `#uniHost` 子树打印每层 `getBoundingClientRect().height`，找第一个"比父元素矮"的层，看它是不是 `HEADER`/`FOOTER`。这一步比读 Univer 源码快一百倍。
- **修法**：宿主页所有布局规则改 id 选择器（`#appHd` / `#footer` / `body > main`）；取页面 footer 用 `getElementById`——`document.querySelector('footer')` 会命中 Univer 内部那个（它在 DOM 里更靠前，量出来的 268 是 Univer 的，不是页面的，会误导排查）。
- **同类风险**：任何"Univer 会自己渲染同名标签"的宿主页都适用（`header/footer/main/section/table/button/input`）。一眼判据：面板里出现不属于你设计的横条、或者某块区域高度恰好等于你页面上某个条的高度。

## 表 / 排序 / 筛选 / 数据验证模块（2026-11-05 晚，新模块，含 5 个假 FAIL 翻案）

### 「接口返回成功」≠「能力存在」——排序命令空转
Univer 的表内排序命令 `sheet.command.set-table-sort-state` 执行返回 `true`，但数据一行都没动。源码里 `sortInfo` 只有 `setSortInfo` / `toJSON`，全库没有任何消费方去重排。
**纪律**：判 PASS 必须有"结果值"作证（排序后取值对比），不能拿 `executeCommand` 的布尔返回当证据。任何"命令/接口返回 true"的项，都要补一步"读回数据看有没有变"。

### 注释不可信：`@todo not support yet` 可能是过时的
`ITableOptions.showHeader` 标着 `@todo not support yet`，实测**生效**；同组的 `hasTotalRow` / `showFooter` 实测**确实不生效**。同一份注释里真假混着。
**纪律**：注释只当线索，每个字段都要真实调用一次；测"开关生效"要同页建两个对象（不传参 vs 传参）对比，否则"读回 false"可能只是默认值。

### Univer 能力分「功能包 + UI 包」两层，缺一层会伪装成产品缺陷

| 层 | 包 | 缺了的表现 |
| :--- | :--- | :--- |
| 功能 | `sheets-table` / `-filter` / `-data-validation` | Facade 方法直接 undefined |
| 界面 | `sheets-table-ui` / `sheets-filter-ui` | Facade 全通，但列头不画筛选漏斗、表没有可视化编辑 |

还要 `import '<pkg>/facade'` 与 `import '<pkg>/lib/index.css'`。**做截图证据前必须确认两层都装了**，否则会把"我没装包"写成"产品不支持"（本轮差点犯）。

### 静默失效三例（不报错但结果不对）
1. **Univer 颜色筛选**：内部 `new ColorKit(bg.rgb).toRgbString()` 比对，必须传 `rgb(0,255,255)`（无空格）。传 `#00ffff` 不报错，但一行都匹配不上。
2. **SJS `top10Condition`**：正确参数是 `{ type: Top10ConditionType.top, expected: N }`；误传 `compareType` 不报错，结果永远是"后 N 项"（最小的 N 个），容易被判成"不支持前 N 项"。
3. **Univer 日期验证**：`requireDateOnOrAfter` 必须吃 `Date` 对象，传字符串会在内部 `toLocaleDateString` 处抛错。

### 反过来：SJS 也有"方法名不存在但能力存在"的情况
按颜色筛选没有 `rf.addFilterItemByColor`，要走 `new Condition(ConditionType.colorCondition, {compareType: ColorCompareType.backgroundColor, expected: '#FFC7CE'})` → `addFilterItem` → `filter`。
**纪律**：查 SJS 能力先看 `ConditionType` 等**枚举**，不只看类的方法列表；方法列表在压缩产物里还会被混淆成 `V_r` / `zct` 这类名字，`Object.getOwnPropertyNames` 靠不住。

### 用例写法铁律（本轮修掉 8 条假 PASS）
```js
// 错误：能力缺失被记成 PASS
return '不支持按颜色筛选'; throw new Error('…');

// 正确：先算、先读，把观测值拼进错误消息，只在最后 throw
var vis = [...]; if (vis.join(',') === '乙') return '可见=' + vis.join(',');
throw new Error('可见=' + vis.join(',') + '（期望 乙）');
```

### 造样本：openpyxl 汇总行合法性
只写 `t.totalsRowCount = 1` 而 `tableColumn` 上没有 `totalsRowFunction` / `totalsRowLabel`，Excel 会**直接拒绝打开**（COM 只报"不能取得类 Workbooks 的 Open 属性"，无任何提示）。必须先 `t._initialise_columns()` 再逐列声明。

### UI 入口判定：自动化点不开 ≠ 按钮不可用（2026-11-14，用户当场用截图推翻）

**现象**：探针点 Univer Ribbon 的「排序」按钮，数据不变、无菜单、无报错，据此写成"普通表格在 UI 上无排序入口"。用户随手一点，弹出 5 项菜单：当前区域升序 / 拓展区域升序 / 当前区域降序 / 拓展区域降序 / **自定义排序**。整节结论作废。

**四个独立原因叠在一起**，任何一个单独出现都会让人误判：

| # | 原因 | 表现 | 正确做法 |
| :-: | :--- | :--- | :--- |
| 1 | **split button 主按钮与箭头是两个热区** | 点 rect 中心落在主按钮上，不弹菜单 | 点 `rect.x + rect.w - 6`（右端），或分两次分别验证 |
| 2 | **标签在 `aria-label` 而非 `title`** | `[title="排序"]` 命中 0 个 | `querySelectorAll('[title], [aria-label]')`，取值 `title \|\| aria-label` |
| 3 | **菜单 portal 到 `body`** | `#uniHost *` 里查不到菜单项 | 查询根节点用 `document.body *`，再用坐标/文本过滤 |
| 4 | **菜单在 pointerdown 展开、pointerup 关闭** | `mouse.click()` 的 down/up 间隔太短，点完即关 | `move → down → wait 600~900ms → （截图/查询）→ up`，或 `down → move(菜单项) → up` |

**纪律（写进流程，不靠记性）**：
1. 判"某产品 UI 上没有 X"之前，先 `dump` 出当前 tab 下全部 `[title]`/`[aria-label]`，**用清单说话**，不凭印象列按钮。本轮就是清单里明明有「筛选」「排序」，却被前一轮的错误印象覆盖了。
2. 按钮 `disabled=false` 且 `cursor:pointer` 就默认"入口存在"，能力另测；`disabled=true` 要换选区/换页面再验一次（可能是上下文条件，不是永久禁用）。
3. **自动化点不开时，必须补一次人工点击或用产品文档核实**，不能把"我的脚本没点开"写成"产品没做"。这类误判与"拿 SJS 的 `sortRange` 去找 Univer 的排序"是同一类错误：**用 A 的交互方式否定 B 的能力**。
4. 截图要覆盖"菜单展开态"。只截主界面会看不出菜单/浮层，反而给错误结论背书。

**同轮另两个坑**：
- `FWorksheet` 上没有 `getTable()`（`typeof sh.getTable === 'undefined'`），不能靠它取表对象；表内判定要从 `FRange`/表模型侧走。
- `getRange().createFilter()` 的 `getRange()` 返回 `undefined`，但筛选**真实生效**（隐藏行可逐行验证）。**Facade 的返回值不可信，一律以结果数据为准**。

**结局（人工截图补齐）**：「自定义排序」对话框实测支持「+ 添加排序条件」（多列分层）与「标题不参与排序」勾选；排序依据下拉只有列值，无按颜色、按图标，也无 Excel「选项」里的自定义序列与按行排序。由此得出一条通用判据——**UI 与 API 的覆盖面可以不一致，必须分别取证**：Univer 的「排除标题行」UI 有勾选框、API 的 `hasTitle` 却固定为 `false`，只测一层就会写错结论。凡遇「某能力 API 不支持」，都要再问一句「UI 上有没有」。反之亦然。

## 数字格式 / 格式化模块（2026-11-05 晚，用户当场指出 SJS 侧代码是错的）

### 「设置-呈现」两段式能力，判据必须落在呈现结果上
数字格式、条件格式、样式、边框这类能力都是"先 set 再渲染"。
如果判据取 `getFormatter()` / `getNumberFormat()` 的返回值，那测的只是**格式串能不能存进去**——存什么读回什么，永远相等，永远 PASS。
```js
// 错：往返一致 = PASS，掩盖一切渲染缺陷
sheet.setFormatter(r, c, '[DBNum2][$-804]G/通用格式');
return sheet.getFormatter(r, c);          // 原样返回，看不出 SJS 渲染成 "/通用格式"

// 对：判显示文本
sheet.setFormatter(r, c, '[DBNum2][$-804]General');
return sheet.getText(r, c);               // 壹仟贰佰叁拾肆.伍陆柒捌
```
Univer 侧对应 `FRange.getDisplayValue()`。期望值必须来自 Excel 实测，不能凭格式串推。

### 格式串两家不通用（会互相"静默降级"）

| 场景 | Excel 中文版 | SpreadJS | Univer |
| :--- | :--- | :--- | :--- |
| 中文大写 | `[DBNum2][$-804]G/通用格式`（COM **拒绝** `General`） | 只认 `[DBNum2][$-804]General`，喂 `G/通用格式` 渲染成 `/通用格式` | 两种都不认，显示原数字 |
| `!` 转义 | `0!.0,"万"` → `0.1万`（`!` 让 `.` 成字面量、两个 0 都算整数位） | `1!.2万`（当普通字符） | `1!.2万` |
**纪律**：测"某串支不支持"要分开测"我方产品文档写法"和"Excel 存盘写法"，混在一起会把写法错误算成产品缺陷。

### SpreadJS 默认 culture 是 en-us
`aaaa` 渲染成 `Monday`。必须先 `GC.Spread.Common.CultureManager.culture('zh-cn')`（v19 内置，不需要额外资源包），否则中文日期对比必然出假差异。
同理，拿到"某侧不支持中文"的结论前，先确认区域设置 / 资源包 / UI 包这些**环境前提**。

### 演示截图里的 `#########`
SpreadJS 列宽不足会整格渲染成 `#`。对比两家的格式效果前必须显式 `setColumnWidth`，
否则两侧默认列宽不同（SJS 72px / Univer 略宽），同一串格式会呈现不同结果，截图不可用。

### Excel COM 取基线的两个 locale 坑（中文版）
1. `NumberFormat` 只认中文颜色名 `[红色]` / `[蓝色]`；传 `[Red]` 直接抛「不能设置类 Range 的 NumberFormat 属性」。但**文件里**存的 `[Red]` 能正常打开 —— 是 COM setter 的 locale 限制，不是 Excel 不支持。
2. `General` 也被拒（要写 `常规`）。取基线前先把格式串做一遍"英→中"映射，否则 oracle 本身是错的。


## 表模块复审修正（2026-11-05 晚，用户连提两条）

### 同族方法名只差一个词，漏一个就把自家能力写没了
SpreadJS `GC.Spread.Sheets.Tables.Table` 上有两个极易混淆的方法：

| 方法 | 真正作用 | 示例 |
| :--- | :--- | :--- |
| `setColumnDataFormula(index, formula)` | **表内计算列**（数据区公式，整列自动填充，新增行继承） | `t.setColumnDataFormula(2, '=[@数量]*[@单价]')` |
| `setColumnFormula(index, formula)` | **页脚汇总公式**（totals row） | `t.showFooter(true); t.setColumnFormula(1, '=SUBTOTAL(109,[金额])')` |

本轮先只查到 `setColumnFormula`，就下了"SpreadJS 没有列级计算列 API，只能逐行 `setFormula`"的结论 —— 错，而且是把自家产品判低了（用户用 spreadjs-mcp 纠正）。
**纪律**：①查 API 要看整族方法（`setColumn*` 全列一遍），只看名字猜语义必翻车；②官方示例里看它给**哪一行**赋公式——落在数据区还是页脚行，一眼就分清；③涉及"我方有没有"的结论，必须至少有一次运行时实测，不能只凭文档印象。
对照：Univer 侧是 `setTableColumnFormula(tableId, columnId, formula)`，`columnId` 得先从 `getSubTableInfos()` 取；SpreadJS 直接用列索引，且实测新增行自动继承。**两家都有，差别在易用度**。

### 报告统计一律写通过率
用户明确要求：不要用"X 通 Y 不通""X/Y 不通"做主指标，**所有比率统一为通过率 = PASS ÷ 可判项**，分组表列头写「通过 / 通过率」，差额用 pp。
反面样例（本轮初版踩过）：
- `Univer 28 通 18 不通（60.9%）` —— 60.9% 是通过率，但挨着"18 不通"读起来像失败率；
- `### B 组 表样式：Univer 2/4 不通` —— 分母还写错了（该组 6 条用例，2 通过 → 33.3%）。
生成脚本（`gen_api_report.py`）里的表头也一并改，否则重新生成会把旧口径带回来。

### openpyxl 造表三个坑（本轮全部踩到）
1. **计算列必须写全限定结构化引用**：单元格里写 `=[@数量]*[@单价]`，Excel 打开整列 `#REF!`。xlsx 里存的是 `TblSales[[#This Row],[数量]]*TblSales[[#This Row],[单价]]`，`[@…]` 只是编辑栏的显示形式。同时在 `tableColumn` 上写 `<calculatedColumnFormula>`，Excel 才认它是计算列（新增行自动继承）。
   ```python
   t._initialise_columns()                       # 否则列名是 Column1/Column2
   t.tableColumns[3].calculatedColumnFormula = TableFormula(attr_text=CALC)
   ```
2. **不写缓存值**：openpyxl 只输出 `<f>`、不输出 `<v>`。Excel 会重算（`fullCalcOnLoad="1"`），但**读缓存值的引擎（SpreadJS / Univer 导入、在线预览）看到 0 或空**。收尾必须用真 Excel `CalculateFullRebuild` + `Save` 一遍。
3. **`[#This Row]` 只能在表内用**。放到表外单元格（如 A18 写 `=SUM(TblSales[[#This Row],[金额]])`）必 `#VALUE!`——它不是"表外引用写法"，别拿它做外部引用用例。

### PS 脚本的编码坑（第二次踩）
`-File x.ps1` 方式按 ANSI 码页（中文机器是 GBK）读脚本。UTF-8 存的中文路径会变成乱码，`Workbooks.Open` 直接找不到文件。
**解法**：这类脚本一律纯 ASCII，路径走 `C://temp//xxx.xlsx` 这种临时位置，中文路径的搬运放在 Python 侧做（见 `scripts/finalize_sample.ps1`）。

## 打印 / PDF 导出模块（2026-11-05 深夜，含三个致命假阴性）

### `all` 包不含 charts / shapes / print / pdf —— 缺插件伪装成"不支持"

`gc.spread.sheets.all.min.js` 只是核心，**图表、形状、打印、PDF 各是独立插件**。缺 print → `sheet.printInfo is not a function`；缺 charts/shapes → `sheet.charts` 为 `undefined`，此时 **`spread.import()` 遇到 xlsx 里的图表会静默丢弃**，统计出来 `charts:0 shapes:0`，看起来像"产品不支持导入图表"——本轮差点据此写错结论。

```html
<!-- 顺序有硬依赖：charts 内部引用 Shapes.ShapeBaseData，必须 shapes 先、charts 后 -->
<script src="vendor/gc.spread.sheets.shapes.min.js"></script>
<script src="vendor/gc.spread.sheets.charts.min.js"></script>
<!-- print 先、pdf 后 -->
<script src="vendor/gc.spread.sheets.print.min.js"></script>
<script src="vendor/gc.spread.sheets.pdf.min.js"></script>
```

顺序错了报 `Cannot read properties of undefined (reading 'ShapeBaseData')`。
**判据**：任何"浮动元素为 0"的统计，先断言 `typeof sheet.charts !== 'undefined'`，再谈结论。

### `spread.import` 是回调式，`await` 它不会报错但什么也不导

v19 签名 `import(file, successCallback, errorCallback, importOptions)`，返回 `void`。写 `await sp.import(blob, opts)` 拿到 `undefined`，**没有异常**，页面只剩一个空的 `Sheet1`（200×20）——这是本轮最大的假阴性来源，浪费了一整轮。

```js
await new Promise((res, rej) => {
  sp.import(new Blob([buf]), res, rej, { fileType: GC.Spread.Sheets.FileType.excel });
});
```

同类陷阱：`sh.getComment()` 在 v19 已移除，读批注用 `sh.comments.get(r, c)`。
**通则**：凡"导入了但结果是空表/空数据"，先查签名是否回调式，别急着怀疑文件。

### SJS 导 PDF 中文默认乱码，必须注册 TTF

PDF 只嵌 14 种西文标准字体，不注册就是 `VÞ'h´_br¶`。修法（**只认 `.ttf/.otf`，`.ttc` 不行**——`simsun.ttc` / `msyh.ttc` 会失败）：

```js
const fb = await fetch('/fonts/simhei.ttf').then(r => r.arrayBuffer());
const F = GC.Spread.Sheets.PDF.PDFFontsManager;
F.registerFont('SimHei', { normal: fb, bold: fb });
F.fallbackFont = () => fb;          // 兜底，防止漏配的字体族走默认
```

注册后 PDF 21 KB → 31 KB（只嵌用到的字形子集）。**这是配置成本不是缺陷**，但交付/报价必须提前说明，否则客户第一次导出就是乱码。

### fitToPage 不缩放浮动元素（唯一真实缺陷）

Excel 里 `FitToWidth=1` 会把整页含图表缩到一页宽；SJS 读回 `fitPagesWide=1` 也对，但导出 PDF 里**表格区域缩放、图表按原始尺寸绘制，右侧被页边裁掉**。规避：按目标纸宽反算图表尺寸，或把图表压在打印区域内。

### printInfo 的 getter/setter 二合一 + 两个参数陷阱

`pi.orientation(2)` 设、`pi.orientation()` 读，共 56 个方法（35 项打印语义）。两个坑：

1. **`paperSize(8)` 传裸数字不生效**，读回仍是 Letter `{width:850,height:1100,kind:1}`。必须传对象 `{width:842,height:1191,kind:8}` 或 `PaperKind.a3`。
2. **`zoomFactor` 收小数**：`0.8`=80%；传 `80` 被钳到 `4`，不报错。

读回比较也有坑：`paperSize` 读回对象带额外内部字段，`Object.keys(a).every(k => same(a[k],b[k]))` 只校验显式设置的键，再加 JSON 序列化兜底双通道比对，否则会误判 FAIL。

### Univer 打印无 getter，只能靠输出物反推

`getPrintConfig` / `getPrintRenderConfig` **都不存在**。验收降级为：设置不抛错 + 输出物形态变化（A3 纵向 1123×1587、`maxRowsEachPage=10` 时 40 行 → 5 页）。这对"配置是否真落库"是硬伤，必须在报告里单列为已知边界。

`print()` 会离屏渲染到 `.printing-canvas-container` 里的 canvas，可在 `window.print()` 前轮询抓这个 canvas 得到"打印预览"图。

### Univer 页眉页脚：占位符语法不同，且缺一个开关整块被忽略

| 用途 | Excel / SJS | Univer |
| :-- | :-- | :-- |
| 当前页 | `&P` | `@Page` |
| 总页数 | `&N` | `@TotalPage` |
| 日期 | `&D` | `@DateA` |
| 工作表名 | `&A` | `@WorksheetTitle` |

传 `&P/&N` 输出就是字面量 `&P / &N`（不报错）。且**必须同时传 `isCustomHeaderFooter: true`**，否则 `headerFooterSetting` 整块被忽略，只剩 `headerFooter` 枚举（页眉只显示 sheet 名）——症状是"我配的页眉没生效"，实际是缺开关。

### Univer 的打印是 Pro 付费插件

`@univerjs-pro/sheets-print`，无授权输出带水印。OSS 侧无图表/形状/图片/批注/迷你图，所以"打印保真"在 OSS 缺的不是打印能力，是被打印的对象——**做保真矩阵时必须把这类项标"无此元素"而不是 FAIL**，否则统计会被拉成误导性的低分。

### 打印样本造数：COM 只能补 openpyxl 写不了的，且自身不稳

openpyxl 能写数据/样式/条件格式/批注/超链接/打印设置，写不了形状与迷你图，这两个用 COM 补。COM 侧两个坑：

1. `SparklineGroups.Add(1, "Data!B4:E4")` 报 `0x800A03EC`：①必须 `$sheet.Activate()`；②Location 必须**单格**；③不接受多行区间 → 逐行循环 `Add(1, "Data!B${i}:E${i}")`（`${i}` 花括号必需，`"$i:"` 会被解析成 PS 盘符）。
2. `fitToHeight = 0` 时饼图会被切到下一页，设 `1` 才压进一页。造完先用 Excel 真机导一次 PDF 数页数，别信"应该是一页"。

收尾清理：脚本异常退出会留无窗口 Excel 进程，`Get-Process EXCEL | Where MainWindowHandle -eq 0 | Kill`。


## 工作簿 / 工作表层造样本（2026-11-06，用户定版"样本先行"后第一个模块）

### openpyxl 写的批注，Excel 读成 1 个字符、重存时整条丢失
症状链：`Range.Comment.Text().Length` 读回 1；`SaveAs` 之后 `xl/comments*.xml` 与 `.vml` 全部消失。openpyxl 的 `Comment` 与 `shapeId` 分配让 Excel 只建了空文本框。
**解法**：批注一律由 COM `Range.AddComment(text)` 写。文本不能写进 PS 脚本（见下条），落成 UTF-8 文件后 `[IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)` 读入。
**推论**：凡"Excel 要重存"的样本，openpyxl 只负责单元格/样式/结构，**任何带 VML 或 DrawingML 关联的东西（批注、形状、图片）都交给 COM**。

### `CustomDocumentProperties.Add()` 在没有 custom.xml 时抛 NullReferenceException
空引用不是权限问题，是包里还没有 `docProps/custom.xml` 这个部件，COM 集合拿不到宿主。
**解法**：Python 在包级注入三件套，再让 Excel 重存：①`docProps/custom.xml` 本体（`<Properties>` + `<property fmtid=... pid=2 name=...>`）②`[Content_Types].xml` 加 `Override PartName="/docProps/custom.xml"` ③`_rels/.rels` 加 `Type=.../custom-properties" Target="docProps/custom.xml"`。缺任一件 Excel 都会报修复。

### Excel COM 的读回值有四个"看着像缺陷其实是我读错"的坑
- `Outline.SummaryRow`：`0` = 汇总在**上**（`xlSummaryAbove`），`1` = 在下。与直觉相反，判"summaryAbove"要断言 0。
- `Rows.Item(n).OutlineLevel`：未分组行返回 **1** 而不是 0，整表有 +1 基线。只能判**相对差**（`r3 - r1 == 1` / `r7 - r1 == 2`），或直接看 XML 的 `outlineLevel`。
- `Comment.Text` 是**方法**：写成 `.Text.Length` 得到 `$null`（断言静默变 False），必须 `.Text().Length`。
- PS 5.1 无法 late-bind `BuiltinDocumentProperties.Item('Title')`（NullReferenceException）。docProps 一律用 Python 读 `docProps/core.xml` 断言，不要走 COM。

### Excel 365 打开文件会注入隐藏定义名称 `_xlfn.SINGLE`
指向 `#NAME?`，导致"定义名称数量"凭空 +1。断言写成"我设的两个名字都在"，不要写 `Count == 2`。

### OOXML 的保护语义是反的，"禁止"根本不写属性
`<sheetProtection sheet="1" formatCells="0" insertRows="0" sort="0"/>`——属性值 `0` 才是**允许**；禁止项直接**缺省**（OOXML 默认 true）。所以包级断言要分两组：允许项断言 `xxx="0"` 存在，禁止项断言 `xxx=` 不出现。
单元格锁定也不在 sheet XML，而是 `styles.xml` 的 `<protection locked="0"/>`；断言"解锁区存在"要查 styles。
密码落成旧式 4 位十六进制 `password="DE3A"`，不是 `hashValue`。

### 其他两条
- `SaveAs` 必须反斜杠路径（`C:/temp/x.xlsx`）。正斜杠会报 `不能访问文件"C:///temp/BBF31320"`，看起来像权限问题。
- 列宽往返有字体度量误差（设 26 → Excel 读 25.22）：断言给容差（25～27），不要等值。

## 工作表 / 工作簿 API 探针（2026-11-06，127 项双端跑批）

这一节全部是"**我写错 API 导致把『产品有』判成『产品没有』**"的实例。首轮 47 秒跑出 19+10 处 mismatch，七成是写法问题，不是产品问题。

### SJS `spread.search()` 只返回首个命中，不是结果集
文档印象里像 `foundCount` + `getFoundCell`，实测返回的是**一个扁平对象** `{foundRowIndex, foundColumnIndex, foundSheetIndex, foundString, searchFoundFlag}` —— 没有计数、没有迭代器。
**影响**：判"查找全部"要自己多次改写 `startSheetIndex` + 已搜范围反复调；报告里不能写"支持查找全部"。
附带：`searchString` 传 **`RegExp` 对象不抛错、也不命中**（读回 -1）——这是产品 bug，我在报告里判"未达成"，不判"无 API"。

### SJS `rowOutlines` 三处语义错位
- `find(row)` 不是"查某行属于哪个组"，返回 `null`。要读 `getLevel(row)`（**0 基**，未分组返回 **-1**）。
- `group(2,5)` 后 `getLevel(3) === 0`（第一级是 0，不是 1）。
- `ungroupRange` 后 `toJSON().itemsData` 里的条目**还在**，只是 `info.level` 变 `-1` —— 判"已解散"要看 -1，不能看条目是否消失。

### SJS `setValue()` 不进撤销栈
API 直写不算一次可撤销操作：`getUndoStack().length === 0`，`undo()` 无效。必须 `commandManager().execute({cmd:'editCell', sheetName, row, col, newValue})`。**这一条对客户是硬约束**："代码批量写值 + 用户 Ctrl+Z 撤回"必须统一走命令通道。

### SJS 事件绑得上，但 API 直调不派发
`ActiveSheetChanged` / `SheetChanged` / `RangeGroupStateChanged` 用 `spread.setActiveSheet()` / `addSheet()` / `changeSheetPosition()` 调用时实测触发 **0 次**，只在 UI 路径（点标签、点 +/-）触发。事件名全在 `SJS.Events` 里、`bind` 也返回成功 —— 所以**不能因为"读到 0 次"就判"没有事件"**，要判"绑得上、API 路径不派发"，并在报告里写成行为差异。

### Univer 侧五个方法名/语义错位

| 我写错的 | 正确 |
| :-- | :-- |
| `sel.getRange()` | `sel.getActiveRange()`（另有 `getActiveRangeList()` / `getCurrentCell()`） |
| `range.getRowCount()` | **没有**；`FRange` 用 `getRow()` / `getColumn()`（`IRange` 才有行列数） |
| `executeCommand('sheet.command.set-range-values', {cellValue, range})` | 参数键是 **`value`**，写成 `cellValue` 抛 `Cannot read properties of undefined (reading 'startRow')` |
| `createWorkbook(snapshot)` 直接用 `wb.save()` 的返回值 | 必须先 **`delete snap.id`**，否则 `cannot create a unit with the same unit id` |
| `setFrozenRows(2,5)` 以为冻结 3 行 | 含端点是 **4 行**；`getFrozenRows()` 返回**起始行 6**，行数读 `getFreeze().ySplit` |

`wb.createSheet()` **不存在** —— 建表是 **`insertSheet(name)`**（`FWorkbook` 方法面里没有 create）。

### Univer `autoResizeRows` / `autoFitRow` / `setRowAutoHeight` 语义不同
`autoResizeRows` 对**换行文本不撑高**（读到 0），`autoFitRow` 才撑。三套方法不能混用；自适应前后各加一次 `tick(150)` 再读。

### 漏装插件 = 方法 `undefined`，会被判成"没这个能力"

| 能力 | 需要的包 | 入口 |
| :-- | :-- | :-- |
| 工作表背景图 | `@univerjs/sheets-drawing` + `/facade` | `setBackgroundImage` |
| 焦点行列高亮 | `@univerjs/sheets-crosshair-highlight` + `/facade`（**不在 Core Preset**） | `univerAPI.setCrosshairHighlightEnabled` |
| 查找替换 | `@univerjs/find-replace` **和** `@univerjs/sheets-find-replace` + `/facade`（内核 + 表格适配，缺一个入口就不存在） | `univerAPI.createTextFinderAsync` |
| 行/列分组 | `@univerjs-pro/sheets-outline`（**Pro 专属**，OSS 无） | `sh.addRowOutline` |

find-replace 装齐后入口与 10 个方法齐全，但 `findAll()` 返回 0 —— 查 `node_modules/@univerjs/sheets-find-replace/lib/es/facade.js` 确认逻辑是 `if (!this._state.findCompleted || !this._complete) return []`，而 `FindReplaceModel.start()` 依赖的 `IFindReplaceService.getProviders()` 在 OSS bundle 里没注册。**这属于装配问题，不能判 FAIL**，记 `ENTRY_OK` 并在 note 里写明"真实命中待完整 Pro 应用复核"。

### 线上 `office.univer.ai` 已改版（2026-11-06）
旧地址 `/zh-CN/playground/sheets` 返回 **401 / "Showcase not found"**。新结构：
- 示例页 `/zh-CN/showcase/sheets/<slug>`（198 个 sheets 示例）
- 编辑器实体在 iframe → `/zh-CN/playground/sheets/<slug>`
- **直接打开 playground 深链**（如 `/zh-CN/playground/sheets/outline`）才能拿到 `window.univerAPI`

线上页首屏常 20～40s 才挂上 API，且**会偶发不挂**。脚本要写"等 45s → reload 一次 → 再等 60s"，否则会把"页没加载完"当成"环境不支持"。另外每个示例页只装该例所需插件 —— **在某页上探不到某方法，不构成该能力的缺失证据**；只能反过来用（在对应示例页上探到了 = Pro 有）。

### Pro 的保护模型不叫 `protect()`
`FWorksheet` 上没有 `protect`。正确入口是 **`sh.getWorksheetPermission()`** 返回 `FWorksheetPermission`，方法为 `protect` / `unprotect` / `setPoint` / `protectRanges` / `setReadOnly` / `setEditable` / `canEdit` / `canEditCell` / `getPoint`。工作簿级同理 `wb.getWorkbookPermission()`。**没有密码概念**。权限点位：工作表 22 个（`WorksheetPermissionPoint`）+ 工作簿 22 个（`WorkbookPermissionPoint`）。

### Pro 的透视表在 FWorkbook / FUniver，不在 FWorksheet
`sh.generatePivotTable` 是 `undefined`；正确入口 `wb.addPivotTable(...)`、`wb.getPivotTableById(...)`、`univerAPI.generatePivotTable(...)`。**按 `sh.*` 探会得出"Pro 无透视"的错误结论。**

### Pro 的行分组只有 6 个方法
`addRowOutline` / `addColumnOutline` / `removeDimensionOutline(id)` / `setDimensionOutlineCollapsed(id,bool)` / `clearDimensionOutlines(axis,from,to)` / `getDimensionOutlines(axis)`。outline 对象字段只有 `{id, unitId, subUnitId, axis, start, end, collapsed}` —— **无 `level`、无 `summaryBelow`**，层级靠区间包含关系推导。`univerAPI.Event` 52 项里**没有任何 outline/group/dimension 事件**。嵌套实测 8 级无问题（`DimensionOutlineErrorReason.MAX_DEPTH` 常量存在但未触发）。

## 打印 / PDF 字段矩阵（2026-11-06，SJS 35 项 vs Univer Pro 19 字段 / 35 用例）

### Univer Pro 的打印配置**能读回**，只是不走 getter
`getPrintConfig` / `getPrintRenderConfig` / `getLayoutConfig` / `getRenderConfig` 四个全 `undefined`，看起来像"没有读回通道"——**这个结论是错的**。真正的读回口是事件载荷：

```js
univerAPI.addEvent(univerAPI.Event.BeforeSheetPrintConfirm, (e) => {
  window.__cap = { layout: e.layoutConfig, render: e.renderConfig };
  e.cancel = true;             // 拿到就取消，不进入实际渲染
});
wb.updatePrintConfig({ paperSize: univerAPI.Enum.PrintPaperSize.A3 });
wb.print();                    // 必须触发一次，配置才会从事件里出来
```

`SheetPrintConfirmed` / `SheetPrintCanceled` / `BeforeSheetPrintCanceled` 同样带 `{layoutConfig, renderConfig}`。原理：`updatePrintConfig` 内部就是 `injector.get(ISheetPrintManagerService).updateLayoutConfig(cfg)`，事件处把服务上的 `layoutConfig` / `renderConfig` getter 塞进载荷。
**别去 hack injector**：`univerAPI._injector.dependencyCollection.dependencyMap` 里的键是压缩后的类，按名字匹配不到服务；走事件才是正路。

### `updatePrintConfig` 会静默吸收不存在的字段
`{ __bogusField: 'X' }` 不报错，且**原样出现在读回的配置里**。所以：
- "设了没报错" ≠ 支持；
- "读回有这个键" ≠ 支持。

判定必须以类型定义里的字段表为准（`@univerjs-pro/sheets-print` 的 `ISheetPrintLayoutConfig` 12 + `ISheetPrintRenderConfig` 7 = 19 个），再逐个实测。类比：这跟 SJS `printInfo` 上不存在的字段会直接抛错，是完全不同的失败模式。

### `pageSizeCustom` 是坏的：会改写调用方对象
传 `{w:500,h:700}`，函数返回后**你的对象就变成 `{w:394,h:588}`**（固定减 106/112，随页边距/页眉页脚变化），连续 `print()` 还会累减成负数（实测 −3556）。PaperType 枚举里也没有 `Custom` 项。结论：Univer 没有可用的自定义纸张尺寸入口，只能选 11 种标准纸。

### 两家的读回语义不同，比较函数要分家写
- SJS：`printInfo` 是 getter/setter 二合一，`pi.x(v)` 设、`pi.x()` 读，**未知字段直接抛错**。
- Univer：只写 + 事件读回；**对象键序会变**（`{top,right,bottom,left}` vs `{top,bottom,left,right}`），比较前必须按 key 排序，否则会误判成 FAIL。本轮就栽过一次，误报了 `marginCustom`。
- 数值钳制也不同：SJS `zoomFactor` 传 `80` 被钳到 4（不报错），Univer `customScale` 传 `800` 原样接受（也不报错）。**两边都不报错，但语义完全相反**，不要拿"没抛错"当"设置成功"。

### 输出侧验证别用页数，用画布尺寸
打印预览的画布会残留在 DOM 里，页数会累加（1→2→4→8），按像素去重又会把内容相同的页合并掉。**可靠的硬指标是画布尺寸**：A4 794×1124 → A3 1123×1587、横向 1124×794。改纸张/方向后尺寸变 = 生效；页数需要干净环境（每次重开页面）才能用。

### 字段能设 ≠ 输出正确（打印模块最典型）
Pro 上 `freeze: [PrintFreeze.Row]` 读回完全正常，但第 2 页仍然没有表头；图表 `insertChart` 成功、编辑器里显示正常、打印出来只有圆角外框。**合同里"API 支持"和"功能可用"必须分开写**，否则验收时被输出打回。

### SJS ↔ Univer 打印字段对位速查
- 对等 20 项：打印区域（`subUnitIds.range` 四个角）、网格线、方向、纸张、适配页宽/页高（`scale: FitWidth/FitHeight`）、缩放比（`scale: Custom` + `customScale`）、页眉页脚六槽、页眉页脚结构、页边距、居中（`hAlign`/`vAlign`）、水印。
- 只剩粗粒度 4 项：重复标题行/列只能 `freeze: [Row]` / `[Column]` 整行整列，**不能指定第几行**。
- 完全没有 11 项：外边框、黑白打印、页序、页码范围、起始页码、首页不同、奇偶页不同、页眉图片、最佳列宽、最佳行高、打印质量。

---

## 图表模块（2026-11-06，103 项 PASS/FAIL 探针）

### Univer 新建的表面灌数据，`sheet.command.set-range-values` 会静默失败
在 `wb.insertSheet('__chart')` 出来的新表上执行：

```js
await univerAPI.executeCommand('sheet.command.set-range-values',
  { value: vals, range: { startRow: 0, startColumn: 0, endRow: 6, endColumn: 2 } });
// 之后 sheet.getRange('A1').getValue() === null  ← 数据根本没写进去
```

**后果极隐蔽**：图表照常创建成功（`getId()` 有值、`dataSource.range` 还是你设的 A1:C7），但 `config.series` 是空数组 → 所有「系列样式 / 数据标签 / 标记 / 次坐标轴」类用例读回全是 `undefined`。**一次误操作能让 20 条用例集体假 FAIL。**

正确写法：`sheet.getRange('A1:D7').setValues(vals)`（走 Facade 的 FRange，带 unitId）。跑批前务必回读一格断言：`init()` 里返回 `seedB2: sheet.getRange('B2').getValue()`，非 undefined 才继续。

### 饼图/圆环的 builder 上没有 `setSeries`
`sh.newChart(T.Pie).setSeries(...)` → `TypeError: b.setSeries is not a function`。饼图的标签入口是 **`setPieLabel({visible, position, contentType, fontSize, color, family})`**（扁平对象，不要再包一层 `{label:{...}}`），读回路径是 `getInfo().config.pie.label`。`contentType` 是位掩码：CategoryName 2 / SeriesName 4 / Value 8 / Percentage 16。

饼图其余专属入口：`setDoughnutHole(0.4)`、`setExplosion(0.2)`、`setRosePie(true)`、`setHalfPie`、`setPieComposite`、`setLabelLineVisible`。

### 判断「有没有某个方法」必须用对 builder
`setThreeD` 只挂在 **3D 图型**（`newChart(T.Column3D)`）的 builder 上，拿 2D `newChart(T.Column)` 去判断会得到「没有」→ 造出假 FAIL。同理 `setSeriesAxis` / `setSeriesType` 只在 `Combination` builder 上。**判断能力面时，builder 要从实际要用的图型取。**

### Univer 改图表位置要走实例方法，改尺寸两条路都行
```js
const b = chart.toBuilder(); b.setPosition({row:3, column:6});
await chart.update(b.build());     // ← anchor 仍是 {row:1,column:4}，静默无效
chart.setPosition({row:3, column:6});  // ← 正确，anchor 立刻变 {row:3,column:6}
```
尺寸倒没这问题：`toBuilder().setSize(500,240) + update()` 与 `chart.setSize()` 都生效。

### `exportImage()` 有时序，且空结果会被缓存
图表首帧没渲染完就调 `chart.exportImage()` → 返回 `undefined`，**再重试也还是 undefined**（同一个结果被缓存）。必须等 ≥500ms（实测 700ms 稳）；成功时返回 40KB+ 的 `dataURL` 字符串。判 PASS 用 `typeof r === 'string' && r.length > 500`。

### Univer 系列只存「补丁」，不是完整状态
`getInfo().config.series` 里没被设置过的系列只有 `{selector:{index:2}}`，没有任何样式字段。取值必须 `pick = (c,i) => (c.getInfo().config.series||[]).find(x=>x.selector?.index===i) || {}`，把 `undefined` 当「未设置」，别当「读回失败」。`setSeries` 的 selector 传 `1` 或 `{index:1}` 都可以。

### 图例关闭时 `config.legend` 存的是布尔 `false`
不是 `{visible:false}`。所以 `P(i,'config.legend.visible')` 会得到 `undefined` → 误判成「读不回」。正确判法：`const v = cfg.legend; got = v === false ? false : v && v.visible`。另外**只传 `{visible:true}` 是 no-op**（读回 undefined），要带 position / 字体等其它字段才会落成对象。

### Univer 图表字段面的硬缺口（类型定义确认，不是漏测）
- `IChartAxisSpec`：无 `visible` / `majorUnit` / `minorUnit` / `crossPoint` / `minorGridLine`（只有 `gridLine`）；轴标题只有 `titleStyle: IChartFontSpec`，**无 rotate**。
- `IChartTitleOptions`：无 `background`。
- `IChartLegendOptions`：无 `background` / `border`。
- `IChartSeriesLabelSpec`：无 `backgroundColor` / `separator`；**`family` 字段在类型里有，但运行时被丢弃**（同批 `fontSize` / `color` 正常落库，轴标签的 `family` 又正常——只在系列标签这条路上丢）。
- `IChartAppearanceSpec`：只有 `background` / `border` / `font`，**无 plotArea 样式**；`IChartPlotAreaSpec` 只是布局矩形（left/top/width/height）。
- `IChartTrendlineSpec`：无 `forward` / `backward`。
- `FChart` 无 `setName` / `getName`（图表不可命名）；`setZOrder()` 有 setter 但 `getInfo()` 无层级字段（只有 config/dataSource/position/size/anchor）。

### SJS 图表侧四个「字段在但赋值被忽略」
`title().bold`（整对象赋值也仍 false）、轴标题 `bold`（对象里根本没这字段）、`majorGridLine.dashStyle`（试过 `dashStyle` / `style` / `lineType` 三种都读回 0）、`legend` 的 `bold`。另外 `series().set()` 会把 `errorBars` 丢掉（读回 null）。

### SJS 图表侧三个 API 形态坑
- `chart.position({x,y})` **只读**，写入要用 `chart.x(120)` / `chart.y(90)`。
- 轴标题必须**整对象赋值一次**：`a.primaryValue.title = {text, fontFamily, fontSize, color, textDirection}` → `chart.axes(a)`；之后单属性赋值才生效。
- 颜色读回是 `rgb(0,0,255)` 字符串，设置值是 `#0000FF` → 用 canvas `fillStyle` 归一化后比，否则一次误报 4 条。
- 枚举成员名带前缀：`MarkerStyle.MarkerStyleDiamond`、`CategoryType.TimeScale`、`DisplayBlanksAs.gaps`；`DataLabelPosition.above(8)` 对柱形图会被归一到 6。

### 图表探针的宿主页要独立，且插件加载顺序固定
不要塞进主对照台（`index.html` 加载 Univer bundle，老表状态会污染）。独立页只加载三件，**顺序不可换**：`gc.spread.sheets.all.min.js` → `gc.spread.sheets.shapes.min.js` → `gc.spread.sheets.charts.min.js`（charts 内部引用 `Shapes.ShapeBaseData`）。

### SJS `chart.colorScheme()` 入参是 ColorScheme 对象，不是编号（本轮最大假 FAIL 来源）
```js
const C = GC.Spread.Sheets.Charts;
chart.colorScheme(C.ColorSchemes.colorfulPalette2);                     // ✅ 内置：17 套具名方案
chart.colorScheme([C.ColorRule.cycle, ['#FF0000', '#00AA00', '#0000FF']]); // ✅ 自定义色数组
chart.colorScheme(1);   // ❌ 不报错、不生效、读回 undefined —— 会被误判成「可读不可写」
```
读回：`chart.colorScheme()` 返回值与设置值**逐字一致**（含 `[规则, ['#...']]` 两元结构），可直接 `JSON.stringify` 比较；`C.getColor(scheme, i, count)` 按索引取实际颜色；`charts.add(..., 第 9 参)` 也能直接给方案。
**两个产品层注意**：① 换方案**不会刷新已有系列颜色**（官方明确：只对新系列生效），要即时生效须 `getColor()` 回写每个 `series()[i].backColor`；② 只有配色层，没有 Excel 那种带字体/效果的「图表样式」，也没有 `chart.theme`。

### Univer `exportImage()` 需要「图表所在表为活动表」+「首帧已渲染」
两个条件缺一就返回 `undefined`，而且**结果会被缓存**，同一实例再怎么重试都是 `undefined`：
```js
U.getActiveWorkbook().setActiveSheet(chartSheet);  // ← 前面的用例会把活动表切走（实测切到示例表）
await wait(700);                                    // ← 首帧渲染
const dataUrl = await chart.exportImage();          // 成功时 40KB+ 字符串
```
排查顺序：先确认 `getActiveSheet().getSheetName()` 是不是图表所在表（本轮就卡在这里，误当成产品缺陷写进了报告）。稳妥做法：`fresh()` 里每次建图前都 `setActiveSheet`。

### 判「有没有这个能力」看能力，不看方法名
`chart.colorScheme()` 之于「主题 / 调色板」、`ColorSchemes` 之于 `theme`——命名不同但 set+get 双通道齐全就是 PASS。**写「✕ 不成立」又在说明里写「有等价能力」是自相矛盾的**，用户一眼会抓住。要么改判定，要么把说明改成「只有部分等价，缺口在 X」。

### 「无 API」结论前的五问（本轮沉淀）
判 FAIL 前逐条排除：① 数据 / 环境没就绪（活动表、渲染时序、新表灌数据是否真的写进去了）；② 对象调错（2D vs 3D builder、builder vs chart 实例）；③ 通道走错（直调 vs `toBuilder().update()`）；④ **入参类型错**（编号 vs 对象、字符串 vs 数组）；⑤ **方法名不同**（`colorScheme`/`palette`、`ColorSchemes`/`theme`）。五条都排除，才轮到拿 `.d.ts` 说「字段确实不存在」。

## 单工程双引擎对比台（2026-11-06，六模块统一复跑，最多假阴性的一轮）

架构与命令见 `references/workbench-architecture.md`。以下是这一轮踩出的 15 条，**前 10 条每条都造成过成批的假 FAIL**，11～15 条会让证据图/证据表直接失真。

### 1. 插件与 Facade 必须成对——漏一件只让方法变 undefined / 命令未注册

Univer 每个能力 = npm 包 + `registerPlugin` + `/facade` 子路径。三种残缺各有各的症状：

| 缺什么 | 症状 | 本轮代价 |
| :-- | :-- | :-- |
| 只 import facade、没 `registerPlugin` | `redi: Expect 1 dependency item(s) for id "xm"`；`command "sheet.command.addDataValidation" is not registered` | 表模块 Univer 28 → **13** |
| 只 `registerPlugin`、没 import facade | `newConditionalFormattingRule` / `createTextFinderAsync` / `setCrosshairHighlightEnabled` / `setBackgroundImage` 直接 `undefined` | 格式模块 95 → **63**；工作表 94 → **85** |
| advanced preset 不含某能力，也没单装 | 与上一条同症状（方法 undefined） | 查找替换、焦点高亮、背景图三项 |

**固定动作**：注册写进 `for` 循环逐个 `try`（一个 `already registered` 抛出来会中断后面全部），并把结果记进 `window.__proExtraPlugins` 作为自证清单：
```js
Object.keys(window.__proExtraPlugins).filter(k => window.__proExtraPlugins[k] !== 'ok')  // 非空 = 工程缺件，不是产品缺陷
```

### 2. `updatePrintConfig` 静默吸收未知字段——字段白名单只能来自 `.d.ts`

设 `pageOrder`、`blackAndWhite` 这类**类型定义里根本没有的字段**，不报错，读回时**原样回显**。第一版拿"读回有这个键"当 PASS，一口气误判 10 项（P10/P11/P14～P16/P20/P21/P32～P34）。

正确做法：白名单写死在代码里，来源是 `ISheetPrintLayoutConfig`（12 字段）/ `ISheetPrintRenderConfig`（7 字段）：
```js
const MODEL = {
  layout: ['area','subUnitIds','paperSize','direction','scale','customScale',
           'freeze','margin','pageSizeCustom','maxRowsEachPage','maxColumnsEachPage','marginCustom'],
  render: ['gridlines','hAlign','vAlign','headerFooter','headerFooterSetting',
           'isCustomHeaderFooter','watermark'],
}
absent(kind, field) { return MODEL[kind].includes(field) ? null : { ok:false, note:'类型定义里没有该字段，设了只会被静默回显' } }
```
另两个附带陷阱：**读回必须触发一次 `print()`**（走 `BeforeSheetPrintConfirm` 事件载荷取配置，取完 `e.cancel = true`）；**对象键序会变**，比较前先按 key 排序。

### 3. Univer `setFormulas` 必须带前导 `=`

`rg.setFormulas([['SUM(A1:A6)']])` **静默写空**——不报错、不抛异常，单元格就是空的。整批 500 条因此全判 FAIL。写成 `[['=SUM(A1:A6)']]` 才生效。

同类：一致性判定前必须做**值归一**（日期→Excel 序列号、`¥2,328.39`/`50%`→数值、`TRUE`→1），否则 19 条精度类、18 条类型语义类全被误判为不一致。

### 4. 公式 oracle 必须 `data_only=True` 取 Excel 缓存值

`load_workbook(SRC, data_only=True)` 才能拿到 Excel 算好的值；不加则是公式字符串。而且**要取"Excel 结果文本"那一列**，不是"活公式"那一列——样本里往往一列是 `=ABS(-3.5)`、一列是 `3.5`，抽错列等于自己给自己出题。

另一个坑：**复算位置**。把公式写到「对比表」的空列复算，SJS 只能对 289/500；必须写在**数据区所在的表**（「Excel实测」A1:A6/B1:B6 在那儿），才 462/500。写脚本前先确认 `getUsedRange()`，数据区不在首表就不要用首表。

### 5. 样本上下文会污染断言——三处假 FAIL 全是它

| 项 | 断言 | 真因 | 改法 |
| :-- | :-- | :-- | :-- |
| L1 已用区域 | 起点必须是 `1,0` | 样本首表第 1 行有内容 | 改判"覆盖 seed 单元格"，不要求起点坐标 |
| L7 选区 | `setSelection(3,2)` 读回 `3,2` | 样本首表只有 4 列，越界被**吸附回 `3,0`** | 在临时空表 `__pL7` 上测 |
| I6 定义名称 | 用固定名 | 样本自带同名定义名称，抛 `invalid custom name` | 用带时间戳的独有名 `ZZ_TaxRate_<6位>` |

**通则**：探针的断言要写"语义达成"而不是"坐标/字面相等"；凡是要越界、要同名、要空表的用例，先开一张临时表再测。

### 6. 行数不够会造出整批"假失败"

`Range is out of bounds. Max rows: 20`。探针行号推到 140，样本主表只有 20 行。第一版只把 SJS 侧撑到 400 行，Univer 侧照旧抛错。

```js
if (sh.getRowCount() < 400) sh.setRowCount(400)          // SJS
if (ush.getMaxRows && ush.getMaxRows() < 400) ush.setRowCount(400)   // Univer
```
**两侧都要撑**，只改一侧等于把环境的错记成产品的错。

### 7. `window.__wb` 导入后必须重挂

导入走 `createWorkbook(snapshot)`，旧簿被 dispose，但 `window.__wb` 还指着它。症状是探针报 `Range is out of bounds` 或表名不符，很像是"导入失败"。

```js
APP.uni.createWorkbook(snap)
await sleep(400)
window.__wb = APP.uni.getActiveWorkbook()   // 漏了这行，整批跑飞
```
同理 init 后也要挂一次：advanced preset **不自动建簿**，`getActiveWorkbook()` 初始是 null。

### 8. 截图可比性三件套：`setActiveSheetIndex` + 按表名对齐 + 等够时间

- `setActiveSheet(sheet 对象)` **不刷新画布**，要用 `setActiveSheetIndex(0)`。
- 图表样本两家表顺序不同（SJS 首表"组合图"、Univer 首表"新式图表"），必须**按表名**在对方工作簿里找同名表，不能各取第 0 张。
- 等 1200ms 不够，改 1600ms，否则截到空白画布。

```js
APP.sjs.setActiveSheetIndex(0)
const first = APP.sjs.getSheet(0).name()
const target = (first && wb.getSheetByName(first)) || wb.getSheets()[0]
wb.setActiveSheet(target)
```

### 9. NA 不是一句话——三种理由要分开写

`NA` 至少有三个来源，混成一句"Excel 无此能力"文档就解释不通：
- **Excel 真没有**（圆角边框、省略符、内边距、自动合并、样式标志位、折角、图标、单元格状态）
- **别处已覆盖**（导入保真类 → 走数据层比对，不在 API 表重复计分）
- **属效果层**（渲染综合效果 → 走人工复核，不进 API 通过率）

### 10. 判定口径反复一次，两侧同时失真

格式模块第一版把 9 项"Excel 本身没有的能力"用 `none()` 记成 **Univer FAIL**。表面只是 Univer 掉分，实际同时失真：一边拿 Excel 都没有的功能扣对方分，一边给 SJS 的独有能力记 PASS。挪进 NA 后：可判 133 → **124**、SJS FAIL 5 → **2**、Univer PASS 保持 95。

**判据**：NA 的唯一正当理由是"基准方（Excel）自己也没有"。凡是为了"看起来更严格"把基准方没有的项记 FAIL，都是在给自己挖坑。

### 11. SpreadJS 视口复位有三个坑——不复位，往返截图对不上（连改两版才成）

症状：截图停在**上次浏览到的行**（本机实测停在 `244 INTERCEPT`），和另一侧第一屏根本没法比。三次尝试：

| 写法 | 结果 | 原因 |
| :-- | :-- | :-- |
| `sh.setActiveCell(0, 0)` | 仍停在 244 行 | 只移焦点，**不滚视口** |
| `sh.setTopRow(0)` | `TypeError: sh.getTopRow is not a function` | v19 **没有** 这套 API |
| `sh.showCell(0, 0, VerticalPosition.top, …)` | 仍停在 244 行 | **有冻结行时** `showCell(0,0)` 作用在冻结区，滚不动下面的滚动视口 |
| `sh.showRow(sh.frozenRowCount(), VerticalPosition.top)` | `getViewportTopRow(1)` 243 → **4** ✅ | 把**第一个非冻结行**顶到顶部才对 |

另一个容易被当成 bug 的事实：**xlsx 里存了上次的滚动位置，SpreadJS 导入时会还原，Univer 不还原**。所以"两边打开后行位置不同"未必是缺陷，但截图必须先归一化再比。

排查口诀：先 `Object.keys`/`typeof` 确认方法存在（`getTopRow` 这类老文档里的名字在 v19 已删），再查 `frozenRowCount()`，最后用 `getViewportTopRow(1)` 验证滚动视口（`(0)` 是冻结区，永远 0，看不出效果）。

### 12. `page.evaluate(fn, a, b)` 只接受一个参数

写 `await page.evaluate((s, u) => APP.importOne(s, u, u), side, rel)` 会抛 `Too many arguments`，六个模块的往返抓取**全部**失败。包成对象即可：

```js
await page.evaluate(({ s, u }) => APP.importOne(s, u, u), { s: side, u: rel });
```

### 13. SpreadJS 试用版导出会插一张 `Evaluation Version` 表

`format-sjs.xlsx` 13 张表 vs `format-uni.xlsx` 12 张：多的那张是试用版授权提示，**插在第一张**。后果有两个：
- 截图取 `setActiveSheetIndex(0)` 就截到空白授权页，看起来像"导出后全丢"；
- 表数统计会差 1，被误读成"导出多写了一张表"。

`alignViews` 里要按**首次导入时的表名**定位，找不到才取第一张非 `/evaluation/i` 的表；报告里要写明"多一张授权表，正式授权版无此表，比对时已剔除"。

### 14. 单模块重跑会把汇总 JSON 冲掉

`capture_io.js` 原来只写本次模块的 `all`，重跑 `format` 就把 `_io-all.json` 里其他五个模块抹了，证据图页脚统计随即变成 `?`。改成先读旧文件再 `Object.assign(merged, all)`。

**通用做法**：任何"每模块产一份、再汇总一份"的管线，写入前都要先 merge，不要 `writeFileSync(JSON.stringify(本次))`。

### 15. 逐条表截断不能停在半个括号/半个字段名上

探针会把整段 JSON 塞进说明（`sample={"id":"tid-a1","subUnitId":"IZiV18…`），生成表格时按字符数硬切，读者看到的是：

> `Univer 按底色 rgb(255,199,206) 隐藏行=[1,3] … 颜色筛选可用（颜色须 rgb 形`

三处一起改才对：
1. 探针侧：`brief(s, n)` 里先把长 JSON 收成 `{…}`，再回退半个英文词 `[A-Za-z0-9_]{2,}$`，末尾补 `…`；
2. 生成器侧：`clip()` 同样收边界（`{` / `（` / `“` 回退 + 尾巴去标点）；
3. 表格侧：`|` 必须转义 `\|`、换行压空格，否则一行错会带歪整张表。

改完要**从源头重跑**（`run_cases.js <模块>`）再重生成附录，只在生成器里打补丁治不了 cases.json 里的半截原文。

---

## 第二部分·补（2026-11-06 加 Designer / 打印界面 / PDF 输出时踩的）

### 16. Designer 的三个入口陷阱

**① 资源包必须排在上之前。** `gc.spread.sheets.designer.resource.cn.min.js` 在前、`gc.spread.sheets.designer.all.min.js` 在后；反了会在 Designer 初始化时抛 `Cannot read properties of undefined (reading 'DR')`，`GC.Spread.Sheets.Designer.Designer` 根本注册不上，页面里只剩一个空 host——**报错信息完全不提加载顺序，极难联想**。对照实验才定案：写两个只差顺序的测试页，`res_first → {ok:true}` / `designer_first → 抛 DR`。

**② 探针取 Workbook 要走 `designer.getWorkbook()`。** `new GC.Spread.Sheets.Designer.Designer(host)` 之后 `APP.sjs = designer.getWorkbook()`，其余探针不用改。Designer 实例上**没有** `execute()`，`Designer.Commands` 是空对象，`Object.keys(designer)` 只有一个 `__`——别在这些地方找命令入口。

**③ 打印预览在「文件」后台面板里，选择器 `.gc-designer-fileMenuButton`。** 点它 → 在 `.gc-list-control-item` 列表里找文本为 `打印` 的项点进去 → 右侧是实时预览，底部有「第 X 页 / 共 Y 页」翻页条。`_runRibbonCommand(CommandNames.FileMenuButton)` 调了不报错但**什么也不会发生**，别被"调用成功"骗了。

### 17. Univer 打印读回要 `addEvent`，不能把 `Event.Xxx` 当函数调

`univerAPI.Event.BeforeSheetPrintConfirm` 是**事件名常量（string）**，当函数调会报 `is not a function`。正确写法：

```js
const sub = univerAPI.addEvent(univerAPI.Event.BeforeSheetPrintConfirm, (e) => {
  got = { layout: e.layoutConfig, render: e.renderConfig };
  e.cancel = true;          // 拦掉真打印，否则弹浏览器打印对话框
});
await univerAPI.getActiveWorkbook().print();   // 必须真触发一次才有载荷
sub && sub.dispose && sub.dispose();
```

可挂的 6 个：`BeforeSheetPrintOpen / BeforeSheetPrintConfirm / BeforeSheetPrintCanceled / SheetPrintOpen / SheetPrintConfirmed / SheetPrintCanceled`。

另外：`wb.openPrintDialog()` 在 `page.evaluate` 里会抛 `[CommandService]: Command handler should not return a promise` —— **对话框其实已经正常打开了**，这是 evaluate 把返回值当 promise 处理导致的假警报，不要据此判"打不开"。

### 18. `spread.savePDF` 的成功回调给的是 Blob，不是 base64

文档写 `function (blob) {}`，但若按老经验把返回值直接塞进 `page.evaluate`，**Blob 过桥会变成 `{}`**：JSON 里看不到 `bytes`、写出来的文件 0 字节，**且不报任何错**。必须在页面内 `FileReader.readAsDataURL(blob)` 转出来再返回。

### 19. SJS 导出中文 PDF 必须先 `registerFont`，否则真乱码

`gc.spread.sheets.pdf` 只内置 14 种**西文**标准字体。不注册中文时，内嵌字体表里只有 `Times-Roman` / `Times-Bold`，中文被映射到西文字形 → 乱码（Chromium 与 pymupdf **两个渲染器一致乱码**，所以不是渲染器问题）。

```js
const fb = await fetch('vendor/fonts/simhei.ttf').then((r) => r.arrayBuffer());
const F = GC.Spread.Sheets.PDF.PDFFontsManager;
['SimHei', '宋体', '微软雅黑', '等线'].forEach((n) => F.registerFont(n, { normal: fb, bold: fb }));
```

注册名要和**单元格里写的字体名对得上**才命中，所以同一份 ttf 挂多个中文名。注册后内嵌字体变成 `CZZZZZ+SimHei`，中文正常、文本层可提取。

### 20. pymupdf 会把合法的中文 PDF 渲染成乱码——PDF 结论必须换第二个渲染器复核

`sjs-print.pdf` 注册字体之后：文本层完全正确、**Chromium 打开完全正常**，但 pymupdf 渲染仍是乱码（TrueType 子集取错字形）。本轮差点据此写下"SJS 导出 PDF 中文乱码"的错误结论——与"无头截图看不见 ≠ 没渲染"是同一类错误。

定版做法：**像素比对用 Chromium**（`render_pdf_pages.js`：`file:///...#page=N&zoom=100&toolbar=0` + 截白纸区域）；**文本层与页尺寸用 pymupdf**（`get_text()` / `page.rect`，这两项可信）。

### 21. 比 PDF 文本层要按「字符集」比，不能按顺序比

形状文字（Rectangle / Arrow / Text box）在 PDF 里的绘制次序两家不同，按字符串顺序比会得到一堆"内容缺失"的假差异（实测一度看起来少了整页图表标题，其实只是顺序不同）。

正确姿势：`Counter(re.sub(r'\s+','',text))` 后做 `ca - cb` / `cb - ca`，看**多出/少了哪些字**。这样才能读出真差异，例如"Excel 第 3 页多放 3 行、SJS 放到第 4 页"。

配套经验：**图表文字可能完全不进文本层**（SJS 把图表标签画成矢量），所以"文本层没有图表标签"≠"图没画出来"，必须看图确认。

### 22. Excel COM 读 `PageSetup` 的三个硬要求

1. **脚本文件必须存成 UTF-8 with BOM**。Windows PowerShell 5.1 按 ANSI 读无 BOM 文件，含中文路径直接变 `鎵撳嵃-PDF-澶氬`，报"找不到文件"。
2. **只在已激活的工作表上可靠读**：`foreach ($ws in $wb.Worksheets) { $ws.Activate(); $ws.PageSetup... }`。
3. 单位是 **point**（1 inch = 72 pt），SpreadJS 是 **1/100 inch**；比较前先统一，否则会得出"页边距不一致"的假结论。

## 2026-11-06 晚补：条件格式 UI 包、公式读数口径、曲面图（几条都会直接写出错结论）

### 23. Univer 条件格式只装 model 包 → API 全通、画布全空（与缺 `-filter-ui` 同一类坑）

`@univerjs/sheets-conditional-formatting` 只有**数据与命令**，数据条/色阶/图标集的**绘制在 `@univerjs/sheets-conditional-formatting-ui`**（含 canvas 渲染器与其 CSS）。

漏装 UI 包时的表现极具误导性：

- `addConditionalFormattingRule()` 不报错；
- `getConditionalFormattingRules()` 读回**完全正确**（颜色 `#638EC6`、三色阶 `#F8696B,#FFEB84,#63BE7B`、图标集 `3TrafficLights1` 全对）；
- 画布与打印预览**一个都不画**。

定因两步（缺一不可）：① 空白簿 + API 新建三类规则 → 排除"导入丢数据"；② 官网 playground 跑**同一段代码**能画 → 分离"我们的构建"与"产品能力"。两步都做完才敢下"是构建少了包"的结论。

### 24. 官网 playground 的真实入口是深链 `/playground/sheets/<slug>`

- `office.univer.ai/zh-CN/playground/sheets/<slug>`：直接深链，页面上有 `window.univerAPI`，可以跑 API 对照。
- `office.univer.ai/zh-CN/showcase/sheets/<slug>`：外壳页，编辑器在 iframe 里，`window.univerAPI` 拿不到。
- `univer.ai/showcase/...`：404，显示 `#REF!`。
- `/exchange`：纯格式转换 demo，**没有编辑器**，图表在服务端转换阶段就丢——不要拿它当"Univer 能不能显示"的证据。

### 25. Univer 会给公式结果自动套数字格式——判数值前必须先把复算列设成「常规」

`getValues()` 在某些单元格上返回的是**套了格式的显示值**：负金额 → `(¥188.71)`、比例 → `8%`、金额 → `¥5,299.07`。直接跟 Excel oracle 比，会把这 11 个财务函数全判成"算不准"，而实际原始值是 `-188.71233644011` vs Excel `-188.712336440109`，只差 1e-12。

定版做法：写公式**之前**两侧都把复算列设成常规格式（`sh.setFormatter(r, c, 'General')` / `range.setNumberFormat('General')`），再读再比。探针 `probe_formula_readback.js` 同时读 `getValues()` / `getDisplayValue()` / 快照 `cellData.v`，一眼看出差异来自哪一层。

连带教训：**"读回是字符串"要先怀疑格式，再怀疑计算**。带 `¥`、`%`、括号负数、千分位的读回值，一律先归一再判。

### 26. 日期序列号换算不能用 `getTime() - Date.UTC(1899,11,30)`

`getTime()` 是 UTC 毫秒。东八区的 `new Date(2020,2,15)`（本地零点）实际是 `2020-03-14T16:00Z`，减出来是 `43904.667` 而不是 `43905` → `DATE`/`EDATE`/`EOMONTH` 被误判成"不一致"。

正确写法：拿**本地墙钟字段**重建 UTC 时间。

```js
const wall = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(),
  d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds());
const serial = Math.round((wall - Date.UTC(1899, 11, 30)) / 86400000 * 1e9) / 1e9;
```

精度也要留够：收到 `1e-6` 会把 `TIME(10,20,30)` 的 `0.4309027778` 截成 `0.430903`，同样假阴性。**1e-9 起**。

### 27. oracle 本身就是错误值的用例要单独判（`NA()` 这类）

`NA()` 在 Excel 里返回 `#N/A`，引擎返回 `#N/A` 就是**对的**。如果判定逻辑写成"返回错误值 = 不支持"，两家都会在 `NA` 上记一条"不支持"，客户一眼就能看出报告写错了。

做法：造 oracle 时把这类值记成哨兵（本项目用 `'ERR'`），判定时 `exp === 'ERR'` 且读回是错误值 → PASS。

### 28. SpreadJS 曲面图表：对象在、类型对，画出来是柱形

`surface` / `surfaceWireframe` / `surfaceTopView` / `surfaceTopViewWireframe` 在 `ChartType` 枚举里存在（44～47），导入后 `chart.chartType()` 也读回 44～47，**但画布画的是簇状柱形**。官方文档《保留不支持的图表》把四型列在"当前不支持"清单，并写明 `preserveUnsupportedChart` 默认 `false` 时"导入会将不受支持的图表转换为柱形图"。

所以"导入对象数 = 源文件图表数"**不能**作为"图表还原"的结论。要看三件事：对象数、`chartType()` 读回、**画布实际画什么**。同一个清单里的 `bubble3DEffect`、`boxWhisker`、`funnel`、`paretoLine`、`regionMap`、`waterfall`、`histogram` 都要按这个办法逐项验。

> 同一份样本里 Univer 的错法是"4 张曲面图对象根本没建"。两家都拿不到正确曲面图，这一项不能当卖点。

### 29. 局部取证：缩放优先于滚动

要截"某张表下方的一排图表"，用滚轮往下滚两侧的取景**对不上**（两个编辑器的滚动步长与行高换算不同，实测滚 4 步后一边还在数据区、一边已经在图表区）。

改成**两侧同缩放**：SJS `sheet.zoom(0.25)`（下限 0.25）、Univer `sheet.zoom(0.25)`，一屏看全后再截。`shot_sheet.js` 的第 4 个参数就是缩放比。

### 30. Windows 脚本文件的双 BOM 坑

`.ps1` 必须 UTF-8 **with BOM**（否则中文路径按 ANSI 读成乱码），但**不能有双 BOM**：文件头出现 `EF BB BF EF BB BF` 时，PowerShell 5.1 报 `ParserError: 赋值表达式无效`，报错行还指向第一行无关内容。

修法：`raw.decode('utf-8-sig').lstrip('\ufeff')` 后再以 `utf-8-sig` 重写，写之前断言开头只有一个 BOM。

### 31. 逐项判定里"对象结构"也要防错：Univer 的 `rule` 是对象不是数组

`getConditionalFormattingRules()` 返回的 `rule` 是 `{type, config}`，不是 `[{...}]`。按数组 `map` 会抛 `(x.rule || []).map is not a function`。

另外：**导入的规则带 `ranges` + `order` + `stopIfTrue`，API 新建的规则只有 `{type,config,cfId}`**。缺 `ranges` 时渲染器会抛 `reading 'map'` / `reading 'type'` —— 这反而证明渲染通道是通的，是"数据不合法"而不是"能力不存在"，别急着判 FAIL。

### 32. 报告写法硬口径（2026-11-06 三条 + 2026-11-14 补两条）

**结构固定四段式，顺序不许倒装**：①结论（先行）→ ②导入导出对比（截图 + 差异清单，只留本模块相关）→ ③三方/功能截图（Excel / SJS / Univer 拼一张）→ ④API 完整度验证（分组通过率 + FAIL 逐条 + 附录 A）。**不设「用例设计」章**，分组与 NA 边界压进标题下的「测了什么」小表（测试项 / 对象 / 量什么 / 结果在哪节）。

1. **导入导出不一致的地方必须有对比截图**，不能只有文字表格。截图要两侧同表、同缩放、同取景，并标注"哪张是哪张"。
2. **不写"与上一版报告的区别"**。读者只看得到当前这一版，版本对比是内部过程，进报告只会稀释结论。
3. **公式这类没有可见样式变化的模块不加图**，但必须写清①Excel 样本文件在工程里的路径、②三方（Excel 基准 / SJS / Univer）在同一张表上的差异。"都对/都错/仅谁错"的四列交叉表比两张一模一样的截图有用得多。
4. **只陈述本次结果，不记历史过程**（2026-11-14）：不得出现"上一版错了 / 本轮纠错 / 已作废 / 踩过的坑 / 改判轨迹"，改写为「口径 / 前置条件 / 复现要求」。本身份文档（pitfalls）作为内部资产可以记坑，但**交付报告里必须抹掉叙事痕迹**。
5. **不写测试方法**（2026-11-14）：采集步骤、判定条件、复现命令、边界表一律删；能进代码注释就进代码注释。删章后**下一章编号不动**，避免交叉引用连带返工（范例：`docs/10-设计器与Ribbon.md` 717 → 573 行）。

配套：**判定层的错会污染整份报告**。本轮的 11 个财务函数"差异"、4 个日期函数"差异"、`NA()`"不支持"，全是判定写法问题，不是产品差异。数字落进报告前，先拿 3～5 个典型 case 逐个人工核对一遍（读回原值、Excel 原值、判定结论三者对齐）。

### 33. 样本自带的「说明文本框」能把 Univer 的图表标题整条吃掉（2026-11-07，代价最大的一次误判）

给每个用例加一条灰色说明文本框（Excel `Shapes.AddTextbox`，writing 到 drawing 里是 `xdr:sp`）是很好用的自解释办法，但 **Univer 的 xlsx 导入器只要同一个 drawing 里存在 `sp` 锚点，就会把图表的富文本标题整条丢掉**：`getInfo().config.title` 变成空、画布上也不画。

它把我坑到什么程度：我先看到"91 张图只有 8 张有标题"，又发现这 8 张正好是标题只含 1 个文本片段的，于是写下"多段标题 82/82 全丢，零例外"，还画了 A5（多段，丢）vs B1（单段，正常）的对照图当证据。**这条结论完全错**。

钉死它的办法（只改 `xl/drawings/drawing1.xml`，其余字节不动）：

| 变体 | 图表 | 文本框 | 有标题 |
| --- | ---: | ---: | ---: |
| 原样本 | 91 | 91 | 8 |
| 删掉全部文本框 | 91 | 0 | 91 |
| 只留 1/2/4/8/…/90 张图 | ≤90 | 0 | 全部正常 |
| 文本框挪到图表前面 | 91 | 91 | 8（顺序无关） |

结论：与"标题有几个文本片段"无关，与"同表有没有 `sp`"有关。

**通用教训**：①**样本里的任何装饰元素都可能成为被测对象的干扰项**，加注释框、加标题行、加图例说明前先问一句"这会不会被算进被测对象所在的那个 XML 容器"；②结论一旦呈现出"全丢/零例外"这种过于干净的规律，先怀疑自己的样本构造，别急着写进报告；③做**单变量对照实验**（只改一个变量、其余字节不动）比反复看截图快得多。

**工程后果**（图表模块最终采用的做法）：
- 测量用干净样本 `samples/…-无标注.xlsx`（91 图 0 文本框），带注释的原样本只作交付物；
- Univer 单图截图改为「从样本派生只含该图的变体工作簿」`out/variants/{caseId}.xlsx`，顺带把文本框一起丢掉；
- 派生脚本要把工作表 `dimension` 撑开（`A1:AE1000`），否则 Univer 按默认 150 行建表，锚在第 193/543/894 行的图会被夹到最后一行、高度退化成 1px。

### 34. Univer 侧"藏邻居图"三条路全不通（图表单图取景）

想截"只要这一张图"的画面，Univer 没有一条省事的路：

1. `chart.remove()` 掉其余图表 → **幸存的那张不再绘制**。`getCharts()` 数量对、`getInfo().size/position` 都对，画布上就是空的。所以"先删后截"不行。
2. `chart.setAbsolutePosition({x: 原值-20000, y: 原值-20000})` → **负坐标被夹到 0**，90 张图全叠到左上角，剪裁框里一片花。要挪只能往**正方向**挪。
3. 想按 `getInfo().size` 精确剪裁 → `size` 和实际渲染尺寸对不上（多图时读到的值偏大/偏小），剪出来不是多一条网格线就是切掉图例。

最终走的路：**派生单图变体工作簿**（见坑 33），一张表只有一张图，剪裁框按 `size + 6px` 即可，不用做任何"藏/挪"动作。

顺带记一条：`getInfo().size` 在**刚导入、布局还没稳定**时读会偏大（A7 第一次 792，稳定后 721.6），滚动后至少等 1.5s 再读，或者干脆读两次取第二次。

### 35. SpreadJS 侧不要截 Designer 视口，直接用 `chart.toImageSrcAsync()`

Designer 里截单图有三个死结：`workbook.setActiveSheet(i)` 不生效（非活动表 `getCellRect` 一律 `{}`，要键盘 Ctrl+PageDown 切）、滚动是异步懒渲染（滚十几屏后 420ms 不够，截出来是空白网格）、同表邻居图会叠进剪裁框。

`chart.toImageSrcAsync()` 返回图表控件自己渲染的 PNG（data URL），**和 Excel 的 `Chart.Export` 同源同口径**，一次调用解决全部三个问题：没有网格线、没有邻居、没有滚动依赖。代价是它只给你图本身，拿不到"图在表上的位置"——而位置本来就是 API 读回的事，不该靠截图。

> 这条和"API 层由 AI 判、效果层由人工判"的分工完全对齐：截图只负责"画出来像不像"，几何问题交给读回值。

## 2026-11-07 补：采集脚本没采 ≠ 对方没有 API（本模块最贵的一次自伤）

### 36. 「无读回口」这个结论，必须先用类型定义把方法全列出来才敢写

**事故**：打印报告里「打印区域」一项写成 `SJS 无读回口（按工作表整体打印）`，判定「无法比对」。用户凭印象反驳「我印象中是可以的」——一查，不但有，而且**读回来的值与 Excel 逐字一致**。

真因不是 API 藏得深，是**自己的探针脚本没采这几个字段**：`tools/probe_print_elements.js` 的 sjs 段列了 orientation/paperSize/margin/header/footer/centering/showGridLine/repeatRow*/blackAndWhite/pageOrder/firstPageNumber，**唯独漏了 `rowStart / rowEnd / columnStart / columnEnd`**。于是"我的字段表里没有"被写成了"它没有"。更该挨打的是：同一份报告 §5 的 P1～P4「打印区域-四角」SJS 全判 PASS——**自相矛盾就摆在同一份文件里，成稿前没对账**。

**正确顺序（判「无 API」前必走）**：

1. **先列全方法名**，不要凭探针脚本的字段表反推能力边界：
   ```bash
   # SJS：从 d.ts 里把类的全部方法抓出来
   python - << 'PY'
   import io,re
   s=io.open('.../spread-sheets/dist/gc.spread.sheets.d.ts',encoding='utf-8',errors='ignore').read()
   i=s.find('export class PrintInfo'); seg=s[i:i+60000]
   j=seg.find('export class ',1); seg = seg[:j] if j>0 else seg
   print(sorted(set(re.findall(r'^\s{8,}(\w+)\s*\(',seg,re.M))))
   PY
   ```
   `PrintInfo` 实测 32 个方法，一眼就能看出漏了谁。
2. **字段表与方法表做差集**，差集里的每一项都要单独实测一次，不能直接记「无」。
3. **同一模块内做一致性对账**：元素级表格（§4.1）判「无/无法比对」的项，必须去 API 用例表（§5）确认对应用例也是 FAIL/NA。**两张表互相打架 = 必有一张错。**
4. 措辞也要分清三档，不能混着写：
   - **无此方法**（读写都没有）→ `PrintInfo` 没有 `draft` / `printErrors` / `printComments`
   - **有 setter 没 getter**（写进去读不回）
   - **读得到但导入不回填**（本例 Univer：字段有、`area` 回落 `CurrentSheet`）

### 37. SpreadJS 打印区域的三条通道（含一条文档没直说的）

```js
const pi = sheet.printInfo();
pi.rowStart(); pi.rowEnd(); pi.columnStart(); pi.columnEnd();   // 0-based，-1 = 未设置

// NameInfo 的方法是 getName() / getExpression()，没有 getFormula()！
const ni = sheet.getCustomName('Print_Area');                    // 返回 NameInfo 实例
const a1 = GC.Spread.Sheets.CalcEngine.expressionToFormula(sheet, ni.getExpression());
// → 'Data!$A$1:$I$27'
```

双向联动实测：`pi.rowStart(2)/rowEnd(20)/columnStart(1)/columnEnd(5)` 之后，名称公式立刻变 `$B$3:$F$21`；还原后回到 `Data!$A$1:$I$27`。所以打印区域是**存在工作表自定义名称 `Print_Area` 里**的具名固定区域，插删行列会跟着变，还能在公式里引用（`=ROWS(Print_Area)`）。

三个坑：
- `getCustomName` 返回的 NameInfo **没有 `formula` 属性也没有 `getFormula()`**，只有 `getExpression()`；直接 `ni.formula` 得到 `undefined`，会误判成"读不到"。
- 表达式对象是**循环引用**的，直接 `JSON.stringify` 会抛 `Converting circular structure to JSON`；必须过 `expressionToFormula` 或 `String()`。
- `pi.rowStart()` 等四个方法**在没有打印区域时返回 -1**，而打印区域 = 整表时返回 `0/rowCount-1/0/columnCount-1` —— 两者看起来都"像整表"，**判「没读进来」之前先跟 Excel 的 `PageSetup.PrintArea` 对一遍**。

### 38. 样本设计缺陷会被误读成对方缺陷（打印区域实例）

本样本三张表打印区域都**等于或略大于**数据区（Data 27×9 ↔ `A1:I27`、Mixed 63×8 ↔ `A1:H63`、Visuals 42×17 ↔ `A1:R44` 往外多 2 行 1 列），所以打印输出上**没有可观察的裁切效果**。当时的处置是"截不出图 → 记无法比对"，正确处置应该是"**样本没能体现该特性 → 样本不合格，去补样本**"。

这就是「样本先行」五条硬要求里第 ③ 条（元素刻意铺满）的反面案例：要证明"打印区域生效"，样本必须让打印区域**明显小于**数据区，否则打了等于没打。


## 2026-11-07 补二：字段名对不上，不等于没有能力

### 39. 「配置里没有 repeatRow」→ 判「不能重复标题行」（打印模块最容易被反问的一条）

Univer 打印配置只有 12 个 layout 字段，确实没有 `repeatRowStart/End`、`repeatColumnStart/End`。据此判 P5～P8 FAIL，并在报告里写"只有整行/整列冻结开关，无法指定第几行"。用户一句反问就推翻了：**冻结前 3 行，第 2 页就重复这 3 行**。

真相是三层结构：

| 层 | 事实 |
| --- | --- |
| 打印配置 `freeze` | 恒为 `["Row","Column"]`，**只是"允许重复"的总开关**，与工作表冻不冻结无关 |
| 工作表冻结 | `setFrozenRows(n)` / `setFrozenColumns(m)` 定义**重复哪几行/列** |
| 效果 | 冻结 3 行 → 第 2 页重复第 1～3 行；冻结 1 列 + 横向分页 → 每页左端重复第 1 列（都有 canvas 截图） |

**这就是"字段名对不上"被误写成"没有能力"的典型。** 判 FAIL 之前先问自己：这个能力在产品里会不会有**另一条入口**？

三条硬规矩：
1. **能力按效果判，字段按键判**，两张表分开记：`配置字段表`（有没有 `repeatRow`）与 `能力表`（能不能让第 2 页重复表头）不是一回事。
2. **读回一致不算证明，输出变化才算**。`freeze` 设进去能读回，但读回值恒为 `["Row","Column"]`——它根本不携带"第几行"的信息，光看读回永远发现不了机制。
3. 边界必须一起写，否则等于把缺陷藏起来：冻结区**只能从第 1 行/列开始**（Excel 的 `$3:$5` 做不到）；**打印重复与编辑器冻结窗格是同一个状态**（Excel 的 `Print_Titles` 是独立设置）——这是真正的语义差异，"有/没有"是错的问法。

### 40. 证据图三要素：哪张表、第几页、什么状态（缺一个就会被追问）

用户看打印报告时的原话：「截图中没有说明 Univer 打印时，第三个 sheet 的重复行列头没有生效」。图里其实写了 `Mixed`，但没写"第三个 sheet"，也没解释**为什么图上看不到列头**（Mixed 只有 8 列、一页放得下，列本来就不跨页）。

规矩：每张证据图的标题或 `<sub>` 必须含
- **对象**：哪个 sheet（并标出它在本样本里的序号，`第三个 sheet「Mixed」`比`Mixed`清楚）
- **位置**：第几页（跨页类特性必须写"第 2 页顶部"）
- **状态**：导入后 / 手工设置后 / 导出重开后

外加一条：**本样本体现不了的维度，要显式声明并指路**——"列头在这张样本上本来就不会跨页，列方向见 §3.3.7 的强制分页实验"。不声明，读者会以为漏测。

## 2026-11-07 补三：拼图裁到哪里 ≠ 内容到哪里

### 41. 「截图里没看见」和「截图里裁不到」是两件事（迷你图事故）

打印报告原先写「条件格式、迷你图、合计行三家都打得出来」，还配了图 3-8「三家的打印输出里都在」。用户反问：「我观测的是打印的时候迷你图消失，为什么你说存在？」

复查三分钟就翻案：Univer 的打印输出里，趋势列被列宽挤到了**第 2 页**，而我那张拼图只抓了第 1 页、按墨迹裁到「日期」列就收边。**「我没裁到」被写成了「它在那儿」。**

错法比错误更值得记：

1. **拼图的裁剪带不是证据边界**。用 `crop_ink` / 固定比例裁图时，裁掉的区域等于没测。凡是要判「有 / 无」的元素，必须**先看整页**，确认该元素落在哪一页、哪一列，再决定裁哪里。
2. **跨页元素要按页遍历**。分页是打印的常态：一张 9 列的表在窄纸型下就是会分成 2 页。只抓第 1 页等于只测了前 8 列。取证脚本必须 `for page in pages`，且核对「总页数 × 每页内容」能拼回完整表。
3. **一张图里塞多个结论，必然有一个是猜的**。图 3-8 同时号称证明「条件格式」和「迷你图」两件事，结果后者纯属跟着前者顺推。**一个元素一张图、一句结论**，宁可多一张图。

复核的正确顺序（判「某元素打印时在不在」）：
- ① 抓**全部页**，不预设它在第 1 页；
- ② 找**表头**定位：元素属于哪一列/哪一行，先确认表头在不在（本例表头「趋势」在，格内空白 → 是渲染缺失，不是列丢失）；
- ③ 排除布局因素：**改列宽 / 压窄邻列**把该列拽进第 1 页再看一次（本例放宽到 200px、整表一页，仍空白 → 不是"太窄画不下"）；
- ④ 反查数据层：编辑器视图有没有、导出文件 XML 里在不在（本例 `x14:sparklineGroup` 42 处都在）；
- ⑤ 只有 ①～④ 都指向同一结论，才写「不支持」。

### 42. 同一页里的元素要分开判，别用「元素」这个词打包

条件格式（数据条/色阶/图标集）在 Univer 打印里**画得出来**，迷你图**画不出来**——同一页、同一批渲染管线，两种结果。所以：

- 「Univer 打印丢元素」会被条件格式截图反驳；
- 「Univer 打印元素没问题」会被迷你图截图反驳。

**元素级结论必须逐元素落表**（条件格式一项、迷你图一项、图表一项、合计行一项），不许出现「X、Y、Z 三家都在」这种合并句。合并句一旦有一样错，整句都不可信。

修完的写法（可直接抄）：`元素要分两种说法：条件格式、图表、合计行三家都打得出来；但迷你图 Univer 打印时整列空白——数据在、编辑器能画、放宽列宽也画不出，缺陷在打印渲染。`

### 43. 「批注」是两种格式，判「导入丢批注」前先分清 Note 与 Threaded Comment

Excel 的「批注」在两套 OOXML 里、两个界面名称下：

| 形态 | OOXML | Excel 界面名 | 通道 |
| --- | --- | --- | --- |
| 传统批注 | `xl/comments1.xml` + VML 红三角 | 备注（Note） | SpreadJS `sheet.comments` |
| 新式讨论 | `xl/threadedComments/*.xml` + `xl/persons/person.xml` | 批注 / 评论（可回复、@、已解决） | SpreadJS `sheet.threadedComments`；Univer `sheet.getComments()` |

实测（Excel 365 COM 造一份两种并存的样本，A1 `AddCommentThreaded`、A2 `AddComment`）：**SpreadJS 两条都还原；Univer 只还原 A1，A2 丢失**。所以对外口径必须是「Univer 缺传统批注（Note）通道」，不能写「Univer 不支持批注」——后者会被 `addCommentAsync` 一行代码当场打回。

判定顺序（四条缺一不可）：

1. **先造样本，两种格式都造**。`AddCommentThreaded` 只有 Excel 365 有，Excel 2021 直接抛错，别用 2021 造。
2. **先测 API 能力，再测导入还原**。`range.addCommentAsync()` → `range.getComment()` → `sheet.getComments()` 三步都对，说明能力在；此时导入仍为 0，才轮到"导入层"下结论。
3. **插件先补齐**。Univer 需要 `@univerjs/thread-comment` + `@univerjs/thread-comment-ui` + `@univerjs/sheets-thread-comment` + `@univerjs/sheets-thread-comment-ui`，加两个 `/facade` side-effect import。漏装只会让 `getComments` 是 `undefined`，不报错——与第 10 条同源。
4. **落盘层核对 XML**。传统批注数 `xl/comments*.xml` 的 `<comment>`，新式讨论数 `xl/threadedComments/*.xml` 的 `<threadedComment>`，两个计数分开报。Univer 导出时会把讨论批注**同时**写成 `threadedComments` 和兼容用的 `comments1.xml`（正文是 `[Threaded comment]` 占位串），后者不是真 Note，不能计入传统批注数。

附带事实：Univer 导出的讨论批注，SpreadJS 能通过 `threadedComments.all()` 正常读回；SJS 的批注模块在 `gc.spread.sheets.all.min.js` 全量包里，不需要另加包。

## 2026-11-07 补四：打印设置落不落盘、层叠顺序怎么判、文档里别写波浪号

### 44. 「改完设置生效」要分两件事：运行时生效 / 导出落盘

打印模块的 PASS 只证明"当前会话里 set 了、get 得到"。客户真正关心的是**改完存回文件还带不带走**，所以必须并测一步：改设置 → 导出 xlsx → 解包读 `pageSetup` / `headerFooter` / `_xlnm.Print_Titles`。

| 动作 | SJS 导出 XML | Univer 导出 XML |
| --- | --- | --- |
| 导入原样本，原样导出 | `orientation="landscape"` / paperSize 9 | 同左 |
| 改成纵向再导出 | **`orientation="portrait"`** | **仍是 `landscape`** |

结论：Univer 的页面设置是**纯运行时状态**，既不读文件也不写文件。三处连带影响：

1. 话术禁止写"会把你的设置弄丢"——三家导出的文件都是完好的（这一点要靠解包证明）。只能写"**导进来要重设一遍，改完存回去也不会带走**"。
2. 报告里"导入 2/8"与"字段层 19/35"必须分开陈述，否则读者会以为"能设就算通"。
3. 验收清单要加一条：**"设置是否持久化"单列**，与"能否设置"并列。

### 45. 浮动对象：SJS 是单一有序集合，Univer 按类型拆三个集合——所以「层叠顺序不同」通常是无依据的

模型差异（实测 `out/data/shape-zorder.json`）：

| 侧 | 模型 | 能不能表达跨类型前后关系 |
| --- | --- | --- |
| SpreadJS | `shapes.all()` 单一有序集合，7 个对象顺序与 Excel `drawing1.xml` 逐一对上（Chart1→Chart2→Chart3→Image4→矩形5→圆柱6→文本框7） | 能，顺序即层序 |
| Univer | `getShapes()` 3 / `getCharts()` 3 / `getImages()` 1 **三个类型集合** | 不能，跨类型前后关系不由文件对象顺序表达 |

三条规矩：

1. **判「层叠顺序」必须先确认两侧都画出来了**。本次样本上 Univer 的 3 个图表对象存在（`getCharts()` 返回 3 个）但画布没有像素，遮挡关系根本无从观察——正确写法是"本样本无从比对"，不是"顺序不同"。
2. **看不到遮挡就别谈层序**。只在 128×63px 范围内轻微重叠、肉眼无遮挡时，任何"谁压住谁"的描述都是硬凑。
3. **几何读回先探再写**。Univer 的 shape / chart 只有写入口（`setPosition` / `setSize` / `setAbsolutePosition`），没有对应的几何 getter；判"位置没还原"之前先确认能不能读回，读不回就只能用截图判，且要写明"仅画面观测"。

### 46. Markdown 里成对的单波浪号也会变成删除线——范围连接符一律用 `～` 或 `-`

交付文档里写用例编号区间 `1～4、14～16`，渲染出来是"1、4、14、16"中间一段带横线：GFM 把**成对出现的单个 `～`** 也当删除线定界符。

- 正确写法：全角 `1～4、14～16`，或半角连字符 `1-4、14-16`。
- 只有一处 `～` 不会触发，所以抽样看一两行发现不了，必须全库扫描：`grep -rn '～' --include='*.md' .`。
- 故意要删除线（标注"该说法已作废"）才用双波浪号包裹，并且正文别处不要再出现单波浪号，否则配对关系会错位。

## 2026-11-07 补五：公式模块定稿（判定尺子 / 插件配对 / 性能口径）

### 47. 判定尺子不标定就出数，会跑出反直觉排名

- 事故：用户要求"直接判和 Excel 一样不一样"，第一版按 **double 严格逐位相同**（`===`）重算：SpreadJS 281/484 = 58.06%、Univer 289/482 = 59.96%——**被普遍认为精度差的那家反而更高**。
- 根因：Excel 内部 17 位，两家都存不下 17 位，`=PI()` 在两边都被判"不一样"。这个尺子比的是"谁输出位数接近 17 位"，不是"谁算得准"，会把正确答案判成错误。
- 标定三步（写在 `references/formula-module.md` §2.1）：①必然正确的用例（`=PI()` / `=1+1`）两家都要过；②人为差 1e-3 的对照两家都要不过；③排名反直觉就先怀疑尺子。
- 定稿口径：**数值按 Excel 自身的 15 位有效数字归一后逐字比，文本（含复数）原样逐字比**。措辞必须写"**位数归一，不是容差**"——15 位以内差一位就判不一致。

### 48. 一加相对容差，真差异就全消失了

- 设 1e-9 相对容差后：`BESSEL*` 四条（差 1e-8）全判 PASS；复数解析成实部/虚部再按 1e-9 比，26 条也全过。报告结论会变成"两家都对、复数没问题"，客户做逐位对账时一导出就露。
- 正确做法：**先逐字比，再按误差幅度分档**（≤1e-12 / ≤1e-9 / ≤1e-6 / ≤1e-3 / >1e-3 / 文本或形状）。这样既保留"差一位就算不一致"的严格性，又能说清"差 1e-15 和差 1e-6 不是一回事"。
- 本轮实测：两家失分**没有一条超过 1e-3**；SJS 37 条里 27 条在 1e-12 以内，Univer 157 条里 104 条在 1e-12 以内、52 条在 1e-9 以内。

### 49. 支持度与准确度必须两个分母，`#NAME?` 只扣支持度

- 混成一个百分比会同时冤枉两家：一家可能函数少但算得准，另一家函数全但位数差。
- 定稿：支持度 ÷ **504**（519 个函数 − 15 条无 oracle，含易失、`IMAGE` 与 4 个特殊函数）；准确度 ÷（该家支持数 − 5 易失 − `IMAGE`）= SJS 488 / Univer 485。
- 只算支持度的两类：**易失 5 条**（`NOW`/`TODAY`/`RAND`/`RANDBETWEEN`/`RANDARRAY`，值每次重算都变）与 **`IMAGE`**（图片对象，Excel 本机也给不出标量答案，错误码 `0x800A0803`）。
- 数字落地：SJS 494/504 = 98.02% · 448/488 = 91.80%；Univer 491/504 = 97.42% · 324/485 = 66.80%。支持度只差 3 个，准确度差 25.0 pp——这是本模块唯一值得讲的一句话。

### 50. SpreadJS 的 addon 也要求版本锁死：版本差一点，表现和"没装"一样

- `GETPIVOTDATA` 不装 `@grapecity-software/spread-sheets-pivot-addon` 返回 `#REF!`；**插件版本必须与主包完全一致**（本轮 19.2.2 对 19.2.2）。跨小版本注册失败，症状与未加载完全相同，很容易写成"函数不支持"。
- 加载顺序：**core → pivot addon → resources-zh → designer**。Designer 必须最后接管 Ribbon/UI。
- ~~副作用：装插件后透视表表头变成英文 `Row Labels` / `Grand Total`（不走 Designer 中文资源包）~~ **该结论已于 2026-11-09 推翻**：英文的真因是缺核心中文资源包 `@grapecity-software/spread-sheets-resources-zh`，补上并排在 pivot 插件之后即全中文（连面板与单元格文案一起变）。机制与证据见本条后的 **59**。
- 装好后两条公式与 Excel 逐格相同（500 / 1520），透视表结果区 `A3:B7` 数值全对。

### 51. "包里 grep 得到函数名"不等于"这个函数能算"

- Univer 侧 `GETPIVOTDATA` 在 `@univerjs/engine-formula` 里有函数名映射、函数描述、15 个语言包，但**没有任何计算实现**；`@univerjs-pro/engine-pivot`、`sheets-pivot` 里也搜不到。
- 只在函数清单里看到条目就写"支持"会翻车；反过来，只看 `#NAME?` 就写"不支持透视表"也会翻车——透视表能力是好的。
- 三层取证：①**插件在**（`@univerjs-pro/sheets-pivot` 随 advanced preset 注册，`workbook.addPivotTable` 是 function）；②**透视表被解析**（`getPivotTableByCell(2,0)` 返回 `xlsxPivot1`，结果区 500/530/490/1520 全对）；③**函数无实现**（无参/单条件/双条件三条全 `#NAME?`）。取证完只能说"透视表有、这个取值函数没有"。
- 与原则 10 的关系：原则 10 是"方法 undefined 先怀疑装配"，本条是"装配齐了仍算不出，才叫函数没实现"。

### 52. 性能：同页连跑会退化，写入耗时不是引擎指标

- **同页面连续跑**：Univer 第 2 轮重算从 11 s 涨到 22 s（旧簿 `dispose()` 后公式 worker 仍留着上一份数据）。**每轮必须新开页面**，否则倍数凭空翻倍。
- **取中位不取均值**：6 次极差占中位 SJS 45%（首轮冷启动）/ Univer 19%，均值会被一次抖动带偏。
- **写入 ≠ 引擎**：SJS 写入 3.2 s，其中大部分是 Designer 每次 `setFormula` 通知 Ribbon/状态栏的成本（裸 Workbook 1.1 s）；Univer 2.5 s。比引擎只看重算那一栏（2,402 vs 12,271 ms = 5.11×），写入单列为"带真实外壳的工程成本"。
- **条数按实际写公式的列算**：`Data` 5 列数据但只写 4 列公式，按 5 列会多算 10,000 条（53,263 → 43,263）。负载变了必须重跑，不能沿用旧耗时。
- 跑完先做**等价性先验**：抽 9 个格（明细/汇总/链式/动态数组）双端取值逐字相同，才能排除"一边少算了一批"。

### 53. 不设 `CALC_PRECISION` 就比，等于拿两把尺子量

- 事故：只在"比较时"把 Excel 归一到 15 位，SpreadJS 仍按完整 double 算、拼文本默认 17 位——26 条复数用例有 15 条被判"不一致"，报告里写成"SpreadJS 复数只有 42.31% 逐字一致"，看着像算法差。
- 真相：Excel 的复数文本是 15 位有效数字拼的，SpreadJS 拼 17 位，**同一条正确结果必然不同**。设 `GC.Spread.CalcEngine.Functions.CALC_PRECISION = 15` 后从 11/26 提到 24/26，只剩 `IMCOT`/`IMTAN` 实部末位（1e-17 量级）。
- 纪律：①它是**静态属性**，必须在 `new Workbook` / `new Designer` **之前**设，建完再改无效；②对比台启动段与跑批脚本都要默认设，不能只在报告附录里"另跑一档"；③Univer 没有对应开关，它本来就只写 12 位左右，所以只是"把强的一方拉到和弱的一方同一把尺子"，不是给谁放水；④旧数据要复现就显式退回（`?prec=default` / `--prec-default`），别默默改口径。

### 54. 能测的函数被排除在分母外，客户第一个问"凭什么"

- 事故：`GETPIVOTDATA`/`SHEET`/`SHEETS`/`IMAGE` 因为"要另造样本"被单列，总分母只报 500；用户直接质疑"这几个明明可以测，为什么排除"。
- 正确做法：**分母 = 有 Excel oracle 的函数数**。能造样本测出真值的，无论样本多特殊，都要进分母（7 条用例 → 4 个函数，一个函数一票）；进不了分母的只有"本机给不出 oracle"那 15 条。
- 只有一种例外可以只算支持度：**没有标量真值**（`IMAGE`，Excel 自己也只回错误码）——和易失函数同一种处理，写清"只计支持度"即可。
- 落地：SJS 494 = 490 + 4、Univer 491 = 488 + 3；准确度 SJS 448 = 445 + 3、Univer 324 = 322 + 2。特殊函数那页改成"证据页"，不再另立分母。

### 55. Designer 里批量写公式不挂事件，测出来的写入耗时不是引擎的

- 事故：对比台宿主是 Designer 包裹的 Workbook，批量灌 43,263 条公式时只挂了计算（`suspendCalcService`）和绘制（`suspendPaint`），没挂事件。每条 `setFormula` 都触发一次 `CellChanged`，Designer 收到就刷 Ribbon / 状态栏重算按钮——4.3 万条 = 4.3 万次 UI 通知，写公式被测成 3.2 s，报告里还写成"SpreadJS 写入比 Univer 慢（0.76×）"。
- 真相：补上 `spread.suspendEvent()`（整段写完再 `spread.resumeEvent()`），写入从 **3,160 ms 掉到 1,134 ms**，是原来的 1/2.8；重算一格没变（2,540 → 2,417 ms），说明差的这 2 s 全是事件通知，不是计算。照最佳实践写，Designer 下的写入和裸 Workbook 同一量级（1.0–1.4 s），比 Univer 的 2.6 s 快 2.06×。
- 纪律：①批量写公式/值，**三样一起挂起**：`suspendEvent()` + `suspendPaint()` + `suspendCalcService()`；②恢复顺序照官方示例：`resumeEvent()` → `resumePaint()` → `resumeCalcService()`；③`suspendEvent` 是 **Workbook 级** API，Designer 里用 `designer.getWorkbook()` 拿到的实例上调（`Worksheet` 上也有）；④挂起期间的事件**不会补发**，靠事件驱动的外部状态（Ribbon、公式栏、自定义监听）要在 resume 后自己刷一次；⑤Univer 没有等价开关，只能整段计时，报告里要标注这个口径差；⑥**写入耗时先问"挂了吗"**，再谈谁快。

### 56. 绘图层"看着一样"要拆成 API / 导入 / 导出三层，否则结论会反

- 事故：先只按 API 判，写成"Univer 无旋转、无翻转"；查 XML 又发现它导出里 `rot`/`flipH` 全是 0，进一步确认。实际按层拆开是三种不同的病：**API 层 Univer 有**（`setTransform({rotation})`、`{flipX:true}` 可设可读，C 组与 I4 全过）；**导入层不解析**（`getSnapshot().transform.rotation` 恒 0、`flipX/flipY` 恒 false，45° 对象尺寸被重算成 80×165、-30° 直接渲染成正矩形）；**导出层不写**（`xfrm` 上没有任何 rot/flip）。SpreadJS 反过来：没有翻转 API，但导入的 `flipH/flipV` 会原样写回导出文件。
- 纪律一：**同一能力必须报三层**——能不能设（API）、读进来还在不在（导入）、写出去还在不在（导出）。只报一层会把"能用但不能存"写成"没有"。
- 纪律二：**绘图层要用元素级对账，`diff_io.py` 那种部件级统计没有分辨力**（三家都是 9 个 `drawingN.xml`）。要数 `xdr:pic/sp/cxnSp/grpSp`、`editAs`、`srcRect`、`rot`/`flipH`、`outerShdw`、`glow`、`gradFill`、`blipFill`、`grayscl`、`hlinkClick`、`media` 去重。
- 纪律三：**先分清"表示法不同"和"丢失"**。①锚定：Excel 用 `twoCellAnchor`+`editAs` 表达三态，Univer 换成 `oneCellAnchor`/`absoluteAnchor`，总数不变，是等价写法；②连接线：`xdr:cxnSp` 换成 `xdr:sp`+`prstGeom prst="straightConnector1"`，样本无 `stCxn/endCxn`（未绑端点），视觉等价；③未裁剪的图片写 `srcRect="0,0,0,0"`、每个形状显式写 `noFill`/`prstDash solid` 都是冗余不是失真——**计数涨了不代表差异，要看值**。
- 纪律四：**对象总数不可跨引擎直接比**。`shapes.all()` 只到顶层（组合算 1 个），`getShapes()` 把组合成员平铺；SJS 的 `shapes` 里还混着 `PictureShape`，Univer 的图片在 `getImages()`。要么按"顶层对象"统一，要么把 shapes/images 拆开报，不能拿 44 vs 46 直接写"Univer 多两个"。
- 纪律五：**`statsUni` 一类采集函数里的字段可能是写死的占位**（本项目 `shapes: 0` 就这么留了两轮），判"导入后没有对象"前先确认统计真的读了引擎。
- 纪律六：截图取证要按**表名**对齐两侧（同一工作簿里两家表顺序会不同），并按表截而不是只截首表——绘图对象的证据全在各分组表里，索引页什么都证明不了。缩放：SJS 是 `spread.getActiveSheet().zoom(f)`（`spread.zoom` 不存在，循环内 `const sh` 到循环外取会静默失败），Univer 是 `ws.zoom(f)`（无 `setZoomRatio`）。

### 57. 引用 Excel 的值，位数不能超过 Excel 能显示的位数

- 事故：差异说明自动拼成「Excel 1.0471975511965976」，客户照着在 Excel 里敲 `=ACOS(0.5)`，格子里是 `1.047197551`，当场质疑数据是编的。查下来三层数字全对，只是**说错了层**：
  - 存：`1.0471975511965976`（IEEE 双精度，oracle 用 COM `Value2` 取到的就是它）
  - 最多能显示：`1.0471975511966`（Excel **任何格式**上限 = 15 位有效数字，16 位小数格式也只会显示成这样）
  - 常规格式实际显示：`1.047197551`（10 位，随列宽还会变）
- 真相：`=ACOS(0.5)-PI()/3` 在 Excel 里得 `0`，计算没有任何问题；问题在"对外引用写了几位"。
- 纪律：①**判定用 15 位，说明文字也用 15 位**（Excel 能显示的上限），17 位只留在 oracle 证据文件；②有 `disp()` 这类"对外渲染"函数就统一走它，别在字符串拼接里直接 `String(v)`——这是本次唯一漏掉的地方（判定走 `numText`，说明走 `str`）；③客户报"你的数和我看到的不一样"，先问**他看的是存储值还是显示值**，别急着改数；④配套 xlsx 的「对比表」保留存值，但表头必须写明"存储值；格子里最多显示 15 位"，口径页再写清三者区别。
- 附带发现：`IMPORT_ONLY` 把 `formula` 设为"只导入"后，`APP.runCases` 一律返回 `null`，`node tools/run_cases.js formula` 静默失败——脚本只打印"没有结果"，读的却是上一次的旧 JSON，看上去像"重跑过、数字没变"。跑批入口要能绕过 UI 闸门（`runCases(m, {force:true})`），并且**跑完必须核对文件里的 `at` 时间戳**，不能只看数字对不对。

### 58. 一侧的导入记录少了 `via` / `convMs`，说明它跑的是旧链路——数字可能差一个数量级

工作台里所有模块的 Univer 导入都会写 `{ms, waitMs, via:'local', convMs}`。**某一模块只写 `{ms}`，就是切换测量链路之前留下的旧值**，不能当现行数字用。

实测：打印模块原记 `3069 ms`（浏览器内 `importSheetToSnapshotAsync` 路径、冷启动一次），换成本机 exchange 后同一样本 **171 ms**（其中转换 34 ms），5 轮中位 **117 ms**，相差 26 倍。

判定与复测顺序：

1. **先看记录里的链路标记**，别先怀疑引擎：带 `via:'local'` + `convMs` 的才是现行口径。
2. **转换与装载分开报**（`convMs` = 服务端纯转换；`ms - convMs` = 传输 + 建模）。
3. **小样本再压一次纯转换**排除固定开销（28 KB 样本 20 ms 上下 ⇒ 无固定开销，成本随单元格数走）。
4. **写清单次还是中位、含不含首帧**：单次在 107～222 ms 之间波动，中位才稳。
5. 复测脚本 `tools/probe_import_ms.js <module>`：每轮新开页面、两侧各导一次、取 N 轮中位 → `out/data/import-ms-<module>.json`。

同源坑：只重跑 `capture_io` 不够——图里的耗时文字是 `compose_io_figure.py` 从 json 读出来画的，**json 更新后必须重画图**（`--doc` 与 `--rt` 两套都要）。

## 2026-11-09 补六：数据透视表模块（71 条双端跑批，含多条结论翻案）

用例分组：P-A 对象与数据源 8 / P-B 字段布局 8 / P-C 聚合与值显示方式 17 / P-D 分组 6 / P-E 排序筛选折叠 7 / P-F 计算字段与计算项 5 / P-G 刷新与改源 4 / P-H 透视图 4 / P-I 切片器 5 / P-J 函数·文案·文件级 7 = **71**。定稿口径：**PASS = 能读回 + 能对上**；导入/导出双口径 = **对象数 + 逐格值**。结果 **SJS 62/71（87.3%）、Univer 52/71（73.2%）**。

### 59. 「切了 culture 还是英文」几乎都是缺核心中文资源包，不是插件内嵌英文

- 事故：透视表字段面板、值字段设置弹框、单元格表头全英文（`PivotTable Fields` / `Row Labels` / `Grand Total`）。工程已经 `CultureManager.culture('zh-cn')`、也引了 `gc.spread.sheets.designer.resource.cn.min.js`，于是第一版归因写成「透视引擎自带英文词表，切 culture 不生效，改不了」——**错了，能改**，而且这条错归因已经流进 SKILL 原则 22 和 `formula-module.md`，本轮一并纠正。
- 三层证据（缺一不可）：
  1. `designer.resource.cn.min.js` 只覆盖 Ribbon 与普通对话框；核心 UI 词条在 `@grapecity-software/spread-sheets-resources-zh`（`gc.spread.sheets.resources.zh.min.js`，19.2.2 约 318 KB）。
  2. 包内静态统计：`gc.spread.pivot.pivottables.min.js` 里 CJK 字符 **0** 个、`zh-cn` 出现 **0** 次 → 它自己不可能有中文。
  3. 运行时真因：`CultureManager.JPt['zh-cn'].PivotTables` **有 358 个键**，值却是 `PivotTable Fields / Row Labels / Value Field Settings`。插件取词走 `new GC.Spread.Common.ResourceManager(GC.Spread.Pivot.SR, 'PivotTables')`，而 `GC.Spread.Pivot.SR.langs` 只有 `['en']`——zh-cn 槽位被英文 SR 兜底**填满**，既走不到「缺词条」分支，也不会被翻译。这正是「明明设了文化却没用」的机制。
- 修复与验证（同页不刷新）：引入官方包后 `SR.langs` 变 `['en','zh']`，同一 358 键同槽位读出 `数据透视表字段 / 行标签 / 列标签 / 总计 / 值字段设置`；面板、Ribbon「数据透视表分析」页签、单元格表头同时变中文。
- **加载顺序是硬约束**：包内注册带 `if (GC.Spread.Pivot && GC.Spread.Pivot.SR)` 守卫，放在 pivot 插件**之前**会整段静默跳过且不自愈。顺序固定 **core → pivot 插件 → resources-zh → designer**。
- 纪律：①「设了 culture 没用」先数**包里有没有中文字符**、再读**运行时 JPt 槽位里的值**，两步就能分清「没实现」和「资源没装」；②「内嵌英文、改不了」这种归因必须答得出「那 358 个键的值是什么」，答不出就是猜；③资源包版本跟主包锁死同版本；④**装了语言包要重跑全量**，文案会变（见 60）。
- 附带（取证入口）：Designer 的透视字段面板没有公开命令通道，`designer.executeCommand()` / `designer.getCommandManager().execute()` 都不通，只有内部方法 `designer._runRibbonCommand('pivotTablePanel')` 能打开；截图前先把 `setActiveCell` 落到透视表区域内，否则面板与表对不上。

### 60. 装语言包会改变单元格文案——主网格比对必须首尾都剥「汇总/合计」

- 事故：装上中文资源包后，SJS 的值字段标题从 `金额` 变「汇总 金额合计」（中文包用前置模板），Excel 是「金额合计汇总」（后缀）。`cmpGrid` 原来只剥尾部，于是把文案差异算成**数值**差异，`P-B-05` 翻 FAIL（总分 62 → 61，修好归一后回到 62）。
- 正确做法：文案归一**首尾都剥**，用 `do { t = 剥(t) } while (t 变化)` 剥到不变为止（「金额合计汇总」要剥两层），纯文案差异不进入数值判定，只由专门那条（`P-J-07`）度量。
- 纪律：①语言包/主题包会改比对文本，装了包必须重跑全量并逐条看翻案项，不能只跑新用例；②「文案」与「数值」要各有一条用例兜住，否则归一化要么吃掉真差异、要么把文案差异算成数值差异；③一轮里出现「已经 PASS 的项莫名翻 FAIL」，先查本轮改了什么公共函数。

### 61. 判「偶发成功」的接口，单次采样等于掷硬币

- 事故：Univer `insertSlicer()` 上一轮 1/3 成功、报告写 PASS；下一轮同一份代码 3/3 抛 `source-missing`，又翻 FAIL。复跑统计：同一链路连开 **8 次，只有 1~2 次**能出选项。
- 纪律：①这类接口固定跑 **3 轮、3/3 才 PASS**；结论写「能渲染，但**不稳定**」+ 成功率区间，不写「支持」也不写「不支持」；②「不稳定」必须带样本数（8 次 / 1~2 次），否则读者分不清是环境抖动还是产品缺陷；③同一模块前后两轮结论打架，先怀疑采样，别急着改判定口径。
- 一般化：凡是「第一次不行、重试就行」的接口，判定式必须是 `轮次 + 成功率`，单次结果只能当线索。

### 62. 「create 抛错」和「render 抛错」要分开判——兜底错误文案会把你带偏

- 事故：Univer 透视表切片器一度被写成「不支持」。拆开是两条通道：
  - **对象通道正常**：`getSlicers()` = 2、`getSource()` 带 `pivotCacheId` / `sourceFieldId`、`setSelection()` 能驱动透视表重算过滤。
  - **UI 通道失败**：打开选项列表时 `getInfo()` 抛 `field-missing`，界面回落到语言包 `error.unsupportedTarget`，显示「**该数据源不支持请求的控件。**」
- 坑就在这句文案：它读起来像「数据源/能力不支持」，实际是取选项那一步拿不到字段。照文案写结论必错。
- 纪律：①控件类能力（切片器、迷你图、表单控件）一律按「**创建 → 数据链路 → UI 渲染**」三段分别取证，一段失败不许推到整个能力；②**语言包里的兜底文案不能当根因**，要看抛出点（`getInfo` 还是 `insertSlicer`）；③新建表上同样抛 `source-missing` 的，才叫入口不可用；导入对象能读能驱动过滤的，只能写「UI 不稳定」。

### 63. 主题/配色类能力不能读单元格属性判，必须比渲染像素

- 事故：三方透视表配色差点写成「都没有配色」。因为 xlsx 里只有 `<pivotTableStyleInfo name="PivotStyleLight16" showRowStripes="1"/>`，没有任何内联样式，Excel COM 取 `Interior.Color` 全是 16777215（白）。
- 真相：样式名是**主题引用**，由各自渲染器现算——
  - SJS：按主题算出表头/总计 `#C0E6F5`，与 Excel 截图**逐像素一致**；
  - Univer：走**硬编码模板**（列头 `#1D32E9` 深蓝白字、左上角 `#ECEEFF`、明细行 `#F4F4F5`），不读 `pivotTableStyleInfo`、不跟主题、还丢了隔行条纹（`showRowStripes="1"` 无效）。
- 附带：Univer 的配色只在**透视对象建成那一刻**套一次。计算字段那张表对象丢失后整表 **0 个带色格**——配色结论必须跟「对象还在不在」一起报。
- 纪律：①`pivotTableStyleInfo` / `tableStyle` / 主题名这类**引用式样式**，属性读回一律不作判据，判据只能落到渲染像素（关键格 RGB + 截图）；②「Excel 读出来是白」不等于「Excel 没配色」，它只是没把主题展开；③配色差异要指出是「主题驱动」还是「硬编码」——前者能跟肤换色，后者是产品缺陷。

### 64. 参数名差一个词就把能力判死：`baseFieldName` 不是 `baseField`

- 事故：「占父级%」SJS 整列空白，直接写进报告当缺口。根因是参数拼成 `baseField`——**不报错、不生效**，等于没设基本字段。
- 改正后：多层行字段（区域 → 城市）出 `0.25 / 0.75`，与 Excel 一致；**单层行字段仍整列空白**（Excel 出 100%），这才是真缺口。导入表那条「累计%」原来 1600%，补 `baseFieldName='区域'` 后变 **35.38%**——顺带证明导入侧没还原基本字段（属**解析缺口**，不是 API 缺口）。
- 枚举值一并记：`PivotShowDataAs.percentOfParent = 7`、`percentOfParentCol = 6`、`percentOfRunningTotal = 11`。
- 纪律：①判「不支持」前，参数名必须从 `.d.ts` 或官方 API 页**抄**下来，不许凭印象拼；②setter 不抛错 ≠ 生效，判据只能是**结果值**；③一个功能有多种模式（多层 / 单层）时分模式报，别用一条失败覆盖整项。

### 65. 「整表空白」通常是观测口径错：对象没了 ≠ 值没了

- 事故：Univer 计算字段那张表原写「整表空白」。实拍是**数值都在、透视对象丢了**（计算字段表退化成静态值），左上角连表头都没了。
- 对策：导入判定永远**双口径**——对象数（22 / 22 / 21）+ 逐格值（1067 / 1067 / 1067）。只报值会漏掉「结构塌陷」，只报对象会漏掉「值算错」。
- 纪律：①写「空白」前分别读一次对象数和值，两者都空才叫空白；②截图里看不见内容，先排除裁切与列宽（Univer 总计列显示 `4.2E+06` 是列宽不足，不是值错）；③「结构没了但值还在」必须用一句话说明白，别让读者以为是两张不同的表。

### 66. Excel 基线截图只能用后台 `PrintWindow`，屏幕截图会拍到别人

- 事故：用 `CopyFromScreen` 截 Excel：黑屏、抢前台把测试页面挤到后面、把输入法候选框拍进图里，重跑几次都不一样。
- 正确做法：COM 里取 `Application.ActiveWindow.Hwnd` 下的 `EXCEL7` 子窗口，`PrintWindow(hwnd, hdc, 2)`（`PW_RENDERFULLCONTENT`）后台抓图——不激活窗口、不带 Ribbon/任务栏，宽度也稳定。配套：Excel 文档窗口先最大化，让三家面板宽高比接近。
- 拼图侧：自动裁白要**排除行号列标那条灰边**，否则要么裁进内容、要么留一条灰带；按面板设 cap 宽高比，别统一拉伸；图注写清「哪张表 / 什么状态」，别只写产品名。
- 纪律：基线截图不许用屏幕抓图；一张合成图里三家必须同尺度、同裁切规则。

### 67. 由脚本重生成的表格会吃掉尾部章节

- 事故：`gen_case_tables.js` 刷新附录 A（逐条用例表）时，把后面的 `## 附录 B`（人工复核清单）和 `## 附录 C`（复现与证据）整段截掉——脚本按自己的结束标记截断，写完就没了，文档少半页才发现。
- 纪律：①生成器只替换自己那一段，**保留其后的原文**（先定位「下一个 `## ` 标题」的位置，做拼接而不是覆盖）；②生成后立刻 `grep -n '^## '` 核对章节数；③凡「脚本生成 + 手工续写」混排的文档都要有这条保护，否则每次重生成都要人工补一遍。

### 68. 人工判定类用例不落工程，就是「报告里的一张表」

- 事故：打印元素矩阵 21 项判完，判定写在生成脚本的 `ITEMS` 常量里，产物只有一份 JSON；报告 §2.8.1 是另抄一遍（序号还写成 1~21，与 JSON 的 `E01~E21` 对不上）；对比台界面里完全看不到这 21 项。结果：改一条判定要动脚本、要记得重出报告、还要记得号；三处任一处漏改，报告与数据就永久性对不上，而且没人发现。
- 正确做法（三层）：①判定表抽成 `out/data/print-matrix-items.json` 当**唯一数据源**，脚本只读不算；②脚本同时出 `print-elements.js` 前端清单，**复核页**（`打印元素矩阵.html`）读它渲染，点行展开三家证据图；③跑批文件里只挂 summary（`print-cases.json` 的 `elements`），不并入 PASS/FAIL 分母。
- 三条纪律：①**人工判定类必须单列一张表**——塞进 PASS/FAIL 表会被 `verdict()` 渲染成「—」，且读者会默认是一个分母；②ID 贯穿「判定表 / 复核页 / 报告表」三处，脚本只能保证前两处同源，**报告是手写的，发报告前逐项核**；③页数、页映射与判定是同一条链，样本布局一动就全链重跑（`rebuild_print_matrix.ps1`），只改一处会让"某页有没有这个元素"集体失效。

## 2026-11-12 补七：剪贴板模块（有头 + 哨兵法，含一条界面判定的翻案）

### 69. 右键菜单项带「▸」不等于「只走子菜单」——SpreadJS 的「选择性粘贴...」是拆分菜单

- 事故：报告写成「SpreadJS 右键『选择性粘贴』是子菜单、对话框只在 Ribbon 下拉里，两处口径不同，客户会问」，被用户当场否掉。这个结论来自看截图里的箭头，没实际点。
- 实测：DOM 是 `gc-ui-contextmenu-splitMenu-container` 里挂两颗——`gc-ui-contextmenu-splitMenu-content`（文字区）+ 兄弟节点 `.gc-ui-contextmenu-sup-container`（右侧箭头）。**点文字区直接弹出 `gc-sjs-paste-special-dialog-container`（366×311，与 Ribbon 弹出的是同一个）**，点箭头才展开子菜单。Excel 的「选择性粘贴(S)…」同样右键即弹（等价 `Ctrl+Alt+V`）。
- 纪律：①「带箭头」只是样式提示，判行为必须**真实点一次**，再看**对话框容器类名**是否出现（SJS 是 `.gc-overlay-gc-dialog2` + `.gc-popup-gc-dialog2` + `.gc-sjs-paste-special-dialog-container` 三件套）；②凡是写「A 入口有、B 入口没有」的结论，两条入口都要点到，且要落到同一个 DOM 证据；③顺带记：右键那行「粘贴选项：」的 6 个图标是 `gc-ui-contextmenu-groupitem`，不是菜单项，按 menuitem 枚举会漏。

### 70. SpreadJS 右键菜单的定位三坑：多套预渲染容器 / 必须先取焦点 / 坐标点击与元素点击分工

- 坑 1（最费时间）：页面上常驻**多套** `.gc-ui-contextmenu-container`（预渲染复用），`.first()` 命中的是不可见那套 → `boundingBox()` 直接返回 `null`、`locator.click()` 卡满 6000 ms 超时报 actionability（不是「找不到元素」，容易误判成选择器错）。正确做法：`page.evaluate` 遍历所有容器里的候选元素，**按 `getBoundingClientRect()` 的可见尺寸自己挑**，再拿中心坐标。
- 坑 2：右击前必须**先用左键点一下目标格取焦点**，否则 `mouse.down({button:'right'})` 根本不进 Host，菜单一次都开不出来（现象是「菜单没开、控制台也没错」）。另外不要用 `mouse.click({button:'right'})`，实测开不出菜单；要 `move → wait 180ms → down(right) → wait 200ms → up(right)`。
- 坑 3：菜单项的点击方式**不统一**——`.gc-spread-pasteAll` 这类必须元素点击（坐标点击能命中 DOM，但 SJS 把命令绑在内部指针事件上，命令不跑）；拆分菜单项反过来，locator 卡超时只剩坐标真点可用。先挑到可见元素，再按项决定用哪种。
- 纪律：写「点了没反应」之前，先把三件事分开取证——菜单开没开 / 有没有开在别处 / 点的是不是可见那一套。

### 71. 无头浏览器的 `navigator.clipboard` 不是 Windows 剪贴板，通道类用例必须换口径

- 事故：`headless: true` 下取证「SpreadJS 右键复制后剪贴板为空」，当成结论写进报告。换有头重跑，真相是：`Ctrl+C` 写 63 B 文本 + 2803 B HTML（HTML 里带 `gc-sjs-clipboard="true"`），**右键复制一个字节都不写**（20 B 哨兵原样留存，`HTML Format` 项根本不出现）；Univer 两条路写出的内容**逐字节相同**。
- 正确口径：`headless: false` + 真 Windows 剪贴板（Python `win32clipboard` 枚举 `EnumClipboardFormats`，逐格式记字节数）+ **哨兵法**：每轮先把系统剪贴板写成哨兵（如 `WINSENTINEL-CTX-0713`），于是「引擎里贴出源区数据」= 走内部缓存、「贴出哨兵」= 走系统剪贴板；再补一步「复制完把剪贴板换回哨兵」，两条通道就完全分离。
- 纪律：①headless 只能证明「Chromium 自己的进程内缓冲没被写」，不能当「系统剪贴板没被写」；②“贴不出内容”还有第三种可能——它压根不读 `navigator.clipboard`，只吃原生 `paste` 事件里的 `clipboardData`（SJS 右键「粘贴」就是这样，程序化点击不产生该事件 → 记 NA，别记 FAIL）；③Excel 侧做基准时 `Visible=False` 的实例不跑消息泵，`IDataObject` 不渲染，Chrome 取 `text/html` 拿到空 —— 要 `Visible=True`；且 `CutCopyMode=False`（按 Esc）会**清空系统剪贴板**，之后 `Paste()` 报「类 Worksheet 的 Paste 方法无效」。

### 72. 记 NA 之前先反证一次：入口查错会让能力凭空消失

- 事故一：剪贴板 J4-08「线程批注」给 SpreadJS 记了 NA，理由写「该引擎无线程批注读回入口」。真实原因只是探针查了 `cell.commentThreaded()`——19.2 里这个方法是 `undefined`，线程批注在 `sheet.threadedComments` 集合上。换入口后源区读到「线程批注 D4」、粘贴到目标格也读到，**改判 PASS**（NA 5→4、可判 71→72、通过率 85.9%→86.1%）。
- 事故二（同一模块早前）：J5-02「样式跟随」判 Univer FAIL，也是读回入口选错——`getCellStyle()` 返回 `{}`、`getCellStyleData()` 才有 `bg/bl`；修入口后 20/20 一致，改判 PASS。
- 正确做法：NA 只能有两种理由——「本引擎确实没有这个对象/字段」（要有 typeof / 枚举证据）与「基准建不起来」（要有 Excel 侧证据）。写之前强制自检三问：①这个入口名是从 `.d.ts` / 官方 API 页抄的，还是凭印象？②同义入口试过几个（单数/复数、cell 级 vs sheet 级、集合 `.all()` vs 单条 `.get()`）？③把同一个读法用在**已知有值**的样本上，能读出来吗（正控）？
- 纪律：**探测不到 ≠ 不支持**；探测不到时先怀疑自己，再怀疑对方。改判要同时在报告里写清「原判什么、为什么错、改后多少」，否则分母悄悄变了没人发现。

### 73. 「官方 demo 有、本地没有」先查插件集，再查版本

- 事故：拿本地工程 Univer 的单元格右键菜单（15 项，且「添加评论」显示成 `sheets-thread-comment-ui.menu.addComment`）与 `office.univer.ai/zh-CN/exchange`（18 项）对比，第一反应是「版本落后 / 功能缺失」。实际是**自建 bundle 少装了 4 个包**：`@univerjs/sheets-hyper-link-ui`（装了 model 没注册 -ui）、`@univerjs/sheets-sort-ui`（压根没装）、`@univerjs/sheets-note` + `sheets-note-ui`（没装）、以及 locale 只 `mergeLocales(core, advanced)` 导致后注册的包文案直出 key。
- 判据：Univer 的菜单项由各包的 `menuSchema`（`SheetsSortUIMenuSchema` / `SheetsHyperLinkUIMenuSchema` / `SheetsNoteUIMenuSchema` …）注册，**注册了 UI 插件才有菜单项**；命令与能力往往在 model 包里（`FRange.sort()` 早就能用，只是没有菜单入口）。
- 纪律：①「demo 有、本地无」先 `ls node_modules/@univerjs | grep <关键字>` 与入口文件的 register 列表对账；②一个能力要分三层查——**model 包（命令/数据）、UI 包（菜单/面板）、facade（API 方法）**，任一层缺失的现象都是「没有」，但原因完全不同；③界面文案变 key 原文 = locale 没合并，属于最容易在客户演示翻车的一类；④别把 bundle 配置差异写成产品差异。
- 附带坑：`npm i` 会按 `package.json` 重解析依赖树，**把没写进 package.json 的间接依赖删掉**（本项目被删过 `react` / `react-dom` / `@univerjs/sheets-hyper-link-ui`，随后 esbuild 报 `Could not resolve "react"`）。动依赖前先备份 `package.json` 与 bundle。

### 74. 菜单项 / 能力「不存在」的第三种原因：测试环境没装对应的包

- 事故：`docs/02` 与 `docs/12` 都把「Univer 传统批注（Excel 的备注，`comments1.xml`）导入即丢」写成产品缺口，持续了两轮。真实原因是 bundle 里**没有 `@univerjs/sheets-note`**，`SHEET_NOTE_PLUGIN` 这个 resource 根本不存在，`range.getNote()` 也就无从读起；补装后同一份样本立刻读到「这是 Excel 传统批注」，且导出件多出 `xl/comments2.xml` + `vmlDrawing2.vml`。
- 判断链：`range.getNote()` 是 `undefined` ≠ 「引擎不支持批注」，也可能只是**插件没注册**。先 `typeof` 一遍 + 看快照 `resources` 里有没有对应的 `XXX_PLUGIN` 节点，再下结论。
- 纪律：①凡是写「导入即丢 / 不支持」，必须同时报**被测 bundle 的插件清单**（本项目的口径是 `window.__proExtraPlugins` 全量落盘）；②补装插件后要重跑**受影响的整组**用例，并把「原判 / 改为 / 分母变化」写进报告——本轮 J4-07 由 NA 改 FAIL，Univer 侧 NA 7→6、可判 69→70、通过率 55.1%→54.3%；③插件补齐不等于结论变好：J4-07 从「测不到」变成「测到了，但粘贴不带批注」，是更差但更真实的结果。

### 75. 对照前先对齐「操作」，不是对齐「用例名」

- 事故：J1-08b 名为「整列复制（列宽是否跟随）」，三方实际操作却不同——Excel 基准跑的是 `Range(R26:R28)` 三格（列内片段），SpreadJS 跑的是 `Range(-1, c, -1, 1)`（真整列），Univer 跑的是 `R26:R28`。结果 Excel 目标列宽 8.11（不跟随）、SJS 169→169（跟随），于是判成「SJS 多做了列宽 → FAIL」；而 Univer 因为和 Excel 都"不跟随"被判 PASS。
- 实测真相（`tools/_probe_j108b_excel.py`）：Excel **片段复制** R26:R28 → T26:T28，目标列宽 8.11 **不跟随**；**整列复制** R:R → T:T，目标列宽 18.00 **跟随**。用户一句「Excel 整列复制也会带列宽」直接推翻原判。
- 改判：三方统一为整列后，SJS PASS、Univer FAIL（`getColumnWidth` 源 150.22 → 目标 70.56，不跟随）。SJS 66/10（86.8%）、Univer 38/36（51.4%）。
- 纪律：①用例名里的「整行 / 整列」必须落到 `Rows(n).Copy()` / `Columns(n).Copy()` / `Range(-1,c,-1,1)` 这类真整行整列 API，`Range(A26:A28)` 是片段不是整列；②写基准脚本时把**被测侧的调用方式抄一遍**再改成宿主 API，不要各写各的；③判据是「两侧的布尔量是否一致」（跟不跟随），但前提是**两侧问的是同一个问题**；④结论方向反了（「比 Excel 多做」）时优先怀疑操作不对等，而不是急着判 FAIL。

## 2026-11-14 补八：两条 FAIL 的自我翻案（探针构造 + 默认规则）

### 76. 对象参数用官方构造器；`x.y is not a function` 先怀疑自己传错类型

- 事故：表模块 B7「自定义表样式」连续两轮判 SpreadJS FAIL，写明「`customTableThemes.all()` 已含 `MyTblTheme`，按该名称建表抛 `v.h5 is not a function`」，并据此写成「注册通道存在、套用通道中断」，占掉 1/7 与 2 个百分点。
- 真相：探针写的是 `th.headerRowStyle({ backColor: '#FFD966', foreColor: '#7F6000' })`——给 `TableTheme` 的样式槽位传了**普通对象**。官方示例一律传 `new GC.Spread.Sheets.Tables.TableStyle(backColor, foreColor, ...)`。对照实验（`tools/_diag_b7_root.js`，同一页面连跑两遍）：
  - 普通对象 → `tables.add(...)` 抛 `v.h5 is not a function`
  - `TableStyle` 实例 → `add=ok repaint=ok toJSON=ok 含主题=true 样式读回=mytbltheme A1底色=#FFD966`
- 四条官方路径全部可用（`tools/_diag_b7_api2.js`，每条单独一页避免区域相交）：建表时第 6 参传样式名 / 建表后 `table.style(name)` / 建表时直接传 `TableTheme` 对象 / `spread.defaultTableTheme(name)` 后建表。`getStyleName()` 读回小写 `mytbltheme`，`toJSON` 保留主题，A1 底色 `#FFD966`，画布像素由灰 `rgb(178,178,178)` 变黄。
- 注意 `table.styleName` 在当前版本是**字符串属性**不是方法（`typeof t.styleName === 'string'`，值为 `medium2` 这类），文档里的 `table.styleName("custom0")` 写法在这个 build 上不成立——写用例前先 `typeof` 一遍，别照抄文档。
- 识别信号：① 报错是压缩后的内部短名（`v.h5` / `e.h0`），不是业务语义；② 抛错点在你调用那一行的下游；③ 一句话里同时出现「注册成功」和「抛错」，很可能是自己把参数塞错了类型。
- 纪律：① SDK 的样式/主题/配置类参数一律用官方构造器（`TableStyle` / `LineBorder` / `Style` / `TableTheme`），不要图省事传字面量对象；② 判 FAIL 前跑一次「最小对照」——同一条路径换成官方写法再跑一遍，两次结果不同即是探针问题；③ 报告里写 `v.h5 is not a function` 这种无名报错等于没写，必须先定位到自己的哪一行。

### 77. 默认行为不符 Excel ≠ 能力缺失：先找补齐参数

- 事故：D14「中文排序规则」判定 FAIL，理由「`sortRange` 默认返回码点序（研发 < 财务 < 销售），与 Excel 拼音序不符」，并与 Univer 并列成「两家同错」。
- 真相：码点序只是**默认**。官方 API 文档的示例本身就用拼音比较——`sortRange(..., [{ index: 0, ascending: true, compareFunction: pinyin }])`，其中 `pinyin = (a,b) => a.localeCompare(b, 'zh')`。实测（`tools/_probe_b7_d14.js`）：默认序 `研发 研发 财务 销售 销售`，补 `compareFunction` 后 `财务 销售 销售 研发 研发`，与 Excel 逐值相同。Univer 侧同一参数被静默丢弃（D4 已证），所以补齐路径只有 SpreadJS 有。
- 判法：**默认规则差异写进「差异说明」，不写进「缺口」**。判定按「能否用公开入口得到与 Excel 一致的结果值」——能，则 PASS，说明列注明「默认码点序，需补 compareFunction」。反过来，若产品连补齐钩子都没有（如 Univer 的 `compareFunction`、`hasTitle`），那才是能力缺口。
- 识别信号：官方参数表里有 `compareFunction` / `locale` / `culture` / `ignoreHidden` 这类钩子，或产品有「选项」类对话框时，默认行为不符先按「可配置」处理，测过再说。
- 纪律：一条判 FAIL 的用例，说明列必须回答两件事——**试过哪些参数**、**补哪个参数能对齐**。只写「默认不一致」不足以判 FAIL。

# API 支持度与友好性探针法（v1，2026-11-05 定版）

回答一个和"导入保真"完全不同的问题：**这套 SDK 有没有代码入口、能不能设、能不能读回、用起来顺不顺手。**

适用场景：客户已经能接受"打开文件长得像"，现在要评估"二次开发能走多远"。图表/格式/条件格式/样式/保护这类模块都适用。

---

## 1. 三个必须先分清的问题

| 问题 | 判据 | 常见误判 |
|:--|:--|:--|
| 有没有入口？ | 运行时 `typeof` + 文档双向查 | 只看文档 → 把"文档有、运行时无"判成支持 |
| 设得进吗？ | set 不抛错 + 读回值等于期望 | 只看 set 不抛错 → 静默失败全被算成"支持" |
| 读得回吗？ | 有可信读回口且值稳定 | 不读 → "调了不知道生效没" |

**核心原则：没有读回口的"支持"最多判 S1⁻，不能算通过。**

---

## 2. 五步管线

### 第 1 步：以 Excel 对象模型拆用例（约 130 项 / 10～12 组）

拆到"一个配置项一行"，不要合并。例：

- 对（可测）：`数据条最小值类型`、`数据条方向`、`图标集反转顺序`
- 错（不可测）：`条件格式支持程度`

每行必须带：`id / 组 / 能力项 / Excel 基线值`。基线值要能在 Excel 里造出来（否则无法判"值不一致"）。

**建议分组**（格式与条件格式实战版，可直接复用）：A 字体 13 / B 填充 8 / C 边框 11 / D 对齐与合并 17 / E 数字格式 25 / F 行高列宽 8 / G 条件格式规则 16 / H 条件格式图形参数 15 / I 条件格式管理 6 / J 样式复用与主题 11 / K 保护与视图 6。

### 第 2 步：写探针表（一个 JS 文件，两家共用一套用例）

每条用例对两家各写一份 `{api, kind, steps, param, doc, set(ctx), read(ctx), exp}`：

| 字段 | 取值 | 报告里怎么用 |
|:--|:--|:--|
| `api` | 公开 API 名（`FRange.setFontColor`） | 直接写进矩阵 |
| `kind` | `prop` / `method` / `builder` / `none` | 衡量"能不能一步到位" |
| `steps` | 完成这件事要几次调用 | 友好性"一步到位"维度 |
| `param` | `typed` / `enum` / `loose` | 友好性"参数类型化"维度 |
| `doc` | 1 有 / 0 无 | 只做"文档覆盖率"参考，不参与 PASS/FAIL |
| `set` | 设值函数（可 async） | — |
| `read` | 读回函数 | — |
| `exp` | 期望值（`undefined` 表示"应该没有值"） | 比对 |
| `assert` | `'noThrow'` = 只判调用不报错 | 用于无读回口的 API：仍算 PASS，但「可读回」列记否 |

模板：`scripts/api_probe_table.js`。

**写法上的三个硬要求：**

1. `kind:'none'` 的项，`set` 抛错要归为"无入口"，不能记为"运行失败"。
2. 每项独立行号，且**行号不得超过工作表行数**（见陷阱 1）。
3. 条件格式这类有状态的组，每项开跑前清空两侧规则，否则规则互相污染。

### 第 3 步：无头跑批

`scripts/api_run_probe.js`：内置静态服务 + Playwright（`channel:'msedge'` 或 `'chrome'`），等两家引擎 ready 后 `page.evaluate(() => window.ApiProbe.runAll())`，落 `api-items.json`。

要点：

- 跑批器里对 Univer 侧每次 set 后 `await tick()`（写是异步命令，见陷阱 3）。
- 跑前把 SpreadJS 行数拉够：`sheet.setRowCount(400)`。
- 结果里要保留 `got / exp / readErr / setOk / ms`，方便回溯"是我的写法错还是引擎错"。

### 第 4 步：定级 + 五维评分

`scripts/api_gen_report.py` 读 `api-items.json`，输出分组定级矩阵与评分。

**判定只有两档（2026-11-05 用户定版，机械化，避免主观）：**

```
noApi 或 set 抛错                → FAIL（备注写清"无公开入口"/"调用即抛错"）
setOk 成功                      → PASS
tierHint = 'S-'（Excel 本没有）  → N/A，不计入分母
```

- 不要再输出 S1/S2/S3/S4/S5：用户明确反馈"分级看着很乱，直接给 pass 和 fail，都遵照公开 api"。
- `doc`（有没有文档）不参与判定，只单列"文档覆盖率"。
- `readable`（能不能读回）不参与判定，单列进五维的「可读回」和附录备注。
- 分母 = PASS + FAIL；N/A 项在附录单独标注（"Excel 本无、我方可做"也是卖点）。

**五维评分（各 0～5，合计 25）：**

| 维度 | 口径 |
|:--|:--|
| 覆盖广度 | PASS 项 / **Excel 基准项** ×5（分母剔除"Excel 本身没有"的 N/A 项，否则对缺能力的一方不公平） |
| 一步到位 | 平均步数映射：1 步 5 / 2 步 3.8 / 3 步 2.4 / ≥4 步 1.2 |
| 可读回 | 能取回等值结果的比例 ×5 |
| 参数类型化 | typed 5 / enum 4 / 裸值或魔法字符串 2，取均值 |
| 错误反馈 | 5 − 0.5×静默失败条数（每条都要有实测证据） |

### 第 5 步：可视化 + 报告

对照台加一个「API 检查」tab（见 `templates/dual-pane-harness/`）：

- 表格列：ID / 组 / 能力项 / Excel 基线 / 两侧 API / **两侧结论**（通过·无入口·不符）/ 两侧读回值 / 备注
- 汇总行直接给"PASS x / FAIL y（无公开入口 n）"
- 支持导出 JSON 留证

报告结构（读者的问法顺序）：

1. 分层口径（保真层 / 可编程层 / 渲染层）+ PASS/FAIL 判定规则
2. 覆盖率 + 分组 PASS/FAIL 矩阵 + 差距最大模块
3. 分族解读（每组 3～5 句，说清差距根因，不说"支持/不支持"就跑）
4. 友好性五维 + **静默失败逐条证据表**（最有说服力的一节）
5. 分模块调用步数
6. 结论与给两家的改进建议
7. 复现命令 + **被纠正的"假失败"留档**
8. 附录：逐项矩阵

---

## 3. 陷阱清单（每条都付过代价）

### 3.1 环境类

1. **SpreadJS 默认工作表只有 100 行**（新版初始 200）。行号超出后：`addSpan` 抛 `Invalid range`，更远处**静默返回 null**；样式写到一个不存在的行也不报错。→ 跑批前 `sheet.setRowCount(400)`。
2. **SpreadJS 试用版会插一个 "Evaluation Version" sheet**，统计时要排除。
3. 本机无 Chrome 时用 `channel:'msedge'`（同 Chromium），装了 Chrome 换 `'chrome'`。

### 3.2 SpreadJS 侧

4. **Style 的属性是"首次赋值才创建"的自有属性**。`typeof new Style().foreColor === 'undefined'`，拼错属性名不报错、只是静默建了个无关键。→ 判"有没有这个能力"不能靠 `typeof style.xxx`。
5. **`conditionalFormats.addRule(new Rule(), ranges)` 不进区域索引**：`count()` +1，但 `getRules(区域)` 返回 0、`removeRuleByRange` 清不掉。只有 `addDataBarRule` / `addIconSetRule` 等便捷方法正常。→ 测参数时用便捷方法建规则，再取回**同一个对象**改参数（实测 `sameRef === true`）。
6. 19.2 压缩版**未导出规则类** `CellValueRule / SpecificTextRule / Top10Rule / TwoScaleRule / ThreeScaleRule`，只有 `NormalConditionRule / StateRule / ScaleRule / DataBarRule / IconSetRule / SparklineRule` 可 `new`。
7. **富文本是"值"不是"方法"**：`setValue(r,c,{richText:[{text,style}]})`，读回必须 `getValue(r,c,SheetArea.viewport,ValueType.richText)`。找 `setRichText` 会得出"不支持"的错误结论。
8. `Style.fontSize` 只收字符串，传裸数字抛 `e.match is not a function`。
9. `removeSpan` **只认两参** `(row,col)`；四参或 Range 对象不报错但取消不了合并；对无 span 的格调用会抛 `Invalid range`。
10. 内部边框落点：`{inside:true}` 的内部横线在首格 `borderBottom`、内部竖线在 `borderRight`，不是 `borderTop/Left`。
11. `priority` 要**先建完所有规则再统一赋值**；交错"加一条改一条"会被引擎重排（实测 3/5/9 变 4/7/9）。
12. `sheet.autoMerge(...)` 调用不报错，但 `getSpan / toJSON / 逐格取值` 都看不到痕迹 → PASS（有公开 API）但「可读回」记否。
13. 单元格状态用 `sheet.cellStates.add(range, CellStatesType, style)`，读回用 `getJSONForCell(r,c)`；`getState` 只反映"此刻是否处于该视觉态"。

### 3.3 Univer 侧

14. **facade 写操作是异步命令**：`setRowHeight(120,20)` 立刻 `getRowHeight` 读到 24（旧值/夹紧值）。→ 每次 set 后 await 一个 tick。
15. **枚举值与 setter 不是一套词**：`Enum.BorderType.BLTR === 'bl_tr'`，但 `setBorder` 只认 `'bltr'`；传 `'bl_tr'` 不报错、静默无效。
16. 对齐只认 3 个字符串：`'left' | 'center' | 'normal'`（**normal = 右对齐**）、`'top' | 'middle' | 'bottom'`；传数字枚举抛 `Invalid horizontal alignment: 6`。
17. **内部边框需要多行/多列区域**：1 行区域上 `setBorder('horizontal')` 静默返回空 `bd`。
18. **条件格式 Facade 在独立入口里**：只注册插件不够，必须 `import '@univerjs/sheets-conditional-formatting/facade'`，否则 `FRange/FWorksheet` 上根本没有 `newConditionalFormattingRule`。
19. `setRowHeight` 有最小值夹紧（设 20 读回 24）；`setWrap(true)` 会自动把行高撑到多行高度，导致 `autoResizeRows` 前后无差值——别据此判"自动行高不支持"。
20. `hideRows / hideColumns` 有写入口但不报错也不变行高，`getHiddenState()` 恒 0、`FRange` 无 `isRowHidden` → PASS 但无可信读回口。
21. **`setNumberFormat` 会把单元格值改成格式化后的字符串**（2026-11-05 实测，影响面极大）：`setValue(0.256)` + `setNumberFormat('0.00%')` 之后 `getValue()` 返回 `"25.60%"`，类型从 number 变 string；千分位/货币/补零/科学计数/日期全部一样；`setNumberFormat('General')` 才会变回数字。→ 任何"设置了格式之后再读值/再导出"的断言都可能踩空；测数字格式时**先读值再套格式**，或套完格式后用 `getDisplayValue()` 断言显示。
22. **`[DBNum1]/[DBNum2]` 中文数字格式不实现**：格式串能设能读回，但显示仍是 `12345`。API 层 PASS、显示层 FAIL——这类"API 通过但显示不对"的项必须单独进渲染专项，不能混进 API 覆盖率。
23. **`aaaa`（星期）跟随运行时 locale**：同一个格式串、同一份文件，本地 OSS（locale=zh-CN）出"星期二"，官方线上环境出"Tuesday"。跨环境对比前先确认 locale。

### 3.5 导入/渲染专项（2026-11-05 增补）

24. **别在没装插件的 demo 里验条件格式**：`office.univer.ai/zh-CN/playground/sheets/charts` 这类页面 `typeof sheet.newConditionalFormattingRule === 'undefined'`，导入带 CF 的文件会**一条规则都不画**，很容易被误判成"导入丢条件格式"。→ 验 CF 必须用带 `UniverSheetsConditionalFormattingPlugin` 的环境，并先 `typeof` 探一下。
25. **"导入丢没丢"用往返导出复核，别靠肉眼**：导入后点「下载文件 → XLSX」，解包比对 `<conditionalFormatting>` 与 `<numFmt>`。规则块数在、格式串在 = 没丢；丢的是子参数（如 dataBar 的 x14 扩展）。本次实测：规则 15/17/3 全在，dataBar 的 gradient/border/negativeBarColor/axisPosition 全丢。
26. **Excel 基线图用 COM 导出 PDF 再转 PNG**：`Worksheet.ExportAsFixedFormat(0, 'C:////temp////x.pdf')` 在 `Application.Visible=$false` 下也能用（不需要 ActiveWindow，那是整表 SelectedSheets 导出才要的），再用 PyMuPDF `get_pixmap(dpi=170)` 转图。比截图稳，也不受锁屏影响。
27. **脚本文件含中文路径会乱码**：以 `-File` 方式跑 .ps1 时按 GBK 读文件，UTF-8 无 BOM 的中文全变乱码。→ 脚本里只写 ASCII，输入文件先复制到 `C:////temp////a.xlsx`，输出也用 ASCII 路径。

### 3.4 通用

21. **esbuild `--global-name` 陷阱**：生成 `var X = (()=>{...})()`，若入口内部先 `window.X = {...}` 再返回 undefined，IIFE 返回值会覆盖全局为 `undefined`——`typeof X === 'undefined'` 但 `Object.keys(window)` 里仍有 `X`。→ 入口自己挂 `window.X`，不要用 `--global-name`。
22. **Playwright `page.evaluate` 只序列化传入函数本身**，外部定义的助手函数在页面里不存在 → 助手必须定义在 evaluate 的函数体内。
23. **不要只用 `headless` 截图判"有没有渲染"**：canvas 渲染器在无头下可能不出图。凡"看不见"的结论必须有头复核。

---

## 4. 已经踩过并必须记住的一条

> **运行时没有某个方法，不等于没有这个能力；能力可能挂在数据结构的形状上。**

实战：A9 上标 / A10 下标 / A13 富文本三项，我先按"`sheet.setRichText` 不存在"判了不支持，查官方文档后发现 SpreadJS 的富文本是**值**——`setValue` 传 `{richText:[...]}`。三项全部改判 PASS。

第二条同样重要：**A 环境看不到 ≠ B 环境不支持**。条件格式在 charts playground 里"整表没画"，看着像导入丢规则；换到带插件的导入页，规则 15/17/3 块全在。下结论前先确认环境装了什么插件、什么版本。

**动作要求**：判"无入口"之前，必须做两件事——①查官方文档该能力的用法示例（不是 API 索引，是示例代码）；②枚举相关命名空间（`Object.keys(GC.Spread.Sheets)` 里出现过 `CellState / CornerPosition / IconPosition` 这类枚举，就是能力的线索）。

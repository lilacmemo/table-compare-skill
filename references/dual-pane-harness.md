# 双栏对照台（工程模板）

用于格式/样式/条件格式这类"看得见、也能用代码设"的模块：同一段设置代码在 SpreadJS 与 Univer 两侧同时执行，直接对比渲染结果。图表/公式模块用脚本跑数据即可；格式、条件格式、打印分页这类**视觉判定为主**的模块用它效率最高。

同一套台子还承载**性能基准**：同一份确定性负载灌进两侧，统一口径计时。功能与性能共用两个引擎实例，靠"重置"切换。

第三个视图是**API 检查**（2026-11-05 新增）：一次跑完 136 项探针，每项走"设 → 读回 → 比对期望值"，按 `set 抛错 / 读回 == exp / doc 有无` 机械定级出 S1～S5，汇总行给覆盖率，筛选器可切"仅某侧无入口"，可导出 JSON 交给 `scripts/api_gen_report.py` 出矩阵与五维友好性评分。探针表在 `js/api-probe.js`，换模块只改 SPEC。它把"两边 API 到底差多少"从几天的文档比对压缩成点一次按钮，且结果自带代码入口名（可直接贴进报告当证据）。

模板本体已随 skill 交付：`templates/dual-pane-harness/`（`index.html` + `js/perf.js` + `js/workload.js` + `js/api-probe.js` + `README.md`）。复制到新工作空间后只需自备 `vendor/` 两个引擎包即可跑起来。

## 结构

```
测试工程/
  index.html                页面骨架 + 功能用例库（CASES 数组）+ 三视图切换
  js/workload.js            性能负载生成器（种子固定，双端同源）
  js/perf.js                双引擎性能适配（setup / write / fullCalc / editRecalc / samples）
  js/api-probe.js           136 项 API 探针表 + 双端跑批器（window.ApiProbe）
  start.cmd                 python -m http.server 8123 + 打开浏览器
  vendor/
    gc.spread.sheets.all.min.js    SpreadJS 19.2.2（@grapecity-software）
    gc.spread.sheets.io.min.js     仅保留运行时依赖，工程不做导入导出
    univer.bundle.js / .css        esbuild 打出的 IIFE，暴露 window.UniverCompare
  samples/                   Excel COM 基线样本 + 生成脚本 + OOXML 补丁脚本
```

## 关键实现

1. **两侧执行器**：`new AsyncFunction('spread','sheet','GC', code)` 与 `new AsyncFunction('univerAPI','wb','sheet', code)`。支持 `await`，`return` 值进日志，异常记 FAIL——异常即"该 API 在该版本不可用/参数不合法"，直接对应判定口径。
2. **Univer bundle 入口**在 `poc-compare/univer_compare_entry.js`：`createUniver({ presets:[UniverSheetsCorePreset({...})], plugins:[UniverSheetsConditionalFormattingPlugin] })`。
   - **core preset 不含条件格式**，必须单独注册 `@univerjs/sheets-conditional-formatting`（OSS 包，非 Pro）。
   - 打包（入口必须放在有 node_modules 的目录下，esbuild 才解析得到依赖）：
     ```bash
     cd poc-compare && ./node_modules/.bin/esbuild univer_compare_entry.js --bundle \
       --format=iife --outfile="../测试工程/vendor/univer.bundle.js" --minify
     ```
3. **SpreadJS 初始化坑**：`new Workbook(host,{sheetCount:1})` 后会有多余空白表，需倒序 `removeSheet`；试用版会追加 "Evaluation Version" 授权页，**删不掉也隐藏不了**，只能填 license。
4. **性能用例的前置条件**：SJS 重建实例要 `destroy()` + 清空 host.innerHTML，再 `new Workbook(host,{sheetCount:N})`；重建后必须回写全局引用（`window.__onSjsReplaced`），否则功能视图还指着已销毁的实例。Univer 用 `createWorkbook({name})` 直接换工作簿，再用 `wb.create(name, rows, cols)` 建表。
5. **性能计时口径**：建表灌值不计时 → 写公式计时（SJS 在 `suspendCalcService` 状态下写，`resumeCalcService` 单独记）→ 全量重算逐轮计时取中位数 → 改一个输入后重算。Univer 用 `executeCalculation()` + `calculationEnd` 事件包 Promise（必须带超时兜底）。跑完读 9 个样本双端比对，证明两边算的是同一份结果。
6. **宿主页 CSS 必须全部走 id 选择器**（2026-11-05 实战踩坑，症状极误导）：Univer 在容器内**自己会渲染 `<header>` 与 `<footer>`**（即使 preset 传了 `header:false, footer:false`）。页面若写裸元素规则 `header{height:52px;background:#1f2937}` / `footer{height:268px}`，会把它们一起改掉，表现为：Univer 面板顶部一条莫名黑条 + 画布只有 257px 高、下方大片空白，**看起来像"Univer 没渲染出产品"**。
   - 定量关系：`canvas 高 = 容器高 − 页面 header 高 − 页面 footer 高 − 1`；换视图容器变 32px，画布也变 32px，但差额恒为 321px——所以"画布跟着容器变了"会骗你以为 resize 没问题。
   - 排查法：遍历 `#uniHost` 子树，打印每层 `getBoundingClientRect().height`，找第一个"比父元素矮"的层，看它是不是 `HEADER`/`FOOTER`。
   - 修法：`#appHd` / `#footer` / `body > main`，别用裸 `header`/`footer`/`main`；同时 `document.querySelector('footer')` 会命中 Univer 内部那个（DOM 里更靠前），取页面 footer 一律用 `getElementById`。

## 样本生成（Excel COM）

- **一个工作簿装全部用例，按组拆 sheet**：比拆成多个文件好用得多——一次拖入官网、一次导入双端、组与组之间不互相干扰；每 group 一个 sheet 名带编号（`G条件格式规则`），便于脚本按名定位。
- 脚本必须存 UTF-8 **BOM**（PS 5.1），否则中文乱码。
- 中文区域下颜色段只认本地化名：`[红色]`/`[蓝色]` 可以，`[Red]`/`[Blue]`/`[Color10]` 直接报"不能设置 NumberFormat"；Excel 保存时会转成标准 OOXML 的 `[Red]`。
- 常规格式不要用 `NumberFormat='General'`（中文环境报错），留默认；文本格式（`@`）必须先设格式再写值。
- **COM 静默失败会污染全部结论**：每个设置都包 try/catch，把失败原因写进该行的"说明"列；生成后解包 XML 数节点（`numFmt`/`xf`/`conditionalFormatting`/`<x14:dataBar>`）做二次核验。
- **失败备注必须写到说明列，不能写数据列**：多个 sheet 的说明列位置不同时，用 `$script:DescCol` 这类全局变量传列号，否则备注会把数据格覆盖成一串错误文本（排查时极难发现，因为规则本身是对的）。

### Excel COM 造不出来的部分 → 改 OOXML

`FormatConditions.Add(Type, Operator, Formula1, Formula2)` 的后 3 个参数在 PowerShell 后期绑定下**传不进** `DISP_E_PARAMNOTFOUND`（传 `[Missing]::Value`、传空串、传四个实参都报"非选择性的参数"）。因此：

- 无法用 COM 创建：**特定文本**（xlTextString）、**日期发生**（xlTimePeriod）两类规则。
- 无法用 COM 设置：DataBar 的 `MinPoint.Type` / `MaxPoint.Type`（报"找不到 Type 的重载"）、`Direction`（任何值都报"不在预期范围内"）。
- 可以绕开的：`AddTop10()` / `AddAboveAverage()` / `AddUniqueValues()` 有无参专用方法，直接用；图标集要用 `wb.IconSets.Item(n)`（不是 `Application.IconSets`，后者返回 Null）；`NegativeBarFormat.Color` 返回 ColorFormat，要写 `.Color.Color`；`TopBottom` 的 bottom = **0**（不是 2），`DupeUnique` 的 unique = **0**（不是 2）。

补法：写一个 `patch_cf.py` 直接改 zip 里的 sheet XML——标准层在 `<conditionalFormatting>` 插 `<cfRule type="containsText">` / `type="timePeriod"`，同时在 `styles.xml` 的 `<dxfs>` 追加 dxf 并同步 `count`；x14 层改 `<x14:cfRule id="{GUID}">` 对应的 `<x14:dataBar>` 的 `direction` 与 `<x14:cfvo>`。**不能用 openpyxl 重新保存**（会把 x14 扩展整段丢掉）。改完用 Excel COM 打开复核规则条数。

## 何时该升级

- 需要"导入保真"两侧同时看 → 接 Univer Pro（`@univerjs-pro/exchange-client` + license + 转换后端）或改走官方 playground 人工对照。**没有授权时不要做互相导入的按钮**，只有一侧能用的功能会误导判断。
- 需要成组回归（几十项一次跑完并出矩阵）→ 保持双栏台做"看效果"，另写 Playwright 脚本调 `CASES` 逐条执行并抓返回值落 JSON，形成可复算的矩阵数据。
- 性能要长期跟踪 → 结果导出 JSON，按"机器 + 日期 + 规模"入库，只看同机相对倍数；换机器不要横比绝对值。

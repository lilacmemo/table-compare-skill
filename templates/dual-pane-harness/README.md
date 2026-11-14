# SpreadJS vs Univer 对照台

一个本地双栏工具：左边 SpreadJS、右边 Univer，**同一份设置代码 / 同一份负载在两侧执行**，肉眼与数字直接对比。三个视图：

- **功能对照**：逐项跑 API 代码，异常即 FAIL（对应"公开 API 调通才算支持"的判定口径）
- **性能基准**：同一份确定性负载灌进两个引擎，统一口径计时（写公式 / 全量重算 / 改值重算）
- **API 检查**：一次跑完 136 项探针（设/读回/抛错三判据），自动出 S 级与汇总，可筛"仅某侧无入口"、可导出 JSON

## 快速开始

双击 `start.cmd`，或在本目录执行：

```bash
python -m http.server 8123 --bind 127.0.0.1
```

浏览器打开 http://127.0.0.1:8123/index.html

## 功能对照视图

| 控件 | 作用 |
| --- | --- |
| 测试项下拉 | 选一条测试项，自动把两侧代码填进底部编辑框，并显示该项的 Excel 基线行为 |
| 执行两侧（Ctrl+Enter） | 分别执行 SpreadJS / Univer 代码，结果显示在右下日志（OK / FAIL + 返回值或异常） |
| 重置 | 两侧清空重建（性能用例跑完后也用它回到干净状态） |
| 授权 | 填 SpreadJS 授权码（存 localStorage），去掉试用水印与评估页 |

代码框可用变量：

- SpreadJS 侧：`spread`（Workbook）、`sheet`（活动表）、`GC`
- Univer 侧：`univerAPI`、`wb`（工作簿）、`sheet`（活动表）

代码可含 `await`（按 async 函数执行），`return` 的值会打印到日志。**抛异常即记 FAIL**。

## 性能基准视图

负载由 `js/workload.js` 生成，种子固定为 42，**双端完全一致**，与之前独立跑的 `poc-compare/perf_workload.js` 是同一套定义：

| 表 | 内容 |
| --- | --- |
| Data | N 行原始数据 + 每行 4 条公式（金额 / 税后 / VLOOKUP / 评级） |
| Report | 120 条 SUMIFS + 120 条 SUMPRODUCT |
| Chain | 1,000 行三级链式依赖（3,000 条公式） |
| Dynamic | UNIQUE / SORT / FILTER / SUMIFS 共 23 条动态数组公式 |

标准规模 9,999 行 = **43,259 条公式**（`Data` 4 条/行 × 9,999 + `Report` 240 + `Chain` 3,000 + `Dynamic` 23）。**公式条数按实际写公式的列算**——`Data` 有 5 列数据但只写 4 列公式，按 5 列统计会多算近 1 万条。

计时口径（与 `poc-compare/perf_run.js` 一致）：

0. **每轮新开页面**——同一页面连跑会让 Univer 第 2 轮重算翻倍（旧簿 `dispose()` 后公式 worker 残留），倍数会凭空变大
1. 建表 + 灌原始值 —— **不计时**
2. 写全部公式 —— 计时（SpreadJS 在 `suspendCalcService` 状态下写，恢复计算单独计时）
3. 强制全量重算 N 轮 —— 每轮独立计时，取中位数（不取均值，单次抖动可达 27%）
4. 改一个输入值后重算 —— 计时
5. 读 9 个核对样本 —— 双端比对，验证是同一份结果

> 写入耗时不能当引擎指标：SpreadJS 侧套了 Designer，每次 `setFormula` 都要通知 Ribbon/状态栏（实测 3.2 s vs 裸 Workbook 1.1 s）。比引擎只看重算那一栏。

结果表可「导出结果」为 JSON，便于多轮追加汇总。跑完后两侧画布会停在 Data 表，回功能对照前点「重置」。

实测参考（2026-11-05，本机 Chrome 无头，9999 行 / 3 轮）：
SpreadJS 重算中位数 3,068 ms，Univer 13,103 ms（4.27×），9/9 样本一致，写公式 952 ms vs 1,502 ms。

## API 检查视图

点顶部「API 检查」→「跑全部用例」，双端各跑 136 项探针，结果表每项一行：

| 列 | 含义 |
| --- | --- |
| 组 / 项 | A~K 组别与用例名 |
| 侧别 | SJS / Univer 的入口名（`api` 字段，直接显示"实际调的什么方法"） |
| 级别 | S1 有文档可编程 / S1⁻ 可读不回 / S2 可编程无文档 / S3 仅界面或调用抛错 / S4 值在无入口 / S5 丢失 |
| 设 / 读回 | 写入是否抛错、读回值是否等于期望值 |
| 步数 | 完成这一件事要几次调用（友好性"一步到位"维度） |

汇总行给 `可编程 x/y（无公开入口 n，不符 0）`；筛选器可切「仅两侧有差异」「仅 SJS 无入口」「仅 Univer 无入口」「仅失败项」；「导出 JSON」出逐项原始记录，交给 `scripts/api_gen_report.py` 生成矩阵与五维评分。

**探针表在 `js/api-probe.js`**（`window.ApiProbe = { SPEC, runAll, runSide }`）。每项 spec 形状：

```js
{ id:'A1', name:'字体名', kind:'prop|method',
  sjs:{ api:'sheet.setStyle → font', kind:'prop', steps:1, param:'typed|enum|loose', doc:1,
        set(c){…}, read(c){…}, exp:'期望读回值' },
  uni:{ … , noApi:true, tierHint:'S3', note:'运行时无公开入口' } }
```

三条判据机械化定级，不做主观判断：**set 抛错 → S3；读回 == exp → 可读回；doc=0 → S2**。`exp` 留 `undefined` 表示"读到非空即算可读"（用于"清除填充/清边框"这类期望为空的项）。换模块时只改 SPEC，跑批与报告脚本不用动。

## 怎么加用例

**功能用例**：编辑 `index.html` 里的 `CASES` 数组：

```js
{ id: 'G12', title: '数据条', excel: 'Excel 基线：…', sjs: `…`, uni: `…` }
```

**性能用例**：在 `js/perf.js` 的引擎对象里加一个方法（如 `sortBench`），再在 `run()` 里串联；
若要换负载，改 `js/workload.js`（改完记得双端同源，且记录口径变化）。

## 已知边界（不要误读）

0. **宿主页面的 CSS 只能用 id 选择器，禁止裸元素选择器**（本模板踩过，症状极其误导）。Univer 内部会渲染自己的 `<header>` 和 `<footer>`，页面里写 `header { height:52px; background:#1f2937 }` / `footer { height:268px }` 会连带命中它们，表现为：①Univer 面板顶部出现一条莫名黑条；②画布只有 257px 高、下面一大片空白，看起来像"产品没渲染出来"。实测公式：`canvas 高 = 容器高 − 页面 header 高 − 页面 footer 高 − 1`，两个视图都严丝合缝。本模板已改成 `#appHd` / `#footer` / `body > main`。排查手法：遍历 `#uniHost` 子树打印每层 `getBoundingClientRect().height`，找"子元素比父元素矮"的那一层，看它是不是 `<header>/<footer>`。
1. **Univer 侧是本地 OSS 构建**（`@univerjs/preset-sheets-core` 1.3.1 + 条件格式插件）。格式与条件格式本身是 OSS 插件；涉及 Pro 专属能力的结论仍以 office.univer.ai 为准。
2. **本工程不做 xlsx 导入导出**：Univer 的 xlsx 通道在 Pro 的 exchange-client 里、依赖转换后端，未授权跑不通；与其做成"只有一侧能用"的半吊子功能，不如直接用 `samples/` 的样本文件到官方 playground 手工对照。
3. **SpreadJS 试用版**会插入 "Evaluation Version" 授权页（删不掉、也隐藏不了），填授权码后消失。
4. 两侧字体渲染引擎不同（都是 canvas 但排版实现不同），**亚像素级字号差异不算缺陷**，判断看格式语义。
5. 性能数字是**同机同页面双实例并存**下测的，绝对值会随机器负载波动；**只看两侧相对倍数**，不要跨机器横比绝对值。

## 样本文件（`samples/`）

`samples/格式与条件格式-全量用例-A至K.xlsx` —— 单个工作簿 12 个 sheet，覆盖 A~K 全部 135 项：

| Sheet | 覆盖 | 项数 |
| --- | --- | --- |
| 00索引 | 组别 / 项数 / 批次 / 覆盖内容 | — |
| A字体 | 字体名/字号/粗斜体/下划线 4 种/删除线/颜色/上下标/主题字体/富文本 | 13 |
| B填充 | 纯色/图案类型/前景背景色/渐变 3 项/无填充 | 8 |
| C边框 | 四边/内部横竖/线型 13 种/颜色/双对角线/斜线表头/清除/合并区/圆角 | 11 |
| D对齐 | 水平 8 种/垂直 5 种/缩进/换行/旋转/竖排/缩小填充/跨列居中/RTL/合并/省略符/内边距/自动合并 | 17 |
| E数字格式 | 常规到 DBNum、自定义占位符、多段/条件段/颜色段/区域码、SJS 独有语法 | 25 |
| F行高列宽 | 行高/列宽/AutoFit/隐藏/默认值/导入保真 | 8 |
| G条件格式规则 | 单元格值/文本/日期/空值/公式/前N/平均/重复唯一/数据条/色阶/图标集/迷你图 | 16 |
| H条件格式参数 | 数据条 min/max/正负色/轴/实心渐变/方向/边框/仅条；色阶；图标集类型/阈值/反转/仅图标 | 15 |
| I条件格式管理 | 优先级/stopIfTrue/查询/修改/删除/多规则叠加 | 6 |
| J样式复用 | 命名样式/复用/部分应用/主题色 tint/主题字体/色板/复制格式/折角/图标/状态 | 11 |
| K保护与视图 | 锁定/公式隐藏/保护/网格线/零值/冻结窗格 | 6 |

**D 组第 21~24 行是最直接的对照**：四行文本完全相同、列宽完全相同，只有属性不同——
`D-ref` 无设置（溢出）/ `D4` 自动换行（折 5 行）/ `D7` 缩小字体填充（单行小字）/ `D15` 省略符（Excel 侧等同 D-ref）。

重新生成（两步，缺一不可）：

```powershell
powershell -ExecutionPolicy Bypass -File samples\build_samples_all.ps1   # Excel COM 造样本
python samples\patch_cf.py                                               # 补 COM 造不出的部分
```

`patch_cf.py` 存在的原因：Excel COM 的 `FormatConditions.Add` 有 3 个可选参数，PowerShell 后期绑定传不进 `DISP_E_PARAMNOTFOUND`，所以「特定文本」「日期发生」两类规则、以及数据条的 min/max 类型与方向无法用 COM 创建，只能直接改 OOXML 补上（改完用 Excel 打开校验过：G 组 15 条规则、H 组 18 条规则齐全）。

## 目录

```
测试工程/
  index.html                  页面骨架 + 功能用例库 + 三视图逻辑
  js/workload.js              性能负载生成器（种子 42，双端同源）
  js/perf.js                  双引擎性能适配（setup / write / fullCalc / editRecalc / samples）
  js/api-probe.js             136 项 API 探针表 + 双端跑批器（window.ApiProbe）
  start.cmd                   一键起本地服务
  vendor/                     SpreadJS 运行时 + Univer bundle
  samples/                    全量基线样本 + 生成脚本 + OOXML 补丁脚本
```

## 重建 Univer bundle

Univer 侧被 esbuild 打成 IIFE（`vendor/univer.bundle.js` + `.css`）。改了 `poc-compare/univer_compare_entry.js` 后：

```bash
cd ../poc-compare
./node_modules/.bin/esbuild univer_compare_entry.js --bundle --format=iife \
  --outfile="../测试工程/vendor/univer.bundle.js" --minify
```

# table-sdk-compare

表格类 SDK 功能对比方法论（以 Excel 对象模型为基准线）。产出带证据的竞争分析报告：SpreadJS vs Univer 已完整跑通 12 个模块。

这个仓库**就是技能本体**——克隆到 `~/.workbuddy/skills/table-sdk-compare` 即可被 WorkBuddy 识别，无需安装步骤。

---

## 1. 前置依赖

| 类别 | 需要 | 缺失后果 |
|---|---|---|
| 系统 | **Windows 10/11 + Microsoft Excel（桌面版）** | 没有 Excel 就没有"基准线"，图表/打印/公式三大模块无法出结论 |
| 运行时 | Node.js ≥ 18、Python ≥ 3.9 | 跑批与出图全断 |
| Node 包 | `playwright`（含 Chrome）、`pngjs`、`esbuild` | 截图闭环、拼图、bundle 构建不可用 |
| Python 包 | `openpyxl`、`Pillow`、`numpy` | xlsx 解析、图像处理不可用 |
| MCP | `spreadjs-mcp`（必装）、`gcexcel-mcp`（可选） | SpreadJS 侧 API 判定失去首要证据通道 |

**Excel 是硬门槛**。Mac 上只能做 API 层对比，报告里必须注明"无 Excel 基准列"，且不得给出"一致/不一致"判定——只能写"两家行为差异"。

## 2. 安装

```bash
git clone <内部仓库地址> ~/.workbuddy/skills/table-sdk-compare
```

克隆完重开一个对话，技能即生效。可用以下方式确认：

```
你：列出 table-sdk-compare 的目录结构
```

后续更新：

```bash
cd ~/.workbuddy/skills/table-sdk-compare && git pull
```

> 不想用 git 也能用：把仓库打成 zip，解压到同一路径，效果一样，只是没有版本追踪。

## 3. MCP 配置（必做，且必须用**自己的** token）

在 `~/.workbuddy/mcp.json` 里加：

```json
{
  "mcpServers": {
    "spreadjs-mcp": {
      "type": "http",
      "url": "https://mcp.grapecity.com.cn/mcp/spreadjs",
      "headers": { "token": "<你自己的 GrapeCity MCP token>" }
    },
    "gcexcel-mcp": {
      "type": "http",
      "url": "https://mcp.grapecity.com.cn/mcp/gcexcel",
      "headers": { "token": "<你自己的 GrapeCity MCP token>" }
    }
  }
}
```

写完后到连接器管理页右上角的自定义连接器入口点"信任"才会激活。

**token 是个人凭据，不要提交进仓库、不要发给别人。** 本仓库不含任何 token。

## 4. 首次自检

```bash
node scripts/bootstrap.js                                   # 查 Node/Python/Excel/Chrome/MCP 是否齐备
node scripts/preflight_univer.js --root <你的工程根> \
  --module core --url <页面地址>                             # 查 Univer 包与插件是否装齐
```

两个脚本都是只读检查，退出码非 0 即有问题。**`preflight_univer.js` 不过，不许开跑**——这是本方法论最贵的一条教训：方法变 `undefined` 是装配问题，不是对方没能力，历史上因此误判过 6 次。

## 5. 目录结构

```
SKILL.md                     方法论主文件（模型读这个）
references/
  verdict-framework.md       判定尺子：S1~S5、PASS/FAIL 红线、话术规则
  dimensions-template.md     维度清单模板（A~N 组）
  pitfalls.md                坑全量（写报告时的反面清单）
  workbench-architecture.md  双引擎对比台架构与插件装配
  api-probe-method.md        API 探测方法
  formula-module.md          公式模块口径（支持度/准确度双分母）
  print-pdf-module.md        打印模块四层口径
  dual-pane-harness.md       双栏对照台搭建
  spreadjs-feature-inventory.md  官方功能清单
  SpreadJS 功能清单（v19.2）.xlsx  清单原始数据
scripts/                     可复用脚本（样本构造/探针/跑批/拼图/自检）
templates/dual-pane-harness/ 可直接起跑的双栏工程骨架
```

## 6. 怎么触发

在对话里给出模块与目标即可，例如：

- "对比 SpreadJS 和 Univer 的数据透视表能力"
- "验证这份 xlsx 在两家的导入保真度"
- "做一个图表模块的竞争分析报告，Excel 为基准"

方法论会按 **样本先行 → 环境门禁 → 计划 → 文档查证 → 跑批 → 四段式报告** 推进。

## 7. 硬规矩（改之前先读 CHANGELOG）

1. **基准线一律 Excel 本体**，没有 Excel 列只能写"行为差异"。
2. **样本先行**：先造 xlsx 样本，再谈测试项。样本本身是交付物。
3. **SpreadJS 侧 API 判定第一动作走 spreadjs-mcp**，不许凭记忆或网页搜索下结论。
4. **Univer 侧跑批前必须过环境门禁**。
5. **报告四段式**：结论先行 → 导入导出对比 → 三方截图 → API 完整度验证。
6. **报告只写本次结果**，不写"上一版错了/本轮纠错/改判轨迹"，也不写测试方法叙事。
7. **API 通 ≠ 效果对**：界面上是降级效果就判 FAIL。

## 8. 口径会变，且必须留痕

改判定口径、改报告结构、改分母，都要：

- 改 `VERSION`
- 在 `CHANGELOG.md` 写清**改了什么、为什么改、影响哪些已出报告**
- 已发布的报告不回溯修改，除非用户明确要求

对外的历史报告请注明生成时的版本号，否则数字无法对账。

// Univer 侧打包入口（esbuild → vendor/univer.bundle.js + .css）
//
// 重要：core preset 只带 sheets / sheets-ui / formula / numfmt 四个插件。
// 下面这些能力都要「插件 + facade」成对引入，缺任一半就表现为「方法不存在」：
//   - 条件格式  @univerjs/sheets-conditional-formatting（+ /facade）
//   - 表 Table  @univerjs/sheets-table（+ /facade 才有 addTable / setTableName …）
//   - 筛选      @univerjs/sheets-filter（+ /facade 才有 FRange.createFilter）
//   - 数据验证  @univerjs/sheets-data-validation（+ /facade 才有 FRange.setDataValidation）
// 排序没有独立包、也没有 Facade API：只有表内排序命令 set-table-sort-state，
// 自由区域排序（SJS 的 sortRange）在本版本没有公开入口 —— 这是实测结论，不是漏引包。
import { createUniver, defaultTheme, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import { UniverSheetsConditionalFormattingPlugin } from '@univerjs/sheets-conditional-formatting'
import { UniverSheetsTablePlugin } from '@univerjs/sheets-table'
import { UniverSheetsFilterPlugin } from '@univerjs/sheets-filter'
import { UniverSheetsDataValidationPlugin } from '@univerjs/sheets-data-validation'
// 界面侧：筛选按钮（列头漏斗）/ 表样式的可视化编辑在 *-ui 包里。
// 只装功能包（sheets-filter）不装 UI 包时，Facade 全都能调，但列头看不到筛选按钮、
// 表样式也没有可视化编辑入口 —— 做证据截图时必须一区分，不能把「没装 UI 包」
// 说成「产品没有」。
import { UniverSheetsFilterUIPlugin } from '@univerjs/sheets-filter-ui'
import { UniverSheetsTableUIPlugin } from '@univerjs/sheets-table-ui'
import '@univerjs/sheets-conditional-formatting/facade'
import '@univerjs/sheets-table/facade'
import '@univerjs/sheets-filter/facade'
import '@univerjs/sheets-data-validation/facade'
import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/sheets-filter-ui/lib/index.css'
import '@univerjs/sheets-table-ui/lib/index.css'

let univerAPI = null

window.UniverCompare = {
  get api() { return univerAPI },

  async init(hostId) {
    const created = createUniver({
      locale: LocaleType.ZH_CN,
      locales: { [LocaleType.ZH_CN]: mergeLocales(UniverPresetSheetsCoreZhCN) },
      theme: defaultTheme,
      presets: [
        UniverSheetsCorePreset({
          container: hostId,
          header: false,
          toolbar: false,
          footer: false,
          contextMenu: false,
          formulaBar: false,
        }),
      ],
      plugins: [
        UniverSheetsConditionalFormattingPlugin,
        UniverSheetsTablePlugin,
        UniverSheetsFilterPlugin,
        UniverSheetsDataValidationPlugin,
        UniverSheetsFilterUIPlugin,
        UniverSheetsTableUIPlugin,
      ],
    })
    univerAPI = created.univerAPI
    window.univerAPI = univerAPI
    window.__wb = univerAPI.createWorkbook({ name: 'compare' })
    return true
  },

  // 清空并重建一个 200 行 × 26 列的空工作簿
  async reset(rows = 200, cols = 26) {
    const old = window.__wb
    if (old) { try { old.dispose && old.dispose() } catch (e) {} }
    const wb = univerAPI.createWorkbook({ name: 'compare' })
    window.__wb = wb
    const sheets = wb.getSheets ? wb.getSheets() : []
    // createWorkbook 自带一个空表，统一改名为 Sheet1 并扩到目标行列
    if (sheets.length) {
      const s = sheets[0]
      try { s.setName('Sheet1') } catch (e) {}
      try { wb.setSheetRowCount && wb.setSheetRowCount(s.getSheetId(), rows) } catch (e) {}
    }
    return true
  },

  // 快照导入（从官方 playground 导出的 JSON 快照可直接粘贴）
  async loadSnapshot(snapshot) {
    const wb = univerAPI.createWorkbook(snapshot)
    window.__wb = wb
    return true
  },
}

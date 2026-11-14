# 生成"格式与条件格式"全量测试用例（A~K 组）单工作簿基线样本（Excel COM）
# 用法: powershell -ExecutionPolicy Bypass -File build_samples_all.ps1
# 输出: 格式与条件格式-全量用例-A至K.xlsx（12 个 sheet：索引 + A~K）
$ErrorActionPreference = 'Stop'
$outDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outFile = Join-Path $outDir '格式与条件格式-全量用例-A至K.xlsx'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

$wb = $excel.Workbooks.Add()
while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
$wb.Sheets.Item(1).Name = '00索引'

function RGBv($r, $g, $b) { return $r + $g * 256 + $b * 65536 }

function NewTab($name) {
    $last = $wb.Sheets.Item($wb.Sheets.Count)
    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $last)
    $ws.Name = $name
    $ws.Cells.Font.Name = '微软雅黑'
    $ws.Cells.Font.Size = 10
    return $ws
}

function Header($ws, $title, $cols) {
    $ws.Cells.Item(1, 1).Value2 = $title
    $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $cols)).Merge()
    $ws.Cells.Item(1, 1).Font.Bold = $true
    $ws.Cells.Item(1, 1).Font.Size = 14
    $ws.Cells.Item(1, 1).Interior.Color = RGBv 31 41 55
    $ws.Cells.Item(1, 1).Font.Color = RGBv 255 255 255
    $ws.Cells.Item(1, 1).HorizontalAlignment = -4131
    $ws.Rows.Item(1).RowHeight = 26
}

function ColHead($ws, $row, $labels) {
    $i = 1
    foreach ($l in $labels) {
        $ws.Cells.Item($row, $i).Value2 = $l
        $i++
    }
    $ws.Range($ws.Cells.Item($row, 1), $ws.Cells.Item($row, $labels.Count)).Font.Bold = $true
    $ws.Range($ws.Cells.Item($row, 1), $ws.Cells.Item($row, $labels.Count)).Interior.Color = RGBv 229 231 235
    $ws.Range($ws.Cells.Item($row, 1), $ws.Cells.Item($row, $labels.Count)).HorizontalAlignment = -4108
    $ws.Rows.Item($row).RowHeight = 20
}

# 标准 5 列项行：编号 | 配置项 | 示例（承载格式） | 数值 | 说明
function Item($ws, $r, $code, $name, $text, $num, $desc) {
    $c1 = $ws.Cells.Item($r, 1)
    $c1.Value2 = [string]$code
    $c1.Font.Name = 'Consolas'
    $c1.Font.Bold = $true
    $c1.Font.Size = 9
    $c1.HorizontalAlignment = -4108
    $ws.Cells.Item($r, 2).Value2 = [string]$name
    if ($null -ne $text) { $ws.Cells.Item($r, 3).Value2 = [string]$text }
    if ($null -ne $num) { $ws.Cells.Item($r, 4).Value2 = [double]$num }
    if ($null -ne $desc) { $ws.Cells.Item($r, 5).Value2 = [string]$desc }
    $ws.Cells.Item($r, 2).VerticalAlignment = -4108
    $ws.Cells.Item($r, 3).VerticalAlignment = -4108
    $ws.Cells.Item($r, 5).VerticalAlignment = -4108
}

$script:DescCol = 5
function Try-Set($ws, $r, $what, [scriptblock]$sb) {
    try { & $sb } catch {
        $dc = $script:DescCol
        $cur = [string]$ws.Cells.Item($r, $dc).Value2
        $ws.Cells.Item($r, $dc).Value2 = ($cur + ' || COM失败(' + $what + ')：' + $_.Exception.Message).Trim()
    }
}

function Try-SetQuiet([scriptblock]$sb) {
    try { & $sb } catch { }
}

function BaseCols($ws, $c3, $c5) {
    $ws.Columns.Item(1).ColumnWidth = 9
    $ws.Columns.Item(2).ColumnWidth = 24
    $ws.Columns.Item(3).ColumnWidth = $c3
    $ws.Columns.Item(4).ColumnWidth = 16
    $ws.Columns.Item(5).ColumnWidth = $c5
    $ws.Columns.Item(5).WrapText = $true
    $ws.Columns.Item(2).WrapText = $true
    $ws.Rows.Item(2).RowHeight = 22
}

# ============================================================
# 00 索引
# ============================================================
$ws = $wb.Sheets.Item(1)
Header $ws '格式与条件格式 · 全量测试用例（A~K 组，Excel 365 基线）' 5
$ws.Cells.Item(2, 1).Value2 = '说明'
$ws.Cells.Item(2, 1).Font.Bold = $true
$ws.Cells.Item(3, 1).Value2 = '本工作簿为 Excel COM 生成的"基线行为"样本：每一行/块对应测试计划中的一个配置项，可直接拖入 SpreadJS 或 Univer 官方 playground，观察导入后是否还原、以及与下方"说明"列描述的 Excel 行为是否一致。'
$ws.Range('A3:E3').Merge()
$ws.Cells.Item(3, 1).WrapText = $true
$ws.Rows.Item(3).RowHeight = 46
$ws.Cells.Item(4, 1).Value2 = '说明列若出现"COM 失败"，表示该项无法用 Excel COM 直接构造（多为 Excel 本身无此对象模型能力），该行仅作 API 对照锚点，不代表 Excel 的行为。'
$ws.Range('A4:E4').Merge()
$ws.Cells.Item(4, 1).WrapText = $true
$ws.Rows.Item(4).RowHeight = 32

ColHead $ws 6 @('Sheet', '组', '项数', '批次', '覆盖内容')
$idx = @(
    @('A字体', 'A 字体', 13, '批次1', '字体名/字号/粗斜体/下划线 4 种/删除线/颜色/上下标/主题字体/富文本'),
    @('B填充', 'B 填充', 8, '批次1', '纯色/图案类型/图案前景背景色/线性渐变/停靠点/渐变类型/无填充'),
    @('C边框', 'C 边框', 11, '批次1', '四边/内部横竖/线型 13 种/颜色/双对角线/斜线表头/清除/合并区/圆角边框'),
    @('D对齐', 'D 对齐与合并', 17, '批次1', '水平 8 种/垂直 5 种/缩进/换行/旋转/竖排/缩小填充/跨列居中/RTL/合并 4 类/省略符/内边距/自动合并'),
    @('E数字格式', 'E 数字格式', 25, '批次1', '常规/数值/货币/会计/日期时间/百分比/分数/科学/文本/DBNum/自定义占位符/多段/条件段/颜色段/区域码'),
    @('F行高列宽', 'F 行高列宽', 8, '批次3', '行高/列宽/自动调行高/自动调列宽/隐藏行列/默认值/导入保真'),
    @('G条件格式规则', 'G 条件格式-规则类型', 16, '批次2', '单元格值/特定文本/日期发生/空值/公式/前N项/前N%/后N/平均/重复/唯一/数据条/色阶/图标集/迷你图'),
    @('H条件格式参数', 'H 条件格式-图形参数', 15, '批次2', '数据条 min/max/正负色/负值轴/实心渐变/方向/边框/仅条；色阶 min/max/中点；图标集类型/阈值/反转/仅图标/综合'),
    @('I条件格式管理', 'I 条件格式-管理', 6, '批次2', '优先级/stopIfTrue/查询命中/修改规则/删除三种/多规则叠加顺序'),
    @('J样式复用', 'J 样式复用与主题', 11, '批次3', '命名样式/样式复用/部分应用/主题色 tint/主题字体/标准色板/复制格式/导入保真/折角/图标/状态样式'),
    @('K保护与视图', 'K 保护与视图', 6, '批次3', '锁定/公式隐藏/工作表保护/网格线/零值显示/冻结窗格')
)
$r = 7
foreach ($x in $idx) {
    $ws.Cells.Item($r, 1).Value2 = $x[0]
    $ws.Cells.Item($r, 2).Value2 = $x[1]
    $ws.Cells.Item($r, 3).Value2 = [int]$x[2]
    $ws.Cells.Item($r, 4).Value2 = $x[3]
    $ws.Cells.Item($r, 5).Value2 = $x[4]
    $ws.Rows.Item($r).RowHeight = 30
    $ws.Cells.Item($r, 5).WrapText = $true
    $ws.Cells.Item($r, 1).Font.Name = 'Consolas'
    $r++
}
$ws.Cells.Item($r + 1, 1).Value2 = '合计 135 项（A12 经核查为 Excel 本身无此能力，不计入；A13/E24 等补充项按计划归组）'
$ws.Cells.Item($r + 1, 1).Font.Bold = $true
$ws.Range($ws.Cells.Item($r + 1, 1), $ws.Cells.Item($r + 1, 5)).Merge()
$ws.Columns.Item(1).ColumnWidth = 18
$ws.Columns.Item(2).ColumnWidth = 24
$ws.Columns.Item(3).ColumnWidth = 8
$ws.Columns.Item(4).ColumnWidth = 10
$ws.Columns.Item(5).ColumnWidth = 96
$ws.Rows.Item(2).RowHeight = 20
$ws.Range('A3:E4').Font.Size = 10

# ============================================================
# A 字体
# ============================================================
$ws = NewTab 'A字体'
Header $ws 'A 组 字体（13 项）' 5
ColHead $ws 2 @('编号', '配置项', '示例（文本，观察字形）', '数值示例', '说明 / Excel 侧限制')
BaseCols $ws 40 46

$aRows = @(
    @('A1', '字体名称 宋体', '宋体字体示例 AaBb 0123 中文', '设为宋体，与本表默认的微软雅黑对照'),
    @('A2', '字号 20 磅', '20 磅字号 Aa 中文', '支持半磅，如 10.5'),
    @('A3', '加粗', '加粗文本 Bold Aa', '字形：加粗'),
    @('A4', '斜体', '斜体文本 Italic Aa', '字形：倾斜'),
    @('A5-1', '下划线：单线', '单下划线示例 Aa', 'xlUnderlineStyleSingle'),
    @('A5-2', '下划线：双线', '双下划线示例 Aa', 'xlUnderlineStyleDouble'),
    @('A5-3', '下划线：会计单线', '会计单下划线 Aa', '线不压字符，跨整格宽'),
    @('A5-4', '下划线：会计双线', '会计双下划线 Aa', '线不压字符，跨整格宽'),
    @('A6', '下划线颜色（蓝，字色深红）', '蓝下划线 + 深红字', 'Excel 下划线颜色随字体色，无独立属性'),
    @('A7', '删除线', '删除线文本 Strike', '单删除线'),
    @('A8', '字体颜色 深红', '深红字 RGB(192,0,0)', 'RGB(192,0,0)'),
    @('A9', '上标（前 4 字）', 'AB12 上标示例', '字符级设置，前 4 字为上标'),
    @('A10', '下标（前 4 字）', 'AB12 下标示例', '字符级设置，前 4 字为下标'),
    @('A11', '主题字体（正文/标题）', '正文主题字体示例', 'Excel COM 无主题字体 API，需 UI 设置（本行不预设）'),
    @('A13', '富文本（单元格内多段）', 'AAAABBBBCCCC 富文本', '前 4 字红色加粗 14 磅 / 中 4 字蓝色斜体 / 后 4 字黑色')
)
$r = 3
foreach ($x in $aRows) { Item $ws $r $x[0] $x[1] $x[2] 1234.5 $x[3]; $r++ }
# A1 -> A15 对应行号 3..17
Try-Set $ws 3 'A1' { $ws.Cells.Item(3, 3).Font.Name = '宋体' }
Try-Set $ws 4 'A2' { $ws.Cells.Item(4, 3).Font.Size = 20 }
Try-Set $ws 5 'A3' { $ws.Cells.Item(5, 3).Font.Bold = $true }
Try-Set $ws 6 'A4' { $ws.Cells.Item(6, 3).Font.Italic = $true }
Try-Set $ws 7 'A5-1' { $ws.Cells.Item(7, 3).Font.Underline = 2 }
Try-Set $ws 8 'A5-2' { $ws.Cells.Item(8, 3).Font.Underline = -4119 }
Try-Set $ws 9 'A5-3' { $ws.Cells.Item(9, 3).Font.Underline = 4 }
Try-Set $ws 10 'A5-4' { $ws.Cells.Item(10, 3).Font.Underline = 5 }
Try-Set $ws 11 'A6' { $ws.Cells.Item(11, 3).Font.Color = RGBv 192 0 0; $ws.Cells.Item(11, 3).Font.Underline = 2 }
Try-Set $ws 12 'A7' { $ws.Cells.Item(12, 3).Font.Strikethrough = $true }
Try-Set $ws 13 'A8' { $ws.Cells.Item(13, 3).Font.Color = RGBv 192 0 0 }
Try-Set $ws 14 'A9' { $ws.Cells.Item(14, 3).Characters(1, 4).Font.Superscript = $true }
Try-Set $ws 15 'A10' { $ws.Cells.Item(15, 3).Characters(1, 4).Font.Subscript = $true }
Try-Set $ws 17 'A13' {
    $t = $ws.Cells.Item(17, 3)
    $t.Characters(1, 4).Font.Color = RGBv 192 0 0
    $t.Characters(1, 4).Font.Bold = $true
    $t.Characters(1, 4).Font.Size = 14
    $t.Characters(5, 4).Font.Color = RGBv 0 51 204
    $t.Characters(5, 4).Font.Italic = $true
}
$ws.Rows.Item(4).RowHeight = 30

# ============================================================
# B 填充
# ============================================================
$ws = NewTab 'B填充'
Header $ws 'B 组 填充（8 项）' 5
ColHead $ws 2 @('编号', '配置项', '示例（观察底色/图案）', '数值示例', '说明 / Excel 侧限制')
BaseCols $ws 34 46

$bRows = @(
    @('B1', '纯色填充 浅黄', '背景 FFF2CC', 'RGB(255,242,204)'),
    @('B2', '图案填充 25% 灰', 'xlPatternGray25', 'xlPatternGray25'),
    @('B2b', '图案填充 棋盘', 'xlPatternChecker', '同一编号 B2 的另一图案'),
    @('B3', '图案前景色 深蓝', 'PatternColor 深蓝', '图案线颜色 = 1F3864'),
    @('B4', '图案背景色 浅灰', 'Interior.Color 作底色', '底色 F2F2F2 + 25% 灰图案'),
    @('B5', '线性渐变（角度）', '（COM 无法构造）', 'Excel COM 的 Interior 不支持渐变，需 UI 手工或直接写 OOXML'),
    @('B6', '渐变停靠点（颜色+位置）', '（COM 无法构造）', '同上；本行仅作 API 对照锚点'),
    @('B7', '渐变类型（射线/矩形/路径/从角）', '（COM 无法构造）', '同上；本行仅作 API 对照锚点'),
    @('B8', '无填充（透明）', '清除填充后应显示白色', 'Interior.Pattern = xlNone')
)
$r = 3
foreach ($x in $bRows) { Item $ws $r $x[0] $x[1] $x[2] 1234.5 $x[3]; $r++ }
Try-Set $ws 3 'B1' { $ws.Cells.Item(3, 3).Interior.Color = RGBv 255 242 204 }
Try-Set $ws 4 'B2' { $ws.Cells.Item(4, 3).Interior.Pattern = -4124 }
Try-Set $ws 5 'B2b' { $ws.Cells.Item(5, 3).Interior.Pattern = 9 }
Try-Set $ws 6 'B3' { $ws.Cells.Item(6, 3).Interior.Pattern = -4124; $ws.Cells.Item(6, 3).Interior.PatternColor = RGBv 31 56 100 }
Try-Set $ws 7 'B4' { $ws.Cells.Item(7, 3).Interior.Pattern = -4124; $ws.Cells.Item(7, 3).Interior.Color = RGBv 242 242 242 }
Try-Set $ws 10 'B8' { $ws.Cells.Item(10, 3).Interior.Pattern = -4142 }
foreach ($rr in 3..10) { $ws.Rows.Item($rr).RowHeight = 26 }

# ============================================================
# C 边框
# ============================================================
$ws = NewTab 'C边框'
Header $ws 'C 组 边框（11 项）' 10
ColHead $ws 2 @('编号', '配置项', '示例区（观察边框）', '', '', '', '', '', '', '说明')
BaseCols $ws 34 46

# C1-C3 一块 3x3 区域
$ws.Cells.Item(3, 1).Value2 = 'C1-C3'
$ws.Cells.Item(3, 1).Font.Bold = $true
$ws.Cells.Item(3, 1).Font.Name = 'Consolas'
$ws.Cells.Item(3, 2).Value2 = '外框 + 内部横竖线'
$ws.Cells.Item(3, 2).Font.Bold = $true
$ws.Cells.Item(3, 10).Value2 = '外框：中粗线深红；内部横线：细虚线灰；内部竖线：细点线灰'
$ws.Cells.Item(3, 10).WrapText = $true
$block = $ws.Range($ws.Cells.Item(3, 3), $ws.Cells.Item(5, 5))
for ($i = 1; $i -le 3; $i++) { for ($j = 3; $j -le 5; $j++) { $ws.Cells.Item(2 + $i, $j).Value2 = "R$i C$($j-2)" } }
Try-Set $ws 3 'C1' {
    foreach ($idx in @(7, 8, 9, 10)) {
        $b = $block.Borders.Item($idx)
        $b.LineStyle = 1; $b.Weight = -4138; $b.Color = RGBv 192 0 0
    }
}
Try-Set $ws 3 'C2' {
    $b = $block.Borders.Item(12)
    $b.LineStyle = -4115; $b.Weight = 2; $b.Color = RGBv 156 163 175
}
Try-Set $ws 3 'C3' {
    $b = $block.Borders.Item(11)
    $b.LineStyle = -4118; $b.Weight = 2; $b.Color = RGBv 156 163 175
}
foreach ($rr in 3..5) { $ws.Rows.Item($rr).RowHeight = 24 }

# C4 线型 13 种（横排）
$ws.Cells.Item(7, 1).Value2 = 'C4'
$ws.Cells.Item(7, 1).Font.Bold = $true
$ws.Cells.Item(7, 1).Font.Name = 'Consolas'
$ws.Cells.Item(7, 2).Value2 = '线型 13 种'
$ws.Cells.Item(7, 2).Font.Bold = $true
$ws.Cells.Item(7, 10).Value2 = 'LineStyle + Weight 组合，逐格对照'
$styles = @(
    @('细实线', 1, 2), @('中实线', 1, -4138), @('粗实线', 1, 4), @('极细实线', 1, 1),
    @('细虚线', -4115, 2), @('中虚线', -4115, -4138), @('细点线', -4118, 2), @('中粗点线', -4118, -4138),
    @('双线', -4119, -4138), @('点划线', 4, 2), @('双点划线', 5, 2), @('斜点划线', 13, 2), @('无边框', -4142, 2)
)
$ci = 2
foreach ($s in $styles) {
    $cell = $ws.Cells.Item(8, $ci)
    $cell.Value2 = $s[0]
    $cell.Font.Size = 8
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment = -4108
    $cell.Borders.LineStyle = [int]$s[1]
    if ([int]$s[1] -ne -4142) { $cell.Borders.Weight = [int]$s[2] }
    $ci++
}
$ws.Rows.Item(8).RowHeight = 40
$ws.Columns.Item(2).ColumnWidth = 16

# C5 边框颜色 / C6-C7 对角线 / C8 斜线表头 / C9 清除 / C10 合并区
$ws.Cells.Item(10, 1).Value2 = 'C5'
$ws.Cells.Item(10, 1).Font.Bold = $true
$ws.Cells.Item(10, 1).Font.Name = 'Consolas'
$ws.Cells.Item(10, 2).Value2 = '边框颜色（蓝粗框）'
$ws.Cells.Item(10, 3).Value2 = '蓝色粗外框'
$ws.Cells.Item(10, 10).Value2 = '颜色与线型独立设置'
$c5 = $ws.Range($ws.Cells.Item(10, 3), $ws.Cells.Item(11, 4))
Try-Set $ws 10 'C5' {
    $c5.Borders.LineStyle = 1
    $c5.Borders.Weight = 4
    $c5.Borders.Color = RGBv 0 51 204
}

$ws.Cells.Item(10, 5).Value2 = 'C6 对角线＼'
$ws.Cells.Item(10, 6).Value2 = 'C7 对角线／'
Try-Set $ws 10 'C6' { $ws.Cells.Item(10, 5).Borders.Item(5).LineStyle = 1; $ws.Cells.Item(10, 5).Borders.Item(5).Weight = -4138; $ws.Cells.Item(10, 5).Borders.Item(5).Color = RGBv 31 56 100 }
Try-Set $ws 10 'C7' { $ws.Cells.Item(10, 6).Borders.Item(6).LineStyle = -4118; $ws.Cells.Item(10, 6).Borders.Item(6).Weight = 2; $ws.Cells.Item(10, 6).Borders.Item(6).Color = RGBv 31 56 100 }
$ws.Cells.Item(10, 10).Value2 = $ws.Cells.Item(10, 10).Value2 + '；C6/C7 为单条对角线（左斜 / 右斜）'
$ws.Cells.Item(10, 5).HorizontalAlignment = -4108
$ws.Cells.Item(10, 6).HorizontalAlignment = -4108

$ws.Cells.Item(13, 1).Value2 = 'C8'
$ws.Cells.Item(13, 1).Font.Bold = $true
$ws.Cells.Item(13, 1).Font.Name = 'Consolas'
$ws.Cells.Item(13, 2).Value2 = '斜线表头（双斜线+分隔文本）'
$ws.Cells.Item(13, 3).Value2 = '科目＼期间'
Try-Set $ws 13 'C8' {
    $h = $ws.Cells.Item(13, 3)
    $h.Borders.Item(5).LineStyle = 1
    $h.Borders.Item(6).LineStyle = 1
    $h.HorizontalAlignment = -4108
}
$ws.Cells.Item(13, 10).Value2 = 'Excel 用两条对角线组合 + 空格分隔文本实现，无专用对象；SJS 亦无专用 API'
$ws.Rows.Item(13).RowHeight = 30

$ws.Cells.Item(15, 1).Value2 = 'C9'
$ws.Cells.Item(15, 1).Font.Bold = $true
$ws.Cells.Item(15, 1).Font.Name = 'Consolas'
$ws.Cells.Item(15, 2).Value2 = '清除边框'
$ws.Cells.Item(15, 3).Value2 = '本格应无任何框线'
Try-Set $ws 15 'C9' { $ws.Cells.Item(15, 3).Borders.LineStyle = -4142 }
$ws.Cells.Item(15, 10).Value2 = '先设框线再清除，验证清除是否彻底'
$ws.Cells.Item(15, 3).Borders.LineStyle = -4142

$ws.Cells.Item(17, 1).Value2 = 'C10'
$ws.Cells.Item(17, 1).Font.Bold = $true
$ws.Cells.Item(17, 1).Font.Name = 'Consolas'
$ws.Cells.Item(17, 2).Value2 = '边框与合并区交互'
$ws.Cells.Item(17, 3).Value2 = '合并区外框应包住整块'
Try-Set $ws 17 'C10' {
    $m = $ws.Range($ws.Cells.Item(17, 3), $ws.Cells.Item(18, 4))
    $m.Merge()
    $m.Borders.LineStyle = 1
    $m.Borders.Weight = -4138
}

$ws.Cells.Item(20, 1).Value2 = 'C11'
$ws.Cells.Item(20, 1).Font.Bold = $true
$ws.Cells.Item(20, 1).Font.Name = 'Consolas'
$ws.Cells.Item(20, 2).Value2 = '圆角边框（SJS V19.2 新增）'
$ws.Cells.Item(20, 3).Value2 = '（Excel 无对应能力）'
$ws.Cells.Item(20, 10).Value2 = 'Excel 无圆角边框对象模型，本行仅作 API 对照锚点；SJS 侧 borderRadius / setBorderRadius'
$ws.Cells.Item(20, 3).Interior.Color = RGBv 245 245 245

$ws.Columns.Item(10).ColumnWidth = 44
$ws.Columns.Item(10).WrapText = $true

# ============================================================
# D 对齐与合并（17 项）
# ============================================================
$ws = NewTab 'D对齐'
Header $ws 'D 组 对齐与合并（17 项）' 5
ColHead $ws 2 @('编号', '配置项', '示例（观察对齐/折行/旋转）', '数值示例', '说明 / Excel 侧限制')
BaseCols $ws 44 48

# ---- D1 水平对齐 8 种：中等长度文本，列宽固定，才能看出左右差异 ----
$ws.Cells.Item(3, 1).Value2 = 'D1'
$ws.Cells.Item(3, 1).Font.Name = 'Consolas'; $ws.Cells.Item(3, 1).Font.Bold = $true
$ws.Cells.Item(3, 2).Value2 = '水平对齐 8 种'
$ws.Cells.Item(3, 5).Value2 = '文本占列宽约一半，便于分辨左/中/右/填充/两端/跨列/分散'
$hAligns = @(
    @('D1-1', '常规', 1), @('D1-2', '靠左', -4131), @('D1-3', '居中', -4108), @('D1-4', '靠右', -4152),
    @('D1-5', '填充', 5), @('D1-6', '两端对齐', -4130), @('D1-7', '跨列居中', 7), @('D1-8', '分散对齐', -4117)
)
$r = 4
foreach ($a in $hAligns) {
    Item $ws $r $a[0] $a[1] '对齐示例：中文 Abc 123' $null ('Excel 常量 ' + $a[2])
    $ws.Rows.Item($r).RowHeight = 24
    $r++
}
# 逐行显式设置水平对齐（避免闭包捕获行号出错）
foreach ($i in 0..7) {
    $row = 4 + $i
    $val = [int]$hAligns[$i][2]
    try { $ws.Cells.Item($row, 3).HorizontalAlignment = $val } catch {
        $cur = [string]$ws.Cells.Item($row, 5).Value2
        $ws.Cells.Item($row, 5).Value2 = $cur + ' || COM失败：' + $_.Exception.Message
    }
}

# ---- D2 垂直对齐 5 种：行高拉高才能看出上下差异 ----
$ws.Cells.Item(13, 1).Value2 = 'D2'
$ws.Cells.Item(13, 1).Font.Name = 'Consolas'; $ws.Cells.Item(13, 1).Font.Bold = $true
$ws.Cells.Item(13, 2).Value2 = '垂直对齐 5 种'
$ws.Cells.Item(13, 5).Value2 = '行高 52，文字仅一行，靠上/居中/靠下差异明显'
$vAligns = @(
    @('D2-1', '靠上', -4160), @('D2-2', '居中', -4108), @('D2-3', '靠下', -4107),
    @('D2-4', '两端对齐', -4130), @('D2-5', '分散对齐', -4117)
)
$r = 14
foreach ($a in $vAligns) {
    Item $ws $r $a[0] $a[1] '垂直对齐示例（行高 52）' $null ('Excel 常量 ' + $a[2])
    $ws.Rows.Item($r).RowHeight = 52
    try { $ws.Cells.Item($r, 3).VerticalAlignment = [int]$a[2] } catch {
        $cur = [string]$ws.Cells.Item($r, 5).Value2
        $ws.Cells.Item($r, 5).Value2 = $cur + ' || COM失败：' + $_.Exception.Message
    }
    $ws.Cells.Item($r, 3).HorizontalAlignment = -4131
    $r++
}

# ---- D3 缩进 ----
Item $ws 19 'D3' '缩进 2 级字符' '缩进两级：文字应从左侧内缩两字符' $null 'IndentLevel = 2'
$ws.Cells.Item(19, 3).IndentLevel = 2
$ws.Rows.Item(19).RowHeight = 24

# ---- D4 / D7 / D15：长文本对照区（关键演示，列宽 44 固定）----
$longText = '自动换行与缩小字体填充必须用长文本才看得出差异：这一段文字故意写到超过列宽四倍以上，用来观察不同设置下文字是被折成多行、被整体缩小到一行、还是直接溢出到右侧空单元格。短文本这两种设置看起来完全一样，所以演示效果会很差。'

$r = 21
Item $ws $r 'D-ref' '对照：无任何设置（基线）' $longText $null '不换行不缩小：文字溢出到右侧空单元格，被右侧内容挡住则截断'
$ws.Cells.Item($r, 3).HorizontalAlignment = -4131
$ws.Cells.Item($r, 3).VerticalAlignment = -4108
$ws.Rows.Item($r).RowHeight = 24

$r++
Item $ws $r 'D4' '自动换行 wordWrap' $longText $null '按列宽折行；行高 96 约显示 5 行，可完整看出折行效果'
try {
    $ws.Cells.Item($r, 3).WrapText = $true
    $ws.Cells.Item($r, 3).HorizontalAlignment = -4131
    $ws.Cells.Item($r, 3).VerticalAlignment = -4160
    $ws.Rows.Item($r).RowHeight = 96
} catch {
    $cur = [string]$ws.Cells.Item($r, 5).Value2
    $ws.Cells.Item($r, 5).Value2 = $cur + ' || COM失败：' + $_.Exception.Message
}

$r++
Item $ws $r 'D7' '缩小字体填充 shrinkToFit' $longText $null '保持单行，自动缩小字号塞进列宽；行高不变（对照 D4 同一段文字）'
try {
    $ws.Cells.Item($r, 3).ShrinkToFit = $true
    $ws.Cells.Item($r, 3).HorizontalAlignment = -4131
    $ws.Cells.Item($r, 3).VerticalAlignment = -4108
    $ws.Rows.Item($r).RowHeight = 24
} catch {
    $cur = [string]$ws.Cells.Item($r, 5).Value2
    $ws.Cells.Item($r, 5).Value2 = $cur + ' || COM失败：' + $_.Exception.Message
}

$r++
Item $ws $r 'D15' '文本省略符 showEllipsis（SJS 独有）' $longText $null 'Excel 无 showEllipsis 对象属性，本行 Excel 侧等同 D-ref；SJS 侧应显示尾部省略号 …'
$ws.Cells.Item($r, 3).HorizontalAlignment = -4131
$ws.Cells.Item($r, 3).VerticalAlignment = -4108
$ws.Rows.Item($r).RowHeight = 24

$ws.Cells.Item($r + 1, 1).Value2 = '↑ D-ref / D4 / D7 / D15 四行文本完全相同、列宽完全相同，只有对齐属性不同，是最直接的对照组'
$ws.Range($ws.Cells.Item($r + 1, 1), $ws.Cells.Item($r + 1, 5)).Merge()
$ws.Cells.Item($r + 1, 1).Font.Bold = $true
$ws.Cells.Item($r + 1, 1).Font.Color = RGBv 0 102 204
$ws.Cells.Item($r + 1, 1).HorizontalAlignment = -4131

$r += 3
Item $ws $r 'D5-1' '旋转 45 度' '旋转45度 Abc' $null 'Orientation = 45'
try { $ws.Cells.Item($r, 3).Orientation = 45; $ws.Rows.Item($r).RowHeight = 48 } catch { }
$r++
Item $ws $r 'D5-2' '旋转 -45 度' '旋转-45度 Abc' $null 'Orientation = -45'
try { $ws.Cells.Item($r, 3).Orientation = -45; $ws.Rows.Item($r).RowHeight = 48 } catch { }
$r++
Item $ws $r 'D5-3' '旋转 90 度（竖排文字方向）' '旋转90度 Abc' $null 'Orientation = 90'
try { $ws.Cells.Item($r, 3).Orientation = 90; $ws.Rows.Item($r).RowHeight = 62 } catch { }
$r++
Item $ws $r 'D6' '竖排文字（东亚，字符直立）' '竖排文字示例' $null 'Orientation = xlVertical(-4166)，字符自上而下直立排列'
try { $ws.Cells.Item($r, 3).Orientation = -4166; $ws.Rows.Item($r).RowHeight = 80 } catch { }
$r++
Item $ws $r 'D8' '跨列居中（不合并）' '跨列居中的标题文字' $null 'HorizontalAlignment = 7（xlCenterAcrossSelection），本行 C:F 为同一跨列区'
try {
    $ws.Cells.Item($r, 3).HorizontalAlignment = 7
    $ws.Range($ws.Cells.Item($r, 4), $ws.Cells.Item($r, 6)).ClearContents() | Out-Null
    $ws.Rows.Item($r).RowHeight = 24
} catch {
    $cur = [string]$ws.Cells.Item($r, 5).Value2
    $ws.Cells.Item($r, 5).Value2 = $cur + ' || COM失败：' + $_.Exception.Message
}
$r++
Item $ws $r 'D9' '文字方向 从右到左（RTL）' 'RTL 从右到左 Abc 123' $null 'ReadingOrder = xlRTL(0 为上下文，-500 为 RTL)'
try { $ws.Cells.Item($r, 3).ReadingOrder = -500; $ws.Rows.Item($r).RowHeight = 24 } catch { }
$r++
Item $ws $r 'D10' '合并后仅保留左上值' '左上角值' $null 'C/D 两列合并；D 列预置"应被清空"，合并后应消失'
$ws.Cells.Item($r, 4).Value2 = '应被清空'
try { $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 4)).Merge(); $ws.Rows.Item($r).RowHeight = 24 } catch { }
$r++
Item $ws $r 'D11' '合并区域读取' '本行为纵向合并 3 行' $null 'C21:C23 合并，用于验证 getMergedRanges / getSpan 能否枚举'
try { $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r + 2, 3)).Merge() } catch { }
$r += 4
Item $ws $r 'D12' '取消合并（已还原为独立格）' '合并后再取消' $null '本行 C 列曾合并 2 格后取消，两侧导入后应显示为普通单元格'
try {
    $mg = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 4))
    $mg.Merge()
    $mg.UnMerge()
    $ws.Rows.Item($r).RowHeight = 24
} catch { }
$r++
Item $ws $r 'D13' '跨行/跨列合并快捷方式' '横向合并（C:E）' $null '横向整行合并；纵向见 D11'
try {
    $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 5)).Merge()
    $ws.Cells.Item($r, 3).HorizontalAlignment = -4108
    $ws.Rows.Item($r).RowHeight = 24
} catch { }
$r++
Item $ws $r 'D14' '合并+自动换行+垂直居中组合' '报表标题常见组合：合并 C:E、开启自动换行、垂直居中、行高 56，文字应上下居中且超过列宽时折行显示' $null '属性组合项，Excel 最常用的标题排版'
try {
    $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 5)).Merge()
    $m = $ws.Cells.Item($r, 3)
    $m.WrapText = $true
    $m.VerticalAlignment = -4108
    $m.HorizontalAlignment = -4108
    $ws.Rows.Item($r).RowHeight = 56
} catch {
    $cur = [string]$ws.Cells.Item($r, 5).Value2
    $ws.Cells.Item($r, 5).Value2 = $cur + ' || COM失败：' + $_.Exception.Message
}
$r++
Item $ws $r 'D16' '单元格内边距 padding（SJS 独有）' '（Excel 无 padding 属性）' $null 'Excel 只能靠缩进近似，无四边 padding；SJS 有 setPadding'
$ws.Cells.Item($r, 3).Interior.Color = RGBv 245 245 245
$r++
Item $ws $r 'D17' '自动合并（SJS 独有）' '相同内容自动合并' $null 'Excel 无对象模型能力；SJS autoMerge（V13）'
$ws.Cells.Item($r, 3).Interior.Color = RGBv 245 245 245

# ============================================================
# E 数字格式（25 项）
# ============================================================
$ws = NewTab 'E数字格式'
Header $ws 'E 组 数字格式（25 项）' 5
ColHead $ws 2 @('编号', '配置项', '原始值', '应用后的显示（本格已套格式）', '格式串 / 说明')
$ws.Columns.Item(1).ColumnWidth = 9
$ws.Columns.Item(2).ColumnWidth = 28
$ws.Columns.Item(3).ColumnWidth = 16
$ws.Columns.Item(4).ColumnWidth = 30
$ws.Columns.Item(5).ColumnWidth = 46
$ws.Columns.Item(5).WrapText = $true
$ws.Rows.Item(2).RowHeight = 22

$fmts = @(
    @('E1', '常规', 1234.5678, 'General', 'Excel 默认格式，不设置（中文环境 NumberFormat="General" 会报错）'),
    @('E2', '数值 2 位小数+千分位', 1234567.891, '#,##0.00', ''),
    @('E3', '货币（负数为括号）', -1234.5, '$#,##0.00_);($#,##0.00)', ''),
    @('E4', '会计专用', 1234.5, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', '货币符号与数字按列对齐，_ 为占位空格'),
    @('E5-1', '日期 短横线', 45678, 'yyyy-mm-dd', ''),
    @('E5-2', '日期 斜杠', 45678, 'm/d/yyyy', ''),
    @('E5-3', '日期 两位年份', 45678, 'mm-dd-yy', ''),
    @('E5-4', '日期 中文年月日+星期', 45678, 'yyyy"年"m"月"d"日" aaaa', '中文环境需要中文引号段；直接写 yyyy年m月d日 会被 Excel 归一成内置格式'),
    @('E6-1', '时间 时:分', 0.5678, 'h:mm', ''),
    @('E6-2', '时间 时:分:秒', 0.5678, 'h:mm:ss', ''),
    @('E7', '百分比 2 位', 0.256, '0.00%', ''),
    @('E7-2', '百分比 0 位', 0.256, '0%', ''),
    @('E8', '分数 最多两位', 3.14159, '# ?/??', ''),
    @('E9', '科学记数', 123456789, '0.00E+00', ''),
    @('E10', '文本格式 @', '00123', '@', '必须先设 @ 再写值，否则 00123 会被转成数字 123'),
    @('E11-1', '中文小写 DBNum1', 12345, '[DBNum1][$-804]G/通用格式', '中文环境可写；Univer 官方明确 DBNum 不支持'),
    @('E11-2', '中文大写 DBNum2', 12345, '[DBNum2][$-804]G/通用格式', '财务大写金额，POC 高频'),
    @('E12', '补零 6 位', 123, '000000', ''),
    @('E13', '# 占位符', 1234.5, '#.##', '不补零'),
    @('E14', '? 占位符（补空格对齐）', 1.5, '0.???', ''),
    @('E15', '@ 文本段', 'A-100', '"编号："@', ''),
    @('E16', '千分位缩放（按千）', 12345678, '#,##0,', '末尾逗号表示按千缩放'),
    @('E17', '多段+负数括号红字', -1234.5, '#,##0.00_);[红色](#,##0.00)', '中文环境颜色名只认 [红色]，[Red] 会直接报错'),
    @('E18', '条件段 [>100]', 250, '[>100]"高"0;[<=100]"低"0', '按值切换显示段'),
    @('E19', '颜色段 [红]/[蓝]', -12.5, '[<0][红色]0.00;[>=0][蓝色]0.00', ''),
    @('E22', '区域性代码 [$-411]', 45678, '[$-411]yyyy-mm-dd', ''),
    @('E23', '中日韩转义符 !（SJS V19.2）', 12345, '0!.0,"万"', 'SJS 独有语法；与 0\.0,"万" 等价，Excel 侧无 ! 语法'),
    @('E25', '格式串内嵌公式（SJS V17）', 12345, 'TEXT(A1,"#,##0")', 'SJS 独有；Excel 的 formatter 不能内嵌公式，本格按文本处理')
)
$r = 3
foreach ($f in $fmts) {
    $c3 = $ws.Cells.Item($r, 3)
    $c4 = $ws.Cells.Item($r, 4)
    $ws.Cells.Item($r, 1).Value2 = [string]$f[0]
    $ws.Cells.Item($r, 1).Font.Name = 'Consolas'
    $ws.Cells.Item($r, 1).Font.Bold = $true
    $ws.Cells.Item($r, 2).Value2 = [string]$f[1]
    $ws.Cells.Item($r, 5).Value2 = [string]$f[4]
    $ok = $true
    if ($f[0] -eq 'E1') {
        if ($f[2] -is [string]) { $c3.Value2 = $f[2] } else { $c3.Value2 = [double]$f[2] }
    } elseif ($f[0] -eq 'E10') {
        $c3.NumberFormat = '@'
        $c3.Value2 = [string]$f[2]
        $c4.Value2 = '@'
    } elseif ($f[0] -eq 'E25') {
        $c3.Value2 = [double]$f[2]
        $c4.Value2 = [string]$f[3]
        $ok = $false
    } else {
        if ($f[2] -is [string]) { $c3.Value2 = $f[2] } else { $c3.Value2 = [double]$f[2] }
        try {
            $c3.NumberFormat = [string]$f[3]
            $c4.Value2 = [string]$c3.NumberFormat
            if ([string]$c3.NumberFormat -ne [string]$f[3]) {
                $ws.Cells.Item($r, 5).Value2 = [string]$f[4] + ' || Excel 归一化后=' + [string]$c3.NumberFormat
            }
        } catch {
            $ok = $false
            $c4.Value2 = '设置失败：' + $f[3]
            $ws.Cells.Item($r, 5).Value2 = [string]$f[4] + ' || COM失败：' + $_.Exception.Message
        }
    }
    $c4.Font.Name = 'Consolas'
    $c4.Font.Size = 9
    $c3.HorizontalAlignment = -4152
    $ws.Rows.Item($r).RowHeight = 22
    $r++
}
$ws.Cells.Item($r + 1, 1).Value2 = '注：D 列显示的是设置后 COM 回读到的真实格式串（与 B 列意图格式串对照，可发现 Excel 的自动归一化）'
$ws.Range($ws.Cells.Item($r + 1, 1), $ws.Cells.Item($r + 1, 5)).Merge()
$ws.Cells.Item($r + 1, 1).Font.Italic = $true

# ============================================================
# F 行高列宽（8 项）
# ============================================================
$ws = NewTab 'F行高列宽'
Header $ws 'F 组 行高列宽（8 项）' 5
ColHead $ws 2 @('编号', '配置项', '示例', '数值示例', '说明 / Excel 侧限制')
BaseCols $ws 46 46

Item $ws 3 'F1' '设置行高 44 磅' '本行行高 44（相邻行默认 14.25）' $null 'setRowHeight / RowHeight，支持小数磅'
$ws.Rows.Item(3).RowHeight = 44

Item $ws 4 'F2' '设置列宽 40 字符' 'C 列宽 = 40' $null 'setColumnWidth / ColumnWidth，单位为默认字体的字符宽'
# C 列宽在 BaseCols 已设 46，这里用 E 列做示范
$ws.Cells.Item(4, 5).Value2 = 'E 列宽被设为 8（对照：C 列 46）'
$ws.Columns.Item(5).ColumnWidth = 8
$ws.Rows.Item(4).RowHeight = 30

Item $ws 6 'F3' '自动调整行高 autoFit' '自动调整行高的演示文本：这一段故意写得比较长，用于让 WrapText 折成多行，然后在行高只有 14.25 的情况下执行 AutoFit，行高应当自动增长到能容纳全部内容为止。' $null '先设 WrapText + 行高 14.25，再 AutoFit，行高应从 14.25 增长到约 60'
$ws.Cells.Item(6, 3).WrapText = $true
$ws.Rows.Item(6).RowHeight = 14.25
$ws.Rows.Item(6).AutoFit()

Item $ws 8 'F4' '自动调整列宽 autoFit' '自动调整列宽演示：这一段文字非常长，用来验证 AutoFit 后列宽是否被撑开到刚好容纳整行文本。' $null '先设列宽 8，再 AutoFit，列宽应从 8 增长到约 60~70'
$ws.Columns.Item(3).ColumnWidth = 8
$ws.Rows.Item(8).AutoFit()
$ws.Columns.Item(3).AutoFit()
# 记下自动调整后的实际值
$ws.Cells.Item(8, 5).Value2 = 'AutoFit 后：C 列宽 = ' + [math]::Round($ws.Columns.Item(3).ColumnWidth, 1) + '，本行行高 = ' + [math]::Round($ws.Rows.Item(8).RowHeight, 1)

Item $ws 10 'F5' '隐藏行' '本行下方第 11 行被隐藏' $null 'setRowVisible(false) / Row.Hidden = true'
$ws.Cells.Item(11, 1).Value2 = '（这一行本应隐藏）'
$ws.Cells.Item(11, 2).Value2 = 'F5 被隐藏的行'
$ws.Rows.Item(11).Hidden = $true

Item $ws 13 'F6' '隐藏列' 'E 列已被隐藏（本行 E 列有值但看不到）' $null 'setColumnVisible(false) / Column.Hidden = true'
$ws.Columns.Item(5).Hidden = $true

Item $ws 15 'F7' '默认行高/列宽（只读属性）' '本表默认行高/列宽见说明列' $null ('Excel 的 Worksheet.StandardHeight/StandardWidth 是只读属性：当前行高=' + [math]::Round($ws.StandardHeight, 2) + ' 磅、列宽=' + [math]::Round($ws.StandardWidth, 2) + ' 字符；改默认值只能逐行逐列设，或用 Rows.RowHeight 全表覆盖')

Item $ws 17 'F8' '行高列宽导入保真' '本工作簿即证据：各 sheet 的行高列宽均为非默认值' $null '导入后逐行比对 RowHeight / ColumnWidth'

# ============================================================
# G 条件格式 - 规则类型（16 项）
# ============================================================
$script:DescCol = 9
$ws = NewTab 'G条件格式规则'
Header $ws 'G 组 条件格式 · 规则类型（16 项）' 9
ColHead $ws 2 @('编号', '规则类型', '数据区 →（规则已应用到本行 C:H）', '', '', '', '', '', '说明 / 预期效果')
$ws.Columns.Item(1).ColumnWidth = 8
$ws.Columns.Item(2).ColumnWidth = 22
foreach ($i in 3..8) { $ws.Columns.Item($i).ColumnWidth = 9 }
$ws.Columns.Item(9).ColumnWidth = 52
$ws.Columns.Item(9).WrapText = $true
$ws.Rows.Item(2).RowHeight = 24

$red = RGBv 255 199 206
$green = RGBv 198 239 206
$yellow = RGBv 255 235 156
$blue = RGBv 189 215 238

function CFRow($ws, $r, $code, $name, $values, $desc) {
    $ws.Cells.Item($r, 1).Value2 = [string]$code
    $ws.Cells.Item($r, 1).Font.Name = 'Consolas'
    $ws.Cells.Item($r, 1).Font.Bold = $true
    $ws.Cells.Item($r, 2).Value2 = [string]$name
    $i = 3
    foreach ($v in $values) {
        if ($null -ne $v) {
            $cell = $ws.Cells.Item($r, $i)
            if ($v -is [string]) { $cell.Value2 = [string]$v } else { $cell.Value2 = [double]$v }
        }
        $ws.Cells.Item($r, $i).HorizontalAlignment = -4108
        $i++
    }
    $ws.Cells.Item($r, 9).Value2 = [string]$desc
    $ws.Rows.Item($r).RowHeight = 26
}

$nums = @(10, 25, 40, 55, 70, 85)
$t0 = [double](Get-Date).Date.ToOADate()
$r = 3

CFRow $ws $r 'G1' '单元格值：大于 50' $nums '55/70/85 应被标红'
Try-Set $ws $r 'G1' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(1, 5, '50')
    $fc.Interior.Color = $red
}
$r++

CFRow $ws $r 'G2' '特定文本：包含"合格"' @('合格', '优秀', '良好', '不合格', '合格', '待定') '含"合格"的两格标绿（"不合格"也算包含）｜此规则类型无法用 Excel COM 创建，已由 OOXML 补丁写入，Excel 中显示正常'
Try-SetQuiet {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(9, 0, '合格')
    $fc.Interior.Color = $green
}
$r++

CFRow $ws $r 'G3' '日期发生：今天' @(($t0 - 3), ($t0 - 1), $t0, $t0, ($t0 + 1), ($t0 + 5)) '今天的两格标黄（其余为日期值，已设日期格式）｜timePeriod 规则由 OOXML 补丁写入'
Try-SetQuiet {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $rng.NumberFormat = 'yyyy-mm-dd'
    $fc = $rng.FormatConditions.Add(11, 2, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value)
    $fc.Interior.Color = $yellow
}
$r++

CFRow $ws $r 'G4' '空值 / 非空' @(1, $null, 3, $null, 5, 6) '空的两格（D/F）应被标色'
Try-Set $ws $r 'G4' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(10)
    $fc.Interior.Color = $blue
}
$r++

CFRow $ws $r 'G5' '公式规则：偶数标色' $nums '10/40/70 为偶数应标色（公式 =MOD(C3,2)=0）'
Try-Set $ws $r 'G5' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(2, [System.Reflection.Missing]::Value, ('=MOD(' + $ws.Cells.Item($r, 3).Address($false, $false) + ',2)=0'))
    $fc.Interior.Color = $blue
}
$r++

CFRow $ws $r 'G6' '前 3 项' $nums '70/85 之外再取 55，共 3 格标色'
Try-Set $ws $r 'G6' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(5)
    $fc.TopBottom = 1
    $fc.Rank = 3
    $fc.Percent = $false
    $fc.Interior.Color = $green
}
$r++

CFRow $ws $r 'G7' '前 30%' $nums '6 格中前 30% ≈ 前 2 项'
Try-Set $ws $r 'G7' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(5)
    $fc.TopBottom = 1
    $fc.Rank = 30
    $fc.Percent = $true
    $fc.Interior.Color = $green
}
$r++

CFRow $ws $r 'G8' '后 2 项' $nums '10/25 应标色'
Try-Set $ws $r 'G8' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(5)
    $fc.TopBottom = 0
    $fc.Rank = 2
    $fc.Percent = $false
    $fc.Interior.Color = $yellow
}
$r++

CFRow $ws $r 'G9' '高于平均值' $nums '平均值 47.5，55/70/85 应标色'
Try-Set $ws $r 'G9' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(12)
    $fc.AboveBelow = 0
    $fc.Interior.Color = $red
}
$r++

CFRow $ws $r 'G10' '重复值' @(1, 2, 2, 3, 3, 3) '2 和 3 全部标色'
Try-Set $ws $r 'G10' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(8)
    $fc.DupeUnique = 1
    $fc.Interior.Color = $red
}
$r++

CFRow $ws $r 'G11' '唯一值' @(1, 2, 2, 3, 3, 3) '只有 1 被标色'
Try-Set $ws $r 'G11' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(8)
    $fc.DupeUnique = 0
    $fc.Interior.Color = $green
}
$r++

CFRow $ws $r 'G12' '数据条' $nums '默认蓝渐变数据条'
Try-Set $ws $r 'G12' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.BarColor.Color = RGBv 99 163 255
}
$r++

CFRow $ws $r 'G13' '色阶 2 色' $nums '红→绿双色渐变'
Try-Set $ws $r 'G13' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddColorScale(2)
    $fc.ColorScaleCriteria(1).Type = 1
    $fc.ColorScaleCriteria(1).FormatColor.Color = RGBv 248 105 107
    $fc.ColorScaleCriteria(2).Type = 2
    $fc.ColorScaleCriteria(2).FormatColor.Color = RGBv 99 190 123
}
$r++

CFRow $ws $r 'G14' '色阶 3 色' $nums '红→黄→绿三段'
Try-Set $ws $r 'G14' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddColorScale(3)
    $fc.ColorScaleCriteria(1).Type = 1
    $fc.ColorScaleCriteria(1).FormatColor.Color = RGBv 248 105 107
    $fc.ColorScaleCriteria(2).Type = 3
    $fc.ColorScaleCriteria(2).Value = 50
    $fc.ColorScaleCriteria(2).FormatColor.Color = RGBv 255 235 132
    $fc.ColorScaleCriteria(3).Type = 2
    $fc.ColorScaleCriteria(3).FormatColor.Color = RGBv 99 190 123
}
$r++

CFRow $ws $r 'G15' '图标集（三色交通灯）' $nums '每格左侧显示红/黄/绿圆灯'
Try-Set $ws $r 'G15' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddIconSetCondition()
    $fc.IconSet = $wb.IconSets.Item(4)
    $fc.ShowIconOnly = $false
}

$r++
CFRow $ws $r 'G16' '迷你图条件格式（SJS 独有）' @(1, 3, 2, 5, 4, 6) 'Excel 无"条件格式内嵌迷你图"对象模型，本行仅作 API 对照锚点'
$ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8)).Interior.Color = RGBv 245 245 245

# ============================================================
# H 条件格式 - 图形参数（15 项）
# ============================================================
$ws = NewTab 'H条件格式参数'
Header $ws 'H 组 条件格式 · 图形参数（15 项）' 9
ColHead $ws 2 @('编号', '参数项', '数据区 →（规则已应用到本行 C:H）', '', '', '', '', '', '说明 / 预期效果')
$ws.Columns.Item(1).ColumnWidth = 8
$ws.Columns.Item(2).ColumnWidth = 26
foreach ($i in 3..8) { $ws.Columns.Item($i).ColumnWidth = 10 }
$ws.Columns.Item(9).ColumnWidth = 50
$ws.Columns.Item(9).WrapText = $true
$ws.Rows.Item(2).RowHeight = 24

$r = 3
CFRow $ws $r 'H1' '数据条 最小值类型=数字 0' $nums '条长以 0 为起点｜min 类型由 OOXML 补丁写入（COM 不可设）'
Try-SetQuiet {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.MinPoint.Type = 4
    $fc.MinPoint.Value = 0
    $fc.BarColor.Color = RGBv 99 163 255
}
$r++

CFRow $ws $r 'H2' '数据条 最大值类型=数字 100' $nums '条长以 100 为满值｜max 类型由 OOXML 补丁写入（COM 不可设）'
Try-SetQuiet {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.MaxPoint.Type = 4
    $fc.MaxPoint.Value = 100
    $fc.BarColor.Color = RGBv 99 163 255
}
$r++

CFRow $ws $r 'H3' '数据条 正条颜色=橙' $nums '条为橙色'
Try-Set $ws $r 'H3' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.BarColor.Color = RGBv 237 125 49
}
$r++

CFRow $ws $r 'H4' '数据条 负条颜色+负值轴' @(-60, -20, 10, 30, 50, 80) '负值显示红色反向条，轴线在 0 处｜负条同色标志与轴位置由 OOXML 补丁写入'
Try-SetQuiet {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.AxisPosition = 1
    $fc.NegativeBarFormat.ColorType = 1
    $fc.NegativeBarFormat.Color.Color = RGBv 255 0 0
}
$r++

CFRow $ws $r 'H5-1' '数据条 渐变填充（默认）' $nums '条内部有渐变'
Try-Set $ws $r 'H5-1' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.BarFillType = 0
    $fc.BarColor.Color = RGBv 99 163 255
}
$r++

CFRow $ws $r 'H5-2' '数据条 实心填充' $nums '与 H5-1 对照，条内部纯色'
Try-Set $ws $r 'H5-2' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.BarFillType = 1
    $fc.BarColor.Color = RGBv 99 163 255
}
$r++

CFRow $ws $r 'H6-1' '数据条 方向 从左到右' $nums '条从单元格左侧生长｜direction 由 OOXML 补丁写入（COM 不可设）'
Try-SetQuiet {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.Direction = 1
    $fc.BarColor.Color = RGBv 99 163 255
}
$r++

CFRow $ws $r 'H6-2' '数据条 方向 从右到左' $nums '条从单元格右侧生长｜direction 由 OOXML 补丁写入（COM 不可设）'
Try-SetQuiet {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.Direction = 2
    $fc.BarColor.Color = RGBv 99 163 255
}
$r++

CFRow $ws $r 'H7' '数据条 边框（黑色实线）' $nums '条外沿有黑色边线'
Try-Set $ws $r 'H7' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.BarBorder.Type = 1
    $fc.BarBorder.Color.Color = RGBv 0 0 0
}
$r++

CFRow $ws $r 'H8' '数据条 仅显示条（隐藏数值）' $nums '数值不显示，只剩色条'
Try-Set $ws $r 'H8' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddDatabar()
    $fc.ShowValue = $false
    $fc.BarColor.Color = RGBv 99 163 255
}
$r++

CFRow $ws $r 'H9' '色阶 最小/最大值类型与颜色' $nums 'min=数字 0 红，max=数字 100 绿'
Try-Set $ws $r 'H9' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddColorScale(2)
    $fc.ColorScaleCriteria(1).Type = 4
    $fc.ColorScaleCriteria(1).Value = 0
    $fc.ColorScaleCriteria(1).FormatColor.Color = RGBv 248 105 107
    $fc.ColorScaleCriteria(2).Type = 4
    $fc.ColorScaleCriteria(2).Value = 100
    $fc.ColorScaleCriteria(2).FormatColor.Color = RGBv 99 190 123
}
$r++

CFRow $ws $r 'H10' '色阶 中间值（50% 黄）' $nums '三段色阶，中点为 50%'
Try-Set $ws $r 'H10' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddColorScale(3)
    $fc.ColorScaleCriteria(1).Type = 1
    $fc.ColorScaleCriteria(1).FormatColor.Color = RGBv 248 105 107
    $fc.ColorScaleCriteria(2).Type = 3
    $fc.ColorScaleCriteria(2).Value = 50
    $fc.ColorScaleCriteria(2).FormatColor.Color = RGBv 255 235 132
    $fc.ColorScaleCriteria(3).Type = 2
    $fc.ColorScaleCriteria(3).FormatColor.Color = RGBv 99 190 123
}
$r++

CFRow $ws $r 'H11' '图标集类型：4 级箭头' $nums '每格显示 4 级方向箭头'
Try-Set $ws $r 'H11' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddIconSetCondition()
    $fc.IconSet = $wb.IconSets.Item(9)
}
$r++

CFRow $ws $r 'H12' '图标集阈值（数字 30/60/90）' $nums '按 30/60/90 分级，>= 为默认比较符'
Try-Set $ws $r 'H12' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddIconSetCondition()
    $fc.IconSet = $wb.IconSets.Item(4)
    $fc.IconCriteria(2).Type = 4
    $fc.IconCriteria(2).Value = 30
    $fc.IconCriteria(2).Operator = 5
    $fc.IconCriteria(3).Type = 4
    $fc.IconCriteria(3).Value = 60
    $fc.IconCriteria(3).Operator = 5
}
$r++

CFRow $ws $r 'H13' '图标集 反转顺序' $nums '与默认顺序相反（数值越大图标越"差"）'
Try-Set $ws $r 'H13' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddIconSetCondition()
    $fc.IconSet = $wb.IconSets.Item(4)
    $fc.ReverseOrder = $true
}
$r++

CFRow $ws $r 'H14' '图标集 仅显示图标' $nums '数值被隐藏，只留图标'
Try-Set $ws $r 'H14' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.AddIconSetCondition()
    $fc.IconSet = $wb.IconSets.Item(4)
    $fc.ShowIconOnly = $true
}
$r++

CFRow $ws $r 'H15' '综合叠加（图标集 + 数据条无法共存，此块为多规则叠加）' $nums '同一区域两条规则同时存在，观察渲染优先级'
Try-Set $ws $r 'H15' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc1 = $rng.FormatConditions.AddDatabar()
    $fc1.BarColor.Color = RGBv 99 163 255
    $fc2 = $rng.FormatConditions.Add(1, 5, '70')
    $fc2.Interior.Color = $red
}
$ws.Cells.Item($r + 1, 1).Value2 = '注：Excel 同一区域的数据条与图标集互斥（后加者替换），本块故意叠加"数据条 + 单元格值规则"验证求值顺序'
$ws.Range($ws.Cells.Item($r + 1, 1), $ws.Cells.Item($r + 1, 9)).Merge()
$ws.Cells.Item($r + 1, 1).Font.Italic = $true

# ============================================================
# I 条件格式 - 管理（6 项）
# ============================================================
$ws = NewTab 'I条件格式管理'
Header $ws 'I 组 条件格式 · 管理（6 项）' 9
ColHead $ws 2 @('编号', '参数项', '数据区 →（规则已应用）', '', '', '', '', '', '说明 / 预期效果')
$ws.Columns.Item(1).ColumnWidth = 8
$ws.Columns.Item(2).ColumnWidth = 28
foreach ($i in 3..8) { $ws.Columns.Item($i).ColumnWidth = 10 }
$ws.Columns.Item(9).ColumnWidth = 52
$ws.Columns.Item(9).WrapText = $true
$ws.Rows.Item(2).RowHeight = 24

$r = 3
CFRow $ws $r 'I1-1' '优先级：先加">30 红"再加">60 绿"' $nums '45 显示红；65 显示绿（后加规则优先级更高，默认胜出）'
Try-Set $ws $r 'I1-1' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc1 = $rng.FormatConditions.Add(1, 5, '30')
    $fc1.Interior.Color = $red
    $fc2 = $rng.FormatConditions.Add(1, 5, '60')
    $fc2.Interior.Color = $green
}
$r++

CFRow $ws $r 'I1-2' '优先级调整：把">30 红"设为最高' $nums '同一组数据，改为 65 也显示红（红规则被 SetFirstPriority 提到最前）'
Try-Set $ws $r 'I1-2' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc1 = $rng.FormatConditions.Add(1, 5, '30')
    $fc1.Interior.Color = $red
    $fc2 = $rng.FormatConditions.Add(1, 5, '60')
    $fc2.Interior.Color = $green
    $fc1.SetFirstPriority()
}
$r++

CFRow $ws $r 'I2' 'stopIfTrue：即使为真则停止后续' $nums '第一条（>30 红）开启 StopIfTrue，65 不再被绿色覆盖'
Try-Set $ws $r 'I2' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc1 = $rng.FormatConditions.Add(1, 5, '30')
    $fc1.Interior.Color = $red
    $fc1.StopIfTrue = $true
    $fc1.SetFirstPriority()
    $fc2 = $rng.FormatConditions.Add(1, 5, '60')
    $fc2.Interior.Color = $green
}
$r++

CFRow $ws $r 'I3' '查询某单元格命中的规则' $nums '无视觉差异；用 FormatConditions / getRules 枚举，本行为查询对象'
$ws.Cells.Item($r, 9).Value2 = '无视觉差异；用 FormatConditions 集合枚举，或 SJS getRules(row,col) / Univer getConditionalFormattingRules'
$r++

CFRow $ws $r 'I4' '修改已有规则（颜色红→紫）' $nums '本条规则创建时为红，随后被修改为紫色，应显示紫色'
Try-Set $ws $r 'I4' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(1, 5, '50')
    $fc.Interior.Color = $red
    $fc.Interior.Color = RGBv 178 140 255
}
$r++

CFRow $ws $r 'I5-1' '删除单条规则' $nums '规则已删除，本行无任何高亮'
Try-Set $ws $r 'I5-1' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(1, 5, '50')
    $fc.Interior.Color = $red
    $fc.Delete()
}
$r++

CFRow $ws $r 'I5-2' '清空区域全部规则' $nums '与上一行同样先加规则再清空，本行无高亮'
Try-Set $ws $r 'I5-2' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc = $rng.FormatConditions.Add(1, 5, '50')
    $fc.Interior.Color = $red
    $rng.FormatConditions.Delete()
}
$r++

CFRow $ws $r 'I6' '多规则叠加求值顺序（3 条）' @(5, 20, 45, 60, 75, 95) '三条规则叠加：>30 黄、>50 蓝、>80 绿，观察最终渲染哪一条'
Try-Set $ws $r 'I6' {
    $rng = $ws.Range($ws.Cells.Item($r, 3), $ws.Cells.Item($r, 8))
    $fc1 = $rng.FormatConditions.Add(1, 5, '30')
    $fc1.Interior.Color = $yellow
    $fc2 = $rng.FormatConditions.Add(1, 5, '50')
    $fc2.Interior.Color = $blue
    $fc3 = $rng.FormatConditions.Add(1, 5, '80')
    $fc3.Interior.Color = $green
}
$ws.Cells.Item($r + 1, 1).Value2 = '注：Excel 默认"后添加的规则优先级更高"，且同一区域多个填充类规则只显示优先级最高的一条'
$ws.Range($ws.Cells.Item($r + 1, 1), $ws.Cells.Item($r + 1, 9)).Merge()
$ws.Cells.Item($r + 1, 1).Font.Italic = $true

# ============================================================
# J 样式复用与主题（11 项）
# ============================================================
$script:DescCol = 5
$ws = NewTab 'J样式复用'
Header $ws 'J 组 样式复用与主题（11 项）' 5
ColHead $ws 2 @('编号', '配置项', '示例', '数值示例', '说明 / Excel 侧限制')
BaseCols $ws 40 48

# J1 命名样式
Item $ws 3 'J1' '命名样式（报表标题）' '本行套用命名样式"报表标题"' $null 'wb.Styles.Add("报表标题") 后由 Range.Style 应用；样式库中可见'
Try-Set $ws 3 'J1' {
    $st = $wb.Styles.Add('报表标题')
    $st.Font.Name = '微软雅黑'
    $st.Font.Size = 14
    $st.Font.Bold = $true
    $st.Font.Color = RGBv 255 255 255
    $st.Interior.Color = RGBv 31 41 55
    $st.HorizontalAlignment = -4108
    $ws.Cells.Item(3, 3).Style = '报表标题'
    $ws.Rows.Item(3).RowHeight = 30
}

# J2 样式实例复用
Item $ws 5 'J2' '样式实例复用（同一命名样式套两块区域）' '第一块区域' $null '下面第 6 行是第二块，两块共用同一命名样式对象'
Try-Set $ws 5 'J2' {
    $ws.Cells.Item(5, 3).Style = '报表标题'
    $ws.Cells.Item(6, 3).Value2 = '第二块区域（同一样式）'
    $ws.Cells.Item(6, 3).Style = '报表标题'
    $ws.Rows.Item(5).RowHeight = 26
    $ws.Rows.Item(6).RowHeight = 26
}

# J3 只应用部分属性
Item $ws 8 'J3' '只应用部分属性（apply* 标志位）' '（Excel 无此概念，SDK 特性）' $null 'Excel 只能整体套 Style；SJS 有 applyFont/applyFill/applyBorder 等标志位'
$ws.Cells.Item(8, 3).Interior.Color = RGBv 245 245 245

# J4 主题色 + tint
$ws.Cells.Item(10, 1).Value2 = 'J4'
$ws.Cells.Item(10, 1).Font.Name = 'Consolas'; $ws.Cells.Item(10, 1).Font.Bold = $true
$ws.Cells.Item(10, 2).Value2 = '主题色 Accent1 + 亮度 -25%'
$ws.Cells.Item(10, 3).Value2 = '主题色 Accent1，TintAndShade -0.25'
$ws.Cells.Item(10, 5).Value2 = '随主题换色；SJS 用 "Accent 1 -60" 字符串，Univer 仅 rgb'
Try-Set $ws 10 'J4' {
    $ws.Cells.Item(10, 3).Interior.ThemeColor = 5
    $ws.Cells.Item(10, 3).Interior.TintAndShade = -0.25
    $ws.Cells.Item(10, 3).Font.Color = RGBv 255 255 255
}

# J5 主题字体
$ws.Cells.Item(12, 1).Value2 = 'J5'
$ws.Cells.Item(12, 1).Font.Name = 'Consolas'; $ws.Cells.Item(12, 1).Font.Bold = $true
$ws.Cells.Item(12, 2).Value2 = '主题字体（正文 中文）'
$ws.Cells.Item(12, 3).Value2 = '主题正文中文字体 +mn-lt'
$ws.Cells.Item(12, 5).Value2 = 'Font.Name 设为 "+mn-lt" 表示正文主题中文字体；随主题切换'
Try-Set $ws 12 'J5' { $ws.Cells.Item(12, 3).Font.Name = '+mn-lt' }

# J6 标准色
$ws.Cells.Item(14, 1).Value2 = 'J6'
$ws.Cells.Item(14, 1).Font.Name = 'Consolas'; $ws.Cells.Item(14, 1).Font.Bold = $true
$ws.Cells.Item(14, 2).Value2 = '主题色板 Accent2/3/4/5/6'
for ($k = 0; $k -lt 5; $k++) {
    $c = $ws.Cells.Item(14, 3 + $k)
    $c.Value2 = 'Accent' + ($k + 2)
    $c.HorizontalAlignment = -4108
    $c.VerticalAlignment = -4108
    try {
        $c.Interior.ThemeColor = 6 + $k
        $c.Font.Color = RGBv 255 255 255
    } catch { }
}
$ws.Cells.Item(14, 5).Value2 = 'Excel 主题色 = Accent1~6 等 12 组；SJS 有 ColorScheme，Univer 仅枚举'
$ws.Rows.Item(14).RowHeight = 26

# J7 复制格式（格式刷）
Item $ws 16 'J7' '复制格式（格式刷的编程等价）' '本格套用下一行的格式（深蓝底白字）' $null '源格在 C17；用 Copy + PasteSpecial(xlPasteFormats)'
Try-Set $ws 16 'J7' {
    $src = $ws.Cells.Item(17, 3)
    $src.Value2 = '格式源（深蓝底白字）'
    $src.Interior.Color = RGBv 31 56 100
    $src.Font.Color = RGBv 255 255 255
    $src.Font.Bold = $true
    $src.Copy() | Out-Null
    $ws.Cells.Item(16, 3).PasteSpecial(-4122) | Out-Null
    $ws.Rows.Item(16).RowHeight = 24
    $ws.Rows.Item(17).RowHeight = 24
}

# J8 导入保真
Item $ws 19 'J8' '样式导入保真与 styleId 复用' '本工作簿即证据：多 sheet 共享同一批样式' $null '解包 xlsx 的 styles.xml，看同一格式是否共用 styleId（s 属性重复率）'
$ws.Cells.Item(19, 3).Interior.Color = RGBv 255 242 204
$ws.Cells.Item(19, 3).Font.Bold = $true

# J9/J10/J11 SJS 独有
Item $ws 21 'J9' '单元格折角（SJS V16 独有）' '（Excel 无折角对象模型）' $null 'SJS 有折角大小/颜色/位置；本行仅作 API 对照锚点'
$ws.Cells.Item(21, 3).Interior.Color = RGBv 245 245 245
Item $ws 23 'J10' '单元格图标（SJS V16 独有）' '（Excel 无单元格内装饰图标）' $null 'SJS 有单元格内/外图标；本行仅作 API 对照锚点'
$ws.Cells.Item(23, 3).Interior.Color = RGBv 245 245 245
Item $ws 25 'J11' '单元格状态样式（SJS V13 独有）' '（Excel 无 readOnly/disabled 状态样式）' $null 'SJS 有只读/禁用/必填等状态样式；本行仅作 API 对照锚点'
$ws.Cells.Item(25, 3).Interior.Color = RGBv 245 245 245

# ============================================================
# K 保护与视图（6 项）
# ============================================================
$ws = NewTab 'K保护与视图'
Header $ws 'K 组 保护与视图（6 项）' 5
ColHead $ws 2 @('编号', '配置项', '示例', '数值示例', '说明 / Excel 侧限制')
BaseCols $ws 40 48

Item $ws 3 'K1' '单元格锁定（Locked）' '本格为锁定状态（需开启工作表保护才生效）' $null 'Range.Locked = true；保护后不可编辑'
$ws.Cells.Item(3, 3).Locked = $true
$ws.Cells.Item(3, 3).Interior.Color = RGBv 255 242 204

Item $ws 5 'K2' '公式隐藏（FormulaHidden）' '本格 =SUM(D6:D8)，保护后公式栏不可见' $null 'Range.FormulaHidden = true；注意 SJS 的属性名是 Style.hidden（语义相反）'
$ws.Cells.Item(5, 3).Formula = '=SUM(D6:D8)'
$ws.Cells.Item(5, 3).FormulaHidden = $true
$ws.Cells.Item(5, 3).Interior.Color = RGBv 226 240 217

Item $ws 7 'K3' '工作表保护（允许选择锁定格/排序/筛选）' '本工作表已保护，密码 1234；C3 不可编辑，C6:C8 可编辑' $null 'Protect(Password, ..., AllowFiltering, AllowSorting...)；保护在脚本最后一步执行'
$ws.Cells.Item(7, 3).Value2 = '此区域为未锁定（可编辑）'
$ws.Cells.Item(7, 3).Locked = $false
$ws.Cells.Item(8, 3).Value2 = '同上，未锁定'
$ws.Cells.Item(8, 3).Locked = $false
foreach ($rr in 6..8) { $ws.Cells.Item($rr, 4).Value2 = [double]($rr * 10) }

Item $ws 11 'K4' '网格线显示（本表已关闭）' '本工作表不显示网格线' $null 'Excel 网格线为 sheetView 级；本表 Activate 后关闭 DisplayGridlines'
Item $ws 13 'K5' '零值显示（本表已关闭）' '本行 D 列写了 0，导入后应显示为空白' 0 'DisplayZeros = false 时 0 显示为空白；本格写了 0，导入后应显示为空'
$ws.Cells.Item(13, 4).Value2 = 0
Item $ws 15 'K6' '冻结窗格（冻结前 3 行 + 前 2 列）' '滚动时第 1~3 行与 A~B 列固定' $null 'ActiveWindow.SplitRow / SplitColumn + FreezePanes = true'

# ---- 视图类设置：需要激活本表 ----
$ws.Activate()
Try-Set $ws 11 'K4' { $excel.ActiveWindow.DisplayGridlines = $false }
Try-Set $ws 13 'K5' { $excel.ActiveWindow.DisplayZeros = $false }
Try-Set $ws 15 'K6' {
    $excel.ActiveWindow.SplitRow = 3
    $excel.ActiveWindow.SplitColumn = 2
    $excel.ActiveWindow.FreezePanes = $true
}

# ---- 保护（最后一步：保护后本表不能再改）----
Try-Set $ws 7 'K3' {
    $ws.Protect('1234', [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, $true, $true, $true)
}

# ============================================================
# 收尾：回到索引页并保存
# ============================================================
$wb.Sheets.Item(1).Activate()
if (Test-Path $outFile) { Remove-Item $outFile -Force }
$wb.SaveAs($outFile, 51)
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output ('DONE -> ' + $outFile)
Write-Output ('Sheets: ' + $wb.Sheets.Count)

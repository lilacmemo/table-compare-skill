# 用 Excel 自身把 G/H/I 三张条件格式表导出成 PDF（作为“Excel 基线”视觉参照）
$ErrorActionPreference = 'Stop'
$src = 'D:\workbuddy\spreadjs vs univer\测试工程\samples\格式与条件格式-全量用例-A至K.xlsx'
$out = 'D:\workbuddy\spreadjs vs univer\测试工程\tools\out\excel-CF-baseline.pdf'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($src, $false, $true)
  $names = @('G条件格式规则', 'H条件格式参数', 'I条件格式管理')
  $sheets = @()
  foreach ($n in $names) { $sheets += $wb.Worksheets.Item($n) }
  # 只选中三张表导出
  $wb.Worksheets.Item('G条件格式规则').Select()
  $wb.Worksheets.Item('H条件格式参数').Select($false)
  $wb.Worksheets.Item('I条件格式管理').Select($false)
  $xl.ActiveSheet.Parent.ActiveWindow.SelectedSheets.ExportAsFixedFormat(0, $out)
  $wb.Close($false)
  Write-Output "OK $out"
} finally {
  $xl.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

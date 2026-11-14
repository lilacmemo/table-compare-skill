# 把「Excel原生图表效果.xlsx」里 63 张图逐个导出 PNG（类型对比的 Excel 基准图）
#   两个坑：
#   1) Chart.Export 前必须 Activate()，否则返回 True 但写出 0 字节（上一轮 87/91 空文件就是这么来的）
#   2) 文件名用 {sheet}__{序号}.png，序号与 drawing anchor 顺序对齐，才能和 manifest 的 idx 对上
param(
  [string]$Xlsx = "D:\workbuddy\spreadjs vs univer\对比工程\samples\Excel原生图表效果.xlsx",
  [string]$OutDir = "D:\workbuddy\spreadjs vs univer\对比工程\out\shots\type-excel"
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

try {
  $wb = $excel.Workbooks.Open($Xlsx, $false, $true)   # ReadOnly
  $result = @()
  foreach ($ws in $wb.Worksheets) {
    $n = $ws.ChartObjects().Count
    if ($n -eq 0) { continue }
    $ws.Activate()
    for ($i = 1; $i -le $n; $i++) {
      $co = $ws.ChartObjects($i)
      $co.Activate()
      $file = Join-Path $OutDir ("{0}__{1:d2}.png" -f $ws.Name, ($i - 1))
      $ok = $co.Chart.Export($file, "PNG")
      $title = ""
      try { $title = $co.Chart.ChartTitle.Text } catch { $title = "" }
      $sz = 0
      if (Test-Path $file) { $sz = (Get-Item $file).Length }
      $result += [pscustomobject]@{
        sheet = $ws.Name; index = $i - 1; title = $title
        file = $file; bytes = $sz; ok = $ok
        width = [math]::Round($co.Width, 2); height = [math]::Round($co.Height, 2)
        chartType = $co.Chart.ChartType
      }
    }
  }
  $result | ConvertTo-Json -Depth 4 | Out-File -FilePath (Join-Path $OutDir "type-charts-index.json") -Encoding utf8
  $empty = ($result | Where-Object { $_.bytes -lt 1000 }).Count
  Write-Output ("exported {0} charts, empty {1}" -f $result.Count, $empty)
}
finally {
  if ($wb) { $wb.Close($false) | Out-Null }
  $excel.Quit() | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

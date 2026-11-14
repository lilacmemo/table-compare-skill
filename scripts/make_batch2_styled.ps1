# build_batch2.ps1 —— 批次2+3样本簿（I 系列样式 + J 高级元素，21 项）
$ErrorActionPreference = 'Continue'
$outDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsx = Join-Path $outDir '批次2系列与高级元素样本.xlsx'
$pngDir = Join-Path $outDir 'batch2_baseline'
New-Item -ItemType Directory -Force -Path $pngDir | Out-Null
if (Test-Path $xlsx) { Remove-Item $xlsx -Force }

function RGB($r, $g, $b) { return [int]($r + $g * 256 + $b * 65536) }

function Set-Range($ws, $addr, $rows) {
    $r = $rows.Count; $c = $rows[0].Count
    $arr = New-Object 'object[,]' $r, $c
    for ($i = 0; $i -lt $r; $i++) { for ($j = 0; $j -lt $c; $j++) { $arr[$i, $j] = $rows[$i][$j] } }
    $ws.Range($addr).Value2 = $arr
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()

function New-Sheet($name) {
    $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $ws.Name = $name
    return $ws
}
function Add-Chart($ws, $type, $left, $top, $dataRange) {
    $co = $ws.Shapes.AddChart2(-1, $type, $left, $top, 520, 320)
    $ch = $co.Chart
    $ch.SetSourceData($dataRange, 2)   # xlColumns
    return $ch
}
function Save-Png($ch, $tag) { try { $ch.Export((Join-Path $pngDir $tag), 'PNG') | Out-Null } catch { Write-Host "PNG FAIL $tag" } }

# ============ 08 系列样式 ============
$ws = New-Sheet('08系列样式')
Set-Range $ws 'A1:B6' @(@('月份','销售'), @('一月',120), @('二月',180), @('三月',150), @('四月',220), @('五月',260))
Set-Range $ws 'D1:E6' @(@('月份','增长'), @('一月',30), @('二月',55), @('三月',48), @('四月',80), @('五月',95))
Set-Range $ws 'G1:H6' @(@('X','Y'), @(1,25), @(2,45), @(3,60), @(4,90), @(5,130))
Set-Range $ws 'J1:K6' @(@('类目','值'), @('A',50), @('B',80), @('C',60), @('D',95), @('E',70))
Set-Range $ws 'A9:B12' @(@('类目','份额'), @('北',0.35), @('上',0.28), @('广',0.22), @('深',0.15))
Set-Range $ws 'D9:E12' @(@('类目','值'), @('A',120), @('B',180), @('C',150), @('D',220))
Set-Range $ws 'G9:H12' @(@('类目','值'), @('A',45), @('B',70), @('C',60), @('D',85))
Set-Range $ws 'J9:K12' @(@('类目','值'), @('A',110), @('B',170), @('C',140), @('D',210))

# 08A 系列填充红 + 边框黑2磅
$ch = Add-Chart $ws 51 20 260 $ws.Range('A1:B6')
try {
    $s = $ch.SeriesCollection(1)
    $s.Format.Fill.ForeColor.RGB = RGB 200 40 40
    $s.Format.Line.ForeColor.RGB = RGB 20 20 20
    $s.Format.Line.Weight = 2
    Write-Host 'OK 08A 系列填充+边框'
} catch { Write-Host "FAIL 08A: $_" }

# 08B 系列透明度 50%
$ch = Add-Chart $ws 51 560 260 $ws.Range('A1:B6')
try {
    $ch.SeriesCollection(1).Format.Fill.Transparency = 0.5
    Write-Host 'OK 08B 系列透明度'
} catch { Write-Host "FAIL 08B: $_" }

# 08C 平滑折线
$ch = Add-Chart $ws 4 1080 260 $ws.Range('D1:E6')
try {
    $ch.SeriesCollection(1).Smooth = $true
    Write-Host 'OK 08C 平滑线'
} catch { Write-Host "FAIL 08C: $_" }

# 08D 折线虚线3磅蓝
$ch = Add-Chart $ws 4 20 620 $ws.Range('D1:E6')
try {
    $s = $ch.SeriesCollection(1)
    $s.Format.Line.ForeColor.RGB = RGB 0 112 192
    $s.Format.Line.Weight = 3
    $s.Format.Line.DashStyle = 4      # msoLineDash
    Write-Host 'OK 08D 线型虚线/粗细/颜色'
} catch { Write-Host "FAIL 08D: $_" }

# 08E 散点圆形标记12号
$ch = Add-Chart $ws 74 560 620 $ws.Range('G1:H6')
try {
    $s = $ch.SeriesCollection(1)
    $s.MarkerStyle = 8                # xlMarkerStyleCircle
    $s.MarkerSize = 12
    $s.MarkerBackgroundColor = RGB 192 0 0
    $s.MarkerForegroundColor = RGB 30 30 30
    Write-Host 'OK 08E 数据标记'
} catch { Write-Host "FAIL 08E: $_" }

# 08F 单数据点着色
$ch = Add-Chart $ws 51 1080 620 $ws.Range('J1:K6')
try {
    $ch.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = RGB 237 125 49
    Write-Host 'OK 08F 单点着色'
} catch { Write-Host "FAIL 08F: $_" }

# 08G 饼图扇区分离 30%
$ch = Add-Chart $ws 5 20 980 $ws.Range('A9:B12')
try {
    $ch.SeriesCollection(1).Points(3).Explosion = 30
    Write-Host 'OK 08G 扇区分离'
} catch { Write-Host "FAIL 08G: $_" }

# 08H gapWidth 50 + overlap -20
$ch = Add-Chart $ws 51 560 980 $ws.Range('D9:E12')
try {
    $cg = $ch.ChartGroups(1)
    $cg.GapWidth = 50
    $cg.Overlap = -20
    Write-Host 'OK 08H gapWidth/overlap'
} catch { Write-Host "FAIL 08H: $_" }

# 08I 圆环内径 40%
$ch = Add-Chart $ws -4120 1080 980 $ws.Range('G9:H12')
try {
    $ch.ChartGroups(1).DoughnutHoleSize = 40
    Write-Host 'OK 08I 圆环内径'
} catch { Write-Host "FAIL 08I: $_" }

# 08J 系列图案填充
$ch = Add-Chart $ws 51 20 1340 $ws.Range('J9:K12')
try {
    $ch.SeriesCollection(1).Format.Fill.Patterned(39)   # msoPatternZigZag
    Write-Host 'OK 08J 图案填充'
} catch { Write-Host "FAIL 08J: $_" }

foreach ($co in $ws.ChartObjects()) { Save-Png $co.Chart ('08_' + $co.Index + '.png') }

# ============ 09 高级元素 ============
$ws = New-Sheet('09高级元素')
Set-Range $ws 'A1:C7' @(@('月份','销量','增长率'), @('一',120,0.10), @('二',180,0.15), @('三',150,0.08), @('四',220,0.22), @('五',260,0.18), @('六',310,0.25))
Set-Range $ws 'E1:F9' @(@('X','Y'), @(1,3.2), @(2,5.4), @(3,7.1), @(4,9.6), @(5,11.2), @(6,13.5), @(7,15.1), @(8,17.8))
Set-Range $ws 'H1:I7' @(@('类目','值'), @('A',60), @('B',110), @('C',85), @('D',130), @('E',95), @('F',150))
Set-Range $ws 'K1:L7' @(@('类目','值'), @('A',55), @('B',90), @('C',70), @('D',120), @('E',100), @('F',140))

# 09A 次轴组合图（柱+线双轴）
$ch = Add-Chart $ws 51 20 260 $ws.Range('A1:C7')
try {
    $s2 = $ch.SeriesCollection(2)
    $s2.ChartType = 4          # xlLine
    $s2.AxisGroup = 2          # xlSecondary
    $s2.Format.Line.ForeColor.RGB = RGB 192 0 0
    Write-Host 'OK 09A 次轴组合图'
} catch { Write-Host "FAIL 09A: $_" }

# 09B 趋势线线性 + 公式 + R²
$ch = Add-Chart $ws 74 560 260 $ws.Range('E1:F9')
try {
    $s = $ch.SeriesCollection(1)
    $s.MarkerStyle = 8
    $tl = $s.Trendlines().Add(-4132)     # xlLinear
    $tl.DisplayEquation = $true
    $tl.DisplayRSquared = $true
    Write-Host 'OK 09B 趋势线线性+公式+R²'
} catch { Write-Host "FAIL 09B: $_" }

# 09C 趋势线移动平均 2 期
$ch = Add-Chart $ws 4 1080 260 $ws.Range('H1:I7')
try {
    $tl = $ch.SeriesCollection(1).Trendlines().Add(6)   # xlMovingAvg
    $tl.Period = 2
    $tl.Format.Line.ForeColor.RGB = RGB 197 90 17
    Write-Host 'OK 09C 移动平均趋势线'
} catch { Write-Host "FAIL 09C: $_" }

# 09D 误差线固定值 50
$ch = Add-Chart $ws 51 20 620 $ws.Range('H1:I7')
try {
    $ch.SeriesCollection(1).ErrorBar(1, 1, 1, 50)     # Y轴/双向/固定值/50
    Write-Host 'OK 09D 误差线固定值'
} catch { Write-Host "FAIL 09D: $_" }

# 09E 误差线标准差 1 倍
$ch = Add-Chart $ws 51 560 620 $ws.Range('H1:I7')
try {
    $ch.SeriesCollection(1).ErrorBar(1, 1, -4155, 1)  # Y轴/双向/标准差/1倍
    Write-Host 'OK 09E 误差线标准差'
} catch { Write-Host "FAIL 09E: $_" }

# 09F 图下数据表
$ch = Add-Chart $ws 51 1080 620 $ws.Range('K1:L7')
try {
    $ch.HasDataTable = $true
    Write-Host 'OK 09F 图下数据表'
} catch { Write-Host "FAIL 09F: $_" }

# 09G 绘图区填充浅黄+边框
$ch = Add-Chart $ws 51 20 980 $ws.Range('K1:L7')
try {
    $pa = $ch.PlotArea
    $pa.Format.Fill.ForeColor.RGB = RGB 255 242 204
    $pa.Format.Line.ForeColor.RGB = RGB 130 130 130
    $pa.Format.Line.Weight = 1.5
    Write-Host 'OK 09G 绘图区样式'
} catch { Write-Host "FAIL 09G: $_" }

# 09H 图表区渐变填充
$ch = Add-Chart $ws 51 560 980 $ws.Range('K1:L7')
try {
    $ca = $ch.ChartArea.Format
    $ca.Fill.ForeColor.RGB = RGB 226 239 218
    $ca.Fill.BackColor.RGB = RGB 255 255 255
    $ca.Fill.TwoColorGradient(1, 1)
    Write-Host 'OK 09H 图表区渐变'
} catch { Write-Host "FAIL 09H: $_" }

# 09I 趋势线前推外推预测（线性 + 前推2期）
$ch = Add-Chart $ws 74 1080 980 $ws.Range('E1:F9')
try {
    $s = $ch.SeriesCollection(1)
    $s.MarkerStyle = 8
    $tl = $s.Trendlines().Add(-4132)
    $tl.Forward = 2
    $tl.Backward = 1
    Write-Host 'OK 09I 趋势线前后推预测'
} catch { Write-Host "FAIL 09I: $_" }

foreach ($co in $ws.ChartObjects()) { Save-Png $co.Chart ('09_' + $co.Index + '.png') }

$wb.SaveAs($xlsx, 51)
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host 'DONE'

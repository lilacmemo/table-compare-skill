# build_batch1.ps1 —— 批次1样式主战场样本簿（B~H 组配置项）v3
# 数据用显式 object[,] 写入（PS 字面量二维数组 marshaling 不可靠）
$ErrorActionPreference = 'Continue'
$outDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsx = Join-Path $outDir '批次1样式样本.xlsx'
$pngDir = Join-Path $outDir 'batch1_baseline'
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
function Add-Chart($ws, $type, $left, $dataRange, $name) {
    $co = $ws.Shapes.AddChart2(-1, $type, $left, 260, 520, 320)
    $ch = $co.Chart
    $ch.SetSourceData($dataRange, 2)   # 2 = xlColumns，按列绘图
    return $ch
}
function Save-Png($ch, $tag) { try { $ch.Export((Join-Path $pngDir $tag), 'PNG') | Out-Null } catch { Write-Host "PNG FAIL $tag" } }

# ============ 01 标题 ============
$ws = New-Sheet('01标题')
Set-Range $ws 'A1:B6' @(@('月份','销售'), @('一月',120), @('二月',180), @('三月',150), @('四月',220), @('五月',260))
$ch = Add-Chart $ws 51 20 $ws.Range('A1:B6') '标题_红色加粗16'
try {
    $ch.HasTitle = $true
    $ch.ChartTitle.Text = '月度销售统计'
    $ch.ChartTitle.Font.Name = '微软雅黑'
    $ch.ChartTitle.Font.Size = 16
    $ch.ChartTitle.Font.Bold = $true
    $ch.ChartTitle.Font.Color = RGB 192 0 0
    Write-Host 'OK 01A 标题(文本/字体/加粗/颜色)'
} catch { Write-Host "FAIL 01A: $_" }
$ch2 = Add-Chart $ws 560 20 $ws.Range('A1:B6') '标题_隐藏'
try { $ch2.HasTitle = $false; Write-Host 'OK 01B 标题隐藏' } catch { Write-Host "FAIL 01B: $_" }
Save-Png $ch '01A_title.png'
Save-Png $ch2 '01B_title_hidden.png'

# ============ 02 轴刻度 ============
$ws = New-Sheet('02轴刻度')
Set-Range $ws 'A1:B7' @(@('月份','值'), @('一',120), @('二',180), @('三',150), @('四',220), @('五',260), @('六',310))
Set-Range $ws 'D1:E7' @(@('序','对数值'), @(1,1), @(2,10), @(3,100), @(4,1000), @(5,5000), @(6,10000))
Set-Range $ws 'G1:H7' @(@('类目','比例'), @('A',0.05), @('B',0.12), @('C',0.18), @('D',0.25), @('E',0.32), @('F',0.40))
Set-Range $ws 'J1:K7' @(@('日期','销量'), @('2026-01-05',30), @('2026-02-05',45), @('2026-03-05',38), @('2026-04-05',60), @('2026-05-05',72), @('2026-06-05',66))
$ch = Add-Chart $ws 51 20 $ws.Range('A1:B7') 'Y轴_0到500_主刻度100'
try {
    $ax = $ch.Axes(2)
    $ax.MinimumScale = 0; $ax.MaximumScale = 500; $ax.MajorUnit = 100
    Write-Host 'OK 02A Y轴最大最小/主刻度单位'
} catch { Write-Host "FAIL 02A: $_" }
$ch2 = Add-Chart $ws 4 540 $ws.Range('D1:E7') 'Y轴_对数刻度'
try { $ch2.Axes(2).ScaleType = -4133; Write-Host 'OK 02B 对数刻度' } catch { Write-Host "FAIL 02B: $_" }
$ch3 = Add-Chart $ws 51 20 $ws.Range('A1:B7') '分类轴_反转'
try { $ch3.Axes(1).ReversePlotOrder = $true; Write-Host 'OK 02C 分类轴反转' } catch { Write-Host "FAIL 02C: $_" }
$ch4 = Add-Chart $ws 51 540 $ws.Range('G1:H7') '刻度标签_百分比格式'
try { $ch4.Axes(2).TickLabels.NumberFormat = '0%'; Write-Host 'OK 02D 标签百分比格式' } catch { Write-Host "FAIL 02D: $_" }
$ws.Range('J2:J7').NumberFormat = 'yyyy-mm-dd'
$ch5 = Add-Chart $ws 4 20 $ws.Range('J1:K7') '日期轴'
try { $ch5.Axes(1).TickLabels.NumberFormat = 'm"月"d"日"'; Write-Host 'OK 02E 日期轴格式' } catch { Write-Host "FAIL 02E: $_" }
Save-Png $ch '02A_axis_scale.png'; Save-Png $ch2 '02B_log.png'; Save-Png $ch3 '02C_reverse.png'; Save-Png $ch4 '02D_numfmt.png'; Save-Png $ch5 '02E_date.png'

# ============ 03 刻度标签与轴线 ============
$ws = New-Sheet('03刻度标签')
Set-Range $ws 'A1:C6' @(@('产品','Q1','Q2'), @('甲',80,95), @('乙',60,70), @('丙',90,85), @('丁',55,65), @('戊',70,75))
$ch = Add-Chart $ws 51 20 $ws.Range('A1:C6') '刻度标签_紫色12号旋转45'
try {
    $ax = $ch.Axes(1)
    $ax.TickLabels.Font.Size = 12
    $ax.TickLabels.Font.Color = RGB 112 48 160
    $ax.TickLabels.Font.Bold = $true
    $ax.TickLabels.Orientation = -45
    $ax.MajorTickMark = 3   # xlTickMarkOutside
    Write-Host 'OK 03A 刻度标签字体/颜色/加粗/旋转-45 + 刻度线朝外'
} catch { Write-Host "FAIL 03A: $_" }
$ch2 = Add-Chart $ws 51 560 $ws.Range('A1:C6') '刻度标签_隐藏_刻度线无'
try {
    $ch2.Axes(1).TickLabelPosition = -4142   # xlNone
    $ch2.Axes(1).MajorTickMark = -4142
    Write-Host 'OK 03B 标签隐藏/刻度线无'
} catch { Write-Host "FAIL 03B: $_" }
$ch3 = Add-Chart $ws 51 1080 $ws.Range('A1:C6') '轴线_红色加粗'
try {
    $ax3 = $ch3.Axes(2)
    $ax3.Format.Line.ForeColor.RGB = RGB 200 0 0
    $ax3.Format.Line.Weight = 2.5
    Write-Host 'OK 03C 轴线颜色/粗细'
} catch { Write-Host "FAIL 03C: $_" }
Save-Png $ch '03A_ticklabels.png'; Save-Png $ch2 '03B_hidden.png'; Save-Png $ch3 '03C_axisline.png'

# ============ 04 网格线 ============
$ws = New-Sheet('04网格线')
Set-Range $ws 'A1:C6' @(@('产品','Q1','Q2'), @('甲',80,95), @('乙',60,70), @('丙',90,85), @('丁',55,65), @('戊',70,75))
$ch = Add-Chart $ws 51 20 $ws.Range('A1:C6') '主网格线_红色虚线2磅'
try {
    $ch.Axes(2).HasMajorGridlines = $true
    $gl = $ch.Axes(2).MajorGridlines
    $gl.Format.Line.ForeColor.RGB = RGB 200 0 0
    $gl.Format.Line.DashStyle = 4      # msoLineDash
    $gl.Format.Line.Weight = 2
    Write-Host 'OK 04A 主横网格线颜色/虚线/粗细'
} catch { Write-Host "FAIL 04A: $_" }
$ch2 = Add-Chart $ws 51 560 $ws.Range('A1:C6') '分类轴网格线_蓝色'
try {
    $ch2.Axes(1).HasMajorGridlines = $true
    $ch2.Axes(1).MajorGridlines.Format.Line.ForeColor.RGB = RGB 0 112 192
    Write-Host 'OK 04B 纵网格线蓝色'
} catch { Write-Host "FAIL 04B: $_" }
$ch3 = Add-Chart $ws 51 1080 $ws.Range('A1:C6') '次网格线'
try {
    $ch3.Axes(2).HasMinorGridlines = $true
    $ch3.Axes(2).MinorUnit = 25
    Write-Host 'OK 04C 次横网格线'
} catch { Write-Host "FAIL 04C: $_" }
Save-Png $ch '04A_grid.png'; Save-Png $ch2 '04B_catgrid.png'; Save-Png $ch3 '04C_minor.png'

# ============ 05 图例 ============
$ws = New-Sheet('05图例')
Set-Range $ws 'A1:C4' @(@('季度','北京','上海'), @('Q1',120,90), @('Q2',150,110), @('Q3',180,140))
$ch = Add-Chart $ws 51 20 $ws.Range('A1:C4') '图例_底部橙色12加粗_填充边框'
try {
    $ch.HasLegend = $true
    $ch.Legend.Position = -4107   # xlLegendPositionBottom
    $ch.Legend.Font.Size = 12
    $ch.Legend.Font.Bold = $true
    $ch.Legend.Font.Color = RGB 197 90 17
    $ch.Legend.Format.Fill.ForeColor.RGB = RGB 255 242 204
    $ch.Legend.Format.Line.ForeColor.RGB = RGB 191 143 0
    Write-Host 'OK 05A 图例位置/字体/加粗/颜色/背景/边框'
} catch { Write-Host "FAIL 05A: $_" }
$ch2 = Add-Chart $ws 51 560 $ws.Range('A1:C4') '图例_隐藏'
try { $ch2.HasLegend = $false; Write-Host 'OK 05B 图例隐藏' } catch { Write-Host "FAIL 05B: $_" }
$ch3 = Add-Chart $ws 51 1080 $ws.Range('A1:C4') '图例_删除第二项'
try {
    $ch3.HasLegend = $true
    $ch3.Legend.LegendEntries(2).Delete()
    Write-Host 'OK 05C 单图例项显隐'
} catch { Write-Host "FAIL 05C: $_" }
Save-Png $ch '05A_legend.png'; Save-Png $ch2 '05B_nolegend.png'; Save-Png $ch3 '05C_delentry.png'

# ============ 06 数据标签 ============
$ws = New-Sheet('06数据标签')
Set-Range $ws 'A1:B7' @(@('月份','金额'), @('一月',12345), @('二月',24500), @('三月',18600), @('四月',30200), @('五月',27800), @('六月',35100))
Set-Range $ws 'D1:E5' @(@('渠道','占比'), @('线上',0.55), @('门店',0.30), @('分销',0.15))
Set-Range $ws 'G1:H7' @(@('月份','销量'), @('一',30), @('二',45), @('三',38), @('四',60), @('五',72), @('六',66))
$ch = Add-Chart $ws 51 20 $ws.Range('A1:B7') '标签_值_外侧_千分位_蓝色加粗'
try {
    $s = $ch.SeriesCollection(1)
    $s.HasDataLabels = $true
    Write-Host 'OK 06A-1 标签显隐'
} catch { Write-Host "FAIL 06A-1: $_" }
try {
    $dl = $s.DataLabels
    $dl.Position = 2          # xlLabelPositionOutsideEnd
    $dl.NumberFormat = '#,##0'
    $dl.Font.Color = RGB 0 112 192
    $dl.Font.Bold = $true
    Write-Host 'OK 06A-2 位置/数字格式/字体'
} catch {
    try { $ch.ApplyDataLabels(2); Write-Host 'OK 06A-2(降级ApplyDataLabels)' } catch { Write-Host "FAIL 06A-2: $_" }
}
$ch2 = Add-Chart $ws 5 560 $ws.Range('D1:E5') '饼图_百分比标签'
try { $ch2.ApplyDataLabels(3); Write-Host 'OK 06B 饼图百分比标签' } catch { Write-Host "FAIL 06B: $_" }
$ch3 = Add-Chart $ws 4 1080 $ws.Range('G1:H7') '标签_类别名'
try { $ch3.ApplyDataLabels(1); Write-Host 'OK 06C 标签类别名' } catch { Write-Host "FAIL 06C: $_" }
$ch4 = Add-Chart $ws 51 560 $ws.Range('G1:H7') '标签_单点独立'
try {
    $s4 = $ch4.SeriesCollection(1)
    $s4.Points(4).HasDataLabel = $true
    $s4.Points(4).DataLabel.Text = '峰值60'
    $s4.Points(4).DataLabel.Font.Color = RGB 192 0 0
    Write-Host 'OK 06D 单数据点独立标签'
} catch { Write-Host "FAIL 06D: $_" }
Save-Png $ch '06A_datalabel.png'; Save-Png $ch2 '06B_pie.png'; Save-Png $ch3 '06C_catname.png'; Save-Png $ch4 '06D_point.png'

# ============ 07 轴标题 ============
$ws = New-Sheet('07轴标题')
Set-Range $ws 'A1:B7' @(@('月份','金额'), @('一月',120), @('二月',180), @('三月',150), @('四月',220), @('五月',260), @('六月',310))
$ch = Add-Chart $ws 51 20 $ws.Range('A1:B7') '轴标题_分类旋转45_数值水平蓝'
try {
    $ch.Axes(1).HasTitle = $true
    $ch.Axes(1).AxisTitle.Text = '统计月份'
    $ch.Axes(1).AxisTitle.Orientation = -45
    $ch.Axes(2).HasTitle = $true
    $ch.Axes(2).AxisTitle.Text = '金额(万元)'
    $ch.Axes(2).AxisTitle.Font.Color = RGB 0 112 192
    $ch.Axes(2).AxisTitle.Font.Bold = $true
    Write-Host 'OK 07A 轴标题文本/旋转/颜色/加粗'
} catch { Write-Host "FAIL 07A: $_" }
Save-Png $ch '07A_axis_titles.png'

# 收尾
$wb.Worksheets.Item('Sheet1').Delete()
$wb.SaveAs($xlsx, 51)
$wb.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "DONE $xlsx"

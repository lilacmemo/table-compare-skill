# Build a chart with explicitly styled elements (fonts, legend, labels, gridlines)
# so we can verify whether Univer's import preserves those configs.
$ErrorActionPreference = 'Stop'
$out = 'D:\workbuddy\spreadjs vs univer\图表配置测试.xlsx'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = '配置测试'

    $data = @(
        @('月份', '华东', '华南', '华北'),
        @('1月', 326, 287, 246),
        @('2月', 342, 301, 238),
        @('3月', 358, 296, 259),
        @('4月', 401, 318, 271),
        @('5月', 389, 335, 263),
        @('6月', 432, 349, 284)
    )
    for ($i = 0; $i -lt $data.Count; $i++) {
        for ($j = 0; $j -lt $data[$i].Count; $j++) {
            $cell = $ws.Cells.Item(($i + 1), ($j + 1))
            $v = $data[$i][$j]
            if ($i -eq 0 -or $j -eq 0) { $cell.Value2 = [string]$v } else { $cell.Value2 = [double]$v }
        }
    }
    $ws.Range('A1:D7').Font.Name = 'Microsoft YaHei'
    $ws.Range('A1:D7').Font.Size = 11
    $ws.Columns.AutoFit() | Out-Null

    # Clustered column chart
    $shape = $ws.Shapes.AddChart2(-1, 51, 20, 140, 620, 380)
    $ch = $shape.Chart
    $ch.SetSourceData($ws.Range('A1:D7'))
    $ch.HasTitle = $true
    $ch.ChartTitle.Text = '配置还原测试图'
    $ch.ChartTitle.Font.Size = 20
    $ch.ChartTitle.Font.Bold = $true
    $ch.ChartTitle.Font.Name = 'Microsoft YaHei'
    $ch.ChartTitle.Font.Color = 255            # red (BGR)

    # Legend: bottom, font 16, blue, bold
    $ch.HasLegend = $true
    $ch.Legend.Position = -4107                # xlBottom
    $ch.Legend.Font.Size = 16
    $ch.Legend.Font.Bold = $true
    $ch.Legend.Font.Name = 'Microsoft YaHei'
    $ch.Legend.Font.Color = 16711680           # blue (BGR)

    # Category axis title + font
    $ch.Axes(1).HasTitle = $true
    $ch.Axes(1).AxisTitle.Text = '月份（月）'
    $ch.Axes(1).AxisTitle.Font.Size = 14
    $ch.Axes(1).AxisTitle.Font.Italic = $true
    $ch.Axes(1).AxisTitle.Font.Name = 'Microsoft YaHei'
    $ch.Axes(1).TickLabels.Font.Size = 13
    $ch.Axes(1).TickLabels.Font.Color = 8388736   # purple-ish BGR

    # Value axis: number format + explicit gridline style
    $ch.Axes(2).HasTitle = $true
    $ch.Axes(2).AxisTitle.Text = '销售额（万元）'
    $ch.Axes(2).AxisTitle.Font.Size = 14
    $ch.Axes(2).AxisTitle.Font.Name = 'Microsoft YaHei'
    $ch.Axes(2).TickLabels.NumberFormat = '#,##0.0'
    try { $ch.Axes(2).HasMajorGridlines = $true
          $ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = 12632256
          $ch.Axes(2).MajorGridlines.Format.Line.Weight = 1.25
          $ch.Axes(2).MajorGridlines.Format.Line.DashStyle = 4 } catch { Write-Output "GRIDLINE_SKIP: $($_.Exception.Message)" }

    # Data labels: value, outside end, font 12
    try {
        $ch.SeriesCollection(1).HasDataLabels = $true
        $ch.SeriesCollection(1).DataLabels.Position = 2
        $ch.SeriesCollection(1).DataLabels.Font.Size = 12
        $ch.SeriesCollection(1).DataLabels.Font.Bold = $true
        $ch.SeriesCollection(1).DataLabels.NumberFormat = '#,##0'
    } catch { Write-Output "DATALABEL_SKIP: $($_.Exception.Message)" }

    # Series styling: fill color + gap width
    try { $ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = 12611584
          $ch.SeriesCollection(1).Format.Line.ForeColor.RGB = 6299648 } catch { Write-Output "SERIES_SKIP: $($_.Exception.Message)" }
    try { $ch.ChartGroups(1).GapWidth = 80
          $ch.ChartGroups(1).Overlap = -10 } catch { Write-Output "GROUP_SKIP: $($_.Exception.Message)" }

    if (Test-Path $out) { Remove-Item $out -Force }
    $wb.SaveAs($out, 51)
    Write-Output "SAVED: $out"
    Write-Output "SHEETS: $($wb.Worksheets.Count)"
    Write-Output "CHARTS: $($ws.ChartObjects().Count)"
    $wb.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
}

# export_excelview.ps1 —— 导出方向"Excel 实际打开展示"核对：
# 打开 4 个 roundtrip 导出文件，逐图 Chart.Export 到 excelview_export/<tag>/，命名与 baseline_export 一致
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent   # 工作区根目录
$pc   = Join-Path $root 'poc-compare'
$out  = Join-Path $pc 'excelview_export'

$jobs = @(
    @{ tag='sjs_b1';    file='sjs_roundtrip_batch1.xlsx'; sheets=@('01标题','02轴刻度','03刻度标签','04网格线','05图例','06数据标签','07轴标题'); prefix='b1' },
    @{ tag='sjs_b2';    file='sjs_roundtrip_batch2.xlsx'; sheets=@('08系列样式','09高级元素'); prefix='b2' },
    @{ tag='univer_b1'; file='univer_roundtrip_batch1.xlsx'; sheets=@('01标题','02轴刻度','03刻度标签','04网格线','05图例','06数据标签','07轴标题'); prefix='b1' },
    @{ tag='univer_b2'; file='univer_roundtrip_batch2.xlsx'; sheets=@('08系列样式','09高级元素'); prefix='b2' }
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$report = @()
try {
    foreach ($job in $jobs) {
        $dir = Join-Path $out $job.tag
        if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        $fp = Join-Path $pc $job.file
        $wb = $xl.Workbooks.Open($fp, 0, $true)
        $total = 0
        foreach ($sn in $job.sheets) {
            $ws = $null
            try { $ws = $wb.Worksheets.Item($sn) } catch { Write-Output ("MISSING SHEET " + $job.tag + " " + $sn); continue }
            $n = 1
            foreach ($co in $ws.ChartObjects()) {
                $p = Join-Path $dir ($job.prefix + '_' + $sn + '_c' + $n + '.png')
                try { $co.Chart.Export($p, 'PNG') | Out-Null } catch { }
                if (!(Test-Path $p) -or (Get-Item $p).Length -eq 0) {
                    # 兜底：Chart.Export 对部分 3D 图写出 0 字节 → CopyPicture 粘贴到临时图表再导出
                    try {
                        $co.CopyPicture()
                        $tmp = $ws.ChartObjects().Add(0, 0, [Math]::Max($co.Width, 100), [Math]::Max($co.Height, 100))
                        $tmp.Activate()
                        $tmp.Chart.Paste()
                        $tmp.Chart.Export($p, 'PNG') | Out-Null
                        $tmp.Delete()
                    } catch { Write-Output ("FALLBACK FAIL " + $job.tag + " " + $sn + " c" + $n + ": " + $_.Exception.Message) }
                }
                if ((Test-Path $p) -and (Get-Item $p).Length -gt 0) { $total++ }
                else { Write-Output ("EXPORT FAIL " + $job.tag + " " + $sn + " c" + $n) }
                $n++
            }
        }
        $report += ("{0}: {1} charts exported" -f $job.tag, $total)
        $wb.Close($false)
    }
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
$report | ForEach-Object { Write-Output $_ }

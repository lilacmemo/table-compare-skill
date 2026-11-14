# export_baseline_fallback.ps1 —— 从源样本文件重新导出基线图（含 CopyPicture 兜底），刷新 baseline_export/
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent
$pc   = Join-Path $root 'poc-compare'
$out  = Join-Path $pc 'baseline_export'

$jobs = @(
    @{ tag='b1'; file='批次1样式样本.xlsx'; sheets=@('01标题','02轴刻度','03刻度标签','04网格线','05图例','06数据标签','07轴标题') },
    @{ tag='b2'; file='批次2系列与高级元素样本.xlsx'; sheets=@('08系列样式','09高级元素') }
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    foreach ($job in $jobs) {
        $wb = $xl.Workbooks.Open((Join-Path $root $job.file), 0, $true)
        foreach ($sn in $job.sheets) {
            $ws = $wb.Worksheets.Item($sn)
            $n = 1
            foreach ($co in $ws.ChartObjects()) {
                $p = Join-Path $out ($job.tag + '_' + $sn + '_c' + $n + '.png')
                try { $co.Chart.Export($p, 'PNG') | Out-Null } catch { }
                if (!(Test-Path $p) -or (Get-Item $p).Length -eq 0) {
                    try {
                        $co.CopyPicture()
                        $tmp = $ws.ChartObjects().Add(0, 0, [Math]::Max($co.Width, 100), [Math]::Max($co.Height, 100))
                        $tmp.Activate()
                        $tmp.Chart.Paste()
                        $tmp.Chart.Export($p, 'PNG') | Out-Null
                        $tmp.Delete()
                    } catch { Write-Output ("FALLBACK FAIL " + $job.tag + " " + $sn + " c" + $n + ": " + $_.Exception.Message) }
                }
                if ((Test-Path $p) -and (Get-Item $p).Length -gt 0) { Write-Output ("OK " + $job.tag + " " + $sn + " c" + $n) }
                else { Write-Output ("FAIL " + $job.tag + " " + $sn + " c" + $n) }
                $n++
            }
        }
        $wb.Close($false)
    }
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

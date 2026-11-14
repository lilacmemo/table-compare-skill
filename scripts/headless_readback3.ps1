$ErrorActionPreference = 'Stop'
$Path = 'D:\workbuddy\spreadjs vs univer\公式对比-Excel365基准519函数-三方-v2.xlsx'
$Out  = 'D:\workbuddy\spreadjs vs univer\poc-compare\_headless_readback.json'

$REJECTED = -2147418111

function Retry([scriptblock]$sb, [int]$max = 120) {
    for ($i = 0; $i -lt $max; $i++) {
        try { return & $sb }
        catch {
            $ex = $_.Exception
            while ($ex -and $ex.InnerException) { $ex = $ex.InnerException }
            $hr = if ($ex) { $ex.HResult } else { 0 }
            if ($hr -eq $REJECTED -or $hr -eq -2147417846) { Start-Sleep -Milliseconds 500 }
            else { throw }
        }
    }
    throw 'retry exhausted'
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

# 自动计算模式：fullCalcOnLoad=1 会让 Excel 打开时全量重算，Open 期间 COM 会被拒，靠 Retry 等待
$wb = Retry { $xl.Workbooks.Open($Path, 0, $true) } 240
Write-Output ("opened " + (Retry { $wb.Name }) + " sheets=" + (Retry { $wb.Worksheets.Count }))
Write-Output ("calcState: " + (Retry { $xl.CalculationState }))

$ws = $wb.Sheets.Item('对比表')
$obj = [System.Collections.ArrayList]@()
for ($r = 5; $r -le 523; $r++) {
    [void]$obj.Add(@{
        name   = (Retry { $ws.Cells.Item($r, 2).Text })
        sjs_sup = (Retry { $ws.Cells.Item($r, 7).Text })
        sjs_j  = (Retry { $ws.Cells.Item($r, 9).Text })
        uni_sup = (Retry { $ws.Cells.Item($r, 10).Text })
        uni_j  = (Retry { $ws.Cells.Item($r, 12).Text })
    })
}
Write-Output ("read rows: " + $obj.Count)

$sum = $wb.Sheets.Item('结论')
$concl = [System.Collections.ArrayList]@()
foreach ($r in 5..38) {
    $a = Retry { $sum.Cells.Item($r, 1).Text }
    $b = Retry { $sum.Cells.Item($r, 2).Text }
    $c = Retry { $sum.Cells.Item($r, 3).Text }
    if ($a -or $b -or $c) { [void]$concl.Add(@{ r = $r; label = $a; sjs = $b; uni = $c }) }
}
Write-Output ("read conclusion rows: " + $concl.Count)

@{ rows = $obj; conclusion = $concl } | ConvertTo-Json -Depth 4 | Out-File -FilePath $Out -Encoding utf8
Write-Output ("written " + $Out)

Retry { $wb.Close($false) } | Out-Null
Retry { $xl.Quit() } | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Write-Output 'DONE'

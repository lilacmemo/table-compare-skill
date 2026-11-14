# fill_live_v2.ps1 -- let Excel itself write the live formulas into the 'Excel实测' sheet.
# ASCII only: PS 5.1 misparses a BOM when the file body is pure ASCII.
param(
    [string]$Book = 'D:\workbuddy\spreadjs vs univer\公式对比-Excel365基准519函数-三方-v2.xlsx',
    [string]$CasesPath = 'D:\workbuddy\spreadjs vs univer\poc-compare\cases_v2.json',
    [string]$DataPath = 'D:\workbuddy\spreadjs vs univer\poc-compare\data_block_v2.json'
)

$ErrorActionPreference = 'Stop'

$casesText = [IO.File]::ReadAllText($CasesPath, [Text.Encoding]::UTF8)
$cases = ConvertFrom-Json -InputObject $casesText
$dataText = [IO.File]::ReadAllText($DataPath, [Text.Encoding]::UTF8)
$data = ConvertFrom-Json -InputObject $dataText

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$xl.EnableEvents = $false
Start-Sleep -Seconds 2

$wb = $null
try {
    $wb = $xl.Workbooks.Open($Book)
    $ws = $wb.Worksheets.Item('Excel实测')
    Write-Output ('sheet: ' + $ws.Name)

    foreach ($c in $data.cells) {
        $val = $c.value
        if ($val -is [string]) { $ws.Range($c.cell).Value2 = [string]$val }
        else { $ws.Range($c.cell).Value2 = [double]$val }
    }
    foreach ($f in $data.formulaCells) {
        $ws.Range($f.cell).Formula = $f.formula
    }

    $xl.Calculation = -4135
    $fail = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $cases.Count; $i++) {
        $er = 2 + $i * 12
        $cell = $ws.Cells.Item($er, 26)
        try { $cell.Formula2 = [string]$cases[$i].formula }
        catch {
            try { $cell.Formula = [string]$cases[$i].formula }
            catch { [void]$fail.Add([string]$cases[$i].name) }
        }
    }
    $xl.Calculation = -4105
    $xl.CalculateFullRebuild()
    Start-Sleep -Milliseconds 3000
    $xl.Calculate()
    Start-Sleep -Milliseconds 1500

    $wb.Save()
    $wb.Close($false)
    Write-Output ('formulas ' + $cases.Count + ' | setFail ' + $fail.Count)
    if ($fail.Count -gt 0) { Write-Output ('fail: ' + ($fail -join ',')) }
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
}

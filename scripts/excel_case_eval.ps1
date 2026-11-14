# excel_case_eval.ps1 -- evaluate every formula case in real Excel 365 (the oracle).
# Writes the shared data block, then each case formula, recalculates, and dumps values to JSON.
# ASCII only on purpose: PS 5.1 misparses BOM when the file body is pure ASCII.
param(
    [string]$CasesPath = 'D:\workbuddy\spreadjs vs univer\poc-compare\cases_v2.json',
    [string]$DataPath = 'D:\workbuddy\spreadjs vs univer\poc-compare\data_block_v2.json',
    [string]$OutPath = 'D:\workbuddy\spreadjs vs univer\poc-compare\excel_eval_v2.json'
)

$ErrorActionPreference = 'Stop'

function De-XlErr([object]$v) {
    # Value2 returns a negative Int32 for error cells: 0x800A0000 | xlErrCode
    if ($null -eq $v) { return $null }
    if ($v -is [int]) {
        $code = $v -band 0xFFFF
        switch ($code) {
            2000 { return '#NULL!' }
            2007 { return '#DIV/0!' }
            2015 { return '#VALUE!' }
            2023 { return '#REF!' }
            2029 { return '#NAME?' }
            2036 { return '#NUM!' }
            2042 { return '#N/A' }
            2043 { return '#GETTING_DATA' }
            2047 { return '#SPILL!' }
            2050 { return '#CALC!' }
            default { return ('#ERR' + $code) }
        }
    }
    return $null
}

$casesText = [IO.File]::ReadAllText($CasesPath, [Text.Encoding]::UTF8)
$cases = ConvertFrom-Json -InputObject $casesText
$dataText = [IO.File]::ReadAllText($DataPath, [Text.Encoding]::UTF8)
$data = ConvertFrom-Json -InputObject $dataText
Write-Output ('parsed cases type=' + $cases.GetType().FullName + ' count=' + $cases.Count + ' | data cells ' + $data.cells.Count)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$xl.EnableEvents = $false
Start-Sleep -Seconds 2

$wb = $null
$results = New-Object System.Collections.ArrayList

try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Eval'

    # ---- data block ----
    foreach ($c in $data.cells) {
        $val = $c.value
        if ($val -is [string]) { $ws.Range($c.cell).Value2 = [string]$val }
        else { $ws.Range($c.cell).Value2 = [double]$val }
    }
    foreach ($f in $data.formulaCells) {
        $ws.Range($f.cell).Formula = $f.formula
    }

    # ---- formulas in column Z ----
    $xl.Calculation = -4135   # manual
    $n = $cases.Count
    $fail = 0
    for ($i = 0; $i -lt $n; $i++) {
        $r = 2 + $i
        $cell = $ws.Cells.Item($r, 26)
        $ok = $true
        try { $cell.Formula2 = [string]$cases[$i].formula }
        catch {
            try { $cell.Formula = [string]$cases[$i].formula }
            catch { $ok = $false; $fail++ }
        }
        [void]$results.Add([ordered]@{
            name = $cases[$i].name; formula = $cases[$i].formula; setOk = $ok;
            value = $null; valueType = $null; err = $null
        })
    }

    $xl.Calculation = -4105   # automatic
    $xl.CalculateFullRebuild()
    Start-Sleep -Milliseconds 2500
    $xl.Calculate()
    Start-Sleep -Milliseconds 1000

    Write-Output ('reading ' + $n + ' cells')
    for ($i = 0; $i -lt $n; $i++) {
        $v = $ws.Cells.Item(2 + $i, 26).Value2
        $e = De-XlErr $v
        $r = $results[$i]
        if (-not $r['setOk']) {
            $r['valueType'] = 'reject'; $r['value'] = $null; $r['err'] = 'SET_REJECT'
            continue
        }
        if ($e) {
            $r['valueType'] = 'error'; $r['err'] = $e; $r['value'] = $e
        }
        elseif ($null -eq $v -or ($v -is [string] -and $v -eq '')) {
            $r['valueType'] = 'empty'; $r['value'] = ''
        }
        elseif ($v -is [bool]) { $r['valueType'] = 'boolean'; $r['value'] = $v }
        elseif ($v -is [double] -or $v -is [decimal] -or $v -is [int] -or $v -is [long]) {
            $r['valueType'] = 'number'; $r['value'] = [double]$v
        }
        else { $r['valueType'] = 'string'; $r['value'] = [string]$v }
    }

    $wb.Close($false)
    Write-Output ('cases ' + $n + ' | setFail ' + $fail)
    $errCount = ($results | Where-Object { $_.valueType -eq 'error' }).Count
    $rejCount = ($results | Where-Object { $_.valueType -eq 'reject' }).Count
    Write-Output ('error cells ' + $errCount + ' | rejected ' + $rejCount)

    $json = $results | ConvertTo-Json -Depth 6
    [IO.File]::WriteAllText($OutPath, $json, (New-Object Text.UTF8Encoding($false)))
    Write-Output ('saved ' + $OutPath)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
}

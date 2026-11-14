# Recalculate a workbook with real Excel so the xlsx carries cached values.
# ASCII only on purpose: powershell -File reads .ps1 with the ANSI codepage,
# so a UTF-8 script containing CJK paths gets mangled. Caller copies the file
# to this ASCII path first, then copies it back.
$ErrorActionPreference = 'Stop'
$p = 'C:\temp\_sample_finalize.xlsx'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($p)
$xl.CalculateFullRebuild()
$wb.Save()
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Output 'recalc ok'

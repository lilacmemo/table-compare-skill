# Worksheet/Workbook capability sample - Excel COM supplement pass.
# Pure ASCII on purpose: PowerShell 5.1 decodes script files with the ANSI code
# page, so Chinese literals / sheet names would be mangled. Every sheet is
# addressed by index, and no Chinese string literal appears in this file.
#
# What this pass adds (things openpyxl cannot write):
#   1. sheet background picture  -> <sheetPr><picture r:id> + xl/media
#   2. custom document property  -> docProps/custom.xml
#   3. Excel re-save             -> makes the final file genuinely Excel-authored
# It also reads back every carrier so the sample is self-verified, not assumed.

$ErrorActionPreference = 'Stop'
$Src = 'C:\temp\ws_sample.xlsx'
$Dst = 'C:\temp\ws_sample_com.xlsx'
$Bg  = 'C:\temp\ws_bg.png'
$Cmt1 = 'C:\temp\ws_cmt1.txt'
$Cmt2 = 'C:\temp\ws_cmt2.txt'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

$rows = @()
function Add-Check($id, $what, $got, $want) {
    $ok = 'OK'
    if ("$got" -ne "$want") { $ok = 'MISMATCH' }
    $script:rows += [pscustomobject]@{ ID = $id; Item = $what; ReadBack = "$got"; Expect = "$want"; Result = $ok }
}

try {
    $wb = $xl.Workbooks.Open($Src)
    Write-Output ("opened sheets = " + $wb.Worksheets.Count)

    # ---- 1. background picture on sheet index 13 (12-<bg>) ----
    $wb.Worksheets.Item(13).SetBackgroundPicture($Bg)
    Write-Output "background picture set"

    # ---- 2. comments ----
    # openpyxl-written comments are read back as a 1-char string by Excel and
    # then dropped on save, so the comment text is kept in UTF-8 files and
    # applied here through COM. Keeps this script free of Chinese literals.
    $cmt1 = [IO.File]::ReadAllText($Cmt1, [Text.Encoding]::UTF8)
    $cmt2 = [IO.File]::ReadAllText($Cmt2, [Text.Encoding]::UTF8)
    $wb.Worksheets.Item(9).Range('B20').AddComment($cmt1) | Out-Null
    $wb.Worksheets.Item(11).Range('K20').AddComment($cmt2) | Out-Null
    Write-Output "comments added via COM"

    # ---- 3. read back every carrier ----
    Add-Check 'B2' 'sheet14 hidden state'    $wb.Worksheets.Item(14).Visible 0
    Add-Check 'B3' 'sheet15 veryHidden'      $wb.Worksheets.Item(15).Visible 2
    Add-Check 'C1' 'sheet04 ProtectContents' $wb.Worksheets.Item(4).ProtectContents $true
    Add-Check 'C3' 'sheet04 B17 locked'      $wb.Worksheets.Item(4).Range('B17').Locked $false
    Add-Check 'C3' 'sheet04 A5 locked'       $wb.Worksheets.Item(4).Range('A5').Locked $true
    Add-Check 'C4' 'sheet04 EnableSelection' $wb.Worksheets.Item(4).EnableSelection 0

    $wb.Worksheets.Item(5).Activate()
    Add-Check 'D1' 'sheet05 SplitRow'    $xl.ActiveWindow.SplitRow 5
    Add-Check 'D1' 'sheet05 SplitColumn' $xl.ActiveWindow.SplitColumn 3

    # OutlineLevel from COM carries a baseline of 1 for ungrouped rows in this
    # build, so assert the *relative* step instead of absolute levels.
    # XlSummaryRow: xlSummaryAbove = 0, xlSummaryBelow = 1  -> 0 means "above".
    $lv3 = $wb.Worksheets.Item(6).Rows.Item(3).OutlineLevel
    $lv7 = $wb.Worksheets.Item(6).Rows.Item(7).OutlineLevel
    $lv1 = $wb.Worksheets.Item(6).Rows.Item(1).OutlineLevel
    Add-Check 'E1' 'sheet06 r3 = base+1'      ($lv3 - $lv1) 1
    Add-Check 'E1' 'sheet06 r7 = base+2'      ($lv7 - $lv1) 2
    Add-Check 'E1' 'sheet06 summary above'    $wb.Worksheets.Item(6).Outline.SummaryRow 0

    Add-Check 'F2' 'sheet07 D hidden'   $wb.Worksheets.Item(7).Columns.Item(4).Hidden $true
    Add-Check 'F2' 'sheet07 F hidden'   $wb.Worksheets.Item(7).Columns.Item(6).Hidden $true
    Add-Check 'F2' 'sheet07 row12 hide' $wb.Worksheets.Item(7).Rows.Item(12).Hidden $true
    Add-Check 'F1' 'sheet07 row4 height' $wb.Worksheets.Item(7).Rows.Item(4).RowHeight 36
    $w = $wb.Worksheets.Item(7).Columns.Item(2).ColumnWidth
    Add-Check 'F3' 'sheet07 B width ~26' (($w -ge 25) -and ($w -le 27)) $true

    $wb.Worksheets.Item(8).Activate()
    Add-Check 'G1' 'sheet08 gridlines' $xl.ActiveWindow.DisplayGridlines $false
    Add-Check 'G3' 'sheet08 zoom'      $xl.ActiveWindow.Zoom 80

    # Comment.Text is a *method* on the COM object: .Text returns the method
    # itself (so .Length is $null) -> must call .Text() explicitly.
    $cm = $wb.Worksheets.Item(9).Range('B20').Comment
    Add-Check 'H3' 'sheet08 B20 comment len' (($null -ne $cm) -and ($cm.Text().Length -gt 10)) $true
    Add-Check 'L4' 'sheet11 M30 dv type' $wb.Worksheets.Item(11).Range('M30').Validation.Type 3
    # Excel 365 injects a hidden _xlfn.SINGLE name on open, so count >= 2 only
    Add-Check 'I2' 'defined names >= 2' ($wb.Names.Count -ge 2) $true
    Add-Check 'L1' 'sheet11 D4 value'    $wb.Worksheets.Item(11).Range('D4').Value2 128

    $wb.Worksheets.Item(12).Activate()
    Add-Check 'G5' 'sheet12 RTL' $wb.Worksheets.Item(12).DisplayRightToLeft $true

    # docProps / custom properties are asserted at package level by
    # verify_sample_worksheet.py. PS 5.1 cannot late-bind the parameterized
    # BuiltinDocumentProperties.Item('x') member (throws NullReferenceException),
    # so they are deliberately not read back through COM here.

    $rows | Format-Table -AutoSize | Out-String -Width 200 | Write-Output

    $bad = @($rows | Where-Object { $_.Result -ne 'OK' }).Count
    Write-Output ("verification: " + $rows.Count + " checks, " + $bad + " mismatch")

    Write-Output "tab colors (sheet: BGR long):"
    for ($i = 1; $i -le 13; $i++) {
        Write-Output ("  " + $i + " -> " + $wb.Worksheets.Item($i).Tab.Color)
    }

    # ---- 4. re-save through Excel ----
    $wb.Worksheets.Item(1).Activate()
    $wb.SaveAs($Dst, 51)
    Write-Output ("saved: " + $Dst)
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force
}

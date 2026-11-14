# Add the elements openpyxl cannot write (shapes, text box, sparklines),
# recalc formulas, save, and export the whole workbook to PDF.
# Pure ASCII by design: PS 5.1 reads .ps1 as GBK and mangles non-ASCII paths.
param(
  [string]$InFile = 'C:\temp\print-sample.xlsx',
  [string]$PdfOut = 'C:\temp\print-excel.pdf'
)
$ErrorActionPreference = 'Stop'
if (Test-Path $PdfOut) { Remove-Item $PdfOut -Force }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($InFile)

# ---- shapes on Visuals ----
$vs = $wb.Worksheets.Item('Visuals')
$rect = $vs.Shapes.AddShape(1, 520, 60, 150, 70)
$rect.Fill.ForeColor.RGB = 13561798
$rect.Line.ForeColor.RGB = 192
$rect.TextFrame2.TextRange.Text = 'Rectangle'
$rect.TextFrame2.TextRange.Font.Size = 11

$arrow = $vs.Shapes.AddShape(13, 690, 60, 120, 50)
$arrow.Fill.ForeColor.RGB = 255
$arrow.TextFrame2.TextRange.Text = 'Arrow'

$tb = $vs.Shapes.AddTextbox(1, 520, 150, 300, 60)
$tb.TextFrame2.TextRange.Text = 'Text box: shapes must survive PDF export'
$tb.TextFrame2.TextRange.Font.Size = 12
$tb.TextFrame2.TextRange.Font.Bold = -1
$tb.Line.Visible = 0
$tb.Fill.Visible = 0

# ---- sparklines on Data (openpyxl has no sparkline support) ----
# SparklineGroups.Add only works on the ACTIVE sheet, and Location must be a
# single cell (Excel fills down from SourceData row by row). Both are COM quirks.
$dt = $wb.Worksheets.Item('Data')
$dt.Activate()
# Add() rejects a multi-row Location/SourceData pair (0x800A03EC) in this build,
# so add one sparkline per row. Note ${i} braces: "$i:" would parse as a PS drive.
for ($i = 4; $i -le 23; $i++) {
  $dt.Range("I$i").SparklineGroups.Add(1, "Data!B${i}:E${i}") | Out-Null
}

# ---- recalc so cached values exist for any consumer ----
$xl.CalculateFullRebuild()
$wb.Save()

# ---- export PDF (worksheet-level export breaks on hidden app; use workbook) ----
$wb.ExportAsFixedFormat(0, $PdfOut)

$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
[GC]::Collect()

Write-Output ('PDF -> ' + $PdfOut)
Write-Output ('PDF size = ' + (Get-Item $PdfOut).Length)

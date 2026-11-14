# Excel 本体读取打印设置 + 导出 PDF（基准）
#   用法：powershell -File tools/excel-print-oracle.ps1
# 产出：
#   out/data/print-excel-oracle.json   Excel 自己报告的 PageSetup 值（判定基准）
#   out/pdf/excel-print.pdf            Excel 打印输出（效果层基准）
# 注意：Excel 的 PageSetup 只能在**已激活**的工作表上可靠读取，必须逐表 Activate。

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$sample = Join-Path $root 'samples\打印-PDF-多元素样本.xlsx'
$outJson = Join-Path $root 'out\data\print-excel-oracle.json'
$outPdfDir = Join-Path $root 'out\pdf'
if (-not (Test-Path $outPdfDir)) { New-Item -ItemType Directory -Path $outPdfDir | Out-Null }
$outPdf = Join-Path $outPdfDir 'excel-print.pdf'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($sample)

$rows = @()
foreach ($ws in $wb.Worksheets) {
    $ws.Activate()
    $ps = $ws.PageSetup
    $rows += [ordered]@{
        sheet              = $ws.Name
        # xlPortrait=1 / xlLandscape=2
        orientation        = [int]$ps.Orientation
        # xlPaperA4 = 9
        paperSize          = [int]$ps.PaperSize
        fitToPagesWide     = [int]$ps.FitToPagesWide
        fitToPagesTall     = [int]$ps.FitToPagesTall
        zoom               = [bool]$ps.Zoom
        marginTop          = [math]::Round([double]$ps.TopMargin, 4)
        marginBottom       = [math]::Round([double]$ps.BottomMargin, 4)
        marginLeft         = [math]::Round([double]$ps.LeftMargin, 4)
        marginRight        = [math]::Round([double]$ps.RightMargin, 4)
        marginHeader       = [math]::Round([double]$ps.HeaderMargin, 4)
        marginFooter       = [math]::Round([double]$ps.FooterMargin, 4)
        centerHorizontally = [bool]$ps.CenterHorizontally
        centerVertically   = [bool]$ps.CenterVertically
        printGridlines     = [bool]$ps.PrintGridlines
        headerLeft         = [string]$ps.LeftHeader
        headerCenter       = [string]$ps.CenterHeader
        headerRight        = [string]$ps.RightHeader
        footerLeft         = [string]$ps.LeftFooter
        footerCenter       = [string]$ps.CenterFooter
        footerRight        = [string]$ps.RightFooter
        printArea          = [string]$ps.PrintArea
        centerFooterNote   = ''
    }
}

# 打印整个工作簿 → PDF（效果层基准）
$wb.ExportAsFixedFormat(0, $outPdf)   # 0 = xlTypePDF

$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

$obj = [ordered]@{ sample = (Split-Path -Leaf $sample); pdf = $outPdf; sheets = $rows }
$obj | ConvertTo-Json -Depth 6 | Out-File -FilePath $outJson -Encoding utf8
Write-Output "OK $outJson"
Write-Output "PDF $outPdf"

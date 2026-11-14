param(
    [string]$RowsPath = "D:\workbuddy\spreadjs vs univer\poc-compare\_prec_rows.json",
    [string]$BookPath = "D:\workbuddy\spreadjs vs univer\公式对比-Excel365基准519函数-三方-v2.xlsx"
)

$ErrorActionPreference = 'Stop'
$rows = ConvertFrom-Json -InputObject (Get-Content -Raw -Encoding UTF8 -Path $RowsPath)
Write-Output ("rows loaded: " + $rows.Count)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $stub = $xl.Workbooks.Add()                # 先有打开的工作簿才能设置 Calculation
    $xl.Calculation = -4135                    # xlCalculationManual：打开目标簿不自动重算
    $wb = $xl.Workbooks.Open($BookPath)
    $stub.Close($false)
    Write-Output "chk: opened"

    $exists = $false
    for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
        if ($wb.Sheets.Item($i).Name -eq '精度验证-CALC_PRECISION') { $exists = $true }
    }
    if ($exists) { $wb.Sheets.Item('精度验证-CALC_PRECISION').Delete() }

    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
    $ws.Name = '精度验证-CALC_PRECISION'
    Write-Output "chk: sheet created"

    # ---- 标题与说明 ----
    $ws.Cells.Item(1,1) = 'SpreadJS 精度差异修复验证：GC.Spread.CalcEngine.Functions.CALC_PRECISION = 15'
    $ws.Cells.Item(1,1).Font.Bold = $true
    $ws.Cells.Item(1,1).Font.Size = 13

    $notes = @(
        '参数作用：将 SpreadJS 计算引擎的数值精度设为 15 位有效数字（对齐 Excel 365 的显示精度），在创建 Workbook 前设置。',
        '验证方法：21 条用例（19 个精度差异函数 + BASE/DOLLAR 两个非精度对照组）在 CALC_PRECISION=15 下重算，与 Excel 365 oracle 结果对比。',
        '判定口径：与「对比表」完全一致 —— 数值 1e-9 相对容差 / 文本精确比对 / 不支持（#NAME?）不比较。',
        '实测环境：SpreadJS 19.2.5（npm @grapecity-software，Playwright 无头）｜ Excel 365 基线 ｜ 2026-11-05。',
        '结论：12/19 修复（IM* 复数函数 12 个）；BESSELI/J/K、T.INV.2T/TINV 为算法级数值差异（设置后值完全不变）；IMCOT/IMTAN 为末位舍入差；对照组 BASE/DOLLAR 不变，证明该参数只影响数值精度、不影响格式/区域差异。'
    )
    for ($i = 0; $i -lt $notes.Count; $i++) {
        $ws.Cells.Item(2 + $i, 1) = $notes[$i]
        $ws.Cells.Item(2 + $i, 1).Font.Size = 9
    }
    $ws.Cells.Item(8,1) = '统计（单独汇总）：'
    $ws.Cells.Item(8,1).Font.Bold = $true
    $ws.Cells.Item(9,1) = '精度差异函数'
    $ws.Cells.Item(9,2) = 19
    $ws.Cells.Item(10,1) = '设置 CALC_PRECISION=15 后修复'
    $ws.Cells.Item(10,2) = 12
    $ws.Cells.Item(11,1) = '未修复（算法级/末位舍入）'
    $ws.Cells.Item(11,2) = 7
    $ws.Cells.Item(12,1) = 'SJS 一致数（464 →）'
    $ws.Cells.Item(12,2) = 476
    $ws.Cells.Item(13,1) = 'SJS 一致率（94.7% →）'
    $ws.Cells.Item(13,2) = 0.952
    $ws.Cells.Item(13,2).NumberFormat = '0.0%'
    $ws.Cells.Item(14,1) = 'SJS 支持但不一致数（21 →）'
    $ws.Cells.Item(14,2) = 9
    Write-Output "chk: notes written"

    # ---- 明细表 ----
    $hdr = @('函数名','用例公式','Excel 365 结果','SJS 默认结果','SJS 设置 CALC_PRECISION=15 后','默认判定','设置后判定','是否修复')
    $HDRROW = 17
    for ($c = 0; $c -lt $hdr.Count; $c++) {
        $cell = $ws.Cells.Item($HDRROW, 1 + $c)
        $cell.Value2 = $hdr[$c]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 15917535
    }
    $r = $HDRROW
    foreach ($row in $rows) {
        $r++
        $ws.Cells.Item($r, 1) = $row.name
        $ws.Cells.Item($r, 2) = $row.formula
        $ws.Cells.Item($r, 3) = [string]$row.excel
        $ws.Cells.Item($r, 4) = [string]$row.default
        $ws.Cells.Item($r, 5) = [string]$row.prec15
        $ws.Cells.Item($r, 6) = $row.j_default
        $ws.Cells.Item($r, 7) = $row.j_prec
        if ($row.j_default -ne '一致' -and $row.j_prec -eq '一致') {
            $ws.Cells.Item($r, 8) = '修复'
            $ws.Cells.Item($r, 8).Interior.Color = 13561798
        } elseif ($row.name -eq 'BASE' -or $row.name -eq 'DOLLAR') {
            $ws.Cells.Item($r, 8) = '对照（非精度差异）'
            $ws.Cells.Item($r, 8).Interior.Color = 16770698
        } else {
            $ws.Cells.Item($r, 8) = '未修复'
            $ws.Cells.Item($r, 8).Interior.Color = 13551615
        }
    }
    $ws.Columns.Item(2).ColumnWidth = 24
    $ws.Columns.Item(3).ColumnWidth = 40
    $ws.Columns.Item(4).ColumnWidth = 40
    $ws.Columns.Item(5).ColumnWidth = 40
    $ws.Columns.Item(6).ColumnWidth = 12
    $ws.Columns.Item(7).ColumnWidth = 12
    $ws.Columns.Item(8).ColumnWidth = 18
    Write-Output "chk: table written"

    $wb.Save()
    Write-Output "chk: save returned"

    $saved = ($ws.Cells.Item($HDRROW + 1, 1).Value2)
    Write-Output ("saved check A18: " + $saved)
    $wb.Close($false)
    Write-Output "closed"
} catch {
    Write-Output ("ERR: " + $_.Exception.Message + " @ " + $_.InvocationInfo.PositionMessage)
    throw
} finally {
    try { $xl.Quit() } catch { Write-Output ("quit err: " + $_.Exception.Message) }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

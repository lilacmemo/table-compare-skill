# 由 Excel 本体(COM 自动化)创建全部图表类型的工作簿
$ErrorActionPreference = 'Stop'
$out = "D:\workbuddy\spreadjs vs univer\Excel原生图表效果.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
$wb = $excel.Workbooks.Add()

# 图例位置常量
$xlBottom = -4107; $xlRight = -4152; $xlTop = -4160

# ---------- 通用函数 ----------
function Add-Data($ws, [int]$row, $headers, $rows) {
    $nR = $rows.Count + 1
    $nC = $headers.Count
    $arr = New-Object 'object[,]' $nR, $nC
    for ($j = 0; $j -lt $nC; $j++) { $arr[0, $j] = $headers[$j] }
    for ($i = 0; $i -lt $rows.Count; $i++) {
        for ($j = 0; $j -lt $nC; $j++) { $arr[($i + 1), $j] = $rows[$i][$j] }
    }
    $rng = $ws.Range($ws.Cells.Item($row, 1), $ws.Cells.Item($row + $nR - 1, $nC))
    $rng.Value2 = $arr
    $ws.Range($ws.Cells.Item($row, 1), $ws.Cells.Item($row, $nC)).Font.Bold = $true
    return ($row + $nR - 1)
}

function Add-Ch($ws, $type, $rng, $title, $catT, $valT, $legend, $labels, [int]$idx, [int]$baseTop) {
    $col = $idx % 3
    $row = [math]::Floor($idx / 3)
    $left = 12 + $col * 345
    $top = $baseTop + $row * 252
    $shape = $null
    try {
        $shape = $ws.Shapes.AddChart2(-1, $type, $left, $top, 330, 238)
    } catch {
        Write-Output "ADD2_FAIL type=$type : $($_.Exception.Message)"
        try {
            $shape = $ws.Shapes.AddChart2(-1, 51, $left, $top, 330, 238)
            $shape.Chart.ChartType = $type
        } catch {
            Write-Output "SKIP type=$type : $($_.Exception.Message)"
            return
        }
    }
    $ch = $shape.Chart
    try { $ch.SetSourceData($rng) } catch { Write-Output "SRC_FAIL type=$type : $($_.Exception.Message)" }
    try { $ch.HasTitle = $true; $ch.ChartTitle.Text = $title } catch {}
    if ($catT) { try { $ax = $ch.Axes(1); $ax.HasTitle = $true; $ax.AxisTitle.Text = $catT } catch {} }
    if ($valT) { try { $ax = $ch.Axes(2); $ax.HasTitle = $true; $ax.AxisTitle.Text = $valT } catch {} }
    if ($legend -ne 0) {
        try { $ch.HasLegend = $true; $ch.Legend.Position = $legend } catch {}
    } else {
        try { $ch.HasLegend = $false } catch {}
    }
    if ($labels) { try { $ch.ApplyDataLabels() } catch {} }
}

function Add-ChX($ws, $type, $rng, $title, $catT, $valT, $legend, $labels, [int]$idx, [int]$baseTop, [string]$mode) {
    $col = $idx % 3
    $row = [math]::Floor($idx / 3)
    $left = 12 + $col * 345
    $top = $baseTop + $row * 252
    $shape = $null; $ch = $null
    if ($mode -eq 'sourcedata') {
        # 先建空图绑定数据，再切换类型（曲面图/股价图的可靠路径）
        try {
            $co = $ws.ChartObjects().Add($left, $top, 330, 238)
            $ch = $co.Chart
            $ch.SetSourceData($rng)
            $ch.ChartType = $type
        } catch {
            Write-Output "SD_FAIL type=$type : $($_.Exception.Message)"
            try {
                $ws.Activate(); $rng.Select()
                $shape = $ws.Shapes.AddChart2(-1, $type, $left, $top, 330, 238)
                $ch = $shape.Chart
            } catch { Write-Output "SKIP type=$type : $($_.Exception.Message)"; return }
        }
    } else {
        # 新式图表：先选中数据区域再插入（模拟 Excel 界面操作）
        try {
            $ws.Activate()
            $rng.Select()
            $shape = $ws.Shapes.AddChart2(-1, $type, $left, $top, 330, 238)
            $ch = $shape.Chart
        } catch {
            Write-Output "SEL_FAIL type=$type : $($_.Exception.Message)"
            try {
                $co = $ws.ChartObjects().Add($left, $top, 330, 238)
                $ch = $co.Chart
                $ch.SetSourceData($rng)
                $ch.ChartType = $type
            } catch { Write-Output "SKIP type=$type : $($_.Exception.Message)"; return }
        }
    }
    try { $ch.HasTitle = $true; $ch.ChartTitle.Text = $title } catch {}
    if ($catT) { try { $ax = $ch.Axes(1); $ax.HasTitle = $true; $ax.AxisTitle.Text = $catT } catch {} }
    if ($valT) { try { $ax = $ch.Axes(2); $ax.HasTitle = $true; $ax.AxisTitle.Text = $valT } catch {} }
    if ($legend -ne 0) {
        try { $ch.HasLegend = $true; $ch.Legend.Position = $legend } catch {}
    } else {
        try { $ch.HasLegend = $false } catch {}
    }
    if ($labels) { try { $ch.ApplyDataLabels() } catch {} }
    try { if ($ch.SeriesCollection().Count -eq 0) { Write-Output "EMPTY type=$type" } } catch {}
}

function New-Sheet($name) {
    $ws = $wb.Worksheets.Add()
    $ws.Name = $name
    $ws.Cells.Item(1, 1) = "$name（Excel 本体生成）"
    $ws.Cells.Item(1, 1).Font.Bold = $true
    $ws.Cells.Item(1, 1).Font.Size = 12
    return $ws
}

$months = @('1月', '2月', '3月', '4月', '5月', '6月')
$salesRows = @(
    @('1月', 326, 287, 246), @('2月', 342, 301, 238), @('3月', 358, 296, 259),
    @('4月', 401, 318, 271), @('5月', 389, 335, 263), @('6月', 432, 349, 284)
)

# ========== 1. 柱形图 ==========
$ws = New-Sheet('柱形图')
$end = Add-Data $ws 3 @('月份', '华东', '华南', '华北') $salesRows
$src = $ws.Range($ws.Cells.Item(3, 1), $ws.Cells.Item($end, 4))
$base = ($end + 2) * 15
$specs = @(
    @(51,  '簇状柱形图', $true),
    @(52,  '堆积柱形图', $true),
    @(53,  '百分比堆积柱形图', $true),
    @(54,  '三维簇状柱形图', $true),
    @(55,  '三维堆积柱形图', $true),
    @(56,  '三维百分比堆积柱形图', $true),
    @(-4100, '三维柱形图', $true)
)
for ($i = 0; $i -lt $specs.Count; $i++) {
    Add-Ch $ws $specs[$i][0] $src $specs[$i][1] '月份' '销售额（万元）' $xlBottom $false $i $base
}

# ========== 2. 条形图与圆柱/圆锥/棱锥 ==========
$ws = New-Sheet('条形图与三维变体')
$end = Add-Data $ws 3 @('月份', '华东', '华南', '华北') $salesRows
$src = $ws.Range($ws.Cells.Item(3, 1), $ws.Cells.Item($end, 4))
$base = ($end + 2) * 15
$specs = @(
    @(57,  '簇状条形图', $true),
    @(58,  '堆积条形图', $true),
    @(59,  '百分比堆积条形图', $true),
    @(60,  '三维簇状条形图', $true),
    @(61,  '三维堆积条形图', $true),
    @(62,  '三维百分比堆积条形图', $true),
    @(92,  '簇状圆柱图', $true),
    @(99,  '簇状圆锥图', $true),
    @(106, '簇状棱锥图', $true)
)
for ($i = 0; $i -lt $specs.Count; $i++) {
    Add-Ch $ws $specs[$i][0] $src $specs[$i][1] '月份' '销售额（万元）' $xlBottom $false $i $base
}

# ========== 3. 折线图 ==========
$ws = New-Sheet('折线图')
$end = Add-Data $ws 3 @('月份', '华东', '华南', '华北') $salesRows
$src = $ws.Range($ws.Cells.Item(3, 1), $ws.Cells.Item($end, 4))
$base = ($end + 2) * 15
$specs = @(
    @(4,     '折线图', $true),
    @(65,    '带数据标记的折线图', $true),
    @(63,    '堆积折线图', $true),
    @(64,    '百分比堆积折线图', $true),
    @(66,    '带数据标记的堆积折线图', $true),
    @(67,    '带数据标记的百分比堆积折线图', $true),
    @(-4101, '三维折线图', $true)
)
for ($i = 0; $i -lt $specs.Count; $i++) {
    Add-Ch $ws $specs[$i][0] $src $specs[$i][1] '月份' '销售额（万元）' $xlBottom $false $i $base
}

# ========== 4. 饼图与圆环图 ==========
$ws = New-Sheet('饼图与圆环图')
$pieRows = @(@('产品A', 1860), @('产品B', 1420), @('产品C', 980), @('产品D', 640), @('产品E', 410))
$end = Add-Data $ws 3 @('产品', '年销售额（万元）') $pieRows
$src = $ws.Range($ws.Cells.Item(3, 1), $ws.Cells.Item($end, 2))
$base = ($end + 2) * 15
$specs = @(
    @(5,     '饼图', $true),
    @(69,    '分离型饼图', $true),
    @(-4102, '三维饼图', $true),
    @(70,    '分离型三维饼图', $true),
    @(68,    '复合饼图', $true),
    @(71,    '复合条饼图', $true),
    @(-4120, '圆环图', $true),
    @(80,    '分离型圆环图', $true)
)
for ($i = 0; $i -lt $specs.Count; $i++) {
    Add-Ch $ws $specs[$i][0] $src $specs[$i][1] $null $null $xlRight $true $i $base
}

# ========== 5. 散点图 ==========
$ws = New-Sheet('散点图')
$scRows = @(
    @(5, 42, 38), @(8, 58, 52), @(12, 82, 75), @(15, 95, 91),
    @(20, 128, 110), @(25, 150, 142), @(30, 171, 158)
)
$end = Add-Data $ws 3 @('广告费用（万元）', '渠道一销售额', '渠道二销售额') $scRows
$src = $ws.Range($ws.Cells.Item(3, 1), $ws.Cells.Item($end, 3))
$base = ($end + 2) * 15
$specs = @(
    @(-4169, '散点图（仅数据标记）'),
    @(74,    '带直线的散点图'),
    @(75,    '无数据标记的直线散点图'),
    @(72,    '带平滑线的散点图'),
    @(73,    '无数据标记的平滑线散点图')
)
for ($i = 0; $i -lt $specs.Count; $i++) {
    Add-Ch $ws $specs[$i][0] $src $specs[$i][1] '广告费用（万元）' '销售额（万元）' $xlBottom $false $i $base
}

# ========== 6. 面积图 ==========
$ws = New-Sheet('面积图')
$end = Add-Data $ws 3 @('月份', '华东', '华南', '华北') $salesRows
$src = $ws.Range($ws.Cells.Item(3, 1), $ws.Cells.Item($end, 4))
$base = ($end + 2) * 15
$specs = @(
    @(1,     '面积图', $true),
    @(76,    '堆积面积图', $true),
    @(77,    '百分比堆积面积图', $true),
    @(-4098, '三维面积图', $true),
    @(78,    '三维堆积面积图', $true),
    @(79,    '三维百分比堆积面积图', $true)
)
for ($i = 0; $i -lt $specs.Count; $i++) {
    Add-Ch $ws $specs[$i][0] $src $specs[$i][1] '月份' '销售额（万元）' $xlBottom $false $i $base
}

# ========== 7. 雷达图 / 气泡图 / 曲面图 ==========
$ws = New-Sheet('雷达气泡曲面图')
$radarRows = @(
    @('画质', 4.5, 3.9), @('音效', 3.8, 4.3), @('续航', 4.2, 3.6),
    @('价格', 3.5, 4.4), @('外观', 4.0, 3.7)
)
$end1 = Add-Data $ws 3 @('能力维度', '产品A', '产品B') $radarRows
$radarSrc = $ws.Range($ws.Cells.Item(3, 1), $ws.Cells.Item($end1, 3))
$bubRows = @(
    @('产品A', 24.5, 8.2, 1860), @('产品B', 18.2, 12.5, 1420), @('产品C', 12.8, 15.3, 980),
    @('产品D', 9.6, 6.8, 640), @('产品E', 6.3, 21.4, 410)
)
$end2 = Add-Data $ws ($end1 + 3) @('产品', '市场份额（%）', '销售增长率（%）', '销售额（万元）') $bubRows
$bubSrc = $ws.Range($ws.Cells.Item($end1 + 3, 2), $ws.Cells.Item($end2, 4))
$surfRows = @(
    @('华东', 326, 342, 358, 401, 389), @('华南', 287, 301, 296, 318, 335),
    @('华北', 246, 238, 259, 271, 263), @('西南', 173, 186, 192, 181, 205),
    @('东北', 152, 165, 171, 178, 184)
)
$end3 = Add-Data $ws ($end2 + 3) @('地区', '1月', '2月', '3月', '4月', '5月') $surfRows
$surfSrc = $ws.Range($ws.Cells.Item($end2 + 3, 1), $ws.Cells.Item($end3, 6))
$base = ($end3 + 2) * 15
Add-Ch $ws -4151 $radarSrc '雷达图' $null '评分' $xlBottom $false 0 $base
Add-Ch $ws 81    $radarSrc '带数据标记的雷达图' $null '评分' $xlBottom $false 1 $base
Add-Ch $ws 82    $radarSrc '填充雷达图' $null '评分' $xlBottom $false 2 $base
Add-Ch $ws 15    $bubSrc '气泡图（气泡大小=销售额）' '市场份额（%）' '销售增长率（%）' $xlBottom $false 3 $base
Add-Ch $ws 87    $bubSrc '三维气泡图' '市场份额（%）' '销售增长率（%）' $xlBottom $false 4 $base
Add-ChX $ws 83 $surfSrc '三维曲面图' '月份' '销售额（万元）' $xlRight $false 5 $base 'sourcedata'
Add-ChX $ws 84 $surfSrc '三维曲面图（线框）' '月份' '销售额（万元）' $xlRight $false 6 $base 'sourcedata'
Add-ChX $ws 85 $surfSrc '曲面俯视图（等高线）' '月份' '销售额（万元）' $xlRight $false 7 $base 'sourcedata'
Add-ChX $ws 86 $surfSrc '曲面俯视图（线框）' '月份' '销售额（万元）' $xlRight $false 8 $base 'sourcedata'

# ========== 8. 股价图 ==========
$ws = New-Sheet('股价图')
$stockBase = @(
    @('2026/9/21', 28.4, 29.1, 27.9, 28.8), @('2026/9/22', 28.8, 29.5, 28.5, 29.2),
    @('2026/9/23', 29.2, 30.2, 29.0, 29.9), @('2026/9/24', 29.9, 30.5, 29.3, 29.6),
    @('2026/9/25', 29.6, 30.1, 29.1, 30.0), @('2026/9/28', 30.0, 31.0, 29.8, 30.6)
)
$volRows = @(@('2026/9/21', 8200000), @('2026/9/22', 9100000), @('2026/9/23', 7600000),
             @('2026/9/24', 8800000), @('2026/9/25', 9500000), @('2026/9/28', 10200000))
# 块A：日期+开盘+最高+最低+收盘
$endA = Add-Data $ws 3 @('日期', '开盘', '最高', '最低', '收盘') $stockBase
$srcA = $ws.Range($ws.Cells.Item(3, 1), $ws.Cells.Item($endA, 5))
# 块B：日期+成交量+最高+最低+收盘
$rB = $endA + 3
$rowsB = @()
for ($i = 0; $i -lt 6; $i++) { $rowsB += ,@($stockBase[$i][0], $volRows[$i][1], $stockBase[$i][2], $stockBase[$i][3], $stockBase[$i][4]) }
$endB = Add-Data $ws $rB @('日期', '成交量', '最高', '最低', '收盘') $rowsB
$srcB = $ws.Range($ws.Cells.Item($rB, 1), $ws.Cells.Item($endB, 5))
# 块C：日期+成交量+开盘+最高+最低+收盘
$rC = $endB + 3
$rowsC = @()
for ($i = 0; $i -lt 6; $i++) { $rowsC += ,@($stockBase[$i][0], $volRows[$i][1], $stockBase[$i][1], $stockBase[$i][2], $stockBase[$i][3], $stockBase[$i][4]) }
$endC = Add-Data $ws $rC @('日期', '成交量', '开盘', '最高', '最低', '收盘') $rowsC
$srcC = $ws.Range($ws.Cells.Item($rC, 1), $ws.Cells.Item($endC, 6))
# 块D：日期+最高+最低+收盘（供 HLC 股价图使用）
$rD = $endC + 3
$rowsD = @()
for ($i = 0; $i -lt 6; $i++) { $rowsD += ,@($stockBase[$i][0], $stockBase[$i][2], $stockBase[$i][3], $stockBase[$i][4]) }
$endD = Add-Data $ws $rD @('日期', '最高', '最低', '收盘') $rowsD
$srcD = $ws.Range($ws.Cells.Item($rD, 1), $ws.Cells.Item($endD, 4))
$base = ($endD + 2) * 15
Add-ChX $ws 88 $srcD '股价图：最高-最低-收盘' '日期' '价格（元）' $xlBottom $false 0 $base 'sourcedata'
Add-ChX $ws 89 $srcA '股价图：开盘-最高-最低-收盘' '日期' '价格（元）' $xlBottom $false 1 $base 'sourcedata'
Add-ChX $ws 90 $srcB '股价图：成交量-最高-最低-收盘' '日期' '价格（元）' $xlBottom $false 2 $base 'sourcedata'
Add-ChX $ws 91 $srcC '股价图：成交量-开盘-最高-最低-收盘' '日期' '价格（元）' $xlBottom $false 3 $base 'sourcedata'

# ========== 9. 新式图表（Excel 2016+）==========
$ws = New-Sheet('新式图表')
$wfRows = @(
    @('期初库存', 520), @('采购入库', 340), @('销售出库', -410),
    @('运输损耗', -25), @('盘盈调整', 18), @('期末库存', 443)
)
$endWf = Add-Data $ws 3 @('项目', '金额（万元）') $wfRows
$srcWf = $ws.Range($ws.Cells.Item(3, 1), $ws.Cells.Item($endWf, 2))
$rH = $endWf + 3
$treeRows = @(
    @('电子', '手机', 860), @('电子', '平板', 420), @('家电', '冰箱', 530),
    @('家电', '洗衣机', 380), @('服饰', '男装', 260), @('服饰', '女装', 310)
)
$endTr = Add-Data $ws $rH @('类别', '子类', '销售额（万元）') $treeRows
$srcTree = $ws.Range($ws.Cells.Item($rH, 1), $ws.Cells.Item($endTr, 3))
$rHs = $endTr + 3
$scores = @(62,68,71,74,75,76,78,79,80,82,83,85,86,88,89,90,91,93,95,97,45,52,58,64,69,73,77,81,84,87,55,60,66,72,78,84,90,96,50,70)
$scoreArr = New-Object 'object[,]' $scores.Count, 1
for ($i = 0; $i -lt $scores.Count; $i++) { $scoreArr[$i, 0] = $scores[$i] }
$rngHs = $ws.Range($ws.Cells.Item($rHs, 1), $ws.Cells.Item($rHs + $scores.Count, 1))
$rngHs.Cells.Item(1, 1) = '成绩'
$rngHs.Offset(1, 0).Resize($scores.Count, 1).Value2 = $scoreArr
$endHs = $rHs + $scores.Count
$srcHs = $ws.Range($ws.Cells.Item($rHs, 1), $ws.Cells.Item($endHs, 1))
$rPa = $endHs + 3
$paretoRows = @(@('尺寸偏差', 45), @('划痕', 32), @('色差', 21), @('气泡', 14), @('变形', 8), @('其他', 5))
$endPa = Add-Data $ws $rPa @('缺陷类型', '数量') $paretoRows
$srcPa = $ws.Range($ws.Cells.Item($rPa, 1), $ws.Cells.Item($endPa, 2))
$rBw = $endPa + 3
$boxA = @(12,15,16,17,18,19,20,21,22,23,25,28); $boxB = @(8,10,12,14,16,18,20,23,26,29,32,35)
$bwRows = @()
for ($i = 0; $i -lt 12; $i++) { $bwRows += ,@($boxA[$i], $boxB[$i]) }
$endBw = Add-Data $ws $rBw @('A组得分', 'B组得分') $bwRows
$srcBw = $ws.Range($ws.Cells.Item($rBw, 1), $ws.Cells.Item($endBw, 2))
$rFu = $endBw + 3
$funnelRows = @(@('访问店铺', 10000), @('加入购物车', 4200), @('提交订单', 2600), @('支付成功', 1900), @('复购客户', 620))
$endFu = Add-Data $ws $rFu @('阶段', '人数') $funnelRows
$srcFu = $ws.Range($ws.Cells.Item($rFu, 1), $ws.Cells.Item($endFu, 2))
$base = ($endFu + 2) * 15
Add-ChX $ws 119 $srcWf   '瀑布图（库存变动）' $null '金额（万元）' 0 $false 0 $base 'select'
Add-ChX $ws 117 $srcTree '树状图（层级占比）' $null $null $xlRight $false 1 $base 'select'
Add-ChX $ws 120 $srcTree '旭日图（层级占比）' $null $null $xlRight $false 2 $base 'select'
Add-ChX $ws 118 $srcHs   '直方图（成绩分布）' '成绩' '频数' $xlBottom $false 3 $base 'select'
Add-ChX $ws 122 $srcPa   '排列图（帕累托图）' '缺陷类型' '数量' $xlBottom $false 4 $base 'select'
Add-ChX $ws 121 $srcBw   '箱形图（两组得分分布）' $null '得分' $xlBottom $false 5 $base 'select'
Add-ChX $ws 123 $srcFu   '漏斗图（转化流程）' $null '人数' $xlBottom $false 6 $base 'select'

# ========== 10. 组合图（双坐标轴）==========
$ws = New-Sheet('组合图')
$comboRows = @()
$targets = @(520, 560, 600, 640, 680, 720); $actuals = @(498, 585, 612, 601, 705, 748)
for ($i = 0; $i -lt 6; $i++) { $comboRows += ,@($months[$i], $targets[$i], $actuals[$i], [math]::Round($actuals[$i] / $targets[$i], 4)) }
$end = Add-Data $ws 3 @('月份', '目标销售额（万元）', '实际销售额（万元）', '达成率') $comboRows
$ws.Range($ws.Cells.Item(3, 4), $ws.Cells.Item($end, 4)).NumberFormat = '0%'
$src = $ws.Range($ws.Cells.Item(3, 1), $ws.Cells.Item($end, 4))
$base = ($end + 2) * 15
$co = $ws.ChartObjects().Add(12, $base, 500, 300)
$ch = $co.Chart
$ch.ChartType = 51
$ch.SetSourceData($src)
$ch.HasTitle = $true
$ch.ChartTitle.Text = '组合图：销售额（柱）+ 达成率（线，次坐标轴）'
try { $s3 = $ch.SeriesCollection(3); $s3.ChartType = 65; $s3.AxisGroup = 2 } catch {}
try { $ax = $ch.Axes(1); $ax.HasTitle = $true; $ax.AxisTitle.Text = '月份' } catch {}
try { $ax = $ch.Axes(2); $ax.HasTitle = $true; $ax.AxisTitle.Text = '销售额（万元）' } catch {}
try { $ax2 = $ch.Axes(2, 2); $ax2.HasTitle = $true; $ax2.AxisTitle.Text = '达成率'; $ax2.TickLabels.NumberFormat = '0%' } catch {}
try { $ch.HasLegend = $true; $ch.Legend.Position = $xlBottom } catch {}

# ---------- 收尾：列宽、校验、删除默认表、保存 ----------
foreach ($w in $wb.Worksheets) { try { $w.UsedRange.Columns.AutoFit() | Out-Null } catch {} }
foreach ($w in $wb.Worksheets) {
    Write-Output "SHEET $($w.Name): $($w.ChartObjects().Count) charts"
}
foreach ($w in @($wb.Worksheets)) {
    if ($w.Name -like 'Sheet*') { $w.Delete() }
}
$wb.SaveAs($out, 51)
$wb.Close($false)
Write-Output "DONE: $out"
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

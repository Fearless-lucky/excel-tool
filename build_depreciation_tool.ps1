$ErrorActionPreference = 'Stop'

function New-Color([int]$r, [int]$g, [int]$b) {
    return $r + 256 * $g + 65536 * $b
}

function Set-RangeStyle($range, $fill, $fontColor, $fontSize, $bold, $hAlign, $vAlign) {
    $range.Interior.Color = $fill
    $range.Font.Name = '微软雅黑'
    $range.Font.Color = $fontColor
    $range.Font.Size = [int][Math]::Round($fontSize)
    $range.Font.Bold = $bold
    $range.HorizontalAlignment = $hAlign
    $range.VerticalAlignment = $vAlign
}

function Set-LightBorder($range, $color) {
    $range.Borders.LineStyle = 1
    $range.Borders.Color = $color
    $range.Borders.Weight = 2
}

function Set-CardBorder($range, $color) {
    foreach ($edge in @(7, 8, 9, 10)) {
        $range.Borders.Item($edge).LineStyle = 1
        $range.Borders.Item($edge).Color = $color
        $range.Borders.Item($edge).Weight = 2
    }
}

function Add-ToolButton($sheet, $cellRange, $name, $caption, $macro, $fill, $white) {
    $area = $sheet.Range($cellRange)
    $shape = $sheet.Shapes.AddShape(5, $area.Left + 4, $area.Top + 3, $area.Width - 8, $area.Height - 6)
    $shape.Name = $name
    $shape.OnAction = "modDepreciation.$macro"
    $shape.Fill.ForeColor.RGB = $fill
    $shape.Line.Visible = 0
    $shape.Placement = 1
    $shape.AlternativeText = $caption
    try {
        $shape.TextFrame2.TextRange.Text = $caption
        $shape.TextFrame2.TextRange.Font.Name = '微软雅黑'
        $shape.TextFrame2.TextRange.Font.Size = 11
        $shape.TextFrame2.TextRange.Font.Bold = -1
        $shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $white
        $shape.TextFrame2.VerticalAnchor = 3
        $shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        $shape.TextFrame2.MarginLeft = 5
        $shape.TextFrame2.MarginRight = 5
        $shape.TextFrame2.MarginTop = 2
        $shape.TextFrame2.MarginBottom = 2
    } catch {
        $shape.TextFrame.Characters().Text = $caption
        $shape.TextFrame.Characters().Font.Name = '微软雅黑'
        $shape.TextFrame.Characters().Font.Size = 11
        $shape.TextFrame.Characters().Font.Bold = $true
        $shape.TextFrame.Characters().Font.Color = $white
        $shape.TextFrame.HorizontalAlignment = -4108
        $shape.TextFrame.VerticalAlignment = -4108
    }
    try {
        $shape.Shadow.Visible = -1
        $shape.Shadow.ForeColor.RGB = (New-Color 18 42 68)
        $shape.Shadow.Transparency = 0.8
        $shape.Shadow.OffsetX = 1
        $shape.Shadow.OffsetY = 2
    } catch {}
}

$toolPath = 'C:\Users\admin\Desktop\excel tool\月度折旧生成工具.xlsm'
$sourceDir = 'C:\Users\admin\Desktop\excel tool\VBA源码'
$backupPath = 'C:\tmp\月度折旧生成工具_v1.2升级前_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsm'
Copy-Item -LiteralPath $toolPath -Destination $backupPath -Force

$navy = New-Color 20 44 72
$blue = New-Color 45 105 190
$teal = New-Color 23 143 148
$green = New-Color 31 145 98
$slate = New-Color 88 103 122
$canvas = New-Color 242 246 250
$white = New-Color 255 255 255
$text = New-Color 35 49 65
$muted = New-Color 102 116 133
$line = New-Color 207 217 228
$paleBlue = New-Color 229 238 249
$paleTeal = New-Color 229 245 245
$paleGreen = New-Color 229 245 237
$inputBlue = New-Color 0 80 200
$altRow = New-Color 248 250 253

$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$app.ScreenUpdating = $false

try {
    $wb = $app.Workbooks.Open($toolPath)
    $ws = $wb.Worksheets.Item('操作首页')

    $isNewLayout = ([string]$ws.Range('C10').Value2 -eq '固定资产类别')
    $savedPath = $ws.Range('B5').Value2
    $savedSheet = $ws.Range('B6').Value2
    $savedPrev = $ws.Range('E6').Value2
    $savedCurr = $ws.Range('G6').Value2
    if ($isNewLayout) {
        $savedChanges = $ws.Range('A11:M20').Value2
    } else {
        $savedChanges = $ws.Range('A16:G25').Value2
    }
    $moduleText = [System.IO.File]::ReadAllText(
        (Join-Path $sourceDir 'modDepreciation.bas'),
        [System.Text.Encoding]::UTF8
    )
    $moduleText = [regex]::Replace($moduleText, '(?m)^Attribute VB_Name = "modDepreciation"\r?\n', '')
    $mainModule = $wb.VBProject.VBComponents.Item('modDepreciation')
    if ($mainModule.CodeModule.CountOfLines -gt 0) {
        $mainModule.CodeModule.DeleteLines(1, $mainModule.CodeModule.CountOfLines)
    }
    $mainModule.CodeModule.AddFromString($moduleText)

    foreach ($className in @('CChangeItem', 'CResultRow')) {
        try {
            $oldClass = $wb.VBProject.VBComponents.Item($className)
            $wb.VBProject.VBComponents.Remove($oldClass)
        } catch {}
        $classText = [System.IO.File]::ReadAllText(
            (Join-Path $sourceDir ($className + '.cls')),
            [System.Text.Encoding]::UTF8
        )
        $classBodyStart = $classText.IndexOf('Option Explicit')
        if ($classBodyStart -lt 0) { throw "类模块 $className 缺少 Option Explicit。" }
        $classComponent = $wb.VBProject.VBComponents.Add(2)
        $classComponent.Name = $className
        $classComponent.CodeModule.AddFromString($classText.Substring($classBodyStart))
    }

    $thisWorkbookModule = $wb.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
    if ($thisWorkbookModule.CountOfLines -gt 0) {
        $thisWorkbookModule.DeleteLines(1, $thisWorkbookModule.CountOfLines)
    }
    $thisWorkbookModule.AddFromString(@'
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    modDepreciation.InitializeHome
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    '导入路径、预览和状态均为本次会话临时信息，退出工具时无需保存。
    Me.Saved = True
    On Error GoTo 0
End Sub
'@)

    foreach ($shape in @($ws.Shapes)) { $shape.Delete() }
    try { $ws.Cells.Validation.Delete() } catch {}
    try { $ws.Cells.FormatConditions.Delete() } catch {}
    $ws.Cells.UnMerge()
    $ws.Cells.Clear()

    $ws.Cells.Font.Name = '微软雅黑'
    $ws.Cells.Font.Size = 9.5
    $ws.Range('A1:M27').Interior.Color = $canvas
    $ws.Range('A1:M27').Font.Color = $text
    $ws.Range('A1:M27').VerticalAlignment = -4108

    $widths = @{
        A = 12; B = 13; C = 16; D = 9; E = 8; F = 12; G = 10;
        H = 12; I = 12; J = 13; K = 14; L = 18; M = 15
    }
    foreach ($column in $widths.Keys) {
        $ws.Columns.Item($column).ColumnWidth = $widths[$column]
    }

    $heights = @{
        1 = 30; 2 = 17; 3 = 5; 4 = 22; 5 = 23; 6 = 23; 7 = 23; 8 = 5;
        9 = 22; 10 = 28; 11 = 18; 12 = 18; 13 = 18; 14 = 18; 15 = 18;
        16 = 18; 17 = 18; 18 = 18; 19 = 18; 20 = 18; 21 = 18; 22 = 5;
        23 = 22; 24 = 24; 25 = 24; 26 = 25; 27 = 17
    }
    foreach ($row in $heights.Keys) {
        $ws.Rows.Item([int]$row).RowHeight = $heights[$row]
    }

    # WPS 通过 COM 创建相邻合并单元格时会错误扩展合并范围，
    # 顶部信息区改为连续单元格布局，避免路径和规则被写入错误的合并区域。
    foreach ($merge in @(
        'A9:M9', 'A21:M21', 'A23:M23', 'A24:C25', 'D24:F25',
        'G24:I25', 'J24:M25', 'B26:M26', 'A27:M27'
    )) {
        $ws.Range($merge).Merge()
    }

    $ws.Range('A1').Value2 = '月度折旧生成工具'
    $ws.Range('A2').Value2 = '完全离线  ·  不修改上月原文件  ·  完整复制历史工作簿  ·  自动校验后生成'
    Set-RangeStyle $ws.Range('A1:M1') $navy $white 19 $true -4131 -4108
    $ws.Range('A1:M1').IndentLevel = 1
    Set-RangeStyle $ws.Range('A2:M2') $navy (New-Color 205 220 236) 9 $false -4131 -4108
    $ws.Range('A2:M2').IndentLevel = 2

    $ws.Range('A4').Value2 = '01  上月文件'
    $ws.Range('H4').Value2 = '02  折旧规则'
    foreach ($section in @('A4:G4', 'H4:M4', 'A9:M9', 'A23:M23')) {
        Set-RangeStyle $ws.Range($section) $paleBlue $navy 10.5 $true -4131 -4108
        $ws.Range($section).IndentLevel = 1
        $ws.Range($section).Borders.Item(9).LineStyle = 1
        $ws.Range($section).Borders.Item(9).Color = $blue
        $ws.Range($section).Borders.Item(9).Weight = 3
    }

    $ws.Range('A5').Value2 = '文件路径'
    $ws.Range('A6').Value2 = '最新月表'
    $ws.Range('D6').Value2 = '上月期间'
    $ws.Range('F6').Value2 = '本月期间'
    $ws.Range('A7').Value2 = '自动识别名称为 YYYY.MM 的最大月份；默认输出到源文件所在文件夹。'
    # WPS 可能把期间或工作表名作为数值返回；首页输入区统一按文本回填。
    $ws.Range('B5').NumberFormat = '@'
    $ws.Range('B6').NumberFormat = '@'
    $ws.Range('E6').NumberFormat = '@'
    $ws.Range('G6').NumberFormat = '@'
    $ws.Range('B5').Value2 = [string]$savedPath
    $ws.Range('B6').Value2 = [string]$savedSheet
    $ws.Range('E6').Value2 = [string]$savedPrev
    $ws.Range('G6').Value2 = [string]$savedCurr
    foreach ($label in @('A5', 'A6', 'D6', 'F6')) {
        Set-RangeStyle $ws.Range($label) $canvas $text 9.5 $true -4131 -4108
    }
    foreach ($input in @('B5:G5', 'B6:C6', 'E6', 'G6')) {
        Set-RangeStyle $ws.Range($input) $white $inputBlue 9.5 $false -4131 -4108
        Set-LightBorder $ws.Range($input) $line
    }
    Set-RangeStyle $ws.Range('A7:G7') $paleTeal $muted 8.5 $false -4131 -4108
    $ws.Range('A7:G7').IndentLevel = 1
    Set-CardBorder $ws.Range('A5:G7') $line

    $ws.Range('H5').Value2 = '新增起提'
    $ws.Range('H6').Value2 = '最低净值'
    $ws.Range('H7').Value2 = '报废/停用当月'
    $ws.Range('J5').Value2 = '当月开始计提'
    $ws.Range('J6').Value2 = '最低净值为 0'
    $ws.Range('J7').Value2 = '当月仍计提折旧'
    foreach ($label in @('H5:I5', 'H6:I6', 'H7:I7')) {
        Set-RangeStyle $ws.Range($label) $paleTeal $text 9 $true -4131 -4108
        $ws.Range($label).IndentLevel = 1
    }
    foreach ($input in @('J5:M5', 'J6:M6', 'J7:M7')) {
        Set-RangeStyle $ws.Range($input) $white $inputBlue 9.5 $true -4131 -4108
        Set-LightBorder $ws.Range($input) $line
        $ws.Range($input).IndentLevel = 1
    }
    Set-CardBorder $ws.Range('H5:M7') $line

    $ruleValidations = @(
        @('J5', '当月开始计提,次月开始计提', '决定新增资产从本月还是次月开始计提折旧。'),
        @('J6', '最低净值为 0,按预计净残值', '决定折旧最低保留 0 元，或保留每项资产预计净残值。'),
        @('J7', '当月仍计提折旧,当月不计提折旧', '决定报废或停用发生当月是否仍计提一次折旧。')
    )
    foreach ($item in $ruleValidations) {
        $cell = $ws.Range($item[0])
        $cell.Validation.Add(3, 1, 1, $item[1])
        $cell.Validation.IgnoreBlank = $false
        $cell.Validation.InCellDropdown = $true
        $cell.Validation.ShowInput = $true
        $cell.Validation.InputTitle = '规则说明'
        $cell.Validation.InputMessage = $item[2]
        $cell.Validation.ShowError = $true
        $cell.Validation.ErrorTitle = '请选择有效选项'
        $cell.Validation.ErrorMessage = '请从单元格右侧的下拉列表中选择。'
    }
    $ws.Range('J5').Value2 = '当月开始计提'
    $ws.Range('J6').Value2 = '最低净值为 0'
    $ws.Range('J7').Value2 = '当月仍计提折旧'

    $ws.Range('A9').Value2 = '03  本月资产变动（无变动时整表留空）'
    $ws.Range('A10').Value2 = '变动类型'
    $ws.Range('B10').Value2 = '财产编号'
    $ws.Range('C10').Value2 = '固定资产类别'
    $ws.Range('D10').Value2 = '计量单位'
    $ws.Range('E10').Value2 = '数量'
    $ws.Range('F10').Value2 = '原值'
    $ws.Range('G10').Value2 = '使用年限'
    $ws.Range('H10').Value2 = '预计净残值'
    $ws.Range('I10').Value2 = '月均折旧'
    $ws.Range('J10').Value2 = '购买日期'
    $ws.Range('K10').Value2 = '保存地点'
    $ws.Range('L10').Value2 = '使用部门及个人'
    $ws.Range('M10').Value2 = '说明'
    Set-RangeStyle $ws.Range('A10:M10') $navy $white 9 $true -4108 -4108
    Set-LightBorder $ws.Range('A10:M10') $navy
    $ws.Range('A10:M10').WrapText = $true
    Set-RangeStyle $ws.Range('A11:M20') $white $inputBlue 9 $false -4131 -4108
    Set-LightBorder $ws.Range('A11:M20') $line
    for ($r = 11; $r -le 20; $r++) {
        if (($r % 2) -eq 0) { $ws.Range("A${r}:M${r}").Interior.Color = $altRow }
    }

    if ($null -ne $savedChanges) {
        $rowLow = $savedChanges.GetLowerBound(0)
        $rowHigh = $savedChanges.GetUpperBound(0)
        $colLow = $savedChanges.GetLowerBound(1)
        $colHigh = $savedChanges.GetUpperBound(1)
        for ($sourceRow = $rowLow; $sourceRow -le $rowHigh; $sourceRow++) {
            for ($sourceCol = $colLow; $sourceCol -le $colHigh; $sourceCol++) {
                $savedValue = $savedChanges.GetValue($sourceRow, $sourceCol)
                if ($null -eq $savedValue -or [string]$savedValue -eq '') { continue }
                $targetRow = 11 + ($sourceRow - $rowLow)
                if ($isNewLayout) {
                    $targetCol = 1 + ($sourceCol - $colLow)
                } else {
                    $oldToNew = @(1, 2, 6, 7, 8, 9, 13)
                    $targetCol = $oldToNew[$sourceCol - $colLow]
                }
                $ws.Cells.Item($targetRow, $targetCol).Value2 = $savedValue
            }
        }
    }

    $ws.Range('A11:A20').HorizontalAlignment = -4108
    $ws.Range('B11:B20').NumberFormat = '@'
    $ws.Range('E11:E20').NumberFormat = '0.00'
    $ws.Range('F11:F20').NumberFormat = '#,##0.00;[Red]-#,##0.00;–'
    $ws.Range('G11:G20').NumberFormat = '0.00'
    $ws.Range('H11:I20').NumberFormat = '#,##0.00;[Red]-#,##0.00;–'
    $ws.Range('J11:J20').NumberFormat = 'yyyy-mm-dd'
    $ws.Range('M11:M20').WrapText = $true

    $ws.Range('A11:A20').Validation.Add(3, 1, 1, '无变动,新增,调整,报废,停用')
    $ws.Range('A11:A20').Validation.IgnoreBlank = $true
    $ws.Range('A11:A20').Validation.InCellDropdown = $true
    $ws.Range('A11:A20').Validation.ShowError = $true
    $ws.Range('A11:A20').Validation.ErrorTitle = '变动类型无效'
    $ws.Range('A11:A20').Validation.ErrorMessage = '请选择：无变动、新增、调整、报废或停用。'

    foreach ($rangeAddress in @('E11:E20', 'F11:F20', 'G11:G20', 'I11:I20')) {
        $rng = $ws.Range($rangeAddress)
        $rng.Validation.Add(2, 1, 5, '0')
        $rng.Validation.IgnoreBlank = $true
        $rng.Validation.ShowError = $true
        $rng.Validation.ErrorTitle = '数值无效'
        $rng.Validation.ErrorMessage = '请输入大于 0 的数字，或留空。'
    }
    $ws.Range('H11:H20').Validation.Add(2, 1, 7, '0')
    $ws.Range('H11:H20').Validation.IgnoreBlank = $true
    $ws.Range('H11:H20').Validation.ShowError = $true
    $ws.Range('H11:H20').Validation.ErrorTitle = '数值无效'
    $ws.Range('H11:H20').Validation.ErrorMessage = '请输入不小于 0 的数字，或留空。'

    $ws.Range('A21').Value2 = '新增：类别、单位、原值、年限、购买日期、地点、部门必填，数量空白默认 1；月均折旧空白时自动计算。调整：只填需更新项。'
    Set-RangeStyle $ws.Range('A21:M21') $paleTeal $muted 8.5 $false -4131 -4108
    $ws.Range('A21:M21').IndentLevel = 1
    Set-CardBorder $ws.Range('A10:M21') $line

    $ws.Range('A23').Value2 = '04  执行操作'
    Add-ToolButton $ws 'A24:C25' 'btnSelectSource' '选择上月折旧表' 'SelectSource' $blue $white
    Add-ToolButton $ws 'D24:F25' 'btnCheckPreview' '检查并预览' 'CheckAndPreview' $teal $white
    Add-ToolButton $ws 'G24:I25' 'btnGenerateMonth' '生成本月折旧表' 'GenerateMonth' $green $white
    Add-ToolButton $ws 'J24:M25' 'btnClearAll' '清空' 'ClearAll' $slate $white

    $ws.Range('A26').Value2 = '当前状态'
    $ws.Range('B26').Value2 = '请先选择上月折旧表。'
    Set-RangeStyle $ws.Range('A26') $navy $white 9.5 $true -4108 -4108
    Set-RangeStyle $ws.Range('B26:M26') $paleGreen $green 9.5 $true -4131 -4108
    Set-LightBorder $ws.Range('A26:M26') (New-Color 194 224 207)
    $ws.Range('B26:M26').IndentLevel = 1
    $ws.Range('B26:M26').WrapText = $true

    $ws.Range('A27').Value2 = '隐私说明：工具完全离线，不联网、不上传数据；上月原文件保持不变，结果另存为新的 .xlsx 文件。'
    Set-RangeStyle $ws.Range('A27:M27') $canvas $muted 8 $false -4108 -4108

    $ws.Tab.Color = $blue
    $ws.PageSetup.PrintArea = '$A$1:$M$27'
    $ws.PageSetup.Orientation = 2
    $ws.PageSetup.Zoom = $false
    $ws.PageSetup.FitToPagesWide = 1
    $ws.PageSetup.FitToPagesTall = 1
    $ws.PageSetup.LeftMargin = $app.InchesToPoints(0.2)
    $ws.PageSetup.RightMargin = $app.InchesToPoints(0.2)
    $ws.PageSetup.TopMargin = $app.InchesToPoints(0.25)
    $ws.PageSetup.BottomMargin = $app.InchesToPoints(0.25)

    $ws.Activate()
    try {
        $app.ActiveWindow.DisplayGridlines = $false
        $app.ActiveWindow.Zoom = 100
        $app.ActiveWindow.ScrollRow = 1
        $app.ActiveWindow.ScrollColumn = 1
        $app.ActiveWindow.FreezePanes = $false
    } catch {}

    try {
        $wb.BuiltinDocumentProperties.Item('Title').Value = '月度折旧生成工具'
        $wb.BuiltinDocumentProperties.Item('Comments').Value = '完全离线；不修改源文件；自动校验并输出完整月度折旧表。'
    } catch {}
    $wb.Save()
    $wb.Close($true)

    Write-Output "BUILT|$toolPath"
    Write-Output "BACKUP|$backupPath"
} finally {
    try { if ($wb) { $wb.Close($false) } } catch {}
    $app.ScreenUpdating = $true
    $app.DisplayAlerts = $true
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
}

$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\native_excel_helper.ps1"

$tool = 'C:\Users\admin\Desktop\excel tool\月度折旧生成工具.xlsm'
$source = 'C:\Users\admin\Desktop\excel tool\折旧表_脱敏.xlsx'
$qaRoot = 'C:\tmp\depreciation_qa_excel16'

function Get-ExcelSnapshot {
    param(
        $ExcelApplication,
        [string]$Path
    )
    if (-not (Test-Path -LiteralPath $Path)) { return 'NONE' }
    $book = $ExcelApplication.Workbooks.Open($Path, 0, $true)
    try {
        $sheet = $book.Worksheets.Item('2026.02')
        $totalRow = 0
        for ($row = 3; $row -le 50; $row++) {
            if ([string]$sheet.Cells.Item($row, 1).Value2 -eq '合计') {
                $totalRow = $row
                break
            }
        }
        $rows = @()
        for ($row = 3; $row -lt $totalRow; $row++) {
            $rows += (
                [string]$sheet.Cells.Item($row, 3).Text + ':' +
                [string]$sheet.Cells.Item($row, 1).Text + ':' +
                [string]$sheet.Cells.Item($row, 4).Text + ':' +
                [string]$sheet.Cells.Item($row, 5).Value2 + ':' +
                [string]$sheet.Cells.Item($row, 8).Text + ':' +
                [string]$sheet.Cells.Item($row, 9).Text + ':' +
                [string]$sheet.Cells.Item($row, 11).Text + ':' +
                [string]$sheet.Cells.Item($row, 13).Value2 + ':' +
                [string]$sheet.Cells.Item($row, 14).Value2 + ':' +
                [string]$sheet.Cells.Item($row, 15).Value2
            )
        }
        return (
            'ROWS=' + ($rows -join ',') +
            '|TOTAL=' + $totalRow +
            '|NFORMULA=' + [string]$sheet.Cells.Item($totalRow, 14).Formula +
            '|OFORMULA=' + [string]$sheet.Cells.Item($totalRow, 15).Formula +
            '|FREEZE=' + [string]$ExcelApplication.ActiveWindow.FreezePanes +
            '|ORIENTATION=' + [string]$sheet.PageSetup.Orientation
        )
    } finally {
        $book.Close($false)
    }
}

function Invoke-ExcelScenario {
    param(
        $ExcelApplication,
        [string]$Name,
        [scriptblock]$Setup
    )

    $folder = Join-Path $qaRoot $Name
    $resolved = [System.IO.Path]::GetFullPath($folder)
    if (-not $resolved.StartsWith('C:\tmp\depreciation_qa_excel16\')) {
        throw "Unsafe QA path: $resolved"
    }
    New-Item -ItemType Directory -Path $resolved -Force | Out-Null
    $output = Join-Path $resolved '折旧表_2026.02.xlsx'
    if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output }

    $book = $ExcelApplication.Workbooks.Open($tool)
    try {
        $homeSheet = $book.Worksheets.Item('操作首页')
        $homeSheet.Range('A11:M20').ClearContents()
        $homeSheet.Range('J5').Value2 = '当月开始计提'
        $homeSheet.Range('J6').Value2 = '最低净值为 0'
        $homeSheet.Range('J7').Value2 = '当月仍计提折旧'

        $loaded = $ExcelApplication.Run("'月度折旧生成工具.xlsm'!LoadSourceFile", $source)
        & $Setup $homeSheet
        $previewErrors = $ExcelApplication.Run("'月度折旧生成工具.xlsm'!RunPreview", $false)
        $generated = $ExcelApplication.Run("'月度折旧生成工具.xlsm'!GenerateCore", $resolved, $false)
        $status = $homeSheet.Range('B26').Text
    } finally {
        $book.Close($false)
    }

    $snapshot = Get-ExcelSnapshot -ExcelApplication $ExcelApplication -Path $output
    Write-Output (
        'EXCEL16|' + $Name +
        '|LOAD=' + $loaded +
        '|PREVIEW_ERRORS=' + $previewErrors +
        '|GENERATED=' + $generated +
        '|OUTPUT=' + (Test-Path -LiteralPath $output) +
        '|STATUS=' + $status +
        '|' + $snapshot
    )
}

New-Item -ItemType Directory -Path $qaRoot -Force | Out-Null
$excelApplication = Get-NativeExcelApplication -TitleLike '*'
$oldAlerts = $excelApplication.DisplayAlerts
$oldScreen = $excelApplication.ScreenUpdating
$oldSecurity = $excelApplication.AutomationSecurity

try {
    $excelApplication.DisplayAlerts = $false
    $excelApplication.ScreenUpdating = $false
    $excelApplication.AutomationSecurity = 1

    if ($null -ne $excelApplication.ActiveWorkbook) {
        if ($excelApplication.ActiveWorkbook.Name -eq '月度折旧生成工具.xlsm') {
            $excelApplication.ActiveWorkbook.Close($false)
        }
    }

    Invoke-ExcelScenario $excelApplication 'baseline' {
        param($homeSheet)
    }

    Invoke-ExcelScenario $excelApplication 'new_complete_auto_dep' {
        param($homeSheet)
        $homeSheet.Range('J6').Value2 = '按预计净残值'
        $homeSheet.Range('A11').Value2 = '新增'
        $homeSheet.Range('B11').NumberFormat = '@'
        $homeSheet.Range('B11').Value2 = '0002'
        $homeSheet.Range('C11').Value2 = '通用设备'
        $homeSheet.Range('D11').Value2 = '台'
        $homeSheet.Range('E11').Value2 = 2
        $homeSheet.Range('F11').Value2 = 12000
        $homeSheet.Range('G11').Value2 = 5
        $homeSheet.Range('H11').Value2 = 1200
        $homeSheet.Range('J11').Value2 = '2026-02-01'
        $homeSheet.Range('K11').Value2 = '设备间'
        $homeSheet.Range('L11').Value2 = '综合部/张三'
        $homeSheet.Range('M11').Value2 = '自动计算月均折旧'
    }

    Invoke-ExcelScenario $excelApplication 'new_next_month' {
        param($homeSheet)
        $homeSheet.Range('J5').Value2 = '次月开始计提'
        $homeSheet.Range('A11').Value2 = '新增'
        $homeSheet.Range('B11').NumberFormat = '@'
        $homeSheet.Range('B11').Value2 = '0003'
        $homeSheet.Range('C11').Value2 = '办公设备'
        $homeSheet.Range('D11').Value2 = '台'
        $homeSheet.Range('F11').Value2 = 6000
        $homeSheet.Range('G11').Value2 = 3
        $homeSheet.Range('J11').Value2 = '2026年2月'
        $homeSheet.Range('K11').Value2 = '办公室'
        $homeSheet.Range('L11').Value2 = '办公室/李四'
    }

    Invoke-ExcelScenario $excelApplication 'adjust_full_fields' {
        param($homeSheet)
        $homeSheet.Range('J6').Value2 = '按预计净残值'
        $homeSheet.Range('A11').Value2 = '调整'
        $homeSheet.Range('B11').NumberFormat = '@'
        $homeSheet.Range('B11').Value2 = '0001'
        $homeSheet.Range('C11').Value2 = '房屋及构筑物（调整）'
        $homeSheet.Range('D11').Value2 = '栋'
        $homeSheet.Range('E11').Value2 = 1
        $homeSheet.Range('G11').Value2 = 45
        $homeSheet.Range('H11').Value2 = 10000
        $homeSheet.Range('I11').Value2 = 100
        $homeSheet.Range('J11').Value2 = '1988年1月'
        $homeSheet.Range('K11').Value2 = '新址'
        $homeSheet.Range('L11').Value2 = '资产部/王五'
        $homeSheet.Range('M11').Value2 = '资料及折旧参数调整'
    }

    Invoke-ExcelScenario $excelApplication 'scrap_no_dep' {
        param($homeSheet)
        $homeSheet.Range('J7').Value2 = '当月不计提折旧'
        $homeSheet.Range('A11').Value2 = '报废'
        $homeSheet.Range('B11').NumberFormat = '@'
        $homeSheet.Range('B11').Value2 = '0001'
        $homeSheet.Range('M11').Value2 = '本月报废'
    }

    Invoke-ExcelScenario $excelApplication 'duplicate_asset_blocked' {
        param($homeSheet)
        $homeSheet.Range('A11').Value2 = '新增'
        $homeSheet.Range('B11').NumberFormat = '@'
        $homeSheet.Range('B11').Value2 = '0001'
        $homeSheet.Range('C11').Value2 = '通用设备'
        $homeSheet.Range('D11').Value2 = '台'
        $homeSheet.Range('E11').Value2 = 1
        $homeSheet.Range('F11').Value2 = 5000
        $homeSheet.Range('G11').Value2 = 5
        $homeSheet.Range('H11').Value2 = 0
        $homeSheet.Range('J11').Value2 = '2026-02-01'
        $homeSheet.Range('K11').Value2 = '办公室'
        $homeSheet.Range('L11').Value2 = '办公室'
    }

    Invoke-ExcelScenario $excelApplication 'missing_required_blocked' {
        param($homeSheet)
        $homeSheet.Range('A11').Value2 = '新增'
        $homeSheet.Range('B11').NumberFormat = '@'
        $homeSheet.Range('B11').Value2 = '0099'
        $homeSheet.Range('F11').Value2 = 5000
        $homeSheet.Range('G11').Value2 = 5
    }
} finally {
    $excelApplication.AutomationSecurity = $oldSecurity
    $excelApplication.ScreenUpdating = $oldScreen
    $excelApplication.DisplayAlerts = $oldAlerts
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApplication)
}

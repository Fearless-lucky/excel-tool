$ErrorActionPreference = 'Stop'

$tool = 'C:\Users\admin\Desktop\excel tool\月度折旧生成工具.xlsm'
$source = 'C:\Users\admin\Desktop\excel tool\折旧表_脱敏.xlsx'
$qaRoot = 'C:\tmp\depreciation_qa_v12'

function Get-OutputSnapshot([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return 'NONE' }
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    try {
        $wb = $app.Workbooks.Open($path, 0, $true)
        $ws = $wb.Worksheets.Item('2026.02')
        $totalRow = 0
        for ($r = 3; $r -le 50; $r++) {
            if ([string]$ws.Cells.Item($r, 1).Value2 -eq '合计') {
                $totalRow = $r
                break
            }
        }
        $rows = @()
        for ($r = 3; $r -lt $totalRow; $r++) {
            $rows += (
                [string]$ws.Cells.Item($r, 3).Text + ':' +
                [string]$ws.Cells.Item($r, 1).Text + ':' +
                [string]$ws.Cells.Item($r, 4).Text + ':' +
                [string]$ws.Cells.Item($r, 5).Value2 + ':' +
                [string]$ws.Cells.Item($r, 8).Text + ':' +
                [string]$ws.Cells.Item($r, 9).Text + ':' +
                [string]$ws.Cells.Item($r, 11).Text + ':' +
                [string]$ws.Cells.Item($r, 13).Value2 + ':' +
                [string]$ws.Cells.Item($r, 14).Value2 + ':' +
                [string]$ws.Cells.Item($r, 15).Value2
            )
        }
        $snapshot = 'ROWS=' + ($rows -join ',') +
            '|TOTAL=' + $totalRow +
            '|NFORMULA=' + [string]$ws.Cells.Item($totalRow, 14).Formula +
            '|OFORMULA=' + [string]$ws.Cells.Item($totalRow, 15).Formula +
            '|FREEZE=' + [string]$app.ActiveWindow.FreezePanes +
            '|ORIENTATION=' + [string]$ws.PageSetup.Orientation
        $wb.Close($false)
        return $snapshot
    } finally {
        try { if ($wb) { $wb.Close($false) } } catch {}
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}

function Invoke-Scenario {
    param(
        [string]$Name,
        [scriptblock]$Setup
    )

    $folder = Join-Path $qaRoot $Name
    $resolved = [System.IO.Path]::GetFullPath($folder)
    if (-not $resolved.StartsWith('C:\tmp\depreciation_qa_v12\')) {
        throw "Unsafe QA path: $resolved"
    }
    New-Item -ItemType Directory -Path $resolved -Force | Out-Null
    $output = Join-Path $resolved '折旧表_2026.02.xlsx'
    if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output }

    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    try {
        try { $app.AutomationSecurity = 1 } catch {}
        $wb = $app.Workbooks.Open($tool)
        $homeSheet = $wb.Worksheets.Item('操作首页')
        $homeSheet.Range('A11:M20').ClearContents()
        $homeSheet.Range('J5').Value2 = '当月开始计提'
        $homeSheet.Range('J6').Value2 = '最低净值为 0'
        $homeSheet.Range('J7').Value2 = '当月仍计提折旧'

        $loaded = $app.Run("'月度折旧生成工具.xlsm'!LoadSourceFile", $source)
        & $Setup $homeSheet
        $previewErrors = $app.Run("'月度折旧生成工具.xlsm'!RunPreview", $false)
        $generated = $app.Run("'月度折旧生成工具.xlsm'!GenerateCore", $resolved, $false)
        $status = $homeSheet.Range('B26').Text
        $wb.Close($false)

        $snapshot = Get-OutputSnapshot $output
        Write-Output (
            'SCENARIO|' + $Name +
            '|LOAD=' + $loaded +
            '|PREVIEW_ERRORS=' + $previewErrors +
            '|GENERATED=' + $generated +
            '|OUTPUT=' + (Test-Path -LiteralPath $output) +
            '|STATUS=' + $status +
            '|' + $snapshot
        )
    } finally {
        try { if ($wb) { $wb.Close($false) } } catch {}
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}

New-Item -ItemType Directory -Path $qaRoot -Force | Out-Null

Invoke-Scenario -Name 'baseline' -Setup {
    param($homeSheet)
}

Invoke-Scenario -Name 'new_complete_auto_dep' -Setup {
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

Invoke-Scenario -Name 'new_next_month' -Setup {
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

Invoke-Scenario -Name 'adjust_full_fields' -Setup {
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

Invoke-Scenario -Name 'scrap_no_dep' -Setup {
    param($homeSheet)
    $homeSheet.Range('J7').Value2 = '当月不计提折旧'
    $homeSheet.Range('A11').Value2 = '报废'
    $homeSheet.Range('B11').NumberFormat = '@'
    $homeSheet.Range('B11').Value2 = '0001'
    $homeSheet.Range('M11').Value2 = '本月报废'
}

Invoke-Scenario -Name 'duplicate_asset_blocked' -Setup {
    param($homeSheet)
    $homeSheet.Range('A11').Value2 = '新增'
    $homeSheet.Range('B11').NumberFormat = '@'
    $homeSheet.Range('B11').Value2 = '0001'
    $homeSheet.Range('C11').Value2 = '通用设备'
    $homeSheet.Range('D11').Value2 = '台'
    $homeSheet.Range('F11').Value2 = 5000
    $homeSheet.Range('G11').Value2 = 5
    $homeSheet.Range('J11').Value2 = '2026-02-01'
    $homeSheet.Range('K11').Value2 = '办公室'
    $homeSheet.Range('L11').Value2 = '办公室'
}

Invoke-Scenario -Name 'missing_required_blocked' -Setup {
    param($homeSheet)
    $homeSheet.Range('A11').Value2 = '新增'
    $homeSheet.Range('B11').NumberFormat = '@'
    $homeSheet.Range('B11').Value2 = '0099'
    $homeSheet.Range('F11').Value2 = 5000
    $homeSheet.Range('G11').Value2 = 5
}

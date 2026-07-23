$ErrorActionPreference = 'Stop'

$tool = 'C:\Users\admin\Desktop\excel tool\月度折旧生成工具.xlsm'
$pdf = 'C:\tmp\月度折旧生成工具_首页预览.pdf'
$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
try {
    try { $app.AutomationSecurity = 3 } catch {}
    $wb = $app.Workbooks.Open($tool, 0, $true)
    $ws = $wb.Worksheets.Item('操作首页')
    $ws.ExportAsFixedFormat(0, $pdf)
    $wb.Close($false)
    Write-Output $pdf
} finally {
    try { if ($wb) { $wb.Close($false) } } catch {}
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
}

$ErrorActionPreference = 'Stop'

$dir = 'C:\Users\admin\Desktop\excel tool\VBA源码'
$utf8 = New-Object System.Text.UTF8Encoding($false)

$module = [System.IO.File]::ReadAllText((Join-Path $dir 'modDepreciation.bas'), [System.Text.Encoding]::UTF8)
$module = [regex]::Replace($module, '(?m)^Attribute VB_Name = "modDepreciation"\r?\n', '')
[System.IO.File]::WriteAllText((Join-Path $dir 'modDepreciation_可直接复制.txt'), $module, $utf8)

foreach ($className in @('CChangeItem', 'CResultRow')) {
    $source = [System.IO.File]::ReadAllText((Join-Path $dir ($className + '.cls')), [System.Text.Encoding]::UTF8)
    $start = $source.IndexOf('Option Explicit')
    if ($start -lt 0) { throw "$className 缺少 Option Explicit。" }
    [System.IO.File]::WriteAllText(
        (Join-Path $dir ($className + '_可直接复制.txt')),
        $source.Substring($start),
        $utf8
    )
}

Write-Output 'VBA copy-friendly sources synchronized.'

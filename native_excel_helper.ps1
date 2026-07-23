if (-not ('NativeExcelBridge' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class NativeExcelBridge
{
    public delegate bool EnumWindowProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(
        IntPtr hwnd,
        uint dwObjectId,
        ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out object ppvObject);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(
        IntPtr hwndParent,
        EnumWindowProc callback,
        IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(
        IntPtr hwnd,
        StringBuilder className,
        int maxCount);

    public static IntPtr FindDescendantByClass(IntPtr root, string expectedClass)
    {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(root, delegate(IntPtr hwnd, IntPtr lParam)
        {
            var name = new StringBuilder(256);
            GetClassName(hwnd, name, name.Capacity);
            if (String.Equals(name.ToString(), expectedClass, StringComparison.OrdinalIgnoreCase))
            {
                found = hwnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
'@
}

function Get-NativeExcelApplication {
    param(
        [string]$TitleLike = '*'
    )

    $processes = @(
        Get-Process -Name EXCEL -ErrorAction Stop |
            Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like $TitleLike }
    )
    if ($processes.Count -ne 1) {
        throw "需要唯一的 Microsoft Excel 窗口，实际找到 $($processes.Count) 个。"
    }

    $iidDispatch = [guid]'00020400-0000-0000-C000-000000000046'
    $nativeObject = $null
    $objIdNativeOm = [uint32]::MaxValue - 15
    $excelGridHandle = [NativeExcelBridge]::FindDescendantByClass(
        $processes[0].MainWindowHandle,
        'EXCEL7'
    )
    if ($excelGridHandle -eq [IntPtr]::Zero) {
        throw '未找到 Microsoft Excel 的 EXCEL7 工作表窗口。'
    }
    $hr = [NativeExcelBridge]::AccessibleObjectFromWindow(
        $excelGridHandle,
        $objIdNativeOm,
        [ref]$iidDispatch,
        [ref]$nativeObject
    )
    if ($hr -ne 0 -or $null -eq $nativeObject) {
        throw ('无法取得 Microsoft Excel 对象模型，HRESULT=0x{0:X8}' -f $hr)
    }
    return $nativeObject.Application
}

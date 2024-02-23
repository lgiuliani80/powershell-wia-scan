Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class Win32Apis
{
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(
        IntPtr hWnd,
        int nCmdShow
    );

    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
}
"@ -Language CSharp -PassThru | Out-Null

$consoleWindow = [Win32Apis]::GetConsoleWindow()
$SW_MINIMIZE = 6
[Win32Apis]::ShowWindow($consoleWindow, $SW_MINIMIZE)


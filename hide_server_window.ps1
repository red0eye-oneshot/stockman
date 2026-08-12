# 현재 실행 중인 Python 서버 창을 숨깁니다 (서버는 계속 실행)
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

$hidden = 0
$names = @("py","python","pythonw")
foreach ($name in $names) {
    $procs = Get-Process -Name $name -ErrorAction SilentlyContinue
    foreach ($p in $procs) {
        if ($p.MainWindowHandle -ne [IntPtr]::Zero) {
            [Win32]::ShowWindow($p.MainWindowHandle, 0) | Out-Null
            Write-Host "숨김 완료: $($p.Name) (PID: $($p.Id))"
            $hidden++
        }
    }
}
if ($hidden -eq 0) { Write-Host "숨길 Python 창이 없습니다." }
Start-Sleep -Seconds 2

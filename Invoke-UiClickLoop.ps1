# TABARC-Code
# Purpose:
#   Force a target UI window to foreground, then:
#   click (353,981) -> type "continue" -> click (842,977),
#   wait random 60–80 seconds, repeat for ~15 minutes.
#
# Notes:
# - Absolute pixel coordinates (primary display)
# - Target window must exist and be stable
# - Pixel automation is brittle by nature

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class Win32 {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
}
"@

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Mouse {
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")] public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

    public const int LEFTDOWN = 0x02;
    public const int LEFTUP   = 0x04;
}
"@

Add-Type -AssemblyName System.Windows.Forms

function Invoke-LeftClick {
    param([int]$X, [int]$Y)

    [Mouse]::SetCursorPos($X, $Y) | Out-Null
    Start-Sleep -Milliseconds 80
    [Mouse]::mouse_event([Mouse]::LEFTDOWN, 0, 0, 0, 0)
    Start-Sleep -Milliseconds 40
    [Mouse]::mouse_event([Mouse]::LEFTUP, 0, 0, 0, 0)
}

function Get-MainWindowHandleByProcessName {
    param([string]$ProcessName)

    $p = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } |
        Select-Object -First 1

    if (-not $p) { return [IntPtr]::Zero }
    return [IntPtr]$p.MainWindowHandle
}

function Set-ForegroundWindowSafe {
    param([IntPtr]$Hwnd)

    if ($Hwnd -eq [IntPtr]::Zero) { return $false }

    if ([Win32]::IsIconic($Hwnd)) {
        [Win32]::ShowWindow($Hwnd, 9) | Out-Null
    } else {
        [Win32]::ShowWindow($Hwnd, 5) | Out-Null
    }

    $fg = [Win32]::GetForegroundWindow()
    $nullPid = 0
    $fgTid = [Win32]::GetWindowThreadProcessId($fg, [ref]$nullPid)
    $curTid = [Win32]::GetCurrentThreadId()

    [Win32]::AttachThreadInput($curTid, $fgTid, $true) | Out-Null
    $ok = [Win32]::SetForegroundWindow($Hwnd)
    [Win32]::AttachThreadInput($curTid, $fgTid, $false) | Out-Null

    Start-Sleep -Milliseconds 150
    return $ok
}

# ---------- CONFIG ----------
$TargetProcessName = "notepad"   # change this

$FirstClick  = @{ X = 353; Y = 981 }
$SecondClick = @{ X = 842; Y = 977 }
$TextToType  = "continue"

$RunMinutes = 15
$WaitMinSeconds = 60
$WaitMaxSeconds = 80
# ----------------------------

Start-Sleep -Milliseconds 800

$sw = [System.Diagnostics.Stopwatch]::StartNew()

while ($sw.Elapsed.TotalMinutes -lt $RunMinutes) {

    if ($TargetProcessName) {
        $h = Get-MainWindowHandleByProcessName $TargetProcessName
        if ($h -eq [IntPtr]::Zero) { Start-Sleep 2; continue }
        if (-not (Set-ForegroundWindowSafe $h)) { Start-Sleep 2; continue }
    }

    Invoke-LeftClick @FirstClick
    Start-Sleep -Milliseconds 150

    [System.Windows.Forms.SendKeys]::SendWait($TextToType)
    Start-Sleep -Milliseconds 150

    Invoke-LeftClick @SecondClick

    Start-Sleep -Seconds (Get-Random -Minimum $WaitMinSeconds -Maximum ($WaitMaxSeconds + 1))
}

$sw.Stop()

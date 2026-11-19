# 실제 Windows 바탕화면에서 위젯의 창 순서 동작을 확인한다.
#   1) 평소에는 다른 창 아래에 깔려 있고, 클릭해도 앞으로 나오지 않는다
#   2) Win+D(바탕화면 보기) 상태에서는 바탕화면 위에 보인다
#   3) 바탕화면 보기를 끝내면 다시 다른 창 아래로 내려간다
#   4) 바탕화면 보기 중에 타일을 눌러 연 프로그램을 가리지 않는다
# 사용: pwsh scripts/smoke-desktop.ps1 -Exe <deskboard.exe 경로> [-Screenshot <png 경로>]
param(
    [Parameter(Mandatory)][string]$Exe,
    [string]$Screenshot = ""
)
$ErrorActionPreference = "Stop"

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class Win {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindowW(string cls, string title);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT p);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
    [DllImport("user32.dll")] public static extern IntPtr GetTopWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] public static extern IntPtr GetWindowLongPtr(IntPtr h, int index);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern IntPtr GetShellWindow();
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
    public static string Class(IntPtr h) { var s = new StringBuilder(256); GetClassNameW(h, s, 256); return s.ToString(); }
    public static string Title(IntPtr h) { var s = new StringBuilder(256); GetWindowTextW(h, s, 256); return s.ToString(); }
    public static bool TopMost(IntPtr h) { return (GetWindowLongPtr(h, -20).ToInt64() & 0x8) != 0; }
    public static uint Pid(IntPtr h) { uint pid; GetWindowThreadProcessId(h, out pid); return pid; }
    public static IntPtr RootAt(int x, int y) { return GetAncestor(WindowFromPoint(new POINT { X = x, Y = y }), 2); }
    public static void Click(int x, int y) { SetCursorPos(x, y); mouse_event(2, 0, 0, 0, UIntPtr.Zero); mouse_event(4, 0, 0, 0, UIntPtr.Zero); }
}
'@

[Win]::SetProcessDPIAware() | Out-Null
$script:failures = 0
$started = @()

function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name $detail" }
    else { Write-Host "FAIL  $name $detail"; $script:failures++ }
}

function Describe([IntPtr]$h) {
    if ($h -eq [IntPtr]::Zero) { return "(none)" }
    "0x{0:X} [{1}] '{2}'" -f $h.ToInt64(), [Win]::Class($h), [Win]::Title($h)
}

# 조건이 참이 될 때까지 기다린다. 끝내 안 되면 false.
function WaitFor([scriptblock]$condition, [double]$seconds = 10) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        if (& $condition) { return $true }
        Start-Sleep -Milliseconds 100
    }
    return [bool](& $condition)
}

# 위에서 아래로 창 순서를 적는다. 판단에 필요한 창만 남긴다.
function DumpOrder([string]$label) {
    Write-Host "--- window order: $label (foreground: $(Describe ([Win]::GetForegroundWindow())))"
    $shellPid = [Win]::Pid([Win]::GetShellWindow())
    $h = [Win]::GetTopWindow([IntPtr]::Zero)
    $index = 0
    while ($h -ne [IntPtr]::Zero -and $index -lt 3000) {
        $class = [Win]::Class($h)
        $title = [Win]::Title($h)
        $visible = [Win]::IsWindowVisible($h)
        if ($visible -or $class -in "WorkerW", "Progman" -or $title -like "DeskBoard*") {
            $r = New-Object Win+RECT
            [Win]::GetWindowRect($h, [ref]$r) | Out-Null
            $flags = @()
            if ($visible) { $flags += "visible" }
            if ([Win]::TopMost($h)) { $flags += "topmost" }
            if ([Win]::IsIconic($h)) { $flags += "minimized" }
            if ([Win]::Pid($h) -eq $shellPid) { $flags += "shell" }
            Write-Host ("  #{0,-4} 0x{1:X8} {2,-28} {3,-22} {4}x{5}@{6},{7} {8}" -f $index, $h.ToInt64(), $class, "'$title'", ($r.R - $r.L), ($r.B - $r.T), $r.L, $r.T, ($flags -join ","))
        }
        $h = [Win]::GetWindow($h, 2)
        $index++
    }
}

# 위젯 자리를 덮는 평범한 창을 하나 띄운다.
function OpenCover([string]$title, [int]$x, [int]$y, [int]$w, [int]$h) {
    $code = "Add-Type -AssemblyName System.Windows.Forms; `$f = New-Object Windows.Forms.Form; `$f.Text = '$title'; " +
            "`$f.StartPosition = 'Manual'; `$f.SetBounds($x, $y, $w, $h); [Windows.Forms.Application]::Run(`$f)"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
    $process = Start-Process powershell.exe -PassThru -ArgumentList "-NoProfile", "-WindowStyle", "Hidden", "-EncodedCommand", $encoded
    $script:started += $process
    $found = WaitFor { $c = [Win]::FindWindowW($null, $title); $c -ne [IntPtr]::Zero -and [Win]::IsWindowVisible($c) } 30
    if (-not $found) { throw "cover window '$title' did not appear" }
    return [Win]::FindWindowW($null, $title)
}

# 창 순서에서 $a 가 $b 보다 위에 있는가
function IsAbove([IntPtr]$a, [IntPtr]$b) {
    $h = [Win]::GetTopWindow([IntPtr]::Zero)
    for ($i = 0; $h -ne [IntPtr]::Zero -and $i -lt 5000; $i++) {
        if ($h -eq $a) { return $true }
        if ($h -eq $b) { return $false }
        $h = [Win]::GetWindow($h, 2)
    }
    return $false
}

# 제목에 cmd.exe 가 들어간 콘솔 창 가운데 테스트 전에 없던 것
function FindNewConsole($known) {
    $h = [Win]::GetTopWindow([IntPtr]::Zero)
    for ($i = 0; $h -ne [IntPtr]::Zero -and $i -lt 5000; $i++) {
        if ([Win]::IsWindowVisible($h) -and [Win]::Title($h) -like "*cmd.exe*" -and $known -notcontains $h) { return $h }
        $h = [Win]::GetWindow($h, 2)
    }
    return [IntPtr]::Zero
}

function ConsoleWindows() {
    $list = @()
    $h = [Win]::GetTopWindow([IntPtr]::Zero)
    for ($i = 0; $h -ne [IntPtr]::Zero -and $i -lt 5000; $i++) {
        if ([Win]::Title($h) -like "*cmd.exe*") { $list += $h }
        $h = [Win]::GetWindow($h, 2)
    }
    return $list
}

try {
    Write-Host "OS: $((Get-CimInstance Win32_OperatingSystem).Caption) $([Environment]::OSVersion.Version)"
    $shell = [Win]::GetShellWindow()
    Write-Host "Shell window: $(Describe $shell)"
    if ($shell -eq [IntPtr]::Zero) { throw "no desktop shell in this session, so the desktop behaviour cannot be tested" }

    # 깨끗한 환경(CI)에서만 돈다. 쓰고 있는 보드를 덮어쓰지 않는다.
    $dataDir = Join-Path $env:APPDATA "com.koko8787.deskboard"
    $boardFile = Join-Path $dataDir "board.json"
    if (Test-Path $boardFile) { throw "$boardFile already exists; this test only runs on a machine without DeskBoard data" }
    New-Item -ItemType Directory -Force $dataDir | Out-Null
    $cmdPath = Join-Path $env:SystemRoot "System32\cmd.exe"
    $seed = @{
        version = 1
        categories = @(@{ id = "c1"; name = "Test"; tiles = @(@{ id = "t1"; name = "cmd"; target = $cmdPath; kind = "program" }) })
        notes = @(@{ id = "n1"; text = "smoke test"; x = 16; y = 16; w = 220; h = 160; color = "yellow"; z = 1 })
        settings = @{ split = 0.56; tileSize = "m"; opacity = 0.86; locked = $true }
    }
    $seed | ConvertTo-Json -Depth 8 | Set-Content -Path $boardFile -Encoding utf8

    $app = Start-Process -FilePath $Exe -PassThru
    $started += $app
    $appeared = WaitFor { $w = [Win]::FindWindowW($null, "DeskBoard"); $w -ne [IntPtr]::Zero -and [Win]::IsWindowVisible($w) } 60
    Check "widget window appears" $appeared
    if (-not $appeared) { DumpOrder "widget missing"; throw "widget window did not appear" }
    $widget = [Win]::FindWindowW($null, "DeskBoard")
    Start-Sleep -Seconds 4   # 화면이 그려질 시간

    $r = New-Object Win+RECT
    [Win]::GetWindowRect($widget, [ref]$r) | Out-Null
    $width = $r.R - $r.L
    $height = $r.B - $r.T
    Write-Host "Widget: $(Describe $widget) ${width}x${height} at $($r.L),$($r.T)"
    $exStyle = [Win]::GetWindowLongPtr($widget, -20).ToInt64()
    Check "widget is kept out of Alt+Tab and the taskbar" (($exStyle -band 0x80) -ne 0)
    Check "widget starts below other windows, not always-on-top" (-not [Win]::TopMost($widget))

    # 덮개 창은 위젯의 오른쪽 60% 를 덮는다. 왼쪽은 클릭해 볼 수 있게 남겨 둔다.
    $coverX = $r.L + [int]($width * 0.4)
    $probeX = $r.L + [int]($width * 0.75)
    $probeY = $r.T + [int]($height * 0.5)
    $cover = OpenCover "DeskBoardCover1" $coverX $r.T ($width - [int]($width * 0.4)) $height
    $atProbe = { [Win]::RootAt($probeX, $probeY) }

    # 1) 평소
    Check "normal: another window covers the widget" (WaitFor { (& $atProbe) -eq $cover } 5) (Describe (& $atProbe))
    [Win]::Click($r.L + [int]($width * 0.12), $r.B - 30)
    Start-Sleep -Milliseconds 800
    Write-Host "after clicking the widget, foreground: $(Describe ([Win]::GetForegroundWindow()))"
    Check "normal: clicking the widget does not bring it in front" ((& $atProbe) -eq $cover) (Describe (& $atProbe))
    DumpOrder "normal"

    # 2) 바탕화면 보기 (Win+D 와 같은 동작)
    $shellApp = New-Object -ComObject Shell.Application
    $shellApp.ToggleDesktop()
    Check "show desktop: widget is visible on the desktop" (WaitFor { (& $atProbe) -eq $widget } 10) (Describe (& $atProbe))
    Check "show desktop: widget is not minimized or hidden" ([Win]::IsWindowVisible($widget) -and -not [Win]::IsIconic($widget))
    DumpOrder "desktop shown"

    if ($Screenshot) {
        Add-Type -AssemblyName System.Drawing
        $bitmap = New-Object System.Drawing.Bitmap $width, $height
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($r.L, $r.T, 0, 0, $bitmap.Size)
        $bitmap.Save($Screenshot, [System.Drawing.Imaging.ImageFormat]::Png)
        # 넣어 둔 노란 메모가 그려졌는지로 화면이 제대로 떴는지 확인한다.
        $yellow = 0
        for ($y = 0; $y -lt $height; $y += 4) {
            for ($x = 0; $x -lt $width; $x += 4) {
                $c = $bitmap.GetPixel($x, $y)
                if ($c.R -ge 240 -and $c.G -ge 215 -and $c.G -le 245 -and $c.B -ge 110 -and $c.B -le 160) { $yellow++ }
            }
        }
        $graphics.Dispose(); $bitmap.Dispose()
        Check "show desktop: the yellow note is drawn" ($yellow -ge 300) "($yellow yellow samples)"
    }

    # 3) 바탕화면 보기 끝
    $shellApp.ToggleDesktop()
    Check "restored: other windows cover the widget again" (WaitFor { (& $atProbe) -eq $cover } 10) (Describe (& $atProbe))
    Check "restored: widget is no longer always-on-top" (WaitFor { -not [Win]::TopMost($widget) } 5)
    DumpOrder "restored"

    # 4) 바탕화면 보기 중에 위젯의 타일을 눌러 프로그램을 연다
    $shellApp.ToggleDesktop()
    Check "show desktop again: widget is visible" (WaitFor { (& $atProbe) -eq $widget } 10) (Describe (& $atProbe))
    $consolesBefore = ConsoleWindows
    # 첫 타일의 가운데쯤 (테두리 1 + 왼쪽 여백 12 / 헤더 36 + 여백 6 + 카테고리 이름 26)
    [Win]::Click($r.L + 48, $r.T + 100)
    $opened = WaitFor { (FindNewConsole $consolesBefore) -ne [IntPtr]::Zero } 15
    Check "tile click opens the program" $opened
    if ($opened) {
        $console = FindNewConsole $consolesBefore
        Write-Host "opened: $(Describe $console)"
        Check "opened program is in front of the widget" (WaitFor { IsAbove $console $widget } 10)
    }
    Write-Host "widget always-on-top after opening: $([Win]::TopMost($widget))"
    DumpOrder "program opened from the shown desktop"
    Get-Process cmd -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -gt $app.StartTime } | ForEach-Object { $script:started += $_ }

    Check "app is still running" (-not $app.HasExited)
}
catch {
    Write-Host "ERROR $_"
    $script:failures++
}
finally {
    foreach ($process in $started) {
        if ($process -and -not $process.HasExited) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue }
    }
}

if ($script:failures -gt 0) { Write-Host "`n$($script:failures) check(s) failed"; exit 1 }
Write-Host "`nAll checks passed"

param(
    [switch]$SelfTest
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$nativeCode = @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class NativeWin {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int command);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
"@

Add-Type -TypeDefinition $nativeCode

$script:ffmpegProcess = $null
$script:audioProcess = $null
$script:activeRecording = $null
$script:selectedWindow = $null
$script:savedBounds = @{}
$script:recordingsDir = Join-Path $PSScriptRoot "recordings"
$script:configPath = Join-Path $PSScriptRoot "settings.json"
$script:loopbackAudioSources = @{}
$script:loopbackAudioPrefix = "Loopback: "
$script:noAudioLabel = "No audio (video only)"

$HWND_TOP = [IntPtr]::Zero
$HWND_BOTTOM = [IntPtr]1
$SW_RESTORE = 9
$SW_SHOWNOACTIVATE = 4
$SWP_NOSIZE = 0x0001
$SWP_NOMOVE = 0x0002
$SWP_NOACTIVATE = 0x0010
$SWP_SHOWWINDOW = 0x0040

function Get-WindowTitle {
    param([IntPtr]$Hwnd)

    if (-not [NativeWin]::IsWindow($Hwnd)) {
        return ""
    }

    $length = [NativeWin]::GetWindowTextLength($Hwnd)
    if ($length -le 0) {
        return ""
    }

    $builder = New-Object System.Text.StringBuilder ($length + 1)
    [void][NativeWin]::GetWindowText($Hwnd, $builder, $builder.Capacity)
    return $builder.ToString().Trim()
}

function Get-ProcessNameFromWindow {
    param([IntPtr]$Hwnd)

    [uint32]$processIdValue = 0
    [void][NativeWin]::GetWindowThreadProcessId($Hwnd, [ref]$processIdValue)
    if ($processIdValue -le 0) {
        return ""
    }

    try {
        return (Get-Process -Id ([int]$processIdValue) -ErrorAction Stop).ProcessName
    }
    catch {
        return ""
    }
}

function Get-OpenWindows {
    $windows = New-Object System.Collections.ArrayList

    [NativeWin]::EnumWindows({
        param([IntPtr]$hWnd, [IntPtr]$lParam)

        if (-not [NativeWin]::IsWindowVisible($hWnd)) {
            return $true
        }

        $title = Get-WindowTitle -Hwnd $hWnd
        if ([string]::IsNullOrWhiteSpace($title)) {
            return $true
        }

        $rect = New-Object NativeWin+RECT
        [void][NativeWin]::GetWindowRect($hWnd, [ref]$rect)
        $width = $rect.Right - $rect.Left
        $height = $rect.Bottom - $rect.Top

        if ($width -lt 80 -or $height -lt 60) {
            return $true
        }

        $processName = Get-ProcessNameFromWindow -Hwnd $hWnd
        $display = "{0}    [{1}]    hwnd=0x{2:X}" -f $title, $processName, $hWnd.ToInt64()

        [void]$windows.Add([pscustomobject]@{
            Hwnd = $hWnd
            Title = $title
            ProcessName = $processName
            Left = $rect.Left
            Top = $rect.Top
            Width = $width
            Height = $height
            Display = $display
        })

        return $true
    }, [IntPtr]::Zero) | Out-Null

    return $windows | Sort-Object ProcessName, Title
}

function Get-DShowAudioDevices {
    $ffmpeg = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if (-not $ffmpeg) {
        return @()
    }

    $output = & $ffmpeg.Source -hide_banner -list_devices true -f dshow -i dummy 2>&1
    $devices = New-Object System.Collections.ArrayList
    foreach ($line in $output) {
        $text = $line.ToString()
        if ($text -match '"(.+)"\s+\(audio\)') {
            [void]$devices.Add($matches[1])
        }
    }

    return $devices | Sort-Object -Unique
}

function Get-PythonCommand {
    return Get-Command python -ErrorAction SilentlyContinue
}

function Test-LoopbackAudioAvailable {
    $python = Get-PythonCommand
    if (-not $python) {
        return $false
    }

    $probe = & $python.Source -c "import soundcard, numpy" 2>$null
    return $LASTEXITCODE -eq 0
}

function Test-FfmpegFilterAvailable {
    param([string]$FilterName)

    $ffmpeg = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if (-not $ffmpeg) {
        return $false
    }

    $probe = & $ffmpeg.Source -hide_banner -h ("filter=" + $FilterName) 2>&1
    return $LASTEXITCODE -eq 0 -and (($probe | Out-String) -match [regex]::Escape($FilterName))
}

function Get-LoopbackAudioSources {
    $python = Get-PythonCommand
    if (-not $python) {
        return @()
    }

    $scriptPath = Join-Path $PSScriptRoot "loopback_audio_recorder.py"
    if (-not (Test-Path $scriptPath)) {
        return @()
    }

    try {
        $json = & $python.Source $scriptPath --list-json 2>$null
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($json)) {
            return @()
        }
        $parsed = $json | ConvertFrom-Json
        foreach ($item in $parsed) {
            $item
        }
    }
    catch {
        return @()
    }
}

function Load-Settings {
    if (-not (Test-Path $script:configPath)) {
        return
    }

    try {
        $settings = Get-Content -LiteralPath $script:configPath -Raw | ConvertFrom-Json
        if ($settings.RecordingsDir -and (Test-Path $settings.RecordingsDir)) {
            $script:recordingsDir = [string]$settings.RecordingsDir
        }
    }
    catch {
    }
}

function Save-Settings {
    try {
        $settings = [pscustomobject]@{
            RecordingsDir = $script:recordingsDir
        }
        $settings | ConvertTo-Json | Set-Content -LiteralPath $script:configPath -Encoding UTF8
    }
    catch {
    }
}

function Get-UniqueOutputPath {
    param(
        [string]$Directory,
        [string]$Extension
    )

    $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $candidate = Join-Path $Directory ($stamp + $Extension)
    if (-not (Test-Path $candidate)) {
        return $candidate
    }

    for ($i = 2; $i -lt 1000; $i++) {
        $next = Join-Path $Directory ("{0}-{1:00}{2}" -f $stamp, $i, $Extension)
        if (-not (Test-Path $next)) {
            return $next
        }
    }

    return Join-Path $Directory ("{0}-{1}{2}" -f $stamp, [guid]::NewGuid().ToString("N").Substring(0, 8), $Extension)
}

function Set-MonitoringState {
    param(
        [bool]$Enabled,
        [System.Windows.Forms.TextBox]$LogBox
    )

    if ($script:audioProcess -and -not $script:audioProcess.HasExited) {
        try {
            if ($Enabled) {
                $script:audioProcess.StandardInput.WriteLine("monitor on")
            }
            else {
                $script:audioProcess.StandardInput.WriteLine("monitor off")
            }
            if ($LogBox) {
                $LogBox.AppendText(("Listening monitor: {0}" -f $(if ($Enabled) { "on" } else { "off" })) + [Environment]::NewLine)
            }
        }
        catch {
            if ($LogBox) {
                $LogBox.AppendText("Could not change listening monitor." + [Environment]::NewLine)
            }
        }
    }
}

function Quote-ProcessArg {
    param([string]$Arg)

    if ($null -eq $Arg -or $Arg.Length -eq 0) {
        return '""'
    }

    if ($Arg -notmatch '[\s"]') {
        return $Arg
    }

    $escaped = $Arg -replace '(\\*)"', '$1$1\"'
    $escaped = $escaped -replace '(\\+)$', '$1$1'
    return '"' + $escaped + '"'
}

function New-ArgumentString {
    param([System.Collections.IEnumerable]$Args)

    return ($Args | ForEach-Object { Quote-ProcessArg ([string]$_) }) -join " "
}

function Stop-ProcessNicely {
    param(
        [System.Diagnostics.Process]$Process,
        [int]$TimeoutMs = 5000
    )

    if ($null -eq $Process) {
        return
    }

    try {
        if (-not $Process.HasExited) {
            try {
                $Process.StandardInput.WriteLine("q")
            }
            catch {
            }

            if (-not $Process.WaitForExit($TimeoutMs)) {
                $Process.Kill()
                [void]$Process.WaitForExit(2000)
            }
        }
    }
    catch {
        try {
            if (-not $Process.HasExited) {
                $Process.Kill()
            }
        }
        catch {
        }
    }
}

function Start-LoopbackAudioRecorder {
    param(
        [string]$OutputPath,
        [string]$SourceName,
        [bool]$MonitorOn,
        [System.Windows.Forms.TextBox]$LogBox
    )

    $python = Get-PythonCommand
    if (-not $python) {
        [System.Windows.Forms.MessageBox]::Show("python was not found in PATH.", "Window Back Recorder") | Out-Null
        return $false
    }

    $scriptPath = Join-Path $PSScriptRoot "loopback_audio_recorder.py"
    if (-not (Test-Path $scriptPath)) {
        [System.Windows.Forms.MessageBox]::Show("loopback_audio_recorder.py was not found.", "Window Back Recorder") | Out-Null
        return $false
    }

    $args = New-Object System.Collections.ArrayList
    [void]$args.Add($scriptPath)
    [void]$args.Add($OutputPath)
    if (-not [string]::IsNullOrWhiteSpace($SourceName)) {
        [void]$args.Add("--source")
        [void]$args.Add($SourceName)
    }
    if ($MonitorOn) {
        [void]$args.Add("--monitor-on")
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $python.Source
    $psi.Arguments = New-ArgumentString $args
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.EnableRaisingEvents = $true

    $proc.add_OutputDataReceived({
        param($sender, $eventArgs)
        if ($eventArgs.Data -and $script:logBox) {
            try {
                [void]$script:logBox.BeginInvoke([Action]{
                    $script:logBox.AppendText("[audio] " + $eventArgs.Data + [Environment]::NewLine)
                })
            }
            catch {
            }
        }
    })

    $proc.add_ErrorDataReceived({
        param($sender, $eventArgs)
        if ($eventArgs.Data -and $script:logBox) {
            try {
                [void]$script:logBox.BeginInvoke([Action]{
                    $script:logBox.AppendText("[audio] " + $eventArgs.Data + [Environment]::NewLine)
                })
            }
            catch {
            }
        }
    })

    try {
        if (-not $proc.Start()) {
            return $false
        }
        $script:audioProcess = $proc
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()
        $LogBox.AppendText("Loopback audio started: $OutputPath" + [Environment]::NewLine)
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Could not start loopback audio") | Out-Null
        return $false
    }
}

function Invoke-MuxLoopbackRecording {
    param(
        [pscustomobject]$Recording,
        [System.Windows.Forms.TextBox]$LogBox
    )

    if ($null -eq $Recording -or $Recording.Mode -ne "loopback") {
        return
    }

    $ffmpeg = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if (-not $ffmpeg) {
        $LogBox.AppendText("Could not mux: ffmpeg not found." + [Environment]::NewLine)
        return
    }

    if (-not (Test-Path $Recording.VideoTemp)) {
        $LogBox.AppendText("Could not mux: video temp file missing." + [Environment]::NewLine)
        return
    }

    if (-not (Test-Path $Recording.AudioTemp)) {
        $LogBox.AppendText("Could not mux: audio temp file missing." + [Environment]::NewLine)
        return
    }

    $args = New-Object System.Collections.ArrayList
    [void]$args.Add("-hide_banner")
    [void]$args.Add("-y")
    [void]$args.Add("-i")
    [void]$args.Add($Recording.VideoTemp)
    [void]$args.Add("-i")
    [void]$args.Add($Recording.AudioTemp)
    [void]$args.Add("-map")
    [void]$args.Add("0:v:0")
    [void]$args.Add("-map")
    [void]$args.Add("1:a:0")
    [void]$args.Add("-c:v")
    [void]$args.Add("copy")
    [void]$args.Add("-c:a")
    [void]$args.Add("aac")
    [void]$args.Add("-b:a")
    [void]$args.Add("160k")
    [void]$args.Add("-shortest")
    [void]$args.Add($Recording.FinalOutput)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ffmpeg.Source
    $psi.Arguments = New-ArgumentString $args
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true

    $LogBox.AppendText("Muxing audio and video..." + [Environment]::NewLine)
    $proc = [System.Diagnostics.Process]::Start($psi)
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if ($proc.ExitCode -eq 0) {
        $LogBox.AppendText("Final file: " + $Recording.FinalOutput + [Environment]::NewLine)
        try {
            Remove-Item -LiteralPath $Recording.VideoTemp -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $Recording.AudioTemp -ErrorAction SilentlyContinue
        }
        catch {
        }
    }
    else {
        $LogBox.AppendText("Mux failed with code " + $proc.ExitCode + [Environment]::NewLine)
        if ($stderr) {
            $LogBox.AppendText($stderr + [Environment]::NewLine)
        }
        $LogBox.AppendText("Temp video: " + $Recording.VideoTemp + [Environment]::NewLine)
        $LogBox.AppendText("Temp audio: " + $Recording.AudioTemp + [Environment]::NewLine)
    }
}

function Save-WindowBounds {
    param([IntPtr]$Hwnd)

    if (-not [NativeWin]::IsWindow($Hwnd)) {
        return
    }

    $rect = New-Object NativeWin+RECT
    [void][NativeWin]::GetWindowRect($Hwnd, [ref]$rect)
    $script:savedBounds[$Hwnd.ToInt64()] = [pscustomobject]@{
        Left = $rect.Left
        Top = $rect.Top
        Width = $rect.Right - $rect.Left
        Height = $rect.Bottom - $rect.Top
    }
}

function Send-TargetToBack {
    param([IntPtr]$Hwnd)

    if (-not [NativeWin]::IsWindow($Hwnd)) {
        return $false
    }

    if ([NativeWin]::IsIconic($Hwnd)) {
        [void][NativeWin]::ShowWindow($Hwnd, $SW_RESTORE)
    }

    [void][NativeWin]::ShowWindow($Hwnd, $SW_SHOWNOACTIVATE)
    return [NativeWin]::SetWindowPos($Hwnd, $HWND_BOTTOM, 0, 0, 0, 0, ($SWP_NOMOVE -bor $SWP_NOSIZE -bor $SWP_NOACTIVATE -bor $SWP_SHOWWINDOW))
}

function Bring-TargetForward {
    param([IntPtr]$Hwnd)

    if (-not [NativeWin]::IsWindow($Hwnd)) {
        return $false
    }

    if ([NativeWin]::IsIconic($Hwnd)) {
        [void][NativeWin]::ShowWindow($Hwnd, $SW_RESTORE)
    }

    [void][NativeWin]::SetWindowPos($Hwnd, $HWND_TOP, 0, 0, 0, 0, ($SWP_NOMOVE -bor $SWP_NOSIZE -bor $SWP_SHOWWINDOW))
    return [NativeWin]::SetForegroundWindow($Hwnd)
}

function Compact-TargetToCorner {
    param([IntPtr]$Hwnd)

    if (-not [NativeWin]::IsWindow($Hwnd)) {
        return $false
    }

    Save-WindowBounds -Hwnd $Hwnd
    if ([NativeWin]::IsIconic($Hwnd)) {
        [void][NativeWin]::ShowWindow($Hwnd, $SW_RESTORE)
    }

    $area = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $width = [Math]::Min(520, [Math]::Max(320, [int]($area.Width * 0.28)))
    $height = [Math]::Min(320, [Math]::Max(200, [int]($area.Height * 0.28)))
    $x = $area.Right - $width - 18
    $y = $area.Bottom - $height - 18

    [void][NativeWin]::ShowWindow($Hwnd, $SW_SHOWNOACTIVATE)
    return [NativeWin]::SetWindowPos($Hwnd, $HWND_BOTTOM, $x, $y, $width, $height, ($SWP_NOACTIVATE -bor $SWP_SHOWWINDOW))
}

function Restore-TargetBounds {
    param([IntPtr]$Hwnd)

    if (-not [NativeWin]::IsWindow($Hwnd)) {
        return $false
    }

    $key = $Hwnd.ToInt64()
    if (-not $script:savedBounds.ContainsKey($key)) {
        return $false
    }

    $bounds = $script:savedBounds[$key]
    [void][NativeWin]::ShowWindow($Hwnd, $SW_SHOWNOACTIVATE)
    return [NativeWin]::SetWindowPos($Hwnd, $HWND_BOTTOM, $bounds.Left, $bounds.Top, $bounds.Width, $bounds.Height, ($SWP_NOACTIVATE -bor $SWP_SHOWWINDOW))
}

function Stop-Recording {
    param([System.Windows.Forms.TextBox]$LogBox)

    if ($null -eq $script:ffmpegProcess -and $null -eq $script:audioProcess) {
        return
    }

    $recording = $script:activeRecording

    Stop-ProcessNicely -Process $script:ffmpegProcess -TimeoutMs 6000
    Stop-ProcessNicely -Process $script:audioProcess -TimeoutMs 6000

    $script:ffmpegProcess = $null
    $script:audioProcess = $null

    if ($LogBox) {
        $LogBox.AppendText("Capture stopped." + [Environment]::NewLine)
    }

    if ($recording -and $recording.Mode -eq "loopback") {
        Invoke-MuxLoopbackRecording -Recording $recording -LogBox $LogBox
    }

    $script:activeRecording = $null
}

function Start-Recording {
    param(
        [pscustomobject]$Window,
        [string]$AudioDevice,
        [int]$Fps,
        [bool]$DrawMouse,
        [bool]$MonitorOn,
        [System.Windows.Forms.TextBox]$LogBox
    )

    if (($null -ne $script:ffmpegProcess -and -not $script:ffmpegProcess.HasExited) -or ($null -ne $script:audioProcess -and -not $script:audioProcess.HasExited)) {
        [System.Windows.Forms.MessageBox]::Show("A recording is already running.", "Window Back Recorder") | Out-Null
        return $false
    }

    $ffmpeg = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if (-not $ffmpeg) {
        [System.Windows.Forms.MessageBox]::Show("ffmpeg was not found in PATH.", "Window Back Recorder") | Out-Null
        return $false
    }

    if (-not [NativeWin]::IsWindow($Window.Hwnd)) {
        [System.Windows.Forms.MessageBox]::Show("The selected window no longer exists. Refresh and select it again.", "Window Back Recorder") | Out-Null
        return $false
    }

    $title = Get-WindowTitle -Hwnd $Window.Hwnd
    if ([string]::IsNullOrWhiteSpace($title)) {
        [System.Windows.Forms.MessageBox]::Show("The selected window has no title now. Try refreshing the list.", "Window Back Recorder") | Out-Null
        return $false
    }

    if (-not (Test-Path $script:recordingsDir)) {
        [void](New-Item -ItemType Directory -Path $script:recordingsDir -Force)
    }

    $output = Get-UniqueOutputPath -Directory $script:recordingsDir -Extension ".mp4"
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($output)
    $isLoopback = -not [string]::IsNullOrWhiteSpace($AudioDevice) -and $script:loopbackAudioSources.ContainsKey($AudioDevice)
    $hasDShowAudio = -not [string]::IsNullOrWhiteSpace($AudioDevice) -and $AudioDevice -ne $script:noAudioLabel -and -not $isLoopback
    $videoOutput = $output
    $audioTemp = $null
    $loopbackSourceName = $null

    if ($isLoopback) {
        $videoOutput = Join-Path $script:recordingsDir ($baseName + ".video.mkv")
        $audioTemp = Join-Path $script:recordingsDir ($baseName + ".audio.wav")
        $loopbackSourceName = [string]$script:loopbackAudioSources[$AudioDevice]
    }

    $args = New-Object System.Collections.ArrayList
    $useGfxCapture = Test-FfmpegFilterAvailable -FilterName "gfxcapture"
    [void]$args.Add("-hide_banner")
    [void]$args.Add("-y")

    if ($useGfxCapture) {
        $source = "gfxcapture=hwnd={0}:capture_cursor={1}:display_border=false:max_framerate={2}:width=-2:height=-2" -f $Window.Hwnd.ToInt64(), $(if ($DrawMouse) { "true" } else { "false" }), $Fps
        [void]$args.Add("-f")
        [void]$args.Add("lavfi")
        [void]$args.Add("-i")
        [void]$args.Add($source)
    }
    else {
        [void]$args.Add("-thread_queue_size")
        [void]$args.Add("512")
        [void]$args.Add("-f")
        [void]$args.Add("gdigrab")
        [void]$args.Add("-draw_mouse")
        [void]$args.Add($(if ($DrawMouse) { "1" } else { "0" }))
        [void]$args.Add("-framerate")
        [void]$args.Add([string]$Fps)
        [void]$args.Add("-i")
        [void]$args.Add("title=$title")
    }

    if ($hasDShowAudio) {
        [void]$args.Add("-thread_queue_size")
        [void]$args.Add("512")
        [void]$args.Add("-f")
        [void]$args.Add("dshow")
        [void]$args.Add("-i")
        [void]$args.Add("audio=$AudioDevice")
        [void]$args.Add("-map")
        [void]$args.Add("0:v:0")
        [void]$args.Add("-map")
        [void]$args.Add("1:a:0")
    }
    else {
        [void]$args.Add("-map")
        [void]$args.Add("0:v:0")
        [void]$args.Add("-an")
    }

    [void]$args.Add("-vf")
    if ($useGfxCapture) {
        [void]$args.Add("hwdownload,format=bgra,pad=ceil(iw/2)*2:ceil(ih/2)*2,format=yuv420p")
    }
    else {
        [void]$args.Add("pad=ceil(iw/2)*2:ceil(ih/2)*2,format=yuv420p")
    }
    [void]$args.Add("-c:v")
    [void]$args.Add("libx264")
    [void]$args.Add("-preset")
    [void]$args.Add("veryfast")
    [void]$args.Add("-crf")
    [void]$args.Add("23")
    [void]$args.Add("-pix_fmt")
    [void]$args.Add("yuv420p")

    if ($hasDShowAudio) {
        [void]$args.Add("-c:a")
        [void]$args.Add("aac")
        [void]$args.Add("-b:a")
        [void]$args.Add("160k")
    }

    [void]$args.Add($videoOutput)

    $argumentString = New-ArgumentString $args

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ffmpeg.Source
    $psi.Arguments = $argumentString
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.EnableRaisingEvents = $true

    $proc.add_ErrorDataReceived({
        param($sender, $eventArgs)
        if ($eventArgs.Data -and $script:logBox) {
            try {
                [void]$script:logBox.BeginInvoke([Action]{
                    $script:logBox.AppendText($eventArgs.Data + [Environment]::NewLine)
                })
            }
            catch {
            }
        }
    })

    try {
        if (-not $proc.Start()) {
            return $false
        }
        $script:ffmpegProcess = $proc
        $proc.BeginErrorReadLine()

        if ($isLoopback) {
            if (-not (Start-LoopbackAudioRecorder -OutputPath $audioTemp -SourceName $loopbackSourceName -MonitorOn $MonitorOn -LogBox $LogBox)) {
                Stop-ProcessNicely -Process $proc -TimeoutMs 3000
                $script:ffmpegProcess = $null
                return $false
            }

            $script:activeRecording = [pscustomobject]@{
                Mode = "loopback"
                FinalOutput = $output
                VideoTemp = $videoOutput
                AudioTemp = $audioTemp
                AudioSource = $loopbackSourceName
            }
        }
        elseif ($hasDShowAudio) {
            $script:activeRecording = [pscustomobject]@{
                Mode = "dshow"
                FinalOutput = $output
            }
        }
        else {
            $script:activeRecording = [pscustomobject]@{
                Mode = "video-only"
                FinalOutput = $output
            }
        }

        $LogBox.AppendText("Recording started: $output" + [Environment]::NewLine)
        $LogBox.AppendText("Target title: $title" + [Environment]::NewLine)
        $LogBox.AppendText(("Video engine: {0}" -f $(if ($useGfxCapture) { "gfxcapture" } else { "gdigrab fallback" })) + [Environment]::NewLine)
        if ($isLoopback) {
            $LogBox.AppendText("Audio: " + $AudioDevice + ", muxed after stop" + [Environment]::NewLine)
            $LogBox.AppendText(("Listening monitor starts: {0}" -f $(if ($MonitorOn) { "on" } else { "off" })) + [Environment]::NewLine)
        }
        elseif ($hasDShowAudio) {
            $LogBox.AppendText("Audio device: $AudioDevice" + [Environment]::NewLine)
        }
        else {
            $LogBox.AppendText("Audio: none" + [Environment]::NewLine)
        }
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Could not start ffmpeg") | Out-Null
        return $false
    }
}

if ($SelfTest) {
    Write-Output "Native window API loaded."
    $windowCount = @(Get-OpenWindows).Count
    Write-Output ("Visible windows found: {0}" -f $windowCount)
    $audioDevices = @(Get-DShowAudioDevices)
    Write-Output ("DirectShow audio devices found: {0}" -f $audioDevices.Count)
    Write-Output ("System loopback audio available: {0}" -f (Test-LoopbackAudioAvailable))
    $loopbacks = @(Get-LoopbackAudioSources)
    Write-Output ("Loopback sources found: {0}" -f $loopbacks.Count)
    Write-Output ("FFmpeg gfxcapture available: {0}" -f (Test-FfmpegFilterAvailable -FilterName "gfxcapture"))
    $ffmpeg = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if ($ffmpeg) {
        Write-Output ("ffmpeg: {0}" -f $ffmpeg.Source)
    }
    else {
        Write-Output "ffmpeg: not found"
        exit 1
    }
    exit 0
}

Load-Settings
[void][System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "Window Back Recorder"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1060, 740)
$form.MinimumSize = New-Object System.Drawing.Size(820, 560)

$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = "Top"
$topPanel.Height = 112
$form.Controls.Add($topPanel)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh windows"
$refreshButton.Location = New-Object System.Drawing.Point(12, 12)
$refreshButton.Size = New-Object System.Drawing.Size(120, 28)
$topPanel.Controls.Add($refreshButton)

$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "Start recording"
$startButton.Location = New-Object System.Drawing.Point(142, 12)
$startButton.Size = New-Object System.Drawing.Size(120, 28)
$topPanel.Controls.Add($startButton)

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = "Stop"
$stopButton.Enabled = $false
$stopButton.Location = New-Object System.Drawing.Point(272, 12)
$stopButton.Size = New-Object System.Drawing.Size(80, 28)
$topPanel.Controls.Add($stopButton)

$sendBackButton = New-Object System.Windows.Forms.Button
$sendBackButton.Text = "View off"
$sendBackButton.Location = New-Object System.Drawing.Point(362, 12)
$sendBackButton.Size = New-Object System.Drawing.Size(92, 28)
$topPanel.Controls.Add($sendBackButton)

$bringFrontButton = New-Object System.Windows.Forms.Button
$bringFrontButton.Text = "View on"
$bringFrontButton.Location = New-Object System.Drawing.Point(464, 12)
$bringFrontButton.Size = New-Object System.Drawing.Size(92, 28)
$topPanel.Controls.Add($bringFrontButton)

$compactButton = New-Object System.Windows.Forms.Button
$compactButton.Text = "Compact corner"
$compactButton.Location = New-Object System.Drawing.Point(566, 12)
$compactButton.Size = New-Object System.Drawing.Size(118, 28)
$topPanel.Controls.Add($compactButton)

$restoreButton = New-Object System.Windows.Forms.Button
$restoreButton.Text = "Restore size"
$restoreButton.Location = New-Object System.Drawing.Point(694, 12)
$restoreButton.Size = New-Object System.Drawing.Size(100, 28)
$topPanel.Controls.Add($restoreButton)

$soundMixerButton = New-Object System.Windows.Forms.Button
$soundMixerButton.Text = "Sound mixer"
$soundMixerButton.Location = New-Object System.Drawing.Point(804, 12)
$soundMixerButton.Size = New-Object System.Drawing.Size(104, 28)
$topPanel.Controls.Add($soundMixerButton)

$autoBackCheck = New-Object System.Windows.Forms.CheckBox
$autoBackCheck.Text = "Send target behind other windows after start"
$autoBackCheck.Checked = $true
$autoBackCheck.Location = New-Object System.Drawing.Point(12, 46)
$autoBackCheck.Size = New-Object System.Drawing.Size(280, 20)
$topPanel.Controls.Add($autoBackCheck)

$drawMouseCheck = New-Object System.Windows.Forms.CheckBox
$drawMouseCheck.Text = "Record mouse cursor"
$drawMouseCheck.Checked = $true
$drawMouseCheck.Location = New-Object System.Drawing.Point(306, 46)
$drawMouseCheck.Size = New-Object System.Drawing.Size(150, 20)
$topPanel.Controls.Add($drawMouseCheck)

$listenCheck = New-Object System.Windows.Forms.CheckBox
$listenCheck.Text = "Listen to captured audio"
$listenCheck.Checked = $false
$listenCheck.Location = New-Object System.Drawing.Point(470, 46)
$listenCheck.Size = New-Object System.Drawing.Size(170, 20)
$topPanel.Controls.Add($listenCheck)

$fpsLabel = New-Object System.Windows.Forms.Label
$fpsLabel.Text = "FPS"
$fpsLabel.Location = New-Object System.Drawing.Point(652, 48)
$fpsLabel.Size = New-Object System.Drawing.Size(32, 18)
$topPanel.Controls.Add($fpsLabel)

$fpsBox = New-Object System.Windows.Forms.NumericUpDown
$fpsBox.Minimum = 5
$fpsBox.Maximum = 60
$fpsBox.Value = 30
$fpsBox.Location = New-Object System.Drawing.Point(686, 45)
$fpsBox.Size = New-Object System.Drawing.Size(60, 22)
$topPanel.Controls.Add($fpsBox)

$audioLabel = New-Object System.Windows.Forms.Label
$audioLabel.Text = "Audio"
$audioLabel.Location = New-Object System.Drawing.Point(762, 48)
$audioLabel.Size = New-Object System.Drawing.Size(42, 18)
$topPanel.Controls.Add($audioLabel)

$audioCombo = New-Object System.Windows.Forms.ComboBox
$audioCombo.DropDownStyle = "DropDownList"
$audioCombo.Location = New-Object System.Drawing.Point(808, 45)
$audioCombo.Size = New-Object System.Drawing.Size(236, 24)
$topPanel.Controls.Add($audioCombo)

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = "Save folder"
$outputLabel.Location = New-Object System.Drawing.Point(12, 82)
$outputLabel.Size = New-Object System.Drawing.Size(80, 18)
$topPanel.Controls.Add($outputLabel)

$outputDirBox = New-Object System.Windows.Forms.TextBox
$outputDirBox.Location = New-Object System.Drawing.Point(96, 78)
$outputDirBox.Size = New-Object System.Drawing.Size(724, 22)
$outputDirBox.Text = $script:recordingsDir
$topPanel.Controls.Add($outputDirBox)

$browseFolderButton = New-Object System.Windows.Forms.Button
$browseFolderButton.Text = "Browse"
$browseFolderButton.Location = New-Object System.Drawing.Point(828, 76)
$browseFolderButton.Size = New-Object System.Drawing.Size(76, 26)
$topPanel.Controls.Add($browseFolderButton)

$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Text = "Open"
$openFolderButton.Location = New-Object System.Drawing.Point(912, 76)
$openFolderButton.Size = New-Object System.Drawing.Size(72, 26)
$topPanel.Controls.Add($openFolderButton)

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = "Fill"
$split.Orientation = "Horizontal"
$split.SplitterDistance = 365
$form.Controls.Add($split)

$windowList = New-Object System.Windows.Forms.ListView
$windowList.Dock = "Fill"
$windowList.View = "Details"
$windowList.FullRowSelect = $true
$windowList.GridLines = $true
$windowList.HideSelection = $false
[void]$windowList.Columns.Add("Title", 560)
[void]$windowList.Columns.Add("Process", 140)
[void]$windowList.Columns.Add("Size", 110)
[void]$windowList.Columns.Add("HWND", 120)
$split.Panel1.Controls.Add($windowList)

$bottomPanel = New-Object System.Windows.Forms.Panel
$bottomPanel.Dock = "Fill"
$split.Panel2.Controls.Add($bottomPanel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Pick a window, then start recording. Do not minimize or hide the target window."
$statusLabel.Dock = "Top"
$statusLabel.Height = 28
$bottomPanel.Controls.Add($statusLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$script:logBox = $logBox
$logBox.Multiline = $true
$logBox.ScrollBars = "Both"
$logBox.ReadOnly = $true
$logBox.WordWrap = $false
$logBox.Dock = "Fill"
$bottomPanel.Controls.Add($logBox)

function Refresh-WindowList {
    $currentHwnd = $null
    if ($windowList.SelectedItems.Count -gt 0) {
        $currentHwnd = $windowList.SelectedItems[0].Tag.Hwnd.ToInt64()
    }

    $windowList.BeginUpdate()
    $windowList.Items.Clear()
    foreach ($win in Get-OpenWindows) {
        if ($win.Title -eq $form.Text) {
            continue
        }

        $item = New-Object System.Windows.Forms.ListViewItem($win.Title)
        [void]$item.SubItems.Add($win.ProcessName)
        [void]$item.SubItems.Add(("{0}x{1}" -f $win.Width, $win.Height))
        [void]$item.SubItems.Add(("0x{0:X}" -f $win.Hwnd.ToInt64()))
        $item.Tag = $win
        [void]$windowList.Items.Add($item)

        if ($currentHwnd -and $win.Hwnd.ToInt64() -eq $currentHwnd) {
            $item.Selected = $true
        }
    }
    $windowList.EndUpdate()
}

function Refresh-AudioDevices {
    $audioCombo.Items.Clear()
    $script:loopbackAudioSources = @{}
    [void]$audioCombo.Items.Add($script:noAudioLabel)

    foreach ($source in Get-LoopbackAudioSources) {
        $label = $script:loopbackAudioPrefix + [string]$source.name
        $suffix = 2
        while ($script:loopbackAudioSources.ContainsKey($label)) {
            $label = $script:loopbackAudioPrefix + [string]$source.name + " ($suffix)"
            $suffix++
        }
        $script:loopbackAudioSources[$label] = [string]$source.name
        [void]$audioCombo.Items.Add($label)
    }

    foreach ($device in Get-DShowAudioDevices) {
        [void]$audioCombo.Items.Add($device)
    }

    if ($script:loopbackAudioSources.Count -gt 0) {
        $audioCombo.SelectedItem = @($script:loopbackAudioSources.Keys)[0]
    }
    else {
        $audioCombo.SelectedIndex = 0
    }
}

function Apply-OutputFolderFromUi {
    $candidate = $outputDirBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        [System.Windows.Forms.MessageBox]::Show("Choose a save folder first.", "Window Back Recorder") | Out-Null
        return $false
    }

    try {
        $fullPath = [System.IO.Path]::GetFullPath($candidate)
        if (-not (Test-Path $fullPath)) {
            [void](New-Item -ItemType Directory -Path $fullPath -Force)
        }
        $script:recordingsDir = $fullPath
        $outputDirBox.Text = $fullPath
        Save-Settings
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Could not use that save folder: " + $_.Exception.Message, "Window Back Recorder") | Out-Null
        return $false
    }
}

function Get-SelectedWindow {
    if ($windowList.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select a target window first.", "Window Back Recorder") | Out-Null
        return $null
    }

    return $windowList.SelectedItems[0].Tag
}

$refreshButton.Add_Click({
    Refresh-WindowList
})

$browseFolderButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Choose where recordings are saved"
    $dialog.ShowNewFolderButton = $true
    if (Test-Path $outputDirBox.Text) {
        $dialog.SelectedPath = $outputDirBox.Text
    }
    elseif (Test-Path $script:recordingsDir) {
        $dialog.SelectedPath = $script:recordingsDir
    }

    if ($dialog.ShowDialog($form) -eq "OK") {
        $outputDirBox.Text = $dialog.SelectedPath
        [void](Apply-OutputFolderFromUi)
        $statusLabel.Text = "Save folder set: $script:recordingsDir"
    }
})

$openFolderButton.Add_Click({
    if (Apply-OutputFolderFromUi) {
        Start-Process explorer.exe -ArgumentList ('"{0}"' -f $script:recordingsDir)
    }
})

$soundMixerButton.Add_Click({
    Start-Process "ms-settings:apps-volume"
})

$listenCheck.Add_CheckedChanged({
    Set-MonitoringState -Enabled $listenCheck.Checked -LogBox $logBox
})

$windowList.Add_SelectedIndexChanged({
    if ($windowList.SelectedItems.Count -gt 0) {
        $script:selectedWindow = $windowList.SelectedItems[0].Tag
        $statusLabel.Text = "Selected: " + $script:selectedWindow.Display
    }
})

$sendBackButton.Add_Click({
    $win = Get-SelectedWindow
    if ($win) {
        if (Send-TargetToBack -Hwnd $win.Hwnd) {
            $statusLabel.Text = "Target sent behind other windows. It is still alive; avoid minimizing it."
        }
    }
})

$bringFrontButton.Add_Click({
    $win = Get-SelectedWindow
    if ($win) {
        [void](Bring-TargetForward -Hwnd $win.Hwnd)
        $statusLabel.Text = "Target brought to the front."
    }
})

$compactButton.Add_Click({
    $win = Get-SelectedWindow
    if ($win) {
        if (Compact-TargetToCorner -Hwnd $win.Hwnd) {
            $statusLabel.Text = "Target compacted to the lower-right corner and sent back."
        }
    }
})

$restoreButton.Add_Click({
    $win = Get-SelectedWindow
    if ($win) {
        if (Restore-TargetBounds -Hwnd $win.Hwnd) {
            $statusLabel.Text = "Target size restored and sent back."
        }
        else {
            $statusLabel.Text = "No saved size for this target yet."
        }
    }
})

$startButton.Add_Click({
    $win = Get-SelectedWindow
    if (-not $win) {
        return
    }

    if (-not (Apply-OutputFolderFromUi)) {
        return
    }

    Save-WindowBounds -Hwnd $win.Hwnd
    $audio = [string]$audioCombo.SelectedItem
    $started = Start-Recording -Window $win -AudioDevice $audio -Fps ([int]$fpsBox.Value) -DrawMouse $drawMouseCheck.Checked -MonitorOn $listenCheck.Checked -LogBox $logBox
    if ($started) {
        if ($autoBackCheck.Checked) {
            [void](Send-TargetToBack -Hwnd $win.Hwnd)
        }
        $startButton.Enabled = $false
        $stopButton.Enabled = $true
        $browseFolderButton.Enabled = $false
        $outputDirBox.Enabled = $false
        $statusLabel.Text = "Recording. The target can be covered by other windows, but do not minimize or close it."
    }
})

$stopButton.Add_Click({
    Stop-Recording -LogBox $logBox
    $startButton.Enabled = $true
    $stopButton.Enabled = $false
    $browseFolderButton.Enabled = $true
    $outputDirBox.Enabled = $true
    $statusLabel.Text = "Stopped. Recordings are saved under: $script:recordingsDir"
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({
    if ($script:ffmpegProcess -and $script:ffmpegProcess.HasExited) {
        $exitCode = $script:ffmpegProcess.ExitCode
        Stop-Recording -LogBox $logBox
        $startButton.Enabled = $true
        $stopButton.Enabled = $false
        $browseFolderButton.Enabled = $true
        $outputDirBox.Enabled = $true
        $statusLabel.Text = "ffmpeg exited with code $exitCode. Check the log below."
        $logBox.AppendText("ffmpeg exited with code $exitCode" + [Environment]::NewLine)
    }
})
$timer.Start()

$form.Add_FormClosing({
    if (($script:ffmpegProcess -and -not $script:ffmpegProcess.HasExited) -or ($script:audioProcess -and -not $script:audioProcess.HasExited)) {
        $answer = [System.Windows.Forms.MessageBox]::Show("A recording is running. Stop it and close?", "Window Back Recorder", "YesNo", "Warning")
        if ($answer -ne "Yes") {
            $_.Cancel = $true
            return
        }
        Stop-Recording -LogBox $logBox
    }
})

Refresh-AudioDevices
Refresh-WindowList

[void][System.Windows.Forms.Application]::Run($form)

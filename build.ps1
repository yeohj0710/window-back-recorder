$ErrorActionPreference = "Stop"

$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) {
    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path $csc)) {
    throw "Could not find csc.exe from .NET Framework."
}

$exeName = (-join @(
    [char]0xBC31, [char]0xADF8, [char]0xB77C, [char]0xC6B4, [char]0xB4DC,
    [char]0x0020,
    [char]0xC601, [char]0xC0C1,
    [char]0x0020,
    [char]0xB179, [char]0xD654,
    [char]0x0020,
    [char]0xD504, [char]0xB85C, [char]0xADF8, [char]0xB7A8
)) + ".exe"
$output = Join-Path $repo $exeName
$source = Join-Path $repo "WindowBackRecorderApp.cs"
$frameworkDir = Split-Path -Parent $csc
$wpfDir = Join-Path $frameworkDir "WPF"
$icon = Join-Path $repo "app.ico"

& $csc `
    "/nologo" `
    "/target:winexe" `
    "/out:$output" `
    "/win32icon:$icon" `
    "/reference:$(Join-Path $wpfDir 'PresentationCore.dll')" `
    "/reference:$(Join-Path $wpfDir 'PresentationFramework.dll')" `
    "/reference:$(Join-Path $wpfDir 'WindowsBase.dll')" `
    "/reference:$(Join-Path $frameworkDir 'System.Xaml.dll')" `
    "/reference:System.Windows.Forms.dll" `
    "/reference:System.Drawing.dll" `
    "/reference:System.Web.Extensions.dll" `
    "$source"

if ($LASTEXITCODE -ne 0) {
    throw "csc.exe failed with exit code $LASTEXITCODE"
}

Write-Output "Built $output"

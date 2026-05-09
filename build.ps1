$ErrorActionPreference = "Stop"

$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) {
    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path $csc)) {
    throw "Could not find csc.exe from .NET Framework."
}

$output = Join-Path $repo "WindowBackRecorder.exe"
$source = Join-Path $repo "WindowBackRecorderApp.cs"
$frameworkDir = Split-Path -Parent $csc
$wpfDir = Join-Path $frameworkDir "WPF"

& $csc `
    "/nologo" `
    "/target:winexe" `
    "/out:$output" `
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

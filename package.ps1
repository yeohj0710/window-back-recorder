$ErrorActionPreference = "Stop"

$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$dist = Join-Path $repo "dist"
$appName = -join @(
    [char]0xBC31, [char]0xADF8, [char]0xB77C, [char]0xC6B4, [char]0xB4DC,
    [char]0x0020,
    [char]0xC601, [char]0xC0C1,
    [char]0x0020,
    [char]0xB179, [char]0xD654,
    [char]0x0020,
    [char]0xD504, [char]0xB85C, [char]0xADF8, [char]0xB7A8
)
$supportFolderName = "$appName`_$([char]0xC790)$([char]0xB8CC)"
$exeName = "$appName.exe"
$zipName = "$appName.zip"
$stageRoot = Join-Path $dist $appName
$support = Join-Path $stageRoot $supportFolderName
$zipPath = Join-Path $dist $zipName

& (Join-Path $repo "build.ps1")
if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE"
}

if (Test-Path $stageRoot) {
    Remove-Item -LiteralPath $stageRoot -Recurse -Force
}
if (-not (Test-Path $dist)) {
    New-Item -ItemType Directory -Path $dist | Out-Null
}

New-Item -ItemType Directory -Path $support | Out-Null

Copy-Item -LiteralPath (Join-Path $repo $exeName) -Destination (Join-Path $stageRoot $exeName) -Force

$supportFiles = @(
    "README.md",
    "LICENSE",
    "requirements.txt",
    "loopback_audio_recorder.py",
    "app.ico",
    "app_icon.png"
)

foreach ($file in $supportFiles) {
    Copy-Item -LiteralPath (Join-Path $repo $file) -Destination (Join-Path $support $file) -Force
}

if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path (Join-Path $stageRoot "*") -DestinationPath $zipPath -Force
Write-Output "Created $zipPath"

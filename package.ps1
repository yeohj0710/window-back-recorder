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
$bin = Join-Path $support "bin"
$zipPath = Join-Path $dist $zipName
$helperBuild = Join-Path $repo "build\loopback_audio_recorder"
$helperDist = Join-Path $helperBuild "dist"
$helperWork = Join-Path $helperBuild "work"
$helperExe = Join-Path $helperDist "loopback_audio_recorder.exe"

& (Join-Path $repo "build.ps1")
if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE"
}

$ffmpeg = Get-Command ffmpeg -ErrorAction SilentlyContinue
if (-not $ffmpeg) {
    throw "ffmpeg.exe was not found. Install FFmpeg on the build machine before packaging."
}

$previousErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "Continue"
python -m PyInstaller --version > $null 2> $null
$hasPyInstaller = $LASTEXITCODE -eq 0
$ErrorActionPreference = $previousErrorActionPreference

if (-not $hasPyInstaller) {
    Write-Output "Installing PyInstaller for the build machine..."
    python -m pip install --user pyinstaller
    if ($LASTEXITCODE -ne 0) {
        throw "Could not install PyInstaller."
    }
}

if (Test-Path $helperBuild) {
    Remove-Item -LiteralPath $helperBuild -Recurse -Force
}

python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --name loopback_audio_recorder `
    --distpath "$helperDist" `
    --workpath "$helperWork" `
    --specpath "$helperBuild" `
    (Join-Path $repo "loopback_audio_recorder.py")

if ($LASTEXITCODE -ne 0 -or -not (Test-Path $helperExe)) {
    throw "Could not build loopback audio helper."
}

if (Test-Path $stageRoot) {
    Remove-Item -LiteralPath $stageRoot -Recurse -Force
}
if (-not (Test-Path $dist)) {
    New-Item -ItemType Directory -Path $dist | Out-Null
}

New-Item -ItemType Directory -Path $support | Out-Null
New-Item -ItemType Directory -Path $bin | Out-Null

Copy-Item -LiteralPath (Join-Path $repo $exeName) -Destination (Join-Path $stageRoot $exeName) -Force
Copy-Item -LiteralPath $ffmpeg.Source -Destination (Join-Path $bin "ffmpeg.exe") -Force
Copy-Item -LiteralPath $helperExe -Destination (Join-Path $bin "loopback_audio_recorder.exe") -Force

$supportFiles = @(
    "README.md",
    "LICENSE",
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

$ErrorActionPreference = "Stop"

$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $repo
$dist = Join-Path $repo "dist"
$releaseDir = Join-Path $repo "$([char]0xBC30)$([char]0xD3EC)$([char]0xC6A9)"
$appName = -join @(
    [char]0xBC31, [char]0xADF8, [char]0xB77C, [char]0xC6B4, [char]0xB4DC,
    [char]0x0020,
    [char]0xC601, [char]0xC0C1,
    [char]0x0020,
    [char]0xB179, [char]0xD654,
    [char]0x0020,
    [char]0xD504, [char]0xB85C, [char]0xADF8, [char]0xB7A8
)
$supportFolderName = "$([char]0xB0B4)$([char]0xBD80)$([char]0xD30C)$([char]0xC77C)"
$guideName = "$([char]0xC0AC)$([char]0xC6A9)$([char]0xC124)$([char]0xBA85)$([char]0xC11C).html"
$recordingsFolderName = "$([char]0xB179)$([char]0xD654) $([char]0xC644)$([char]0xB8CC)$([char]0xB41C) $([char]0xB3D9)$([char]0xC601)$([char]0xC0C1)"
$exeName = "$appName.exe"
$installerName = "$appName $([char]0xC124)$([char]0xCE58).exe"
$zipName = "$appName.zip"
$stageRoot = Join-Path $dist $appName
$support = Join-Path $stageRoot $supportFolderName
$bin = Join-Path $support "bin"
$recordings = Join-Path $stageRoot $recordingsFolderName
$zipPath = Join-Path $dist $zipName
$installerPath = Join-Path $dist $installerName
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
New-Item -ItemType Directory -Path $recordings | Out-Null

Copy-Item -LiteralPath (Join-Path $projectRoot $exeName) -Destination (Join-Path $stageRoot $exeName) -Force
Copy-Item -LiteralPath (Join-Path $projectRoot $guideName) -Destination (Join-Path $stageRoot $guideName) -Force
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

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    $recordingsEntryName = "$recordingsFolderName/"
    $existing = $archive.Entries | Where-Object { $_.FullName -eq $recordingsEntryName } | Select-Object -First 1
    if (-not $existing) {
        $archive.CreateEntry($recordingsEntryName) | Out-Null
    }
}
finally {
    $archive.Dispose()
}

$csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) {
    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path $csc)) {
    throw "Could not find csc.exe from .NET Framework."
}

$frameworkDir = Split-Path -Parent $csc
$installerSource = Join-Path $repo "InstallerApp.cs"
$icon = Join-Path $repo "app.ico"

if (Test-Path $installerPath) {
    Remove-Item -LiteralPath $installerPath -Force
}

& $csc `
    "/nologo" `
    "/target:winexe" `
    "/out:$installerPath" `
    "/win32icon:$icon" `
    "/resource:$zipPath,Payload.Zip" `
    "/reference:System.Windows.Forms.dll" `
    "/reference:System.Drawing.dll" `
    "/reference:$(Join-Path $frameworkDir 'System.IO.Compression.dll')" `
    "$installerSource"

if ($LASTEXITCODE -ne 0) {
    throw "Installer build failed with exit code $LASTEXITCODE"
}

if (Test-Path $releaseDir) {
    Remove-Item -LiteralPath $releaseDir -Recurse -Force
}
New-Item -ItemType Directory -Path $releaseDir | Out-Null
Copy-Item -LiteralPath $installerPath -Destination (Join-Path $releaseDir $installerName) -Force

Write-Output "Created $zipPath"
Write-Output "Created $installerPath"
Write-Output "Copied installer to $releaseDir"

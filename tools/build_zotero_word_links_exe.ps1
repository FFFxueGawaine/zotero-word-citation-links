param(
    [string]$SourceDir = "",
    [string]$OutputExe = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot
if (-not $SourceDir) {
    $SourceDir = Join-Path $RepoRoot "install"
}
if (-not $OutputExe) {
    $OutputExe = Join-Path $RepoRoot "dist\zotero-word-links-installer.exe"
}

if (-not (Get-Command iexpress.exe -ErrorAction SilentlyContinue)) {
    throw "IExpress was not found on this machine."
}

if (-not (Test-Path -LiteralPath $SourceDir)) {
    throw "Source directory not found: $SourceDir"
}

$requiredFiles = @(
    "install_wrapper.cmd",
    "install_zotero_word_links.ps1",
    "ZoteroWordHyperlinks.bas"
)

foreach ($name in $requiredFiles) {
    $path = Join-Path $SourceDir $name
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Required file not found: $path"
    }
}

$outputFolder = Split-Path -Parent $OutputExe
New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("zotero_word_links_iexpress_" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

$sedPath = Join-Path $tempRoot "build.sed"
$sourceDirEscaped = [System.IO.Path]::GetFullPath($SourceDir)
$outputExeEscaped = [System.IO.Path]::GetFullPath($OutputExe)

$sed = @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=0
UseLongFileName=1
InsideCompressed=1
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=Install finished.
TargetName=$outputExeEscaped
FriendlyName=Zotero Word Citation Links Installer
AppLaunched=install_wrapper.cmd
PostInstallCmd=<None>
AdminQuietInstCmd=install_wrapper.cmd
UserQuietInstCmd=install_wrapper.cmd
SourceFiles=SourceFiles
[SourceFiles]
SourceFiles0=$sourceDirEscaped\
[SourceFiles0]
install_wrapper.cmd=
install_zotero_word_links.ps1=
ZoteroWordHyperlinks.bas=
"@

Set-Content -LiteralPath $sedPath -Value $sed -Encoding ASCII

try {
    & iexpress.exe /N /Q $sedPath | Out-Null
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path -LiteralPath $OutputExe)) {
    throw "IExpress did not create the installer: $OutputExe"
}

Get-Item -LiteralPath $OutputExe | Select-Object FullName,Length,LastWriteTime

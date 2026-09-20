param(
    [ValidateSet('Debug','Release')]
    [string]$Configuration = 'Release',
    [string]$Generator = '',
    [switch]$BuildOnly,
    [switch]$Reconfigure
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'common.ps1')

$Root = $PSScriptRoot
$BuildDir = Join-Path $Root 'build'

Stop-JustNotepad
Initialize-CMake

$Generator = Resolve-Generator -Requested $Generator
if ($Generator -notlike 'Visual Studio*') { Initialize-MsvcEnvironment | Out-Null }

Write-Host "Generator: $Generator   Configuration: $Configuration" -ForegroundColor Cyan

if ($Reconfigure -or (Test-NeedsConfigure -Root $Root -BuildDir $BuildDir -Generator $Generator)) {
    Invoke-CMakeConfigure -Root $Root -BuildDir $BuildDir -Generator $Generator -Configuration $Configuration
} else {
    Write-Host "Build system is up to date; building changes only." -ForegroundColor DarkGray
}

Invoke-CMakeBuild -BuildDir $BuildDir -Configuration $Configuration

$ExePath = Get-ExePath -BuildDir $BuildDir -Generator $Generator -Configuration $Configuration
if (-not (Test-Path $ExePath)) { throw "Executable not found: $ExePath" }

Show-PluginSummary -OutputDir (Split-Path $ExePath)
Write-Host "Built: $ExePath" -ForegroundColor Green

if (-not $BuildOnly) {
    Write-Host "Starting app..." -ForegroundColor Cyan
    Start-Process -FilePath $ExePath
}

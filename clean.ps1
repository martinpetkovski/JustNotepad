$ErrorActionPreference = 'Stop'

# Resolve paths
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildDir = Join-Path $Root 'build'

# Ensure no running instance blocks deletion
Get-Process "Just Notepad" -ErrorAction SilentlyContinue | Stop-Process -Force

# Clean build folder
if (Test-Path $BuildDir) {
    Write-Host "Cleaning build directory..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $BuildDir
    Write-Host "Build directory cleaned." -ForegroundColor Green
} else {
    Write-Host "Build directory does not exist." -ForegroundColor Yellow
}

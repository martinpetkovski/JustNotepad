$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'common.ps1')

$BuildDir = Join-Path $PSScriptRoot 'build'

Stop-JustNotepad

if (Test-Path $BuildDir) {
    Write-Host "Cleaning build directory..." -ForegroundColor Yellow
    # Retry: the linker or an antivirus scanner can still hold a handle briefly.
    for ($i = 1; $i -le 5; $i++) {
        try {
            Remove-Item -Recurse -Force $BuildDir -ErrorAction Stop
            break
        } catch {
            if ($i -eq 5) { throw "Could not remove '$BuildDir': $($_.Exception.Message)" }
            Start-Sleep -Milliseconds 500
        }
    }
    Write-Host "Build directory cleaned." -ForegroundColor Green
} else {
    Write-Host "Build directory does not exist." -ForegroundColor Yellow
}

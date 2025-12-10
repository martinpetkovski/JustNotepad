param(
    [ValidateSet('Debug','Release')]
    [string]$Configuration = 'Release',
    [string]$Generator = ''
)

$ErrorActionPreference = 'Stop'

# Resolve paths
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildDir = Join-Path $Root 'build'

# Ensure no running instance blocks linking
Get-Process JustNotepad -ErrorAction SilentlyContinue | Stop-Process -Force

# Create build folder
if (-not (Test-Path $BuildDir)) { New-Item -ItemType Directory -Path $BuildDir | Out-Null }

# Determine generator if not specified
if (-not $Generator) {
    # Prefer Ninja if available, else Visual Studio
    $ninja = Get-Command ninja -ErrorAction SilentlyContinue
    if ($ninja) { $Generator = 'Ninja' }
    else { $Generator = 'Visual Studio 17 2022' }
}

Write-Host "Using generator: $Generator; Configuration: $Configuration"

# Configure
Push-Location $BuildDir
try {
    # Configure
    cmake -G "$Generator" -DCMAKE_BUILD_TYPE=$Configuration "$Root"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "CMake configure failed (exit code $LASTEXITCODE)." -ForegroundColor Red
        exit $LASTEXITCODE
    }

    # Build
    cmake --build . --config $Configuration
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Build failed (exit code $LASTEXITCODE). Not running the app." -ForegroundColor Red
        exit $LASTEXITCODE
    }

    # Resolve executable path based on generator
    if ($Generator -like 'Visual Studio*') {
        $ExePath = Join-Path $BuildDir "$Configuration/JustNotepad.exe"
    } else {
        $ExePath = Join-Path $BuildDir 'JustNotepad.exe'
    }

    if (-not (Test-Path $ExePath)) {
        throw "Executable not found: $ExePath"
    }

    Write-Host "Starting app: $ExePath" -ForegroundColor Cyan
    & $ExePath
}
finally {
    Pop-Location
}

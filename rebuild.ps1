param(
    [ValidateSet('Debug','Release')]
    [string]$Configuration = 'Release',
    [string]$Generator = '',
    [switch]$BuildOnly
)

$ErrorActionPreference = 'Stop'

# Resolve paths
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildDir = Join-Path $Root 'build'

# Ensure no running instance blocks linking
Get-Process "Just Notepad" -ErrorAction SilentlyContinue | Stop-Process -Force

# Clean build folder
if (Test-Path $BuildDir) {
    Write-Host "Cleaning build directory..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $BuildDir
}

# Create build folder
if (-not (Test-Path $BuildDir)) { New-Item -ItemType Directory -Path $BuildDir | Out-Null }

# Check for CMake
if (-not (Get-Command cmake -ErrorAction SilentlyContinue)) {
    $cmakePaths = @(
        "C:\Program Files\CMake\bin\cmake.exe",
        "C:\Program Files (x86)\CMake\bin\cmake.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin\cmake.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin\cmake.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin\cmake.exe"
    )
    
    $cmakePath = $cmakePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if ($cmakePath) {
        Write-Host "Found CMake at $cmakePath"
        $env:PATH = "$([System.IO.Path]::GetDirectoryName($cmakePath));$env:PATH"
    } else {
        Write-Error "CMake not found. Please install CMake or add it to your PATH."
        exit 1
    }
}

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
        $ExePath = Join-Path $BuildDir "$Configuration/Just Notepad.exe"
    } else {
        $ExePath = Join-Path $BuildDir 'Just Notepad.exe'
    }

    if (-not (Test-Path $ExePath)) {
        throw "Executable not found: $ExePath"
    }

    if (-not $BuildOnly) {
        Write-Host "Starting app: $ExePath" -ForegroundColor Cyan
        & $ExePath
    }
}
finally {
    Pop-Location
}

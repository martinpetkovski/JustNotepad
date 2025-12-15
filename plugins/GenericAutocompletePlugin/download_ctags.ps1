$ErrorActionPreference = 'Stop'

$installDir = Join-Path $PSScriptRoot "ctags"
$ctagsExe = Join-Path $installDir "ctags.exe"

if (Test-Path $ctagsExe) {
    Write-Host "Ctags already installed at $ctagsExe"
    exit 0
}

Write-Host "Finding latest Ctags release..."
try {
    $latestRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/universal-ctags/ctags-win32/releases/latest"
    $asset = $latestRelease.assets | Where-Object { $_.name -match "x64.zip$" -and $_.name -notmatch "clang" } | Select-Object -First 1
    
    if (-not $asset) {
        # Fallback to any x64 zip if specific exclusion fails
        $asset = $latestRelease.assets | Where-Object { $_.name -match "x64.zip$" } | Select-Object -First 1
    }

    if (-not $asset) {
        throw "Could not find x64 zip asset in latest release."
    }

    $url = $asset.browser_download_url
    $zipPath = Join-Path $PSScriptRoot "ctags.zip"

    Write-Host "Downloading $url..."
    Invoke-WebRequest -Uri $url -OutFile $zipPath

    Write-Host "Extracting to $installDir..."
    if (Test-Path $installDir) { Remove-Item -Recurse -Force $installDir }
    Expand-Archive -Path $zipPath -DestinationPath $installDir -Force

    Remove-Item $zipPath
    Write-Host "Ctags installed successfully."
}
catch {
    Write-Error "Failed to download/install Ctags: $_"
    exit 1
}

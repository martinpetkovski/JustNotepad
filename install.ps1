# Installs (or updates) JustNotepad for the current user, no admin needed:
#
#   irm https://raw.githubusercontent.com/martinpetkovski/JustNotepad/main/install.ps1 | iex
#
# What it does, and nothing else:
#   1. asks GitHub for the latest release,
#   2. downloads the zip and its published SHA-256, and refuses to continue if they differ,
#   3. puts JustNotepad.exe in %LOCALAPPDATA%\Programs\JustNotepad,
#   4. adds a Start menu shortcut, and starts it.
# Remove it by deleting that folder and the shortcut.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Repo = 'martinpetkovski/JustNotepad'
$Dest = Join-Path $env:LOCALAPPDATA 'Programs\JustNotepad'
$Temp = Join-Path ([IO.Path]::GetTempPath()) ('JustNotepadInstall-' + [Guid]::NewGuid().ToString('N'))

Write-Host 'Finding the latest JustNotepad release...'
$release = Invoke-RestMethod -UseBasicParsing -Headers @{ 'User-Agent' = 'JustNotepad-Installer' } `
    -Uri "https://api.github.com/repos/$Repo/releases/latest"
$zip = $release.assets | Where-Object { $_.name -like '*.zip' } | Select-Object -First 1
$sha = $release.assets | Where-Object { $_.name -like '*.sha256' } | Select-Object -First 1
if (-not $zip) { throw "Release $($release.tag_name) has no zip to install." }
if (-not $sha) { throw "Release $($release.tag_name) has no published checksum; refusing to install it." }

New-Item -ItemType Directory -Force $Temp | Out-Null
try {
    $zipPath = Join-Path $Temp $zip.name
    Write-Host "Downloading $($zip.name)..."
    Invoke-WebRequest -UseBasicParsing -Uri $zip.browser_download_url -OutFile $zipPath
    $expected = ((Invoke-WebRequest -UseBasicParsing -Uri $sha.browser_download_url).Content -split '\s+')[0].Trim().ToLowerInvariant()
    $actual = (Get-FileHash -Algorithm SHA256 -LiteralPath $zipPath).Hash.ToLowerInvariant()
    if ($expected -notmatch '^[0-9a-f]{64}$' -or $actual -ne $expected) {
        throw "The download does not match its published SHA-256 (expected $expected, got $actual). Nothing was installed."
    }
    Write-Host "SHA-256 verified: $actual"

    # Close a running copy so its exe can be replaced.
    Get-Process JustNotepad -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -and $_.Path.StartsWith($Dest, [StringComparison]::OrdinalIgnoreCase) } |
        ForEach-Object { $_.CloseMainWindow() | Out-Null; $_.WaitForExit(5000) | Out-Null }

    $unpacked = Join-Path $Temp 'unpacked'
    Expand-Archive -LiteralPath $zipPath -DestinationPath $unpacked -Force
    New-Item -ItemType Directory -Force $Dest | Out-Null
    Copy-Item -Path (Join-Path $unpacked '*') -Destination $Dest -Recurse -Force
} finally {
    Remove-Item -Recurse -Force $Temp -ErrorAction SilentlyContinue
}

$exe = Join-Path $Dest 'JustNotepad.exe'
$lnk = Join-Path ([Environment]::GetFolderPath('Programs')) 'JustNotepad.lnk'
$shell = New-Object -ComObject WScript.Shell
$s = $shell.CreateShortcut($lnk)
$s.TargetPath = $exe
$s.WorkingDirectory = $Dest
$s.Description = 'JustNotepad'
$s.Save()

Write-Host "Installed JustNotepad $($release.tag_name) to $Dest" -ForegroundColor Green
Start-Process $exe

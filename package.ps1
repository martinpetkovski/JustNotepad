param(
    [int]$MajorVersion = 1,
    [ValidateSet('Debug','Release')]
    [string]$Configuration = 'Release',
    [string]$Generator = '',
    [switch]$SkipBuild,
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'common.ps1')

$Root = $PSScriptRoot
$BuildDir = Join-Path $Root 'build'

Write-Host "Starting package creation..." -ForegroundColor Green

# --------------------------------------------------------------------------
# 1. Work out the release version
# --------------------------------------------------------------------------
Write-Host "Step 1: Calculating release version..." -ForegroundColor Cyan
$Date = Get-Date
$BaseVersion = "{0}.{1:yy}{1:MM}{1:dd}" -f $MajorVersion, $Date

$DailyBuild = 1
$Pattern = "^v" + [Regex]::Escape($BaseVersion) + "\.(\d+)$"
foreach ($tag in @(git tag -l "v$BaseVersion.*")) {
    if ($tag -match $Pattern) {
        $num = [int]$Matches[1]
        if ($num -ge $DailyBuild) { $DailyBuild = $num + 1 }
    }
}

$Version = "$BaseVersion.$DailyBuild"
$TagName = "v$Version"

if (@(git tag -l $TagName).Count -gt 0) {
    throw "Tag $TagName already exists. Aborting before building."
}

Write-Host "Version: $Version   Tag: $TagName"

# --------------------------------------------------------------------------
# 2. Stamp Version.h
# --------------------------------------------------------------------------
$VersionHeaderPath = Join-Path $Root 'src\Version.h'
$VersionParts = $Version.Split('.')
$Major    = $VersionParts[0]
$DatePart = $VersionParts[1]
$Daily    = $VersionParts[2]
$Year     = $DatePart.Substring(0, 2)
$MonthDay = $DatePart.Substring(2)

$VersionContent = @"
#pragma once
#define APP_VERSION_STRING "$Version"
#define APP_VERSION_MAJOR $Major
#define APP_VERSION_MINOR $Year
#define APP_VERSION_PATCH $MonthDay
#define APP_VERSION_BUILD $Daily

#define APP_COMPANY_NAME "Martin Petkovski"
#define APP_FILE_DESCRIPTION "Just a notepad."
#define APP_INTERNAL_NAME "JustNotepad"
#define APP_LEGAL_COPYRIGHT "Copyright (C) 2025-2026 Martin Petkovski"
#define APP_ORIGINAL_FILENAME "JustNotepad.exe"
#define APP_PRODUCT_NAME "Just Notepad"
"@
Set-Content -Path $VersionHeaderPath -Value $VersionContent -Encoding UTF8
Write-Host "Stamped src/Version.h with $Version"

# --------------------------------------------------------------------------
# 3. Build
# --------------------------------------------------------------------------
Initialize-CMake
$Generator = Resolve-Generator -Requested $Generator

if ($SkipBuild) {
    Write-Host "Step 2: Skipping build as requested." -ForegroundColor Yellow
} else {
    Write-Host "Step 2: Building project..." -ForegroundColor Cyan
    & (Join-Path $Root 'build.ps1') -Configuration $Configuration -Generator $Generator -BuildOnly
}

$OutputDir = Get-OutputDir -BuildDir $BuildDir -Generator $Generator -Configuration $Configuration
$ExePath   = Join-Path $OutputDir 'Just Notepad.exe'
if (-not (Test-Path $ExePath)) { throw "Executable not found: $ExePath" }

# The stamped version must actually be in the binary we are about to ship.
$BuiltVersion = (Get-Item $ExePath).VersionInfo.FileVersion
if ($BuiltVersion) {
    Write-Host "Binary reports version: $BuiltVersion" -ForegroundColor DarkGray
}

# --------------------------------------------------------------------------
# 4. Tests
# --------------------------------------------------------------------------
if ($SkipTests) {
    Write-Host "Step 3: Skipping tests as requested." -ForegroundColor Yellow
} else {
    Write-Host "Step 3: Running tests..." -ForegroundColor Cyan
    Push-Location $BuildDir
    try {
        & ctest -C $Configuration --output-on-failure
        if ($LASTEXITCODE -ne 0) { throw "Tests failed with exit code $LASTEXITCODE" }
    } finally {
        Pop-Location
    }
}

# --------------------------------------------------------------------------
# 5. Release artifacts
# --------------------------------------------------------------------------
Write-Host "Step 4: Creating release artifacts..." -ForegroundColor Cyan

$ReleasesRoot = Join-Path $Root 'releases'
if (-not (Test-Path $ReleasesRoot)) { New-Item -ItemType Directory -Path $ReleasesRoot | Out-Null }

$ReleaseDirName = "JustNotepad-$TagName"
$ReleaseDir = Join-Path $ReleasesRoot $ReleaseDirName
if (Test-Path $ReleaseDir) { Remove-Item -Recurse -Force $ReleaseDir }
New-Item -ItemType Directory -Path $ReleaseDir | Out-Null
New-Item -ItemType Directory -Path (Join-Path $ReleaseDir 'plugins') | Out-Null

# NOTE: the file name here is what the in-app updater overwrites on the user's
# disk, so it must stay stable across releases.
Copy-Item -LiteralPath $ExePath -Destination (Join-Path $ReleaseDir 'JustNotepad.exe')

$BuiltPlugins = Join-Path $OutputDir 'plugins'
if (Test-Path $BuiltPlugins) {
    Copy-Item -Path (Join-Path $BuiltPlugins '*') -Destination (Join-Path $ReleaseDir 'plugins') -Recurse -Force
    $count = @(Get-ChildItem -Path (Join-Path $ReleaseDir 'plugins') -Filter '*.dll').Count
    Write-Host "Packaged $count plugin(s)."
} else {
    Write-Host "Warning: no plugins were built at $BuiltPlugins" -ForegroundColor Yellow
}

$BuiltMacros = Join-Path $OutputDir 'macros'
if (Test-Path $BuiltMacros) {
    New-Item -ItemType Directory -Path (Join-Path $ReleaseDir 'macros') | Out-Null
    Copy-Item -Path (Join-Path $BuiltMacros '*') -Destination (Join-Path $ReleaseDir 'macros') -Recurse -Force
}

$ZipFile = Join-Path $ReleasesRoot "$ReleaseDirName.zip"
if (Test-Path $ZipFile) { Remove-Item -Force $ZipFile }
Compress-Archive -Path (Join-Path $ReleaseDir '*') -DestinationPath $ZipFile
Write-Host "Created release archive: $ZipFile" -ForegroundColor Green

# --------------------------------------------------------------------------
# 6. Commit the version bump and tag it
# --------------------------------------------------------------------------
if (@(git status --porcelain -- 'src/Version.h').Count -gt 0) {
    Write-Host "Committing version bump..."
    git add -- 'src/Version.h'
    git commit -m "Bump version to $Version" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Failed to commit the version bump" }
}

Write-Host "Creating git tag $TagName..."
git tag -a $TagName -m "Release $Version"
if ($LASTEXITCODE -ne 0) { throw "Failed to create tag" }

Write-Host "Package created successfully!" -ForegroundColor Green

return [PSCustomObject]@{
    Version = $Version
    TagName = $TagName
    ZipFile = $ZipFile
}

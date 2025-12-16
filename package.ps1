param(
    [int]$MajorVersion = 1
)

$ErrorActionPreference = 'Stop'

Write-Host "Starting Package Creation..." -ForegroundColor Green

# 1. Calculate release version (Moved to start)
Write-Host "Step 1: Calculating release version..." -ForegroundColor Cyan
$Date = Get-Date
$BaseVersion = "{0}.{1:yy}{1:MM}{1:dd}" -f $MajorVersion, $Date

# Calculate daily build number
$DailyBuild = 1
$Pattern = "^v" + [Regex]::Escape($BaseVersion) + "\.(\d+)$"
$Tags = git tag -l "v$BaseVersion.*"

foreach ($tag in $Tags) {
    if ($tag -match $Pattern) {
        $num = [int]$matches[1]
        if ($num -ge $DailyBuild) {
            $DailyBuild = $num + 1
        }
    }
}

$Version = "$BaseVersion.$DailyBuild"
$TagName = "v$Version"

Write-Host "Version: $Version"
Write-Host "Tag: $TagName"

# Update Version.h
$VersionHeaderPath = Join-Path $PSScriptRoot "src\Version.h"
$VersionParts = $Version.Split('.')
$Major = $VersionParts[0]
$DatePart = $VersionParts[1]
$Daily = $VersionParts[2]

# Split date part (yyMMdd) into Year (yy) and MonthDay (MMdd) to fit 16-bit integers
$Year = $DatePart.Substring(0, 2)
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
Set-Content -Path $VersionHeaderPath -Value $VersionContent
Write-Host "Updated Version.h with version $Version and detailed info"

# 2. Build everything (including plugins)
Write-Host "Step 2: Building project..." -ForegroundColor Cyan
.\build.ps1 -Configuration Release -BuildOnly
if ($LASTEXITCODE -ne 0) { 
    Write-Error "Build failed with exit code $LASTEXITCODE"
    exit 1 
}

# 3. Run tests
Write-Host "Step 3: Running tests..." -ForegroundColor Cyan
Push-Location build
try {
    ctest -C Release --output-on-failure
    if ($LASTEXITCODE -ne 0) { 
        Write-Error "Tests failed with exit code $LASTEXITCODE"
        exit 1 
    }
}
finally {
    Pop-Location
}

# 4. Create Release Artifacts
Write-Host "Step 4: Creating release artifacts..." -ForegroundColor Cyan

# Ensure releases folder exists
$ReleasesRoot = "releases"
if (-not (Test-Path $ReleasesRoot)) { New-Item -ItemType Directory -Path $ReleasesRoot | Out-Null }

$ReleaseDirName = "JustNotepad-$TagName"
$ReleaseDir = Join-Path $ReleasesRoot $ReleaseDirName

if (Test-Path $ReleaseDir) { Remove-Item -Recurse -Force $ReleaseDir }
New-Item -ItemType Directory -Path $ReleaseDir | Out-Null
New-Item -ItemType Directory -Path "$ReleaseDir\plugins" | Out-Null

# Copy Executable
Copy-Item "build\Release\Just Notepad.exe" -Destination "$ReleaseDir\JustNotepad.exe"

# Copy Plugins (including subfolders like 'highlight')
Copy-Item "build\Release\plugins\*" -Destination "$ReleaseDir\plugins" -Recurse

# Copy Macros
if (Test-Path "build\Release\macros") {
    New-Item -ItemType Directory -Path "$ReleaseDir\macros" | Out-Null
    Copy-Item "build\Release\macros\*" -Destination "$ReleaseDir\macros" -Recurse
}

# Create Zip
$ZipFile = Join-Path $ReleasesRoot "$ReleaseDirName.zip"
if (Test-Path $ZipFile) { Remove-Item -Force $ZipFile }
Compress-Archive -Path "$ReleaseDir\*" -DestinationPath $ZipFile
Write-Host "Created release archive: $ZipFile" -ForegroundColor Green

# Check if tag exists
$existingTag = git tag -l $TagName
if ($existingTag) {
    Write-Error "Tag $TagName already exists. Aborting."
    exit 1
}

# Create tag
Write-Host "Creating git tag..."
git tag -a $TagName -m "Release $Version"
if ($LASTEXITCODE -ne 0) { 
    Write-Error "Failed to create tag"
    exit 1 
}

Write-Host "Package created successfully!" -ForegroundColor Green

# Return info object
return [PSCustomObject]@{
    Version = $Version
    TagName = $TagName
    ZipFile = $ZipFile
}

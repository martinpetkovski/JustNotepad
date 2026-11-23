# Publishes a release: refreshes the website from README.md, works out the
# version, stamps it, builds, tests, zips, commits the version bump, tags,
# pushes, and creates the GitHub release.
#
#   .\publish.ps1                 full release
#   .\publish.ps1 -DryRun         build, test and zip only; nothing committed, tagged or pushed
#   .\publish.ps1 -SkipTests
#   .\publish.ps1 -MajorVersion 3
#
# Version scheme: <major>.<yyMMdd>.<n>, n = release number for that day.
param(
    [int]$MajorVersion = 2,
    [switch]$SkipTests,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$Root = $PSScriptRoot
$VersionHeader = Join-Path $Root 'src\win\version.h'
$stamped = $false

Push-Location $Root
try {
    # ----------------------------------------------------------------------
    # Pre-flight: refuse to release from a dirty tree or without a remote.
    # ----------------------------------------------------------------------
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw "git was not found on PATH." }
    $dirty = @(git status --porcelain | Where-Object { $_ -notmatch '^\?\?' })
    if ($dirty.Count -gt 0) {
        Write-Host ($dirty -join "`n") -ForegroundColor Yellow
        throw "The working tree has uncommitted changes. Commit or stash them before releasing."
    }
    $Branch = (git rev-parse --abbrev-ref HEAD).Trim()
    $origin = git config --get remote.origin.url
    if (-not $origin) { throw "No 'origin' remote configured." }
    $haveGh = [bool](Get-Command gh -ErrorAction SilentlyContinue)
    if (-not $haveGh -and -not $DryRun) {
        Write-Warning "GitHub CLI ('gh') not found: the tag will be pushed, but the release must be created by hand."
    }

    # ----------------------------------------------------------------------
    # 0. Website: copy README.md into docs/index.html (the page renders it
    #    with marked), so najjak.com/JustNotepad matches the release.
    # ----------------------------------------------------------------------
    Write-Host "Step 0: Updating docs/index.html from README.md..." -ForegroundColor Cyan
    $HtmlPath = Join-Path $Root 'docs\index.html'
    if (Test-Path $HtmlPath) {
        $bt = [string][char]0x60
        $Readme = Get-Content -Path (Join-Path $Root 'README.md') -Raw -Encoding UTF8
        # Inside a JavaScript template string, ` and ${ must be escaped.
        $Escaped = $Readme.Replace($bt, ('\' + $bt)).Replace('${', '\${')
        $Html = Get-Content -Path $HtmlPath -Raw -Encoding UTF8
        $StartMarker = 'const markdown = ' + $bt
        $EndMarker = $bt + ';'
        $Start = $Html.IndexOf($StartMarker)
        $End = if ($Start -ge 0) { $Html.IndexOf($EndMarker, $Start + $StartMarker.Length) } else { -1 }
        if ($Start -lt 0 -or $End -le $Start) {
            Write-Warning "No markdown placeholder found in docs/index.html; leaving it untouched."
        } else {
            $NewHtml = $Html.Substring(0, $Start + $StartMarker.Length) + $Escaped + $Html.Substring($End)
            if ($NewHtml -ne $Html) {
                [IO.File]::WriteAllText($HtmlPath, $NewHtml, (New-Object Text.UTF8Encoding($false)))
                Write-Host "Updated docs/index.html"
                if (-not $DryRun) {
                    git add -- 'docs/index.html'
                    git commit -m "Update docs/index.html from README.md" | Out-Null
                    if ($LASTEXITCODE -ne 0) { throw "Failed to commit docs/index.html." }
                }
            }
        }
    }

    # ----------------------------------------------------------------------
    # 1. Version
    # ----------------------------------------------------------------------
    Write-Host "Step 1: Calculating release version..." -ForegroundColor Cyan
    git fetch --tags --quiet origin 2>$null
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
    if (@(git tag -l $TagName).Count -gt 0) { throw "Tag $TagName already exists." }
    Write-Host "Version: $Version   Tag: $TagName   Branch: $Branch"

    # ----------------------------------------------------------------------
    # 2. Stamp src\win\version.h (numbers without leading zeros: the resource
    #    compiler would read 0920 as octal)
    # ----------------------------------------------------------------------
    $Year = [int]$Date.ToString('yy')
    $MonthDay = [int]$Date.ToString('MMdd')
    $VersionContent = @"
#pragma once
// Rewritten by publish.ps1 for each release.
#define APP_VERSION_STRING "$Version"
#define APP_VERSION_MAJOR $MajorVersion
#define APP_VERSION_MINOR $Year
#define APP_VERSION_PATCH $MonthDay
#define APP_VERSION_BUILD $DailyBuild
"@
    [IO.File]::WriteAllText($VersionHeader, ($VersionContent -replace "`r?`n", "`n") + "`n", (New-Object Text.UTF8Encoding $false))
    $stamped = $true
    Write-Host "Stamped src\win\version.h with $Version"

    # ----------------------------------------------------------------------
    # 3. Build and test
    # ----------------------------------------------------------------------
    Write-Host "Step 2: Building..." -ForegroundColor Cyan
    & (Join-Path $Root 'build.ps1')
    if ($LASTEXITCODE -ne 0) { throw "Build failed." }
    $ExePath = Join-Path $Root 'build\JustNotepad.exe'
    if (-not (Test-Path $ExePath)) { throw "Executable not found: $ExePath" }
    $BuiltVersion = (Get-Item $ExePath).VersionInfo.FileVersion
    Write-Host "Binary reports version: $BuiltVersion" -ForegroundColor DarkGray
    if ($BuiltVersion -ne $Version) { throw "The built exe reports version '$BuiltVersion', expected '$Version'." }

    if ($SkipTests) {
        Write-Host "Step 3: Skipping tests as requested." -ForegroundColor Yellow
    } else {
        Write-Host "Step 3: Running tests..." -ForegroundColor Cyan
        & (Join-Path $Root 'build.ps1') -Test
        if ($LASTEXITCODE -ne 0) { throw "Tests failed." }
    }

    # ----------------------------------------------------------------------
    # 4. Release archive
    # ----------------------------------------------------------------------
    Write-Host "Step 4: Creating the release archive..." -ForegroundColor Cyan
    $ReleasesRoot = Join-Path $Root 'releases'
    New-Item -ItemType Directory -Force $ReleasesRoot | Out-Null
    $ZipFile = Join-Path $ReleasesRoot "JustNotepad-$TagName.zip"
    if (Test-Path $ZipFile) { Remove-Item -Force $ZipFile }
    Compress-Archive -LiteralPath $ExePath -DestinationPath $ZipFile
    Write-Host "Created $ZipFile" -ForegroundColor Green

    if ($DryRun) {
        git checkout --quiet -- 'src/win/version.h'
        $stamped = $false
        Write-Host "Dry run: stopped before committing, tagging and pushing." -ForegroundColor Yellow
        return
    }

    # ----------------------------------------------------------------------
    # 5. Commit the version bump, tag, push, release
    # ----------------------------------------------------------------------
    Write-Host "Step 5: Publishing to GitHub..." -ForegroundColor Cyan
    git add -- 'src/win/version.h'
    git commit -m "Bump version to $Version" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Failed to commit the version bump." }
    $stamped = $false   # committed: nothing to roll back from here on

    git tag -a $TagName -m "Release $Version"
    if ($LASTEXITCODE -ne 0) { throw "Failed to create tag $TagName." }

    # The tag must be reachable from the remote branch, so push the commits first.
    Write-Host "Pushing $Branch..."
    git push origin $Branch
    if ($LASTEXITCODE -ne 0) { throw "Failed to push branch '$Branch'." }
    Write-Host "Pushing tag $TagName..."
    git push origin $TagName
    if ($LASTEXITCODE -ne 0) { throw "Failed to push tag $TagName." }

    if ($haveGh) {
        Write-Host "Creating the GitHub release and uploading the archive..."
        gh release create $TagName $ZipFile --title "Release $Version" --generate-notes
        if ($LASTEXITCODE -ne 0) { throw "gh release create failed. The tag is pushed; create the release by hand and attach $ZipFile." }
        Write-Host "Release $Version published." -ForegroundColor Green
    } else {
        Write-Host "ACTION REQUIRED: create the $TagName release on GitHub and upload '$ZipFile'." -ForegroundColor Yellow
    }
}
finally {
    # If something failed between stamping and committing, put the header back.
    if ($stamped) { git checkout --quiet -- 'src/win/version.h' 2>$null }
    Pop-Location
}

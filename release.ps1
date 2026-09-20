param(
    [int]$MajorVersion = 1,
    [string]$Generator = '',
    [switch]$SkipClean,
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'common.ps1')

$Root = $PSScriptRoot
Push-Location $Root
try {
    # ----------------------------------------------------------------------
    # Pre-flight: refuse to release from a dirty tree or without a remote.
    # ----------------------------------------------------------------------
    $dirty = @(git status --porcelain | Where-Object { $_ -notmatch '^\?\?' })
    if ($dirty.Count -gt 0) {
        Write-Host ($dirty -join "`n") -ForegroundColor Yellow
        throw "The working tree has uncommitted changes. Commit or stash them before releasing."
    }

    $Branch = (git rev-parse --abbrev-ref HEAD).Trim()
    if (-not (git remote get-url origin 2>$null)) { throw "No 'origin' remote configured." }

    # ----------------------------------------------------------------------
    # Keep docs/index.html in sync with README.md
    # ----------------------------------------------------------------------
    Write-Host "Updating docs/index.html from README.md..." -ForegroundColor Cyan
    $HtmlPath = Join-Path $Root 'docs/index.html'
    if (Test-Path $HtmlPath) {
        $bt = [string][char]0x60
        $ReadmeContent = Get-Content -Path (Join-Path $Root 'README.md') -Raw
        $ReadmeContentEscaped = $ReadmeContent.Replace($bt, ('\' + $bt)).Replace('${', '\${')

        $HtmlContent = Get-Content -Path $HtmlPath -Raw
        $StartMarker = 'const markdown = ' + $bt
        $EndMarker = $bt + ';'
        $StartIndex = $HtmlContent.IndexOf($StartMarker)

        if ($StartIndex -lt 0) {
            Write-Warning "Could not find the markdown placeholder in docs/index.html; leaving it untouched."
        } else {
            $EndIndex = $HtmlContent.IndexOf($EndMarker, $StartIndex + $StartMarker.Length)
            if ($EndIndex -le $StartIndex) {
                Write-Warning "Malformed markdown placeholder in docs/index.html; leaving it untouched."
            } else {
                $NewHtmlContent = $HtmlContent.Substring(0, $StartIndex + $StartMarker.Length) +
                                  $ReadmeContentEscaped +
                                  $HtmlContent.Substring($EndIndex)
                Set-Content -Path $HtmlPath -Value $NewHtmlContent -Encoding UTF8 -NoNewline
                Write-Host "Updated docs/index.html"
            }
        }

        if (@(git status --porcelain -- 'docs/index.html').Count -gt 0) {
            git add -- 'docs/index.html'
            git commit -m "Update docs/index.html from README.md" | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "Failed to commit docs/index.html" }
        }
    }

    # ----------------------------------------------------------------------
    # Clean, then build + test + package (package.ps1 does the build)
    # ----------------------------------------------------------------------
    Invoke-DependencyScripts -Root $Root
    if (-not $SkipClean) { & (Join-Path $Root 'clean.ps1') }

    $packageArgs = @{
        MajorVersion  = $MajorVersion
        Configuration = 'Release'
        Generator     = $Generator
    }
    if ($SkipTests) { $packageArgs['SkipTests'] = $true }

    $ReleaseInfo = & (Join-Path $Root 'package.ps1') @packageArgs
    if (-not $ReleaseInfo) { throw "package.ps1 did not return release information." }

    $TagName = $ReleaseInfo.TagName
    $ZipFile = $ReleaseInfo.ZipFile
    $Version = $ReleaseInfo.Version

    if (-not (Test-Path $ZipFile)) { throw "Release archive is missing: $ZipFile" }

    # ----------------------------------------------------------------------
    # Publish
    # ----------------------------------------------------------------------
    Write-Host "Publishing to GitHub..." -ForegroundColor Cyan

    # The tag must be reachable from the remote branch, so push the commits first.
    Write-Host "Pushing $Branch..."
    git push origin $Branch
    if ($LASTEXITCODE -ne 0) { throw "Failed to push branch '$Branch'" }

    Write-Host "Pushing tag $TagName..."
    git push origin $TagName
    if ($LASTEXITCODE -ne 0) { throw "Failed to push tag $TagName" }

    if (Get-Command gh -ErrorAction SilentlyContinue) {
        Write-Host "Creating GitHub release and uploading the archive..."
        gh release create $TagName $ZipFile --title "Release $Version" --generate-notes
        if ($LASTEXITCODE -ne 0) { throw "gh release create failed." }
        Write-Host "Release created and asset uploaded." -ForegroundColor Green
        Write-Host "The in-app updater will now see $TagName." -ForegroundColor Green
    } else {
        Write-Warning "GitHub CLI ('gh') not found, so the release has no attached archive."
        Write-Warning "The in-app updater needs the .zip asset to work."
        Write-Host "ACTION REQUIRED: upload '$ZipFile' to the $TagName release on GitHub." -ForegroundColor Yellow
    }

    Write-Host "Release $Version completed successfully!" -ForegroundColor Green
}
finally {
    Pop-Location
}

param(
    [int]$MajorVersion = 1
)

$ErrorActionPreference = 'Stop'

# Clean and Rebuild
Write-Host "Cleaning..." -ForegroundColor Cyan
.\clean.ps1

Write-Host "Rebuilding..." -ForegroundColor Cyan
.\rebuild.ps1 -Configuration Release

# Call the package script
$ReleaseInfo = .\package.ps1 -MajorVersion $MajorVersion

$TagName = $ReleaseInfo.TagName
$ZipFile = $ReleaseInfo.ZipFile
$Version = $ReleaseInfo.Version

# 4. Push to GitHub / Create Release
Write-Host "Step 4: Publishing to GitHub..." -ForegroundColor Cyan

# Always push the tag first so gh release create finds it
Write-Host "Pushing tag to remote..."
git push origin $TagName
if ($LASTEXITCODE -ne 0) { 
    Write-Error "Failed to push tag to remote"
    exit 1 
}

if (Get-Command gh -ErrorAction SilentlyContinue) {
    Write-Host "GitHub CLI (gh) found. Creating release and uploading asset..."
    
    # Create release and upload asset
    # --generate-notes automatically adds the changelog from git commits
    gh release create $TagName $ZipFile --title "Release $Version" --generate-notes
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Release created and asset uploaded successfully!" -ForegroundColor Green
    } else {
        Write-Error "Failed to create release with GitHub CLI."
        exit 1
    }
} else {
    Write-Warning "GitHub CLI ('gh') not found."
    Write-Warning "The release on GitHub will ONLY contain source code."
    Write-Warning "To automatically upload the zip file, please install GitHub CLI: https://cli.github.com/"
    
    Write-Host "Tag pushed to GitHub."
    Write-Host "ACTION REQUIRED: Please manually upload '$ZipFile' to the release on GitHub." -ForegroundColor Yellow
}

Write-Host "Release $Version completed successfully!" -ForegroundColor Green

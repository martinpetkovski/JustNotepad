param($filePath)
if ($filePath) {
    $dir = Split-Path -Path $filePath
    Invoke-Item $dir
}


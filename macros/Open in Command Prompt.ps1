param($filePath)
if ($filePath) {
    $dir = Split-Path -Path $filePath
    Start-Process cmd.exe -WorkingDirectory $dir
}


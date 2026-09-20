# Shared helpers for build.ps1 / rebuild.ps1 / clean.ps1 / package.ps1 / release.ps1
# Dot-source this file:  . (Join-Path $PSScriptRoot 'common.ps1')

function Get-RepoRoot {
    return $PSScriptRoot
}

# Kills any running instance so the linker can replace the binaries.
function Stop-JustNotepad {
    foreach ($name in @('Just Notepad', 'JustNotepad')) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Host "Stopping running instance: $($_.Name) (PID $($_.Id))" -ForegroundColor DarkGray
            try { $_ | Stop-Process -Force -ErrorAction Stop } catch { }
        }
    }
    Start-Sleep -Milliseconds 200
}

# Makes sure cmake is callable. Throws if it cannot be found.
function Initialize-CMake {
    if (Get-Command cmake -ErrorAction SilentlyContinue) { return }

    $candidates = @()
    foreach ($pf in (Get-ProgramFilesRoots)) {
        $candidates += (Join-Path $pf 'CMake\bin\cmake.exe')
    }
    # Anything shipped inside a Visual Studio installation.
    foreach ($vs in (Get-VsInstallPaths)) {
        $candidates += Join-Path $vs 'Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin\cmake.exe'
    }

    $found = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $found) {
        throw "CMake not found. Install it from https://cmake.org/download/ or add it to PATH."
    }
    Write-Host "Found CMake at $found" -ForegroundColor DarkGray
    $env:PATH = "$([System.IO.Path]::GetDirectoryName($found));$env:PATH"
}

# Program Files roots that actually exist on this machine.
function Get-ProgramFilesRoots {
    $roots = @()
    foreach ($v in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432)) {
        if ($v -and (Test-Path $v) -and ($roots -notcontains $v)) { $roots += $v }
    }
    return $roots
}

# All Visual Studio installation roots, newest first (via vswhere when present).
function Get-VsInstallPaths {
    $paths = @()

    foreach ($pf in (Get-ProgramFilesRoots)) {
        $vswhere = Join-Path $pf 'Microsoft Visual Studio\Installer\vswhere.exe'
        if (-not (Test-Path $vswhere)) { continue }
        $out = & $vswhere -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
                          -latest -prerelease -property installationPath 2>$null
        foreach ($line in @($out)) {
            if ($line -and ($paths -notcontains $line.Trim())) { $paths += $line.Trim() }
        }
        if ($paths) { break }
    }

    if (-not $paths) {
        foreach ($pf in (Get-ProgramFilesRoots)) {
            foreach ($year in @('2022', '2019')) {
                foreach ($edition in @('Enterprise', 'Professional', 'Community', 'BuildTools')) {
                    $p = Join-Path $pf "Microsoft Visual Studio\$year\$edition"
                    if (Test-Path $p) { $paths += $p }
                }
            }
        }
    }
    return $paths
}

# Ninja needs cl.exe on PATH; a plain PowerShell window does not have it.
# Imports the MSVC developer environment into this process if needed.
function Initialize-MsvcEnvironment {
    if (Get-Command cl.exe -ErrorAction SilentlyContinue) { return $true }

    foreach ($vs in (Get-VsInstallPaths)) {
        $devShell = Join-Path $vs 'Common7\Tools\Microsoft.VisualStudio.DevShell.dll'
        if (Test-Path $devShell) {
            Write-Host "Importing MSVC environment from $vs" -ForegroundColor DarkGray
            Import-Module $devShell -ErrorAction SilentlyContinue
            try {
                Enter-VsDevShell -VsInstallPath $vs -SkipAutomaticLocation `
                                 -DevCmdArguments '-arch=x64 -host_arch=x64' | Out-Null
            } catch {
                Write-Host "Could not enter the VS developer shell: $($_.Exception.Message)" -ForegroundColor Yellow
            }
            if (Get-Command cl.exe -ErrorAction SilentlyContinue) { return $true }
        }
    }
    return $false
}

# Picks Ninja when it is usable, otherwise the Visual Studio generator.
function Resolve-Generator {
    param([string]$Requested)

    if ($Requested) { return $Requested }

    if (Get-Command ninja -ErrorAction SilentlyContinue) {
        if (Initialize-MsvcEnvironment) { return 'Ninja' }
        Write-Host "Ninja found but the MSVC compiler is not available; using the Visual Studio generator instead." -ForegroundColor Yellow
    }
    return 'Visual Studio 17 2022'
}

# Single-config generators put binaries in build/, multi-config ones in build/<Config>/.
function Get-OutputDir {
    param(
        [Parameter(Mandatory)][string]$BuildDir,
        [Parameter(Mandatory)][string]$Generator,
        [Parameter(Mandatory)][string]$Configuration
    )
    if ($Generator -like 'Visual Studio*' -or $Generator -like 'Xcode*') {
        return (Join-Path $BuildDir $Configuration)
    }
    return $BuildDir
}

function Get-ExePath {
    param(
        [Parameter(Mandatory)][string]$BuildDir,
        [Parameter(Mandatory)][string]$Generator,
        [Parameter(Mandatory)][string]$Configuration
    )
    return (Join-Path (Get-OutputDir -BuildDir $BuildDir -Generator $Generator -Configuration $Configuration) 'Just Notepad.exe')
}

# True when the cache was made by a different generator, or when the build
# system is older than CMakeLists.txt (a new source file was added, etc.).
function Test-NeedsConfigure {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$BuildDir,
        [Parameter(Mandatory)][string]$Generator
    )

    $cache = Join-Path $BuildDir 'CMakeCache.txt'
    if (-not (Test-Path $cache)) { return $true }

    $cachedGenerator = $null
    $match = Select-String -Path $cache -Pattern '^CMAKE_GENERATOR:INTERNAL=(.*)$' | Select-Object -First 1
    if ($match) { $cachedGenerator = $match.Matches.Groups[1].Value.Trim() }
    if ($cachedGenerator -and $cachedGenerator -ne $Generator) {
        Write-Host "Cache was generated with '$cachedGenerator' but '$Generator' was requested; reconfiguring." -ForegroundColor Yellow
        return $true
    }

    # CMakeLists.txt files newer than the cache mean targets/sources changed.
    $cacheTime = (Get-Item $cache).LastWriteTimeUtc
    $lists = @(Get-ChildItem -Path $Root -Filter 'CMakeLists.txt' -Recurse -ErrorAction SilentlyContinue |
               Where-Object { $_.FullName -notlike "$BuildDir*" })
    foreach ($l in $lists) {
        if ($l.LastWriteTimeUtc -gt $cacheTime) {
            Write-Host "$($l.FullName) changed since the last configure; reconfiguring." -ForegroundColor Yellow
            return $true
        }
    }
    return $false
}

function Invoke-CMakeConfigure {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$BuildDir,
        [Parameter(Mandatory)][string]$Generator,
        [Parameter(Mandatory)][string]$Configuration
    )

    $cmakeArgs = @('-S', $Root, '-B', $BuildDir, '-G', $Generator)
    if ($Generator -like 'Visual Studio*') {
        $cmakeArgs += @('-A', 'x64')
    } else {
        $cmakeArgs += "-DCMAKE_BUILD_TYPE=$Configuration"
    }

    Write-Host "Configuring: cmake $($cmakeArgs -join ' ')" -ForegroundColor Yellow
    & cmake @cmakeArgs
    if ($LASTEXITCODE -ne 0) { throw "CMake configure failed (exit code $LASTEXITCODE)." }
}

function Invoke-CMakeBuild {
    param(
        [Parameter(Mandatory)][string]$BuildDir,
        [Parameter(Mandatory)][string]$Configuration,
        [int]$Jobs = 0
    )

    $cmakeArgs = @('--build', $BuildDir, '--config', $Configuration)
    if ($Jobs -gt 0) { $cmakeArgs += @('--parallel', $Jobs) }
    else { $cmakeArgs += '--parallel' }

    Write-Host "Building Just Notepad and all plugins..." -ForegroundColor Cyan
    & cmake @cmakeArgs
    if ($LASTEXITCODE -ne 0) { throw "Build failed (exit code $LASTEXITCODE)." }
}

function Show-PluginSummary {
    param([Parameter(Mandatory)][string]$OutputDir)

    $pluginsDir = Join-Path $OutputDir 'plugins'
    if (Test-Path $pluginsDir) {
        $plugins = @(Get-ChildItem -Path $pluginsDir -Filter '*.dll' -ErrorAction SilentlyContinue)
        Write-Host "Built $($plugins.Count) plugin(s)." -ForegroundColor Green
    } else {
        Write-Host "Warning: no plugins directory at $pluginsDir" -ForegroundColor Yellow
    }
}

function Invoke-DependencyScripts {
    param([Parameter(Mandatory)][string]$Root)

    $pluginRoot = Join-Path $Root 'plugins'
    if (-not (Test-Path $pluginRoot)) { return }

    $scripts = @(Get-ChildItem -Path $pluginRoot -Recurse -Filter 'download_*.ps1' -ErrorAction SilentlyContinue)
    if (-not $scripts) { return }

    Write-Host "Fetching plugin dependencies..." -ForegroundColor Cyan
    foreach ($script in $scripts) {
        Write-Host "  $($script.Name)" -ForegroundColor DarkGray
        try {
            & $script.FullName
            if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
                Write-Host "  $($script.Name) exited with $LASTEXITCODE; continuing." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "  $($script.Name) failed: $($_.Exception.Message); continuing." -ForegroundColor Yellow
        }
    }
}

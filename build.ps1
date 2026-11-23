# Builds JustNotepad with the Visual Studio C++ compiler. No CMake needed.
#
#   .\build.ps1          release build  -> build\JustNotepad.exe
#   .\build.ps1 -Run     build, then start the editor
#   .\build.ps1 -Test    build and run the core tests and benchmark
#   .\build.ps1 -Debug   unoptimized build with debug symbols
param(
    [switch]$Run,
    [switch]$Test,
    [switch]$Debug
)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$out = Join-Path $root 'build'
New-Item -ItemType Directory -Force (Join-Path $out 'obj') | Out-Null
New-Item -ItemType Directory -Force (Join-Path $out 'test') | Out-Null

# Find the compiler unless this is already a Developer PowerShell / Command Prompt.
if (-not (Get-Command cl.exe -ErrorAction SilentlyContinue)) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path $vswhere)) {
        throw 'Visual Studio (or the Build Tools) with the C++ workload was not found.'
    }
    $vs = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
    if (-not $vs) { throw 'No Visual Studio installation with the C++ tools was found.' }
    $vcvars = Join-Path $vs 'VC\Auxiliary\Build\vcvars64.bat'
    cmd /c "`"$vcvars`" >nul 2>&1 && set" | ForEach-Object {
        if ($_ -match '^([^=]+)=(.*)$') { Set-Item -Path "env:$($matches[1])" -Value $matches[2] }
    }
}

$core = Get-ChildItem (Join-Path $root 'src\core\*.cpp') | ForEach-Object { $_.FullName }
$common = @('/nologo', '/std:c++20', '/utf-8', '/W3', '/EHsc', '/DUNICODE', '/D_UNICODE', '/DNOMINMAX')
if ($Debug) {
    $opt = @('/Od', '/Zi', '/MTd')
    $link = @('/DEBUG')
} else {
    # /MT links the C++ runtime statically: one exe, no redistributable to install.
    $opt = @('/O2', '/GL', '/Gw', '/Gy', '/MT', '/DNDEBUG')
    $link = @('/LTCG', '/OPT:REF', '/OPT:ICF')
}
$objDir = ($out -replace '\\', '/')

if ($Test) {
    $exe = Join-Path $out 'core_tests.exe'
    & cl @common @opt "/Fo:$objDir/test/" "/Fd:$objDir/test/" "/Fe:$exe" (Join-Path $root 'tests\core_tests.cpp') @core /link @link
    if ($LASTEXITCODE) { exit $LASTEXITCODE }
    & $exe bench
    exit $LASTEXITCODE
}

$res = Join-Path $out 'app.res'
& rc /nologo /i (Join-Path $root 'docs') /i (Join-Path $root 'src\win') /fo $res (Join-Path $root 'src\win\app.rc')
if ($LASTEXITCODE) { exit $LASTEXITCODE }

$exe = Join-Path $out 'JustNotepad.exe'
# comdlg32 and shell32 are only needed once a dialog opens or a file is
# dropped, so they are delay-loaded to keep them out of the startup path.
& cl @common @opt "/Fo:$objDir/obj/" "/Fd:$objDir/obj/" "/Fe:$exe" (Join-Path $root 'src\win\main.cpp') (Join-Path $root 'src\win\vcs.cpp') (Join-Path $root 'src\win\updater.cpp') (Join-Path $root 'src\win\theme.cpp') @core $res `
    /link @link /SUBSYSTEM:WINDOWS /MANIFEST:NO user32.lib gdi32.lib comdlg32.lib shell32.lib winhttp.lib advapi32.lib comctl32.lib uxtheme.lib dwmapi.lib gdiplus.lib ole32.lib delayimp.lib `
    /DELAYLOAD:comdlg32.dll /DELAYLOAD:shell32.dll
if ($LASTEXITCODE) { exit $LASTEXITCODE }

Write-Host "Built $exe ($([math]::Round((Get-Item $exe).Length / 1KB)) KB)"
if ($Run) { & $exe }

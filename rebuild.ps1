param(
    [ValidateSet('Debug','Release')]
    [string]$Configuration = 'Release',
    [string]$Generator = '',
    [switch]$BuildOnly
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'common.ps1')

$Root = $PSScriptRoot

Invoke-DependencyScripts -Root $Root

& (Join-Path $Root 'clean.ps1')

$buildArgs = @{
    Configuration = $Configuration
    Generator     = $Generator
    Reconfigure   = $true
}
if ($BuildOnly) { $buildArgs['BuildOnly'] = $true }

& (Join-Path $Root 'build.ps1') @buildArgs

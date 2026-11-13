<#
.SYNOPSIS
  Builds the solution and runs the tests.

.EXAMPLE
  ./build.ps1            # build + test (Debug)
  ./build.ps1 -Release   # build + test (Release), add-in .xll lands in src/ERechnung.ExcelAddIn/bin/Release
  ./build.ps1 -NoTest
#>
param([switch]$Release, [switch]$NoTest)

$ErrorActionPreference = 'Stop'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:DOTNET_NOLOGO = '1'

# Find a dotnet with an SDK: PATH first, then the per-user install used on dev machines
# (installed via: dotnet-install.ps1 -Channel 8.0 -InstallDir $env:LOCALAPPDATA\Microsoft\dotnet-sdk).
$candidates = @('dotnet', (Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet-sdk\dotnet.exe'))
$dotnet = $null
foreach ($c in $candidates) {
    try { if ((& $c --list-sdks 2>$null | Measure-Object).Count -gt 0) { $dotnet = $c; break } } catch { }
}
if (-not $dotnet) { throw 'No .NET SDK found. Install .NET 8 SDK (see AGENTS.md).' }

$cfg = if ($Release) { 'Release' } else { 'Debug' }
Push-Location $PSScriptRoot
try {
    & $dotnet build ERechnung.sln -c $cfg -nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed ($LASTEXITCODE)" }
    if (-not $NoTest) {
        & $dotnet test ERechnung.sln -c $cfg --no-build -nologo
        if ($LASTEXITCODE -ne 0) { throw "Tests failed ($LASTEXITCODE)" }
    }
} finally { Pop-Location }

<#
.SYNOPSIS
  Installs the official KoSIT validator (with the XRechnung configuration) and a
  portable Java runtime into a per-user folder, so the add-in can validate
  invoices against the official Schematron rules.

.DESCRIPTION
  Target folder (default): %LOCALAPPDATA%\ERechnungExcel\kosit
    jre\            portable Eclipse Temurin JRE (only if -SkipJre is not given)
    validator.jar   KoSIT validator (standalone jar)
    config\         XRechnung validator configuration (scenarios.xml + resources)
    versions.txt    which versions were installed

  The add-in looks in this folder automatically. Nothing is installed system-wide,
  no admin rights are needed. Re-run the script to update to newer releases.

.PARAMETER ValidatorVersion
  Tag of https://github.com/itplr-kosit/validator, e.g. "v1.6.3". Default: latest.

.PARAMETER ConfigVersion
  Tag of https://github.com/itplr-kosit/validator-configuration-xrechnung. Default: latest.
#>
[CmdletBinding()]
param(
    [string]$TargetDir = (Join-Path $env:LOCALAPPDATA 'ERechnungExcel\kosit'),
    [string]$ValidatorVersion = 'latest',
    [string]$ConfigVersion = 'latest',
    [int]$JavaMajor = 21,
    [switch]$SkipJre
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # Invoke-WebRequest is very slow with the progress bar
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Get-Release([string]$repo, [string]$tag) {
    $url = if ($tag -eq 'latest') { "https://api.github.com/repos/$repo/releases/latest" }
           else { "https://api.github.com/repos/$repo/releases/tags/$tag" }
    Invoke-RestMethod $url -UseBasicParsing -Headers @{ 'User-Agent' = 'setup-kosit' }
}

New-Item -ItemType Directory -Force $TargetDir | Out-Null
$tmp = Join-Path $TargetDir '_download'
New-Item -ItemType Directory -Force $tmp | Out-Null

# --- KoSIT validator --------------------------------------------------------
$rel = Get-Release 'itplr-kosit/validator' $ValidatorVersion
$asset = $rel.assets | Where-Object { $_.name -like '*-standalone.jar' } | Select-Object -First 1
if (-not $asset) { throw "No standalone jar found in validator release $($rel.tag_name)" }
Write-Host "Validator  $($rel.tag_name): $($asset.browser_download_url)"
Invoke-WebRequest $asset.browser_download_url -OutFile (Join-Path $TargetDir 'validator.jar') -UseBasicParsing
$validatorTag = $rel.tag_name

# --- XRechnung configuration -----------------------------------------------
$rel = Get-Release 'itplr-kosit/validator-configuration-xrechnung' $ConfigVersion
$asset = $rel.assets | Where-Object { $_.name -like '*.zip' } | Select-Object -First 1
if (-not $asset) { throw "No zip found in configuration release $($rel.tag_name)" }
Write-Host "XRechnung config $($rel.tag_name): $($asset.browser_download_url)"
$zip = Join-Path $tmp 'config.zip'
Invoke-WebRequest $asset.browser_download_url -OutFile $zip -UseBasicParsing
$configDir = Join-Path $TargetDir 'config'
if (Test-Path $configDir) { Remove-Item $configDir -Recurse -Force }
Expand-Archive $zip -DestinationPath $configDir
# Some releases wrap everything in one sub-folder: flatten so scenarios.xml sits in config\
$scen = Get-ChildItem $configDir -Recurse -Filter 'scenarios.xml' | Select-Object -First 1
if (-not $scen) { throw 'scenarios.xml not found in the configuration archive' }
if ($scen.DirectoryName -ne (Resolve-Path $configDir).Path) {
    Get-ChildItem $scen.DirectoryName -Force | Move-Item -Destination $configDir -Force
}
$configTag = $rel.tag_name

# --- portable Java runtime -------------------------------------------------
$javaInfo = 'system java'
if (-not $SkipJre) {
    $jreDir = Join-Path $TargetDir 'jre'
    $jreZip = Join-Path $tmp 'jre.zip'
    $jreUrl = "https://api.adoptium.net/v3/binary/latest/$JavaMajor/ga/windows/x64/jre/hotspot/normal/eclipse?project=jdk"
    Write-Host "Java runtime: Temurin $JavaMajor JRE"
    Invoke-WebRequest $jreUrl -OutFile $jreZip -UseBasicParsing
    $jreTmp = Join-Path $tmp 'jre'
    Expand-Archive $jreZip -DestinationPath $jreTmp -Force
    $inner = Get-ChildItem $jreTmp -Directory | Select-Object -First 1   # e.g. jdk-21.0.5+11-jre
    if (Test-Path $jreDir) { Remove-Item $jreDir -Recurse -Force }
    Move-Item $inner.FullName $jreDir
    $javaInfo = "Temurin $($inner.Name)"
}

Remove-Item $tmp -Recurse -Force
@(
    "validator=$validatorTag"
    "xrechnung-config=$configTag"
    "java=$javaInfo"
    "installed=$(Get-Date -Format s)"
) | Set-Content (Join-Path $TargetDir 'versions.txt') -Encoding UTF8

Write-Host "KoSIT validator installed to $TargetDir"

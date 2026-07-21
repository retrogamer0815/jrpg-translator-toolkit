param(
    [string]$LaunchBoxRoot = $env:LAUNCHBOX_ROOT,
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [string]$Version = "0.1.0-preview",
    [switch]$IncludeSymbols,
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path

& (Join-Path $projectDir "build.ps1") `
    -LaunchBoxRoot $LaunchBoxRoot `
    -Configuration $Configuration `
    -SkipTests:$SkipTests

$outputDir = Join-Path $projectDir "bin\$Configuration\net9.0-windows"
$distDir = Join-Path $projectDir "dist"
$packageDir = Join-Path $distDir "JRPG Translator Integration"
$archivePath = Join-Path $distDir "JRPG_Translator_LaunchBox_Plugin_v$Version.zip"

if (-not (Test-Path -LiteralPath (Join-Path $outputDir "JrpgTranslator.LaunchBox.dll"))) {
    throw "Plugin output was not found under $outputDir"
}

if (Test-Path -LiteralPath $packageDir) {
    Remove-Item -LiteralPath $packageDir -Recurse -Force
}
New-Item -ItemType Directory -Path $packageDir -Force | Out-Null

Copy-Item -LiteralPath (Join-Path $outputDir "JrpgTranslator.LaunchBox.dll") -Destination $packageDir
$depsFile = Join-Path $outputDir "JrpgTranslator.LaunchBox.deps.json"
if (Test-Path -LiteralPath $depsFile) {
    Copy-Item -LiteralPath $depsFile -Destination $packageDir
}
if ($IncludeSymbols) {
    $pdbFile = Join-Path $outputDir "JrpgTranslator.LaunchBox.pdb"
    if (Test-Path -LiteralPath $pdbFile) {
        Copy-Item -LiteralPath $pdbFile -Destination $packageDir
    }
}
Copy-Item -LiteralPath (Join-Path $projectDir "README.md") -Destination $packageDir

if (Test-Path -LiteralPath $archivePath) {
    Remove-Item -LiteralPath $archivePath -Force
}
Compress-Archive -LiteralPath $packageDir -DestinationPath $archivePath -CompressionLevel Optimal

Write-Host "Created $archivePath"

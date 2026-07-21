param(
    [string]$LaunchBoxRoot = $env:LAUNCHBOX_ROOT,
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dotnet = $null
$searchRoot = $projectDir
while ($searchRoot) {
    $candidate = Join-Path $searchRoot ".dotnet-sdk\dotnet.exe"
    if (Test-Path -LiteralPath $candidate) {
        $dotnet = $candidate
        break
    }
    $parent = Split-Path -Parent $searchRoot
    if (-not $parent -or $parent -eq $searchRoot) {
        break
    }
    $searchRoot = $parent
}
if (-not $dotnet) {
    $dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($dotnetCommand) {
        $dotnet = $dotnetCommand.Source
    }
}

if (-not $LaunchBoxRoot) {
    $LaunchBoxRoot = @(
        (Join-Path $env:USERPROFILE "LaunchBox"),
        (Join-Path ${env:ProgramFiles} "LaunchBox")
    ) | Where-Object {
        Test-Path -LiteralPath (Join-Path $_ "Core\Unbroken.LaunchBox.Plugins.dll")
    } | Select-Object -First 1
}

$env:DOTNET_CLI_HOME = Join-Path $projectDir ".dotnet-home"
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = "1"
$env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"

if (-not $dotnet) {
    throw "A .NET 9 SDK was not found. Install it or place a workspace SDK in .dotnet-sdk."
}

if (-not $LaunchBoxRoot -or -not (Test-Path -LiteralPath (Join-Path $LaunchBoxRoot "Core\Unbroken.LaunchBox.Plugins.dll"))) {
    throw "LaunchBox plugin API not found. Pass -LaunchBoxRoot or set LAUNCHBOX_ROOT."
}

& $dotnet build (Join-Path $projectDir "JrpgTranslatorLaunchBox.csproj") `
    --nologo `
    --configuration $Configuration `
    --property:LaunchBoxRoot=$LaunchBoxRoot

if ($LASTEXITCODE -ne 0) {
    throw "Plugin build failed with exit code $LASTEXITCODE"
}

if (-not $SkipTests) {
    & $dotnet run --project (Join-Path $projectDir "Tests\SmokeTest.csproj") `
        --configuration $Configuration `
        --property:LaunchBoxRoot=$LaunchBoxRoot

    if ($LASTEXITCODE -ne 0) {
        throw "Plugin smoke test failed with exit code $LASTEXITCODE"
    }
}

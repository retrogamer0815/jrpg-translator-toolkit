param([string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-plugin-test-' + [Guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$sources = @(
    (Join-Path $repo 'integrations/launchbox/OwnedProcess.cs'),
    (Join-Path $repo 'integrations/launchbox/GameRuntimePlugin.cs'),
    (Join-Path $PSScriptRoot 'launchbox_runtime_stubs.cs'),
    (Join-Path $PSScriptRoot 'launchbox_ownership_harness.cs')
)
Add-Type -Path $sources -CompilerOptions '/nullable:enable'
[LaunchBoxOwnershipHarness]::Run('C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe', (Join-Path $PSScriptRoot 'launchbox_ownership_helper.ahk'), $output.FullName)

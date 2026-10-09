param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-capture-tests-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$source = [IO.File]::ReadAllText((Join-Path $repo 'bin/overlay.ahk')).Replace("`r`n", "`n")
function Get-Function([string]$name) {
    $body = [regex]::Match($source, '(?ms)^' + $name + '\([^\n]*\)\s*\{.*?^\}').Value
    if (!$body) { throw "Missing production function: $name" }
    return $body
}
foreach ($line in @('OnMessage(0x0201, Region_LButtonDown)', 'OnMessage(0x0202, Region_LButtonUp)')) {
    if ([regex]::Matches($source, [regex]::Escape($line)).Count -ne 1) { throw "Expected one persistent monitor: $line" }
}
if ((Get-Function 'CancelAnyPick').Contains('OnMessage(')) { throw 'Do not unregister picker monitors inside mouse callbacks.' }
$generated = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'capture_picker_harness.ahk'))
$globals = [regex]::Match($source, '(?ms)^global __GDI_Ready :=.*?^global __WindowHighlightGui :=[^\n]*').Value
if (!$globals) { throw 'Picker runtime globals were not found.' }
$generated = $generated.Replace('; @RUNTIME_GLOBALS@', $globals)
$names = @(
    'CapPickRun', 'CapPickFlag', 'CapPickHotIf', 'CapPickArrow', 'CapPickArrowCore',
    'CapPickConfirm', 'CapPickConfirmCore', 'CapPickCancel', 'CapPickModifierKey',
    'CapPickSignalCompletion', 'CapPickCreateHud', 'CapPickPositionHud', 'CapPickPlaceRegionAboveHud',
    'CapPickUpdateHud', 'CapPickStartState', 'CapPickEndSession', 'CancelCapturePick',
    'CapPickTick', 'CapPickTickCore', 'CapPickWholePixels', 'CapPickVelocity',
    'CapPickInitializeButtonStates', 'CapPickHandlePendingControllerButtons', 'CapPickControllerKey',
    'CancelAnyPick', 'BeginPickHide', 'EndPickHide', 'CapPickFailSafe',
    'CapPickClampRegionRect', 'CapPickApplyRegionRect', 'CapPickInitialRegionRect',
    'StartPickRegion', 'StartPickRegionCore', 'RegionMessageIsCurrent', 'RegionMessagePoint',
    'Region_LButtonDown', 'RegionBeginDrag', 'Region_LButtonUp', 'RegionEndDrag',
    'DrawBand', 'DrawBandCore', 'DestroySelOverlay', 'FinishRegionPick', 'FinishRegionPickRect',
    'StartPickWindow', 'StartPickWindowCore', 'PulseHover', 'PulseHoverCore',
    'PickWindowClick', 'PickWindowClickCore', 'FinishWindowCandidate',
    'CapPickCandidateIndex', 'CapPickSetWindowCandidate', 'CapPickCycleWindow'
)
foreach ($name in $names) {
    $body = Get-Function $name
    # Isolate only desktop/hardware boundaries: all picker lifecycle and native
    # mouse-message code is production code, using real but off-screen GUIs.
    $body = $body.Replace('MonitorGetWorkArea(', 'FixtureWorkArea(')
    $body = $body.Replace('ToolTip(', 'FixtureToolTip(')
    $body = $body.Replace('MouseGetPos', 'FixtureMouseGetPos')
    $body = $body.Replace('hud.Show("NA AutoSize")', 'hud.Show("NA AutoSize x-30000 y-30000")')
    $body = $body.Replace('g.Show("x" vsx', 'g.Show("NA x" vsx')
    $body = $body.Replace('CapPickCreateHud()', 'FixtureCreateHud()')
    if ($name -eq 'CapPickCreateHud') { $body = $body.Replace('FixtureCreateHud() {', 'CapPickCreateHud() {') }
    $body = $body.Replace('"Explainer"', '"Picker fixture other overlay"').Replace('"Translator"', '"Picker fixture translator"')
    $generated += "`n" + $body
}
$main = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$captureWatch = [regex]::Match($main, '(?ms)^CPBigBoxWatchCapture\([^\r\n]*\)\s*\{.*?^\}').Value
if (!$captureWatch) { throw 'Missing Big Box capture watcher.' }
$generated += "`n" + $captureWatch
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$script = Join-Path $output.FullName 'capture-picker-generated.ahk'
[IO.File]::WriteAllText($script, $generated, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'capture-picker.stdout.txt'
$stderr = Join-Path $output.FullName 'capture-picker.stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $script + '"')) -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
if (!$process.WaitForExit(30000)) { $process.Kill(); throw "Capture picker test timed out. Artifacts: $output" }
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or $result.Contains('Warning:') -or !$result.Contains('PASS:')) {
    throw "$result`n$errors`nArtifacts: $output"
}
Write-Output $result.Trim()
Write-Output "Test artifacts: $output"

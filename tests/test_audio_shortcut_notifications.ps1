param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$SourcePath,
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-audio-shortcuts-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
if (!$SourcePath) { $SourcePath = Join-Path $repo 'JRPG Translator.ahk' }
$source = [IO.File]::ReadAllText($SourcePath)
$generated = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'audio_shortcut_notifications_harness.ahk'))
foreach ($name in @('StartStopAudio', 'StartAudio', 'StopAudio', 'ToggleAudioFromButton', 'CPControllerDispatchAction')) {
    $body = [regex]::Match($source, '(?ms)^' + $name + '\([^\r\n]*\)\s*\{.*?^\}').Value
    if (!$body) { throw "Missing production function: $name" }
    $generated += "`r`n" + $body
}
$binding = [regex]::Match($source, '(?ms)^Rebind_StartStopAudio\([^\r\n]*\)\s*\{.*?^\}').Value
if (!$binding.Contains('SafeCall(StartStopAudio)')) { throw 'Configured keyboard/JoyToKey shortcut must reach the notification handler.' }
$toast = [regex]::Match($source, '(?ms)^Toast\([^\r\n]*\)\s*\{.*?^\}').Value
if (!$toast.Contains('NoActivate x20 y20') -or !$toast.Contains('+AlwaysOnTop') -or !$toast.Contains('CPToastGui.Opt("+E0x20")')) {
    throw 'Shortcut notification must reuse the top-left, non-activating, click-through toast.'
}
$constructor = [regex]::Match($toast, 'CPToastGui := Gui\([^\r\n]+').Value
if ($constructor.Contains('+E0x20') -or
    $toast.IndexOf('CPToastGui.Opt("+E0x20")') -lt $toast.IndexOf('CPToastText.Redraw()') -or
    !$toast.Contains('WinSetTransparent(255, CPToastGui.Hwnd)')) {
    throw 'Paint the opaque layered toast before enabling click-through; do not restore the Show-time paint stall.'
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$script = Join-Path $output.FullName 'audio-shortcuts-generated.ahk'
[IO.File]::WriteAllText($script, $generated, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'stdout.txt'
$stderr = Join-Path $output.FullName 'stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $script + '"')) -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
if (!$process.WaitForExit(30000)) { $process.Kill(); throw "Audio shortcut test timed out: $output" }
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or $result.Contains('Warning:') -or !$result.Contains('PASS:')) {
    throw "$result`n$errors`nArtifacts: $output"
}
Write-Output $result.Trim()
Write-Output 'PASS: hotkey binding and existing toast presentation guards.'
Write-Output "Test artifacts: $output"

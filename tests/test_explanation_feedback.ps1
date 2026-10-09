param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-explanation-feedback-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$main = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$overlay = [IO.File]::ReadAllText((Join-Path $repo 'bin/overlay.ahk'))
$generated = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'explanation_feedback_harness.ahk'))
foreach ($entry in @(
    @($main, @('ToastState', 'ToastDestroy', 'Toast', 'ToastPresent', 'ToastPosition', 'GetWindowDPI', 'CPColorRef')),
    @($overlay, @('IsWindowTopmost', 'ToggleTop'))
)) {
    foreach ($name in $entry[1]) {
        $body = [regex]::Match($entry[0], '(?ms)^' + $name + '\([^\r\n]*\)\s*\{.*?^\}').Value
        if (!$body) { throw "Missing production function: $name" }
        if ($name -eq 'ToggleTop') {
            if ($body.Contains('WinGetID("A")')) { throw 'Active-window lookup must tolerate full-screen transitions.' }
            $body = $body.Replace('WinExist("Explainer")', 'WinExist("Synthetic nonexistent feedback overlay")')
        }
        if ($name -eq 'ToastPosition') { $body = $body.Replace('ToastPosition(', 'ProductionToastPosition(') }
        $body = $body.Replace('DllCall("user32\GetForegroundWindow", "ptr")', 'TestForegroundWindow()')
        $generated += "`r`n" + $body
    }
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$script = Join-Path $output.FullName 'explanation-feedback-generated.ahk'
[IO.File]::WriteAllText($script, $generated, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'stdout.txt'
$stderr = Join-Path $output.FullName 'stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $script + '"')) -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
if (!$process.WaitForExit(30000)) { $process.Kill(); throw "Test timed out: $output" }
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or $result.Contains('Warning:') -or !$result.Contains('PASS:')) { throw "$result`n$errors`nArtifacts: $output" }
Write-Output $result.Trim()

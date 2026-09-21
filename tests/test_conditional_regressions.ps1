param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-condition-tests-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$mainSource = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$overlaySource = [IO.File]::ReadAllText((Join-Path $repo 'bin\overlay.ahk'))
$fallbackBlock = [regex]::Match($mainSource, '(?ms)^; safety: if INI was empty on first run\r?\n(.*?)^; --- Explainer bounds').Groups[1].Value
if (!$fallbackBlock) { throw 'Explanation model fallback block not found' }
$generated = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'conditional_regressions_harness.ahk'))
$generated = $generated.Replace('; @MODEL_FALLBACKS@', $fallbackBlock)
foreach ($name in @('RefreshAllBg', 'EraseAnyBg')) {
    # The blank line/function boundary also handles the old unindented blocks.
    $body = [regex]::Match($overlaySource, '(?ms)^' + $name + '\([^\r\n]*\)\s*\{.*?^\}\r?\n\r?\n(?=[A-Za-z_]|;|$)').Value
    if (!$body) { throw "Overlay function missing: $name" }
    # Only replace OS boundaries, keeping the actual production control flow.
    $generated += "`r`n" + $body.Replace('DllCall(', 'TestDllCall(').Replace('MakeBrush(', 'TestMakeBrush(')
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$script = Join-Path $output.FullName 'conditions-generated.ahk'
[IO.File]::WriteAllText($script, $generated, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'conditions.stdout.txt'
$stderr = Join-Path $output.FullName 'conditions.stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $script + '"')) -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
if (!$process.WaitForExit(30000)) { $process.Kill(); throw "Conditional regression test timed out. Artifacts: $output" }
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or $result.Contains('Warning:') -or !$result.Contains('PASS:')) {
    throw "$result`n$errors`nArtifacts: $output"
}
# Prevent this specific unsupported single-line-if assignment pattern from
# returning, without flagging intentional assignments inside conditions.
foreach ($file in @('JRPG Translator.ahk', 'bin\overlay.ahk', 'launchers\JRPG Study Launcher.ahk')) {
    $matches = Select-String -LiteralPath (Join-Path $repo $file) -Pattern '^\s*if\b.*\)\s+[A-Za-z_]\w*\s*:='
    if ($matches) { throw "Single-line if assignment found: $matches" }
}
Write-Output $result.Trim()
Write-Output 'PASS: single-line conditional assignment guard.'
Write-Output "Test artifacts: $output"

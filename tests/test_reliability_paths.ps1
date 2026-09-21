param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-reliability-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$main = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$overlay = [IO.File]::ReadAllText((Join-Path $repo 'bin/overlay.ahk'))
$generated = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'reliability_paths_harness.ahk'))
foreach ($entry in @(
    @($main, @('SaveApiEnv', 'CPReplaceApiEnvFile', 'ExplainNow', 'ExplainNowCore', 'StudyWindowRevealFinished', 'StudyWindowRevealCore', 'StudyWindowHasHandle')),
    @($overlay, @('flushTranslate', 'FlushTranslateCore', 'RemoveAcceptedCaptures', 'FlushBufferedScreenshots'))
)) {
    foreach ($name in $entry[1]) {
        $body = [regex]::Match($entry[0], '(?ms)^' + $name + '\([^\r\n]*\)\s*\{.*?^\}').Value
        if (!$body) { throw "Missing production function: $name" }
        if ($name -eq 'SaveApiEnv') { $body = $body.Replace('CPReplaceApiEnvFile(tmp, envPath)', 'TestReplaceApiEnvFile(tmp, envPath)') }
        if ($name -eq 'ExplainNowCore') {
            if (!$body.Contains('try Run(cmd, A_ScriptDir, "Hide", &pid)')) { throw 'Explanation launch boundary changed' }
            $body = $body.Replace('try Run(cmd, A_ScriptDir, "Hide", &pid)', 'try TestExplainRun(cmd, A_ScriptDir, "Hide", &pid)').Replace('while ProcessExist(pid)', 'while TestExplainAlive(pid)')
        }
        if ($name -eq 'FlushTranslateCore') {
            if (!$body.Contains('try Run(cmd, A_ScriptDir, "Hide", &__TranslationPid)')) { throw 'Translation launch boundary changed' }
            $body = $body.Replace('try Run(cmd, A_ScriptDir, "Hide", &__TranslationPid)', 'try TestTranslateRun(cmd, A_ScriptDir, "Hide", &__TranslationPid)')
        }
        $body = $body.Replace('ToolTip(', 'TestTip(')
        $generated += "`r`n" + $body
    }
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$script = Join-Path $output.FullName 'reliability-generated.ahk'
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

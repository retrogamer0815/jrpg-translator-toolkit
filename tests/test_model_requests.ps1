param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-model-requests-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$source = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$script = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'model_request_harness.ahk'))
$overlay = [IO.File]::ReadAllText((Join-Path $repo 'bin/overlay.ahk'))
$script += "`r`n" + [regex]::Match($overlay, '(?ms)^WriteTranslationTerminalResult\(.*?^}').Value
$script += "`r`n" + [regex]::Match($source, '(?ms)^StudyCandidatesRunBridge\(.*?^}').Value.Replace('CPModelRequestRun(', 'TestBridgeRun(')
foreach ($match in [regex]::Matches($source, '(?ms)^(?:CPModelRequest\w+|CPModelFailureSummary|CPTranslationFailureStatus|CPPollTranslationStatus)\(.*?^}')) {
    $script += "`r`n" + $match.Value.Replace('DllCall("user32\MessageBeep", "uint", 0x30)', 'TestWarningCue()')
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'model_request_worker.ahk') -Destination $output.FullName
$path = Join-Path $output.FullName 'generated.ahk'
[IO.File]::WriteAllText($path, $script, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'stdout.txt'
$stderr = Join-Path $output.FullName 'stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $path + '"')) -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
if (!$process.WaitForExit(20000)) { $process.Kill(); throw "Harness timeout: $output" }
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or !$result.Contains('PASS:')) { throw "$result`n$errors`nArtifacts: $output" }
Write-Output $result.Trim()

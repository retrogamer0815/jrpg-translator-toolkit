param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-profile-delete-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$source = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$generated = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'profile_delete_notifications_harness.ahk'))
foreach ($name in @('DeleteSelectedGameProfile', 'GameProfilePath', 'GameProfileSafeName',
    'DeletePromptProfile', 'DeleteExplainPromptProfile', 'StudyReaderDeleteNewVersionPrompt',
    'PromptFilePath', 'ExplainProfilePath')) {
    $body = [regex]::Match($source, '(?ms)^' + $name + '\([^\r\n]*\)\s*\{.*?^\}').Value
    if (!$body) { throw "Missing production function: $name" }
    $generated += "`r`n" + $body
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$script = Join-Path $output.FullName 'profile-delete-generated.ahk'
[IO.File]::WriteAllText($script, $generated, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'stdout.txt'
$stderr = Join-Path $output.FullName 'stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $script + '"')) -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
if (!$process.WaitForExit(30000)) { $process.Kill(); throw "Profile deletion test timed out: $output" }
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or $result.Contains('Warning:') -or !$result.Contains('PASS:')) {
    throw "$result`n$errors`nArtifacts: $output"
}
Write-Output $result.Trim()
Write-Output "Test artifacts: $output"

param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-anki-results-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$source = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$generated = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'anki_results_harness.ahk'))
foreach ($name in @('StudyAnkiReadAddResult', 'StudyAnkiShowUncertainAdd')) {
    $body = [regex]::Match($source, '(?ms)^' + $name + '\([^\r\n]*\)\s*\{.*?^\}').Value
    if (!$body) { throw "Missing production function: $name" }
    $generated += "`r`n" + $body
}
foreach ($name in @('StudyReaderAddReviewedAnkiNote', 'StudyReaderGenerateVocabularyExample')) {
    $body = [regex]::Match($source, '(?ms)^' + $name + '\([^\r\n]*\)\s*\{.*?^\}').Value
    if (!$body.Contains('saAddState.Get("uncertain", false)')) { throw "Missing uncertain-result guard: $name" }
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$script = Join-Path $output.FullName 'anki-results-generated.ahk'
[IO.File]::WriteAllText($script, $generated, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'stdout.txt'
$stderr = Join-Path $output.FullName 'stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $script + '"')) -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
if (!$process.WaitForExit(30000)) { $process.Kill(); throw "Anki result test timed out: $output" }
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or $result.Contains('Warning:') -or !$result.Contains('PASS:')) {
    throw "$result`n$errors`nArtifacts: $output"
}
Write-Output $result.Trim()
Write-Output "Test artifacts: $output"

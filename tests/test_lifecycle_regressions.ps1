param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-lifecycle-tests-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$main = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$overlay = [IO.File]::ReadAllText((Join-Path $repo 'bin\overlay.ahk'))
function Get-TestFunction([string]$source, [string]$name) {
    $body = [regex]::Match($source, '(?ms)^' + $name + '\([^\r\n]*\)\s*\{.*?^\}').Value
    if (!$body) { $body = [regex]::Match($source, '(?m)^' + $name + '\([^\r\n]*\)\s*=>[^\r\n]*(?:\r?\n +[^\r\n]+)?').Value }
    if (!$body) { throw "Missing production function: $name" }
    return $body
}
$mainReader = Get-TestFunction $main 'ReadIniInt'
$overlayReader = Get-TestFunction $overlay 'ReadIniInt'
if ($mainReader.Replace("`r`n", "`n") -cne $overlayReader.Replace("`r`n", "`n")) {
    throw 'The standalone main and overlay integer readers must have identical validation.'
}
foreach ($source in @($main, $overlay)) {
    if ($source -match 'Integer\((IniRead|Load|LoadCfg)\(') {
        throw 'Numeric INI values must go through validated readers, not direct Integer conversions.'
    }
}
$generated = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'lifecycle_regressions_harness.ahk'))
foreach ($name in @('ReadIniInt', 'LoadInt', 'GameProfileReadInt', 'StudyLibraryStateAlive', 'StudyCandidatesGuiAlive', 'StudyReaderRestoreCopyButton', 'StudyWindowEnableLiveResize', 'StudyDialogFocusIfAlive')) {
    $generated += "`r`n" + (Get-TestFunction $main $name)
}
foreach ($name in @('LoadCfgInt', 'ClearOverlayOutputFile', 'ClearAudioFile', 'ClearOcrFile', 'ClearExplainerFile', 'TryReadOverlayFile', 'PollAudioSubtitle', 'PollOcrFile', 'PollExplainerFile', 'LoadOverlayBounds')) {
    $generated += "`r`n" + (Get-TestFunction $overlay $name)
}
# Test the actual startup assignments as well as the reader in isolation.
$mainSettings = @()
foreach ($name in @('fontSize', 'fontSize_EW', 'overlayTrans', 'overlayTrans_EW', 'capMaxKB', 'ewX', 'ewY', 'ewW', 'ewH')) {
    $line = [regex]::Match($main, '(?m)^' + $name + '\s*:= [^\r\n]+').Value
    if (!$line) { throw "Missing main startup assignment: $name" }
    $mainSettings += $line
}
$overlaySettings = @()
foreach ($name in @('FONT_SIZE', 'FONT_BOLD', 'Cap_MaxKB', 'IsTop')) {
    $line = [regex]::Match($overlay, '(?m)^global ' + $name + '\s*:= [^\r\n]+').Value
    if (!$line) { throw "Missing overlay startup assignment: $name" }
    $overlaySettings += $line -replace '^global ', ''
}
$generated = $generated.Replace('; @MAIN_SETTINGS@', ($mainSettings -join "`r`n"))
$generated = $generated.Replace('; @OVERLAY_SETTINGS@', ($overlaySettings -join "`r`n"))
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$script = Join-Path $output.FullName 'lifecycle-generated.ahk'
[IO.File]::WriteAllText($script, $generated, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'lifecycle.stdout.txt'
$stderr = Join-Path $output.FullName 'lifecycle.stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $script + '"')) -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
if (!$process.WaitForExit(30000)) { $process.Kill(); throw "Lifecycle test timed out. Artifacts: $output" }
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or $result.Contains('Warning:') -or !$result.Contains('PASS:')) {
    throw "$result`n$errors`nArtifacts: $output"
}
Write-Output $result.Trim()
Write-Output 'PASS: numeric reader parity and direct-conversion guard.'
Write-Output "Test artifacts: $output"

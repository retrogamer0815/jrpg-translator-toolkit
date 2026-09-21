param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-bounds-tests-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$source = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$generated = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'explainer_bounds_harness.ahk'))
foreach ($name in @('SaveExplainerBoundsIfChanged', 'StartExplainerBoundsWatcher', 'StopExplainerBoundsWatcher')) {
    $body = [regex]::Match($source, '(?ms)^' + $name + '\(\)\s*\{.*?^\}').Value
    if (!$body) { throw "Production function missing: $name" }
    if ($name -eq 'SaveExplainerBoundsIfChanged') {
        # Keep the production control flow intact. Inject only OS/file boundaries
        # so races and failed writes are deterministic and never touch user data.
        $testBody = $body.Replace('WinExist(', 'TestWinExist(').Replace('WinGetPos ', 'TestWinGetPos ').Replace('IniWrite(', 'TestIniWrite(')
        $generated += "`r`n" + $testBody
        # A second copy exercises real Windows geometry and INI persistence,
        # using a unique caption which cannot match the user's Explainer.
        $generated += "`r`n" + $body.Replace($name + '()', 'TestRealBoundsSave()').Replace('WinExist("Explainer")', 'WinExist(TestWindowTitle)')
    } else {
        $generated += "`r`n" + $body
    }
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$script = Join-Path $output.FullName 'bounds-generated.ahk'
[IO.File]::WriteAllText($script, $generated, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'bounds.stdout.txt'
$stderr = Join-Path $output.FullName 'bounds.stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $script + '"')) -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
if (!$process.WaitForExit(20000)) {
    $process.Kill()
    throw "Explainer bounds test timed out. Artifacts: $output"
}
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or $result.Contains('Warning:') -or !$result.Contains('PASS:')) {
    throw "$result`n$errors`nArtifacts: $output"
}
Write-Output $result.Trim()
Write-Output "Test artifacts: $output"

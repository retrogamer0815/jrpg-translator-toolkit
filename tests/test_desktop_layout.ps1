param(
    [switch]$StudyOnly,
    [switch]$DialogsOnly,
    [switch]$ShutdownOnly,
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-desktop-tests-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
if (@($StudyOnly, $DialogsOnly, $ShutdownOnly).Where({ [bool]$_ }).Count -gt 1) {
    throw 'Choose only one of StudyOnly, DialogsOnly or ShutdownOnly.'
}
$repo = Split-Path -Parent $PSScriptRoot
$source = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
New-Item -ItemType Directory -Path (Join-Path $output.FullName 'assets') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repo 'assets/bigbox-logo.png') -Destination (Join-Path $output.FullName 'assets')
Copy-Item -LiteralPath (Join-Path $repo 'assets/desktop-logo.png') -Destination (Join-Path $output.FullName 'assets')
Copy-Item -LiteralPath (Join-Path $repo 'JRPG Translator.ahk') -Destination $output.FullName
# Opaque, differently shaped captures exercise replacing the native Picture
# bitmap. These are synthetic test assets, not screenshots from a real library.
Add-Type -AssemblyName System.Drawing
foreach ($fixture in @(@('landscape', 720, 540), @('wide', 1600, 200), @('portrait', 240, 900))) {
    $bitmap = [Drawing.Bitmap]::new($fixture[1], $fixture[2])
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $halfW = [int]($fixture[1] / 2); $halfH = [int]($fixture[2] / 2)
        $graphics.FillRectangle([Drawing.Brushes]::Red, 0, 0, $halfW, $halfH)
        $graphics.FillRectangle([Drawing.Brushes]::Lime, $halfW, 0, $halfW, $halfH)
        $graphics.FillRectangle([Drawing.Brushes]::Blue, 0, $halfH, $halfW, $halfH)
        $graphics.FillRectangle([Drawing.Brushes]::Yellow, $halfW, $halfH, $halfW, $halfH)
        $bitmap.Save((Join-Path $output.FullName ('study-preview-' + $fixture[0] + '.bmp')), [Drawing.Imaging.ImageFormat]::Bmp)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}
# Initialize only literal UI state. Production startup is after ExitApp and
# never runs: no overlays, global hotkeys, personal settings, audio or API requests.
# Native wheel bursts are delivered only to the synthetic GUI and its controls.
$uiGlobals = [regex]::Matches($source.Substring(0, $source.IndexOf('CPRegisterCanvasMessages() {')), '(?m)^global CP[^\r\n]*') | ForEach-Object { $_.Value }
$harness = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'desktop_layout_harness.ahk'))
$captureSource = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'bigbox_navigation_harness.ahk'))
$capture = [regex]::Match($captureSource, '(?ms)^TestCapture\(.*?^}').Value.Replace('CPBigBoxGui', 'ui')
$capture = $capture.Replace('SendMessage(0x0317, dc, 0x1C, ui.Hwnd)', 'DllCall("user32\PrintWindow", "ptr", ui.Hwnd, "ptr", dc, "uint", 1)')
$generated = $harness.Replace('; @UI_GLOBALS@', ($uiGlobals -join "`r`n")).Replace('; @CAPTURE@', $capture)
# Exercise the actual Study control construction and event wiring without its
# bridge/database/startup calls. All paths/settings below belong to this fixture.
foreach ($kind in @('Library', 'Reader')) {
    $start = if ($kind -eq 'Library') { '    slGuiOptions := slBigBoxPresentation' } else { '    srReaderOptions := srWantBigBox' }
    $end = if ($kind -eq 'Library') { '    if slBigBoxPresentation {\r?\n        slBigBoxBounds :=' } else { '    if srWantBigBox {\r?\n        srBigBoxBounds :=' }
    $bodyStart = $source.IndexOf($start)
    $rest = $source.Substring($bodyStart)
    $bodyEnd = [regex]::Match($rest, $end).Index
    if ($bodyStart -lt 0 -or $bodyEnd -lt 1) { throw "Study $kind constructor markers not found" }
    $body = $rest.Substring(0, $bodyEnd).Replace('    StudyReaderBindHotkeys(srState)', '')
    $generated = $generated.Replace('; @STUDY_' + $kind.ToUpper() + '_CONTROLS@', $body)
}
# Only constructors are lifted; bridge calls and real Anki operations never run.
foreach ($kind in @('ANKI', 'CANDIDATES', 'CHAPTER', 'COLUMNS', 'FILTERS', 'DETAILS', 'MANAGER', 'ARCHIVES', 'STORAGE',
    'RECOMMENDATION_CONFIRM', 'RECOMMENDATION_PREFERENCES', 'RECOMMENDATION_PROMPT',
    'ANKI_CONNECTION', 'BULK_DETAILS', 'NEW_VERSION')) {
    $start = if ($kind -eq 'CANDIDATES') { '    scGuiOptions := scWantBigBox' } else { '    ; @DESKTOP_' + $kind + '_CONTROLS_BEGIN@' }
    $bodyStart = $source.IndexOf($start)
    $bodyEnd = $source.IndexOf('    ; @DESKTOP_' + $kind + '_CONTROLS_END@', $bodyStart)
    if ($bodyStart -lt 0 -or $bodyEnd -le $bodyStart) { throw "Study $kind constructor markers not found" }
    $generated = $generated.Replace('; @STUDY_' + $kind + '_CONTROLS@', $source.Substring($bodyStart, $bodyEnd - $bodyStart))
}
$studyCapture = $capture.Replace('TestCapture(name, width, height)', 'TestDesktopStudyCapture(testGui, name, width, height)').Replace('global ui', '').Replace('ui.Hwnd', 'testGui.Hwnd')
$generated += "`r`n" + $studyCapture
# Exercise refresh race/error handling with the production body but a synthetic
# catalog query. The real provider/API-key/network helper is never called.
$refreshBody = [regex]::Match($source, '(?ms)^ModelPickerRefresh\(.*?^}').Value
if (-not $refreshBody) { throw 'Model picker refresh function not found' }
$generated += "`r`n" + $refreshBody.Replace('ModelPickerRefresh(', 'TestModelPickerRefresh(').Replace('ModelCatalogQueryWithFeedback(', 'TestModelCatalogQuery(').Replace('CPAdaptiveOwnedMessage(', 'TestModelRefreshNotice(')
foreach ($kind in @('CHOICE', 'CONTEXT')) {
    $start = $source.IndexOf('    ; @SHARED_' + $kind + '_POPUP_CONTROLS_BEGIN@')
    $end = $source.IndexOf('    ; @SHARED_' + $kind + '_POPUP_CONTROLS_END@', $start)
    if ($start -lt 0 -or $end -le $start) { throw "Shared $kind popup markers not found" }
    $generated = $generated.Replace('; @SHARED_' + $kind + '_POPUP_CONTROLS@', $source.Substring($start, $end - $start))
}
if ($StudyOnly) {
    $generated = $generated.Replace('; @STUDY_ONLY@', 'TestDesktopStudyWindows()' + "`r`n" + '    ui.Destroy()' + "`r`n" + '    FileAppend("PASS: " TestAssertions " desktop study assertions.`n", "*")' + "`r`n" + '    ExitApp(0)')
}
if ($ShutdownOnly) {
    $generated = $generated.Replace('; @STUDY_ONLY@', 'DesktopTestExitWithLiveWindow()')
}
if ($DialogsOnly) {
    $generated = $generated.Replace('; @STUDY_ONLY@', 'TestDesktopStudyDialogs()' + "`r`n" + '    ui.Destroy()' + "`r`n" + '    FileAppend("PASS: " TestAssertions " desktop dialog assertions.`n", "*")' + "`r`n" + '    ExitApp(0)')
}
$script = Join-Path $output.FullName 'desktop-generated.ahk'
[IO.File]::WriteAllText($script, $generated, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'desktop.stdout.txt'
$stderr = Join-Path $output.FullName 'desktop.stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $script + '"')) -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
if (!$process.WaitForExit(90000)) {
    $process.Kill()
    throw "Desktop test timed out. Logs: $output"
}
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or $result.Contains('Warning:') -or $result.Contains('FAIL:')) { throw "$result`n$errors`nArtifacts: $output" }
$namingChecks = 0
$namingBounds = Join-Path $output.FullName 'naming-visibility.txt'
if (Test-Path -LiteralPath $namingBounds) {
    # Geometry alone misses a static background painting over the form. Check
    # each control's interior for its light text, excluding native borders.
    foreach ($line in [IO.File]::ReadAllLines($namingBounds)) {
        $parts = $line.Split('|')
        $bitmap = [Drawing.Bitmap]::new((Join-Path $output.FullName $parts[0]))
        try {
            $visibleText = $false
            $left = [int]$parts[2]; $top = [int]$parts[3]
            $right = $left + [int]$parts[4]; $bottom = $top + [int]$parts[5]
            for ($y = $top; $y -lt $bottom -and -not $visibleText; $y++) {
                for ($x = $left; $x -lt $right; $x++) {
                    $pixel = $bitmap.GetPixel($x, $y)
                    if ($pixel.R -gt 140 -and $pixel.G -gt 140 -and $pixel.B -gt 140) {
                        $visibleText = $true
                        break
                    }
                }
            }
            if (-not $visibleText) { throw "Naming control text is covered or missing: $($parts[0]) $($parts[1]). Artifacts: $output" }
            $namingChecks++
        } finally { $bitmap.Dispose() }
    }
    Write-Output "PASS: $namingChecks naming-control visibility checks."
}
if ($StudyOnly -or $ShutdownOnly -or $DialogsOnly) {
    Write-Output $result.Trim()
    Write-Output "Test artifacts: $output"
    return
}
Add-Type -AssemblyName System.Drawing
$scrollPixelChecks = 0
foreach ($bounds in [IO.File]::ReadAllLines((Join-Path $output.FullName 'scroll-bounds.txt'))) {
    $parts = $bounds.Split('|')
    $beforeImage = [Drawing.Bitmap]::new((Join-Path $output.FullName ($parts[0] + '-before.png')))
    $afterImage = [Drawing.Bitmap]::new((Join-Path $output.FullName ($parts[0] + '.png')))
    try {
        $side = [int]$parts[1]; $top = [int]$parts[2]; $bottom = [int]$parts[3]
        # Compare pixels rather than just geometry/paint counts. Controls outside
        # their clipping region must never be drawn over the fixed chrome.
        $regions = @(
            [Drawing.Rectangle]::new(0, 0, $beforeImage.Width, $top),
            [Drawing.Rectangle]::new(0, $top, $side, $bottom - $top),
            [Drawing.Rectangle]::new(0, $bottom, $beforeImage.Width, $beforeImage.Height - $bottom)
        )
        foreach ($region in $regions) {
            $pixelImages = foreach ($snapshot in @($beforeImage, $afterImage)) {
                $cropped = $snapshot.Clone($region, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
                $pixelStream = [IO.MemoryStream]::new()
                try {
                    $cropped.Save($pixelStream, [Drawing.Imaging.ImageFormat]::Png)
                    [Convert]::ToBase64String($pixelStream.ToArray())
                } finally {
                    $pixelStream.Dispose()
                    $cropped.Dispose()
                }
            }
            if ($pixelImages[0] -cne $pixelImages[1]) { throw "Scroll changed pixels outside the content viewport: $($parts[0]) $region. Artifacts: $output" }
            $scrollPixelChecks++
        }
    } finally {
        $beforeImage.Dispose()
        $afterImage.Dispose()
    }
}
Write-Output $result.Trim()
Write-Output "PASS: $scrollPixelChecks fixed-chrome pixel comparisons."
Write-Output "Test artifacts: $output"
# Also cover ExitApp with the main GUI still alive. The layout run explicitly
# destroys it, which otherwise hides global-release / WM_NCDESTROY regressions.
& $PSCommandPath -ShutdownOnly -AutoHotkey $AutoHotkey -OutputDirectory (Join-Path $output.FullName 'shutdown')

param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-bigbox-tests-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$source = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$sourceNormalized = $source.Replace("`r`n", "`n")
$stage5Requirements = [ordered]@{
    'Big Box Study routing' = 'OpenStudyLibraryWindow(true, true)'
    'presentation-aware Library signature' = 'OpenStudyLibraryWindow(slStandalone := false, slBigBoxPresentation := false)'
    'fullscreen Library window style' = '+AlwaysOnTop -Caption +ToolWindow -DPIScale +OwnDialogs'
    'fullscreen Library resize branch' = 'if StudyLibraryBigBoxPresentation(slState) {'
    'desktop-bounds protection' = "StudyLibrarySaveBounds(slState) {`n    global iniPath`n    if StudyLibraryBigBoxPresentation(slState)`n        return"
    'dashboard return action' = 'Back to Dashboard'
    'fullscreen Reader presentation predicate' = 'StudyReaderBigBoxPresentation(srState)'
    'fullscreen Reader layout' = 'StudyReaderResizeBigBox(srState, srGui, srWidth, srHeight)'
    'Reader return action' = 'Back to Library'
    'Reader desktop-bounds protection' = "StudyReaderSaveBounds(srState) {`n    global iniPath`n    if StudyReaderBigBoxPresentation(srState)`n        return"
    'fullscreen Review presentation predicate' = 'StudyCandidatesBigBoxPresentation(scState)'
    'fullscreen Review layout' = 'StudyCandidatesResizeBigBox(scState, scGui, scWidth, scHeight)'
    'visible Review page navigation' = 'StudyCandidatesSwitchBigBoxPage.Bind(scState, -1)'
    'fullscreen recommendation workflow' = 'StudyCandidatesRecommendationBigBoxShell('
    'Reader returns to fullscreen Review' = 'CPStudyReaderState["returnToCandidates"] := scState'
    'shared fullscreen Study table mode' = 'StudyBigBoxTableModeCreate(scGui, scColors)'
    'Library table-mode entry' = 'slTableModeControls := StudyBigBoxTableModeCreate('
    'Review table-mode entry' = 'scTableModeControls := StudyBigBoxTableModeCreate(scGui, scColors)'
    'table-mode controller column browsing' = 'return StudyBigBoxTableModeScroll('
    'fullscreen Library column editor' = 'return StudyLibraryOpenBigBoxColumns(slState)'
    'column editor working-copy save' = 'StudyLibraryBigBoxColumnsPersist(slEditor)'
    'responsive column-detail transition' = 'slEditor["gui"].GetClientPos(,, &slClientW, &slClientH)'
    'Review row action sheet' = 'StudyCandidatesShowSelectedActions(scState, *)'
    'controller-first Add to Anki action' = 'if scPreferAddToAnki {'
    'non-blocking controller row-action popup' = 'StudyCandidatesShowSelectedActionsDeferred.Bind('
    'stable popup row base colors' = 'CPThemedChoicePopupResetRows(cpPopupState)'
    'seamless fullscreen Study handoff' = 'Keep the dashboard visible and foreground while the Library performs'
    'controller release transition guard' = 'CPControllerSurfaceTransitionBlocks(cpNavState, targetHwnd)'
}
foreach ($requirement in $stage5Requirements.GetEnumerator()) {
    if (!$sourceNormalized.Contains($requirement.Value)) {
        throw "Stage 5 source requirement missing: $($requirement.Key)"
    }
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'audio_input_fixture.ahk') -Destination $output.FullName -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'audio_runtime_fixture.ahk') -Destination $output.FullName -Force
Add-Type -AssemblyName System.Drawing
$artPaths = foreach ($spec in @(@('box', 180, 300), @('logo', 400, 90))) {
    $artPath = Join-Path $output.FullName ($spec[0] + '.png')
    $bitmap = [Drawing.Bitmap]::new($spec[1], $spec[2])
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear([Drawing.Color]::SteelBlue)
        $bitmap.Save($artPath, [Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
    $artPath
}

# Extract the production dashboard and native navigation functions verbatim.
# Only external application services are stubbed, so no personal settings,
# overlays, controller bindings, API requests or library databases are touched.
$dashboard = [regex]::Match($source, '(?ms)^CPBigBoxDashboardAlive\(\).*?(?=^CPApplyPresentationModeTransition\()').Value
if (!$dashboard) { throw 'Dashboard source block was not found.' }
$globals = [regex]::Matches($source, '(?m)^global CPBigBox[^\r\n]*') | ForEach-Object { $_.Value }
$functions = @('CPPalette', 'CPSetWindowCloaked', 'StudyLibraryImageDimensions',
    'CPShowDialogFocusCues', 'CPControllerNavigationState', 'CPControllerResetNavigation',
    'CPControllerBeginSurfaceTransition', 'CPControllerFinishSurfaceTransition',
    'CPControllerSurfaceTransitionBlocks',
    'CPControllerDispatchNavigation', 'CPControllerHandleNavigation',
    'AutoPersist', 'UpdateVars', 'SaveAll', 'ApplyShotSettings', 'ExplainPromptChanged',
    'ToggleModelControls', 'ToggleExplanationControls', 'ToggleAudioControls',
    'CPExplanationPreference', 'CPSetExplanationPreference', 'CPExplanationPreferenceChanged',
    'StudyLibrarySaveToggleChanged', 'CPScreenshotPreference', 'CPSetScreenshotPreference',
    'CPScreenshotPreferenceChanged', 'CPSetAudioDeviceSelection', 'SpeakerChanged',
    'RefreshSpeakerList', 'SetAudioTestStatus', 'TestAudioInput', 'AudioInputJobBusy',
    'AudioInputJobControls', 'AudioInputJobStart', 'AudioInputJobDispose', 'AudioInputJobCancel',
    'AudioInputJobPoll', 'AudioInputApplyTestResult', 'AudioInputApplyDeviceResult',
    'CPSetCaptureMaxKB', 'CPMaxPngAdjustSyncValue', 'StartAudioCore', 'CPSuspendBigBoxForCapture',
    'CPColorHexToHSV', 'CPColorHSVToHex', 'StartOverlayAdjustmentCore', 'CPFinishOverlayAdjustment',
    'GetWindowDPI', 'CPColorRef', 'CPColorGradientWriteVertex', 'CPColorGradientFillRect',
    'CPDrawControllerColorGradient', 'CPControllerColorGradientCustomDraw',
    'CPRegisterControllerColorGradients', 'CPUnregisterControllerColorGradients', 'CPControllerColorDeferredGradientRedraw',
    'HotkeyPretty', 'NormalizeHotkey', 'CPControllerTokenDisplay',
    'GameProfileSafeName', 'GameProfilePath', 'ListGameProfiles',
    'GlossaryProfileDir', 'GlossaryJP2ENPath', 'GlossaryEN2ENPath', 'GlossaryPath',
    'GlossaryKindLabel', 'GlossaryHeader', 'ListGlossaryProfiles', 'GlossaryReadDocument',
    'GlossaryCloneEntries', 'GlossaryNormalizeSource', 'GlossaryDuplicateSummary',
    'GlossaryValidateEntries', 'GlossaryBuildText', 'GlossaryEnsureFile', 'ArrHas',
    'SaveTextAtomic', 'ArrayIndexOf', 'ArrIndexOf',
    'ModelListNaturalCompare', 'ModelListSort', 'SetComboItems',
    'SetComboToExistingItem', 'RefreshModelCombos', 'ModelAlreadyAdded',
    'ModelCatalogParseOutput', 'CPNormalizeApiSecret', 'CPDotEnvValue',
    'UpdatePathsDirtyState',
    'ListPromptProfiles', 'RefreshPromptProfilesList',
    'ListExplainPromptProfiles', 'RefreshExplainPromptProfilesList',
    'AboutVersionInfo')
$extra = foreach ($name in $functions) {
    $match = [regex]::Match($source, '(?ms)^' + $name + '\([^\n]*\{.*?^}')
    if (!$match.Success) { throw "Function not found: $name" }
    $match.Value
}
$harness = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'bigbox_navigation_harness.ahk'))
$desktopTabs = [regex]::Match($source, '(?m)^tabNames := (\[[^\r\n]+\])')
if (!$desktopTabs.Success) { throw 'Desktop tab definitions were not found.' }
$generated = $harness.Replace('; @BIGBOX_GLOBALS@', ($globals -join "`r`n"))
$generated = $generated.Replace('; @DESKTOP_TABS@', ('global TestDesktopTabNames := ' + $desktopTabs.Groups[1].Value))
$generated = $generated.Replace('; @BIGBOX_SOURCE@', ($dashboard + "`r`n" + ($extra -join "`r`n")))
$generatedPath = Join-Path $output.FullName 'bigbox-navigation-generated.ahk'
[IO.File]::WriteAllText($generatedPath, $generated, [Text.UTF8Encoding]::new($true))

$studyFunctions = @(
    'CPHwndIsCombo', 'CPComboDropped', 'CPShowCombo', 'CPHwndIsFocusable',
    'CPApplyOwnedDialogTheme',
    'StudyBigBoxFocusFrameKeys', 'StudyBigBoxFocusFrameHide',
    'StudyBigBoxFocusFrameUpdate', 'StudyBigBoxFocusFrameFocused',
    'StudyBigBoxFocusFrameWatch', 'StudyBigBoxFocusFrameStart',
    'StudyBigBoxFocusFrameStop', 'StudyBigBoxFocusFrameEnsure',
    'CPGetHwndRect', 'CPThemedChoicePopupFinish',
    'CPThemedChoicePopupResetRows',
    'CPThemedChoicePopupPaintFocus', 'CPThemedChoicePopupFocus',
    'StudyControllerSurfaces',
    'StudyControllerSurfaceForWindow', 'StudyControllerSurfaceIsRoot',
    'StudyControllerEnumFocusableProc', 'StudyControllerFocusableHwnds',
    'StudyControllerFocusedHwnd', 'StudyControllerSetFocus',
    'StudyControllerControlClass', 'StudyControllerIsReadOnlyEdit',
    'StudyControllerClearReadOnlyEditSelection',
    'StudyControllerIsReadOnlyMultilineEdit',
    'StudyControllerScrollReadOnlyEdit',
    'StudyControllerSendKey', 'StudyControllerSendControlKey',
    'StudyControllerComboPreviewActive',
    'StudyControllerComboNeedsConfirmation',
    'StudyControllerBeginComboSelection',
    'StudyControllerCommitComboSelection',
    'StudyControllerCancelComboSelection',
    'StudyControllerListViewCanMove',
    'StudyControllerMoveFocus', 'StudyControllerMove',
    'StudyControllerActivate', 'StudyControllerCancel',
    'StudyControllerSwitchPage', 'StudyControllerHandleChoicePopup',
    'StudyControllerDispatchNavigation',
    'CPControllerNavigationTarget',
    'StudyCandidatesDestroyGui',
    'StudyCandidatesRestoreLibraryFocus',
    'StudyCandidatesInitialRefresh',
    'StudyCandidatesInitialRefreshComplete',
    'StudyCandidatesBigBoxPresentation',
    'StudyCandidatesTabChanged',
    'StudyCandidatesBigBoxPageData',
    'StudyCandidatesBigBoxColumnWidths',
    'StudyCandidatesApplyBigBoxTheme',
    'StudyControllerContextPopupPoint',
    'StudyCandidatesShowSelectedActions',
    'StudyCandidatesShowSelectedActionsDeferred',
    'StudyBigBoxTableModeCreate',
    'StudyBigBoxTableModeActive',
    'StudyBigBoxTableModeLists',
    'StudyBigBoxTableModeCurrentList',
    'StudyBigBoxTableModeText',
    'StudyBigBoxTableModeUpdate',
    'StudyBigBoxTableModeApplyTheme',
    'StudyBigBoxTableModeScroll',
    'StudyBigBoxTableModeRestoreScroll',
    'StudyBigBoxTableModeSet',
    'StudyLibraryGroupFocused', 'StudyLibraryLibraryChanged',
    'StudyLibraryBigBoxColumnWidths',
    'StudyLibraryColumnDefaultWidths',
    'StudyLibraryBigBoxColumnsShowMode',
    'StudyLibraryBigBoxColumnsMove',
    'StudyLibraryBigBoxColumnsWidth',
    'StudyLibraryBigBoxColumnsToggle',
    'StudyLibraryBigBoxColumnsResetWidth',
    'StudyReaderBigBoxPresentation',
    'StudyAnkiDialogAlive', 'StudyAnkiCloseDialog', 'StudyAnkiDiscover'
)
$studySource = foreach ($name in $studyFunctions) {
    $match = [regex]::Match($source, '(?ms)^' + $name + '\([^\n]*\{.*?^}')
    if (!$match.Success) { throw "Study controller function not found: $name" }
    $match.Value
}
$studyHarness = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'study_controller_harness.ahk'))
$studyGenerated = $studyHarness.Replace('; @STUDY_SOURCE@', ($studySource -join "`r`n"))
$studyGeneratedPath = Join-Path $output.FullName 'study-controller-generated.ahk'
[IO.File]::WriteAllText($studyGeneratedPath, $studyGenerated, [Text.UTF8Encoding]::new($true))

function Invoke-AhkTest([string]$Script, [string]$Name, [string[]]$ExtraArguments = @()) {
    $stdout = Join-Path $output.FullName ($Name + '.stdout.txt')
    $stderr = Join-Path $output.FullName ($Name + '.stderr.txt')
    $arguments = @('/ErrorStdOut', ('"' + $Script + '"')) + $ExtraArguments
    $process = Start-Process -FilePath $AutoHotkey -ArgumentList $arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    if (!$process.WaitForExit(30000)) {
        $process.Kill()
        throw "$Name timed out; stopped only its test process."
    }
    $process.WaitForExit()
    $out = [IO.File]::ReadAllText($stdout)
    $err = [IO.File]::ReadAllText($stderr)
    if ($process.ExitCode -ne 0 -or $err -or $out.Contains('Warning:')) {
        $tail = (($out.Trim() -split '\r?\n') | Select-Object -Last 16) -join "`n"
        $warnings = [regex]::Matches($out, '(?m)^.*Warning:.*(?:\r?\n.*){0,2}') | ForEach-Object { $_.Value }
        throw "$Name failed (exit $($process.ExitCode))`n$($warnings -join "`n")`n$tail`n$err`nFull log: $stdout"
    }
    if ($out) { Write-Output (($out.Trim() -split '\r?\n')[-1]) }
}
Invoke-AhkTest (Join-Path $PSScriptRoot 'bigbox_syntax_check.ahk') 'syntax'
Invoke-AhkTest (Join-Path $PSScriptRoot 'overlay_syntax_check.ahk') 'overlay-syntax'
Invoke-AhkTest $generatedPath 'navigation' @(('"' + $artPaths[0] + '"'), ('"' + $artPaths[1] + '"'))
Invoke-AhkTest $studyGeneratedPath 'study-controller'
Write-Output "Test artifacts: $($output.FullName)"

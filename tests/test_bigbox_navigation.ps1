param(
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-bigbox-tests-' + [Guid]::NewGuid().ToString('N'))),
    [switch]$StudyOnly
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$source = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$sourceNormalized = $source.Replace("`r`n", "`n")
$profileSaveSource = [regex]::Match($sourceNormalized, '(?ms)^GameProfileSave\([^\n]*\{.*?^}').Value
$profileApplySource = [regex]::Match($sourceNormalized, '(?ms)^GameProfileApply\([^\n]*\{.*?^}').Value
if (!$profileSaveSource.Contains('GameProfileSaveStartup(path)') -or
    !$profileApplySource.Contains('GameProfileApplyStartup(path)') -or
    !$sourceNormalized.Contains("if (CP_START_PROFILE != `"`")`n    CPApplyExternalProfile(CP_START_PROFILE)") -or
    $sourceNormalized.LastIndexOf('CPApplyExternalProfile(CP_START_PROFILE)') -gt
        $sourceNormalized.IndexOf('startupOverlays := CPStartupOverlayPlan(CP_STUDY_START_MODE, CP_START_TRANSLATOR)')) {
    throw 'Startup overlay choices must round-trip through Profiles and be read after the launch Profile is applied.'
}
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
    'Review-owned Add to Anki workflow' = 'scAnkiState := StudyCandidatesAnkiState(scState, scCandidate)'
    'direct controller column browsing' = 'StudyControllerScrollTable(studyFocused, direction,'
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
$candidateAddSource = [regex]::Match(
    $sourceNormalized, '(?ms)^StudyCandidatesAddSelected\([^\n]*\{.*?^}'
).Value
if (!$candidateAddSource -or $candidateAddSource.Contains('OpenStudyReader(') -or
    !$candidateAddSource.Contains('StudyCandidatesAnkiState(scState, scCandidate)')) {
    throw 'Review for Anki must own Add to Anki directly without opening the Study Reader.'
}
if ($sourceNormalized -match 'StudyBigBoxTableMode|Table mode\.\.\.') {
    throw 'Removed fullscreen Table mode controls or routing were reintroduced.'
}
$libraryOpenSource = [regex]::Match($sourceNormalized, '(?ms)^OpenStudyLibraryWindow\([^\n]*\{.*?^}').Value
if ($libraryOpenSource -match '(?:slList|CPStudyLibraryState\["list"\])\.Focus\(\)' -or
    [regex]::Matches($libraryOpenSource, 'StudyLibraryFocusOnOpen\(').Count -ne 4) {
    throw 'Library startup/reopen must use presentation-aware initial focus before and after activation.'
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
$assetOutput = New-Item -ItemType Directory -Path (Join-Path $output.FullName 'assets') -Force
Copy-Item -LiteralPath (Join-Path $repo 'assets/bigbox-logo.png') -Destination $assetOutput.FullName -Force
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
    'CPSetCaptureMaxKB', 'CPMaxPngAdjustSyncValue', 'AudioSessionPid', 'AudioSessionClear',
    'StartAudioCore', 'CPSuspendBigBoxForCapture',
    'CPColorHexToHSV', 'CPColorHSVToHex', 'StartOverlayAdjustmentCore', 'CPFinishOverlayAdjustment',
    'GetWindowDPI', 'CPColorRef', 'CPColorGradientWriteVertex', 'CPColorGradientFillRect',
    'CPDrawControllerColorGradient', 'CPControllerColorGradientCustomDraw',
    'CPRegisterControllerColorGradients', 'CPUnregisterControllerColorGradients', 'CPControllerColorDeferredGradientRedraw',
    'HotkeyPretty', 'NormalizeHotkey', 'CPControllerTokenDisplay',
    'GameProfileSafeName', 'GameProfilePath', 'ListGameProfiles', 'GameProfileReadInt',
    'CPStartupOverlayOptions', 'CPStartupOverlayPlan', 'CPStartupOverlayIndex',
    'CPStartupOverlaysSync', 'CPSetStartupOverlays', 'CPStartupOverlaysChanged',
    'GameProfileSaveStartup', 'GameProfileApplyStartup',
    'GlossaryProfileDir', 'GlossaryJP2ENPath', 'GlossaryEN2ENPath', 'GlossaryPath',
    'GlossaryKindLabel', 'GlossaryHeader', 'ListGlossaryProfiles', 'GlossaryReadDocument',
    'GlossaryCloneEntries', 'GlossaryNormalizeSource', 'GlossaryDuplicateSummary',
    'GlossaryValidateEntries', 'GlossaryBuildText', 'GlossaryEnsureFile', 'ArrHas',
    'SaveTextAtomic', 'ArrayIndexOf', 'ArrIndexOf',
    'ModelListNaturalCompare', 'ModelListSort', 'SetComboItems',
    'SetComboToExistingItem', 'RefreshModelCombos', 'ModelAlreadyAdded',
    'ModelCatalogParseOutput', 'CPNormalizeApiSecret', 'CPDotEnvValue',
    'UpdatePathsDirtyState',
    'CPNativePickerOwner', 'CPNativePicker', 'CPNativeFileSelect', 'CPNativeDirSelect',
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
$audioLanguages = [regex]::Match($source, '(?ms)^audioTargetLangs := (\[.*?^\])')
if (!$audioLanguages.Success) { throw 'Audio output language definitions were not found.' }
$generated = $generated.Replace('; @AUDIO_LANGUAGES@', ('global TestAudioTargetLangs := ' + $audioLanguages.Groups[1].Value))
$generated = $generated.Replace('; @BIGBOX_SOURCE@', ($dashboard + "`r`n" + ($extra -join "`r`n")))
$generatedPath = Join-Path $output.FullName 'bigbox-navigation-generated.ahk'
[IO.File]::WriteAllText($generatedPath, $generated, [Text.UTF8Encoding]::new($true))

$studyFunctions = @(
    'StudyDesktopRegistry', 'StudyDesktopContext',
    'StudyLibraryFocusOnOpen', 'StudyLibraryImageCounterText', 'StudyLibraryApplyBigBoxFonts',
    'StudyLibrarySafeName', 'StudyLibraryDirectoryForName', 'StudyLibraryListNames',
    'StudyLibraryRefreshLibrarySelector', 'StudyLibraryFormatBytes',
    'StudyLibraryCreateNew', 'StudyLibraryOpenNew', 'StudyLibraryManagerSelectedName',
    'StudyLibraryManagerUpdateActions', 'StudyLibraryRefreshManager', 'StudyLibraryManagerSwitch',
    'StudyLibraryManagerNew', 'StudyLibraryManagerOpenFolder', 'StudyLibraryManagerRenameApply',
    'StudyLibraryManagerRename', 'StudyLibraryManagerArchive', 'StudyLibraryArchiveEntries',
    'StudyLibraryArchiveSelectedEntry', 'StudyLibraryArchiveUpdateActions', 'StudyLibraryRefreshArchives',
    'StudyLibraryArchiveOpenFolder', 'StudyLibraryArchiveRestoreApply', 'StudyLibraryArchiveRestore',
    'StudyLibraryOpenArchives', 'StudyLibraryOpenManager',
    'StudyLibraryManagementForm', 'StudyLibraryManagementChildAlive', 'StudyLibraryManagementClose',
    'StudyLibraryQueueManagementAction', 'StudyLibraryRunManagementAction', 'StudyLibraryManagementMessage',
    'StudyLibraryManagementColumns', 'StudyLibraryManagementListActions',
    'StudyLibraryOpenBigBoxManagement', 'StudyLibraryOpenManagementName',
    'CPControllerBeginSurfaceTransition', 'CPControllerFinishSurfaceTransition', 'CPControllerSurfaceTransitionBlocks',
    'CPBigBoxDashboardDpiScale',
    'StudyReaderVocabularyEntries', 'StudyReaderVocabularyPickerAlive',
    'StudyReaderVocabularyPickerChanged', 'StudyReaderVocabularyPickerClose',
    'StudyReaderVocabularyPickerChoose', 'StudyReaderVocabularyPickerOpenReview',
    'StudyReaderVocabularyPickerReturn', 'StudyCandidatesAnkiState',
    'StudyReaderOpenVocabularyPicker', 'StudyReaderShowAnkiMenu',
    'StudyReaderCloseAnkiAddDialog', 'StudyReaderQueueAnkiAction', 'StudyReaderRunAnkiAction',
    'StudyReaderAddReviewedAnkiNote', 'StudyReaderAnkiMessage', 'StudyLibraryOwnedMessage',
    'StudyLibraryOwnedMessageClose', 'StudyLibraryOwnedMessageBigBoxResize', 'CPMeasureWrappedTextHeight',
    'StudyLibraryActiveProfileName', 'StudyLibraryChapterSettingsPath', 'StudyLibraryChapterSettingsKey',
    'StudyLibraryCurrentChapter', 'StudyLibraryWriteCurrentChapter', 'StudyLibraryChapterHistorySection',
    'StudyLibraryReadChapterHistory', 'StudyLibraryWriteChapterHistory', 'StudyLibraryRememberChapter',
    'StudyLibrarySetChapterComboChoices', 'StudyLibraryRefreshCurrentChapterDisplay',
    'StudyLibraryCloseCurrentChapterDialog', 'StudyLibraryChapterMessage', 'StudyLibrarySaveCurrentChapter',
    'StudyLibraryQueueChapterAction', 'StudyLibraryRunChapterAction',
    'StudyLibraryRemoveSavedChapter', 'StudyLibraryClearChapterHistory', 'StudyLibraryOpenCurrentChapter',
    'GameProfileSafeName', 'IniWriteRetry',
    'StudyReaderAnkiScreenshotToggle', 'StudyReaderAnkiScreenshotToggleText',
    'StudyReaderAnkiScreenshotPreferenceChanged', 'StudyReaderAnkiPreviewBigBoxResize',
    'StudyReaderHasGeneratedExample', 'StudyReaderCurrentScreenshot', 'StudyReaderRemoveJapaneseReadings',
    'StudyAnkiChooseText', 'StudyAnkiDeckScope', 'StudyAnkiTextInList', 'StudyLibraryImageDimensions',
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
    'StudyControllerIsMultilineEdit', 'StudyControllerScrollMultilineEdit',
    'StudyControllerSendKey', 'StudyControllerSendControlKey',
    'StudyControllerComboPreviewActive',
    'StudyControllerComboNeedsConfirmation',
    'StudyControllerBeginComboSelection',
    'StudyControllerCommitComboSelection',
    'StudyControllerCancelComboSelection',
    'StudyControllerListViewCanMove', 'StudyControllerScrollTable',
    'StudyControllerMoveFocus', 'StudyControllerMove',
    'StudyControllerActivate', 'StudyControllerCancel',
    'StudyControllerSwitchPage', 'StudyControllerHandleChoicePopup',
    'StudyControllerDispatchNavigation',
    'StudyLibraryDatePickerRegistry', 'StudyLibraryDatePickerShutdown', 'StudyLibraryDatePickerStep',
    'StudyLibraryDatePickerPartText', 'StudyLibraryDatePickerRefresh',
    'StudyLibraryDatePickerAdjust', 'StudyLibraryDatePickerClose',
    'StudyLibraryDatePickerNavigate', 'StudyLibraryDatePickerKeyDown',
    'StudyLibraryDatePickerDestroyed', 'StudyLibraryOpenDatePicker',
    'StudyLibraryBigBoxFormState', 'StudyLibraryBigBoxFormAdd',
    'StudyLibraryBigBoxFormShow', 'StudyCandidatesRecommendationBigBoxShell',
    'StudyCandidatesRecommendationBigBoxShow',
    'StudyCandidatesRecommendationBigBoxApplyFonts',
    'StudyCandidatesRecommendationBigBoxApplyTheme',
    'StudyCandidatesRecommendationDefaults',
    'StudyCandidatesRecommendationInstructionsPath', 'SaveTextAtomic',
    'StudyCandidatesRecommendationLevelIndex', 'StudyCandidatesRecommendationStyleIndex',
    'StudyCandidatesRecommendationSyncBasicSettings',
    'StudyCandidatesRecommendationDialogClose', 'StudyCandidatesRecommendationAdvancedDraft',
    'StudyCandidatesApplyRecommendationDialogTheme',
    'StudyCandidatesRecommendationToggleText', 'StudyCandidatesRecommendationSetupBigBoxToggles',
    'StudyCandidatesRecommendationAdvancedRestore', 'StudyCandidatesRecommendationAdvancedClose',
    'StudyCandidatesRecommendationCustomize', 'StudyCandidatesRecommendationConfirm',
    'StudyCandidatesRecommendationPromptPreview', 'StudyCandidatesRecommendationPreviewClose',
    'StudyCandidatesRecommendationDefaultInstructions',
    'StudyCandidatesRecommendationNormalizeInstructions',
    'StudyCandidatesRecommendationPreviewMode', 'StudyCandidatesRecommendationPreviewToggle',
    'StudyCandidatesRecommendationPreviewRestore',
    'StudyCandidatesRecommendationShowPrompt',
    'StudyLibraryCompactDateTime', 'StudyLibraryBigBoxPresentation',
    'StudyLibraryClearButtonHover',
    'StudyLibraryOpenFilters', 'StudyLibraryApplyFilters',
    'StudyLibraryDateModeIndex', 'StudyLibraryDateModeFromIndex',
    'StudyLibraryChoiceIndex', 'StudyLibraryDateControlsChanged',
    'StudyLibraryClearFilters', 'StudyLibraryClearFiltersAndClose',
    'StudyLibraryCloseDialog',
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
    'StudyLibraryGroupFocused', 'StudyLibraryLibraryChanged',
    'StudyLibraryAddVersionDisplay', 'StudyLibrarySyncVersionNavigation',
    'StudyLibraryBigBoxColumnWidths',
    'StudyLibraryColumnFilterActive', 'StudyLibraryColumnTitle',
    'StudyLibraryApplyColumns', 'StudyLibraryApplyBigBoxColumns',
    'StudyLibraryEnsureInternalColumnHidden', 'StudyLibraryApplyHeaderIndicators',
    'CPStudyHeaderText',
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
    $match = [regex]::Match($source, '(?ms)^' + $name + '\([^{]*\{.*?^}')
    if (!$match.Success) { throw "Study controller function not found: $name" }
    $match.Value
}
$formResize = [regex]::Match($source, '(?ms)^StudyCandidatesRecommendationBigBoxResize\([^{]*\{.*?^}').Value
if (!$formResize) { throw 'Study form layout function not found.' }
$studySource += $formResize.Replace('StudyCandidatesRecommendationBigBoxResize(', 'TestProductionStudyFormResize(')
$libraryResize = [regex]::Match($source, '(?ms)^StudyLibraryResizeBigBox\([^{]*\{.*?^}').Value
if (!$libraryResize) { throw 'Fullscreen Library layout function not found.' }
$studySource += $libraryResize.Replace('StudyLibraryResizeBigBox(', 'TestProductionLibraryResize(')
$ankiPreview = [regex]::Match($source, '(?ms)^StudyReaderOpenReviewedAnkiDialog\([^{]*\{.*?^}').Value
if (!$ankiPreview) { throw 'Anki preview function not found.' }
$studySource += $ankiPreview.Replace('StudyReaderOpenReviewedAnkiDialog(', 'TestProductionAnkiPreview(')
# Real persistence against a temporary control.ini, separate from modal-test stubs.
foreach ($name in @('StudyCandidatesRecommendationLoadSettings', 'StudyCandidatesRecommendationSaveSettings',
    'StudyLibraryHexEncode', 'StudyLibraryHexDecode')) {
    $body = [regex]::Match($source, '(?ms)^' + $name + '\([^\n]*\{.*?^}').Value
    if (!$body) { throw "Persistence test function not found: $name" }
    $body = $body.Replace($name + '(', 'TestProduction' + $name + '(')
    if ($name.StartsWith('StudyCandidates')) {
        $body = $body.Replace('StudyLibraryHexEncode(', 'TestProductionStudyLibraryHexEncode(')
        $body = $body.Replace('StudyLibraryHexDecode(', 'TestProductionStudyLibraryHexDecode(')
    }
    $studySource += $body
}
$studyHarness = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'study_controller_harness.ahk'))
$studyGenerated = $studyHarness.Replace('; @STUDY_SOURCE@', ($studySource -join "`r`n"))
# Render only the synthetic test window, never the user's desktop.
$captureFunction = [regex]::Match($harness, '(?ms)^TestCapture\(.*?^}').Value
if (!$captureFunction) { throw 'Synthetic window capture helper not found.' }
$captureFunction = $captureFunction.Replace('TestCapture(name, width, height)', 'TestStudyCapture(testGui, name, width, height)')
$captureFunction = $captureFunction.Replace('global CPBigBoxGui', '').Replace('CPBigBoxGui', 'testGui')
$studyGenerated += "`r`n" + $captureFunction
foreach ($measureHelper in @('TestTextHeight', 'TestFontHeight')) {
    $studyGenerated += "`r`n" + [regex]::Match($harness, '(?ms)^' + $measureHelper + '\(.*?^}').Value
}
# Use the real tab-container setup so focus-stop regressions cannot be hidden
# by a fixture that independently removes the native tab stop.
$candidateTabs = [regex]::Match($source, '(?ms)^    scTabs := scGui\.Add\(.*?(?=^    scTabs\.SetFont)').Value
if (!$candidateTabs) { throw 'Review tab-container setup was not found.' }
$candidateTabs = $candidateTabs.Replace('scTabs', 'candidateTabs').Replace('scGui', 'candidateGui').Replace('scWantBigBox', 'candidateWantBigBox')
$studyGenerated = $studyGenerated.Replace('; @CANDIDATE_TAB_SETUP@', $candidateTabs)
$studyGeneratedPath = Join-Path $output.FullName 'study-controller-generated.ahk'
[IO.File]::WriteAllText($studyGeneratedPath, $studyGenerated, [Text.UTF8Encoding]::new($true))

function Invoke-AhkTest([string]$Script, [string]$Name, [string[]]$ExtraArguments = @()) {
    $stdout = Join-Path $output.FullName ($Name + '.stdout.txt')
    $stderr = Join-Path $output.FullName ($Name + '.stderr.txt')
    $arguments = @('/ErrorStdOut', ('"' + $Script + '"')) + $ExtraArguments
    $process = Start-Process -FilePath $AutoHotkey -ArgumentList $arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    $timeoutMs = if ($Name -eq 'study-controller') { 120000 }
        elseif ($Name -eq 'navigation') { 60000 }
        else { 30000 }
    if (!$process.WaitForExit($timeoutMs)) {
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
if (!$StudyOnly) {
    Invoke-AhkTest $generatedPath 'navigation' @(('"' + $artPaths[0] + '"'), ('"' + $artPaths[1] + '"'))
}
Invoke-AhkTest $studyGeneratedPath 'study-controller'
Write-Output "Test artifacts: $($output.FullName)"

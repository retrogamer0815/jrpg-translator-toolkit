param(
    [switch]$StudyOnly,
    [switch]$DialogsOnly,
    [switch]$ShutdownOnly,
    [switch]$ResizeOnly,
    [switch]$AudioFeedbackOnly,
    [switch]$ProfileFeedbackOnly,
    [switch]$ChapterOnly,
    [switch]$PromptUnsavedOnly,
    [switch]$GenerationActivityOnly,
    [string]$AutoHotkey = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('jrpg-desktop-tests-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
if (@($StudyOnly, $DialogsOnly, $ShutdownOnly, $ResizeOnly, $AudioFeedbackOnly, $ProfileFeedbackOnly, $ChapterOnly, $PromptUnsavedOnly, $GenerationActivityOnly).Where({ [bool]$_ }).Count -gt 1) {
    throw 'Choose only one focused test mode.'
}
$repo = Split-Path -Parent $PSScriptRoot
$source = [IO.File]::ReadAllText((Join-Path $repo 'JRPG Translator.ahk'))
$pickerRequirements = [ordered]@{
    'shared picker ownership' = 'GuiFromHwnd(ownerHwnd).Opt("+OwnDialogs")'
    'fullscreen modal suspension' = 'CPBigBoxModalDepth += 1'
    'desktop capture route' = 'return OpenModernCapturePicker()'
    'modern capture choices' = '["Capture region", "Capture window", "Cancel"]'
    'Study export owner' = 'slState["gui"].Hwnd, "S16"'
}
foreach ($requirement in $pickerRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Picker-polish source requirement missing: $($requirement.Key)"
    }
}
if ([regex]::Matches($source, '\bFileSelect\(').Count -ne 1 -or
    [regex]::Matches($source, '\bDirSelect\(').Count -ne 1 -or
    [regex]::Matches($source, '\bCPNativeFileSelect\(').Count -ne 2 -or
    [regex]::Matches($source, '\bCPNativeDirSelect\(').Count -ne 2) {
    throw 'Every file/folder picker entry point must use the shared owned wrapper.'
}
$modernActionMenuFunctions = @(
    'CPDesktopAudioModelMenu',
    'CPDesktopExplanationModelMenu',
    'CPDesktopExplanationPromptMenu',
    'CPDesktopModelMenu',
    'CPDesktopPromptMenu'
)
foreach ($functionName in $modernActionMenuFunctions) {
    $body = [regex]::Match($source, '(?ms)^' + [regex]::Escape($functionName) + '\([^\r\n]*\)\s*\{.*?^\}').Value
    if (!$body -or !$body.Contains('CPDesktopActionMenu(') -or $body.Contains('Menu()')) {
        throw "Modern desktop action must use the themed popup: $functionName"
    }
}
if ([regex]::Matches($source, '\bMenu\(\)').Count -ne 0) {
    throw 'Desktop popup actions must not fall back to unthemed native menus.'
}
$promptMenuRequirements = [ordered]@{
    'Game Text prompt edit is inside Manage' = '[OpenPromptEditor, NewPromptProfile, DeletePromptProfile]'
    'Explanation prompt edit is inside Manage' = '[OpenExplainPromptEditor_Multi, NewExplainPromptProfile, DeleteExplainPromptProfile]'
    'prompt manager labels its contextual edit action' = '["Edit selected prompt…", "New prompt…", "Delete selected prompt…"]'
}
foreach ($requirement in $promptMenuRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Prompt-management source requirement missing: $($requirement.Key)"
    }
}
$modernOnlyDesktopRequirements = [ordered]@{
    'legacy preference migration' = 'IniWrite(1, iniPath, "cfg_control", "modernLayout")'
    'modern-only active state' = 'return IsSet(CPDesktop) && CPDesktop.Get("ready", false)'
    'live modern resize route' = 'return CPDesktopLiveResize()'
    'settled modern resize route' = 'return CPDesktopRelayout()'
    'fullscreen return requests normal desktop bounds' = 'CPShowControlPanelReady(true, true)'
    'fullscreen return uses preferred desktop viewport' = 'targetW := Max(clientW, CPPreferredViewportW)'
    'fullscreen return relayouts before reveal' = 'ResizeUI(ui, 0, restoredClientW, restoredClientH)'
}
foreach ($requirement in $modernOnlyDesktopRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Modern-only desktop source requirement missing: $($requirement.Key)"
    }
}
if ($source.Contains('CPDesktopToggleLayout') -or
    $source.Contains('"Use classic layout"') -or
    $source.Contains('"Use modern layout"') -or
    $source.Contains('CPDesktop["modern"]') -or
    [regex]::Matches($source, '\bCPDesktopClassicResize\(').Count -ne 0) {
    throw 'Selectable classic desktop controls, state, menus, and resize code must stay retired.'
}
$modelSourceRequirements = [ordered]@{
    'online source opens directly' = 'c["online"].OnEvent("Click", CPModelDialogChooseSource.Bind(s, "online"))'
    'manual source opens directly' = 'c["manual"].OnEvent("Click", CPModelDialogChooseSource.Bind(s, "manual"))'
}
foreach ($requirement in $modelSourceRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Model-source source requirement missing: $($requirement.Key)"
    }
}
$pageRegistryRequirements = [ordered]@{
    'nine-page registry constructor' = 'CPDesktopCreatePageRegistry()'
    'shared page lookup' = 'CPDesktopPage(page := 0)'
    'semantic control registration' = 'descriptor["controls"][key] := ctrl'
    'shared visibility helper' = 'CPDesktopSetControlGroupVisible(controls, visible)'
    'registry layout dispatch' = 'currentPage["layout"].Call(extentW)'
    'registry navigation identity' = 'active := descriptor["navKey"]'
}
foreach ($requirement in $pageRegistryRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Desktop page-registry source requirement missing: $($requirement.Key)"
    }
}
$profileHeaderRequirements = [ordered]@{
    'header profile dropdown' = 'chrome["profile"] := ui.AddDropDownList'
    'header profile selection wiring' = 'chrome["profile"].OnEvent("Change", CPDesktopProfileSelectionChanged)'
    'header profile list synchronization' = 'CPDesktopRefreshProfileSelector(force := false)'
    'header Manage profiles action' = 'if choice = "Manage profiles…"'
    'profile switch dirty-state comparison' = 'GameProfileHasUnsavedChanges(name)'
    'profile switch save choice' = '"Save and switch", "Switch without saving", "Cancel"'
}
foreach ($requirement in $profileHeaderRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Desktop profile-selector source requirement missing: $($requirement.Key)"
    }
}
if ($source.Contains('["profile", "Current settings", (*) => CPDesktopNavigate(7)]')) {
    throw 'Retired header profile shortcut button remains'
}
$screenshotMigrationRequirements = [ordered]@{
    'direct screenshot-page constructor' = 'CPDesktopCreateScreenshotPage()'
    'owned screenshot registry group' = 'CPDesktop["shot"], [], CPDesktopLayoutScreenshot'
    'Game Text page title and subtitle' = '[1, "screenshot", "Game Text Translation", "Capture and translate text from your game.", "screenshot"'
    'Game Text sidebar label' = '["screenshot", "Game Text", (*) => CPDesktopNavigate(1)]'
    'Game Text translation icon' = 'Map("screenshot", 0xF2B7, "audioPage", 0xE767'
    'primary capture action label' = 'btnST.Text := "Capture && Translate"'
    'standalone capture action label' = 'btnTS.Text := "Make Capture"'
    'queued capture translation label' = 'btnSTO.Text := "Translate Captures"'
    'provider alias registration' = 'CPDesktopPageRegisterControl(1, "providerChoice"'
    'capture alias registration' = 'CPDesktopPageRegisterControl(1, "captureTranslate"'
    'model menu wiring' = 'shot["models"].OnEvent("Click", CPDesktopModelMenu)'
    'prompt menu wiring' = 'shot["prompts"].OnEvent("Click", CPDesktopPromptMenu)'
    'shared selection wiring' = 'ddlPrompt.OnEvent("Change", CPScreenshotAISelectionChanged)'
    'pre-construction startup preference guard' = 'translator := IsSet(chkOpenTW)'
    'pre-construction combo guard' = 'if IsSet(ddlProv)'
    'post-construction combo initialization' = 'Modern Screenshot, Audio, Explanation, Overlay, Terminology, Profiles,'
}
foreach ($requirement in $screenshotMigrationRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Screenshot-page migration source requirement missing: $($requirement.Key)"
    }
}
foreach ($retiredScreenshotControl in @('CPDesktopShotControls', 'btnIMG_Add', 'btnIMG_Del',
    'btnIMG_GM_Add', 'btnIMG_GM_Del', 'btnPrNew', 'btnPrDel', 'btnPrEdit')) {
    if ($source.Contains($retiredScreenshotControl)) {
        throw "Retired Game Text Translation control remains: $retiredScreenshotControl"
    }
}
$audioMigrationRequirements = [ordered]@{
    'direct audio-page constructor' = 'CPDesktopCreateAudioPage()'
    'owned audio registry group' = 'CPDesktop["audioPage"], [], CPDesktopLayoutAudio'
    'provider alias registration' = 'CPDesktopPageRegisterControl(2, "providerChoice"'
    'device alias registration' = 'CPDesktopPageRegisterControl(2, "listenDevice"'
    'diagnostic alias registration' = 'CPDesktopPageRegisterControl(2, "result"'
    'shared audio selection wiring' = 'ddlTR.OnEvent("Change", CPAudioAISelectionChanged)'
    'device-list initialization' = 'PopulateSpeakersList(speakerName)'
    'pre-construction audio combo guard' = 'if IsSet(ddlAProv)'
}
foreach ($requirement in $audioMigrationRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Audio-page migration source requirement missing: $($requirement.Key)"
    }
}
foreach ($retiredAudioControl in @('CPDesktopAudioControls', 'btnA_GM_Add', 'btnA_GM_Del',
    'btnTR_Add', 'btnTR_Del', 'lblLiveInput', 'lblLiveTranslation', 'txtAudioHelp')) {
    if ($source.Contains($retiredAudioControl)) {
        throw "Retired Audio Translation control remains: $retiredAudioControl"
    }
}
$explanationMigrationRequirements = [ordered]@{
    'direct explanation-page constructor' = 'CPDesktopCreateExplanationPage()'
    'owned explanation registry group' = 'CPDesktop["explanationPage"], [], CPDesktopLayoutExplanation'
    'explanation provider alias registration' = 'CPDesktopPageRegisterControl(4, "providerChoice"'
    'explanation prompt alias registration' = 'CPDesktopPageRegisterControl(4, "promptChoice"'
    'explanation action alias registration' = 'CPDesktopPageRegisterControl(4, "createExplanation"'
    'explanation preference alias registration' = 'CPDesktopPageRegisterControl(4, "saveLibrary"'
    'shared explanation selection wiring' = 'ddlEProv.OnEvent("Change", CPExplanationAISelectionChanged)'
    'pre-construction explanation prompt guard' = 'if IsSet(ddlEPr)'
    'explanation prompt initialization' = 'RefreshExplainPromptProfilesList(explainPromptProfile)'
    'post-construction alias initialization' = 'Modern Screenshot, Audio, Explanation, Overlay, Terminology, Profiles,'
}
foreach ($requirement in $explanationMigrationRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Explanation-page migration source requirement missing: $($requirement.Key)"
    }
}
foreach ($retiredExplanationControl in @('CPDesktopExplanationControls', 'btnEGem_Add', 'btnEGem_Del',
    'btnEOpenAI_Add', 'btnEOpenAI_Del', 'btnEPrNew', 'btnEPrDel', 'btnEPrEdit', 'txtExplainSaveInfo')) {
    if ($source.Contains($retiredExplanationControl)) {
        throw "Retired Explanation control remains: $retiredExplanationControl"
    }
}
$overlayMigrationRequirements = [ordered]@{
    'direct overlay-page constructor' = 'CPDesktopCreateOverlayPages()'
    'owned Translator overlay registry group' = 'CPDesktop["overlayPages"][3], [], CPDesktopLayoutOverlay.Bind(3)'
    'owned Explainer overlay registry group' = 'CPDesktop["overlayPages"][5], [], CPDesktopLayoutOverlay.Bind(5)'
    'overlay opacity alias registration' = 'CPDesktopPageRegisterControl(page, "opacitySlider"'
    'overlay font alias registration' = 'CPDesktopPageRegisterControl(page, "fontChoice"'
    'overlay position alias registration' = 'CPDesktopPageRegisterControl(page, "moveResize"'
    'Translator opacity wiring' = 'slTrans.OnEvent("Change", CPTranslatorOpacityChanged)'
    'Explainer opacity wiring' = 'slTrans_EW.OnEvent("Change", CPExplainerOpacityChanged)'
    'shared move/resize wiring' = 'bindings["position"].OnEvent("Click", StartOverlayAdjustment.Bind(title))'
    'post-construction overlay font initialization' = 'LoadFontsIntoCombo_EW()'
}
foreach ($requirement in $overlayMigrationRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Overlay-window migration source requirement missing: $($requirement.Key)"
    }
}
foreach ($retiredOverlayScaffold in @('CPDesktopOverlayLegacy', 'cpBeforeOverlay', 'twLabelX', 'ewLabelX')) {
    if ($source.Contains($retiredOverlayScaffold)) {
        throw "Retired Overlay Windows scaffold remains: $retiredOverlayScaffold"
    }
}
$terminologyMigrationRequirements = [ordered]@{
    'direct terminology-page constructor' = 'CPDesktopCreateTerminologyPage()'
    'owned terminology registry group' = 'CPDesktop["organizePages"][6], [], CPDesktopLayoutOrganize.Bind(6)'
    'terminology enable alias registration' = 'CPDesktopPageRegisterControl(6, "enabled"'
    'local glossary alias registration' = 'CPDesktopPageRegisterControl(6, "localChoice"'
    'model glossary alias registration' = 'CPDesktopPageRegisterControl(6, "modelChoice"'
    'terminology enable wiring' = 'chkUseTerminologyOverrides.OnEvent("Click", TerminologyOverridesChanged)'
    'local glossary selection wiring' = 'ddlENG.OnEvent("Change", CPTerminologyProfileChanged.Bind("en"))'
    'model glossary selection wiring' = 'ddlJPG.OnEvent("Change", CPTerminologyProfileChanged.Bind("jp"))'
    'terminology manager wiring' = 'btnJPG_Edit.OnEvent("Click", CPTerminologyManage.Bind("jp"))'
    'terminology profile initialization' = 'RefreshGlossaryProfilesList(jp2enGlossaryProfile, en2enGlossaryProfile)'
}
foreach ($requirement in $terminologyMigrationRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Terminology-page migration source requirement missing: $($requirement.Key)"
    }
}
foreach ($retiredTerminologyScaffold in @('CPDesktopOrganizeLegacy[6]', 'CPDesktopCaptureOrganizeControls(6',
    'txtGlossaryHelp1', 'txtGlossaryHelp2', 'txtGlossaryHelp3', 'txtGlossaryHelp4')) {
    if ($source.Contains($retiredTerminologyScaffold)) {
        throw "Retired Terminology scaffold remains: $retiredTerminologyScaffold"
    }
}
$profilesMigrationRequirements = [ordered]@{
    'direct profiles-page constructor' = 'CPDesktopCreateProfilesPage()'
    'owned profiles registry group' = 'CPDesktop["organizePages"][7], [], CPDesktopLayoutOrganize.Bind(7)'
    'profile selector alias registration' = 'CPDesktopPageRegisterControl(7, "profileChoice"'
    'startup selector alias registration' = 'CPDesktopPageRegisterControl(7, "startupChoice"'
    'profile state alias registration' = 'CPDesktopPageRegisterControl(7, "profileState"'
    'profile selection wiring' = 'ddlGameProfile.OnEvent("Change", GameProfileUpdateSummary)'
    'startup selection wiring' = 'ddlStartupOverlays.OnEvent("Change", CPStartupOverlaysChanged)'
    'profile creation wiring' = 'btnGameProfileAdd.OnEvent("Click", CreateGameProfile)'
    'profile save wiring' = 'btnGameProfileSave.OnEvent("Click", SaveSelectedGameProfile)'
    'profile apply wiring' = 'btnGameProfileApply.OnEvent("Click", ApplySelectedGameProfile)'
    'profile deletion wiring' = 'btnGameProfileDelete.OnEvent("Click", DeleteSelectedGameProfile)'
    'startup selection initialization' = 'CPStartupOverlaysSync()'
    'profile list initialization' = 'RefreshGameProfilesList()'
}
foreach ($requirement in $profilesMigrationRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Profiles-page migration source requirement missing: $($requirement.Key)"
    }
}
foreach ($retiredProfilesScaffold in @('CPDesktopOrganizeLegacy[7]', 'CPDesktopCaptureOrganizeControls(7',
    'lblGameProfilesTitle', 'txtGameProfileIntro', 'txtGameProfileGlobal', 'txtStartupOverlayHelp',
    'txtGameProfileDetails')) {
    if ($source.Contains($retiredProfilesScaffold)) {
        throw "Retired Profiles scaffold remains: $retiredProfilesScaffold"
    }
}
$controlsMigrationRequirements = [ordered]@{
    'direct controls-page constructor' = 'CPDesktopCreateControlsPage()'
    'owned controls registry group' = 'CPDesktop["organizePages"][8], [], CPDesktopLayoutOrganize.Bind(8)'
    'keyboard selector alias registration' = 'CPDesktopOrganizeButton(8, "keyboard", "Keyboard"'
    'controller selector alias registration' = 'CPDesktopOrganizeButton(8, "gamepad", "Controller"'
    'keyboard binding alias registration' = 'CPDesktopPageRegisterControl(8, "keyboardBinding_" action'
    'controller binding alias registration' = 'CPDesktopPageRegisterControl(8, "controllerBinding_" action'
    'controller option alias registration' = 'CPDesktopPageRegisterControl(8, "controllerEnabled"'
    'keyboard change wiring' = 'hkBtnChg[action].OnEvent("Click", Hotkey_Row_Change.Bind(action))'
    'keyboard disable wiring' = 'hkBtnDis[action].OnEvent("Click", Hotkey_Row_Disable.Bind(action))'
    'keyboard default wiring' = 'hkBtnDef[action].OnEvent("Click", Hotkey_Row_Default.Bind(action))'
    'controller assignment wiring' = 'CPControllerAssignButtons[action].OnEvent("Click", CPControllerAssign.Bind(action))'
    'controller disable wiring' = 'CPControllerDisableButtons[action].OnEvent("Click", CPControllerDisable.Bind(action))'
    'binding initialization' = 'CPControllerLoadBindings()'
    'view initialization' = 'CPSetControlsView(IniRead(iniPath, "controller_inputs", "view", "keyboard"), false)'
    'view-specific registry visibility' = 'for ctrl in CPControlsKeyboardControls'
}
foreach ($requirement in $controlsMigrationRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Controls-page migration source requirement missing: $($requirement.Key)"
    }
}
foreach ($retiredControlsScaffold in @('CPDesktopOrganizeLegacy[8]', 'CPDesktopCaptureOrganizeControls(8',
    'CPLayoutControllerOptions(', 'txtControllerDpadNote', 'controlsActionX', 'controlsBindingX',
    'controllerOptionsX', 'controllerTopY', 'lblKeyboardAction', 'lblControllerAction')) {
    if ($source.Contains($retiredControlsScaffold)) {
        throw "Retired Controls scaffold remains: $retiredControlsScaffold"
    }
}
$apiKeysMigrationRequirements = [ordered]@{
    'direct API Keys-page constructor' = 'CPDesktopCreateApiKeysPage()'
    'owned API Keys registry group' = 'CPDesktop["organizePages"][9], [], CPDesktopLayoutOrganize.Bind(9)'
    'in-app entry alias registration' = 'CPDesktopPageRegisterControl(9, "inAppEntry"'
    'Gemini key alias registration' = 'CPDesktopPageRegisterControl(9, "geminiKey"'
    'OpenAI key alias registration' = 'CPDesktopPageRegisterControl(9, "openAIKey"'
    'masked Gemini field' = 'ui.AddEdit("x0 y0 w420 h34 Hidden Password")'
    'in-app entry wiring' = 'cbApiInApp.OnEvent("Click", CPApiInAppChanged)'
    'key dirty-state wiring' = 'eGemini.OnEvent("Change", UpdateEnvDirty)'
    'key save wiring' = 'btnSaveEnv.OnEvent("Click", SaveApiEnv)'
    'key delete wiring' = 'btnDelEnv.OnEvent("Click", DeleteEnvFile)'
    'environment-variable wiring' = 'btnOpenEnvVars.OnEvent("Click", OpenWindowsEnvironmentVariables)'
    'existing key initialization' = 'prefOpenAI := ParseEnvLine(envBody, "OPENAI_API_KEY")'
    'named enablement synchronizer' = 'ToggleApiKeyControls(*) {'
    'inline save feedback' = 'CPApiKeysSetNotice("In-app API keys saved.")'
    'inline removal feedback' = 'CPApiKeysSetNotice("In-app API keys removed.")'
}
foreach ($requirement in $apiKeysMigrationRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "API Keys-page migration source requirement missing: $($requirement.Key)"
    }
}
if (!$source.Contains('c["about"].OnEvent("Click", CPDesktopAppearanceAbout.Bind(s))')) {
    throw 'Appearance dialog About wiring is missing'
}
foreach ($retiredApiKeysScaffold in @('CPDesktopOrganizeLegacy[9]', 'CPDesktopCaptureOrganizeControls(9',
    'txtApiHelp1', 'txtApiHelp2', 'txtApiHelp3', 'txtApiHelp4', 'ToggleApiKeyControls :=',
    'CPDesktopPageRegisterControl(9, "aboutAction"')) {
    if ($source.Contains($retiredApiKeysScaffold)) {
        throw "Retired API Keys scaffold remains: $retiredApiKeysScaffold"
    }
}
$apiSaveSource = [regex]::Match($source, '(?ms)^SaveApiEnv\([^\r\n]*\)\s*\{.*?^\}').Value
$apiDeleteSource = [regex]::Match($source, '(?ms)^DeleteEnvFile\([^\r\n]*\)\s*\{.*?^\}').Value
if (!$apiSaveSource -or !$apiDeleteSource -or
    $apiSaveSource.Contains('Toast(') -or $apiDeleteSource.Contains('Toast(')) {
    throw 'Desktop API-key feedback must remain inside the Settings page, without a floating toast window.'
}
$advancedIniRequirements = [ordered]@{
    'Python executable load' = 'pythonExe       := Load("pythonExe",        defPython)'
    'direct-output default off' = 'defDirectModelOutput := 0'
    'debug default off' = 'defDebugMode := 0'
    'direct-output INI load' = 'directModelOutput := LoadInt("directModelOutput", defDirectModelOutput, "cfg") ? 1 : 0'
    'debug INI load' = 'debugMode := LoadInt("debugMode", defDebugMode, "cfg")'
    'Python executable persistence' = 'IniWrite(pythonExe,       iniPath, "cfg", "pythonExe")'
    'direct-output persistence' = 'IniWrite(directModelOutput, iniPath, "cfg", "directModelOutput")'
    'debug persistence' = 'IniWrite(debugMode, iniPath, "cfg", "debugMode")'
    'retired visibility-key cleanup' = 'IniDelete(iniPath, "cfg", "showPathsTab")'
}
foreach ($requirement in $advancedIniRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Advanced INI source requirement missing: $($requirement.Key)"
    }
}
foreach ($retiredPathsScaffold in @('CPDesktopCreatePathsPage', 'CPDesktopLayoutPaths',
    'CPBigBoxOpenPathEditor', 'CPBigBoxTogglePathOption', 'CPBigBoxWritePath',
    'UpdatePathsDirtyState', 'SaveEditedPaths', 'ConfirmUnsavedPaths', 'pathsTab',
    'CPDesktopPageRegisterControl(10', '[10, "paths"', '"path_python"')) {
    if ($source.Contains($retiredPathsScaffold)) {
        throw "Retired Paths UI remains: $retiredPathsScaffold"
    }
}
$audioRuntimeRequirements = [ordered]@{
    'worker-owned session marker' = 'EnvSet("AUDIO_SESSION_FILE", gAudioSessionFile)'
    'immediate footer refresh' = 'CPDesktopRefreshStatus()'
    'shutdown-owned audio cleanup' = 'try StopAudioCore(false, false)'
    'compositor-owned toast image' = 'ToastPresent(CPToastGui, CPToastText, background, foreground)'
}
foreach ($requirement in $audioRuntimeRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Audio runtime source requirement missing: $($requirement.Key)"
    }
}
$audioStartBody = [regex]::Match($source, '(?ms)^StartAudioCore\([^\r\n]*\)\s*\{.*?^\}').Value
$audioToggleBody = [regex]::Match($source, '(?ms)^ToggleAudioFromButton\([^\r\n]*\)\s*\{.*?^\}').Value
if (!$audioStartBody -or $audioStartBody.Contains('RunWait(') -or $audioStartBody.Contains('Toast(') -or
    !$audioToggleBody.Contains('AudioIsRunning(true)')) {
    throw 'Audio start/stop must recover worker state without a blocking retry or corner toast.'
}
$desktopLayoutBody = [regex]::Match($source, '(?ms)^CPDesktopLayout\([^\r\n]*\)\s*\{.*?^\}').Value
if (!$desktopLayoutBody) {
    throw 'Desktop layout function not found for page-registry validation.'
}
foreach ($legacyGroup in @('CPDesktopShotControls', 'CPDesktopAudioControls',
    'CPDesktopExplanationControls', 'CPDesktopOverlayLegacy', 'CPDesktopOrganizeLegacy')) {
    if ($desktopLayoutBody.Contains($legacyGroup)) {
        throw "Desktop layout must dispatch through the page registry, not $legacyGroup."
    }
}
if (!$source.Contains('(state & 0x200) = 0') -or
    !$source.Contains('borderInset := Max(1, Ceil(penWidth / 2))')) {
    throw 'Desktop owner-draw must hide pointer focus cues and keep every focus-border edge in bounds.'
}
$navigationPaintRequirements = [ordered]@{
    'painted focus redraw helper' = 'CPDesktopRefreshPaintedButton(hwnd, forceRedraw := true)'
    'painted focus entry path' = 'if CPDesktopRefreshPaintedButton(hwnd)'
    'painted focus restore path' = 'if CPDesktopRefreshPaintedButton(CPFocusVisualHwnd)'
    'controller shares keyboard navigation' = 'CPNavMove(command)'
}
foreach ($requirement in $navigationPaintRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Desktop navigation-paint source requirement missing: $($requirement.Key)"
    }
}
$desktopActivationRequirements = [ordered]@{
    'owner-drawn button activation' = 'DllCall("user32\PostMessageW", "ptr", hwnd, "uint", 0x00F5'
    'keyboard root cancel fallback' = 'CPNavCancel()'
    'controller root cancel fallback' = 'CPNavCancel(cancelFallback)'
}
foreach ($requirement in $desktopActivationRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Desktop activation source requirement missing: $($requirement.Key)"
    }
}
$sectionNavigationRequirements = [ordered]@{
    'modern sidebar page order' = 'desktopOrder := [1, 2, 4, 3, 5, "study", 7, 8, 6, 9]'
    'available-page filtering' = 'if ArrayIndexOf(CPTabVisiblePages, page)'
    'section switching uses semantic order' = 'pages := CPDesktopSectionNavigationPages()'
    'Study is a focus-only stop' = 'if nextPage = "study"'
}
foreach ($requirement in $sectionNavigationRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Desktop section-navigation source requirement missing: $($requirement.Key)"
    }
}
$candidate720pRequirements = [ordered]@{
    '720p-safe Review minimum' = '"Review saved sentences and vocabulary before adding them to Anki.", 960, 620)'
    'compact Review threshold' = 'compact := h < 760'
    'expanded compact assessment' = 'summaryH := compact ? 64 : 44'
    '720p-safe initial Review height' = 'StudyDesktopDialogShow(scState, 1040, 640, scScopeDdl)'
    'taskbar-aware borderless maximize' = 'StudyDesktopConstrainMaximize(hwnd, minMaxInfo)'
    'single-click candidate handler' = 'StudyCandidatesItemSelected(scState, scList, scRow, scSelected)'
    'stale deselection guard' = 'StudyCandidatesSelectionSettled.Bind(scState, scRevision)'
    'sentence single-click wiring' = 'scSentenceList.OnEvent("ItemSelect", StudyCandidatesItemSelected.Bind(scState))'
    'vocabulary single-click wiring' = 'scVocabularyList.OnEvent("ItemSelect", StudyCandidatesItemSelected.Bind(scState))'
    'sentence pointer-row wiring' = 'scSentenceList.OnEvent("Click", StudyCandidatesRowClicked.Bind(scState))'
    'vocabulary pointer-row wiring' = 'scVocabularyList.OnEvent("Click", StudyCandidatesRowClicked.Bind(scState))'
    'settled candidate-tab repaint' = 'SetTimer(StudyDesktopCandidatesRedrawTabs.Bind(s, index), -1)'
}
foreach ($requirement in $candidate720pRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "720p Review-layout source requirement missing: $($requirement.Key)"
    }
}
$studyDialog720pRequirements = [ordered]@{
    'taskbar-aware Study dialog sizing' = 'StudyDesktopDialogFitToWorkArea(g, preferredW, preferredH)'
    'responsive explanation generation dialog' = 'StudyDesktopDialogShow(srNewState, 920, 740, srProvider, true)'
    'compact explanation generation threshold' = 'compact := h < 700'
    'responsive explanation prompt dialog' = 's["controls"]["editor"], s["mode"] = "edit")'
    'prompt edit initial caret' = 'StudyReaderPromptPlaceInitialCaret(s)'
}
foreach ($requirement in $studyDialog720pRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "720p Study-dialog source requirement missing: $($requirement.Key)"
    }
}
$studyLibraryPointerRequirements = [ordered]@{
    'keyboard and controller row-detail wiring' = 'slList.OnEvent("ItemFocus", StudyLibraryGroupFocused.Bind(slState))'
    'exact pointer row-detail wiring' = 'slList.OnEvent("Click", StudyLibraryGroupFocused.Bind(slState))'
}
foreach ($requirement in $studyLibraryPointerRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Study Library pointer-selection source requirement missing: $($requirement.Key)"
    }
}
$studyLibrarySelectorRequirements = [ordered]@{
    'integrated management action' = 'slItems.Push("Manage Study Libraries…")'
    'management action restores active selection' = 'StudyLibraryRefreshLibrarySelector(slState, slState["libraryName"])'
    'management action opens existing workflow' = 'StudyLibraryOpenManager(slState)'
    'separated final dropdown action' = 'CPComboSeparatorBefore[slState["libraryDdl"].Hwnd] := slManageIndex'
}
foreach ($requirement in $studyLibrarySelectorRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Study Library selector source requirement missing: $($requirement.Key)"
    }
}
if ($source.Contains('slNewLibraryButton :=')) {
    throw 'Retired standalone Study Library management button remains'
}
$hotkeyRequirements = [ordered]@{
    'initial themed shortcut display' = 's["controls"]["editor"].Text := display != "" ? display : "None"'
    'shortcut display z-order repair' = 'CPHotkeyDialogPresentEditor(s)'
    'in-app shortcut feedback' = 'CPHotkeySetNotice(notice)'
    'shortcut removal wording' = '"Keyboard shortcut removed · "'
    'shortcut save wording' = '"Keyboard shortcut saved · "'
    'shortcut default wording' = '"Default keyboard shortcut restored · "'
}
foreach ($requirement in $hotkeyRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Keyboard-shortcut source requirement missing: $($requirement.Key)"
    }
}
$desktopInputFieldRequirements = [ordered]@{
    'shared dark input frame' = 'CPDesktopRegisterPaint(frame, "fieldFrame", "panel")'
    'legacy input edge removal' = 'ctrl.Opt("-Border -E0x200 -VScroll")'
    'overlay font-size field frame' = 'CPDesktopRegisterInputField(page, "fontSize", edFSize)'
    'desktop checkbox painter' = 'CPDesktopPrepareNativePaint(ctrl, "checkbox")'
    'desktop numeric-stepper painter' = 'CPDesktopPrepareNativePaint(spinner, "spinner")'
    'keyboard binding field frame' = '8, "keyboardBinding_" action, hkEdits[action], "keyboard"'
    'controller binding field frame' = 'CPControllerBindingEdits[action], "controller"'
    'Gemini key field frame' = 'CPDesktopRegisterInputField(9, "geminiKey", eGemini)'
    'OpenAI key field frame' = 'CPDesktopRegisterInputField(9, "openAIKey", eOpenAI)'
}
foreach ($requirement in $desktopInputFieldRequirements.GetEnumerator()) {
    if (!$source.Contains($requirement.Value)) {
        throw "Desktop input-field source requirement missing: $($requirement.Key)"
    }
}
foreach ($functionName in @('Hotkeys_OnApply', 'Hotkeys_OnRevert')) {
    $body = [regex]::Match($source, '(?ms)^' + $functionName + '\([^\r\n]*\)\s*\{.*?^\}').Value
    if (!$body -or $body.Contains('ToolTip(')) {
        throw "Keyboard-shortcut feedback must stay inside the app UI: $functionName"
    }
}
$output = New-Item -ItemType Directory -Path $OutputDirectory -Force
New-Item -ItemType Directory -Path (Join-Path $output.FullName 'assets') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repo 'assets/bigbox-logo.png') -Destination (Join-Path $output.FullName 'assets')
Copy-Item -LiteralPath (Join-Path $repo 'assets/bigbox-game-placeholder.png') -Destination (Join-Path $output.FullName 'assets')
Copy-Item -LiteralPath (Join-Path $repo 'assets/desktop-logo.png') -Destination (Join-Path $output.FullName 'assets')
Copy-Item -LiteralPath (Join-Path $repo 'JRPG Translator.ahk') -Destination $output.FullName
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'audio_runtime_fixture.ahk') -Destination $output.FullName
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
if ($ChapterOnly) {
    $generated = $generated.Replace('; @STUDY_ONLY@', 'TestDesktopChapterSizing()' + "`r`n" + '    ui.Destroy()' + "`r`n" + '    FileAppend("PASS: " TestAssertions " chapter sizing assertions.`n", "*")' + "`r`n" + '    ExitApp(0)')
}
if ($PromptUnsavedOnly) {
    $generated = $generated.Replace('; @STUDY_ONLY@', 'TestPromptUnsavedChanges()' + "`r`n" + '    ui.Destroy()' + "`r`n" + '    FileAppend("PASS: " TestAssertions " unsaved-prompt assertions.`n", "*")' + "`r`n" + '    ExitApp(0)')
    $generated += "`r`n" + [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'prompt_unsaved_harness.ahk'))
}
if ($GenerationActivityOnly) {
    $generated = $generated.Replace('; @STUDY_ONLY@', 'TestGenerationActivity()' + "`r`n" + '    ui.Destroy()' + "`r`n" + '    FileAppend("PASS: " TestAssertions " generation activity assertions.`n", "*")' + "`r`n" + '    ExitApp(0)')
    $generated += "`r`n" + [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'generation_activity_harness.ahk'))
    # Run production request lifecycles against synthetic process results only:
    # no provider, real database, or Anki is contacted.
    foreach ($name in @('StudyReaderGenerateVocabularyExample', 'StudyReaderGenerateNewVersion')) {
        $body = [regex]::Match($source, '(?ms)^' + $name + '\(.*?^}').Value
        if (!$body) { throw "Missing generation function: $name" }
        $generated += "`r`n" + $body.Replace($name + '(', 'Test' + $name + '(').
            Replace('RunWait(', 'TestGenerationWait(').
            Replace('StudyReaderAnkiMessage(', 'TestGenerationNotice(').
            Replace('CPThemedOwnedMessage(', 'TestGenerationNotice(').
            Replace('StudyLibraryLoadGroup(', 'TestGenerationLoadGroup(')
    }
}
if ($ShutdownOnly) {
    $generated = $generated.Replace('; @STUDY_ONLY@', 'DesktopTestExitWithLiveWindow()')
}
if ($DialogsOnly) {
    $generated = $generated.Replace('; @STUDY_ONLY@', 'TestDesktopStudyDialogs()' + "`r`n" + '    ui.Destroy()' + "`r`n" + '    FileAppend("PASS: " TestAssertions " desktop dialog assertions.`n", "*")' + "`r`n" + '    ExitApp(0)')
}
if ($ResizeOnly) {
    $generated = $generated.Replace('; @STUDY_ONLY@', 'TestDesktopResizing()' + "`r`n" + '    ui.Destroy()' + "`r`n" + '    FileAppend("PASS: " TestAssertions " desktop resize assertions.`n", "*")' + "`r`n" + '    ExitApp(0)')
}
if ($AudioFeedbackOnly) {
    $generated = $generated.Replace('; @STUDY_ONLY@', 'if TestAudioFeedbackBegin()' + "`r`n" + '        return')
    $generated += "`r`n" + [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'audio_feedback_ui_harness.ahk'))
}
if ($ProfileFeedbackOnly) {
    # Never discover/send commands to the user's real overlay windows.
    $isolatedSource = $source.Replace('WinExist(title)', 'WinExist("Synthetic profile " title)').Replace('WinExist("Translator")', 'WinExist("Synthetic profile Translator")').Replace('WinExist("Explainer")', 'WinExist("Synthetic profile Explainer")')
    [IO.File]::WriteAllText((Join-Path $output.FullName 'JRPG Translator.ahk'), $isolatedSource, [Text.UTF8Encoding]::new($true))
    $generated = $generated.Replace('; @STUDY_ONLY@', 'if TestProfileFeedbackBegin()' + "`r`n" + '        return')
    $generated += "`r`n" + [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'audio_feedback_ui_harness.ahk'))
    $generated += "`r`n" + [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'profile_feedback_ui_harness.ahk'))
}
$script = Join-Path $output.FullName 'desktop-generated.ahk'
[IO.File]::WriteAllText($script, $generated, [Text.UTF8Encoding]::new($true))
$stdout = Join-Path $output.FullName 'desktop.stdout.txt'
$stderr = Join-Path $output.FullName 'desktop.stderr.txt'
$process = Start-Process -FilePath $AutoHotkey -ArgumentList @('/ErrorStdOut', ('"' + $script + '"')) -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
$testTimeout = if ($AudioFeedbackOnly -or $ProfileFeedbackOnly) { 60000 } else { 300000 }
if (!$process.WaitForExit($testTimeout)) {
    $process.Kill()
    throw "Desktop test timed out. Logs: $output"
}
$process.WaitForExit()
$result = [IO.File]::ReadAllText($stdout)
$errors = [IO.File]::ReadAllText($stderr)
if ($process.ExitCode -ne 0 -or $errors -or $result.Contains('Warning:') -or $result.Contains('FAIL:')) { throw "$result`n$errors`nArtifacts: $output" }
$namingChecks = 0
if ($GenerationActivityOnly) {
    foreach ($line in [IO.File]::ReadAllLines((Join-Path $output.FullName 'activity-bounds.txt'))) {
        $parts = $line.Split('|')
        $first = [Drawing.Bitmap]::new((Join-Path $output.FullName ("activity-$($parts[0])-1.png")))
        $second = [Drawing.Bitmap]::new((Join-Path $output.FullName ("activity-$($parts[0])-2.png")))
        try {
            $changed = 0
            for ($y = [int]$parts[2]; $y -lt ([int]$parts[2] + [int]$parts[4]); $y++) {
                for ($x = [int]$parts[1]; $x -lt ([int]$parts[1] + [int]$parts[3]); $x++) {
                    if ($first.GetPixel($x, $y) -ne $second.GetPixel($x, $y)) { $changed++ }
                }
            }
            if ($changed -eq 0) { throw "Activity bar is not animating: $($parts[0])" }
            Write-Output "PASS: $($parts[0]) activity bar animates ($changed changed pixels)."
        } finally { $first.Dispose(); $second.Dispose() }
    }
}
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
if ($StudyOnly -or $ShutdownOnly -or $DialogsOnly -or $ResizeOnly -or $AudioFeedbackOnly -or $ProfileFeedbackOnly -or $ChapterOnly -or $PromptUnsavedOnly -or $GenerationActivityOnly) {
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

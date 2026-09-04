#Requires AutoHotkey v2.0
#SingleInstance Off
#Warn All, StdOut
#NoTrayIcon

global CPStudyLibraryState := 0
global CPStudyReaderState := 0
global CPStudyCandidateState := 0
global __CP_STUDY_NAV_ITEMS := []
global CPStudySyntheticKeyDepth := 0
global CPStudyComboTransactions := Map()
global CPControllerLastNativeNavigationAt := Map()
global CPThemedDialogHwnds := Map()
global controlDarkMode := true
global CPThemeBrushWindow := 1
global ui := 0, CPBigBoxGui := 0
global TestCount := 0
global TestLibraryOpens := 0
global TestCandidateOpens := 0
global TestCandidateActions := 0
global TestCandidateActionRow := 0
global TestCandidateActionPrefersAnki := false
global TestColumnRelayouts := 0
global TestLibraryCloses := 0
global TestReaderCloses := 0
global TestCandidateCloses := 0
global TestReaderSteps := []
global TestSyntheticKeys := []
global TestButtonClicks := 0
global TestAnkiStateToClose := 0
global TestAnkiCloseDuringBridge := false
global TestCandidateScopeChanges := 0
global TestCandidateAiChanges := 0
global TestCandidateAsyncStarts := 0
global TestCandidateAsyncCallback := 0
global TestCandidateSnapshotsRead := 0
global TestLibraryGroupLoads := 0
global TestLibraryChanges := 0

; @STUDY_SOURCE@

try {
    libraryGui := Gui("+AlwaysOnTop", "Study controller library fixture")
    libraryList := libraryGui.AddListView("x10 y10 w300 h150", ["Entry"])
    libraryList.Add(, "First")
    libraryList.Add(, "Second")
    libraryList.Add(, "Third")
    libraryButton := libraryGui.AddButton("x330 y10 w120 h35", "Action")
    libraryEdit := libraryGui.AddEdit(
        "x10 y175 w300 h55 ReadOnly Multi VScroll",
        TestMultilineText()
    )
    libraryBottomButton := libraryGui.AddButton(
        "x10 y240 w120 h30", "Bottom action"
    )
    libraryDdl := libraryGui.AddDropDownList(
        "x330 y55 w120 Hidden", ["First", "Second", "Third"]
    )
    libraryDdl.Choose(1)
    libraryButton.OnEvent("Click", (*) => TestRecordButtonClick())
    libraryGui.Show("x80 y80 w470 h285")
    CPStudyLibraryState := Map(
        "gui", libraryGui,
        "list", libraryList,
        "libraryDdl", libraryDdl,
        "libraryName", "First",
        "ankiButton", libraryButton,
        "closed", false
    )
    libraryDdl.OnEvent(
        "Change", StudyLibraryLibraryChanged.Bind(CPStudyLibraryState)
    )
    WinActivate("ahk_id " libraryGui.Hwnd)
    libraryList.Modify(1, "Select Focus Vis")
    libraryList.Focus()
    Sleep(30)

    TestAssert(StudyControllerSurfaceForWindow(
        libraryGui.Hwnd, &librarySurface
    ) && librarySurface["kind"] = "library",
        "Library is a controller navigation root")
    TestAssert(StudyControllerSurfaceIsRoot(libraryGui.Hwnd),
        "Library root is distinguished from owned dialogs")
    TestAssert(CPControllerNavigationTarget(libraryGui.Hwnd) = libraryGui.Hwnd,
        "Controller polling recognizes the Library foreground window")

    focusReturnGui := Gui(
        "+Owner" libraryGui.Hwnd, "Hidden candidate focus fixture"
    )
    focusReturnList := focusReturnGui.AddListView(
        "x10 y10 w150 h45", ["Candidate"]
    )
    focusReturnVocabularyList := focusReturnGui.AddListView(
        "x10 y60 w150 h45", ["Vocabulary"]
    )
    focusReturnGui.Show("x600 y80 w170 h115")
    focusReturnHwnd := focusReturnGui.Hwnd
    WinActivate("ahk_id " focusReturnGui.Hwnd)
    focusReturnState := Map(
        "gui", focusReturnGui,
        "sentenceList", focusReturnList,
        "vocabularyList", focusReturnVocabularyList,
        "libraryState", CPStudyLibraryState,
        "guiClosed", false
    )
    StudyCandidatesDestroyGui(focusReturnState)
    TestAssert(focusReturnState["guiClosed"]
        && !DllCall("user32\IsWindow", "ptr", focusReturnHwnd, "int"),
        "A closing asynchronous Review destroys its native window immediately")
    TestAssert(StudyCandidatesRestoreLibraryFocus(focusReturnState),
        "Closing a busy Review immediately returns navigation to its Library")
    libraryList.Focus()

    focusReloadState := Map(
        "suspend", false,
        "groups", [Map("id", 41), Map("id", 42)],
        "currentGroupId", 41
    )
    StudyLibraryGroupFocused(focusReloadState, libraryList, 1)
    TestAssert(TestLibraryGroupLoads = 0,
        "Refocusing the loaded Library row does not launch another detail bridge")
    StudyLibraryGroupFocused(focusReloadState, libraryList, 2)
    TestAssert(TestLibraryGroupLoads = 1
        && focusReloadState["currentGroupId"] = 42,
        "Focusing a different Library row still loads its explanation")

    StudyControllerDispatchNavigation("Down", libraryGui.Hwnd)
    TestAssert(libraryList.GetNext(0, "F") = 2,
        "D-pad Down moves the active Library table row")
    libraryList.Modify(3, "Select Focus Vis")
    StudyControllerDispatchNavigation("Down", libraryGui.Hwnd)
    TestAssert(StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryEdit.Hwnd,
        "D-pad Down exits the Library table after its last real row")
    libraryList.Modify(1, "Select Focus Vis")
    libraryList.Focus()
    StudyControllerDispatchNavigation("Up", libraryGui.Hwnd)
    TestAssert(StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryButton.Hwnd,
        "D-pad Up exits the Library table before its first real row")
    libraryList.Modify(2, "Select Focus Vis")
    libraryList.Focus()
    StudyControllerDispatchNavigation("Activate", libraryGui.Hwnd)
    TestAssert(TestLibraryOpens = 1,
        "A opens the selected Library explanation")

    libraryButton.Focus()
    StudyControllerDispatchNavigation("Activate", libraryGui.Hwnd)
    Sleep(30)
    TestAssert(TestButtonClicks = 1,
        "A activates ordinary Study buttons")
    StudyControllerDispatchNavigation("Left", libraryGui.Hwnd)
    TestAssert(StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryList.Hwnd,
        "Spatial D-pad navigation moves from an action to the table")

    libraryEdit.Focus()
    beforeLine := SendMessage(0x00CE, 0, 0, libraryEdit.Hwnd) ; EM_GETFIRSTVISIBLELINE
    StudyControllerDispatchNavigation("Down", libraryGui.Hwnd)
    afterLine := SendMessage(0x00CE, 0, 0, libraryEdit.Hwnd)
    TestAssert(afterLine > beforeLine,
        "D-pad scrolls read-only Study text without moving focus")
    SendMessage(0x0115, 7, 0, libraryEdit.Hwnd) ; WM_VSCROLL, SB_BOTTOM
    libraryEdit.Focus()
    StudyControllerDispatchNavigation("Down", libraryGui.Hwnd)
    TestAssert(
        StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryBottomButton.Hwnd,
        "D-pad Down exits a read-only Study field at its scroll boundary"
    )
    SendMessage(0x0115, 6, 0, libraryEdit.Hwnd) ; WM_VSCROLL, SB_TOP
    libraryEdit.Focus()
    StudyControllerDispatchNavigation("Up", libraryGui.Hwnd)
    TestAssert(StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryList.Hwnd,
        "D-pad Up exits a read-only Study field at its scroll boundary")
    SendMessage(0x00B1, 0, -1, libraryEdit.Hwnd) ; EM_SETSEL, select all
    StudyControllerSetFocus(libraryGui.Hwnd, libraryEdit.Hwnd)
    librarySelection := SendMessage(
        0x00B0, 0, 0, libraryEdit.Hwnd
    ) ; EM_GETSEL for short fixture text
    TestAssert((librarySelection & 0xFFFF) = 0
        && ((librarySelection >> 16) & 0xFFFF) = 0,
        "Controller focus does not select all read-only Study text")

    libraryDdl.Visible := true
    libraryDdl.Focus()
    StudyControllerDispatchNavigation("Activate", libraryGui.Hwnd)
    TestAssert(CPComboDropped(libraryDdl.Hwnd)
        && StudyControllerComboPreviewActive(libraryDdl.Hwnd),
        "A opens the Library selector as a pending controller selection")
    StudyControllerDispatchNavigation("Down", libraryGui.Hwnd)
    Sleep(100)
    TestAssert(libraryDdl.Value = 2
        && TestLibraryChanges = 0
        && CPComboDropped(libraryDdl.Hwnd),
        "D-pad previews another Library without selecting or closing it")
    StudyControllerDispatchNavigation("Activate", libraryGui.Hwnd)
    TestAssert(TestLibraryChanges = 1
        && CPStudyLibraryState["libraryName"] = "Second"
        && !CPComboDropped(libraryDdl.Hwnd)
        && !StudyControllerComboPreviewActive(libraryDdl.Hwnd),
        "A explicitly applies the highlighted Library exactly once")
    StudyControllerDispatchNavigation("Activate", libraryGui.Hwnd)
    StudyControllerDispatchNavigation("Down", libraryGui.Hwnd)
    StudyControllerDispatchNavigation("Cancel", libraryGui.Hwnd)
    TestAssert(libraryDdl.Value = 2
        && CPStudyLibraryState["libraryName"] = "Second"
        && TestLibraryChanges = 1
        && !CPComboDropped(libraryDdl.Hwnd)
        && !StudyControllerComboPreviewActive(libraryDdl.Hwnd),
        "B cancels a pending Library choice without changing the Library")
    libraryDdl.Visible := false
    libraryList.Focus()

    bigBoxColumnState := Map("columns", [
        Map("key", "updated", "width", 100, "visible", true),
        Map("key", "source", "width", 200, "visible", true),
        Map("key", "tags", "width", 80, "visible", false)
    ])
    bigBoxWidths := StudyLibraryBigBoxColumnWidths(
        bigBoxColumnState, 800, 2
    )
    TestAssert(bigBoxWidths[1] = 200
        && bigBoxWidths[2] = 600
        && bigBoxWidths[3] = 0,
        "Fullscreen Library columns scale physically and fill spare table width")
    bigBoxColumnState["bigBoxPresentation"] := true
    bigBoxColumnState["tableModeActive"] := true
    bigBoxColumnState["tableModeSurface"] := "library"
    tableModeWidths := StudyLibraryBigBoxColumnWidths(
        bigBoxColumnState, 800, 1
    )
    TestAssert(tableModeWidths[1] + tableModeWidths[2] > 800
        && tableModeWidths[3] = 0,
        "Library table mode preserves readable visible columns for horizontal browsing")
    columnDraft := [
        Map("key", "updated", "width", 125, "defaultWidth", 125,
            "visible", true, "required", true),
        Map("key", "source", "width", 260, "defaultWidth", 260,
            "visible", true, "required", false),
        Map("key", "tags", "width", 150, "defaultWidth", 150,
            "visible", true, "required", false)
    ]
    columnEditor := Map(
        "draft", columnDraft, "selectedIndex", 2, "closed", false
    )
    StudyLibraryBigBoxColumnsMove(columnEditor, 1)
    TestAssert(columnEditor["selectedIndex"] = 3
        && columnEditor["draft"][3]["key"] = "source",
        "Fullscreen column layout can move the chosen column without losing it")
    StudyLibraryBigBoxColumnsWidth(columnEditor, 20)
    TestAssert(columnEditor["draft"][3]["width"] = 280,
        "Fullscreen column layout adjusts the chosen width in readable steps")
    StudyLibraryBigBoxColumnsToggle(columnEditor)
    TestAssert(!columnEditor["draft"][3]["visible"],
        "Fullscreen column layout stages optional-column visibility")
    StudyLibraryBigBoxColumnsResetWidth(columnEditor)
    TestAssert(columnEditor["draft"][3]["width"] = 260,
        "Fullscreen column layout restores the selected default width")
    columnEditor["selectedIndex"] := 1
    StudyLibraryBigBoxColumnsToggle(columnEditor)
    TestAssert(columnEditor["draft"][1]["visible"],
        "Fullscreen column layout keeps required columns visible")
    columnGui := Gui("+AlwaysOnTop", "Column detail layout fixture")
    columnHeading := columnGui.AddText("x10 y10 w300 h24", "Heading")
    columnHelp := columnGui.AddText("x10 y38 w300 h24", "Help")
    columnDetailTitle := columnGui.AddText("x10 y66 w300 h24 Hidden", "Detail")
    columnDetailInfo := columnGui.AddText("x10 y94 w300 h24 Hidden", "Info")
    columnChoice := columnGui.AddButton("x10 y122 w100 h32", "Column")
    columnActions := []
    Loop 6
        columnActions.Push(columnGui.AddButton(
            "x10 y" (122 + A_Index * 36) " w100 h32 Hidden",
            "Action " A_Index
        ))
    columnBack := columnGui.AddButton("x10 y350 w100 h32 Hidden", "Back")
    columnReset := columnGui.AddButton("x120 y350 w100 h32", "Reset")
    columnNotice := columnGui.AddText("x230 y350 w180 h32 Hidden", "")
    columnModeEditor := Map(
        "gui", columnGui, "closed", false, "mode", "overview",
        "draft", columnDraft, "selectedIndex", 1,
        "columnButtons", [columnChoice], "actionButtons", columnActions,
        "controls", Map(
            "heading", columnHeading, "help", columnHelp,
            "detailTitle", columnDetailTitle, "detailInfo", columnDetailInfo,
            "back", columnBack, "resetAll", columnReset,
            "notice", columnNotice
        )
    )
    columnGui.Show("x20 y20 w440 h430")
    TestColumnRelayouts := 0
    StudyLibraryBigBoxColumnsShowMode(columnModeEditor, "detail", false)
    TestAssert(TestColumnRelayouts = 1 && columnDetailTitle.Visible
        && !columnHeading.Visible && columnActions[1].Visible,
        "Opening a fullscreen column detail re-runs its responsive layout")
    columnGui.Destroy()
    TestAssert(StudyReaderBigBoxPresentation(
        Map("bigBoxPresentation", true)
    ), "Fullscreen Reader presentation is explicit in its state")
    TestAssert(!StudyReaderBigBoxPresentation(Map()),
        "Desktop Reader state is not mistaken for fullscreen presentation")
    TestAssert(StudyCandidatesBigBoxPresentation(
        Map("bigBoxPresentation", true)
    ), "Fullscreen Review presentation is explicit in its state")
    TestAssert(!StudyCandidatesBigBoxPresentation(Map()),
        "Desktop Review state is not mistaken for fullscreen presentation")
    sentencePage := StudyCandidatesBigBoxPageData(1)
    vocabularyPage := StudyCandidatesBigBoxPageData(2)
    TestAssert(sentencePage["index"] = 1
        && sentencePage["title"] = "Sentences"
        && InStr(sentencePage["hint"], "1 / 2"),
        "Fullscreen Review visibly identifies the sentence page")
    TestAssert(vocabularyPage["index"] = 2
        && vocabularyPage["title"] = "Vocabulary"
        && InStr(vocabularyPage["hint"], "2 / 2"),
        "Fullscreen Review visibly identifies the vocabulary page")
    sentenceWidths := StudyCandidatesBigBoxColumnWidths(
        "sentences", 1200, 1.5
    )
    vocabularyWidths := StudyCandidatesBigBoxColumnWidths(
        "vocabulary", 1200, 1.5
    )
    TestAssert(sentenceWidths.Length = 5
        && sentenceWidths[4] >= 450,
        "Fullscreen sentence table gives its flexible width to Japanese source")
    TestAssert(vocabularyWidths.Length = 6
        && vocabularyWidths[3] >= 360,
        "Fullscreen vocabulary table gives its flexible width to meaning/context")
    sentenceTableWidths := StudyCandidatesBigBoxColumnWidths(
        "sentences", 1200, 1, true
    )
    vocabularyTableWidths := StudyCandidatesBigBoxColumnWidths(
        "vocabulary", 1200, 1, true
    )
    sentenceTableTotal := 0
    vocabularyTableTotal := 0
    for tableWidth in sentenceTableWidths
        sentenceTableTotal += tableWidth
    for tableWidth in vocabularyTableWidths
        vocabularyTableTotal += tableWidth
    TestAssert(sentenceTableTotal > 1200
        && vocabularyTableTotal > 1200,
        "Review table mode preserves readable columns for horizontal browsing")

    tableModeControls := StudyBigBoxTableModeCreate(
        libraryGui, Map("surface", "303030")
    )
    tableModeFooter := libraryGui.AddText(
        "x10 y275 w440 h20", "Normal Library footer"
    )
    CPStudyLibraryState["bigBoxPresentation"] := true
    CPStudyLibraryState["bigBoxLayoutScale"] := 1
    CPStudyLibraryState["bigBoxControls"] := Map("footer", tableModeFooter)
    CPStudyLibraryState["tableModeSurface"] := "library"
    CPStudyLibraryState["tableModeControls"] := tableModeControls
    CPStudyLibraryState["tableModeNormalControls"] := [
        libraryButton, libraryEdit, libraryBottomButton,
        tableModeControls["enter"]
    ]
    CPStudyLibraryState["tableModeActive"] := false
    TestAssert(StudyBigBoxTableModeSet(CPStudyLibraryState, true),
        "The shared Library table mode opens")
    TestAssert(StudyBigBoxTableModeActive(CPStudyLibraryState)
        && !libraryButton.Visible
        && tableModeControls["exit"].Visible,
        "Table mode hides the normal layout and exposes its dedicated actions")
    TestAssert(StudyBigBoxTableModeCurrentList(CPStudyLibraryState).Hwnd
        = libraryList.Hwnd
        && StudyBigBoxTableModeText(CPStudyLibraryState)["title"]
            = "Library table",
        "The shared table mode targets and identifies the Library table")
    libraryList.Modify(2, "Select Focus Vis")
    StudyBigBoxTableModeUpdate(CPStudyLibraryState)
    TestAssert(InStr(tableModeControls["status"].Text, "Row 2 of 3"),
        "Table mode reports the focused row")
    WinActivate("ahk_id " libraryGui.Hwnd)
    libraryList.Focus()
    TestAssert(StudyControllerDispatchNavigation(
        "Right", libraryGui.Hwnd
    ), "D-pad Left/Right is consumed as table column browsing")
    StudyControllerDispatchNavigation("Cancel", libraryGui.Hwnd)
    TestAssert(!StudyBigBoxTableModeActive(CPStudyLibraryState)
        && libraryButton.Visible
        && !tableModeControls["exit"].Visible
        && tableModeFooter.Text = "Normal Library footer",
        "B exits table mode and restores the normal Library layout")

    rowGui := Gui("+Owner" libraryGui.Hwnd, "Study row navigation fixture")
    rowUpperLeft := rowGui.AddButton("x100 y10 w100 h30", "Upper left")
    rowUpperRight := rowGui.AddButton("x310 y10 w100 h30", "Upper right")
    rowLeft := rowGui.AddButton("x10 y55 w100 h30", "Left")
    rowCenter := rowGui.AddButton("x160 y55 w100 h30", "Center")
    rowRight := rowGui.AddButton("x310 y55 w100 h30", "Right")
    rowGui.Show("x620 y160 w430 h105")
    WinActivate("ahk_id " rowGui.Hwnd)
    rowCenter.Focus()
    StudyControllerDispatchNavigation("Left", rowGui.Hwnd)
    TestAssert(StudyControllerFocusedHwnd(rowGui.Hwnd) = rowLeft.Hwnd,
        "D-pad Left stays in its visual row instead of choosing a diagonal control")
    rowCenter.Focus()
    StudyControllerDispatchNavigation("Right", rowGui.Hwnd)
    TestAssert(StudyControllerFocusedHwnd(rowGui.Hwnd) = rowRight.Hwnd,
        "D-pad Right stays in its visual row instead of choosing a diagonal control")
    StudyControllerDispatchNavigation("Right", rowGui.Hwnd)
    TestAssert(StudyControllerFocusedHwnd(rowGui.Hwnd) = rowRight.Hwnd,
        "D-pad Right stops at the end of a visual row")
    rowGui.Destroy()

    popupGui := Gui(
        "+Owner" libraryGui.Hwnd " +ToolWindow -Caption +Border",
        "Study choice menu fixture"
    )
    popupRows := [
        popupGui.AddText("x3 y3 w180 h30 +Tabstop", "First choice"),
        popupGui.AddText("x3 y34 w180 h30 +Tabstop", "Second choice"),
        popupGui.AddText("x3 y65 w180 h30 +Tabstop", "Third choice")
    ]
    popupState := Map(
        "closed", false, "result", 0, "gui", popupGui,
        "colors", Map(
            "focus", "264F73",
            "surface", "303030",
            "text", "ECEDEF",
            "accentFocus", "0F5F9E",
            "accentText", "FFFFFF"
        ),
        "rows", popupRows, "focusIndex", 1
    )
    popupGui.Show("x420 y340 w186 h100")
    popupHwnd := popupGui.Hwnd
    CPThemedChoicePopupRegistry()[popupHwnd] := popupState
    CPThemedChoicePopupResetRows(popupState)
    CPThemedChoicePopupFocus(popupState, 1)
    WinActivate("ahk_id " popupGui.Hwnd)
    Sleep(30)
    TestAssert(CPControllerNavigationTarget(popupGui.Hwnd) = popupGui.Hwnd,
        "Controller polling recognizes a Study choice menu")
    StudyControllerDispatchNavigation("Down", popupGui.Hwnd)
    TestAssert(popupState["focusIndex"] = 2,
        "D-pad Down moves directly through a Study choice menu")
    StudyControllerDispatchNavigation("Activate", popupGui.Hwnd)
    TestAssert(popupState["closed"] && popupState["result"] = 2,
        "A selects the highlighted Study choice-menu entry")
    CPThemedChoicePopupRegistry().Delete(popupHwnd)

    ankiGui := Gui("+Owner" libraryGui.Hwnd, "Anki lifecycle fixture")
    ankiStatus := ankiGui.AddText("x10 y10 w220 h30", "Checking...")
    ankiDeck := ankiGui.AddDropDownList("x10 y45 w180", [])
    ankiModel := ankiGui.AddDropDownList("x10 y80 w180", [])
    ankiGui.Show("x620 y340 w250 h130")
    ankiState := Map(
        "gui", ankiGui,
        "libraryState", Map("outputDir", A_Temp),
        "status", ankiStatus,
        "deckDdl", ankiDeck,
        "modelDdl", ankiModel,
        "closed", false,
        "decks", [], "models", [], "modelFields", Map()
    )
    TestAnkiStateToClose := ankiState
    TestAnkiCloseDuringBridge := true
    StudyAnkiDiscover(ankiState)
    TestAnkiCloseDuringBridge := false
    TestAssert(ankiState["closed"] && !StudyAnkiDialogAlive(ankiState),
        "Closing the Anki dialog during discovery safely discards its result")

    hiddenThemeGui := Gui(, "Study hidden theme fixture")
    hiddenThemeGui.AddText("x10 y10 w90 h20", "Hidden")
    hiddenThemeGui.Show("Hide w120 h80")
    TestAssert(CPApplyOwnedDialogTheme(hiddenThemeGui),
        "The dialog theme pass supports a valid hidden Study window")
    hiddenThemeGui.Destroy()

    destroyedThemeGui := Gui(, "Destroyed theme fixture")
    destroyedThemeGui.Show("Hide w120 h80")
    destroyedThemeGui.Destroy()
    TestAssert(!CPApplyOwnedDialogTheme(destroyedThemeGui),
        "The dialog theme pass safely ignores an already destroyed GUI")

    initialRefreshGui := Gui(, "Initial candidate refresh fixture")
    initialScope := initialRefreshGui.AddDropDownList(
        "x10 y10 w150", ["New", "All", "Ignored"]
    )
    initialScope.Choose(1)
    initialRefreshButton := initialRefreshGui.AddButton(
        "x170 y10 w80 h30", "Refresh"
    )
    initialStatus := initialRefreshGui.AddText(
        "x10 y48 w240 h25", "Waiting"
    )
    initialTabs := initialRefreshGui.AddTab3(
        "x10 y80 w240 h100", ["Sentences", "Vocabulary"]
    )
    initialTabs.UseTab(1)
    initialSentenceList := initialRefreshGui.AddListView(
        "x20 y110 w210 h45", ["Sentence"]
    )
    initialTabs.UseTab(2)
    initialVocabularyList := initialRefreshGui.AddListView(
        "x20 y110 w210 h45", ["Vocabulary"]
    )
    initialTabs.UseTab()
    initialRefreshGui.Show("Hide w270 h195")
    initialRefreshState := Map(
        "gui", initialRefreshGui,
        "scopeDdl", initialScope,
        "scope", "new",
        "refreshButton", initialRefreshButton,
        "status", initialStatus,
        "tabs", initialTabs,
        "sentenceList", initialSentenceList,
        "vocabularyList", initialVocabularyList,
        "closeRequested", false,
        "initialRefreshCallback", 0
    )
    StudyCandidatesInitialRefresh(initialRefreshState)
    TestAssert(TestCandidateAsyncStarts = 1
        && IsObject(TestCandidateAsyncCallback),
        "The initial candidate snapshot starts asynchronously")
    TestAssert(!initialRefreshButton.Enabled
        && initialStatus.Text = "Building local candidate lists...",
        "The Review window remains responsive with a visible loading state")
    TestCandidateAsyncCallback.Call(0)
    TestAssert(TestCandidateSnapshotsRead = 1 && initialRefreshButton.Enabled,
        "A completed initial snapshot populates Review and restores Refresh")
    initialRefreshState["closeRequested"] := true
    initialRefreshGui.Destroy()
    TestCandidateAsyncCallback.Call(0)
    TestAssert(TestCandidateSnapshotsRead = 1,
        "Closing during the initial candidate snapshot discards its later UI result")

    ownedGui := Gui("+Owner" libraryGui.Hwnd, "Owned Study dialog fixture")
    ownedDdl := ownedGui.AddDropDownList(
        "x10 y10 w150", ["First", "Second", "Third"]
    )
    ownedDdl.Choose(1)
    ownedButton := ownedGui.AddButton("x10 y50 w100 h30", "Cancel")
    ownedGui.Show("x120 y120 w180 h100")
    WinActivate("ahk_id " ownedGui.Hwnd)
    ownedDdl.Focus()
    Sleep(30)
    TestAssert(StudyControllerSurfaceForWindow(
        ownedGui.Hwnd, &ownedSurface
    ) && ownedSurface["kind"] = "library",
        "Owned Study dialogs inherit controller navigation")
    TestAssert(!StudyControllerSurfaceIsRoot(ownedGui.Hwnd),
        "Owned dialog is not mistaken for its Library root")
    TestAssert(CPControllerNavigationTarget(ownedGui.Hwnd) = ownedGui.Hwnd,
        "Controller polling recognizes Library-owned dialogs")
    StudyControllerDispatchNavigation("Activate", ownedGui.Hwnd)
    TestAssert(CPComboDropped(ownedDdl.Hwnd),
        "A opens a native Study dropdown list")
    StudyControllerDispatchNavigation("Down", ownedGui.Hwnd)
    StudyControllerDispatchNavigation("Activate", ownedGui.Hwnd)
    TestAssert(!CPComboDropped(ownedDdl.Hwnd) && ownedDdl.Value = 2,
        "D-pad and A navigate and select a native Study dropdown entry")
    ownedButton.Focus()
    StudyControllerDispatchNavigation("Cancel", ownedGui.Hwnd)
    TestAssert(TestSyntheticKeys.Length
        && TestSyntheticKeys[TestSyntheticKeys.Length] = "Esc",
        "B sends Escape to an owned Study dialog")
    ownedGui.Destroy()

    readerGui := Gui("+AlwaysOnTop", "Study controller reader fixture")
    readerEdit := readerGui.AddEdit(
        "x10 y10 w330 h100 ReadOnly Multi VScroll",
        TestMultilineText()
    )
    readerGui.Show("x600 y80 w360 h140")
    CPStudyReaderState := Map(
        "gui", readerGui,
        "explanation", readerEdit,
        "entryIds", [1, 2],
        "closed", false
    )
    TestAssert(StudyControllerSurfaceForWindow(
        readerGui.Hwnd, &readerSurface
    ) && readerSurface["kind"] = "reader",
        "Study Reader is a controller navigation root")
    StudyControllerDispatchNavigation("PreviousTab", readerGui.Hwnd)
    StudyControllerDispatchNavigation("NextTab", readerGui.Hwnd)
    TestAssert(TestReaderSteps.Length = 2
        && TestReaderSteps[1] = -1 && TestReaderSteps[2] = 1,
        "Reader shoulders browse previous and next explanations")

    candidateGui := Gui("+AlwaysOnTop", "Study controller candidate fixture")
    candidateScope := candidateGui.AddDropDownList(
        "x10 y10 w170", ["New", "All", "Ignored"]
    )
    candidateScope.Choose(1)
    candidateAi := candidateGui.AddDropDownList(
        "x200 y10 w170", ["All statuses", "Recommended", "Unassessed"]
    )
    candidateAi.Choose(1)
    candidateScope.OnEvent("Change", TestCandidateScopeChanged)
    candidateAi.OnEvent("Change", TestCandidateAiChanged)
    candidateTabs := candidateGui.AddTab3(
        "x10 y50 w430 h200", ["Sentences", "Vocabulary"]
    )
    candidateTabs.UseTab(1)
    sentenceList := candidateGui.AddListView(
        "x20 y85 w400 h130", ["Sentence"]
    )
    sentenceList.Add(, "Sentence one")
    sentenceList.Add(, "Sentence two")
    candidateTabs.UseTab(2)
    vocabularyList := candidateGui.AddListView(
        "x20 y85 w400 h130", ["Vocabulary"]
    )
    vocabularyList.Add(, "Word one")
    vocabularyList.Add(, "Word two")
    candidateTabs.UseTab()
    candidateGui.Show("x80 y360 w450 h260")
    candidateTableModeControls := StudyBigBoxTableModeCreate(
        candidateGui, Map("surface", "303030")
    )
    CPStudyCandidateState := Map(
        "gui", candidateGui,
        "scopeDdl", candidateScope,
        "aiFilterDdl", candidateAi,
        "tabs", candidateTabs,
        "sentenceList", sentenceList,
        "vocabularyList", vocabularyList,
        "closeRequested", false,
        "bigBoxPresentation", true,
        "tableModeSurface", "candidates",
        "tableModeControls", candidateTableModeControls,
        "tableModeActive", true
    )
    TestAssert(StudyBigBoxTableModeCurrentList(CPStudyCandidateState).Hwnd
        = sentenceList.Hwnd
        && StudyBigBoxTableModeText(CPStudyCandidateState)["title"]
            = "Sentences table",
        "Review table mode starts on and visibly identifies Sentences")
    candidateTabs.Choose(2)
    StudyBigBoxTableModeUpdate(CPStudyCandidateState)
    TestAssert(StudyBigBoxTableModeCurrentList(CPStudyCandidateState).Hwnd
        = vocabularyList.Hwnd
        && candidateTableModeControls["title"].Text = "Vocabulary table",
        "Review table mode shares its shell with the Vocabulary table")
    CPStudyCandidateState["tableModeActive"] := false
    candidateTabs.Choose(1)
    WinActivate("ahk_id " candidateGui.Hwnd)
    candidateScope.Focus()
    StudyControllerDispatchNavigation("Activate", candidateGui.Hwnd)
    TestAssert(CPComboDropped(candidateScope.Hwnd)
        && StudyControllerComboPreviewActive(candidateScope.Hwnd),
        "A opens Review scope as a pending controller selection")
    StudyControllerDispatchNavigation("Down", candidateGui.Hwnd)
    Sleep(100)
    TestAssert(candidateScope.Value = 2
        && TestCandidateScopeChanges = 0
        && CPComboDropped(candidateScope.Hwnd),
        "D-pad previews Review scope without applying or closing it")
    StudyControllerDispatchNavigation("Activate", candidateGui.Hwnd)
    TestAssert(candidateScope.Value = 2,
        "A keeps the highlighted Review scope while committing")
    TestAssert(TestCandidateScopeChanges = 1,
        "A applies the highlighted Review scope exactly once")
    TestAssert(!CPComboDropped(candidateScope.Hwnd),
        "A closes the Review scope dropdown after committing")
    TestAssert(!StudyControllerComboPreviewActive(candidateScope.Hwnd),
        "A clears the pending Review scope transaction")

    candidateAi.Focus()
    StudyControllerDispatchNavigation("Activate", candidateGui.Hwnd)
    StudyControllerDispatchNavigation("Down", candidateGui.Hwnd)
    TestAssert(candidateAi.Value = 2 && TestCandidateAiChanges = 0,
        "D-pad previews the Review AI filter without applying it")
    StudyControllerDispatchNavigation("Cancel", candidateGui.Hwnd)
    TestAssert(candidateAi.Value = 1
        && TestCandidateAiChanges = 0
        && !CPComboDropped(candidateAi.Hwnd)
        && !StudyControllerComboPreviewActive(candidateAi.Hwnd),
        "B restores the previous Review AI filter without applying it")

    candidateTabs.Choose(1)
    sentenceList.Modify(1, "Select Focus Vis")
    vocabularyList.Modify(1, "Select -Focus Vis")
    StudyCandidatesTabChanged(CPStudyCandidateState)
    StudyControllerDispatchNavigation("NextTab", candidateGui.Hwnd)
    TestAssert(candidateTabs.Value = 2
        && StudyControllerFocusedHwnd(candidateGui.Hwnd) = vocabularyList.Hwnd
        && vocabularyList.GetNext(0, "F") = 1,
        "RB changes Review for Anki to Vocabulary and focuses its selected row")
    StudyControllerDispatchNavigation("Down", candidateGui.Hwnd)
    TestAssert(vocabularyList.GetNext(0, "F") = 2,
        "The first D-pad Down after switching advances to the next vocabulary")
    vocabularyList.Modify(1, "Select Focus Vis")
    StudyControllerDispatchNavigation("Activate", candidateGui.Hwnd)
    Sleep(60)
    TestAssert(TestCandidateActions = 1 && TestCandidateOpens = 0
        && TestCandidateActionRow = 1 && TestCandidateActionPrefersAnki,
        "A opens row actions for the selected fullscreen Anki candidate")
    CPStudyCandidateState["bigBoxPresentation"] := false
    StudyControllerDispatchNavigation("Activate", candidateGui.Hwnd)
    TestAssert(TestCandidateOpens = 1,
        "Desktop Review keeps its direct open action")
    CPStudyCandidateState["bigBoxPresentation"] := true
    StudyControllerDispatchNavigation("PreviousTab", candidateGui.Hwnd)
    TestAssert(candidateTabs.Value = 1
        && StudyControllerFocusedHwnd(candidateGui.Hwnd) = sentenceList.Hwnd,
        "LB changes Review for Anki back to Sentences")

    CPStudyLibraryState["bigBoxPresentation"] := true
    TestAssert(StudyBigBoxFocusFrameStart(CPStudyLibraryState)
        && CPStudyLibraryState["bigBoxFocusFrame"].Length = 4,
        "Fullscreen Study surfaces start dashboard-style focus tracking")
    ; Exercise the independent focus watcher itself. The production fix must
    ; not depend on GuiControl Focus events, which can be missing after a GUI's
    ; hidden fullscreen composition pass on some Windows configurations.
    for studyBoundHwnd, studyBoundCallback
        in CPStudyLibraryState["bigBoxFocusBindings"] {
        try GuiCtrlFromHwnd(studyBoundHwnd).OnEvent(
            "Focus", studyBoundCallback, 0
        )
    }
    CPStudyLibraryState["bigBoxFocusBindings"] := Map()
    candidateGui.Hide()
    readerGui.Hide()
    WinActivate("ahk_id " libraryGui.Hwnd)
    libraryList.Focus()
    StudyBigBoxFocusFrameHide(CPStudyLibraryState)
    libraryButton.Focus()
    Sleep(100)
    studyVisibleFrameParts := 0
    studyBlueFrameParts := 0
    studyFramePixels := ""
    for studyFramePart in CPStudyLibraryState["bigBoxFocusFrame"] {
        if DllCall(
            "user32\IsWindowVisible", "ptr", studyFramePart.Hwnd, "int"
        ) {
            studyVisibleFrameParts += 1
            studyFrameDc := DllCall(
                "user32\GetDC", "ptr", studyFramePart.Hwnd, "ptr"
            )
            if studyFrameDc {
                studyFramePart.GetPos(,, &studyFrameW, &studyFrameH)
                studyFramePixel := DllCall(
                    "gdi32\GetPixel", "ptr", studyFrameDc,
                    "int", Max(0, Floor(studyFrameW / 2)),
                    "int", Max(0, Floor(studyFrameH / 2)), "uint"
                )
                DllCall(
                    "user32\ReleaseDC", "ptr", studyFramePart.Hwnd,
                    "ptr", studyFrameDc
                )
                studyFramePixels .= (studyFramePixels = "" ? "" : ", ")
                    . Format("0x{:06X}", studyFramePixel & 0xFFFFFF)
                ; COLORREF stores RGB(0x62, 0xC7, 0xFF) as 0x00FFC762.
                if ((studyFramePixel & 0xFFFFFF) = 0xFFC762)
                    studyBlueFrameParts += 1
            }
        }
    }
    TestAssert(studyVisibleFrameParts = 4
        && CPStudyLibraryState["bigBoxFocusTarget"] = libraryButton.Hwnd,
        "Fullscreen Study focus draws a visible blue frame around the active control"
            . " (visible=" studyVisibleFrameParts
            . ", target=" CPStudyLibraryState["bigBoxFocusTarget"]
            . ", expected=" libraryButton.Hwnd
            . ", foreground=" DllCall("user32\GetForegroundWindow", "ptr")
            . ", root=" libraryGui.Hwnd . ")")
    TestAssert(studyBlueFrameParts = 4,
        "Fullscreen Study focus frame paints the light-blue selection color ("
            . studyFramePixels . ")")
    StudyBigBoxFocusFrameStop(CPStudyLibraryState)
    CPStudyLibraryState["bigBoxPresentation"] := false

    StudyControllerDispatchNavigation("Cancel", candidateGui.Hwnd)
    StudyControllerDispatchNavigation("Cancel", readerGui.Hwnd)
    WinActivate("ahk_id " libraryGui.Hwnd)
    StudyControllerDispatchNavigation("Cancel", libraryGui.Hwnd)
    TestAssert(TestCandidateCloses = 1 && TestReaderCloses = 1
        && TestLibraryCloses = 1,
        "B uses each Study surface's normal close path")

    try candidateGui.Destroy()
    try readerGui.Destroy()
    try libraryGui.Destroy()
    FileAppend("PASS: " TestCount " Study controller assertions.`n", "*", "UTF-8")
    ExitApp(0)
} catch as testFailure {
    FileAppend("FAIL: " testFailure.Message "`n" testFailure.Stack "`n", "*", "UTF-8")
    ExitApp(1)
}

TestAssert(condition, message) {
    global TestCount
    if !condition
        throw Error(message)
    TestCount += 1
    FileAppend("OK: " message "`n", "*", "UTF-8")
}

TestMultilineText() {
    text := ""
    Loop 20
        text .= (A_Index > 1 ? "`r`n" : "") "Line " A_Index
    return text
}

TestRecordButtonClick(*) {
    global TestButtonClicks
    TestButtonClicks += 1
}

TestCandidateScopeChanged(*) {
    global TestCandidateScopeChanges
    if StudyControllerComboPreviewActive(CPStudyCandidateState["scopeDdl"].Hwnd)
        return
    TestCandidateScopeChanges += 1
}

TestCandidateAiChanged(*) {
    global TestCandidateAiChanges
    if StudyControllerComboPreviewActive(CPStudyCandidateState["aiFilterDdl"].Hwnd)
        return
    TestCandidateAiChanges += 1
}

StudyCandidatesScopeChanged(*) {
    TestCandidateScopeChanged()
}

StudyCandidatesAiFilterChanged(*) {
    TestCandidateAiChanged()
}

StudyCandidatesRefresh(scState, *) {
}

StudyCandidatesStartAsyncBridge(scState, scAction, scCallback, *) {
    global TestCandidateAsyncStarts, TestCandidateAsyncCallback
    TestCandidateAsyncStarts += 1
    TestCandidateAsyncCallback := scCallback
    return true
}

StudyCandidatesReadSnapshot(*) {
    global TestCandidateSnapshotsRead
    TestCandidateSnapshotsRead += 1
    return true
}

CPIsColorSwatchControl(*) => false
CPPalette(*) => Map("window", "202124")
CPApplyDarkTitleBar(*) {
}
CPSetPreferredAppDarkMode(*) {
}
CPAllowDarkModeForWindow(*) {
}
CPApplyWindowScrollbarTheme(*) {
}
CPRefreshThemeBrushes(*) {
}
CPPrepareStudyCombo(*) {
}
CPPrepareStudyListHeader(*) {
}
CPApplyThemeToControl(*) {
}
CPThemedOwnedMessage(*) {
}
CPListViewSaveColumnOrder(*) {
}

CPFocusRingTargetHwnd(hwnd) {
    try {
        parent := DllCall("user32\GetParent", "ptr", hwnd, "ptr")
        if (parent && InStr(WinGetClass("ahk_id " parent), "ComboBox"))
            return parent
    }
    return hwnd
}

CPControllerSendDialogKey(keyName) {
    global TestSyntheticKeys
    TestSyntheticKeys.Push(keyName)
}

CPControllerResetNavigation(*) {
}

CPThemedChoicePopupRegistry() {
    static registry := Map()
    return registry
}

StudyAnkiRunBridge(*) {
    global TestAnkiStateToClose, TestAnkiCloseDuringBridge
    if TestAnkiCloseDuringBridge
        StudyAnkiCloseDialog(TestAnkiStateToClose)
    return true
}

StudyAnkiReadStatus(*) => Map("code", "connected", "version", "6", "message", "")
StudyLibraryReadRows(*) => []
StudyLibraryHexDecode(value) => value
StudyAnkiUpdateConnectionText(*) => true
StudyAnkiApplyProfileMapping(*) {
}

CPControllerKeyboardMirrorActive(*) => false
CPBigBoxDashboardAlive(*) => false
CPBigBoxDashboardDpiScale(*) => 1
CPBigBoxFocusColor(*) => "62C7FF"
CPRegisterColorSwatch(control, *) => control

StudyCandidatesRecommendationBigBoxResize(*) {
    global TestColumnRelayouts
    TestColumnRelayouts += 1
}

StudyCandidatesResizeBigBox(*) {
}

StudyLibraryResizeBigBox(*) {
}

StudyLibraryShowImage(*) {
}

StudyLibraryStateAlive(studyState) {
    if !IsObject(studyState) || !studyState.Has("gui")
        return false
    if studyState.Get("closed", false)
        return false
    try return DllCall(
        "user32\IsWindow", "ptr", studyState["gui"].Hwnd, "int"
    ) != 0
    return false
}

StudyCandidatesGuiAlive(studyState) {
    if !IsObject(studyState) || !studyState.Has("gui")
        return false
    try return DllCall(
        "user32\IsWindow", "ptr", studyState["gui"].Hwnd, "int"
    ) != 0
    return false
}

StudyLibraryOpenSelectedReader(*) {
    global TestLibraryOpens
    TestLibraryOpens += 1
}

StudyLibraryRowGroupId(studyState, row) {
    return studyState["groups"][row]["id"]
}

StudyLibraryLoadGroup(studyState, groupId, *) {
    global TestLibraryGroupLoads
    TestLibraryGroupLoads += 1
    studyState["currentGroupId"] := groupId
}

StudyLibrarySwitchTo(studyState, libraryName, *) {
    global TestLibraryChanges
    TestLibraryChanges += 1
    studyState["libraryName"] := libraryName
    return true
}

StudyLibraryUpdateSelectionActions(*) {
}

StudyCandidatesOpenSelected(*) {
    global TestCandidateOpens
    TestCandidateOpens += 1
}

StudyCandidatesContextMenu(
    scState, scGui, scControl, scRow, scIsRightClick, scX, scY,
    scPreferAddToAnki := false
) {
    global TestCandidateActions, TestCandidateActionRow
    global TestCandidateActionPrefersAnki
    TestCandidateActions += 1
    TestCandidateActionRow := scRow
    TestCandidateActionPrefersAnki := scPreferAddToAnki
}

StudyLibraryBigBoxColumnsRefresh(*) {
}

StudyReaderStepEntry(studyState, direction, *) {
    global TestReaderSteps
    TestReaderSteps.Push(direction)
}

StudyCandidatesUpdateActions(*) {
}

StudyCandidatesUpdateBigBoxPageNavigation(*) {
}

StudyLibraryClose(*) {
    global TestLibraryCloses
    TestLibraryCloses += 1
}

StudyReaderClose(*) {
    global TestReaderCloses
    TestReaderCloses += 1
}

StudyCandidatesClose(*) {
    global TestCandidateCloses
    TestCandidateCloses += 1
}

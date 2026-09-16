#Requires AutoHotkey v2.0
; Fullscreen fixtures must never enter the separate desktop paint path.
StudyDesktopTheme(*) {
    throw Error("Fullscreen Study unexpectedly requested desktop styling")
}
; This suite exercises fullscreen and classic fallbacks. Modern desktop
; constructors, rendering and event wiring run in test_desktop_layout.ps1.
CPDesktopActive() {
    return false
}
StudyDesktopAnkiCreate(*) {
}
StudyDesktopMetadataCreate(*) {
    return 0
}
StudyDesktopManagementCreate(*) {
}
StudyDesktopRecommendationCreate(*) {
}
StudyDesktopRecommendationPreviewText(*) {
    throw Error("Fullscreen/classic Study unexpectedly requested desktop prompt text")
}
StudyDesktopStorageCreate(*) {
    return 0
}
StudyDesktopManagementColumns(*) {
    throw Error("Fullscreen/classic Study unexpectedly requested desktop table sizing")
}
StudyDesktopOpenLibraryName(*) {
    throw Error("Fullscreen/classic Study unexpectedly requested a desktop name dialog")
}
StudyDesktopOpenDatePicker(*) {
    throw Error("Fullscreen/classic Study unexpectedly requested a desktop date picker")
}
StudyDesktopDialogShow(*) {
    throw Error("Fullscreen/classic Study unexpectedly requested a desktop dialog")
}
StudyDesktopMessage(*) {
    throw Error("Fullscreen/classic Study unexpectedly requested a desktop message")
}
#SingleInstance Off
#Warn All, StdOut
#NoTrayIcon

global CPStudyLibraryState := 0
global CPStudyReaderState := 0
global CPStudyCandidateState := 0
global __CP_STUDY_NAV_ITEMS := []
global CPStudySyntheticKeyDepth := 0
global CPStudyComboTransactions := Map()
global CPComboSeparatorBefore := Map()
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
global TestLibraryTableFocusEvents := 0
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
global TestFilterRefreshes := 0
global TestFilterWarnings := 0
global TestRecommendationState := 0
global TestRecommendationSaved := 0
global TestRecommendationCancelOnly := false
global TestRecommendationPreviewCancel := false
global TestVocabularyRows := [], TestVocabularyPreviews := [], TestAnkiMenuChoices := []
global TestDesktopVocabularySelections := 0
global TestVocabularyUseRealReview := false, TestVocabularyReviewState := 0
global TestAnkiCancelReturns := 0
global CPControllerSurfaceTransitionState := ""
global TestHandoffPage := 0, TestHandoffChecks := 0, TestHandoffCancelDiscovery := false
global TestHandoffFailDiscovery := false
global TestChapterReply := "No", TestChapterMessageChecks := 0, TestDesktopMessages := []
global TestDesktopMessageReplies := []
global TestChapterSelectorFocus := 0
global studyLibrariesRoot := A_ScriptDir "\management-fixture\libraries"
global studyLibrariesArchiveRoot := A_ScriptDir "\management-fixture\archives"
global studyLibraryDefaultDir := A_ScriptDir "\management-fixture\default"
global TestManagementReply := "No", TestManagementMessages := 0
global studyLibraryDir := A_ScriptDir "\chapter-workflow-fixture"
global TestAnkiPreviewState := 0, TestAnkiMessageState := 0, TestAnkiReplies := []
global TestAnkiSends := 0, TestAnkiWrittenFiles := Map(), TestScreenshotSaves := []
global TestAnkiExampleRequests := []
global iniPath := A_ScriptDir "\recommendation-settings-fixture\control.ini"
; Modal workflow tests use timers. Fail cleanly instead of leaving an error
; dialog waiting behind a fullscreen fixture when a timer assertion fails.
OnError(TestHarnessError)

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

    ; Exercise real horizontal scroll ranges in both desktop and fullscreen
    ; presentation states, including a final partial step and the next exit.
    horizontalLeftButton := libraryGui.AddButton(
        "x10 y10 w120 h35 Hidden", "Left action")
    for horizontalBigBox in [false, true] {
        CPStudyLibraryState["bigBoxPresentation"] := horizontalBigBox
        CPStudyLibraryState["bigBoxLayoutScale"] := horizontalBigBox ? 1.5 : 1
        libraryList.ModifyCol(1, 1050)
        libraryList.Modify(2, "Select Focus Vis")
        libraryList.Focus()
        SendMessage(0x1014, -10000, 0, libraryList.Hwnd)
        StudyControllerDispatchNavigation("Right", libraryGui.Hwnd)
        horizontalAfterRight := DllCall("user32\GetScrollPos",
            "ptr", libraryList.Hwnd, "int", 0, "int")
        TestAssert(horizontalAfterRight > 0
            && StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryList.Hwnd
            && libraryList.GetNext(0, "F") = 2,
            "Right scrolls overflowing columns without leaving the table or changing row")
        StudyControllerDispatchNavigation("Left", libraryGui.Hwnd)
        TestAssert(DllCall("user32\GetScrollPos", "ptr", libraryList.Hwnd,
            "int", 0, "int") < horizontalAfterRight
            && StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryList.Hwnd,
            "Left scrolls overflowing columns back while retaining table focus")
        SendMessage(0x1014, 10000, 0, libraryList.Hwnd)
        horizontalMax := DllCall("user32\GetScrollPos",
            "ptr", libraryList.Hwnd, "int", 0, "int")
        SendMessage(0x1014, -10, 0, libraryList.Hwnd)
        StudyControllerDispatchNavigation("Right", libraryGui.Hwnd)
        TestAssert(DllCall("user32\GetScrollPos", "ptr", libraryList.Hwnd,
            "int", 0, "int") = horizontalMax
            && StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryList.Hwnd,
            "The final short horizontal step reaches the edge without moving focus")
        StudyControllerDispatchNavigation("Right", libraryGui.Hwnd)
        TestAssert(StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryButton.Hwnd,
            "A subsequent Right at the scroll boundary exits to the right-hand action")
        libraryList.Move(140, 10, 170, 150)
        horizontalLeftButton.Visible := true
        libraryList.Focus()
        SendMessage(0x1014, -10000, 0, libraryList.Hwnd)
        SendMessage(0x1014, 10, 0, libraryList.Hwnd)
        StudyControllerDispatchNavigation("Left", libraryGui.Hwnd)
        TestAssert(DllCall("user32\GetScrollPos", "ptr", libraryList.Hwnd,
            "int", 0, "int") = 0
            && StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryList.Hwnd,
            "The final Left scroll step retains table focus at the left edge")
        StudyControllerDispatchNavigation("Left", libraryGui.Hwnd)
        TestAssert(StudyControllerFocusedHwnd(libraryGui.Hwnd) = horizontalLeftButton.Hwnd,
            "A subsequent Left at the scroll boundary exits to the left-hand action")
        horizontalLeftButton.Visible := false
        libraryList.Move(10, 10, 300, 150)
        libraryList.Focus()
        libraryList.Modify(3, "Select Focus")
        StudyControllerDispatchNavigation("Down", libraryGui.Hwnd)
        TestAssert(StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryEdit.Hwnd,
            "Horizontal overflow still permits Down to exit after the final row")
        libraryList.ModifyCol(1, 180)
        libraryList.Focus()
        StudyControllerDispatchNavigation("Right", libraryGui.Hwnd)
        TestAssert(StudyControllerFocusedHwnd(libraryGui.Hwnd) = libraryButton.Hwnd,
            "A table whose columns fit exits right immediately")
    }
    CPStudyLibraryState["bigBoxPresentation"] := false

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

    TestDateFilters(CPStudyLibraryState)
    TestChapterWorkflow()
    TestManagementWorkflow()
    TestLibraryOpeningAndImageNavigation()
    TestFilterHeaderMarkers()
    TestVersionDisplay()
    TestRecommendationPersistence()
    TestRecommendationWorkflow(CPStudyLibraryState)
    TestReaderVocabularyPicker(CPStudyLibraryState)
    TestCandidateAnkiOwner()
    TestFullscreenAnkiPreview(CPStudyLibraryState)
    WinActivate("ahk_id " libraryGui.Hwnd)
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
    bigBoxColumnState["columns"][2]["width"] := 900
    wideColumnWidths := StudyLibraryBigBoxColumnWidths(
        bigBoxColumnState, 800, 1
    )
    TestAssert(wideColumnWidths[1] = 100 && wideColumnWidths[2] = 900
        && wideColumnWidths[3] = 0,
        "Wide saved Library columns remain available for direct horizontal scrolling")
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
    CPStudyLibraryState["bigBoxPresentation"] := true
    CPStudyLibraryState["bigBoxLayoutScale"] := 1

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
    candidateWantBigBox := true
    ; @CANDIDATE_TAB_SETUP@
    candidateTabs.Move(10, 50, 430, 200)
    candidateTabs.UseTab(1)
    sentenceList := candidateGui.AddListView(
        "x20 y90 w400 h150", ["Sentence"]
    )
    sentenceList.Add(, "Sentence one")
    sentenceList.Add(, "Sentence two")
    candidateTabs.UseTab(2)
    vocabularyList := candidateGui.AddListView(
        "x20 y90 w400 h150", ["Vocabulary"]
    )
    vocabularyList.Add(, "Word one")
    vocabularyList.Add(, "Word two")
    candidateTabs.UseTab()
    candidateGui.Show("x80 y360 w450 h260")
    CPStudyCandidateState := Map(
        "gui", candidateGui,
        "scopeDdl", candidateScope,
        "aiFilterDdl", candidateAi,
        "tabs", candidateTabs,
        "sentenceList", sentenceList,
        "vocabularyList", vocabularyList,
        "closeRequested", false,
        "bigBoxPresentation", true
    )
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

    for reviewPage in [1, 2] {
        candidateTabs.Choose(reviewPage)
        reviewList := reviewPage = 1 ? sentenceList : vocabularyList
        reviewList.Modify(1, "Select Focus Vis")
        StudyCandidatesTabChanged(CPStudyCandidateState)
        StudyControllerDispatchNavigation("Up", candidateGui.Hwnd)
        TestAssert(StudyControllerFocusedHwnd(candidateGui.Hwnd) = candidateAi.Hwnd,
            "One Up from the first Review row exits directly above the table: page " reviewPage)
        TestAssert(reviewList.GetNext() = 1,
            "Leaving the Review table preserves the selected candidate")
        StudyControllerDispatchNavigation("Down", candidateGui.Hwnd)
        TestAssert(StudyControllerFocusedHwnd(candidateGui.Hwnd) = reviewList.Hwnd,
            "Down returns directly to the visible Review list, not its container")
        StudyControllerDispatchNavigation("Down", candidateGui.Hwnd)
        TestAssert(reviewList.GetNext(0, "F") = 2,
            "The next Down advances a row after reentering the Review table")
        reviewList.Delete()
        StudyControllerDispatchNavigation("Up", candidateGui.Hwnd)
        TestAssert(StudyControllerFocusedHwnd(candidateGui.Hwnd) = candidateAi.Hwnd,
            "An empty Review table also exits upward without a container stop")
    }
    TestAssert(!CPHwndIsFocusable(candidateTabs.Hwnd),
        "Fullscreen Review's structural tabs are not a keyboard/controller focus stop")

    ; The same production setup must retain the desktop's visible tab stop.
    candidateWantBigBox := false
    candidateGui := Gui("+ToolWindow", "Desktop Review tab fixture")
    ; @CANDIDATE_TAB_SETUP@
    candidateGui.Show("x80 y80 w850 h520")
    TestAssert(CPHwndIsFocusable(candidateTabs.Hwnd),
        "Desktop Review retains its navigable native tabs")
    candidateGui.Destroy()
    candidateGui := CPStudyCandidateState["gui"]

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
    ; Remove the date-picker test hooks before AHK tears down static registries.
    OnMessage(0x0082, StudyLibraryDatePickerDestroyed, 0)
    OnMessage(0x0100, StudyLibraryDatePickerKeyDown, 0)
    FileAppend("PASS: " TestCount " Study controller assertions.`n", "*", "UTF-8")
    ExitApp(0)
} catch as testFailure {
    FileAppend("FAIL: " testFailure.Message "`n" testFailure.Stack "`n", "*", "UTF-8")
    ExitApp(1)
}

TestHarnessError(failure, *) {
    FileAppend("FAIL: " failure.Message "`n" failure.Stack "`n", "*", "UTF-8")
    OnMessage(0x0082, StudyLibraryDatePickerDestroyed, 0)
    OnMessage(0x0100, StudyLibraryDatePickerKeyDown, 0)
    ExitApp(1)
}

TestAssert(condition, message) {
    global TestCount
    if !condition
        throw Error(message)
    TestCount += 1
    FileAppend("OK: " message "`n", "*", "UTF-8")
}

TestRecommendationPersistence() {
    global iniPath
    SplitPath(iniPath,, &settingsDir)
    DirCreate(settingsDir)
    settings := StudyCandidatesRecommendationDefaults()
    TestProductionStudyCandidatesRecommendationSaveSettings(settings)
    TestAssert(TestProductionStudyCandidatesRecommendationLoadSettings()["promptInstructions"] = "",
        "Missing custom prompt preserves built-in instructions")
    path := StudyCandidatesRecommendationInstructionsPath()
    settings["promptInstructions"] := "Prefer dialogue. 日本語`nSecond line."
    TestProductionStudyCandidatesRecommendationSaveSettings(settings)
    TestAssert(TestProductionStudyCandidatesRecommendationLoadSettings()["promptInstructions"]
        = settings["promptInstructions"], "Instruction edits survive settings save/reload as UTF-8")
    TestProductionStudyCandidatesRecommendationSaveSettings(settings)
    TestAssert(!FileExist(path ".bak"), "Unchanged settings do not replace the prompt or create redundant backups")
    settings["promptInstructions"] := "Revised instructions"
    TestProductionStudyCandidatesRecommendationSaveSettings(settings)
    TestAssert(FileRead(path ".bak", "UTF-8") = "Prefer dialogue. 日本語`nSecond line.",
        "Replacing custom instructions backs up the previous prompt")
    settings["promptInstructions"] := StudyCandidatesRecommendationDefaultInstructions()
    TestProductionStudyCandidatesRecommendationSaveSettings(settings)
    TestAssert(TestProductionStudyCandidatesRecommendationLoadSettings()["promptInstructions"] = ""
        && FileRead(path ".bak", "UTF-8") = "Revised instructions",
        "Restoring defaults removes the override while keeping a recoverable backup")
}

TestRecommendationWorkflow(libraryState) {
    global TestRecommendationSaved
    SetTimer(TestRecommendationConfirmVisible, -100)
    result := StudyCandidatesRecommendationConfirm(
        Map("gui", libraryState["gui"], "bigBoxPresentation", true),
        "Fixture provider", "fixture-model", Map("total", 2, "sentences", 1, "vocabulary", 1))
    TestAssert(!result && IsObject(TestRecommendationSaved)
        && !TestRecommendationSaved["focusGrammar"]
        && InStr(TestRecommendationSaved["promptInstructions"], "Prefer reusable dialogue") = 1,
        "Closing recommendation setup does not generate and preserves applied preferences")
    desktopGui := Gui("+ToolWindow", "Desktop recommendation toggle fixture")
    desktopCheck := desktopGui.AddCheckBox("w220", "Reusable vocabulary")
    StudyCandidatesRecommendationSetupBigBoxToggles(
        Map("bigBoxPresentation", false, "vocabulary", desktopCheck))
    desktopStyle := DllCall("user32\GetWindowLongPtr", "ptr", desktopCheck.Hwnd, "int", -16, "ptr")
    TestAssert(!(desktopStyle & 0x1000) && desktopCheck.Text = "Reusable vocabulary",
        "Desktop recommendation checkboxes keep their original presentation")
    desktopGui.Destroy()
}

TestRecommendationFrame(state, control) {
    control.Focus()
    ; Force one complete production paint before sampling synthetic popup
    ; pixels; DWM can otherwise defer a single unchanged strip past capture.
    state["bigBoxFocusRect"] := ""
    StudyBigBoxFocusFrameUpdate(state, control.Hwnd)
    frameDeadline := A_TickCount + 500
    Loop {
        Sleep(25)
        readyEdges := 0
        if state.Get("bigBoxFocusTarget", 0) = control.Hwnd {
            for part in state["bigBoxFocusFrame"] {
                if !DllCall("user32\IsWindowVisible", "ptr", part.Hwnd, "int")
                    continue
                part.GetPos(,, &sampleW, &sampleH)
                dc := DllCall("user32\GetDC", "ptr", part.Hwnd, "ptr")
                try {
                    if DllCall("gdi32\GetPixel", "ptr", dc,
                        "int", Max(0, sampleW // 2),
                        "int", Max(0, sampleH // 2), "uint") = 0xFFC762
                        readyEdges += 1
                } finally DllCall("user32\ReleaseDC", "ptr", part.Hwnd, "ptr", dc)
            }
        }
        if readyEdges = 4 || A_TickCount >= frameDeadline
            break
    }
    TestAssert(state.Has("bigBoxFocusWatch")
        && state.Get("bigBoxFocusTarget", 0) = control.Hwnd,
        "Recommendation focus watcher follows " state["kind"] " / " control.Type)
    targetRect := Buffer(16, 0)
    DllCall("user32\GetWindowRect", "ptr", control.Hwnd, "ptr", targetRect.Ptr)
    targetX := NumGet(targetRect, 0, "int"), targetY := NumGet(targetRect, 4, "int")
    targetW := NumGet(targetRect, 8, "int") - targetX
    targetH := NumGet(targetRect, 12, "int") - targetY
    thickness := Max(3, Round(3 * CPBigBoxDashboardDpiScale(state["gui"])))
    expected := CPFocusBorderPositions(targetX, targetY, targetW, targetH, thickness)
    visibleBlue := 0
    for index, part in state["bigBoxFocusFrame"] {
        if !DllCall("user32\IsWindowVisible", "ptr", part.Hwnd, "int")
            continue
        part.GetPos(&partX, &partY, &partW, &partH)
        edge := expected[index]
        TestAssert(partX = edge[1] && partY = edge[2]
            && partW = edge[3] && partH = edge[4],
            "Study focus border replaces the active control's neutral border")
        dc := DllCall("user32\GetDC", "ptr", part.Hwnd, "ptr")
        try {
            if DllCall("gdi32\GetPixel", "ptr", dc,
                "int", Max(0, partW // 2),
                "int", Max(0, partH // 2), "uint") = 0xFFC762
                visibleBlue += 1
        } finally DllCall("user32\ReleaseDC", "ptr", part.Hwnd, "ptr", dc)
    }
    virtualX := SysGet(76), virtualY := SysGet(77)
    virtualW := SysGet(78), virtualH := SysGet(79)
    targetOnScreen := targetX >= virtualX && targetY >= virtualY
        && targetX + targetW <= virtualX + virtualW
        && targetY + targetH <= virtualY + virtualH
    TestAssert(visibleBlue = 4 || !targetOnScreen,
        targetOnScreen
            ? "All four blue frame edges are painted on " state["kind"]
            : "Off-screen synthetic focus frame retains all four edge geometries on " state["kind"])
}

TestRecommendationConfirmVisible() {
    global TestRecommendationState, TestRecommendationCancelOnly
    if !TestRecommendationReady(TestRecommendationConfirmVisible, "confirm")
        return
    state := TestRecommendationState
    TestAssert(state["kind"] = "confirm", "Exercise the real Generate recommendations dialog")
    for key in ["generate", "level", "style", "customize", "cancel"]
        TestRecommendationFrame(state, state["controls"][key])
    TestRecommendationCancelOnly := false
    state["controls"]["customize"].Focus()
    SetTimer(TestRecommendationCustomizeVisible, -100)
    StudyCandidatesRecommendationCustomize(state)
    TestAssert(!state["settings"]["focusGrammar"] && state["settings"]["focusVocabulary"],
        "OK applies the chosen native toggle values to recommendation settings")
    TestRecommendationFrame(state, state["controls"]["customize"])
    TestRecommendationCancelOnly := true
    SetTimer(TestRecommendationCustomizeVisible, -100)
    StudyCandidatesRecommendationCustomize(state)
    TestAssert(state["settings"]["focusVocabulary"] && !state["settings"]["focusGrammar"]
        && InStr(state["settings"]["promptInstructions"], "Prefer reusable dialogue") = 1,
        "Cancel discards preference edits without undoing previously applied settings")
    state["controls"]["cancel"].Focus()
    StudyCandidatesRecommendationDialogClose(state, state["gui"], false)
    TestAssert(!state.Has("bigBoxFocusFrame"), "Closing setup immediately destroys its focus overlays")
}

TestRecommendationCustomizeVisible() {
    global TestRecommendationState, TestRecommendationCancelOnly, TestRecommendationPreviewCancel
    if !TestRecommendationReady(TestRecommendationCustomizeVisible, "customize")
        return
    state := TestRecommendationState
    TestAssert(state["kind"] = "customize", "Exercise the real Recommendation preferences dialog")
    if TestRecommendationCancelOnly {
        state["promptInstructions"] := "Discard this preference-level edit"
        state["vocabulary"].Focus()
        StudyControllerDispatchNavigation("Activate", state["gui"].Hwnd)
        Sleep(30)
        state["controls"]["cancel"].Focus()
        StudyCandidatesRecommendationAdvancedClose(state, false)
        return
    }
    for key in ["vocabulary", "grammar", "natural", "reading"] {
        toggle := state[key]
        TestRecommendationFrame(state, toggle)
        style := DllCall("user32\GetWindowLongPtr", "ptr", toggle.Hwnd, "int", -16, "ptr")
        TestAssert(toggle.Type = "CheckBox" && (style & 0x3300) = 0x3300,
            "Large toggle preserves native checkbox semantics: " key)
        StudyControllerDispatchNavigation("Activate", state["gui"].Hwnd)
        Sleep(30)
        TestAssert(toggle.Value = 0 && SubStr(toggle.Text, -4) = "`nOff",
            "Controller A changes both checkbox state and visible Off text: " key)
    }
    TestRecommendationFrame(state, state["additional"])
    TestRecommendationFrame(state, state["controls"]["restore"])
    StudyControllerDispatchNavigation("Activate", state["gui"].Hwnd)
    Sleep(30)
    for key in ["vocabulary", "grammar", "natural", "reading"]
        TestAssert(state[key].Value = 1 && SubStr(state[key].Text, -3) = "`nOn",
            "Restore defaults updates the toggle's visible On text: " key)
    for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
        TestProductionStudyFormResize(state, state["gui"], 0, size[1], size[2])
        state["natural"].GetPos(&nx, &ny, &nw, &nh)
        state["controls"]["guidanceLabel"].GetPos(, &labelY)
        state["additional"].GetPos(, &editorY,, &editorH)
        state["controls"]["hint"].GetPos(, &hintY,, &hintH)
        state["controls"]["ok"].GetPos(, &buttonY)
        TestAssert(nh >= 48 && ny + nh <= labelY && editorY + editorH <= hintY
            && hintY + hintH <= buttonY,
            "Larger recommendation toggles keep guidance and actions separate at " size[1])
    }
    TestProductionStudyFormResize(state, state["gui"], 0, 1280, 720)
    for key in ["natural", "reading"] {
        state[key].Focus()
        StudyControllerDispatchNavigation("Down", state["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(state["gui"].Hwnd) = state["additional"].Hwnd,
            "Down enters the full-width guidance editor from " key)
    }
    state["grammar"].Focus()
    StudyControllerDispatchNavigation("Activate", state["gui"].Hwnd)
    Sleep(30)
    state["controls"]["view"].Focus()
    TestRecommendationPreviewCancel := false
    SetTimer(TestRecommendationPreviewVisible, -100)
    StudyCandidatesRecommendationShowPrompt(state)
    savedInstructions := state["promptInstructions"]
    TestAssert(InStr(savedInstructions, "Prefer reusable dialogue") = 1,
        "Save returns edited instructions to the preferences draft")
    TestRecommendationPreviewCancel := true
    SetTimer(TestRecommendationPreviewVisible, -100)
    StudyCandidatesRecommendationShowPrompt(state)
    TestAssert(state["promptInstructions"] = savedInstructions,
        "Cancel in prompt editor discards changes and default restoration")
    TestRecommendationFrame(state, state["controls"]["ok"])
    StudyCandidatesRecommendationAdvancedClose(state, true)
    TestAssert(!state.Has("bigBoxFocusFrame"), "Closing preferences cleans up its focus overlays")
}

TestRecommendationPreviewVisible() {
    global TestRecommendationState, TestRecommendationPreviewCancel
    if !TestRecommendationReady(TestRecommendationPreviewVisible, "preview")
        return
    state := TestRecommendationState
    TestAssert(state["kind"] = "preview", "Exercise the real recommendation prompt preview")
    editor := state["controls"]["editor"]
    TestAssert(!StudyControllerIsReadOnlyEdit(editor.Hwnd), "Prompt instructions are keyboard-editable")
    if TestRecommendationPreviewCancel {
        StudyCandidatesRecommendationPreviewRestore(state)
        TestAssert(StudyCandidatesRecommendationNormalizeInstructions(editor.Value) = "",
            "Restore default instructions returns to the built-in prompt")
        editor.Value := "Discard this editor change"
        StudyCandidatesRecommendationPreviewClose(state, state["gui"], false)
        return
    }
    TestAssert(SendMessage(0x00B0, 0, 0, editor.Hwnd) = 0,
        "Prompt editor opens with a caret at the start, not selected text")
    TestRecommendationFrame(state, state["controls"]["editor"])
    for key in ["restore", "view", "save", "close"]
        TestRecommendationFrame(state, state["controls"][key])
    editor.Focus()
    editor.Value := "Prefer reusable dialogue. 日本語"
    loop 90
        editor.Value .= "`r`nLearning instruction " A_Index
    SendMessage(0x00B1, 0, 0, editor.Hwnd)
    SendMessage(0x0102, Ord("X"), 0, editor.Hwnd) ; WM_CHAR: real edit processing
    TestAssert(SubStr(editor.Value, 1, 1) = "X", "Keyboard text input modifies prompt instructions")
    SendMessage(0x00B1, 0, 1, editor.Hwnd)
    SendMessage(0x0102, 8, 0, editor.Hwnd) ; Backspace the inserted test character.
    original := editor.Value
    for fullPreview in [false, true] {
        StudyCandidatesRecommendationPreviewMode(state, fullPreview)
        before := editor.Value
        beforeSelection := SendMessage(0x00B0, 0, 0, editor.Hwnd)
        StudyControllerDispatchNavigation("Down", state["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(state["gui"].Hwnd) = editor.Hwnd
            && SendMessage(0x00CE, 0, 0, editor.Hwnd) > 0,
            "D-pad Down scrolls owned prompt field without leaving: preview=" fullPreview)
        TestAssert(editor.Value = before
            && SendMessage(0x00B0, 0, 0, editor.Hwnd) = beforeSelection,
            "Controller scrolling does not move the keyboard caret or modify text")
        SendMessage(0x0115, 7, 0, editor.Hwnd) ; SB_BOTTOM
        StudyControllerDispatchNavigation("Down", state["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(state["gui"].Hwnd) != editor.Hwnd,
            "Down leaves prompt field only after reaching its bottom")
        StudyControllerDispatchNavigation("Up", state["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(state["gui"].Hwnd) = editor.Hwnd,
            "Up from prompt actions returns to the prompt field")
        TestAssert(SendMessage(0x00B0, 0, 0, editor.Hwnd) = 0,
            "Returning to the prompt does not select all text")
        bottomLine := SendMessage(0x00CE, 0, 0, editor.Hwnd)
        StudyControllerDispatchNavigation("Up", state["gui"].Hwnd)
        TestAssert(SendMessage(0x00CE, 0, 0, editor.Hwnd) < bottomLine,
            "D-pad Up scrolls backward through the prompt")
    }
    TestAssert(StudyControllerIsReadOnlyEdit(editor.Hwnd)
        && InStr(editor.Value, "Return only one JSON object")
        && InStr(editor.Value, "Learning instruction 90"),
        "Full preview combines edited instructions with protected automatic output rules")
    StudyCandidatesRecommendationPreviewToggle(state)
    TestAssert(editor.Value = original && !StudyControllerIsReadOnlyEdit(editor.Hwnd),
        "Returning from full preview retains only the editable instructions")
    StudyCandidatesRecommendationPreviewMode(state, true)
    StudyCandidatesRecommendationPreviewClose(state, state["gui"], true)
    TestAssert(!state.Has("bigBoxFocusFrame"), "Closing preview cleans up its focus overlays")
}

TestRecommendationReady(callback, kind) {
    global TestRecommendationState
    state := TestRecommendationState
    foreground := DllCall("user32\GetForegroundWindow", "ptr")
    if (IsObject(state) && state.Get("kind", "") = kind
        && state.Has("bigBoxFocusWatch")
        && (!foreground || foreground = state["gui"].Hwnd))
        return true
    SetTimer(callback, -100)
    return false
}

TestVersionDisplay() {
    global CPStudyLibraryState
    savedLibraryState := CPStudyLibraryState
    versionGui := Gui("+AlwaysOnTop", "Study version display fixture")
    versionGui.BackColor := "202124"
    versionGui.SetFont("s14 cE6E6E6", "Segoe UI")
    versionLabel := versionGui.AddText("x10 y20 w100 h50 +0x200", "Version:")
    versionPrev := versionGui.AddButton("x120 y20 w50 h50", "‹")
    versionView := StudyLibraryAddVersionDisplay(versionGui, "x180 y20 w340 h50")
    versionNext := versionGui.AddButton("x530 y20 w50 h50", "›")
    versionRemove := versionGui.AddButton("x590 y20 w210 h50", "Remove version...")
    versionDdl := versionGui.AddDropDownList("Hidden", ["First", "Second", "Third"])
    versionDdl.Choose(2)
    CPStudyLibraryState := Map("gui", versionGui, "closed", false,
        "versionView", versionView, "versionDdl", versionDdl,
        "previousVersion", versionPrev, "nextVersion", versionNext,
        "removeVersionButton", versionRemove, "versions", [
            Map("version", 1, "created", "2026-09-04 02:48", "preferred", false),
            Map("version", 2, "created", "2026-09-04 15:50", "preferred", false),
            Map("version", 3, "created", "2026-09-05 10:00", "preferred", true,
                "manuallyEdited", true)
        ])
    try {
        versionGui.Show("x80 y80 w820 h130")
        WinActivate("ahk_id " versionGui.Hwnd)
        StudyLibrarySyncVersionNavigation(CPStudyLibraryState)
        TestAssert(versionView.Type = "Text" && !CPHwndIsFocusable(versionView.Hwnd),
            "Shared Library/Reader version display is a non-focusable label, not an editor")
        versionStyle := DllCall("user32\GetWindowLongPtr", "ptr", versionView.Hwnd,
            "int", -16, "ptr")
        TestAssert((versionStyle & 0x4200) = 0x4200 && !(versionStyle & 0x10000),
            "Version text is vertically centered, single-line ellipsized, and has no tab stop")
        for versionFullscreen in [false, true] {
            CPStudyLibraryState["bigBoxPresentation"] := versionFullscreen
            versionPrev.Focus()
            StudyControllerDispatchNavigation("Right", versionGui.Hwnd)
            TestAssert(StudyControllerFocusedHwnd(versionGui.Hwnd) = versionNext.Hwnd,
                "Right skips the informational version name: fullscreen=" versionFullscreen)
            StudyControllerDispatchNavigation("Left", versionGui.Hwnd)
            TestAssert(StudyControllerFocusedHwnd(versionGui.Hwnd) = versionPrev.Hwnd,
                "Left skips the informational version name: fullscreen=" versionFullscreen)
            ; Clicking an informational static label must not create an edit caret.
            SendMessage(0x0201, 1, 0, versionView.Hwnd)
            SendMessage(0x0202, 0, 0, versionView.Hwnd)
            TestAssert(StudyControllerFocusedHwnd(versionGui.Hwnd) != versionView.Hwnd,
                "Clicking the version display does not focus an editor")
        }
        for versionHeight in [30, 50, 76] {
            versionView.Move(,, 340, versionHeight)
            versionLabel.Move(,, 100, versionHeight)
            versionPrev.Move(,, 50, versionHeight)
            versionNext.Move(,, 50, versionHeight)
            versionRemove.Move(,, 210, versionHeight)
            versionRect := CPGetHwndRect(versionView.Hwnd)
            for alignedControl in [versionLabel, versionPrev, versionNext, versionRemove] {
                alignedRect := CPGetHwndRect(alignedControl.Hwnd)
                TestAssert(Abs(versionRect["cy"] - alignedRect["cy"]) < 1,
                    "Version row shares a vertical center at height " versionHeight)
            }
        }
        TestAssert(versionView.Value = "v02  •  2026-09-04 15:50",
            "Version data still updates the shared display label")
        versionDdl.Choose(3)
        StudyLibrarySyncVersionNavigation(CPStudyLibraryState)
        TestAssert(versionView.Value = "v03  •  2026-09-05 10:00  (latest)  •  manually edited"
            && versionPrev.Enabled && !versionNext.Enabled,
            "Latest/manual metadata and arrow boundaries survive the display change")
        CPStudyLibraryState["versions"] := []
        versionDdl.Delete()
        StudyLibrarySyncVersionNavigation(CPStudyLibraryState)
        TestAssert(versionView.Value = "" && !versionPrev.Enabled
            && !versionNext.Enabled && !versionRemove.Enabled,
            "An empty explanation clears the label and disables version actions")
    } finally {
        StudyBigBoxFocusFrameStop(CPStudyLibraryState)
        versionGui.Destroy()
        CPStudyLibraryState := savedLibraryState
        WinActivate("ahk_id " savedLibraryState["gui"].Hwnd)
    }
}

TestFilterHeaderMarkers() {
    headerGui := Gui("+ToolWindow", "Study filter header fixture")
    headerList := headerGui.AddListView("w420 h130", ["Date generated", "Profile", "Japanese source", "ID"])
    headerGui.Show("x100 y100 w440 h150")
    headerList.ModifyCol(1, 65)
    headerOriginalWidth := SendMessage(0x101D, 0, 0, headerList.Hwnd)
    headerState := Map("gui", headerGui, "list", headerList,
        "headerHwnd", SendMessage(0x101F, 0, 0, headerList.Hwnd),
        "closed", false, "sortColumn", 1, "sortDirection", "desc",
        "profileMode", "all", "dateMode", "all", "bigBoxLayoutScale", 1,
        "bigBoxListWidth", 400, "columns", [
            Map("key", "updated", "label", "Date generated", "index", 1, "width", 65, "visible", true),
            Map("key", "profile", "label", "Profile", "index", 2, "width", 110, "visible", true),
            Map("key", "source", "label", "Japanese source", "index", 3, "width", 200, "visible", true)
        ])
    for fullscreen in [false, true] {
        headerState["bigBoxPresentation"] := fullscreen
        for dateMode in ["today", "yesterday", "last24", "last7", "custom", "all"] {
            headerState["dateMode"] := dateMode
            StudyLibraryApplyColumns(headerState)
            expected := (dateMode = "all" ? "" : "▼  ") "Date generated"
            TestAssert(CPStudyHeaderText(headerState["headerHwnd"], 0) = expected,
                "Date header keeps its filter marker before truncated text: " dateMode " / fullscreen=" fullscreen)
            TestAssert(SendMessage(0x101D, 0, 0, headerList.Hwnd) = headerOriginalWidth,
                "Filter indication does not overwrite the user's narrow column width")
            headerFormat := Buffer(A_PtrSize = 8 ? 72 : 48, 0)
            NumPut("uint", 4, headerFormat)
            SendMessage(0x1203, 0, headerFormat.Ptr, headerState["headerHwnd"])
            TestAssert((NumGet(headerFormat, 12 + 2 * A_PtrSize, "int") & 0x600) = 0x200,
                "The separate descending-sort header flag is preserved")
        }
        headerState["profileMode"] := "value"
        StudyLibraryApplyColumns(headerState)
        TestAssert(CPStudyHeaderText(headerState["headerHwnd"], 1) = "▼  Profile"
            && CPStudyHeaderText(headerState["headerHwnd"], 2) = "Japanese source",
            "Other filtered headers use the same leading marker, while unfiltered headers stay plain")
        headerState["profileMode"] := "all"
        StudyLibraryApplyColumns(headerState)
        TestAssert(CPStudyHeaderText(headerState["headerHwnd"], 1) = "Profile",
            "Clearing a metadata filter removes its header marker")
    }
    headerGui.Destroy()
}

TestLibraryTableFocused(*) {
    global TestLibraryTableFocusEvents
    TestLibraryTableFocusEvents += 1
}

TestControlTextWidth(control) {
    dc := DllCall("user32\GetDC", "ptr", control.Hwnd, "ptr")
    font := DllCall("gdi32\SelectObject", "ptr", dc,
        "ptr", SendMessage(0x0031, 0, 0, control.Hwnd), "ptr")
    size := Buffer(8)
    try {
        DllCall("gdi32\GetTextExtentPoint32W", "ptr", dc, "wstr", control.Text,
            "int", StrLen(control.Text), "ptr", size)
        return NumGet(size, 0, "int")
    } finally {
        DllCall("gdi32\SelectObject", "ptr", dc, "ptr", font)
        DllCall("user32\ReleaseDC", "ptr", control.Hwnd, "ptr", dc)
    }
}

TestLibraryOpeningAndImageNavigation() {
    global CPStudyLibraryState, TestLibraryTableFocusEvents
    previous := CPStudyLibraryState
    fixtureGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale", "Library opening/layout fixture")
    fixtureGui.SetFont("s12", "Segoe UI")
    state := Map("gui", fixtureGui, "closed", false, "bigBoxPresentation", true,
        "compactImageNavigation", true, "bigBoxControls",
        StudyCandidatesRecommendationBigBoxShell(fixtureGui, "Study Library", "Browse saved explanations.", "Default", "ACTIVE LIBRARY"),
        "columns", [Map("key", "source", "label", "Japanese source", "index", 1, "width", 200, "visible", true)],
        "headerHwnd", 0, "sortColumn", 1, "sortDirection", "asc")
    state["libraryDdl"] := fixtureGui.AddDropDownList("x10 y10 w150", ["Default", "Fixture"])
    state["libraryDdl"].Choose(1)
    state["list"] := fixtureGui.AddListView("x10 y80 w400 h200 Grid", ["Japanese source", "ID"])
    state["list"].Add(, "Fixture explanation", 1)
    state["list"].Modify(1, "Select Focus")
    state["list"].OnEvent("Focus", TestLibraryTableFocused)
    for key in ["searchButton", "refreshButton", "filterButton", "columnsButton",
        "editDetailsButton", "studyButton", "ankiButton", "currentChapterButton", "previousVersion", "nextVersion",
        "removeVersionButton", "previousImage", "nextImage", "openImage", "exportButton", "storageButton", "bigBoxReturnButton"]
        state[key] := fixtureGui.AddButton("x0 y0 w100 h30", key)
    state["openImage"].Text := "Open full image"
    state["previousImage"].Text := "‹", state["nextImage"].Text := "›"
    for key in ["libraryLabel", "currentChapterStatus", "versionLabel", "versionView", "imageLabel",
        "imageInfo", "sourceLabel", "status", "detailTitle", "metadata", "imageFrame"]
        state[key] := fixtureGui.AddText("x0 y0 w100 h30", "")
    state["imageInfo"].Opt("+0x201")
    state["imageInfo"].Text := StudyLibraryImageCounterText(state, 1, 12)
    state["imageLabel"].Text := "Source screenshot"
    state["detailTitle"].Text := "Saved explanation"
    state["libraryLabel"].Text := "Library:"
    state["search"] := fixtureGui.AddEdit("x0 y0 w100 h30")
    state["source"] := fixtureGui.AddEdit("x0 y0 w100 h50 ReadOnly Multi")
    CPStudyLibraryState := state
    try {
        StudyBigBoxFocusFrameStart(state)
        for repeated in [false, true] {
            fixtureGui.Show("Hide w1280 h720")
            TestProductionLibraryResize(state, fixtureGui, 1280, 720)
            fixtureGui.Show("NA")
            if repeated {
                WinActivate(fixtureGui.Hwnd)
                state["list"].Focus()
                Sleep(30) ; Deliver the deliberate setup Focus event before counting startup events.
                fixtureGui.Hide()
            }
            TestLibraryTableFocusEvents := 0
            ; Same ordering as startup and reuse: show without activation,
            ; prepare the dropdown, activate, and keep the same final focus.
            fixtureGui.Show("NA")
            StudyLibraryFocusOnOpen(state)
            WinActivate(fixtureGui.Hwnd)
            state["list"].Modify(1, "Select Focus Vis")
            StudyLibraryFocusOnOpen(state)
            Sleep(90)
            TestAssert(StudyControllerFocusedHwnd(fixtureGui.Hwnd) = state["libraryDdl"].Hwnd
                && TestLibraryTableFocusEvents = 0, "Fullscreen Library opening never focuses the table: reuse=" repeated)
            TestAssert(state.Get("bigBoxFocusTarget", 0) = state["libraryDdl"].Hwnd,
                "Initial blue frame stays on the Library dropdown")
            TestAssert(state["list"].GetNext() = 1, "The initial explanation stays selected without taking controller focus")
        }
        for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
            fixtureGui.Show("x0 y0 w" size[1] " h" size[2])
            TestProductionLibraryResize(state, fixtureGui, size[1], size[2])
            state["openImage"].GetPos(&ox, &oy, &ow, &oh)
            state["nextImage"].GetPos(&nx, &ny, &nw, &nh)
            state["previousImage"].GetPos(&px, &py, &pw, &ph)
            state["imageInfo"].GetPos(&ix, &iy, &iw, &ih)
            state["imageFrame"].GetPos(&fx, &fy, &fw, &fh)
            state["search"].GetPos(, &searchY, , &searchH)
            state["libraryDdl"].GetPos(, &libraryY, , &libraryH)
            state["metadata"].GetPos(, &metadataY, , &metadataH)
            state["imageLabel"].GetPos(, &imageLabelY)
            TestAssert(TestControlTextWidth(state["openImage"]) + 20 <= ow
                && TestTextHeight(state["openImage"]) <= oh,
                "Open full image fits on one line with padding at " size[1])
            TestAssert(px + pw <= ix && ix + iw <= nx && nx + nw < ox
                && ox + ow <= fx + fw && py = iy && iy = ny && ny = oy,
                "Screenshot controls remain aligned, within the image pane, and non-overlapping at " size[1])
            TestAssert(TestControlTextWidth(state["imageInfo"]) <= iw && pw >= 28,
                "Compact counter stays readable and navigation arrows remain usable at " size[1])
            if size[2] = 720 {
                TestAssert(Abs(searchY - libraryY) <= 2
                    && Abs(searchH - libraryH) <= 1,
                    "720p search field matches the visible Library selector face: search="
                        searchY "/" searchH " library=" libraryY "/" libraryH)
                TestAssert(metadataY + metadataH <= imageLabelY
                    && fh >= 40,
                    "720p metadata clears the screenshot label without shrinking the image area")
            }
            TestStudyCapture(fixtureGui, "library-image-navigation-" size[1] ".png", size[1], size[2])
        }
        TestAssert(StudyLibraryImageCounterText(state, 2, 12) = "2 / 12"
            && StudyLibraryImageCounterText(Map(), 2, 12) = "Screenshot 2 of 12",
            "Only the fullscreen Library opts into compact screenshot counts")
        state["bigBoxPresentation"] := false
        StudyLibraryFocusOnOpen(state)
        TestAssert(StudyControllerFocusedHwnd(fixtureGui.Hwnd) = state["list"].Hwnd,
            "Desktop Library keeps initial table focus for mouse double-click")
    } finally {
        StudyBigBoxFocusFrameStop(state)
        fixtureGui.Destroy()
        CPStudyLibraryState := previous
    }
}

TestManagementAnswer() {
    global TestAnkiMessageState, TestManagementReply, TestManagementMessages
    if !IsObject(TestAnkiMessageState) || TestAnkiMessageState.Get("closed", false) {
        SetTimer(TestManagementAnswer, -20)
        return
    }
    form := TestAnkiMessageState
    TestManagementMessages += 1
    TestAssert(form["kind"] = "studyMessage", "Management confirmations and validation use the fullscreen message page")
    if TestManagementReply = "Yes" {
        StudyControllerSetFocus(form["gui"].Hwnd, form["controls"]["yes"].Hwnd)
        StudyControllerDispatchNavigation("Activate", form["gui"].Hwnd)
    } else
        StudyControllerDispatchNavigation("Cancel", form["gui"].Hwnd)
}

TestManagementWorkflow() {
    global CPStudyLibraryState, studyLibrariesRoot, studyLibrariesArchiveRoot, studyLibraryDefaultDir
    global TestManagementReply, TestAnkiMessageState, TestManagementMessages
    global CPComboSeparatorBefore
    previous := CPStudyLibraryState
    DirCreate(studyLibraryDefaultDir)
    DirCreate(studyLibrariesRoot "\Fixture")
    FileAppend("fixture database content", studyLibrariesRoot "\Fixture\marker.txt")
    rootGui := Gui("+AlwaysOnTop", "Library management root fixture")
    selector := rootGui.AddDropDownList("x10 y10 w260", ["Default", "Fixture"])
    selector.Choose(1)
    root := Map("gui", rootGui, "libraryDdl", selector, "libraryName", "Default",
        "bigBoxPresentation", true, "outputDir", A_ScriptDir "\management-output", "closed", false)
    CPStudyLibraryState := root
    rootGui.Show("w300 h70")
    StudyLibraryRefreshLibrarySelector(root, "Default")
    TestAssert(selector.Text = "Default" && SendMessage(0x146, 0, 0, selector.Hwnd) = 3,
        "Library selector lists real libraries plus one management action")
    TestAssert(CPComboSeparatorBefore.Get(selector.Hwnd, -1) = 2,
        "Library management action is separated after the real libraries")
    selector.Choose(3)
    selector.Focus()
    manager := 0
    try {
        StudyLibraryLibraryChanged(root)
        manager := root.Get("managementChild", 0)
        TestAssert(IsObject(manager) && selector.Text = "Default"
            && root["libraryName"] = "Default",
            "Management dropdown action opens the manager without changing the active library")
        TestAssert(manager["gui"].HasOwnProp("StudyLibraryManagementForm"), "Library manager opens as a fullscreen form")
        TestAssert(StudyLibraryOpenManager(root) = manager, "Duplicate manager launch reuses the existing page")
        TestRecommendationFrame(manager, manager["list"])
        TestAssert(!manager["renameButton"].Enabled && !manager["archiveButton"].Enabled
            && !manager["switchButton"].Enabled, "Default library remains protected and the active library cannot be switched to again")
        for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
            manager["gui"].Show("x0 y0 w" size[1] " h" size[2])
            StudyCandidatesRecommendationBigBoxResize(manager, manager["gui"], 0, size[1], size[2])
            manager["list"].GetPos(&lx, &ly, &lw, &lh)
            manager["switchButton"].GetPos(&bx, &by, &bw, &bh)
            TestAssert(lx + lw < bx && bh >= 38 && by = ly, "Manager table and large row actions are aligned without overlap at " size[1])
            TestAssert(SendMessage(0x101D, 0, 0, manager["list"].Hwnd) > lw * 0.30,
                "Library-name column uses a useful proportional width at " size[1])
            TestStudyCapture(manager["gui"], "library-management-" size[1] ".png", size[1], size[2])
        }
        manager["gui"].Show("x0 y0 w1280 h720")
        manager["list"].Focus()
        StudyControllerDispatchNavigation("Down", manager["gui"].Hwnd)
        Sleep(30)
        TestAssert(StudyLibraryManagerSelectedName(manager) = "Fixture"
            && StudyControllerFocusedHwnd(manager["gui"].Hwnd) = manager["list"].Hwnd,
            "D-pad Down browses rows within the management table")
        StudyControllerDispatchNavigation("Up", manager["gui"].Hwnd)
        Sleep(30)
        TestAssert(StudyLibraryManagerSelectedName(manager) = "Default", "D-pad Up browses back to the first library")
        manager["list"].Modify(0, "-Select")
        manager["list"].Modify(2, "Select Focus Vis")
        StudyLibraryManagerUpdateActions(manager)
        manager["list"].Focus()
        StudyControllerDispatchNavigation("Activate", manager["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(manager["gui"].Hwnd) = manager["switchButton"].Hwnd
            && StudyLibraryManagerSelectedName(manager) = "Fixture", "A reaches selected library actions without changing the selected row")
        TestRecommendationFrame(manager, manager["switchButton"])
        manager["list"].Focus()
        StudyControllerDispatchNavigation("Right", manager["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(manager["gui"].Hwnd) = manager["switchButton"].Hwnd,
            "Right from the library table also reaches the first selected-row action")
        StudyLibraryManagerSwitch(manager)
        TestAssert(root["libraryName"] = "Fixture" && !manager["switchButton"].Enabled,
            "Switch library updates active status without leaving management")
        manager["renameButton"].Focus()
        name := StudyLibraryManagerRename(manager)
        TestAssert(name["controls"]["name"].Value = "Fixture", "Fullscreen rename is prefilled")
        selStart := Buffer(4), selEnd := Buffer(4)
        SendMessage(0x00B0, selStart.Ptr, selEnd.Ptr, name["controls"]["name"].Hwnd)
        TestAssert(NumGet(selStart, 0, "uint") = NumGet(selEnd, 0, "uint"), "Name opens with a caret, not a full-text selection")
        TestRecommendationFrame(name, name["controls"]["name"])
        StudyControllerDispatchNavigation("Cancel", name["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(manager["gui"].Hwnd) = manager["renameButton"].Hwnd,
            "B from name entry returns directly to the originating manager action")
        name := StudyLibraryManagerRename(manager)
        name["controls"]["name"].Value := "Renamed fixture"
        StudyLibraryManagerRenameApply(manager, name["gui"], name["controls"]["name"], "Fixture")
        TestAssert(root["libraryName"] = "Renamed fixture"
            && FileRead(studyLibrariesRoot "\Renamed fixture\marker.txt") = "fixture database content",
            "Rename preserves fixture contents and active library selection")
        manager["archiveButton"].Focus()
        for reply in ["No", "Yes"] {
            TestManagementReply := reply
            TestAnkiMessageState := 0
            SetTimer(TestManagementAnswer, -20)
            StudyLibraryQueueManagementAction(manager, StudyLibraryManagerArchive.Bind(manager))
            Sleep(80)
            while manager.Get("actionRunning", false) || manager.Has("actionTimer")
                Sleep(20)
            TestAssert(!!DirExist(studyLibrariesRoot "\Renamed fixture") = (reply = "No"),
                "Archive is cancel-safe and only moves the selected fixture after explicit confirmation")
        }
        TestAssert(root["libraryName"] = "Default", "Archiving the active fixture returns to Default")
        manager["archivedButton"].Focus()
        archives := StudyLibraryOpenArchives(manager)
        TestAssert(archives["entries"].Length = 1 && archives["restoreButton"].Enabled,
            "Archived-libraries fullscreen page lists the retained fixture")
        TestRecommendationFrame(archives, archives["restoreButton"])
        TestStudyCapture(archives["gui"], "library-archives.png", 1280, 720)
        archives["list"].Focus()
        StudyControllerDispatchNavigation("Activate", archives["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(archives["gui"].Hwnd) = archives["restoreButton"].Hwnd,
            "A on an archived row reaches Restore")
        restore := StudyLibraryArchiveRestore(archives)
        TestAssert(restore["controls"]["name"].Value = "Renamed fixture", "Restore offers the original library name")
        StudyControllerDispatchNavigation("Cancel", restore["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(archives["gui"].Hwnd) = archives["restoreButton"].Hwnd,
            "Cancelling restore returns to Archives, not to the root Library")
        restore := StudyLibraryArchiveRestore(archives)
        StudyLibraryArchiveRestoreApply(archives, restore["gui"], restore["controls"]["name"], archives["entries"][1])
        TestAssert(archives["entries"].Length = 0 && !archives["restoreButton"].Enabled
            && archives["emptyText"].Text != "", "Restoring the last archive leaves a clear empty state")
        TestAssert(StudyControllerFocusedHwnd(archives["gui"].Hwnd) = archives["closeButton"].Hwnd,
            "Restoring the last archive returns focus to Back instead of a disabled Restore action")
        TestAssert(FileRead(studyLibrariesRoot "\Renamed fixture\marker.txt") = "fixture database content",
            "Restore preserves all archived fixture contents")
        StudyControllerDispatchNavigation("Cancel", archives["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(manager["gui"].Hwnd) = manager["archivedButton"].Hwnd,
            "B from Archives restores the Archived libraries button")
        manager["newButton"].Focus()
        create := StudyLibraryOpenNew(root, manager)
        TestStudyCapture(create["gui"], "library-new.png", 1280, 720)
        create["controls"]["name"].Value := "Default"
        TestManagementReply := "No", TestAnkiMessageState := 0
        SetTimer(TestManagementAnswer, -20)
        StudyLibraryCreateNew(root, create["gui"], create["controls"]["name"], manager)
        TestAssert(StudyLibraryStateAlive(create), "Invalid library name shows a fullscreen warning and keeps the name editor open")
        create["controls"]["name"].Value := "New fixture"
        StudyLibraryCreateNew(root, create["gui"], create["controls"]["name"], manager)
        TestAssert(DirExist(studyLibrariesRoot "\New fixture") && root["libraryName"] = "New fixture",
            "Create preserves existing libraries and activates the new fixture")
        TestAssert(StudyControllerFocusedHwnd(manager["gui"].Hwnd) = manager["newButton"].Hwnd,
            "Create returns to its manager action with no root-library focus detour")
        StudyControllerDispatchNavigation("Cancel", manager["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(rootGui.Hwnd) = selector.Hwnd && !root.Has("managementChild"),
            "Closing management restores Library-selector focus and clears its child registration")
        root["bigBoxPresentation"] := false
        StudyLibraryOpenManager(root)
        desktop := GuiFromHwnd(WinExist("Study Libraries ahk_class AutoHotkeyGUI"))
        TestAssert(!desktop.HasOwnProp("StudyLibraryManagementForm") && (WinGetStyle(desktop.Hwnd) & 0xC00000),
            "Desktop library manager retains its normal window layout")
        StudyLibraryCloseDialog(desktop)
        TestAssert(TestManagementMessages = 3, "Both archive decisions and invalid-name validation were exercised")
    } finally {
        if IsObject(manager) && StudyLibraryStateAlive(manager)
            StudyLibraryManagementClose(manager["gui"])
        rootGui.Destroy()
        CPStudyLibraryState := previous
    }
}

TestChapterWorkflow() {
    global CPStudyLibraryState, studyLibraryDir, TestAnkiMessageState
    global TestChapterReply, TestChapterMessageChecks, TestDesktopMessages
    global TestChapterSelectorFocus
    previousLibrary := CPStudyLibraryState
    DirCreate(studyLibraryDir)
    chapterGui := Gui("+AlwaysOnTop", "Chapter workflow library fixture")
    selector := chapterGui.AddDropDownList("x10 y10 w200", ["Fixture"])
    selector.Choose(1)
    selector.OnEvent("Focus", TestChapterSelectorFocused)
    button := chapterGui.AddButton("x230 y10 w180 h40", "Current chapter...")
    status := chapterGui.AddText("x10 y70 w400 h30", "")
    state := Map("gui", chapterGui, "closed", false, "bigBoxPresentation", true,
        "kind", "library", "libraryName", "Chapter fixture", "libraryDdl", selector,
        "currentChapterButton", button, "currentChapterStatus", status,
        "database", studyLibraryDir "\library.sqlite3")
    CPStudyLibraryState := state
    chapterGui.Show("w440 h120")
    StudyBigBoxFocusFrameStart(state)
    profile := StudyLibraryActiveProfileName()
    dialog := 0
    try {
        StudyLibraryWriteChapterHistory(profile, ["First", "Second"], studyLibraryDir)
        StudyLibraryWriteCurrentChapter(profile, "First", studyLibraryDir)
        for mode in [true, false] {
            state["bigBoxPresentation"] := mode
            selector.Focus()
            SendMessage(0x0200, 0, 20 | (20 << 16), button.Hwnd) ; synthetic WM_MOUSEMOVE
            TestAssert(SendMessage(0x00F2, 0, 0, button.Hwnd) & 0x0200,
                "Reproduce the independent native mouse-hover highlight on Current chapter")
            StudyControllerDispatchNavigation("Right", chapterGui.Hwnd)
            TestAssert(!!(SendMessage(0x00F2, 0, 0, button.Hwnd) & 0x0200) = !mode,
                "Fullscreen controller navigation clears the stray hover shade; desktop mouse hover is unchanged")
            SendMessage(0x02A3, 0, 0, button.Hwnd)
            for action in ["cancel", "save", "clear"] {
                button.Focus()
                Sleep(50)
                TestChapterSelectorFocus := 0
                dialog := StudyLibraryOpenCurrentChapter(state)
                hwnd := dialog.Hwnd
                Sleep(50)
                combo := 0
                for control in dialog
                    if control.Type = "ComboBox"
                        combo := control
                TestAssert(button.Enabled && IsObject(combo), "Chapter launch button stays enabled and undimmed while open")
                TestAssert(TestChapterSelectorFocus = 0,
                    "Opening the chapter page never transfers focus to the Library selector")
                duplicate := StudyLibraryOpenCurrentChapter(state)
                TestAssert(duplicate = dialog && state["currentChapterDialog"] = dialog,
                    "Repeated activation reuses the existing chapter page without disabling the button")
                if mode
                    SendMessage(0x0200, 0, 20 | (20 << 16), button.Hwnd)
                if action = "cancel" {
                    if mode
                        StudyControllerDispatchNavigation("Cancel", hwnd)
                    else
                        StudyLibraryCloseCurrentChapterDialog(state, dialog)
                } else {
                    combo.Text := "Saved fixture"
                    StudyLibrarySaveCurrentChapter(state, dialog, combo, profile, studyLibraryDir, action = "clear")
                    TestAssert(StudyLibraryCurrentChapter(profile, studyLibraryDir)
                        = (action = "clear" ? "" : "Saved fixture"), "Chapter save/clear persistence stays intact")
                }
                Sleep(70)
                TestAssert(!DllCall("user32\IsWindow", "ptr", hwnd) && button.Enabled
                    && StudyControllerFocusedHwnd(chapterGui.Hwnd) = button.Hwnd,
                    (mode ? "Fullscreen" : "Desktop") " chapter " action " returns focus to Current chapter")
                TestAssert(TestChapterSelectorFocus = 0 && !state.Has("currentChapterDialog"),
                    "Returning from the chapter page never flashes focus on the selector and releases its page guard")
                if mode
                    TestAssert(!(SendMessage(0x00F2, 0, 0, button.Hwnd) & 0x0200),
                        "Returning from Current chapter leaves no stale light-gray mouse-hover state")
                if mode
                    TestRecommendationFrame(state, button)
            }
        }
        state["bigBoxPresentation"] := true
        StudyLibraryWriteChapterHistory(profile, ["First", "Second"], studyLibraryDir)
        StudyLibraryWriteCurrentChapter(profile, "First", studyLibraryDir)
        dialog := StudyLibraryOpenCurrentChapter(state)
        form := dialog.StudyChapterForm
        combo := form["controls"]["chapter"]
        remove := form["controls"]["remove"]
        for reply in ["No", "Yes"] {
            remove.Focus()
            TestChapterReply := reply, TestAnkiMessageState := 0
            SetTimer(TestAnswerChapterMessage, -30)
            StudyControllerDispatchNavigation("Activate", dialog.Hwnd)
            TestWaitChapterAction(form)
            Sleep(70)
            TestAssert(StudyControllerFocusedHwnd(dialog.Hwnd) = remove.Hwnd,
                "Removal confirmation returns focus to Remove saved")
            history := StudyLibraryReadChapterHistory(profile, studyLibraryDir)
            TestAssert(history.Length = (reply = "No" ? 2 : 1)
                && StudyLibraryCurrentChapter(profile, studyLibraryDir) = (reply = "No" ? "First" : ""),
                "Chapter removal changes history/current assignment only after explicit confirmation")
        }
        StudyLibraryWriteCurrentChapter(profile, "Second", studyLibraryDir)
        historyButton := form["controls"]["history"]
        historyButton.Focus()
        TestChapterReply := "Yes", TestAnkiMessageState := 0
        SetTimer(TestAnswerChapterMessage, -30)
        StudyControllerDispatchNavigation("Activate", dialog.Hwnd)
        TestWaitChapterAction(form)
        TestAssert(StudyLibraryReadChapterHistory(profile, studyLibraryDir).Length = 0
            && StudyLibraryCurrentChapter(profile, studyLibraryDir) = "Second",
            "Clearing remembered chapters preserves the active assignment")
        TestAssert(StudyControllerFocusedHwnd(dialog.Hwnd) = historyButton.Hwnd,
            "Clear-history confirmation restores its launch button")
        StudyLibraryCloseCurrentChapterDialog(state, dialog)
        state["bigBoxPresentation"] := false
        dialog := StudyLibraryOpenCurrentChapter(state)
        before := TestDesktopMessages.Length
        result := StudyLibraryChapterMessage(state, dialog, "Desktop removal fixture", "Remove saved chapter", "yesno", "warning")
        TestAssert(result = "No" && TestDesktopMessages.Length = before + 1
            && TestDesktopMessages[-1]["owner"] = dialog.Hwnd,
            "Desktop chapter confirmation retains the existing themed dialog")
        TestAssert(TestChapterMessageChecks = 3, "All fullscreen chapter confirmations were exercised")
    } finally {
        SetTimer(TestAnswerChapterMessage, 0)
        if IsObject(TestAnkiMessageState) && !TestAnkiMessageState.Get("closed", false)
            StudyLibraryOwnedMessageClose(TestAnkiMessageState, "No")
        if IsObject(dialog)
            StudyLibraryCloseCurrentChapterDialog(state, dialog)
        StudyBigBoxFocusFrameStop(state)
        chapterGui.Destroy()
        CPStudyLibraryState := previousLibrary
    }
}

TestChapterSelectorFocused(*) {
    global TestChapterSelectorFocus
    TestChapterSelectorFocus += 1
}

TestWaitChapterAction(form) {
    deadline := A_TickCount + 8000
    Sleep(60)
    while form.Has("actionTimer") || form.Get("actionRunning", false) {
        if A_TickCount > deadline {
            TestAssert(false, "Chapter action completed within timeout")
            return
        }
        Sleep(25)
    }
}

TestAnswerChapterMessage() {
    global TestAnkiMessageState, TestChapterReply, TestChapterMessageChecks
    if !IsObject(TestAnkiMessageState) || !TestAnkiMessageState.Get("ready", false) {
        SetTimer(TestAnswerChapterMessage, -25)
        return
    }
    form := TestAnkiMessageState
    hwnd := form["gui"].Hwnd
    owner := form["libraryState"]["gui"].Hwnd
    controls := form["controls"]
    TestAssert(!(WinGetStyle("ahk_id " hwnd) & 0xC00000)
        && form["bigBoxPresentation"] && form["result"] = "No"
        && StudyControllerFocusedHwnd(hwnd) = controls["no"].Hwnd,
        "Chapter confirmation is fullscreen and defaults to the safe Cancel action")
    TestAssert(controls["no"].Text = "Cancel"
        && (controls["yes"].Text = "Remove chapter" || controls["yes"].Text = "Clear history")
        && !InStr(form["shell"]["subtitle"].Text, "Anki"),
        "Chapter confirmation uses clear chapter-specific action labels and text")
    TestAssert(!DllCall("user32\IsWindowEnabled", "ptr", owner)
        && DllCall("user32\IsWindowVisible", "ptr", owner),
        "Chapter page remains visible but cannot be changed behind its confirmation")
    for size in [[1280,720], [1920,1080], [3840,2160]] {
        form["gui"].Show("w" size[1] " h" size[2])
        TestProductionStudyFormResize(form, form["gui"], 0, size[1], size[2])
        body := controls["message"]
        body.GetPos(&x, &y, &w, &h)
        controls["no"].GetPos(, &actionY)
        TestAssert(body.Type = "Text" && !CPHwndIsFocusable(body.Hwnd)
            && body.Visible && y + h < actionY && x + w <= size[1],
            "Chapter information remains borderless and fits above actions at " size[1])
    }
    ; Return the fixture to a visible 720p surface before sampling focus-frame
    ; pixels. Leaving a 4K test window clipped by a 720p monitor can prevent
    ; off-screen layered frame edges from being composed even when their
    ; geometry and focus tracking are correct.
    form["gui"].Show("w1280 h720")
    TestProductionStudyFormResize(form, form["gui"], 0, 1280, 720)
    TestRecommendationFrame(form, controls["no"])
    TestChapterMessageChecks += 1
    if TestChapterReply = "No"
        StudyControllerDispatchNavigation("Cancel", hwnd)
    else {
        StudyControllerDispatchNavigation("Left", hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = controls["yes"].Hwnd,
            "D-pad can reach the chapter confirmation action")
        StudyControllerDispatchNavigation("Activate", hwnd)
    }
}

TestDateFilters(libraryState) {
    global TestFilterRefreshes, TestFilterWarnings
    for spec in [
        ["20260131120000", 2, 1, "20260228120000"],
        ["20240229120000", 1, 1, "20250228120000"],
        ["20000228120000", 3, 1, "20000229120000"],
        ["21000228120000", 3, 1, "21000201120000"],
        ["20260430120000", 3, 1, "20260401120000"],
        ["20261205120000", 2, 1, "20260105120000"],
        ["20260101000000", 4, -1, "20260101230000"],
        ["20260101000000", 5, -1, "20260101005900"],
        ["16010101000000", 1, -1, "16010101000000"],
        ["99991231235900", 1, 1, "99991231235900"]
    ]
        TestAssert(StudyLibraryDatePickerStep(spec[1], spec[2], spec[3]) = spec[4],
            "Date wheel respects valid calendar/time bounds: " spec[4])

    filters := libraryState.Clone()
    filters["bigBoxPresentation"] := true
    filters["libraryName"] := "Fixture"
    filters["dateMode"] := "all"
    filters["dateFrom"] := "20260131000000"
    filters["dateTo"] := "20260905121159"
    filters["ankiMode"] := "all"
    for field in ["profile", "chapter", "speaker", "tag"] {
        filters[field "Labels"] := ["Any"]
        filters[field "Choices"] := [Map("mode", "all", "value", "")]
        filters[field "Mode"] := "all"
        filters[field "Filter"] := ""
    }
    form := StudyLibraryOpenFilters(filters)
    filterGui := form["gui"]
    from := form["controls"]["fromDate"]
    to := form["controls"]["toDate"]
    TestAssert(StudyBigBoxRefreshComboFaces(form) = 6,
        "Fullscreen filter focus settling invalidates all six combo faces")
    TestAssert(form.Get("bigBoxRevealPending", false),
        "Fullscreen filter stays cloaked until the combo repaint is ready")
    TestAssert(from.Type = "Button" && to.Type = "Button"
        && from.Enabled && to.Enabled,
        "Fullscreen From/To are accessible themed buttons even for Any time")
    for control in filterGui
        TestAssert(control.Type != "DateTime", "Fullscreen filters contain no native white date fields")
    for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
        filterGui.Show("w" size[1] " h" size[2])
        TestProductionStudyFormResize(form, filterGui, 0, size[1], size[2])
        from.GetPos(&fx, &fy, &fw, &fh)
        to.GetPos(&tx, &ty, &tw, &th)
        TestAssert(fy = ty && fx + fw < tx && tx + tw <= size[1],
            "Fullscreen range buttons remain aligned and separate at " size[1])
        TestAssert(fw > 230 && fh >= 40, "Range targets remain controller-readable at " size[1])
    }
    filterGui.Show("w1280 h720")
    from.Focus()
    SendMessage(0xF5, 0, 0, from.Hwnd)
    Sleep(50)
    picker := 0
    for hwnd, entry in StudyLibraryDatePickerRegistry()
        picker := entry
    TestAssert(IsObject(picker) && picker["wheels"].Length = 5,
        "Activating From opens the production five-part date/time picker")
    TestAssert(!DllCall("user32\IsWindowEnabled", "ptr", filterGui.Hwnd),
        "Date picker owns and disables Filters while active")
    StudyControllerDispatchNavigation("Up", picker["hwnd"])
    TestAssert(SubStr(picker["stamp"], 1, 4) = "2027",
        "Controller Up changes the focused year without moving focus")
    StudyControllerDispatchNavigation("Cancel", picker["hwnd"])
    TestAssert(StudyLibraryDatePickerRegistry().Count = 0
        && form["controls"]["date"].Value = 1
        && from.Text = "2026-01-31 00:00" && filters["dateMode"] = "all",
        "B cancels the date draft without changing either filters or preset")
    TestAssert(StudyControllerFocusedHwnd(filterGui.Hwnd) = from.Hwnd
        && DllCall("user32\IsWindowEnabled", "ptr", filterGui.Hwnd),
        "Cancel restores From focus and re-enables Filters")

    SendMessage(0xF5, 0, 0, from.Hwnd)
    Sleep(50)
    for hwnd, entry in StudyLibraryDatePickerRegistry()
        picker := entry
    StudyControllerDispatchNavigation("Right", picker["hwnd"])
    TestAssert(StudyControllerFocusedHwnd(picker["hwnd"]) = picker["wheels"][2]["value"].Hwnd,
        "Controller Right selects the month wheel")
    StudyControllerDispatchNavigation("Up", picker["hwnd"])
    TestAssert(picker["stamp"] = "20260228000000",
        "Changing month from January 31 clamps the day to February 28")
    StudyLibraryDatePickerKeyDown(0x28, 0, 0x100, picker["wheels"][2]["value"].Hwnd)
    TestAssert(picker["stamp"] = "20260128000000", "Keyboard Down edits the same active wheel")
    StudyControllerDispatchNavigation("Up", picker["hwnd"])
    StudyControllerDispatchNavigation("Activate", picker["hwnd"])
    TestAssert(from.Text = "2026-02-28 00:00" && form["controls"]["date"].Value = 6
        && filters["dateFrom"] = "20260131000000" && TestFilterRefreshes = 0,
        "A updates only the dialog draft and selects Custom range")

    to.Focus()
    SendMessage(0xF5, 0, 0, to.Hwnd)
    Sleep(50)
    for hwnd, entry in StudyLibraryDatePickerRegistry()
        picker := entry
    picker["stamp"] := "20250101000000"
    StudyLibraryDatePickerClose(picker, true)
    SendMessage(0xF5, 0, 0, form["controls"]["apply"].Hwnd)
    Sleep(40)
    TestAssert(TestFilterWarnings = 1 && TestFilterRefreshes = 0
        && DllCall("user32\IsWindow", "ptr", filterGui.Hwnd),
        "Applying a reversed custom range warns without committing or closing")
    SendMessage(0xF5, 0, 0, to.Hwnd)
    Sleep(50)
    for hwnd, entry in StudyLibraryDatePickerRegistry()
        picker := entry
    picker["stamp"] := "20260905121100"
    StudyLibraryDatePickerClose(picker, true)
    SendMessage(0xF5, 0, 0, form["controls"]["apply"].Hwnd)
    Sleep(40)
    TestAssert(TestFilterRefreshes = 1 && filters["dateMode"] = "custom"
        && filters["dateFrom"] = "20260228000000" && filters["dateTo"] = "20260905121159",
        "Apply commits the staged range and includes the complete final minute")
    StudyBigBoxFocusFrameStop(form)

    form := StudyLibraryOpenFilters(filters)
    from := form["controls"]["fromDate"]
    SendMessage(0xF5, 0, 0, from.Hwnd)
    Sleep(50)
    for hwnd, entry in StudyLibraryDatePickerRegistry()
        picker := entry
    picker["gui"].Destroy()
    Sleep(40)
    TestAssert(StudyLibraryDatePickerRegistry().Count = 0 && picker["closed"]
        && DllCall("user32\IsWindowEnabled", "ptr", form["gui"].Hwnd),
        "Destroying a picker cleans its state and re-enables its owner")
    StudyBigBoxFocusFrameStop(form)
    form["gui"].Destroy()
}

TestFullscreenAnkiPreview(libraryState) {
    global CPStudyReaderState, TestAnkiPreviewState, TestAnkiMessageState
    global TestAnkiReplies, TestAnkiSends, TestAnkiWrittenFiles, TestScreenshotSaves
    global TestAnkiCancelReturns
    reader := Gui("+Owner" libraryState["gui"].Hwnd " +AlwaysOnTop -DPIScale", "Anki Reader fixture")
    reader.AddText("x10 y10 w250 h40", "Synthetic source screenshot")
    source := reader.AddEdit("x10 y60 w250 h100 ReadOnly Multi", "お城（おしろ）に行く（いく）。")
    reader.Show("x20 y20 w320 h240")
    imagePath := A_ScriptDir "\anki-source.png"
    TestStudyCapture(reader, "anki-source.png", 320, 240)
    owner := Map("gui", reader, "bigBoxPresentation", true, "libraryName", "Demo library",
        "currentGroupId", 7, "currentVersion", 2, "currentProfile", "Demo profile",
        "currentAnkiStatus", "not_checked", "outputDir", A_ScriptDir,
        "editing", false, "source", source, "sections", [Map("content", TestMultilineText())],
        "mediaIndex", 1, "media", [Map("path", imagePath)])
    CPStudyReaderState := owner
    preview := 0
    try {
        preview := TestProductionAnkiPreview(owner, "vocabulary", "城", "城（しろ） — castle. Complete entry.",
            true, TestAnkiCancelReturn)
        form := preview["bigBoxForm"]
        controls := form["controls"]
        TestAssert(StudyLibraryStateAlive(preview) && preview["bigBoxPresentation"]
            && !(WinGetStyle("ahk_id " preview["gui"].Hwnd) & 0xC00000),
            "Fullscreen Anki preview is borderless and retains the modern form")
        TestAssert(!DllCall("user32\IsWindowEnabled", "ptr", reader.Hwnd)
            && !preview["includeScreenshot"].Visible && controls["screenshotToggle"].Visible,
            "Preview disables its Reader owner and replaces the small checkbox with a large toggle")
        TestAssert(controls["screenshotToggle"].Text = "Include screenshot`nOn",
            "Screenshot toggle clearly describes inclusion and its current state")
        for size in [[1280,720], [1920,1080], [3840,2160]] {
            preview["gui"].Show("w" size[1] " h" size[2])
            TestProductionStudyFormResize(form, preview["gui"], 0, size[1], size[2])
            for key, control in controls {
                if !control.Visible
                    continue
                control.GetPos(&x,&y,&w,&h)
                TestAssert(x >= 0 && y >= 0 && x+w <= size[1] && y+h <= size[2],
                    "Anki preview control fits: " key " at " size[1])
                if control.Type = "Text" || control.Type = "Button"
                    TestAssert(TestTextHeight(control) <= h,
                        "Anki preview text fits: " key " at " size[1] " (" TestTextHeight(control) "/" h ")")
            }
            controls["back"].GetPos(&bx,&by,&bw,&bh)
            controls["status"].GetPos(,&sy)
            controls["image"].GetPos(&ix,&iy,&iw,&ih)
            controls["imageFrame"].GetPos(&fx,&fy,&fw,&fh)
            TestAssert(by+bh < sy && ix >= fx && iy >= fy && ix+iw <= fx+fw && iy+ih <= fy+fh,
                "Editors and fitted screenshot stay within their own panes at " size[1])
            TestAssert(controls["image"].Visible && Abs(iw/ih - 4/3) < 0.025,
                "Anki screenshot keeps its original 4:3 aspect at " size[1])
            TestAnkiExampleNavigation(preview, size[1])
            TestStudyCapture(preview["gui"], "anki-preview-" size[1] ".png", size[1], size[2])
        }
        preview["gui"].Show("w1280 h720")
        TestProductionStudyFormResize(form, preview["gui"], 0, 1280, 720)
        TestRecommendationFrame(form, controls["screenshotToggle"])
        StudyControllerDispatchNavigation("Activate", preview["gui"].Hwnd)
        Sleep(60) ; Native BM_CLICK queues the GUI event callback.
        TestAssert(!preview["includeScreenshot"].Value && InStr(controls["screenshotToggle"].Text,"Off")
            && TestScreenshotSaves.Length = 1 && !TestScreenshotSaves[1],
            "A toggles the screenshot option once and preserves its existing preference behavior")
        TestAssert(preview["deckDdl"].Text = "Demo::Vocabulary",
            "Preview starts with the saved vocabulary destination deck")
        StudyControllerSetFocus(preview["gui"].Hwnd, controls["deck"].Hwnd)
        StudyControllerDispatchNavigation("Activate", preview["gui"].Hwnd)
        TestAssert(CPComboDropped(controls["deck"].Hwnd), "A opens the preview destination menu")
        StudyControllerDispatchNavigation("Up", preview["gui"].Hwnd)
        StudyControllerDispatchNavigation("Cancel", preview["gui"].Hwnd)
        TestAssert(!CPComboDropped(controls["deck"].Hwnd) && !preview["closed"],
            "B closes the destination dropdown before it can close the preview")
        controls["deck"].Choose(2)
        StudyControllerSetFocus(preview["gui"].Hwnd, controls["front"].Hwnd)
        selection := Buffer(8,0)
        SendMessage(0xB0, selection.Ptr, selection.Ptr+4, controls["front"].Hwnd)
        TestAssert(NumGet(selection,0,"uint") = NumGet(selection,4,"uint")
            && !StudyControllerIsReadOnlyEdit(controls["front"].Hwnd),
            "Card front is keyboard-editable with a caret, not selected text")
        SendMessage(0x0102, Ord("a"), 0, controls["front"].Hwnd) ; WM_CHAR
        TestAssert(InStr(controls["front"].Value, "a"), "Keyboard input changes the card front")
        controls["back"].Value := TestMultilineText()
        StudyControllerSetFocus(preview["gui"].Hwnd, controls["back"].Hwnd)
        SendMessage(0x00B1, 0, 0, controls["back"].Hwnd)
        SendMessage(0x00B7, 0, 0, controls["back"].Hwnd)
        StudyControllerDispatchNavigation("Down", preview["gui"].Hwnd)
        TestAssert(StudyControllerFocusedHwnd(preview["gui"].Hwnd) = controls["back"].Hwnd
            && SendMessage(0x00CE, 0, 0, controls["back"].Hwnd) > 0,
            "D-pad Down scrolls a long card back before leaving its editor")
        controls["front"].Value := "お城"
        controls["back"].Value := "お城（おしろ） — edited definition."
        TestAnkiReplies := ["No"]
        TestAnkiMessageState := 0
        SetTimer(TestAnswerAnkiMessage, -30)
        StudyControllerSetFocus(preview["gui"].Hwnd, controls["add"].Hwnd)
        StudyControllerDispatchNavigation("Activate", preview["gui"].Hwnd)
        TestWaitAnkiMessages()
        TestAssert(TestAnkiSends = 0 && !preview["closed"] && controls["front"].Value = "お城"
            && controls["back"].Value = "お城（おしろ） — edited definition.",
            "Declining fullscreen confirmation sends nothing and retains the edited preview")
        TestAnkiReplies := ["Cancel"]
        TestAnkiMessageState := 0
        SetTimer(TestAnswerAnkiMessage, -30)
        StudyReaderQueueAnkiAction(preview, "add")
        TestWaitAnkiMessages()
        TestAssert(TestAnkiSends = 0 && !preview["closed"], "B cancels confirmation, not the preview beneath it")
        TestAnkiReplies := ["Yes", "OK"]
        TestAnkiMessageState := 0
        SetTimer(TestAnswerAnkiMessage, -30)
        StudyReaderQueueAnkiAction(preview, "add")
        StudyReaderQueueAnkiAction(preview, "add")
        TestWaitAnkiMessages()
        TestAssert(TestAnkiSends = 1 && preview["closed"]
            && TestAnkiCancelReturns = 0
            && TestAnkiWrittenFiles["anki_review_front.txt"] = "お城"
            && TestAnkiWrittenFiles["anki_review_back.txt"] = "お城（おしろ） — edited definition.",
            "Only confirmed submission sends the edited card once through the existing bridge")
        TestAssert(DllCall("user32\IsWindowEnabled", "ptr", reader.Hwnd),
            "Completing the workflow re-enables the Reader")
        longMessage := "Long error fixture`n"
        Loop 80
            longMessage .= "Detail " A_Index ": the complete error remains available to read.`n"
        TestAnkiReplies := ["OK"]
        TestAnkiMessageState := 0
        SetTimer(TestAnswerAnkiMessage, -30)
        TestAssert(StudyReaderAnkiMessage(owner, longMessage, "Anki error", "ok", "error") = "OK",
            "Long error details keep the same confirmation and return behavior")
        owner["media"] := []
        preview := TestProductionAnkiPreview(owner, "explanation", "", "")
        TestAssert(!preview["bigBoxForm"]["controls"].Has("example")
            && !preview["bigBoxForm"]["controls"]["screenshotToggle"].Enabled
            && preview["front"].Value = "お城に行く。",
            "Full explanation uses the same fullscreen preview without vocabulary-only actions")
        preview["busy"] := true
        StudyControllerDispatchNavigation("Cancel", preview["gui"].Hwnd)
        TestAssert(!preview["closed"], "Pending bridge work cannot lose its preview controls")
        preview["busy"] := false
        StudyReaderQueueAnkiAction(preview, "add")
        StudyControllerDispatchNavigation("Cancel", preview["gui"].Hwnd)
        Sleep(50)
        TestAssert(preview["closed"] && !preview.Has("actionTimer") && TestAnkiSends = 1
            && !preview["bigBoxForm"].Has("ankiAddState"),
            "B closes the idle preview, cancels queued work, and releases its form state")
        owner["bigBoxPresentation"] := false
        preview := TestProductionAnkiPreview(owner, "vocabulary", "城", "castle")
        TestAssert(!preview.Has("bigBoxForm") && preview["includeScreenshot"].Visible
            && (WinGetStyle("ahk_id " preview["gui"].Hwnd) & 0xC00000),
            "Desktop Anki preview keeps its original window and checkbox")
        for control in preview["gui"] {
            if control.Type = "Button" && control.Text = "Cancel" {
                SendMessage(0xF5, 0, 0, control.Hwnd)
                break
            }
        }
        Sleep(60)
        TestAssert(preview["closed"] && DllCall("user32\IsWindowEnabled", "ptr", reader.Hwnd),
            "Desktop Cancel returns to the Reader for a direct text-selection review")
        TestVocabularyReviewReturn(owner)
    } finally {
        SetTimer(TestAnswerAnkiMessage, 0)
        SetTimer(TestAnswerAnkiMessageSimple, 0)
        if IsObject(TestAnkiMessageState) && !TestAnkiMessageState["closed"]
            StudyLibraryOwnedMessageClose(TestAnkiMessageState, "No")
        if IsObject(preview) && !preview["closed"] {
            preview["busy"] := false
            StudyReaderCloseAnkiAddDialog(preview)
        }
        reader.Destroy()
        CPStudyReaderState := 0
    }
}

TestAnkiExampleNavigation(preview, width) {
    global TestAnkiExampleRequests
    form := preview["bigBoxForm"], controls := form["controls"], hwnd := preview["gui"].Hwnd
    for route in [["front", "Down", "example"], ["example", "Up", "front"],
        ["back", "Up", "example"], ["example", "Down", "back"]] {
        StudyControllerSetFocus(hwnd, controls[route[1]].Hwnd)
        StudyControllerDispatchNavigation(route[2], hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = controls[route[3]].Hwnd,
            "Review follows front / example / back: " route[1] " " route[2] " at " width)
    }
    StudyControllerDispatchNavigation("Up", hwnd)
    TestRecommendationFrame(form, controls["example"])
    beforeRequests := TestAnkiExampleRequests.Length
    StudyControllerDispatchNavigation("Activate", hwnd)
    Sleep(60)
    TestAssert(TestAnkiExampleRequests.Length = beforeRequests + 1
        && TestAnkiExampleRequests[-1] = preview,
        "A reaches the existing example action once, without an AI call, at " width)
    for key in ["front", "back"] {
        editor := controls[key], original := editor.Value
        longEditorText := TestMultilineText()
        Loop 7
            longEditorText .= "`r`n" TestMultilineText()
        editor.Value := longEditorText
        StudyControllerSetFocus(hwnd, editor.Hwnd)
        SendMessage(0xB1, 0, 0, editor.Hwnd)
        SendMessage(0xB7, 0, 0, editor.Hwnd)
        if key = "back"
            SendMessage(0xB6, 0, 10000, editor.Hwnd) ; EM_LINESCROLL to bottom
        exampleBeforeLine := SendMessage(0xCE, 0, 0, editor.Hwnd)
        direction := key = "front" ? "Down" : "Up"
        StudyControllerDispatchNavigation(direction, hwnd)
        exampleAfterLine := SendMessage(0xCE, 0, 0, editor.Hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = editor.Hwnd
            && (key = "front" ? exampleAfterLine > exampleBeforeLine : exampleAfterLine < exampleBeforeLine),
            "Long " key " scrolls before its example focus link at " width)
        SendMessage(0xB6, 0, key = "front" ? 10000 : -10000, editor.Hwnd)
        StudyControllerDispatchNavigation(direction, hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = controls["example"].Hwnd,
            "The " key " scroll boundary exits to Generate example at " width)
        editor.Value := original
    }
    controls["example"].Enabled := false
    try {
        StudyControllerSetFocus(hwnd, controls["front"].Hwnd)
        StudyControllerDispatchNavigation("Down", hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) != controls["example"].Hwnd
            && CPHwndIsFocusable(StudyControllerFocusedHwnd(hwnd)),
            "A disabled example action is skipped safely at " width)
    } finally controls["example"].Enabled := true
}

TestAnkiCancelReturn() {
    global TestAnkiCancelReturns
    TestAnkiCancelReturns += 1
    return true
}

TestCandidateAnkiOwner() {
    global TestReaderCloses, TestAnkiSends, TestAnkiReplies, TestAnkiMessageState
    global TestDesktopMessageReplies
    originalSends := TestAnkiSends
    dataGui := Gui("+AlwaysOnTop", "Anki candidate data fixture")
    source := dataGui.AddEdit("x10 y10 w250 h80 ReadOnly Multi", "お城（おしろ）に行く（いく）。")
    dataGui.Show("x20 y20 w280 h110")
    dataState := Map("gui", dataGui, "closed", false,
        "libraryName", "Demo library", "outputDir", A_ScriptDir,
        "currentGroupId", 7, "currentVersion", 2,
        "currentProfile", "Demo profile", "currentAnkiStatus", "not_checked",
        "editing", false, "versions", [Map("version", 2)],
        "source", source, "sections", [Map("content", TestMultilineText())],
        "mediaIndex", 0, "media", [])
    candidatesGui := Gui("+Owner" dataGui.Hwnd " +AlwaysOnTop", "Anki candidates return fixture")
    candidatesGui.Show("x320 y20 w300 h120")
    candidates := Map("gui", candidatesGui, "closeRequested", false,
        "libraryState", dataState)
    try {
        for bigBox in [false, true] {
            for action in ["cancel", "add"] {
                candidates["bigBoxPresentation"] := bigBox
                beforeCloses := TestReaderCloses
                beforeSends := TestAnkiSends
                ankiOwner := StudyCandidatesAnkiState(candidates, Map("groupId", 7))
                TestAssert(IsObject(ankiOwner) && ankiOwner["gui"] = candidatesGui
                    && ankiOwner["ankiDataState"] = dataState
                    && ankiOwner["ankiCandidateState"] = candidates
                    && ankiOwner["bigBoxPresentation"] = bigBox,
                    "Review builds a Review-owned Anki data state: fullscreen=" bigBox)
                review := TestProductionAnkiPreview(ankiOwner, "vocabulary", "城", "castle")
                if action = "cancel" {
                    for control in review["gui"] {
                        if control.Type = "Button" && control.Text = "Cancel" {
                            SendMessage(0xF5, 0, 0, control.Hwnd)
                            break
                        }
                    }
                    Sleep(80)
                } else {
                    if bigBox {
                        TestAnkiReplies := ["Yes", "OK"]
                        TestAnkiMessageState := 0
                        SetTimer(TestAnswerAnkiMessageSimple, -30)
                    } else
                        TestDesktopMessageReplies := ["Yes", "OK"]
                    StudyReaderAddReviewedAnkiNote(review)
                    if bigBox
                        TestWaitAnkiMessages()
                }
                TestAssert(review["closed"] && StudyCandidatesGuiAlive(candidates)
                    && DllCall("user32\IsWindowEnabled", "ptr", candidatesGui.Hwnd, "int")
                    && TestReaderCloses = beforeCloses
                    && TestAnkiSends = beforeSends + (action = "add" ? 1 : 0),
                    "Candidates-origin " action " returns directly to Review without a Reader: fullscreen=" bigBox)
                if action = "add"
                    TestAnkiSends := originalSends
            }
        }
    } finally {
        SetTimer(TestAnswerAnkiMessage, 0)
        SetTimer(TestAnswerAnkiMessageSimple, 0)
        if IsObject(TestAnkiMessageState) && !TestAnkiMessageState["closed"]
            StudyLibraryOwnedMessageClose(TestAnkiMessageState, "OK")
        TestAnkiSends := originalSends
        TestDesktopMessageReplies := []
        try candidatesGui.Destroy()
        try dataGui.Destroy()
    }
}

TestVocabularyReviewReturn(owner) {
    global TestVocabularyRows, TestVocabularyUseRealReview, TestVocabularyReviewState
    global TestAnkiSends, TestReaderCloses
    global TestHandoffPage, TestHandoffChecks, TestHandoffCancelDiscovery
    global TestHandoffFailDiscovery
    originalReaderCloses := TestReaderCloses
    owner["bigBoxPresentation"] := true
    TestVocabularyRows := [
        ["7", "2", "城（しろ）", "城", "城（しろ） — castle.", "castle"],
        ["7", "2", "行く（いく）", "行く", "行く（いく） — to go.", "to go"]
    ]
    TestVocabularyUseRealReview := true
    try {
        picker := StudyReaderOpenVocabularyPicker(owner)
        StudyControllerDispatchNavigation("Down", picker["gui"].Hwnd)
        Sleep(40)
        for cancelMethod in ["controller", "button"] {
            TestHandoffPage := picker
            beforeHandoffChecks := TestHandoffChecks
            StudyControllerDispatchNavigation("Activate", picker["gui"].Hwnd)
            Sleep(150)
            review := TestVocabularyReviewState
            TestAssert(IsObject(review) && !review["closed"] && review["front"].Value = "行く"
                && IsObject(review.Get("cancelReturn", 0)),
                "Picker opens the real review with its original list return context")
            TestAssert(TestHandoffChecks > beforeHandoffChecks && picker["closed"],
                "The review is composed before the outgoing vocabulary list is destroyed")
            TestAssert(DllCall("user32\IsWindowVisible", "ptr", review["gui"].Hwnd),
                "The prepared review remains visible after the list is silently removed")
            TestHandoffPage := review
            beforeHandoffChecks := TestHandoffChecks
            if cancelMethod = "controller"
                StudyControllerDispatchNavigation("Cancel", review["gui"].Hwnd)
            else
                SendMessage(0xF5, 0, 0, review["bigBoxForm"]["controls"]["cancel"].Hwnd)
            Sleep(100)
            picker := owner["vocabularyPicker"]
            TestAssert(review["closed"] && !review.Has("cancelReturn")
                && StudyReaderVocabularyPickerAlive(picker) && picker["controls"]["list"].GetNext() = 2
                && picker["controls"]["front"].Value = "行く"
                && StudyControllerFocusedHwnd(picker["gui"].Hwnd) = picker["controls"]["list"].Hwnd,
                "Review " cancelMethod " Cancel restores the vocabulary list and selected word")
            TestAssert(TestHandoffChecks > beforeHandoffChecks,
                "The outgoing review stays visible while the returning list is composed")
            TestAssert(TestAnkiSends = 1, "Cancelling a picker review sends no Anki card")
            TestHandoffPage := 0
        }
        TestHandoffPage := picker
        TestHandoffCancelDiscovery := true
        TestVocabularyReviewState := 0
        StudyControllerDispatchNavigation("Activate", picker["gui"].Hwnd)
        Sleep(150)
        TestAssert(picker["closed"] && !IsObject(TestVocabularyReviewState) && !owner["vocabularyPicker"],
            "B during discovery cancels opening without a late review or recreated list")
        TestHandoffCancelDiscovery := false
        TestHandoffPage := 0
        picker := StudyReaderOpenVocabularyPicker(owner)
        TestHandoffPage := picker
        TestHandoffFailDiscovery := true
        StudyControllerDispatchNavigation("Activate", picker["gui"].Hwnd)
        Sleep(150)
        TestAssert(StudyReaderVocabularyPickerAlive(picker) && !picker["pending"]
            && InStr(picker["controls"]["previewHint"].Value, "to review its Anki card")
            && StudyControllerFocusedHwnd(picker["gui"].Hwnd) = picker["controls"]["list"].Hwnd,
            "Discovery failure leaves the original list ready to retry, without recreating it")
        TestHandoffFailDiscovery := false
        TestHandoffPage := 0
        StudyControllerDispatchNavigation("Cancel", picker["gui"].Hwnd)
        TestAssert(picker["closed"], "B from the restored vocabulary list returns to the Reader")
        owner["currentVersion"] := 3
        TestAssert(!StudyReaderVocabularyPickerReturn(owner, 7, 2, "行く", "行く（いく） — to go."),
            "Stale review return cannot reopen vocabulary for a changed explanation version")
        owner["currentVersion"] := 2
    } finally {
        if IsObject(TestVocabularyReviewState) && !TestVocabularyReviewState["closed"]
            StudyReaderCloseAnkiAddDialog(TestVocabularyReviewState)
        if StudyReaderVocabularyPickerAlive(owner.Get("vocabularyPicker", 0))
            StudyReaderVocabularyPickerClose(owner["vocabularyPicker"])
        TestVocabularyUseRealReview := false
        TestVocabularyRows := []
        TestReaderCloses := originalReaderCloses
        TestHandoffPage := 0
        TestHandoffCancelDiscovery := false
        TestHandoffFailDiscovery := false
    }
}

TestWaitAnkiMessages() {
    global TestAnkiReplies, TestAnkiMessageState
    deadline := A_TickCount + 5000
    while A_TickCount < deadline {
        if !TestAnkiReplies.Length && IsObject(TestAnkiMessageState) && TestAnkiMessageState["closed"] {
            Sleep(100)
            return
        }
        Sleep(25)
    }
    TestAssert(false, "Anki confirmation/completion workflow completed within timeout")
}

TestAnswerAnkiMessageSimple() {
    global TestAnkiReplies, TestAnkiMessageState
    if !IsObject(TestAnkiMessageState) || TestAnkiMessageState["closed"]
        || !TestAnkiMessageState.Get("ready", false) {
        SetTimer(TestAnswerAnkiMessageSimple, -25)
        return
    }
    reply := TestAnkiReplies.RemoveAt(1)
    StudyLibraryOwnedMessageClose(TestAnkiMessageState, reply)
    if TestAnkiReplies.Length
        SetTimer(TestAnswerAnkiMessageSimple, -40)
}

TestAnswerAnkiMessage() {
    global TestAnkiReplies, TestAnkiMessageState, TestAnkiSends
    if !IsObject(TestAnkiMessageState) || TestAnkiMessageState["closed"]
        || !TestAnkiMessageState.Get("ready", false) {
        SetTimer(TestAnswerAnkiMessage, -25)
        return
    }
    form := TestAnkiMessageState
    hwnd := form["gui"].Hwnd
    isConfirmation := form["controls"].Has("no")
    action := form["controls"][isConfirmation ? "no" : "ok"]
    body := form["controls"]["message"]
    overflow := form["controls"]["messageOverflow"]
    isLong := InStr(body.Value, "Long error fixture") = 1
    TestAssert(body.Type = "Text" && !CPHwndIsFocusable(body.Hwnd),
        "Normal Anki messages are plain text without a cursor or controller focus stop")
    for size in [[1280,720], [1920,1080], [3840,2160]] {
        form["gui"].Show("w" size[1] " h" size[2])
        TestProductionStudyFormResize(form, form["gui"], 0, size[1], size[2])
        displayed := isLong ? overflow : body
        displayed.GetPos(&x,&y,&w,&h)
        action.GetPos(,&actionY)
        TestAssert(displayed.Visible && x >= 0 && y >= 0 && x+w <= size[1]
            && y+h < actionY && y+h <= size[2], "Anki message fits above its actions at " size[1])
        TestAssert(TestFontHeight(displayed) >= TestFontHeight(action) * 1.25,
            "Anki information text is comfortably larger than the action labels at " size[1])
        style := DllCall("user32\GetWindowLongPtr", "ptr", displayed.Hwnd, "int", -16, "ptr")
        exStyle := DllCall("user32\GetWindowLongPtr", "ptr", displayed.Hwnd, "int", -20, "ptr")
        TestAssert(!(style & 0x800000) && !(exStyle & 0x200),
            "Anki information has no textbox border or sunken edge at " size[1])
        if !isLong
            TestAssert(!overflow.Visible && TestTextHeight(body) <= h,
                "Short messages fit as plain text, without unnecessary scrollbars at " size[1])
    }
    form["gui"].Show("w1280 h720")
    TestProductionStudyFormResize(form, form["gui"], 0, 1280, 720)
    if isLong {
        TestAssert(!body.Visible && CPHwndIsFocusable(overflow.Hwnd),
            "Only overflowing details become a controller-scrollable reading area")
        StudyControllerSetFocus(hwnd, overflow.Hwnd)
        StudyControllerDispatchNavigation("Down", hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = overflow.Hwnd
            && SendMessage(0x00CE, 0, 0, overflow.Hwnd) > 0,
            "Controller can scroll the full borderless error details")
        StudyControllerSetFocus(hwnd, action.Hwnd)
    }
    if isConfirmation {
        TestAssert(TestAnkiSends = 0 && form["result"] = "No"
            && StudyControllerFocusedHwnd(hwnd) = form["controls"]["no"].Hwnd,
            "Confirmation is fullscreen, defaults to Back to review, and precedes any send")
        TestAssert(form["controls"]["no"].Text = "Back to review",
            "Confirmation names the previous page as a review, not a preview")
        TestAssert(InStr(form["controls"]["message"].Value, "Demo::Vocabulary")
            && InStr(form["controls"]["message"].Value, "Do not include"),
            "Confirmation reflects the selected deck and screenshot setting")
        TestRecommendationFrame(form, form["controls"]["no"])
        TestStudyCapture(form["gui"], "anki-confirmation.png", 1280, 720)
    } else {
        TestAssert(!(WinGetStyle("ahk_id " hwnd) & 0xC00000),
            "Anki completion message also keeps the fullscreen presentation")
        if !isLong
            TestStudyCapture(form["gui"], "anki-completion.png", 1280, 720)
    }
    reply := TestAnkiReplies.RemoveAt(1)
    if reply = "Cancel"
        StudyControllerDispatchNavigation("Cancel", hwnd)
    else {
        key := reply = "Yes" ? "yes" : reply = "No" ? "no" : "ok"
        StudyControllerSetFocus(hwnd, form["controls"][key].Hwnd)
        StudyControllerDispatchNavigation("Activate", hwnd)
    }
    if TestAnkiReplies.Length
        SetTimer(TestAnswerAnkiMessage, -40)
}

StudyAnkiLoadMapping(profile) => Map("profile",profile,"deck","Demo","model","Basic",
    "japaneseField","Front","explanationField","Back","addDeck","Demo","vocabularyDeck","Demo::Vocabulary")
StudyAnkiIncludeScreenshotDefault() => true
StudyAnkiSaveIncludeScreenshot(value) {
    global TestScreenshotSaves
    TestScreenshotSaves.Push(value)
}
StudyReaderGenerateVocabularyExample(state, *) {
    global TestAnkiExampleRequests
    TestAnkiExampleRequests.Push(state)
}
StudyReaderWriteAnkiReviewFile(path, value, *) {
    global TestAnkiWrittenFiles
    SplitPath(path, &name)
    TestAnkiWrittenFiles[name] := value
    return true
}
StudyAnkiSaveMapping(*) => 0
StudyAnkiReadAddResult(*) => Map("code","added","message","Fixture card added; no real Anki was contacted.")
StudyReaderRefreshAfterAnkiAdd(*) => 0
GetWindowDPI(*) => 96

TestReaderVocabularyPicker(libraryState) {
    global CPStudyReaderState, TestVocabularyRows, TestVocabularyPreviews
    global TestAnkiMenuChoices, TestDesktopVocabularySelections
    reader := Gui("+Owner" libraryState["gui"].Hwnd " +AlwaysOnTop", "Vocabulary Reader fixture")
    button := reader.AddButton("x10 y10 w160 h40", "Add to Anki…")
    explanation := reader.AddEdit("x10 y60 w400 h140 ReadOnly Multi", TestMultilineText())
    reader.Show("x50 y50 w430 h220")
    originalText := explanation.Value
    SendMessage(0xB6, 0, 3, explanation.Hwnd) ; EM_LINESCROLL
    originalScroll := SendMessage(0xCE, 0, 0, explanation.Hwnd)
    owner := Map("gui", reader, "bigBoxPresentation", false, "addAnkiButton", button,
        "currentGroupId", 7, "currentVersion", 2, "outputDir", A_ScriptDir,
        "editing", false, "libraryName", "Demo", "explanation", explanation)
    CPStudyReaderState := owner
    TestVocabularyRows := [
        ["7", "1", "Stale version", "Wrong", "Wrong", "Wrong"],
        ["99", "2", "Other group", "Wrong", "Wrong", "Wrong"],
        ["7", "2", "抜ける（ぬける）", "抜ける", "抜ける（ぬける） — to slip out. Complete explanation.", "to slip out"],
        ["7", "2", "城（しろ）", "城", "城（しろ） — castle. Complete entry with a reading.", "castle"],
        ["7", "2", "ありがとう", "ありがとう", "ありがとう — thank you", "thank you"]
    ]
    try {
        StudyReaderShowAnkiMenu(owner)
        TestAssert(TestAnkiMenuChoices[2] = "Add selected vocabulary..."
            && TestDesktopVocabularySelections = 1 && !owner.Has("vocabularyPicker"),
            "Desktop Reader retains the existing mouse-selection action")
        owner["bigBoxPresentation"] := true
        StudyReaderShowAnkiMenu(owner)
        TestAssert(TestAnkiMenuChoices[2] = "Choose vocabulary...",
            "Fullscreen Reader menu offers the controller vocabulary picker")
        picker := owner["vocabularyPicker"]
        hwnd := picker["gui"].Hwnd
        controls := picker["controls"]
        TestAssert(!controls.Has("preview") && picker["buttonControls"].Length = 1
            && controls["previewHint"].Type = "Text" && !CPHwndIsFocusable(controls["previewHint"].Hwnd)
            && InStr(controls["previewHint"].Value, "A / Cross / Enter to review")
            && InStr(picker["shell"]["footer"].Value, "Review card"),
            "Fullscreen picker replaces the redundant preview button with a non-focusable activation hint")
        TestAssert(picker["entries"].Length = 3 && controls["list"].GetNext() = 1,
            "Picker selects the first row immediately and rejects stale group/version rows")
        TestAssert(controls["front"].Value = "抜ける"
            && controls["backText"].Value = TestVocabularyRows[3][5],
            "Picker shows a reading-free front and the full original vocabulary entry")
        TestAssert(!TestVocabularyPreviews.Length, "Opening a picker does not contact the Anki preview workflow")
        for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
            picker["gui"].Show("w" size[1] " h" size[2])
            TestProductionStudyFormResize(picker, picker["gui"], 0, size[1], size[2])
            controls["list"].GetPos(&lx, &ly, &lw, &lh)
            controls["backText"].GetPos(&bx, &by, &bw, &bh)
            controls["previewHint"].GetPos(, &py)
            TestAssert(lx + lw < bx && ly + lh < py && by + bh < py,
                "Vocabulary list, preview text and actions do not overlap at " size[1])
            TestStudyCapture(picker["gui"], "vocabulary-picker-" size[1] ".png", size[1], size[2])
            for control in controls {
                item := controls[control]
                item.GetPos(&x, &y, &w, &h)
                TestAssert(x >= 0 && y >= 0 && x + w <= size[1] && y + h <= size[2],
                    "Vocabulary picker control fits: " control " at " size[1])
                if item.Type = "Text" || item.Type = "Button"
                    TestAssert(TestTextHeight(item) <= h,
                        "Vocabulary picker label is not vertically clipped: " control " at " size[1])
            }
            for key in ["eyebrow", "title", "subtitle", "footer"] {
                picker["shell"][key].GetPos(,,, &h)
                TestAssert(TestTextHeight(picker["shell"][key]) <= h,
                    "Vocabulary picker shell text fits: " key " at " size[1]
                        . " (text " TestTextHeight(picker["shell"][key]) ", control " h
                        . ", font " TestFontHeight(picker["shell"][key]) ", DPI " CPBigBoxDashboardDpiScale(picker["gui"]) ")")
            }
        }
        picker["gui"].Show("w1280 h720")
        TestProductionStudyFormResize(picker, picker["gui"], 0, 1280, 720)
        TestRecommendationFrame(picker, controls["list"])
        StudyControllerDispatchNavigation("Down", hwnd)
        Sleep(40)
        TestAssert(controls["list"].GetNext() = 2 && controls["front"].Value = "城",
            "One D-pad Down advances to the second word and updates its preview")
        StudyControllerDispatchNavigation("Right", hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = controls["backText"].Hwnd,
            "D-pad Right reaches the complete entry text")
        StudyControllerDispatchNavigation("Left", hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = controls["list"].Hwnd,
            "D-pad Left returns to the same selected vocabulary row")
        StudyControllerDispatchNavigation("Down", hwnd)
        Sleep(40)
        StudyControllerDispatchNavigation("Down", hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = controls["back"].Hwnd,
            "Down from the final word reaches Back without an obsolete preview-button stop")
        StudyControllerDispatchNavigation("Up", hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = controls["backText"].Hwnd,
            "Up from Back reaches the complete entry directly above it")
        StudyControllerDispatchNavigation("Left", hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = controls["list"].Hwnd,
            "Left from the complete entry returns to the selected word")
        StudyControllerDispatchNavigation("Up", hwnd)
        Sleep(40)
        controls["backText"].Value := TestMultilineText()
        StudyControllerDispatchNavigation("Right", hwnd)
        scrollBefore := SendMessage(0xCE, 0, 0, controls["backText"].Hwnd)
        StudyControllerDispatchNavigation("Down", hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) = controls["backText"].Hwnd
            && SendMessage(0xCE, 0, 0, controls["backText"].Hwnd) > scrollBefore,
            "D-pad scrolls a long complete entry before leaving the text")
        SendMessage(0xB6, 0, 10000, controls["backText"].Hwnd)
        StudyControllerDispatchNavigation("Down", hwnd)
        TestAssert(StudyControllerFocusedHwnd(hwnd) != controls["backText"].Hwnd,
            "Down leaves the complete entry once it reaches the end")
        StudyControllerSetFocus(hwnd, controls["list"].Hwnd)
        StudyReaderVocabularyPickerChanged(picker)
        StudyControllerDispatchNavigation("Activate", hwnd)
        Sleep(80)
        TestAssert(TestVocabularyPreviews.Length = 1 && TestVocabularyPreviews[1][2] = "vocabulary"
            && TestVocabularyPreviews[1][3] = "城" && TestVocabularyPreviews[1][4] = TestVocabularyRows[4][5],
            "Controller A passes only the selected word and full entry to the existing Anki preview")
        TestAssert(picker["closed"] && !owner["vocabularyPicker"] && owner["currentVersion"] = 2
            && owner["currentGroupId"] = 7 && explanation.Value = originalText
            && SendMessage(0xCE, 0, 0, explanation.Hwnd) = originalScroll,
            "Opening preview closes the picker without changing Reader version or text")
        StudyReaderOpenVocabularyPicker(owner)
        picker := owner["vocabularyPicker"]
        StudyControllerDispatchNavigation("Cancel", picker["gui"].Hwnd)
        TestAssert(picker["closed"] && TestVocabularyPreviews.Length = 1
            && StudyControllerFocusedHwnd(reader.Hwnd) = button.Hwnd,
            "Controller B closes the picker and restores Reader focus without adding anything")
        StudyReaderOpenVocabularyPicker(owner)
        picker := owner["vocabularyPicker"]
        Critical("On")
        try {
            StudyReaderVocabularyPickerChoose(picker)
            StudyControllerDispatchNavigation("Cancel", picker["gui"].Hwnd)
        } finally {
            Critical("Off")
        }
        Sleep(80)
        TestAssert(picker["closed"] && !picker.Has("chooseTimer") && TestVocabularyPreviews.Length = 1,
            "Closing before the deferred preview cancels it without leaving a timer behind")
        StudyReaderOpenVocabularyPicker(owner)
        picker := owner["vocabularyPicker"]
        owner["currentVersion"] := 3
        StudyReaderVocabularyPickerOpenReview(picker, 1)
        TestAssert(picker["closed"] && TestVocabularyPreviews.Length = 1,
            "Late selection cannot add vocabulary from a different Reader version")
        owner["currentVersion"] := 2
        TestVocabularyRows := []
        StudyReaderOpenVocabularyPicker(owner)
        picker := owner["vocabularyPicker"]
        TestAssert(!picker["controls"].Has("preview")
            && InStr(picker["controls"]["heading"].Value, "Vocabulary data unavailable")
            && InStr(picker["controls"]["backText"].Value, "scripts\study_library.py")
            && !InStr(picker["controls"]["backText"].Value, "No recognizable"),
            "An older bridge without the vocabulary export reports an update mismatch, not empty vocabulary")
        StudyReaderVocabularyPickerClose(picker)
        FileAppend("", A_ScriptDir "\reader_vocabulary.tsv", "UTF-8-RAW")
        StudyReaderOpenVocabularyPicker(owner)
        picker := owner["vocabularyPicker"]
        TestAssert(!picker["controls"].Has("preview")
            && InStr(picker["controls"]["backText"].Value, "No recognizable Key vocabulary entries")
            && StudyControllerFocusedHwnd(picker["gui"].Hwnd) = picker["controls"]["back"].Hwnd,
            "Empty vocabulary provides a clear message and a focused Back button")
        StudyControllerSetFocus(picker["gui"].Hwnd, picker["controls"]["list"].Hwnd)
        StudyControllerDispatchNavigation("Activate", picker["gui"].Hwnd)
        Sleep(40)
        TestAssert(!picker.Get("pending", false) && TestVocabularyPreviews.Length = 1,
            "A on an empty list does not open an Anki preview")
        StudyReaderVocabularyPickerClose(picker)
        StudyReaderVocabularyPickerOpenReview(picker, 1)
        TestAssert(TestVocabularyPreviews.Length = 1, "Late callback after closing is ignored")
        TestVocabularyRows := [
            ["7", "2", "城（しろ）", "城", "城（しろ） — castle.", "castle"],
            ["7", "2", "抜ける（ぬける）", "抜ける", "抜ける（ぬける） — to slip out.", "to slip out"],
            ["7", "2", "ありがとう", "ありがとう", "ありがとう — thank you", "thank you"]
        ]
        for targetRow in [1, 3] {
            StudyReaderOpenVocabularyPicker(owner)
            picker := owner["vocabularyPicker"]
            Loop targetRow - 1 {
                StudyControllerDispatchNavigation("Down", picker["gui"].Hwnd)
                Sleep(40)
            }
            previewCount := TestVocabularyPreviews.Length
            StudyControllerDispatchNavigation("Activate", picker["gui"].Hwnd)
            Sleep(80)
            TestAssert(TestVocabularyPreviews.Length = previewCount + 1
                && TestVocabularyPreviews[-1][3] = TestVocabularyRows[targetRow][4]
                && TestVocabularyPreviews[-1][4] = TestVocabularyRows[targetRow][5],
                "A previews the highlighted row directly without leaving the table: row " targetRow)
        }
    } finally {
        if owner.Has("vocabularyPicker") && IsObject(owner["vocabularyPicker"])
            StudyReaderVocabularyPickerClose(owner["vocabularyPicker"])
        reader.Destroy()
        CPStudyReaderState := 0
        TestVocabularyRows := []
    }
}

StudyReaderOpenReviewedAnkiDialog(owner, kind, front, back, confident := true, cancelReturn := 0, handoffSource := 0) {
    global TestVocabularyPreviews, TestVocabularyUseRealReview, TestVocabularyReviewState
    if TestVocabularyUseRealReview {
        TestVocabularyReviewState := TestProductionAnkiPreview(owner, kind, front, back, confident, cancelReturn, handoffSource)
        return TestVocabularyReviewState
    }
    TestVocabularyPreviews.Push([owner, kind, front, back])
    return Map("cancelReturn", cancelReturn)
}
StudyReaderOpenAnkiAddDialog(*) {
}
StudyReaderOpenVocabularyAnkiDialog(*) {
    global TestDesktopVocabularySelections
    TestDesktopVocabularySelections += 1
}
CPThemedChoicePopup(owner, anchor, choices, *) {
    global TestAnkiMenuChoices
    TestAnkiMenuChoices := choices
    return 2
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
CPPalette(*) => Map("window", "202124", "surface", "292A2D", "accent", "087ECC",
    "text", "EEEEEE", "muted", "AAAAAA")
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
CPThemedOwnedMessage(owner, message, title := "", buttons := "ok", *) {
    global TestDesktopMessages, TestDesktopMessageReplies
    TestDesktopMessages.Push(Map("owner", owner, "message", message, "title", title, "buttons", buttons))
    if TestDesktopMessageReplies.Length
        return TestDesktopMessageReplies.RemoveAt(1)
    return buttons = "yesno" ? "No" : "OK"
}
Toast(*) => 0
GlossaryOwnedMessage(owner, message, title := "", buttons := "ok", icon := "warning") {
    result := CPThemedOwnedMessage(owner, message, title, buttons, icon)
    return result = "Yes" ? 6 : result = "No" ? 7 : 1
}
StudyLibraryConfiguredName() => "Default"
StudyLibraryManagerSummary(*) => Map("sources", 2, "explanations", 3, "bytes", 1400000)
StudyLibraryManagerSummaryDirectory(*) => Map("sources", 2, "explanations", 3, "bytes", 1400000)
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

StudyAnkiRunBridge(state, action, *) {
    global TestAnkiStateToClose, TestAnkiCloseDuringBridge, TestAnkiSends
    global TestHandoffPage, TestHandoffCancelDiscovery
    global TestHandoffFailDiscovery
    if action = "discover" && IsObject(TestHandoffPage) {
        TestAssert(DllCall("user32\IsWindowVisible", "ptr", TestHandoffPage["gui"].Hwnd)
            && !TestHandoffPage["closed"], "Vocabulary page remains visible during Anki discovery")
        selectedRow := TestHandoffPage["controls"]["list"].GetNext()
        StudyControllerDispatchNavigation("Up", TestHandoffPage["gui"].Hwnd)
        TestAssert(TestHandoffPage["controls"]["list"].GetNext() = selectedRow,
            "The selected word stays fixed during the handoff")
        if TestHandoffCancelDiscovery
            StudyControllerDispatchNavigation("Cancel", TestHandoffPage["gui"].Hwnd)
        if TestHandoffFailDiscovery
            return false
    }
    if action = "add-note"
        TestAnkiSends += 1
    if TestAnkiCloseDuringBridge
        StudyAnkiCloseDialog(TestAnkiStateToClose)
    return true
}

StudyAnkiReadStatus(*) => Map("code", "connected", "version", "6", "message", "")
StudyLibraryReadRows(path, *) {
    global TestVocabularyRows
    if InStr(path, "anki_decks.tsv")
        return [["Demo"], ["Demo::Vocabulary"]]
    if InStr(path, "anki_models.tsv")
        return [["Basic", "Front"], ["Basic", "Back"]]
    return InStr(path, "reader_vocabulary.tsv") ? TestVocabularyRows : []
}
StudyLibraryHexDecode(value) => value
StudyAnkiUpdateConnectionText(*) => true
StudyAnkiApplyProfileMapping(*) {
}

CPControllerKeyboardMirrorActive(*) => false
CPBigBoxDashboardAlive(*) => false
CPBigBoxFocusColor(*) => "62C7FF"
CPRegisterColorSwatch(control, *) => control

StudyCandidatesRecommendationBigBoxResize(state, gui, minmax, width, height) {
    global TestColumnRelayouts, TestRecommendationState, TestAnkiPreviewState, TestAnkiMessageState
    global TestHandoffPage, TestHandoffChecks
    if state.Get("closed", false)
        return
    ; Queued Size notifications may arrive after the successful handoff has
    ; destroyed its outgoing page. Inspect only the live transition surface.
    outgoingHwnd := 0
    if IsObject(TestHandoffPage) && TestHandoffPage["gui"] != gui
        try outgoingHwnd := TestHandoffPage["gui"].Hwnd
    if outgoingHwnd && DllCall("user32\IsWindow", "ptr", outgoingHwnd) {
        TestAssert(DllCall("user32\IsWindowVisible", "ptr", outgoingHwnd),
            "Incoming page layout keeps the outgoing page visible, covering the Reader")
        TestHandoffChecks += 1
    }
    if gui.HasOwnProp("StudyOwnedMessage")
        TestAnkiMessageState := gui.StudyOwnedMessage
    if (state.Get("kind", "") = "ankiPreview") {
        TestAnkiPreviewState := state["ankiAddState"]
        return TestProductionStudyFormResize(state, gui, minmax, width, height)
    }
    if (state.Get("kind", "") = "confirm" || state.Get("kind", "") = "customize"
        || state.Get("kind", "") = "preview") {
        TestRecommendationState := state
        return TestProductionStudyFormResize(state, gui, minmax, width, height)
    }
    if (state.Get("kind", "") = "studyForm" || state.Get("kind", "") = "vocabularyPicker"
        || state.Get("kind", "") = "studyMessage")
        return TestProductionStudyFormResize(state, gui, minmax, width, height)
    TestColumnRelayouts += 1
}

StudyCandidatesRecommendationLoadSettings() => StudyCandidatesRecommendationDefaults()
StudyCandidatesRecommendationSaveSettings(settings) {
    global TestRecommendationSaved
    TestRecommendationSaved := settings.Clone()
}

CPBigBoxMonitorBounds(*) => Map("x", 0, "y", 0, "w", 1280, "h", 720)
CPSetWindowCloaked(*) => 0
CPApplyDialogCheckBoxTheme(*) => 0
CPRegisterMutedControl(*) => 0
StudyLibraryInternalImageRender(*) => 0
StudyLibraryRefresh(*) {
    global TestFilterRefreshes
    TestFilterRefreshes += 1
}
CPAdaptiveOwnedMessage(*) {
    global TestFilterWarnings
    TestFilterWarnings += 1
}

CPDialogDefaultResult(buttons) {
    return buttons = "yesno" ? "No"
        : buttons = "yesnocancel" ? "Cancel" : "OK"
}

StudyCandidatesResizeBigBox(*) {
}

StudyLibraryResizeBigBox(*) {
}
StudyLibraryQueueImageLayout(*) => 0
StudyLibraryRedraw(*) => 0

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

; Native messages go only to the offscreen synthetic GUI, never to other apps.
TestProfileSelection() {
    global ui, CPDesktop, iniPath, TestProfileState, GameProfileLastError
    GameProfileLastError := "Synthetic profile failure"
    TestProfileState := Map("applied", 0, "saved", 0, "notices", 0, "page", 0,
        "dirty", false, "answer", "Cancel", "saveOK", true, "applyOK", true)
    ui.Show("NA x-9000 y-9000 w1120 h760")
    CPDesktopLayout(ui, 0, 1120, 760)
    ctrl := CPDesktop["chrome"]["profile"]
    for active in ["", "Demo profile"] {
        IniWrite(active, iniPath, "game_profiles", "active")
        CPDesktopRefreshProfileSelector(true)
        original := ctrl.Text
        for route in ["native", "controller", "dismiss"] {
            ctrl.Focus()
            CPShowCombo(ctrl.Hwnd, true)
            Sleep(30)
            DesktopAssert(CPComboDropped(ctrl.Hwnd) && IsObject(CPDesktop.Get("profileSelection", 0)), "Opening starts a preview session")
            browsed := false
            for key in [0x28, 0x28, 0x26] {
                if route = "controller"
                    CPControllerDispatchNavigation(key = 0x28 ? "Down" : "Up", ui.Hwnd)
                else
                    SendMessage(0x100, key, 1, ctrl.Hwnd) ; WM_KEYDOWN: Down, Down, Up
                Sleep(25)
                browsed := browsed || ctrl.Text != original
                highlighted := ctrl.Text
                CPDesktopRefreshProfileSelector()
                DesktopAssert(ctrl.Text = highlighted && CPComboDropped(ctrl.Hwnd), "Unchanged status refresh preserves dropdown browsing")
                DesktopAssert(TestProfileState["applied"] = 0 && TestProfileState["notices"] = 0,
                    "Arrow/D-pad browsing does not apply or prompt")
                DesktopAssert(IniRead(iniPath, "game_profiles", "active", "") = active, "Browsing preserves the active profile")
            }
            DesktopAssert(browsed, "Native arrows actually highlight a different choice")
            if route = "native"
                SendMessage(0x100, 0x1B, 1, ctrl.Hwnd)
            else if route = "controller"
                CPControllerDispatchNavigation("Cancel", ui.Hwnd)
            else
                CPShowCombo(ctrl.Hwnd, false)
            Sleep(30)
            DesktopAssert(!CPComboDropped(ctrl.Hwnd) && ctrl.Text = original, "Cancel/dismiss restores the active header label: " route)
        }
        ; A closed combo's type-ahead/arrow selection is not confirmation.
        SendMessage(0x100, 0x28, 1, ctrl.Hwnd)
        Sleep(30)
        DesktopAssert(ctrl.Text = original && TestProfileState["applied"] = 0, "Closed-combo browsing cannot activate a profile")
    }
    ; Exercise A/Enter, native Enter, and native list mouse selection separately.
    for route in ["controller", "native", "mouse"] {
        IniWrite("Demo profile", iniPath, "game_profiles", "active")
        CPDesktopRefreshProfileSelector(true)
        applyBefore := TestProfileState["applied"]
        ctrl.Focus()
        CPControllerDispatchNavigation("Activate", ui.Hwnd) ; A opens the list, it does not switch yet.
        Sleep(30)
        ctrl.Choose("Alternate profile")
        DesktopAssert(TestProfileState["applied"] = applyBefore, "Highlight is preview-only")
        if route = "controller" {
            Critical("On")
            CPControllerDispatchNavigation("Activate", ui.Hwnd)
            CPControllerDispatchNavigation("Activate", ui.Hwnd)
            Critical("Off")
        }
        else if route = "native"
            SendMessage(0x100, 0x0D, 1, ctrl.Hwnd)
        else
            TestProfileMouseSelect(ctrl, ctrl.Value)
        Sleep(100)
        DesktopAssert(TestProfileState["applied"] = applyBefore + 1, "Exactly one profile switch on confirmation: " route)
        DesktopAssert(ctrl.Text = "Alternate profile" && !CPComboDropped(ctrl.Hwnd), "Confirmed profile is displayed and dropdown closes: " route)
    }
    TestProfileState["dirty"] := true
    for answer in ["Cancel", "No", "Yes"] {
        IniWrite("Demo profile", iniPath, "game_profiles", "active")
        CPDesktopRefreshProfileSelector(true)
        TestProfileState["answer"] := answer
        applyBefore := TestProfileState["applied"], notices := TestProfileState["notices"], saved := TestProfileState["saved"]
        CPShowCombo(ctrl.Hwnd, true)
        Sleep(25)
        ctrl.Choose("Alternate profile")
        Sleep(25)
        DesktopAssert(TestProfileState["notices"] = notices, "Unsaved prompt is deferred until confirmation")
        CPDesktopProfileSelectionConfirm(ctrl)
        Sleep(100)
        DesktopAssert(TestProfileState["notices"] = notices + 1, "Confirmation asks about unsaved changes once")
        DesktopAssert(TestProfileState["applied"] = applyBefore + (answer != "Cancel"), "Dirty-profile decision preserved: " answer)
        DesktopAssert(TestProfileState["saved"] = saved + (answer = "Yes"), "Save occurs only for Save and switch")
        DesktopAssert(ctrl.Text = (answer = "Cancel" ? "Demo profile" : "Alternate profile"), "Header matches final active profile")
    }
    TestProfileState["dirty"] := false
    for failure in ["save", "apply"] {
        IniWrite("Demo profile", iniPath, "game_profiles", "active")
        CPDesktopRefreshProfileSelector(true)
        TestProfileState["dirty"] := failure = "save"
        TestProfileState["answer"] := "Yes"
        TestProfileState["saveOK"] := failure != "save"
        TestProfileState["applyOK"] := failure != "apply"
        CPShowCombo(ctrl.Hwnd, true)
        Sleep(25)
        ctrl.Choose("Alternate profile")
        CPDesktopProfileSelectionConfirm(ctrl)
        Sleep(100)
        DesktopAssert(ctrl.Text = "Demo profile", "Failed " failure " restores the active header label")
        DesktopAssert(!IsObject(CPDesktop.Get("profileSelection", 0)), "Failed " failure " clears preview state")
    }
    TestProfileState["dirty"] := false
    TestProfileState["saveOK"] := true, TestProfileState["applyOK"] := true
    applyBefore := TestProfileState["applied"]
    CPShowCombo(ctrl.Hwnd, true)
    Sleep(25)
    CPDesktopProfileSelectionConfirm(ctrl)
    Sleep(75)
    DesktopAssert(TestProfileState["applied"] = applyBefore, "Confirming the active profile is a no-op")
    for confirm in [false, true] {
        CPShowCombo(ctrl.Hwnd, true)
        Sleep(25)
        ctrl.Choose("Manage profiles…")
        Sleep(25)
        DesktopAssert(TestProfileState["page"] = 0, "Browsing Manage profiles does not navigate")
        if confirm
            CPDesktopProfileSelectionConfirm(ctrl)
        else
            CPDesktopProfileSelectionCancel(ctrl)
        Sleep(75)
    }
    DesktopAssert(TestProfileState["page"] = 7 && TestProfileState["applied"] = applyBefore, "Confirmed Manage profiles opens the page without switching")
    ; Refresh between confirmation and the deferred callback invalidates it.
    CPShowCombo(ctrl.Hwnd, true)
    Sleep(25)
    ctrl.Choose("Alternate profile")
    Critical("On")
    CPDesktopProfileSelectionConfirm(ctrl)
    CPDesktopRefreshProfileSelector(true)
    Critical("Off")
    Sleep(75)
    DesktopAssert(TestProfileState["applied"] = applyBefore, "A refreshed list cannot execute a stale pending switch")
    TestProfilesPageSelection()
}

TestProfilesPageSelection() {
    global ui, CPDesktop, ddlGameProfile, txtGameProfileState, iniPath, TestProfileState, btnGameProfileApply
    IniWrite("Demo profile", iniPath, "game_profiles", "active")
    RefreshGameProfilesList("Alternate profile")
    CPDesktopNavigate(7)
    ctrl := ddlGameProfile
    applyBefore := TestProfileState["applied"]
    for route in ["native", "controller"] {
        for ending in ["confirm", "cancel", "dismiss", "mouse"] {
            ctrl.Choose("Alternate profile")
            GameProfileUpdateSummary()
            original := ctrl.Text, originalSummary := txtGameProfileState.Text
            ctrl.Focus()
            CPControllerDispatchNavigation("Activate", ui.Hwnd)
            Sleep(30)
            DesktopAssert(CPComboDropped(ctrl.Hwnd), "Profiles page list opens")
            Loop 9 {
                direction := Mod(A_Index, 2) ? "Down" : "Up"
                if route = "controller"
                    CPControllerDispatchNavigation(direction, ui.Hwnd)
                else
                    TestProfileNativeArrow(ctrl.Hwnd, direction)
                Sleep(35)
                GameProfileUpdateSummary() ; Other refresh callers must also leave the list intact.
                DesktopAssert(CPComboDropped(ctrl.Hwnd), "Profiles page stays open while browsing: " route " press " A_Index)
                DesktopAssert(txtGameProfileState.Text = originalSummary, "Summary is not committed during preview")
                DesktopAssert(IniRead(iniPath, "game_profiles", "active") = "Demo profile"
                    && TestProfileState["applied"] = applyBefore, "Profiles-page browsing never applies a profile")
            }
            choice := ctrl.Text
            DesktopAssert(choice != original, "Profiles-page navigation changed the highlighted item")
            ; The header has its own list; updating it must not cancel this preview.
            CPDesktopRefreshProfileSelector(true)
            DesktopAssert(CPComboDropped(ctrl.Hwnd), "Header refresh leaves the Profiles-page list open")
            if ending = "confirm" {
                if route = "controller"
                    CPControllerDispatchNavigation("Activate", ui.Hwnd)
                else
                    SendMessage(0x100, 0x0D, 1, ctrl.Hwnd)
            } else if ending = "cancel" {
                if route = "controller"
                    CPControllerDispatchNavigation("Cancel", ui.Hwnd)
                else
                    SendMessage(0x100, 0x1B, 1, ctrl.Hwnd)
            } else if ending = "mouse"
                TestProfileMouseSelect(ctrl, ctrl.Value)
            else
                CPShowCombo(ctrl.Hwnd, false)
            Sleep(100)
            expected := ending = "confirm" || ending = "mouse" ? choice : original
            DesktopAssert(!CPComboDropped(ctrl.Hwnd) && ctrl.Text = expected, "Profiles-page close result: " route " " ending)
            DesktopAssert(InStr(txtGameProfileState.Text, "Selected: " expected), "Summary matches the confirmed/cancelled selection")
            DesktopAssert(IniRead(iniPath, "game_profiles", "active") = "Demo profile"
                && TestProfileState["applied"] = applyBefore, "Confirmation only selects; Apply profile is still required")
        }
    }
    ctrl.Choose("Alternate profile")
    GameProfileUpdateSummary()
    SendMessage(0xF5, 0, 0, btnGameProfileApply.Hwnd) ; BM_CLICK, synthetic Apply handler
    Sleep(100)
    DesktopAssert(TestProfileState["applied"] = applyBefore + 1
        && IniRead(iniPath, "game_profiles", "active") = "Alternate profile",
        "Only the separate Apply profile button loads the selected profile")
    DesktopAssert(CPDesktop["chrome"]["profile"].Text = "Alternate profile",
        "Apply profile still synchronizes the header selector")
}

TestProfileMouseSelect(ctrl, index) {
    info := Buffer(A_PtrSize = 8 ? 64 : 52, 0)
    NumPut("uint", info.Size, info)
    DllCall("user32\GetComboBoxInfo", "ptr", ctrl.Hwnd, "ptr", info)
    list := NumGet(info, A_PtrSize = 8 ? 56 : 48, "ptr")
    rect := Buffer(16, 0)
    SendMessage(0x198, index - 1, rect.Ptr, list) ; LB_GETITEMRECT
    point := 8 | ((NumGet(rect, 4, "int") + 4) << 16)
    SendMessage(0x201, 1, point, list)
    SendMessage(0x202, 0, point, list)
}

TestProfileNativeArrow(hwnd, keyName) {
    SendMessage(0x100, GetKeyVK(keyName), 1, hwnd)
}

TestProfileDirty(*) {
    global TestProfileState
    return TestProfileState["dirty"]
}
TestProfileSave(*) {
    global TestProfileState
    TestProfileState["saved"] += 1
    return TestProfileState["saveOK"]
}
TestProfileApply(name, *) {
    global TestProfileState, iniPath
    TestProfileState["applied"] += 1
    if TestProfileState["applyOK"]
        IniWrite(name, iniPath, "game_profiles", "active")
    return TestProfileState["applyOK"]
}
TestProfileNotice(*) {
    global TestProfileState, CPDesktop
    DesktopAssert(!CPComboDropped(CPDesktop["chrome"]["profile"].Hwnd), "Dropdown closes before the modal prompt")
    TestProfileState["notices"] += 1
    return TestProfileState["answer"]
}
TestProfileRefresh(*) {
    CPDesktopRefreshProfileSelector(true)
}
TestProfileNavigate(page) {
    global TestProfileState
    TestProfileState["page"] := page
}
TestProfileNoop(*) {
}

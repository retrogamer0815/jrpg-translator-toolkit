; Exercise actual Windows dialog navigation; never send input to other apps.
TestMessageNavigation() {
    global ui, TestMessageState
    for buttons in ["yesnocancel", "yesno", "ok"] {
        for longMessage in [false, true] {
            message := "Current settings differ from the saved profile. Save before switching?"
            if longMessage
                Loop 60
                    message .= " Long diagnostic information remains scrollable."
            s := StudyDesktopMessageCreate(ui.Hwnd, message, "Switch profile", buttons,
                "Confirmation", "Save and switch", "Switch without saving", "Cancel")
            TestMessageState := s
            try {
                DesktopAssert(s["result"] = CPDialogDefaultResult(buttons), "Safe default result: " buttons)
                for size in [[640, 400], [760, 520], [1040, 700]] {
                    DesktopDialogFixtureShow(s, size[1], size[2])
                    DesktopAssert(s["overflow"].Visible = longMessage, "Scrollable and ordinary message layouts covered")
                    keys := buttons = "yesnocancel" ? ["addButton", "noButton", "cancel"]
                        : buttons = "yesno" ? ["addButton", "cancel"] : ["cancel"]
                    for index, key in keys {
                        if index = keys.Length
                            continue
                        left := s[key], right := s[keys[index + 1]]
                        left.GetPos(&leftX), right.GetPos(&rightX)
                        DesktopAssert(leftX < rightX, "Test follows visible left-to-right button order")
                        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", s["gui"].Hwnd,
                            "ptr", left.Hwnd, "int", false, "ptr") = right.Hwnd, "Tab order matches visual order")
                        for route in ["keyboard", "controller"] {
                            for direction in ["Right", "Left"] {
                                origin := direction = "Right" ? left : right
                                destination := direction = "Right" ? right : left
                                origin.Focus()
                                if route = "keyboard"
                                    TestMessageNativeKey(direction)
                                else
                                    TestMessageControllerDispatch(direction, s["gui"].Hwnd)
                                DesktopAssert(StudyControllerFocusedHwnd(s["gui"].Hwnd) = destination.Hwnd,
                                    route " " direction " follows the visible button row: " buttons)
                                DesktopAssert(!s["closed"], "Navigation never activates a profile action")
                            }
                        }
                    }
                    s["cancel"].Focus()
                    DesktopAssert(StudyControllerFocusedHwnd(s["gui"].Hwnd) = s["cancel"].Hwnd,
                        "Safe Cancel/No/OK focus remains available")
                }
                TestMessageNativeKey("Esc")
                Sleep(25)
                DesktopAssert(s["closed"] && s["result"] = CPDialogDefaultResult(buttons),
                    "Escape preserves the safe default result")
            } finally {
                try s["gui"].Destroy()
            }
        }
    }
    ; Creation order must not change which callback/result belongs to a button.
    for route in ["keyboard", "controller"] {
        for spec in [["addButton", "Yes"], ["noButton", "No"], ["cancel", "Cancel"]] {
            s := StudyDesktopMessageCreate(ui.Hwnd, "Synthetic confirmation", "Switch profile", "yesnocancel",
                "Confirmation", "Save and switch", "Switch without saving", "Cancel")
            TestMessageState := s
            DesktopDialogFixtureShow(s, 760, 520)
            s["cancel"].Focus()
            Loop (spec[1] = "addButton" ? 2 : spec[1] = "noButton" ? 1 : 0)
                TestMessageNativeKey("Left")
            DesktopAssert(StudyControllerFocusedHwnd(s["gui"].Hwnd) = s[spec[1]].Hwnd, "Desired action has real keyboard focus")
            DesktopDialogAssertChrome(s, 760, 520)
            if route = "keyboard"
                TestMessageNativeKey("Enter")
            else
                TestMessageControllerDispatch("Activate", s["gui"].Hwnd)
            Sleep(25)
            DesktopAssert(s["closed"] && s["result"] = spec[2], route " activates the focused action: " spec[2] " (actual " s["result"] ")")
            try s["gui"].Destroy()
        }
    }
    for spec in [["addButton", "Yes"], ["noButton", "No"], ["cancel", "Cancel"]] {
        s := StudyDesktopMessageCreate(ui.Hwnd, "Synthetic confirmation", "Switch profile", "yesnocancel",
            "Confirmation", "Save and switch", "Switch without saving", "Cancel")
        DesktopDialogFixtureShow(s, 760, 520)
        SendMessage(0xF5, 0, 0, s[spec[1]].Hwnd)
        Sleep(25)
        DesktopAssert(s["closed"] && s["result"] = spec[2], "Button retains its result: " spec[2])
        try s["gui"].Destroy()
    }
}

TestMessageNativeKey(direction) {
    global TestMessageState
    g := TestMessageState["gui"]
    vk := GetKeyVK(direction)
    ; MSG offsets differ on x64/x86. IsDialogMessage executes the same native
    ; arrow-key routing used by AHK's message loop, including button tab order.
    msg := Buffer(A_PtrSize = 8 ? 48 : 28, 0)
    NumPut("ptr", StudyControllerFocusedHwnd(g.Hwnd), msg, 0)
    NumPut("uint", 0x100, msg, A_PtrSize)
    NumPut("uptr", vk, msg, A_PtrSize = 8 ? 16 : 8)
    NumPut("ptr", 1, msg, A_PtrSize = 8 ? 24 : 12)
    DllCall("user32\IsDialogMessageW", "ptr", g.Hwnd, "ptr", msg)
    Sleep(10)
}

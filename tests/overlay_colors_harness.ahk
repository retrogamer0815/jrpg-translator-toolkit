; Regression for legacy Background options resetting BS_OWNERDRAW on the color
; buttons. Use an offscreen synthetic desktop, never real settings or overlays.
TestOverlayColors() {
    global ui, tab, CPDesktop, controlDarkMode
    global boxBgHex, txtHex, nameHex, boxBgHex_EW, txtHex_EW
    global TestColorResult := "", TestColorInitial := "", TestColorSaves := 0, TestColorSends := 0
    original := [boxBgHex, txtHex, nameHex, boxBgHex_EW, txtHex_EW, controlDarkMode]
    try {
        ui.Show("NA x-9000 y-9000 w1120 h760")
        for dark in [1, 0] {
            controlDarkMode := dark
            CPApplyControlPanelTheme()
            for page in [3, 5] {
                CPDesktopNavigate(page)
                CPDesktopLayout(ui, 0, 1120, 760)
                title := page = 3 ? "Translator" : "Explainer"
                for field in (page = 3 ? ["bg", "txt", "name"] : ["bg", "txt"]) {
                    nextColor := dark ? "0F27FF" : "E64B29"
                    ; A new color, accepting the same color again, and Cancel.
                    for result in [nextColor, nextColor, ""] {
                        priorColors := TestOverlayColorSnapshot()
                        TestColorResult := result
                        saves := TestColorSaves, sends := TestColorSends
                        TestCPDesktopPickOverlayColor(page, field)
                        DesktopAssert(TestColorInitial = priorColors[title ":" field], "Picker starts with the current color")
                        DesktopAssert(TestColorSaves = saves + (result != "") && TestColorSends = sends + (result != ""),
                            "Accepted colors save/send once; Cancel does neither")
                        for key, value in TestOverlayColorSnapshot()
                            DesktopAssert(value = (key = title ":" field && result != "" ? result : priorColors[key]),
                                "Color edit changes only its own setting: " key)
                        ; No navigation, layout or theme pass between editing and
                        ; these assertions: those passes used to hide the bug.
                        TestOverlayColorState("picker " title ":" field)
                        TestOverlayColorPixels(page)
                    }
                }
                WinGetClientPos(,, &width, &height, ui.Hwnd)
                TestCapture("overlay-colors-" (dark ? "dark-" : "light-") page ".png", width, height)
            }
            ; Repair a clobbered native type even when the RGB did not change.
            ; This protects future callers and the periodic status refresh path.
            for page, controls in CPDesktop["overlayPages"]
                for field in (page = 3 ? ["bg", "txt", "name"] : ["bg", "txt"])
                    controls[field].Opt("Background123456")
            CPDesktopRefreshOverlayColors()
            TestOverlayColorState("same-color style recovery")
            TestOverlayColorPixels(5)

            ; Profile/settings reloads and fullscreen edits may update a hidden
            ; desktop page. Both shared refresh entry points must be safe.
            CPDesktopNavigate(1)
            TestApplyColorValue("name", "A1B2C3")
            TestApplyColorValue_EW("txt", "C3B2A1")
            RefreshColorSwatches()
            RefreshColorSwatches_EW()
            TestOverlayColorState("hidden desktop colors")
            DesktopAssert(tab.Value = 1, "Refreshing hidden colors does not navigate or reveal their page")
        }
    } finally {
        boxBgHex := original[1], txtHex := original[2], nameHex := original[3]
        boxBgHex_EW := original[4], txtHex_EW := original[5], controlDarkMode := original[6]
        CPDesktopRefreshOverlayColors()
        CPApplyControlPanelTheme()
        CPDesktopNavigate(1)
    }
}

TestOverlayColorSnapshot() {
    values := Map()
    for title in ["Translator", "Explainer"]
        for field in (title = "Translator" ? ["bg", "txt", "name"] : ["bg", "txt"])
            values[title ":" field] := CPOverlayPreference(title, field)["value"]
    return values
}

TestOverlayColorState(context) {
    global CPDesktop
    for page, controls in CPDesktop["overlayPages"] {
        for field in (page = 3 ? ["bg", "txt", "name"] : ["bg", "txt"]) {
            ctrl := controls[field], pref := CPOverlayPreference(page = 3 ? "Translator" : "Explainer", field)
            style := DllCall("user32\GetWindowLongPtr", "ptr", ctrl.Hwnd, "int", -16, "ptr")
            DesktopAssert((style & 0xF) = 0xB, context ": preserves owner-drawn color button " page ":" field)
            DesktopAssert(CPDesktop["paint"][ctrl.Hwnd]["color"] = pref["value"], context ": current preview RGB")
            DesktopAssert(ctrl.Text = pref["title"] ": #" StrUpper(pref["value"]), context ": current accessible label")
        }
    }
}

TestOverlayColorPixels(page) {
    global CPDesktop
    for field in (page = 3 ? ["bg", "txt", "name"] : ["bg", "txt"]) {
        ctrl := CPDesktop["overlayPages"][page][field]
        rect := Buffer(16, 0)
        DllCall("user32\GetClientRect", "ptr", ctrl.Hwnd, "ptr", rect)
        width := NumGet(rect, 8, "int"), height := NumGet(rect, 12, "int")
        windowDc := DllCall("user32\GetDC", "ptr", ctrl.Hwnd, "ptr")
        dc := DllCall("gdi32\CreateCompatibleDC", "ptr", windowDc, "ptr")
        bitmap := DllCall("gdi32\CreateCompatibleBitmap", "ptr", windowDc, "int", width, "int", height, "ptr")
        previous := DllCall("gdi32\SelectObject", "ptr", dc, "ptr", bitmap, "ptr")
        try {
            ; Ask the real native button to paint; do not invoke the painter
            ; directly, which would miss loss of the owner-drawn button style.
            SendMessage(0x0317, dc, 0x1C, ctrl.Hwnd) ; WM_PRINT: client, background and children
            scale := GetWindowDPI(ctrl.Hwnd) / 96
            chip := DllCall("gdi32\GetPixel", "ptr", dc, "int", Round(21 * scale), "int", height // 2, "uint")
            background := DllCall("gdi32\GetPixel", "ptr", dc, "int", width - Round(20 * scale), "int", height // 2, "uint")
            DesktopAssert(chip = CPColorRef(CPDesktop["paint"][ctrl.Hwnd]["color"]), "Native paint displays the selected color square")
            DesktopAssert(background = CPColorRef(CPDesktopPalette()["field"]), "Native paint preserves the themed field background")
        } finally {
            DllCall("gdi32\SelectObject", "ptr", dc, "ptr", previous)
            DllCall("gdi32\DeleteObject", "ptr", bitmap)
            DllCall("gdi32\DeleteDC", "ptr", dc)
            DllCall("user32\ReleaseDC", "ptr", ctrl.Hwnd, "ptr", windowDc)
        }
    }
}

TestOverlayColorPicker(initial) {
    global TestColorResult, TestColorInitial
    TestColorInitial := initial
    return TestColorResult
}

TestOverlayColorPersist() {
    global TestColorSaves
    TestColorSaves += 1
}

TestOverlayColorSendTheme() {
    global TestColorSends
    TestColorSends += 1
}

; Tests use isolated offscreen HWNDs: no real pointer/controller input, settings,
; foreground activation, or provider requests.
TestDesktopFocusInput() {
    global ui, tab, CPDesktop, controlDarkMode, CPFocusVisualNavHwnd, CPFocusVisualHwnd
    global ddlEProv, chkGuess, eGemini
    originalDark := controlDarkMode
    for dark in [1, 0] {
        controlDarkMode := dark
        for width in [900, 1120, 1400] {
            ui.Show("NA x-9000 y-9000 w" width " h900")
            for page in [1, 2, 4] {
                CPDesktopNavigate(page)
                CPDesktopLayout(ui, 0, width, 900)
                CPApplyControlPanelTheme()
                controls := CPDesktopPage(page)["controls"]
                link := controls["models"]
                label := controls[page = 2 ? "language" : "prompt"]
                link.GetPos(&lx, &ly, &lw, &lh)
                label.GetPos(&tx, &ty, &tw, &th)
                DesktopAssert(lx + lw <= tx || tx + tw <= lx || ly + lh <= ty || ty + th <= ly,
                    "Model link does not overlap the following label: page " page ", width " width)
                CPSetFocusHwnd(link.Hwnd)
                UpdateCPFocusRing()
                edges := DesktopFocusBorderPixels(link)
                for edge, count in edges
                    DesktopAssert(count > 0, "Full-window paint preserves " edge " model-link border: page " page ", width " width ", theme " dark)
            }
        }
    }
    controlDarkMode := 1
    CPDesktopNavigate(4)
    ui.Show("NA x-9000 y-9000 w900 h760")
    CPDesktopLayout(ui, 0, 900, 760)
    CPApplyControlPanelTheme()
    link := CPDesktop["explanationPage"]["models"]
    CPSetFocusHwnd(link.Hwnd)
    UpdateCPFocusRing()
    CPDesktopFocusInputMessage(0, 0, 0x200, link.Hwnd)
    DesktopAssert(CPDesktopNavigationFocusVisible(link.Hwnd), "Stationary synthetic mouse move does not hide keyboard focus")
    position := CPDesktop["focusPointer"]
    CPDesktopPointerFocusInput(true, [position[1] + 1, position[2]])
    DesktopAssert(CPDesktopNavigationFocusVisible(link.Hwnd), "Tiny pointer jitter does not hide keyboard focus")
    CPDesktopPointerFocusInput(true, [position[1] + 10, position[2]])
    DesktopAssert(!CPDesktopNavigationFocusVisible(link.Hwnd), "Real pointer movement hides navigation cues")
    DesktopAssert(CPFocusedHwnd() = link.Hwnd, "Mouse mode retains actual keyboard focus")
    UpdateCPFocusRing()
    DesktopAssert(!CPFocusVisualNavHwnd && !CPFocusVisualHwnd, "Focus polling cannot restore a mouse-hidden frame")
    for edge, count in DesktopFocusBorderPixels(link)
        DesktopAssert(count = 0, "Pointer mode removes rendered model-link " edge " border")

    for msg in [0x201, 0x204, 0x207, 0x20B, 0x20A, 0x20E, 0xA1] {
        CPSetFocusHwnd(link.Hwnd)
        CPDesktopFocusInputMessage(0, 0, msg, link.Hwnd)
        DesktopAssert(!CPDesktopNavigationFocusVisible(link.Hwnd), "Click/wheel hides navigation cues: " msg)
    }
    ; Exercise registered message routing too, rather than just helper calls.
    CPSetFocusHwnd(link.Hwnd)
    PostMessage(0x204, 0, 0, link.Hwnd)
    Sleep(20)
    DesktopAssert(!CPDesktopNavigationFocusVisible(link.Hwnd), "Posted mouse input reaches the registered mode handler")
    CPDesktopKeyboardTab()
    DesktopAssert(CPDesktopNavigationFocusVisible(link.Hwnd), "Tab restores navigation cues")
    ddlEProv.Focus()
    CPDesktopTabFocusChanged()
    DesktopAssert(CPFocusVisualNavHwnd = ddlEProv.Hwnd, "Tab completion follows the new native focus target")
    CPDesktopFocusInputMessage(0, 0, 0x201, ddlEProv.Hwnd)
    CPControllerDispatchNavigation("Down", ui.Hwnd)
    DesktopAssert(CPDesktopNavigationFocusVisible(ui.Hwnd), "Controller D-pad restores cues after mouse input")
    CPDesktopFocusInputMessage(0, 0, 0x201, CPFocusedHwnd())
    CPNavMove("Up")
    DesktopAssert(CPDesktopNavigationFocusVisible(ui.Hwnd), "Arrow navigation restores cues after mouse input")

    CPSetFocusHwnd(CPDesktop["chrome"]["audioPage"].Hwnd)
    CPNavActivate("Enter")
    Sleep(30)
    DesktopAssert(tab.Value = 2 && CPDesktopNavigationFocusVisible(ui.Hwnd),
        "Controller/Enter activation retains cues; synthesized BM_CLICK is not mouse input")
    mouseButton := CPDesktop["chrome"]["explanation"]
    PostMessage(0x201, 1, 0x00100010, mouseButton.Hwnd)
    PostMessage(0x202, 0, 0x00100010, mouseButton.Hwnd)
    Sleep(30)
    DesktopAssert(tab.Value = 4 && !CPDesktopNavigationFocusVisible(ui.Hwnd),
        "Native mouse click navigates without bringing back the blue focus frame")
    CPDesktopFocusInputMessage(0, 0, 0x201, CPFocusedHwnd())
    CPDesktopNavigate(4)
    DesktopAssert(!CPDesktopNavigationFocusVisible(ui.Hwnd), "Mouse page changes do not restore navigation cues")

    ; A different top-level window must not change the main window's mode.
    other := Gui(), otherButton := other.AddButton(, "Other window")
    CPSetFocusHwnd(link.Hwnd)
    CPDesktopFocusInputMessage(0, 0, 0x201, otherButton.Hwnd)
    DesktopAssert(CPDesktopNavigationFocusVisible(ui.Hwnd), "Unrelated/Big Box windows do not alter desktop input mode")
    other.Destroy()
    TestCapture("explanation-focus-complete.png", Round(900 * GetWindowDPI(ui.Hwnd) / 96), Round(760 * GetWindowDPI(ui.Hwnd) / 96))
    for spec in [[1, chkGuess], [4, ddlEProv], [9, eGemini]] {
        CPDesktopNavigate(spec[1])
        CPDesktopLayout(ui, 0, 900, 760)
        target := spec[2], enabled := target.Enabled
        target.Enabled := true
        CPSetFocusHwnd(target.Hwnd)
        Sleep(20) ; Allow native Edit Focus callbacks to update the parent-painted frame.
        CPEnsureFocusedControlVisible()
        UpdateCPFocusRing()
        painted := CPDesktop["inputFrames"].Get(target.Hwnd, target)
        DesktopAssert(DesktopFocusBorderPixels(painted, true)["all"] > 0,
            "Keyboard mode paints the focused " target.Type " outline")
        CPDesktopFocusInputMessage(0, 0, 0x201, target.Hwnd)
        UpdateCPFocusRing()
        DesktopAssert(DesktopFocusBorderPixels(painted, true)["all"] = 0,
            "Mouse mode removes the focused " target.Type " outline")
        DesktopAssert(CPFocusedHwnd() = target.Hwnd, "Mouse mode retains focus for " target.Type)
        target.Enabled := enabled
    }
    controlDarkMode := originalDark
    CPApplyControlPanelTheme()
}

DesktopFocusBorderPixels(ctrl, includeInterior := false) {
    global ui
    client := Buffer(16, 0)
    DllCall("user32\GetClientRect", "ptr", ui.Hwnd, "ptr", client)
    w := NumGet(client, 8, "int"), h := NumGet(client, 12, "int")
    windowDc := DllCall("user32\GetDC", "ptr", ui.Hwnd, "ptr")
    dc := DllCall("gdi32\CreateCompatibleDC", "ptr", windowDc, "ptr")
    bitmap := DllCall("gdi32\CreateCompatibleBitmap", "ptr", windowDc, "int", w, "int", h, "ptr")
    oldBitmap := DllCall("gdi32\SelectObject", "ptr", dc, "ptr", bitmap, "ptr")
    try {
        DllCall("user32\PrintWindow", "ptr", ui.Hwnd, "ptr", dc, "uint", 1)
        rect := Buffer(16)
        DllCall("user32\GetWindowRect", "ptr", ctrl.Hwnd, "ptr", rect)
        DllCall("user32\MapWindowPoints", "ptr", 0, "ptr", ui.Hwnd, "ptr", rect, "uint", 2)
        l := NumGet(rect, 0, "int"), t := NumGet(rect, 4, "int")
        r := NumGet(rect, 8, "int"), b := NumGet(rect, 12, "int")
        strip := Max(3, Ceil(3 * GetWindowDPI(ui.Hwnd) / 96))
        color := CPColorRef(CPDesktopPalette()["link"])
        if includeInterior
            return Map("all", DesktopCountColor(dc, color, l, t, r, b))
        ; Use center portions of each edge, excluding corners and link text.
        cx := Floor((l + r) / 2), cy := Floor((t + b) / 2)
        return Map("top", DesktopCountColor(dc, color, cx - 10, t, cx + 10, t + strip),
            "bottom", DesktopCountColor(dc, color, cx - 10, b - strip, cx + 10, b),
            "left", DesktopCountColor(dc, color, l, cy - 3, l + Max(2, Round(GetWindowDPI(ui.Hwnd) / 96)), cy + 3),
            "right", DesktopCountColor(dc, color, r - strip, cy - 3, r, cy + 3))
    } finally {
        DllCall("gdi32\SelectObject", "ptr", dc, "ptr", oldBitmap)
        DllCall("gdi32\DeleteObject", "ptr", bitmap)
        DllCall("gdi32\DeleteDC", "ptr", dc)
        DllCall("user32\ReleaseDC", "ptr", ui.Hwnd, "ptr", windowDc)
    }
}

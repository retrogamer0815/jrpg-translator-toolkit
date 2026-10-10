; Real constructors and event wiring, isolated settings, no live Study data.
TestStudyCaptionControls() {
    global controlDarkMode, CPDesktop, ui, TestCaptionEvents := Map(), TestCaptionPaints := Map()
    OnMessage(0x200, TestCaptionObserve)
    OnMessage(0x2A3, TestCaptionObserve)
    OnMessage(0x2B, TestCaptionDrawObserve, -1)
    try {
        for dark in [1, 0] {
            controlDarkMode := dark
            CPRefreshThemeBrushes()
            CPDesktopTheme()
            ui.Show("NA x-12000 y-12000 w1120 h760")
            CPDesktopRelayout()
            for action in ["minimize", "maximize", "close"] {
                hwnd := CPDesktop["chrome"][action].Hwnd
                TestCaptionEvents.Clear()
                TestCaptionPaints.Clear()
                PostMessage(0x200, 0, 8 | (8 << 16), hwnd)
                Sleep(25)
                DesktopAssert(TestCaptionEvents.Get(hwnd "-512", -1) = hwnd, "Main caption still highlights: " action)
                DesktopAssert(TestCaptionEvents[hwnd "-pixel"] = CPColorRef(action = "close" ? "B83B48" : CPDesktopPalette()["selected"]),
                    "Main caption keeps the same hover colors: " action)
                PostMessage(0x2A3, 0, 0, hwnd)
                Sleep(25)
                DesktopAssert(CPDesktop["hover"] = 0, "Main caption hover clears: " action)
            }
            for kind in ["library", "reader"] {
                s := kind = "library" ? TestDesktopStudyLibrary() : TestDesktopStudyReader()
                s["suspend"] := true ; Native list events must not issue bridge requests.
                g := s["gui"], d := s["desktop"], c := d["chrome"], root := g.Hwnd
                g.Opt("+E0x08000000")
                CPSetWindowCloaked(root, true) ; Maximize must not cover the user's desktop.
                g.Show("NA x-12000 y-12000 w1200 h800")
                StudyDesktopResize(s, 1200, 800)
                CPApplyOwnedDialogTheme(g)
                ; Study registers its own prepend draw handler at construction.
                ; Place the read-only probe before it so handled draws are seen.
                OnMessage(0x2B, TestCaptionDrawObserve, 0)
                OnMessage(0x2B, TestCaptionDrawObserve, -1)
                Sleep(50)
                for action in ["minimize", "maximize", "close"] {
                    ctrl := c[action], hwnd := ctrl.Hwnd
                    DesktopAssert(ctrl.Enabled && DesktopShown(ctrl), kind " caption enabled and visible: " action)
                    DesktopAssert(CPDesktopCaptionContext(hwnd) = d, kind " caption resolves to its own window: " action)
                    ctrl.GetPos(&x, &y, &w, &h)
                    point := Buffer(8), scale := GetWindowDPI(root) / 96
                    NumPut("int", Round((x + w / 2) * scale), "int", Round((y + h / 2) * scale), point)
                    child := DllCall("user32\ChildWindowFromPointEx", "ptr", root, "int64", NumGet(point, 0, "int64"), "uint", 7, "ptr")
                    DesktopAssert(child = hwnd, kind " caption is not covered: " action)
                    DllCall("user32\ClientToScreen", "ptr", root, "ptr", point)
                    packed := (NumGet(point, 0, "int") & 0xFFFF) | ((NumGet(point, 4, "int") & 0xFFFF) << 16)
                    DesktopAssert(SendMessage(0x84, 0, packed, root) = 1, kind " caption is clickable client area: " action)
                    TestCaptionEvents.Clear()
                    TestCaptionPaints.Clear()
                    PostMessage(0x200, 0, 8 | (8 << 16), hwnd)
                    Sleep(25)
                    ; Observe during dispatch: TrackMouseEvent immediately queues
                    ; leave because these test windows are deliberately offscreen.
                    DesktopAssert(TestCaptionEvents.Get(hwnd "-512", -1) = hwnd, kind " native mouse move highlights: " action)
                    DesktopAssert(TestCaptionEvents.Get(hwnd "-painted", false), kind " hover repaints without PrintWindow: " action)
                    PostMessage(0x2A3, 0, 0, hwnd)
                    Sleep(25)
                    DesktopAssert(TestCaptionEvents.Get(hwnd "-675", -1) = 0, kind " native mouse leave clears: " action)
                    DesktopAssert(TestCaptionEvents[hwnd "-pixel"] = CPColorRef(action = "close" ? "B83B48" : CPDesktopPalette()["selected"]),
                        kind " hover visibly paints the expected background: " action)
                }
                ; Deterministic queued-leave ordering without a real mouse move.
                Critical "On"
                try {
                    CPDesktopCaptionHover(0, 0, 0x200, c["minimize"].Hwnd)
                    CPDesktopCaptionHover(0, 0, 0x200, c["maximize"].Hwnd)
                    CPDesktopCaptionLeave(0, 0, 0x2A3, c["minimize"].Hwnd)
                    DesktopAssert(d["hover"] = c["maximize"].Hwnd, kind " stale leave preserves the next button highlight")
                    CPDesktopCaptionLeave(0, 0, 0x2A3, c["maximize"].Hwnd)
                    DesktopAssert(d["hover"] = 0 && CPDesktop["hover"] = 0, kind " hover stays local to its window")
                } finally Critical "Off"
                PostMessage(0xF5, 0, 0, c["minimize"].Hwnd) ; BM_CLICK, not a direct action call.
                Sleep(80)
                DesktopAssert(DllCall("user32\IsIconic", "ptr", root), kind " minimize click works")
                g.Restore()
                PostMessage(0xF5, 0, 0, c["maximize"].Hwnd)
                Sleep(80)
                DesktopAssert(DllCall("user32\IsZoomed", "ptr", root), kind " maximize click works")
                PostMessage(0xF5, 0, 0, c["maximize"].Hwnd)
                Sleep(80)
                DesktopAssert(!DllCall("user32\IsZoomed", "ptr", root), kind " second maximize click restores")
                PostMessage(0xF5, 0, 0, c["close"].Hwnd)
                Sleep(80)
                DesktopAssert(!DllCall("user32\IsWindow", "ptr", root), kind " close click reaches the existing close handler")
                DesktopAssert(!StudyDesktopRegistry().Has(root), kind " close releases the window registry")
            }
        }
    } finally {
        OnMessage(0x200, TestCaptionObserve, 0)
        OnMessage(0x2A3, TestCaptionObserve, 0)
        OnMessage(0x2B, TestCaptionDrawObserve, 0)
    }
}

TestCaptionObserve(wParam, lParam, msg, hwnd) {
    global TestCaptionEvents, TestCaptionPaints, ui
    d := CPDesktopCaptionContext(hwnd)
    if !IsObject(d)
        return
    TestCaptionEvents[hwnd "-" msg] := d["hover"]
    if msg != 0x200
        return
    ; Check the real hover-triggered draw BEFORE PrintWindow forces a new one.
    ; A correct hover flag alone does not prove the visible button repainted.
    TestCaptionEvents[hwnd "-painted"] := TestCaptionPaints.Get(hwnd, 0) > 0
    ; Render the real button into an isolated bitmap and sample away from the
    ; glyph/border. This checks visible feedback, not only an internal flag.
    dc := DllCall("gdi32\CreateCompatibleDC", "ptr", 0, "ptr")
    screenDC := DllCall("user32\GetDC", "ptr", 0, "ptr")
    root := d.Get("frameHwnd", ui.Hwnd), rect := Buffer(16)
    DllCall("user32\GetClientRect", "ptr", root, "ptr", rect)
    bmp := DllCall("gdi32\CreateCompatibleBitmap", "ptr", screenDC, "int", NumGet(rect, 8, "int"), "int", NumGet(rect, 12, "int"), "ptr")
    DllCall("user32\ReleaseDC", "ptr", 0, "ptr", screenDC)
    old := DllCall("gdi32\SelectObject", "ptr", dc, "ptr", bmp, "ptr")
    try {
        DllCall("user32\PrintWindow", "ptr", root, "ptr", dc, "uint", 1)
        d["paint"][hwnd]["ctrl"].GetPos(&x, &y)
        scale := GetWindowDPI(root) / 96
        TestCaptionEvents[hwnd "-pixel"] := DllCall("gdi32\GetPixel", "ptr", dc, "int", Round(x * scale) + 6, "int", Round(y * scale) + 6, "uint")
    } finally {
        DllCall("gdi32\SelectObject", "ptr", dc, "ptr", old)
        DllCall("gdi32\DeleteObject", "ptr", bmp)
        DllCall("gdi32\DeleteDC", "ptr", dc)
    }
}

TestCaptionDrawObserve(wParam, lParam, msg, root) {
    global TestCaptionPaints
    if !lParam || NumGet(lParam, 0, "uint") != 4
        return
    hwnd := NumGet(lParam, A_PtrSize = 8 ? 24 : 20, "ptr")
    d := CPDesktopCaptionContext(hwnd)
    if IsObject(d) && d["hover"] = hwnd
        TestCaptionPaints[hwnd] := TestCaptionPaints.Get(hwnd, 0) + 1
}

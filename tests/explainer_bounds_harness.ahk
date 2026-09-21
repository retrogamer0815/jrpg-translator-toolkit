#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
global TestAssertions := 0, TestFound := true, TestReadFails := false
global TestWrites := 0, TestWriteFailsAt := 0, TestGeometry := [40, 60, 640, 480]
global TestStored := Map(), TestWindowTitle := "JRPG bounds test " DllCall("GetCurrentProcessId")
global ewX := 10, ewY := 20, ewW := 300, ewH := 200
global ew_lastX := 10, ew_lastY := 20, ew_lastW := 300, ew_lastH := 200
global ew_bounds_watch_running := false, iniPath := A_ScriptDir "\test-settings.ini"
OnError(TestUnhandledError)
try {
    TestBoundsPersistence()
    TestRealWindow()
    FileAppend("PASS: " TestAssertions " Explainer bounds assertions.`n", "*")
    ExitApp(0)
} catch as err {
    TestUnhandledError(err)
}
ExitApp(1)

TestUnhandledError(err, *) {
    FileAppend("FAIL: " err.Message " | " err.Extra " | " err.What " | line " err.Line "`n", "*")
    ExitApp(1)
}

TestAssert(ok, message) {
    global TestAssertions
    TestAssertions += 1
    if !ok
        throw Error(message)
}

TestWinExist(title) {
    global TestFound
    TestAssert(title = "Explainer" && A_TitleMatchMode = 3, "Lookup must use the exact Explainer caption")
    return TestFound ? 123 : 0
}

TestWinGetPos(&x, &y, &w, &h, title) {
    global TestReadFails, TestGeometry
    TestAssert(title = "ahk_id 123", "Geometry must use the captured window handle")
    if TestReadFails
        throw TargetError("Synthetic overlay closed between lookup and geometry read")
    x := TestGeometry[1], y := TestGeometry[2], w := TestGeometry[3], h := TestGeometry[4]
}

TestIniWrite(value, file, section, key) {
    global TestWrites, TestWriteFailsAt, TestStored, iniPath
    TestWrites += 1
    if TestWrites = TestWriteFailsAt
        throw OSError(5, "Synthetic denied INI write")
    TestAssert(file = iniPath && section = "explainer_bounds", "Writes stay in the bounds section")
    TestStored[key] := value
}

TestBounds(expected, message) {
    global ewX, ewY, ewW, ewH, ew_lastX, ew_lastY, ew_lastW, ew_lastH, TestStored
    TestAssert(ewX = expected[1] && ewY = expected[2] && ewW = expected[3] && ewH = expected[4], message " (public bounds)")
    TestAssert(ew_lastX = expected[1] && ew_lastY = expected[2] && ew_lastW = expected[3] && ew_lastH = expected[4], message " (saved cache)")
    for i, key in ["x", "y", "w", "h"]
        TestAssert(TestStored.Get(key, "missing") = expected[i], message " (INI " key ")")
}

TestBoundsPersistence() {
    global ewX, ewY, ewW, ewH, ew_lastX, ew_lastY, ew_lastW, ew_lastH
    global TestWrites, TestWriteFailsAt, TestGeometry, TestFound, TestReadFails, ew_bounds_watch_running

    ; This first call reproduces the reported unassigned `changed` exception on
    ; the original production body, with already populated persisted bounds.
    SetTitleMatchMode 2
    SaveExplainerBoundsIfChanged()
    TestBounds(TestGeometry, "Move/resize with previously saved bounds")
    TestAssert(A_TitleMatchMode = 2, "The caller's title matching mode is restored")
    writes := TestWrites
    SaveExplainerBoundsIfChanged()
    TestAssert(TestWrites = writes, "Unchanged geometry must not rewrite settings")

    ewX := "", ewY := "", ewW := "", ewH := ""
    ew_lastX := "", ew_lastY := "", ew_lastW := "", ew_lastH := ""
    TestGeometry := [0, -80, 800, 600]
    SaveExplainerBoundsIfChanged()
    TestBounds(TestGeometry, "First save accepts zero and negative screen coordinates")
    ; Cover mixed initialization too, without overwriting existing values merely
    ; because some other public coordinate is empty.
    ewW := ""
    SaveExplainerBoundsIfChanged()
    TestBounds(TestGeometry, "Unchanged geometry seeds missing public bounds")

    Loop 120 {
        TestGeometry := [-A_Index, A_Index, 400 + A_Index, 300 + A_Index]
        SaveExplainerBoundsIfChanged()
        TestBounds(TestGeometry, "Repeated move/resize " A_Index)
    }
    saved := TestGeometry.Clone(), writes := TestWrites
    TestFound := false
    SaveExplainerBoundsIfChanged()
    TestAssert(TestWrites = writes && A_TitleMatchMode = 2, "Absent overlay is ignored and matching mode restored")
    TestFound := true, TestReadFails := true
    SaveExplainerBoundsIfChanged()
    TestAssert(TestWrites = writes, "An overlay closing during a read cannot persist zero bounds")
    TestBounds(saved, "Read failure leaves valid geometry intact")
    TestReadFails := false
    for invalid in [[0, 0, 0, 0], [10, 20, -1, 500], [10, 20, 500, 0], ["", 20, 500, 400]] {
        TestGeometry := invalid
        SaveExplainerBoundsIfChanged()
        TestAssert(TestWrites = writes, "Invalid/incomplete geometry is ignored")
        TestBounds(saved, "Invalid geometry preserves saved bounds")
    }

    TestGeometry := [250, 300, 700, 900]
    TestWriteFailsAt := TestWrites + 3
    SaveExplainerBoundsIfChanged()
    TestAssert(ew_lastX = saved[1] && ew_lastY = saved[2] && ew_lastW = saved[3] && ew_lastH = saved[4], "Partial write must not advance the saved cache")
    TestWriteFailsAt := 0
    SaveExplainerBoundsIfChanged()
    TestBounds(TestGeometry, "Next tick retries the whole geometry after a failed write")

    TestGeometry := [300, 350, 750, 950]
    StartExplainerBoundsWatcher()
    StartExplainerBoundsWatcher()
    TestAssert(ew_bounds_watch_running, "Watcher starts idempotently")
    Sleep 850
    TestBounds(TestGeometry, "Periodic timer persists updated bounds")
    StopExplainerBoundsWatcher()
    StopExplainerBoundsWatcher()
    TestAssert(!ew_bounds_watch_running, "Watcher stops idempotently")
    writes := TestWrites
    TestGeometry := [320, 370, 800, 980]
    Sleep 800
    TestAssert(TestWrites = writes, "Stopped watcher no longer writes")
    SetTimer(SaveExplainerBoundsIfChanged, -1)
    Sleep 80
    TestBounds(TestGeometry, "Adjustment-completion one-shot timer persists updated bounds")
}

TestRealWindow() {
    global TestWindowTitle, iniPath, ewX, ewY, ewW, ewH, ew_lastX, ew_lastY, ew_lastW, ew_lastH
    ; Never activate, move, or query the user's overlay; use a private hidden GUI.
    oldDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    testGui := Gui("+Resize", TestWindowTitle)
    try {
        testGui.Show("Hide x80 y90 w440 h320")
        for bounds in [[80, 90, 460, 360], [0, 0, 640, 480], [-200, -120, 350, 250], [100, 120, 500, 700]] {
            WinMove(bounds[1], bounds[2], bounds[3], bounds[4], "ahk_id " testGui.Hwnd)
            WinGetPos &x, &y, &w, &h, "ahk_id " testGui.Hwnd
            TestRealBoundsSave()
            TestAssert(ewX = x && ewY = y && ewW = w && ewH = h, "Real window move/resize updates public geometry")
            for i, key in ["x", "y", "w", "h"]
                TestAssert(Integer(IniRead(iniPath, "explainer_bounds", key)) = [x, y, w, h][i], "Real INI round trip " key)
        }
        testGui.Destroy()
        TestRealBoundsSave()
        TestAssert(ew_lastX = x && ew_lastY = y && ew_lastW = w && ew_lastH = h, "Closed real window preserves its last bounds")
    } finally {
        try testGui.Destroy()
        DetectHiddenWindows oldDetect
    }
}

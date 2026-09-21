#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
global TestAssertions := 0, TestFailures := 0, TestReal := false
global TestAlive := Map(), TestCreated := 0, TestDeleted := 0, TestNextHandle := 200
global Overlay := TestControl(1), RectOuter := TestControl(2), RectInner := TestControl(3)
global RectPanel := TestControl(4), OutputCtl := TestControl(5)
global BOX_BG := "202020", BDR_IN := "303030", BDR_OUT := "404040"
global hBrushOuter := 0, hBrushInner := 0, hBrushPanel := 0, hBrushEdit := 0
OnError(TestUnhandledError)
for test in [TestModels, TestRefresh, TestErase] {
    try test.Call()
    catch as err {
        TestFailures += 1
        FileAppend("FAIL: " test.Name ": " err.Message "`n", "*")
    }
}
if TestFailures
    ExitApp(1)
try TestRealDrawingResources()
catch as err
    TestUnhandledError(err)
FileAppend("PASS: " TestAssertions " conditional regression assertions.`n", "*")
ExitApp(0)

TestUnhandledError(err, *) {
    FileAppend("FAIL: " err.Message " | " err.Extra " | line " err.Line "`n", "*")
    ExitApp(1)
}
TestAssert(ok, message) {
    global TestAssertions
    TestAssertions += 1
    if !ok
        throw Error(message)
}
TestModels() {
    for provider in ["OpenAI", "Gemini", ""]
        for openai in ["custom-openai", ""]
            for gemini in ["custom-gemini", ""]
                TestModelFallback(provider, openai, gemini)
}
TestModelFallback(explainProvider, explainOpenAIModel, explainGeminiModel) {
    defExplainProvider := "Gemini", defExplainOpenAIModel := "default-openai", defExplainGeminiModel := "default-gemini"
    expectedProvider := explainProvider != "" ? explainProvider : defExplainProvider
    expectedOpenai := explainOpenAIModel != "" ? explainOpenAIModel : defExplainOpenAIModel
    expectedGemini := explainGeminiModel != "" ? explainGeminiModel : defExplainGeminiModel
    ; @MODEL_FALLBACKS@
    reachedNextStatement := true
    TestAssert(explainProvider = expectedProvider, "Valid saved explanation provider must not be reset")
    TestAssert(explainOpenAIModel = expectedOpenai, "Only an empty OpenAI selection gets the fallback")
    TestAssert(explainGeminiModel = expectedGemini, "Only an empty Gemini selection gets the fallback")
    TestAssert(IsSet(reachedNextStatement), "Initialization after the fallbacks always executes")
}
class TestControl {
    __New(hwnd) {
        this.Hwnd := hwnd
    }
    Opt(*) {
    }
}
TestResetBrushes() {
    global hBrushOuter, hBrushInner, hBrushPanel, hBrushEdit, TestAlive, TestCreated, TestDeleted
    hBrushOuter := 0, hBrushInner := 0, hBrushPanel := 0, hBrushEdit := 0
    TestAlive := Map(), TestCreated := 0, TestDeleted := 0
}
TestRefresh() {
    global hBrushOuter, hBrushInner, hBrushPanel, hBrushEdit, TestAlive, TestCreated, TestDeleted
    TestResetBrushes()
    RefreshAllBg()
    TestAssert(TestAlive.Count = 4 && TestCreated = 4 && TestDeleted = 0, "First theme setup creates four brushes")
    RefreshAllBg()
    TestAssert(TestAlive.Count = 4 && TestCreated = 8 && TestDeleted = 4, "Theme refresh must release all four old brushes")
    Loop 100
        RefreshAllBg()
    TestAssert(TestAlive.Count = 4 && TestCreated - TestDeleted = 4, "Repeated theme changes keep a bounded resource count")
    ; The original initializer also tried to cover previously unset globals.
    TestResetBrushes()
    hBrushOuter := unset, hBrushInner := unset, hBrushPanel := unset, hBrushEdit := unset
    RefreshAllBg()
    TestAssert(TestAlive.Count = 4, "Unset brush globals are initialized without an exception")
}
TestErase() {
    global hBrushOuter, hBrushInner, hBrushPanel, hBrushEdit, TestAlive, TestCreated, TestDeleted
    TestResetBrushes()
    ; A missing brush is allocated once, then reused for every later repaint.
    for ctrl in [RectPanel, RectInner, RectOuter, OutputCtl] {
        createdBefore := TestCreated
        TestAssert(EraseAnyBg(500, 0, 0x0014, ctrl.Hwnd) = 1, "Missing brush is created and used to erase")
        TestAssert(TestCreated = createdBefore + 1, "First erase allocates exactly one brush")
        Loop 100
            TestAssert(EraseAnyBg(500, 0, 0x0014, ctrl.Hwnd) = 1, "Cached brush still paints successfully")
        TestAssert(TestCreated = createdBefore + 1, "Repaints must reuse their existing brush")
    }
    TestAssert(TestAlive.Count = 4 && TestDeleted = 0, "Painting neither leaks nor deletes cached brushes")
    createdBefore := TestCreated
    EraseAnyBg(500, 0, 0x0014, 999)
    EraseAnyBg(500, 0, 0x0014, Overlay.Hwnd)
    TestAssert(TestCreated = createdBefore, "Unrelated/top-level window does not allocate child brushes")
}
TestMakeBrush(color) {
    global TestNextHandle, TestCreated, TestAlive, TestReal
    handle := TestReal ? DllCall("gdi32\CreateSolidBrush", "uint", Integer("0x" color), "ptr") : ++TestNextHandle
    TestAssert(handle != 0, "Drawing resource allocation succeeds")
    TestCreated += 1
    TestAlive[handle] := color
    return handle
}
ToHex6(color) => color
SetRichEditBg(*) {
    ; Text formatting is outside these tests; no user overlay is loaded.
}
Dbg(*) {
}
TestDllCall(name, args*) {
    global TestDeleted, TestAlive, TestReal
    if name = "gdi32\DeleteObject" {
        TestAssert(TestAlive.Has(args[2]), "Only a currently owned brush may be deleted")
        TestAlive.Delete(args[2])
        TestDeleted += 1
    } else if name = "User32\FillRect" {
        TestAssert(TestAlive.Has(args[6]), "Painting uses a live cached brush")
    }
    if TestReal
        return DllCall(name, args*)
    return 1
}
TestGdiCount() => DllCall("user32\GetGuiResources", "ptr", DllCall("GetCurrentProcess", "ptr"), "uint", 0, "uint")
TestRealDrawingResources() {
    global Overlay, RectOuter, RectInner, RectPanel, OutputCtl, TestReal, TestAlive, TestCreated, TestDeleted
    global BOX_BG
    TestResetBrushes()
    ; Real GDI calls and private hidden controls, never the user's running app.
    Overlay := Gui("+Resize", "JRPG conditional regression test")
    RectOuter := Overlay.AddText("x0 y0 w300 h200")
    RectInner := Overlay.AddText("x5 y5 w290 h190")
    RectPanel := Overlay.AddText("x10 y10 w280 h180")
    OutputCtl := Overlay.AddText("x15 y15 w270 h170", "Synthetic overlay")
    Overlay.Show("Hide w300 h200")
    dc := DllCall("user32\GetDC", "ptr", Overlay.Hwnd, "ptr")
    TestAssert(dc != 0, "Private hidden window provides a drawing context")
    TestReal := true
    try {
        ; Allow native controls to initialize both theme colors and process
        ; their pending startup messages before measuring resource growth.
        Loop 20 {
            BOX_BG := Mod(A_Index, 2) ? "202020" : "404040"
            RefreshAllBg()
            for ctrl in [RectOuter, RectInner, RectPanel, OutputCtl]
                EraseAnyBg(dc, 0, 0x0014, ctrl.Hwnd)
        }
        Sleep 50
        DllCall("gdi32\GdiFlush")
        baseline := TestGdiCount()
        Loop 200 {
            BOX_BG := Mod(A_Index, 2) ? "202020" : "404040"
            RefreshAllBg()
            Loop 5
                for ctrl in [RectOuter, RectInner, RectPanel, OutputCtl]
                    EraseAnyBg(dc, 0, 0x0014, ctrl.Hwnd)
        }
        Sleep 50
        DllCall("gdi32\GdiFlush")
        afterStress := TestGdiCount()
        FileAppend("GDI stress: 200 theme changes / 4000 child repaints; resources " baseline " -> " afterStress "`n", "*")
        TestAssert(TestAlive.Count = 4 && TestCreated - TestDeleted = 4, "Real theme/repaint stress retains only four brushes")
        TestAssert(afterStress <= baseline + 2, "Real GDI count must not grow during repeated refresh/repaint")
    } finally {
        for handle in TestAlive
            DllCall("gdi32\DeleteObject", "ptr", handle)
        TestAlive.Clear()
        DllCall("user32\ReleaseDC", "ptr", Overlay.Hwnd, "ptr", dc)
        Overlay.Destroy()
        TestReal := false
    }
}

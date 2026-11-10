'==============================================================================
' WebActions.vbs
'   Generic web wrappers: existence check -> highlight -> action -> verify -> log.
'   All waits use DEFAULT_TIMEOUT from .env (seconds, default 10).
'
'   Each wrapper returns True/False (Function) so tests can branch, and logs
'   the step through ReportUtil so the test-case status is tracked.
'==============================================================================

Function DefaultTimeout()
    DefaultTimeout = GetConfigInt("DEFAULT_TIMEOUT", 10)
End Function

Function ObjectReady(obj, stepName)
    Dim ok
    ok = False
    On Error Resume Next
    ok = obj.Exist(DefaultTimeout())
    If Err.Number <> 0 Then ok = False
    Err.Clear
    If ok Then obj.Highlight
    Err.Clear
    On Error GoTo 0
    If Not ok Then LogFail stepName, "Object not found within " & DefaultTimeout() & " sec"
    ObjectReady = ok
End Function


'------------------------------------------------------------------------------
' WebSet - UFT Set, falls back to keyboard typing, verifies the value
'------------------------------------------------------------------------------
Function WebSet(obj, value, stepName)
    Dim actual, dr
    WebSet = False
    If Not ObjectReady(obj, stepName) Then Exit Function

    On Error Resume Next
    obj.Set value
    Wait 0, 300
    actual = obj.GetROProperty("value")
    Err.Clear

    If CStr(actual) <> CStr(value) Then
        obj.Click
        Set dr = CreateObject("Mercury.DeviceReplay")
        dr.KeyDown 29 : dr.PressKey 30 : dr.KeyUp 29      ' Ctrl+A
        dr.PressKey 14                                    ' Backspace
        dr.SendString value
        Wait 0, 300
        actual = obj.GetROProperty("value")
        Err.Clear
    End If
    On Error GoTo 0

    WebSet = (CStr(actual) = CStr(value))
    If WebSet Then
        LogPass stepName, "Value entered"
    Else
        LogFail stepName, "Expected '" & value & "' but field shows '" & actual & "'"
    End If
End Function

'------------------------------------------------------------------------------
' WebSetSecure - for passwords: value never appears in the run results
'------------------------------------------------------------------------------
Function WebSetSecure(obj, plainValue, stepName)
    WebSetSecure = False
    If Not ObjectReady(obj, stepName) Then Exit Function

    On Error Resume Next
    obj.SetSecure Crypt.Encrypt(plainValue)
    If Err.Number <> 0 Then
        Err.Clear
        obj.Set plainValue
    End If
    WebSetSecure = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    LogBoolResult stepName, WebSetSecure
End Function


'------------------------------------------------------------------------------
' WebClick - standard click, falls back to the DOM click()
'------------------------------------------------------------------------------
Function WebClick(obj, stepName)
    WebClick = False
    If Not ObjectReady(obj, stepName) Then Exit Function

    On Error Resume Next
    obj.Click
    If Err.Number <> 0 Then
        Err.Clear
        obj.Object.click
    End If
    WebClick = (Err.Number = 0)
    If WebClick Then
        LogPass stepName, "Clicked"
    Else
        LogFail stepName, "Click failed: " & Err.Description
    End If
    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' WebClickRobust - tries several click techniques; use when a plain click
' is swallowed by the page (custom components, overlays, JS frameworks)
'------------------------------------------------------------------------------
Function WebClickRobust(obj, stepName)
    Dim methods, i
    WebClickRobust = False
    If Not ObjectReady(obj, stepName) Then Exit Function

    methods = Array("Standard", "DomClick", "FireEvent", "MouseReplay", "DeviceReplay")
    For i = 0 To UBound(methods)
        If RunClickMethod(obj, methods(i)) Then
            LogPass stepName, "Clicked (method: " & methods(i) & ")"
            WebClickRobust = True
            Exit Function
        End If
    Next
    LogFail stepName, "All click methods raised errors"
End Function

'------------------------------------------------------------------------------
' ClickAndVerify - clicks obj with each technique until expectedObj appears
'------------------------------------------------------------------------------
Function ClickAndVerify(obj, expectedObj, stepName)
    Dim methods, i
    ClickAndVerify = False
    If Not ObjectReady(obj, stepName) Then Exit Function

    methods = Array("Standard", "DomClick", "FireEvent", "MouseReplay", "DeviceReplay", "Keyboard")
    For i = 0 To UBound(methods)
        On Error Resume Next
        obj.RefreshObject
        Err.Clear
        On Error GoTo 0
        RunClickMethod obj, methods(i)
        If WaitUntilDisplayed(expectedObj, 5) Then
            LogPass stepName, "Expected object displayed (click method: " & methods(i) & ")"
            ClickAndVerify = True
            Exit Function
        End If
    Next
    LogFail stepName, "Expected object did not appear after any click method"
End Function

' Performs ONE click technique. Returns True when it raised no error.
Function RunClickMethod(obj, method)
    Dim origReplay, dr, x, y
    On Error Resume Next
    Err.Clear

    Select Case method
        Case "Standard"
            obj.Click
        Case "DomClick"
            obj.Object.scrollIntoView False
            obj.Object.click
        Case "FireEvent"
            obj.FireEvent "onmouseover"
            obj.FireEvent "onmousedown"
            obj.FireEvent "onmouseup"
            obj.FireEvent "onclick"
        Case "MouseReplay"
            origReplay = Setting.WebPackage("ReplayType")
            Setting.WebPackage("ReplayType") = 2
            obj.Click
            Setting.WebPackage("ReplayType") = origReplay
        Case "DeviceReplay"
            x = obj.GetROProperty("abs_x") + Int(obj.GetROProperty("width") / 2)
            y = obj.GetROProperty("abs_y") + Int(obj.GetROProperty("height") / 2)
            Set dr = CreateObject("Mercury.DeviceReplay")
            dr.MouseMove x, y
            Wait 0, 200
            dr.MouseClick x, y, 0
        Case "Keyboard"
            obj.Object.focus
            Set dr = CreateObject("Mercury.DeviceReplay")
            dr.PressKey 28                                  ' Enter
    End Select

    RunClickMethod = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


'------------------------------------------------------------------------------
' Lists / checkboxes / keys
'------------------------------------------------------------------------------
Function WebSelect(obj, value, stepName)
    WebSelect = False
    If Not ObjectReady(obj, stepName) Then Exit Function
    On Error Resume Next
    obj.Select value
    WebSelect = (Err.Number = 0)
    If WebSelect Then LogPass stepName, "Selected '" & value & "'" Else LogFail stepName, Err.Description
    Err.Clear
    On Error GoTo 0
End Function

' desiredState: "ON" or "OFF"
Function SetCheckboxState(obj, desiredState, stepName)
    SetCheckboxState = False
    If Not ObjectReady(obj, stepName) Then Exit Function
    On Error Resume Next
    obj.Set UCase(desiredState)
    Err.Clear
    On Error GoTo 0
    Wait 0, 300
    SetCheckboxState = (IsCheckboxChecked(obj) = (UCase(desiredState) = "ON"))
    If SetCheckboxState Then
        LogPass stepName, "Checkbox is " & UCase(desiredState)
    Else
        LogFail stepName, "Checkbox state mismatch. Expected: " & UCase(desiredState)
    End If
End Function

Function IsCheckboxChecked(obj)
    Dim v
    On Error Resume Next
    v = LCase(Trim(CStr(obj.GetROProperty("checked"))))
    Err.Clear
    On Error GoTo 0
    IsCheckboxChecked = (v = "true" Or v = "1" Or v = "on")
End Function

Function WebPressEnter(obj, stepName)
    WebPressEnter = False
    If Not ObjectReady(obj, stepName) Then Exit Function
    On Error Resume Next
    obj.Object.focus
    CreateObject("Mercury.DeviceReplay").PressKey 28
    WebPressEnter = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    LogBoolResult stepName, WebPressEnter
End Function


'------------------------------------------------------------------------------
' Reading values
'------------------------------------------------------------------------------
Function WebGetText(obj)
    Dim t
    t = ""
    On Error Resume Next
    t = obj.GetROProperty("innertext")
    If Trim(t) = "" Then t = obj.GetROProperty("text")
    If Trim(t) = "" Then t = obj.GetROProperty("value")
    Err.Clear
    On Error GoTo 0
    WebGetText = Trim(CStr(t))
End Function

' Value of an input, including inputs that only expose it through the DOM
Function WebGetValue(obj)
    Dim v
    v = ""
    On Error Resume Next
    v = obj.Object.value
    If IsNull(v) Or IsEmpty(v) Or Err.Number <> 0 Then v = ""
    Err.Clear
    If Trim(CStr(v)) = "" Then v = obj.GetROProperty("value")
    Err.Clear
    On Error GoTo 0
    WebGetValue = CStr(v)
End Function

' Sets the DOM value and fires input/change - for JS-framework inputs
Sub WebSetDomValue(obj, textVal)
    Dim el, doc, ev, evName
    On Error Resume Next
    Set el = obj.Object
    el.focus
    el.value = textVal
    Set doc = el.ownerDocument
    For Each evName In Array("input", "change")
        Set ev = doc.createEvent("HTMLEvents")
        ev.initEvent evName, True, True
        el.dispatchEvent ev
    Next
    Err.Clear
    On Error GoTo 0
End Sub


'------------------------------------------------------------------------------
' Waiting / scrolling
'------------------------------------------------------------------------------
' True if obj exists and is visible within timeoutSec (polls every 0.5 sec)
Function WaitUntilDisplayed(obj, timeoutSec)
    Dim i, vis
    WaitUntilDisplayed = False
    On Error Resume Next
    For i = 0 To CInt(timeoutSec * 2)
        Err.Clear
        obj.RefreshObject
        If obj.Exist(0) Then
            vis = LCase(CStr(obj.GetROProperty("visible")))
            If vis <> "false" Then
                WaitUntilDisplayed = True
                Exit For
            End If
        End If
        Wait 0, 500
    Next
    Err.Clear
    On Error GoTo 0
End Function

' True once obj no longer exists (e.g. spinner, modal) within timeoutSec
Function WaitUntilGone(obj, timeoutSec)
    Dim i
    WaitUntilGone = False
    On Error Resume Next
    For i = 0 To CInt(timeoutSec * 2)
        obj.RefreshObject
        If Not obj.Exist(0) Then
            WaitUntilGone = True
            Exit For
        End If
        Wait 0, 500
    Next
    Err.Clear
    On Error GoTo 0
End Function

Sub ScrollIntoView(obj)
    On Error Resume Next
    obj.Object.scrollIntoView True
    Err.Clear
    On Error GoTo 0
End Sub

Sub ScrollToBottom(oPage)
    On Error Resume Next
    oPage.RunScript "window.scrollTo(0, document.body.scrollHeight);"
    Err.Clear
    On Error GoTo 0
End Sub

'==============================================================================
' Assertions.vbs
'   Verification helpers. All return True/False and log Pass/Fail.
'==============================================================================

Function VerifyObjectExist(obj, timeoutSec, stepName)
    Dim t, ok
    t = timeoutSec
    If IsEmpty(t) Or CStr(t) = "" Then t = GetConfigInt("DEFAULT_TIMEOUT", 10)
    ok = False
    On Error Resume Next
    ok = obj.Exist(t)
    Err.Clear
    On Error GoTo 0
    If ok Then
        LogPass stepName, "Object exists"
    Else
        LogFail stepName, "Object NOT found within " & t & " sec"
    End If
    VerifyObjectExist = ok
End Function

Function VerifyObjectNotExist(obj, timeoutSec, stepName)
    Dim ok
    ok = WaitUntilGone(obj, timeoutSec)
    If ok Then
        LogPass stepName, "Object is not present"
    Else
        LogFail stepName, "Object is still present after " & timeoutSec & " sec"
    End If
    VerifyObjectNotExist = ok
End Function

' Contains check on the object's visible text
Function VerifyText(obj, expected, stepName)
    Dim actual
    actual = WebGetText(obj)
    VerifyText = VerifyContains(actual, expected, stepName)
End Function

Function VerifyEquals(actual, expected, stepName)
    VerifyEquals = (CStr(actual) = CStr(expected))
    If VerifyEquals Then
        LogPass stepName, "Value = '" & actual & "'"
    Else
        LogFail stepName, "Expected: '" & expected & "' | Actual: '" & actual & "'"
    End If
End Function

Function VerifyContains(actual, expected, stepName)
    VerifyContains = (InStr(1, CStr(actual), CStr(expected), vbTextCompare) > 0)
    If VerifyContains Then
        LogPass stepName, "Found '" & expected & "'"
    Else
        LogFail stepName, "Expected to contain: '" & expected & "' | Actual: '" & actual & "'"
    End If
End Function

Function VerifyTrue(condition, stepName)
    VerifyTrue = (condition = True)
    LogBoolResult stepName, VerifyTrue
End Function

Function VerifyPageTitleContains(expected, stepName)
    Dim title
    title = ""
    On Error Resume Next
    title = GetBrowser().GetROProperty("title")
    Err.Clear
    On Error GoTo 0
    VerifyPageTitleContains = VerifyContains(title, expected, stepName)
End Function

Function VerifyUrlContains(expected, stepName)
    Dim url
    url = ""
    On Error Resume Next
    url = GetBrowser().GetROProperty("url")
    Err.Clear
    On Error GoTo 0
    VerifyUrlContains = VerifyContains(url, expected, stepName)
End Function

'==============================================================================
' Tests.vbs
'   One Sub per test case. The Sub name goes in the Function_Name column of
'   TestData\Execution.csv; the TestData column arrives as dictData.
'
'   The browser is already open on BASE_URL (or dictData("url")) when the Sub
'   starts and is closed after it ends (OPEN_BROWSER_PER_TEST=true).
'==============================================================================

'------------------------------------------------------------------------------
' TC_Sample_OpenApplication
'   Smoke test that works against any site: checks the app opened from .env.
'   TestData (optional):  title:<text expected in the page title>
'------------------------------------------------------------------------------
Sub TC_Sample_OpenApplication(dictData)
    Dim expectedTitle

    VerifyObjectExist GetBrowser(), 10, "Browser is open"
    VerifyUrlContains HostOf(GetConfig("BASE_URL")), "Application URL loaded"

    expectedTitle = GetData(dictData, "title", "")
    If expectedTitle <> "" Then
        VerifyPageTitleContains expectedTitle, "Page title contains '" & expectedTitle & "'"
    End If
End Sub


'------------------------------------------------------------------------------
' TC_Sample_Login  (template - update AppFunctions.Login objects first)
'   TestData (optional):  user:${APP_USERNAME};pass:${APP_PASSWORD}
'------------------------------------------------------------------------------
Sub TC_Sample_Login(dictData)
    Dim status

    status = Login(GetData(dictData, "user", ""), GetData(dictData, "pass", ""))
    RunStep "Login to application", status
    If Not status Then Exit Sub

    ' ... your scenario steps here ...

    RunStep "Logout from application", Logout()
End Sub


'------------------------------------------------------------------------------
' Helpers used by the samples
'------------------------------------------------------------------------------
Function HostOf(url)
    Dim s, p
    s = url
    p = InStr(s, "://")
    If p > 0 Then s = Mid(s, p + 3)
    p = InStr(s, "/")
    If p > 0 Then s = Left(s, p - 1)
    HostOf = s
End Function

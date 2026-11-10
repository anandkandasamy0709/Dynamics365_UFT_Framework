'==============================================================================
' AppFunctions.vbs
'   Business keywords for YOUR application (login, navigation, search...).
'   Keep them small, return True/False, and log through ReportUtil/WebActions.
'
'   The objects below use descriptive programming with PLACEHOLDER properties.
'   Replace them with your application's properties, or with objects from the
'   shared object repository, e.g.
'       Browser("App").Page("Login").WebEdit("Username")
'==============================================================================

'------------------------------------------------------------------------------
' Login - credentials come from .env (APP_USERNAME / APP_PASSWORD) unless
' passed explicitly.
'------------------------------------------------------------------------------
Function Login(username, password)
    Dim oPage, ok
    Login = False

    If Trim(username) = "" Then username = RequireConfig("APP_USERNAME")
    If Trim(password) = "" Then password = RequireConfig("APP_PASSWORD")
    If username = "" Or password = "" Then Exit Function

    Set oPage = GetPage()

    ' TODO: replace the placeholder descriptions with your login page objects
    ok = WebSet(oPage.WebEdit("html id:=username"), username, "Enter username")
    If ok Then ok = WebSetSecure(oPage.WebEdit("html id:=password"), password, "Enter password")
    If ok Then ok = WebClick(oPage.WebButton("html id:=login"), "Click Login")
    If ok Then
        GetBrowser().Sync
        ' TODO: replace with an object that only exists after a successful login
        ok = WaitUntilDisplayed(oPage.WebElement("html id:=home"), GetConfigInt("PAGE_LOAD_TIMEOUT", 60))
    End If

    Login = ok
End Function

Function Logout()
    Dim oPage
    Set oPage = GetPage()
    ' TODO: replace with your logout object(s)
    Logout = WebClick(oPage.Link("text:=Log ?[Oo]ut"), "Click Logout")
End Function

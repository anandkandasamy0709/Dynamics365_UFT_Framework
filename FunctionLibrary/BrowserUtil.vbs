'==============================================================================
' BrowserUtil.vbs
'   Browser lifecycle driven entirely by .env:
'     BASE_URL                    start URL
'     BROWSER                     chrome | edge | firefox
'     BROWSER_ARGS                extra command-line args (e.g. --incognito)
'     CLOSE_BROWSERS_BEFORE_TEST  kill running instances of BROWSER first
'     PAGE_LOAD_TIMEOUT           seconds to wait for the first page
'==============================================================================

Function GetBrowserExe()
    Select Case LCase(Trim(GetConfigOrDefault("BROWSER", "chrome")))
        Case "edge", "msedge":          GetBrowserExe = "msedge.exe"
        Case "firefox", "ff":           GetBrowserExe = "firefox.exe"
        Case Else:                      GetBrowserExe = "chrome.exe"
    End Select
End Function

' Most recently opened browser / its current page
Function GetBrowser()
    Set GetBrowser = Browser("creationtime:=0")
End Function

Function GetPage()
    Set GetPage = Browser("creationtime:=0").Page("micclass:=Page")
End Function


'------------------------------------------------------------------------------
' OpenBrowser([url])  - url optional; defaults to BASE_URL from .env
'------------------------------------------------------------------------------
Sub OpenBrowser(url)
    Dim target, args, timeout

    target = Trim(url)
    If target = "" Then target = RequireConfig("BASE_URL")
    If target = "" Then Err.Raise vbObjectError + 1001, "OpenBrowser", "BASE_URL is not set in .env"

    If GetConfigBool("CLOSE_BROWSERS_BEFORE_TEST", True) Then
        CloseBrowser
        Wait 1
    End If

    args = Trim(GetConfig("BROWSER_ARGS"))
    If args <> "" Then args = args & " "
    timeout = GetConfigInt("PAGE_LOAD_TIMEOUT", 60)

    SystemUtil.Run GetBrowserExe(), args & target
    Reporter.ReportEvent micDone, "Browser", "Launched " & GetBrowserExe() & " -> " & target

    ' Chromium "Restore pages?" bubble after a forced close
    If Window("regexpwndtitle:=Restore pages\?").Exist(3) Then
        Window("regexpwndtitle:=Restore pages\?").Close
        Reporter.ReportEvent micDone, "Browser", "'Restore pages' popup closed"
    End If

    If GetBrowser().Exist(timeout) Then
        GetBrowser().Sync
        On Error Resume Next
        Window("hwnd:=" & GetBrowser().GetROProperty("hwnd")).Maximize
        Err.Clear
        On Error GoTo 0
        Reporter.ReportEvent micDone, "Browser", "Browser ready"
    Else
        Err.Raise vbObjectError + 1002, "OpenBrowser", _
            "Browser not detected within " & timeout & " sec. Is the UFT browser extension installed and enabled?"
    End If
End Sub

Sub NavigateTo(url)
    GetBrowser().Navigate url
    GetBrowser().Sync
    Reporter.ReportEvent micDone, "Browser", "Navigated to " & url
End Sub


'------------------------------------------------------------------------------
' CloseBrowser - closes every running instance of the configured browser
'   CLOSE_ALL_BROWSERS=true also kills chrome/edge/firefox/iexplore
'------------------------------------------------------------------------------
Sub CloseBrowser()
    Dim exes, i, oShell, oWMI, colProc, p

    On Error Resume Next

    If GetConfigBool("CLOSE_ALL_BROWSERS", False) Then
        exes = Array("chrome.exe", "msedge.exe", "firefox.exe", "iexplore.exe")
    Else
        exes = Array(GetBrowserExe())
    End If

    Set oShell = CreateObject("WScript.Shell")
    For i = 0 To UBound(exes)
        oShell.Run "taskkill /F /T /IM " & exes(i), 0, True
    Next
    Wait 1

    ' WMI fallback for anything taskkill missed
    Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
    For i = 0 To UBound(exes)
        Set colProc = oWMI.ExecQuery("Select * from Win32_Process Where Name='" & exes(i) & "'")
        For Each p In colProc
            p.Terminate
        Next
    Next

    Err.Clear
    On Error GoTo 0
    Reporter.ReportEvent micDone, "Browser", "Closed: " & Join(exes, ", ")
End Sub

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


'==============================================================================
' Dynamics 365 Finance & Operations helpers (C9652593)
'   Objects: descriptive programming on data-dyn-controlname (AOT control name).
'==============================================================================

' Waits until the F&O shell is idle (no blocking overlay). True when idle.
Function D365_WaitIdle(timeoutSec)
    Dim i, res, txt, unknownCount, lastTxt
    D365_WaitIdle = False
    unknownCount = 0
    lastTxt = ""
    For i = 0 To CInt(timeoutSec * 2)
        txt = ""
        On Error Resume Next
        Err.Clear
        ' single JS expression (no 'return'), answers the text "busy" or "idle"
        res = GetPage().RunScript("(document.readyState!=='complete'||(document.getElementById('ShellBlockingDiv')!==null&&document.getElementById('ShellBlockingDiv').offsetWidth>0&&window.getComputedStyle(document.getElementById('ShellBlockingDiv')).display!=='none'))?'busy':'idle'")
        If Err.Number = 0 Then txt = LCase(Trim(CStr(res)))
        Err.Clear
        On Error GoTo 0
        lastTxt = txt
        If txt = "idle" Then
            D365_WaitIdle = True
            Exit For
        ElseIf txt <> "busy" Then
            ' RunScript gave no usable answer (error / empty): fall back to the browser's own Sync
            unknownCount = unknownCount + 1
            If unknownCount >= 3 Then
                On Error Resume Next
                GetBrowser().Sync
                D365_WaitIdle = (Err.Number = 0)
                Err.Clear
                On Error GoTo 0
                If D365_WaitIdle Then LogInfo "D365_WaitIdle", "RunScript returned no usable value; used Browser.Sync as idle check"
                Exit For
            End If
        End If
        Wait 0, 500
    Next
    If Not D365_WaitIdle Then LogInfo "D365_WaitIdle", "Not idle after " & timeoutSec & " s; last state = '" & lastTxt & "'"
End Function

' Opens a menu item through the F&O URL: https://<env>/?cmp=<company>&mi=<menuitem>
Function D365_OpenMenuItem(menuItem, stepName)
    Dim url, company, curUrl, p, q, tries, opened
    url = RequireConfig("BASE_URL")
    company = GetConfig("D365_COMPANY")      ' TODO(C9652593): DATA VALIDATION REQUIRED - legal entity (add D365_COMPANY to .env)
    If company = "" Then
        ' no company configured: keep the legal entity of the session that is already open (the F&O URL carries cmp=<company>)
        On Error Resume Next
        curUrl = CStr(GetPage().GetROProperty("url"))
        Err.Clear
        On Error GoTo 0
        p = InStr(1, curUrl, "cmp=", vbTextCompare)
        If p > 0 Then
            company = Mid(curUrl, p + 4)
            q = InStr(company, "&")
            If q > 0 Then company = Left(company, q - 1)
        End If
    End If
    If InStr(url, "?") > 0 Then url = Left(url, InStr(url, "?") - 1)
    If Right(url, 1) <> "/" Then url = url & "/"
    url = url & "?"
    If company <> "" Then url = url & "cmp=" & company & "&"
    url = url & "mi=" & menuItem
    opened = False
    For tries = 1 To 2
        GetBrowser().Navigate url
        GetBrowser().Sync
        D365_WaitIdle GetConfigInt("PAGE_LOAD_TIMEOUT", 60)
        On Error Resume Next
        curUrl = CStr(GetPage().GetROProperty("url"))
        Err.Clear
        On Error GoTo 0
        If InStr(1, curUrl, "mi=" & menuItem, vbTextCompare) > 0 Then
            opened = True
            Exit For
        End If
    Next
    D365_OpenMenuItem = opened
    If opened Then
        LogPass stepName, "Opened menu item " & menuItem
    Else
        LogFail stepName, "Menu item " & menuItem & " not opened; current URL: " & curUrl
    End If
End Function

' Selects the sales order row in the All sales orders grid by its number.
Function D365_SelectSalesOrder(salesId, stepName)
    Dim oPage, oFilter, oRow, ok
    D365_SelectSalesOrder = False
    Set oPage = GetPage()
    ' Quick filter on the grid (Sales order column)
    ' The list form loads after the shell is idle: wait for the quick filter (up to 60 s), trying the known F&O hooks.
    ' TODO(C9652593): verify object QuickFilterControl input with Object Spy
    Set oFilter = oPage.WebEdit("xpath:=//input[@name='QuickFilterControl_Input' or @placeholder='Filter' or @aria-label='Filter' or (ancestor::*[@data-dyn-controlname='QuickFilterControl'] and @type='text')]")
    If Not oFilter.Exist(60) Then
        LogFail stepName, "Quick filter input not found on the All sales orders page (60 s)"
        Exit Function
    End If
    ok = WebSet(oFilter, salesId, stepName & " - quick filter")
    If Not ok Then Exit Function
    WebPressEnter oFilter, stepName & " - apply filter"
    D365_WaitIdle 30
    ' the order number is a link in the Sales order column (seen on the All sales orders page)
    ' The order number cell of the filtered grid (link or text). Try the known shapes one by one, no XPath union.
    ' TODO(C9652593): verify the Sales order grid cell with Object Spy
    ' Screenshot: the order number is a blue link cell in the grid (SalesTableDetails page, Sales order column).
    ' Cell hooks are unconfirmed, so match any visible leaf element whose text is exactly the order number
    ' (outside input boxes: the quick filter box also holds the number), first match only.
    Set oRow = oPage.WebElement("xpath:=(//*[normalize-space(text())='" & salesId & "' and not(self::input) and not(self::option)])[1]")
    If Not oRow.Exist(30) Then Set oRow = oPage.WebElement("xpath:=(//*[normalize-space(.)='" & salesId & "' and not(*) and not(self::input)])[1]")
    If Not oRow.Exist(10) Then Set oRow = oPage.Link("xpath:=(//a[contains(normalize-space(.),'" & salesId & "')])[1]")
    If Not oRow.Exist(DefaultTimeout()) Then
        LogFail stepName, "Sales order " & salesId & " not found in the grid after filtering"
        Exit Function
    End If
    If Not WebClickRobust(oRow, stepName & " - open order") Then Exit Function
    D365_WaitIdle 30
    D365_SelectSalesOrder = True
    LogPass stepName, "Sales order " & salesId & " selected"
End Function

' Opens the Maintain tab of the Sales order action pane and clicks Cancel.
Function D365_ClickMaintainCancel(stepName)
    Dim oPage, oTab, oBtn, ok
    D365_ClickMaintainCancel = False
    Set oPage = GetPage()
    ' TODO(C9652593): verify object Maintain tab / Cancel button names with Object Spy
    ' The Maintain group (Modify / Cancel / Split order) sits on the 'Sales order' action pane tab: make sure that tab is open.
    Set oTab = oPage.WebElement("xpath:=//*[@data-dyn-role='PivotItem' and normalize-space(.)='Sales order'] | //li[normalize-space(.)='Sales order' and @role='tab']")
    If oTab.Exist(10) Then
        ok = WebClick(oTab, stepName & " - Sales order tab")
        If Not ok Then Exit Function
        D365_WaitIdle 15
    End If
    ' Cancel button of the Maintain group. The previous container-based XPath did not match (screenshot shows the
    ' ribbon open with Maintain > Modify / Cancel / Split order), so match the visible 'Cancel' label, outside dialogs.
    ' TODO(C9652593): verify the Cancel button control name with Object Spy
    Set oBtn = oPage.WebElement("xpath:=(//*[@data-dyn-role='ActionPane' or @data-dyn-role='RibbonGroup' or @data-dyn-role='Group' or contains(@class,'ribbon') or contains(@class,'actionPane')]//*[normalize-space(text())='Cancel'])[1]")
    If Not oBtn.Exist(10) Then Set oBtn = oPage.WebElement("xpath:=(//*[normalize-space(text())='Cancel' and not(ancestor::*[@data-dyn-role='Dialog']) and not(self::input)])[1]")
    If Not oBtn.Exist(10) Then Set oBtn = oPage.WebElement("xpath:=(//*[contains(@data-dyn-controlname,'Cancel') and not(ancestor::*[@data-dyn-role='Dialog'])])[1]")
    D365_ClickMaintainCancel = WebClickRobust(oBtn, stepName & " - Cancel")
    D365_WaitIdle 15
End Function

' Reads the message text of the confirmation dialog (empty if none).
Function D365_GetDialogText()
    Dim oMsg
    D365_GetDialogText = ""
    ' Screenshot: a modal with the question text and Yes / No buttons. The data-dyn-role='Dialog' hooks did not match,
    ' so find the innermost element that holds the question text.
    ' TODO(C9652593): verify dialog message object with Object Spy
    Set oMsg = GetPage().WebElement("xpath:=(//*[contains(normalize-space(.),'cancel all line quantities') and not(*[contains(normalize-space(.),'cancel all line quantities')])])[1]")
    If oMsg.Exist(30) Then D365_GetDialogText = Replace(WebGetText(oMsg), Chr(160), " ")
End Function

Function D365_ClickDialogButton(caption, stepName)
    Dim oBtn, oMsg, methods, i
    D365_ClickDialogButton = False
    ' Yes / No buttons of the modal: the element whose text is exactly the caption.
    Set oBtn = GetPage().WebElement("xpath:=(//button[normalize-space(.)='" & caption & "' or @name='" & caption & "' or @aria-label='" & caption & "'])[1]")
    If Not oBtn.Exist(15) Then Set oBtn = GetPage().WebElement("xpath:=(//*[normalize-space(text())='" & caption & "' and not(self::input) and not(self::option)])[1]")
    If Not oBtn.Exist(5) Then
        LogFail stepName, "Dialog button '" & caption & "' not found"
        Exit Function
    End If
    ' A plain click was reported as OK but the dialog stayed open (screenshot), so the click is only accepted
    ' once the question has disappeared; otherwise the next click technique is tried.
    Set oMsg = GetPage().WebElement("xpath:=(//*[contains(normalize-space(.),'cancel all line quantities') and not(*[contains(normalize-space(.),'cancel all line quantities')])])[1]")
    methods = Array("Standard", "DomClick", "MouseReplay", "DeviceReplay", "FireEvent", "Keyboard")
    For i = 0 To UBound(methods)
        If RunClickMethod(oBtn, methods(i)) Then
            D365_WaitIdle 15
            If Not oMsg.Exist(4) Then
                D365_ClickDialogButton = True
                LogPass stepName, "Clicked '" & caption & "' (method: " & methods(i) & "), dialog closed"
                Exit For
            End If
        End If
    Next
    If Not D365_ClickDialogButton Then LogFail stepName, "Dialog still open after clicking '" & caption & "' with all click techniques"
    D365_WaitIdle 30
End Function

' Reads the sales order status from the header / details of the selected order.
Function D365_GetSalesStatus()
    Dim oStat, i, j, shapes, words, w
    D365_GetSalesStatus = ""
    ' The grid shows the status in the Status column (screenshot: 'Canceled' after the cancel). The status cell is not
    ' matched by text() alone (the same cell was only found by the shapes used for 'Open order'), so every shape is tried.
    ' TODO(C9652593): verify the Status column cell with Object Spy
    words = Array("Canceled", "Cancelled")
    For i = 1 To 6                      ' the server-side cancel refreshes the grid after the Infolog: wait for the value
        D365_WaitIdle 10
        For Each w In words
            shapes = Array("//*[not(*) and normalize-space(text())='" & w & "']", "//*[not(*) and normalize-space(.)='" & w & "']", "//*[@value='" & w & "']", "//*[normalize-space(.)='" & w & "' and not(*[normalize-space(.)='" & w & "'])]")
            For j = 0 To UBound(shapes)
                Set oStat = GetPage().WebElement("xpath:=(" & shapes(j) & ")[last()]")
                If oStat.Exist(3) Then
                    D365_GetSalesStatus = Trim(WebGetText(oStat))
                    If D365_GetSalesStatus = "" Then D365_GetSalesStatus = w
                    Exit Function
                End If
            Next
        Next
    Next
    ' not cancelled: report whatever status the row shows
    Set oStat = GetPage().WebElement("xpath:=(//*[not(*) and (normalize-space(text())='Open order' or normalize-space(text())='Delivered' or normalize-space(text())='Invoiced')])[last()]")
    If oStat.Exist(3) Then D365_GetSalesStatus = Trim(WebGetText(oStat))
End Function


' Signs in only when the Microsoft sign-in page is shown; with an existing session / SSO the
' F&O shell is already open and nothing is typed. Credentials come from .env.
Function D365_EnsureLoggedIn()
    Dim oPage, oUser, oNext, oPwd, oSignIn, oStay, oTile, i
    D365_EnsureLoggedIn = False
    Set oPage = GetPage()
    ' "Pick an account" page (remembered Microsoft session): click the tile of APP_USERNAME
    Set oTile = oPage.WebElement("xpath:=//*[@data-test-id='" & RequireConfig("APP_USERNAME") & "' or @data-test-id='" & LCase(RequireConfig("APP_USERNAME")) & "']")
    If oTile.Exist(5) Then
        If Not WebClick(oTile, "Pick account " & RequireConfig("APP_USERNAME")) Then Exit Function
        GetBrowser().Sync
        Set oStay = oPage.WebButton("xpath:=//input[@id='idSIButton9']")
        If oStay.Exist(5) Then oStay.Click      ' "Stay signed in?" prompt
    End If
    Set oUser = oPage.WebEdit("xpath:=//input[@name='loginfmt']")
    If oUser.Exist(5) Then
        If Not WebSet(oUser, RequireConfig("APP_USERNAME"), "Enter username") Then Exit Function
        If Not WebClick(oPage.WebButton("xpath:=//input[@id='idSIButton9']"), "Click Next") Then Exit Function
        Set oPwd = oPage.WebEdit("xpath:=//input[@name='passwd']")
        If Not oPwd.Exist(15) Then
            LogFail "Login", "Password field not shown"
            Exit Function
        End If
        If Not WebSetSecure(oPwd, RequireConfig("APP_PASSWORD"), "Enter password") Then Exit Function
        If Not WebClick(oPage.WebButton("xpath:=//input[@id='idSIButton9']"), "Click Sign in") Then Exit Function
        Set oStay = oPage.WebButton("xpath:=//input[@id='idSIButton9']")
        If oStay.Exist(5) Then oStay.Click      ' "Stay signed in?" prompt
    End If
    GetBrowser().Sync
    D365_EnsureLoggedIn = D365_WaitIdle(GetConfigInt("PAGE_LOAD_TIMEOUT", 60))
    If D365_EnsureLoggedIn Then
        LogPass "Login", "Dynamics 365 session is open"
    Else
        LogFail "Login", "Dynamics 365 did not finish loading"
    End If
End Function

' Finds, at run time, the LAST sales order with status "Open order" in the All sales orders grid (case: "select the last
' available sales order ... with status Open"). Returns the order number, or "" when none is found.
' Only xpath descriptions work in this UFT/Chrome setup (RunScript / ChildObjects returned nothing). The last "Open order"
' cell is found first (that worked); its order number is then read from the same grid row.
Function D365_FindLastOpenSalesOrder(stepName)
    Dim i, oStat, oSo, statX, soCell, soX(2), txt, found, n, shapes, j, k, html
    D365_FindLastOpenSalesOrder = ""
    found = ""
    n = 0
    ' an order-number cell: element whose own text or value starts with SO-
    soCell = "*[starts-with(normalize-space(text()),'SO-') or starts-with(@value,'SO-')]"
    shapes = Array("//*[not(*) and normalize-space(text())='Open order']", "//*[not(*) and normalize-space(.)='Open order']", "//*[@value='Open order']")
    For i = 1 To 4
        D365_WaitIdle 5
        For j = 0 To UBound(shapes)
            statX = "(" & shapes(j) & ")[last()]"
            Set oStat = GetPage().WebElement("xpath:=" & statX)
            If oStat.Exist(5) Then
                n = n + 1
                ' A: the nearest ancestor that holds exactly one order number = the grid row of this status cell
                soX(0) = statX & "/ancestor::*[count(.//" & soCell & ")=1][1]//" & soCell
                ' B: the nearest order-number cell before the status cell in document order
                soX(1) = statX & "/preceding::" & soCell & "[1]"
                ' C: same, but any element whose text contains the number (link with nested spans)
                soX(2) = statX & "/preceding::*[contains(text(),'SO-')][1]"
                For k = 0 To 2
                    Set oSo = GetPage().WebElement("xpath:=" & soX(k))
                    If oSo.Exist(3) Then
                        txt = Trim(Replace(WebGetText(oSo), Chr(160), " "))
                        If InStr(txt, "SO-") > 0 Then
                            found = Mid(txt, InStr(txt, "SO-"))
                            If InStr(found, " ") > 0 Then found = Left(found, InStr(found, " ") - 1)
                            Exit For
                        End If
                    End If
                Next
                If found = "" And i = 1 Then
                    ' evidence for the next round: the markup around the status cell
                    html = ""
                    On Error Resume Next
                    html = CStr(GetPage().WebElement("xpath:=" & statX & "/ancestor::*[3]").GetROProperty("outerhtml"))
                    Err.Clear
                    On Error GoTo 0
                    LogInfo stepName, "Row markup (shape " & j & "): " & Left(Replace(html, vbCrLf, " "), 700)
                End If
            End If
            If found <> "" Then Exit For
        Next
        If found <> "" Then Exit For
    Next
    If found = "" Then
        LogFail stepName, "No sales order with status 'Open order' found in the All sales orders grid (Open order cells located: " & n & ")"
    Else
        D365_FindLastOpenSalesOrder = found
        LogInfo stepName, "Last Open order in the grid: " & found
    End If
End Function

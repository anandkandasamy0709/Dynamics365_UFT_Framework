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
        ' F&O shows the All sales orders list (SalesTableListPage) as the "Sales order details" grid view: mi=SalesTableDetails
        If StrComp(menuItem, "SalesTableListPage", vbTextCompare) = 0 And InStr(1, curUrl, "mi=SalesTableDetails", vbTextCompare) > 0 Then
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

' Returns to the sales order DETAILS page by stepping back in the browser history (wave -> shipment details -> order).
' Used because a click on the order number in the All sales orders list does not open the details form in this environment.
Function D365_BackToOrder(salesId, stepName)
    Dim i
    D365_BackToOrder = False
    For i = 1 To 4
        If InStr(1, D365_ReadFirstMatch(salesId & " : "), salesId, vbTextCompare) > 0 Then
            D365_BackToOrder = True
            LogPass stepName, "Back on the sales order details page of " & salesId
            Exit Function
        End If
        On Error Resume Next
        GetBrowser().Back
        Err.Clear
        On Error GoTo 0
        D365_WaitIdle 30
        D365FO_WaitText "Sales order", 15
    Next
    If InStr(1, D365_ReadFirstMatch(salesId & " : "), salesId, vbTextCompare) > 0 Then D365_BackToOrder = True: LogPass stepName, "Back on the sales order details page of " & salesId
End Function

' After D365_SelectSalesOrder from the All sales orders list: make sure the order DETAILS page (heading "SO-nnnnnn : <customer>") is open.
' The first click on the order number sometimes only selects the grid row and the list stays open (Step 17: no lines grid, no Warehouse > Work details).
Function D365_EnsureOrderDetailsOpen(salesId, stepName)
    Dim i, oLink, js, dr, oldReplay
    D365_EnsureOrderDetailsOpen = False
    For i = 1 To 5
        If InStr(1, D365_ReadFirstMatch(salesId & " : "), salesId, vbTextCompare) > 0 Then
            D365_EnsureOrderDetailsOpen = True
            LogPass stepName, "Sales order details page of " & salesId & " is open"
            Exit Function
        End If
        Set oLink = GetPage().Link("xpath:=(//a[normalize-space(.)='" & salesId & "'])[1]")
        If Not oLink.Exist(3) Then Set oLink = GetPage().WebElement("xpath:=(//*[normalize-space(text())='" & salesId & "' and not(self::input) and not(self::option)])[1]")
        Select Case i
            Case 1
                ' real mouse events (F&O grid links ignore synthetic clicks)
                oldReplay = Setting.WebPackage("ReplayType")
                Setting.WebPackage("ReplayType") = 2
                On Error Resume Next
                If oLink.Exist(5) Then oLink.Click
                Err.Clear
                On Error GoTo 0
                Setting.WebPackage("ReplayType") = oldReplay
            Case 2
                ' double click on the row opens the details in F&O lists
                oldReplay = Setting.WebPackage("ReplayType")
                Setting.WebPackage("ReplayType") = 2
                On Error Resume Next
                If oLink.Exist(5) Then oLink.DblClick 5, 5
                Err.Clear
                On Error GoTo 0
                Setting.WebPackage("ReplayType") = oldReplay
            Case 3
                ' focus the order number and press Enter (opens the selected row)
                On Error Resume Next
                If oLink.Exist(5) Then oLink.Click
                Set dr = CreateObject("Mercury.DeviceReplay")
                dr.PressKey 28
                Err.Clear
                On Error GoTo 0
            Case Else
                js = "(function(){var e=document.querySelectorAll('a,span,div');for(var k=0;k<e.length;k++){var x=e[k];if(x.children.length===0&&(x.textContent||'').trim()==='" & salesId & "'&&x.tagName!=='INPUT'){var a=x.closest('a')||x;a.click();return 'ok';}}return 'none';})()"
                On Error Resume Next
                GetPage().RunScript js
                Err.Clear
                On Error GoTo 0
        End Select
        D365_WaitIdle 30
        Wait 2
    Next
    LogFail stepName, "Sales order details page of " & salesId & " did not open from the list"
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


'==============================================================================
' Dynamics 365 F&O helpers (C9652592 - Creation of Sales Order, order to delivery)
'   Additive: generic control-name / label based helpers, descriptive programming only.
'   Every locator below is unconfirmed: TODO(C9652592): verify with Object Spy.
'==============================================================================

' F&O element by AOT control name
Function D365FO_Ctl(controlName)
    Set D365FO_Ctl = GetPage().WebElement("xpath:=(//*[@data-dyn-controlname='" & controlName & "'])[last()]")
End Function

' Clicks a button / tab / menu item by its visible label. A dialog button wins over a page one.
Function D365FO_ClickLabel(label, stepName)
    Dim obj, shapes, i, tmo
    D365FO_ClickLabel = False
    shapes = Array( _
        "(//*[(@data-dyn-role='Dialog' or @role='dialog')]//*[(@data-dyn-role='CommandButton' or @data-dyn-role='Button' or self::button) and normalize-space(.)='" & label & "'])[last()]", _
        "(//*[(@data-dyn-role='ActionPaneTab' or @data-dyn-role='PivotItem' or @role='tab' or @data-dyn-role='MenuItemButton' or @data-dyn-role='CommandButton' or @data-dyn-role='Button' or self::button) and normalize-space(.)='" & label & "'])[last()]", _
        "(//*[not(*) and normalize-space(text())='" & label & "' and not(self::input) and not(self::option)])[last()]")
    For i = 0 To UBound(shapes)
        Set obj = GetPage().WebElement("xpath:=" & shapes(i))
        tmo = 8
        If i = 0 Then tmo = 3
        If obj.Exist(tmo) Then
            D365FO_ClickLabel = WebClickRobust(obj, stepName)
            D365_WaitIdle 30
            Exit Function
        End If
    Next
    LogFail stepName, "Button / tab '" & label & "' not found"
End Function

' Opens an action pane tab (ribbon). Not the same-named menu buttons of the lines grid (e.g. Sales order lines > Warehouse).
Function D365FO_ClickTab(label, stepName)
    Dim obj
    Set obj = GetPage().WebElement("xpath:=(//*[(@data-dyn-role='ActionPaneTab' or (@role='tab' and ancestor::*[@data-dyn-role='ActionPane'])) and normalize-space(.)='" & label & "'])[last()]")
    If obj.Exist(8) Then
        D365FO_ClickTab = WebClickRobust(obj, stepName)
        D365_WaitIdle 30
    Else
        D365FO_ClickTab = D365FO_ClickLabel(label, stepName)
    End If
End Function

' Clicks a control by its AOT name, falling back to its label
Function D365FO_ClickCtl(controlName, label, stepName)
    Dim obj
    Set obj = D365FO_Ctl(controlName)
    If obj.Exist(5) Then
        D365FO_ClickCtl = WebClickRobust(obj, stepName)
        D365_WaitIdle 30
    Else
        D365FO_ClickCtl = D365FO_ClickLabel(label, stepName)
    End If
End Function

' True when the typed value is what the field shows (text contains it, or equal numbers such as 2 and 2.00)
Function D365FO_ValueMatches(actual, expected)
    D365FO_ValueMatches = False
    If InStr(1, actual, expected, vbTextCompare) > 0 Then D365FO_ValueMatches = True
    If IsNumeric(actual) And IsNumeric(expected) Then
        If CDbl(actual) = CDbl(expected) Then D365FO_ValueMatches = True
    End If
End Function

' Types into an input with real keys; if the value did not arrive (focus lost, e.g. to the browser address bar) sets it directly. True when the field shows it.
Function D365FO_TypeAndVerify(obj, value)
    Dim dr, actual
    obj.Click
    On Error Resume Next
    obj.Object.focus
    Err.Clear
    On Error GoTo 0
    Set dr = CreateObject("Mercury.DeviceReplay")
    dr.KeyDown 29 : dr.PressKey 30 : dr.KeyUp 29       ' Ctrl+A
    dr.SendString CStr(value)
    Wait 0, 500
    dr.PressKey 15                                      ' Tab
    D365_WaitIdle 30
    actual = Trim(Replace(CStr(obj.GetROProperty("value")), Chr(160), " "))
    If Not D365FO_ValueMatches(actual, CStr(value)) Then
        obj.Click
        obj.Set CStr(value)
        Wait 0, 500
        dr.PressKey 15
        D365_WaitIdle 30
        actual = Trim(Replace(CStr(obj.GetROProperty("value")), Chr(160), " "))
    End If
    D365FO_TypeAndVerify = D365FO_ValueMatches(actual, CStr(value))
End Function

' Closes the "Split view" / what-is-new tip bubbles that steal the focus
Sub D365FO_DismissTips()
    Dim o
    Set o = GetPage().WebElement("xpath:=(//*[normalize-space(text())='Remind me later'])[1]")
    If o.Exist(2) Then
        o.Click
        D365_WaitIdle 10
    End If
End Sub

' Types into an F&O input (by control name) with real key events and leaves it with Tab so the lookup / onchange fires
Function D365FO_SetInput(controlName, value, stepName)
    Dim obj, dr
    D365FO_SetInput = False
    Set obj = GetPage().WebEdit("xpath:=(//*[@data-dyn-controlname='" & controlName & "']//input[not(@type='hidden') and not(@type='checkbox')])[last()]")
    If Not obj.Exist(DefaultTimeout()) Then
        ' exact AOT name unconfirmed: fall back to any control whose name contains the key part (grid columns carry prefixes)
        Set obj = GetPage().WebEdit("xpath:=(//*[contains(@data-dyn-controlname,'" & Replace(controlName, "SalesLine_", "") & "')]//input[not(@type='hidden') and not(@type='checkbox')])[last()]")
    End If
    If Not obj.Exist(DefaultTimeout()) Then
        LogFail stepName, "Input '" & controlName & "' not found"
        Exit Function
    End If
    If Not D365FO_TypeAndVerify(obj, value) Then
        LogFail stepName, "Field " & controlName & " does not show " & value & " after typing (shows " & CStr(obj.GetROProperty("value")) & ")"
        Exit Function
    End If
    D365FO_SetInput = True
    LogPass stepName, value & " entered in " & controlName
End Function

' Value of an F&O input by (part of) the control name, "" when missing
Function D365FO_GetInput(controlName)
    Dim obj
    D365FO_GetInput = ""
    Set obj = GetPage().WebEdit("xpath:=(//*[contains(@data-dyn-controlname,'" & controlName & "')]//input[not(@type='hidden')])[last()]")
    If obj.Exist(10) Then D365FO_GetInput = Trim(Replace(CStr(obj.GetROProperty("value")), Chr(160), " "))
End Function

' Polls (no fixed sleep) until an element with this exact text exists
Function D365FO_WaitText(txt, timeoutSec)
    Dim obj, t
    D365FO_WaitText = False
    t = Timer
    Do
        Set obj = GetPage().WebElement("xpath:=(//*[not(*) and normalize-space(text())='" & txt & "'])[last()]")
        If obj.Exist(2) Then
            D365FO_WaitText = True
            Exit Function
        End If
        D365_WaitIdle 10
    Loop While Timer - t < timeoutSec And Timer >= t
End Function

' Polls (no fixed sleep) until an element whose text contains this text exists
Function D365FO_WaitTextContains(part, timeoutSec)
    Dim obj, t
    D365FO_WaitTextContains = False
    t = Timer
    Do
        Set obj = GetPage().WebElement("xpath:=(//*[not(*) and contains(normalize-space(text()),'" & part & "')])[last()]")
        If obj.Exist(2) Then
            D365FO_WaitTextContains = True
            Exit Function
        End If
        D365_WaitIdle 10
    Loop While Timer - t < timeoutSec And Timer >= t
End Function

' Goes Back until the Work page (it has the Shipment tab) is shown again, e.g. from Work completion
Function D365FO_BackToWorkPage(stepName)
    Dim i, t
    D365FO_BackToWorkPage = False
    ' The previous page stays in the DOM (hidden), so label checks such as "Shipment" are true even on the Work completion page.
    ' Use the browser tab title instead: back until it no longer says "Work completion".
    For i = 1 To 4
        ' the "Processing operation" dialog swallows clicks until the operation has finished
        t = Timer
        Do While GetPage().WebElement("xpath:=//*[not(*) and contains(normalize-space(text()),'Processing operation')]").Exist(1) And Timer - t < 90 And Timer >= t
            D365_WaitIdle 10
        Loop
        If InStr(1, GetPage().GetROProperty("title"), "Work completion", vbTextCompare) = 0 Then
            D365FO_BackToWorkPage = True
            Exit Function
        End If
        D365FO_ClickCtl "SystemDefinedCloseButton", "Back", stepName
        D365_WaitIdle 30
        t = Timer
        Do While InStr(1, GetPage().GetROProperty("title"), "Work completion", vbTextCompare) > 0 And Timer - t < 10 And Timer >= t
            D365_WaitIdle 5
        Loop
    Next
    D365FO_BackToWorkPage = (InStr(1, GetPage().GetROProperty("title"), "Work completion", vbTextCompare) = 0)
End Function

' Sales order number from the open order header
Function D365FO_GetSalesIdFromHeader()
    Dim o, t, p
    D365FO_GetSalesIdFromHeader = ""
    ' page title of the order: "SO-218718 : AV Specialists, Inc"
    Set o = GetPage().WebElement("xpath:=(//*[not(*) and starts-with(normalize-space(text()),'SO-') and contains(text(),' : ')])[1]")
    If o.Exist(30) Then
        t = Trim(Replace(WebGetText(o), Chr(160), " "))
        p = InStr(t, " ")
        If p > 0 Then t = Left(t, p - 1)
        D365FO_GetSalesIdFromHeader = t
    End If
End Function

' Create sales order dialog: cancels the "Sales order summary" panel when shown, enters the customer, OK
Function D365_CreateSalesOrder(customer, stepName)
    Dim oSummary
    D365_CreateSalesOrder = False
    D365FO_DismissTips
    Set oSummary = GetPage().WebElement("xpath:=(//*[normalize-space(text())='Sales order summary'])[1]")
    If oSummary.Exist(5) Then
        If Not D365FO_ClickLabel("Cancel", stepName & " - cancel Sales order summary") Then Exit Function
    End If
    ' TODO(C9652592): verify control name SalesTable_CustAccount
    If Not D365FO_SetInput("SalesTable_CustAccount", customer, stepName & " - customer account") Then Exit Function
    D365_CreateSalesOrder = D365FO_ClickLabel("OK", stepName & " - OK")
End Function

' Sales order lines tab > Add line, then item and quantity
Function D365_AddSalesLine(item, qty, stepName)
    D365_AddSalesLine = False
    D365FO_ClickLabel "Sales order lines", stepName & " - lines tab"
    If Not D365FO_ClickLabel("Add line", stepName & " - Add line") Then Exit Function
    ' TODO(C9652592): verify control names SalesLine_ItemId / SalesLine_SalesQty
    If Not D365FO_SetInput("SalesLine_ItemId", item, stepName & " - item number") Then Exit Function
    D365_WaitIdle 30      ' item defaults (site, price, product name) load after the item is accepted
    If Not D365FO_SetInput("SalesLine_SalesQty", qty, stepName & " - quantity") Then Exit Function
    ' The line is only a draft until it is saved: Confirm now / Confirm sales order stay disabled and the MAO allocation is not set.
    D365_AddSalesLine = D365FO_ClickCtl("SystemDefinedSaveButton", "Save", stepName & " - Save the order line")
    D365_WaitIdle 30
End Function

' True when a visible button / tab with this label exists right now (no logging)
Function D365FO_LabelExists(label, timeoutSec)
    D365FO_LabelExists = GetPage().WebElement("xpath:=(//*[(@data-dyn-role='ActionPaneTab' or @data-dyn-role='PivotItem' or @role='tab' or @data-dyn-role='MenuItemButton' or @data-dyn-role='CommandButton' or @data-dyn-role='Button' or self::button) and normalize-space(.)='" & label & "'])[last()]").Exist(timeoutSec)
End Function

' Finds the action pane tab that holds the button (the case does not name it) and opens it
' Polls (refresh + Sell tab) until the label exists; True when found
' Step 9 of the case: "Confirm now (if not enabled click refresh until it gets enabled)". Polls: Sell tab, look for an ENABLED
' Confirm now (no aria-disabled / disabled class), else refresh the order. True when it is enabled.
Function D365_WaitConfirmNowEnabled(timeoutSec)
    Dim t, xp
    ' enabled = the button (or the text inside it) has no aria-disabled='true' on itself or on an ancestor (class names are not used: an ancestor class matched 'disabled' in the last run)
    xp = "(//*[normalize-space(text())='Confirm now' and not(ancestor-or-self::*[@aria-disabled='true'])])[last()]"
    D365_WaitConfirmNowEnabled = False
    t = Timer
    Do
        D365FO_ClickTab "Sell", "Open Sell tab"
        If GetPage().WebElement("xpath:=" & xp).Exist(3) Then D365_WaitConfirmNowEnabled = True: Exit Do
        D365FO_ClickCtl "SystemDefinedRefreshButton", "Refresh", "Refresh the order (wait for Confirm now to be enabled)"
    Loop While Timer - t < timeoutSec And Timer >= t
End Function

Function D365_WaitCompleteButton(timeoutSec)
    Dim t
    D365_WaitCompleteButton = False
    t = Timer
    Do
        D365FO_ClickTab "Sell", "Open Sell tab"
        If D365FO_LabelExists("Complete", 3) Then D365_WaitCompleteButton = True: Exit Do
        D365FO_ClickCtl "SystemDefinedRefreshButton", "Refresh", "Refresh the order (wait for the confirmation to finish)"
    Loop While Timer - t < timeoutSec And Timer >= t
End Function

Function D365FO_OpenTabWith(buttonLabel, tabs)
    Dim i
    D365FO_OpenTabWith = False
    If D365FO_LabelExists(buttonLabel, 3) Then
        D365FO_OpenTabWith = True
        Exit Function
    End If
    For i = 0 To UBound(tabs)
        If D365FO_ClickLabel(tabs(i), "Open tab " & tabs(i) & " to find " & buttonLabel) Then
            If D365FO_LabelExists(buttonLabel, 4) Then
                D365FO_OpenTabWith = True
                Exit Function
            End If
        End If
    Next
End Function

' Polls the allocation id of the line (set asynchronously by the MAO integration after the save). "" when none in time.
' Step 7: the column header of the lines grid (AOT control ...AllocationId / header text) is visible. Short wait only.
' TODO(C9652592): verify the column control name with Object Spy
Function D365_AllocationColumnVisible(timeoutSec)
    Dim t
    D365_AllocationColumnVisible = False
    t = Timer
    Do
        If GetPage().WebElement("xpath:=//*[contains(@data-dyn-controlname,'AllocationId') and (@role='columnheader' or contains(@class,'columnHeader') or ancestor::*[@data-dyn-role='Grid'])]").Exist(2) Then D365_AllocationColumnVisible = True: Exit Do
        If GetPage().WebElement("xpath:=//*[@role='columnheader' and contains(normalize-space(.),'AllocationId')]").Exist(1) Then D365_AllocationColumnVisible = True: Exit Do
    Loop While Timer - t < timeoutSec And Timer >= t
End Function

Function D365_WaitAllocationId(timeoutSec)
    Dim t, v, opened, i, tabs, seen
    opened = False
    tabs = Array("General", "Sourcing", "Setup", "Delivery", "Product", "Packing", "Price and discount", "Address")
    t = Timer
    Do
        v = D365FO_GetInput("AllocationId")
        If v = "" Then v = D365FO_GetByLabel("Allocation id")
        If v <> "" Then Exit Do
        If Not opened Then
            ' the field is not on the grid: it sits in the collapsed Line details section, on one of its tabs
            opened = True
            If D365FO_LabelExists("Line details", 3) Then D365FO_ClickLabel "Line details", "Open Line details"
        End If
        ' the MAO tab of Line details (control SNPOUniteMAO) holds the MAO fields: open it and read the allocation id there
        If D365FO_Ctl("SNPOUniteMAO").Exist(2) Then
            D365FO_ClickCtl "SNPOUniteMAO", "MAO", "Allocation id - open the MAO tab of Line details"
            v = D365FO_GetInput("AllocationId")
            If v = "" Then v = D365FO_GetByLabel("Allocation id")
            If v = "" Then seen = seen & "MAO tab inputs: " & D365FO_InputValuesIn("SNPOUniteMAO") & " | "
            If v <> "" Then Exit Do
        End If
        ' look on every Line details tab (tab content is only rendered once the tab is opened)
        For i = 0 To UBound(tabs)
            If D365FO_LabelExists(tabs(i), 1) Then
                D365FO_ClickLabel tabs(i), "Allocation id - open Line details tab " & tabs(i)
                v = D365FO_GetInput("AllocationId")
                If v = "" Then v = D365FO_GetByLabel("Allocation id")
                seen = seen & tabs(i) & ": " & D365FO_ControlNamesLike("llocation") & " | "
                If v <> "" Then Exit For
            End If
        Next
        If v <> "" Then Exit Do
        D365FO_ClickCtl "SystemDefinedRefreshButton", "Refresh", "Refresh the order"
    Loop While Timer - t < timeoutSec And Timer >= t
    D365_WaitAllocationId = v
    If v = "" Then LogInfo "Allocation id", "Controls with llocation per tab: " & seen & " / SNPOUnite: " & D365FO_ControlNamesLike("SNPOUnite")
End Function

' Evidence helper: AOT control names on the page that contain the text (empty when RunScript gives nothing)
Function D365FO_InputValuesIn(containerPart)
    Dim r
    D365FO_InputValuesIn = ""
    On Error Resume Next
    r = GetPage().RunScript("(function(){var o=[],q=document.querySelectorAll('[data-dyn-controlname*=""" & containerPart & """] input');for(var i=0;i<q.length&&o.length<40;i++){var c=q[i].closest('[data-dyn-controlname]');o.push((c?c.getAttribute('data-dyn-controlname'):'?')+'='+q[i].value);}return o.join('; ');})()")
    If Err.Number = 0 Then D365FO_InputValuesIn = CStr(r)
    Err.Clear
    On Error GoTo 0
End Function

' Evidence helper: visible text / input values of the open lookup list
Function D365FO_LookupText()
    Dim r
    D365FO_LookupText = ""
    On Error Resume Next
    r = GetPage().RunScript("(function(){var q=document.querySelectorAll('[data-dyn-role=""Lookup""],.lookupDock,[class*=""lookup""]');var o=[];for(var i=0;i<q.length;i++){var t=(q[i].innerText||'').replace(/\s+/g,' ');var v=[].map.call(q[i].querySelectorAll('input'),function(x){return x.value}).join(',');o.push(t.substr(0,200)+' {'+v.substr(0,200)+'}');}return o.slice(-3).join(' || ');})()")
    If Err.Number = 0 Then D365FO_LookupText = CStr(r)
    Err.Clear
    On Error GoTo 0
End Function

Function D365FO_ControlNamesLike(part)
    Dim r
    D365FO_ControlNamesLike = ""
    On Error Resume Next
    r = GetPage().RunScript("(function(){var o=[],q=document.querySelectorAll('[data-dyn-controlname]');for(var i=0;i<q.length;i++){var n=q[i].getAttribute('data-dyn-controlname');if(n.indexOf('" & part & "')>=0&&o.indexOf(n)<0)o.push(n);}return o.join(', ');})()")
    If Err.Number = 0 Then D365FO_ControlNamesLike = CStr(r)
    Err.Clear
    On Error GoTo 0
End Function

' Input of a field found by its visible label (controls whose AOT name is not known). Last match = the open dialog / flyout.
Function D365FO_InputByLabel(label)
    Set D365FO_InputByLabel = GetPage().WebEdit("xpath:=(//*[normalize-space(text())='" & label & "' and not(self::input)]/ancestor::*[@data-dyn-controlname][1]//input[not(@type='hidden') and not(@type='checkbox')])[last()]")
End Function

Function D365FO_SetByLabel(label, value, stepName)
    Dim obj, dr
    D365FO_SetByLabel = False
    Set obj = D365FO_InputByLabel(label)
    If Not obj.Exist(DefaultTimeout()) Then
        LogFail stepName, "Field with label '" & label & "' not found"
        Exit Function
    End If
    If Not D365FO_TypeAndVerify(obj, value) Then
        LogFail stepName, "Field " & label & " does not show " & value & " after typing"
        Exit Function
    End If
    D365FO_SetByLabel = True
    LogPass stepName, value & " entered in field " & label
End Function

Function D365FO_GetByLabel(label)
    Dim obj
    D365FO_GetByLabel = ""
    Set obj = D365FO_InputByLabel(label)
    If obj.Exist(5) Then D365FO_GetByLabel = Trim(Replace(CStr(obj.GetROProperty("value")), Chr(160), " "))
End Function

' Payment screen: Add, method, amount (balance read from the screen), OK, Submit
' Picks a value in an F&O lookup field found by label: type it (no Tab), click the matching lookup row, else Tab.
Function D365FO_PickLookupByLabel(label, value, stepName)
    Dim obj, dr, oRow, rowX, lc, uc, k
    D365FO_PickLookupByLabel = False
    lc = "abcdefghijklmnopqrstuvwxyz"
    uc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Set obj = D365FO_InputByLabel(label)
    If Not obj.Exist(DefaultTimeout()) Then
        LogFail stepName, "Field with label " & label & " not found"
        Exit Function
    End If
    ' Typing the name resolves to the first key that starts with the text (e.g. 4 = On Account), so open the lookup list and
    ' click the row whose text contains the value (case-insensitive); the list shows the key and the name in separate cells.
    rowX = "(//*[(not(*) and contains(translate(normalize-space(text()),'" & uc & "','" & lc & "'),'" & LCase(value) & "') and not(self::input) and not(self::option) and not(ancestor::*[@data-dyn-role='ActionPane'])) or (self::input and contains(translate(@value,'" & uc & "','" & lc & "'),'" & LCase(value) & "') and ancestor::*[@data-dyn-role='Lookup' or contains(@class,'lookup')])])[last()]"
    For k = 1 To 3
        obj.Click
        Set dr = CreateObject("Mercury.DeviceReplay")
        dr.KeyDown 56 : dr.PressKey 208 : dr.KeyUp 56       ' Alt+Down opens the lookup list
        D365_WaitIdle 20
        Set oRow = GetPage().WebElement("xpath:=" & rowX)
        If oRow.Exist(6) Then Exit For
        ' the lookup button of THIS field (key events such as Alt+Down do not arrive when the desktop session is not interactive)
        Set oRow = GetPage().WebElement("xpath:=(//*[normalize-space(text())='" & label & "' and not(self::input)]/ancestor::*[@data-dyn-controlname][1]//*[contains(@class,'lookupButton')])[last()]")
        If oRow.Exist(3) Then
            oRow.Click
            D365_WaitIdle 20
            Set oRow = GetPage().WebElement("xpath:=" & rowX)
            If oRow.Exist(6) Then Exit For
        End If
        Set oRow = GetPage().WebElement("xpath:=(//*[contains(@class,'lookupButton')])[last()]")
        If oRow.Exist(2) Then oRow.Click
        D365_WaitIdle 20
        Set oRow = GetPage().WebElement("xpath:=" & rowX)
        If oRow.Exist(4) Then Exit For
    Next
    If Not oRow.Exist(1) Then
        LogFail stepName, "Lookup row containing '" & value & "' not found in the " & label & " list. Lookup content: " & D365FO_LookupText()
        Exit Function
    End If
    WebClickRobust oRow, stepName & " - lookup row " & value
    D365_WaitIdle 20
    D365FO_PickLookupByLabel = True
End Function

Function D365_PayCash(method, stepName)
    Dim amount, nm, oSummary, t, st, bal
    D365_PayCash = False
    ' the Balance is shown on the Sales order summary behind the popup: read it before opening the popup
    amount = D365FO_GetByLabel("Balance")
    If Not D365FO_ClickLabel("Add", stepName & " - Add payment") Then Exit Function
    If Not D365FO_PickLookupByLabel("Payment method", method, stepName & " - payment method") Then Exit Function
    nm = D365FO_GetByLabel("Payment method name")
    VerifyTrue (InStr(1, nm, method, vbTextCompare) > 0), stepName & " - payment method name is " & method & " (shows " & nm & ")"
    If InStr(1, nm, method, vbTextCompare) = 0 Then Exit Function
    If amount <> "" Then
        If Not D365FO_SetByLabel("Payment amount", amount, stepName & " - amount " & amount) Then Exit Function
    Else
        LogFail stepName, "Balance could not be read from the Sales order summary"
        Exit Function
    End If
    If Not D365FO_ClickLabel("OK", stepName & " - OK") Then Exit Function
    If Not D365FO_ClickLabel("Submit", stepName & " - Submit") Then Exit Function
    ' wait until the Sales order summary dialog has closed (payment submitted)
    Set oSummary = GetPage().WebElement("xpath:=(//*[normalize-space(text())='Sales order summary'])[1]")
    D365_PayCash = WaitUntilGone(oSummary, 20)
    If Not D365_PayCash Then
        ' After a successful Submit (Payment status Paid, balance 0) the summary stays open with a Close button: close it
        ' The summary dialog has only Submit / Cancel buttons (no Close). Check the payment really went through, then leave the dialog with Cancel.
        st = D365FO_GetByLabel("Payment status")
        bal = D365FO_GetByLabel("Balance")
        VerifyTrue (StrComp(Trim(st), "Paid", vbTextCompare) = 0), stepName & " - Payment status is Paid (shows " & st & ", balance " & bal & ")"
        If StrComp(Trim(st), "Paid", vbTextCompare) = 0 Then
            D365FO_ClickLabel "Cancel", stepName & " - Leave the Sales order summary (payment already submitted)"
            D365_PayCash = WaitUntilGone(oSummary, 40)
        End If
    End If
    D365_WaitIdle 30
    If Not D365_PayCash Then LogFail stepName, "Sales order summary still open after Submit"
End Function

' Work page: Work > Complete work, user id, Validate work, Complete work
Function D365_CompleteWork(userId, stepName)
    D365_CompleteWork = False
    If Not D365FO_ClickLabel("Work", stepName & " - Work tab") Then Exit Function
    If Not D365FO_ClickLabel("Complete work", stepName & " - Complete work") Then Exit Function
    ' TODO(C9652592): verify control name UserId
    If Not D365FO_SetInput("UserId", userId, stepName & " - user id") Then Exit Function
    If Not D365FO_ClickLabel("Validate work", stepName & " - Validate work") Then Exit Function
    VerifyTrue D365FO_WaitText("Complete work", 30), stepName & " - Complete work button displayed after validation"
    If Not D365FO_ClickLabel("Complete work", stepName & " - Complete work button") Then Exit Function
    ' wait for the confirmation (the Processing operation dialog swallows clicks until then)
    VerifyTrue D365FO_WaitTextContains("completed successfully", 90), stepName & " - Work completed successfully"
    D365_CompleteWork = True
End Function


' Opens an action pane (ribbon) tab by its label. The same label also exists on grid menu buttons (e.g. Sales order lines > Warehouse),
' which sit after the ribbon in the page: take the first match that is not inside a grid / section / group.
Function D365FO_ClickRibbonTab(label, stepName)
    Dim obj
    Set obj = GetPage().WebElement("xpath:=(//*[(@data-dyn-role='ActionPaneTab' or @role='tab' or @data-dyn-role='PivotItem') and normalize-space(.)='" & label & "' and not(ancestor::*[@data-dyn-role='Grid' or @data-dyn-role='Group' or @data-dyn-role='Section' or @data-dyn-role='TabPage'])])[1]")
    If obj.Exist(8) Then
        D365FO_ClickRibbonTab = WebClickRobust(obj, stepName)
        D365_WaitIdle 30
    Else
        Set obj = GetPage().WebElement("xpath:=(//*[normalize-space(text())='" & label & "' and not(ancestor::*[@data-dyn-role='Grid' or @data-dyn-role='Group' or @data-dyn-role='Section' or @data-dyn-role='TabPage'])])[1]")
        If obj.Exist(5) Then
            D365FO_ClickRibbonTab = WebClickRobust(obj, stepName)
            D365_WaitIdle 30
        Else
            D365FO_ClickRibbonTab = D365FO_ClickTab(label, stepName)
        End If
    End If
End Function


' Step 9: click Confirm now, refresh and retry until the order is confirmed (Complete shows on the Sell tab). True when it is.
Function D365_ConfirmNowUntilDone(timeoutSec, stepName)
    Dim t
    D365_ConfirmNowUntilDone = False
    t = Timer
    Do
        D365FO_ClickTab "Sell", "Open Sell tab"
        If D365FO_LabelExists("Complete", 2) Then D365_ConfirmNowUntilDone = True: Exit Do
        D365FO_ClickLabel "Confirm now", stepName
        If D365FO_LabelExists("Complete", 15) Then D365_ConfirmNowUntilDone = True: Exit Do
        D365FO_ClickCtl "SystemDefinedRefreshButton", "Refresh", "Refresh the order (wait for Confirm now to be enabled)"
    Loop While Timer - t < timeoutSec And Timer >= t
End Function


' Step 15: the wave runs in a batch, so the work of the order is created some seconds after "Release to warehouse". The Work page opens
' filtered by the order; while it shows "We didn't find anything to show here" refresh it. True when a work row is listed in time.
Function D365_WaitWorkRows(timeoutSec, stepName)
    Dim t, emptyMsg
    D365_WaitWorkRows = False
    ' Match the message on the deepest element whose whole text contains it (the text may be split over child elements / use a typographic apostrophe);
    ' the earlier not(*)+text() match missed it, so an empty Work grid was reported as "work listed" (Complete work was disabled, no User id field).
    emptyMsg = "xpath:=//*[contains(normalize-space(.),'anything to show here') and not(*[contains(normalize-space(.),'anything to show here')])]"
    t = Timer
    Do
        D365_WaitIdle 20
        ' The grid is first rendered without the "We didn't find anything" message while it loads: a single miss is not proof that
        ' work exists (the work page was reported as listed with an empty grid). Require two misses in a row, 5 s apart.
        If Not GetPage().WebElement(emptyMsg).Exist(3) Then
            Wait 5
            D365_WaitIdle 20
            If Not GetPage().WebElement(emptyMsg).Exist(2) Then D365_WaitWorkRows = True: Exit Do
        End If
        D365FO_ClickCtl "SystemDefinedRefreshButton", "Refresh", stepName & " - refresh until the work is created"
    Loop While Timer - t < timeoutSec And Timer >= t
End Function


' Waits until the status cell of the grid shows the text. The cell is not matched by text() alone (grid cells are inputs: value attribute, or
' nested spans), so the same shapes as D365_GetSalesStatus are tried. Refreshes the list between rounds. True when found.
Function D365_WaitStatusText(txt, timeoutSec)
    Dim t, shapes, j, n
    D365_WaitStatusText = False
    t = Timer
    n = 0
    Do
        D365_WaitIdle 10
        shapes = Array("//*[not(*) and normalize-space(text())='" & txt & "']", "//*[@value='" & txt & "']", "//*[normalize-space(.)='" & txt & "' and not(*[normalize-space(.)='" & txt & "'])]", "//*[@title='" & txt & "' or @aria-label='" & txt & "']")
        For j = 0 To UBound(shapes)
            If GetPage().WebElement("xpath:=(" & shapes(j) & ")[last()]").Exist(2) Then D365_WaitStatusText = True: Exit Function
        Next
        n = n + 1
        If n Mod 3 = 0 Then D365FO_ClickCtl "SystemDefinedRefreshButton", "Refresh", "Refresh the order list (wait for status " & txt & ")"
    Loop While Timer - t < timeoutSec And Timer >= t
End Function


'==============================================================================
' C9652910 helpers (additive siblings; the C9652592 helpers above are unchanged)
'==============================================================================

' Create sales order dialog with the header fields of C9652910: customer, customer requisition, mode of delivery, sales origin, OK.
' Optional values ("" = skip). TODO(C9652910): labels / control names are unconfirmed - verify with Object Spy.
Function D365_CreateSalesOrderFull(customer, requisition, deliveryMode, salesOrigin, stepName)
    Dim oSummary
    D365_CreateSalesOrderFull = False
    D365FO_DismissTips
    Set oSummary = GetPage().WebElement("xpath:=(//*[normalize-space(text())='Sales order summary'])[1]")
    If oSummary.Exist(5) Then
        If Not D365FO_ClickLabel("Cancel", stepName & " - cancel Sales order summary") Then Exit Function
    End If
    If Not D365FO_SetInput("SalesTable_CustAccount", customer, stepName & " - customer account") Then Exit Function
    ' The customer defaults (currency, language, invoice account) arrive asynchronously after the customer is set. OK before they are there is
    ' refused ("Field 'Language' must be filled in", dialog stays open), so wait for the currency first.
    D365_WaitFieldFilled "Currency", 45
    If requisition <> "" Then
        If Not D365FO_SetByLabel("Customer requisition", requisition, stepName & " - customer requisition") Then Exit Function
    End If
    If deliveryMode <> "" Then
        If Not D365FO_SetDialogLookup("DlvMode,DeliveryMode", "Mode of delivery", deliveryMode, stepName & " - mode of delivery") Then Exit Function
    End If
    If salesOrigin <> "" Then
        If Not D365FO_SetDialogLookup("SalesOrigin", "Source", salesOrigin, stepName & " - sales origin") Then Exit Function
    End If
    D365_WaitFieldFilled "Currency", 45
    D365_CreateSalesOrderFull = D365FO_ClickLabel("OK", stepName & " - OK")
    ' dialog still open (OK refused because defaults were still loading): wait for them and click OK again
    If D365_CreateSalesOrderFull Then
        D365_WaitIdle 20
        If D365_CreateOrderDialogOpen() Then
            D365_WaitFieldFilled "Currency", 45
            D365_CreateSalesOrderFull = D365FO_ClickLabel("OK", stepName & " - OK (second try)")
        End If
    End If
End Function

' Polls until the input with this label has a value. True when it has.
Function D365_WaitFieldFilled(label, timeoutSec)
    Dim t
    D365_WaitFieldFilled = False
    t = Timer
    Do
        If Trim(D365FO_GetByLabel(label)) <> "" Then D365_WaitFieldFilled = True: Exit Do
        D365_WaitIdle 5
    Loop While Timer - t < timeoutSec And Timer >= t
End Function

' Value of the first input whose AOT control name contains one of the parts (comma-separated), "" when none
Function D365FO_GetInputAny(parts)
    Dim a, i, v
    D365FO_GetInputAny = ""
    a = Split(parts, ",")
    For i = 0 To UBound(a)
        v = D365FO_GetInput(Trim(a(i)))
        If v <> "" Then
            D365FO_GetInputAny = v
            Exit Function
        End If
    Next
End Function

' Release to warehouse answers with an infolog message: "Wave ... has been created for shipment ...", "Shipment ... has been assigned to wave ..."
' or "... submitted to run in a batch" (which one depends on the wave template)
Function D365_WaitWaveMessage(timeoutSec)
    Dim t
    D365_WaitWaveMessage = False
    t = Timer
    Do
        If D365FO_WaitTextContains("has been created for shipment", 3) Or D365FO_WaitTextContains("has been assigned to wave", 1) Or D365FO_WaitTextContains("submitted to run in a batch", 1) Or D365FO_WaitTextContains("Operation completed", 1) Then
            D365_WaitWaveMessage = True
            Exit Function
        End If
    Loop While Timer - t < timeoutSec And Timer >= t
End Function

' Shipment id of the open Shipment details page: the page heading (SHIP-nnnnnn); the AOT control name of the field is not known
Function D365_ReadShipmentId()
    Dim r
    D365_ReadShipmentId = ""
    On Error Resume Next
    r = GetPage().RunScript("(function(){var m=(document.body.innerText||'').match(/SHIP-\d+/);return m?m[0]:'';})()")
    If Err.Number = 0 Then D365_ReadShipmentId = Trim(CStr(r))
    Err.Clear
    On Error GoTo 0
End Function

' Shipment details page of the order (Sales order lines > Warehouse > Shipment details)
Function D365_OpenShipmentDetails(stepName)
    Dim t
    D365_OpenShipmentDetails = False
    ' The shipment is created asynchronously after "Release to warehouse": the menu item stays greyed out (click is a no-op) until then.
    ' Click it, and when no shipment page (heading SHIP-nnnnnn) opens, refresh the order and try again.
    t = Timer
    Do
        If D365FO_ClickLabel("Warehouse", stepName & " - Warehouse dropdown of the Sales order lines") Then
            D365FO_ClickLabel "Shipment details", stepName & " - Shipment details"
            D365_WaitIdle 20
            If D365_WaitShipmentId(8) <> "" Then D365_OpenShipmentDetails = True: Exit Do
        End If
        D365FO_ClickCtl "SystemDefinedRefreshButton", "Refresh", stepName & " - refresh until the shipment exists"
    Loop While Timer - t < 240 And Timer >= t
End Function

Function D365_WaitShipmentId(timeoutSec)
    Dim t, v
    t = Timer
    Do
        v = D365_ReadShipmentId()
        If v <> "" Then Exit Do
    Loop While Timer - t < timeoutSec And Timer >= t
    D365_WaitShipmentId = v
End Function

' First match of a regular expression in the page text (e.g. the wave id in the page heading)
Function D365_ReadFirstMatch(pattern)
    Dim r
    D365_ReadFirstMatch = ""
    On Error Resume Next
    r = GetPage().RunScript("(function(){var m=(document.body.innerText||'').match(/" & pattern & "/);return m?m[0]:'';})()")
    If Err.Number = 0 Then D365_ReadFirstMatch = Trim(CStr(r))
    Err.Clear
    On Error GoTo 0
End Function

' Scrolls every scrollable F&O form container to its bottom so virtualised / below-the-fold sections render
Sub D365_ScrollFormDown()
    On Error Resume Next
    GetPage().RunScript "(function(){var e=document.querySelectorAll('*');for(var i=0;i<e.length;i++){var x=e[i];if(x.scrollHeight>x.clientHeight+20&&x.clientHeight>200){var o=getComputedStyle(x).overflowY;if(o==='auto'||o==='scroll'){x.scrollTop=x.scrollHeight;}}}return 'ok';})()"
    Err.Clear
    On Error GoTo 0
    D365_WaitIdle 10
End Sub

' Shipment details > wave id link > select the shipment line > Process
' TODO(C9652910): wave link / grid selection are unconfirmed - verify with Object Spy
Function D365_ProcessWaveOfShipment(shipmentId, stepName)
    Dim oWave, oRow, oFound, i, waveId, waveOpened
    D365_ProcessWaveOfShipment = False
    Set oWave = GetPage().WebElement("xpath:=(//a[contains(@data-dyn-controlname,'WaveId')] | //*[contains(@data-dyn-controlname,'WaveId')]//a)[last()]")
    If Not oWave.Exist(3) Then
        ' control name of the Wave ID link is not known: it shows the wave number (WAVE-nnnnnn), as a link or as a read-only field value
        Set oWave = GetPage().WebElement("xpath:=(//a[starts-with(normalize-space(.),'WAVE-')] | //input[starts-with(@value,'WAVE-')] | //*[not(*) and starts-with(normalize-space(text()),'WAVE-')])[1]")
    End If
    If Not oWave.Exist(15) Then
        LogFail stepName, "Wave id link not found on the shipment details page"
        Exit Function
    End If
    waveId = ""
    On Error Resume Next
    waveId = Trim(CStr(oWave.GetROProperty("innertext")))
    If waveId = "" Then waveId = Trim(CStr(oWave.GetROProperty("value")))
    Err.Clear
    On Error GoTo 0
    If Not WebClickRobust(oWave, stepName & " - wave id link") Then Exit Function
    D365_WaitIdle 30
    If Not D365FO_WaitText("Waves", 30) Then
        LogFail stepName, "Waves page not displayed"
        Exit Function
    End If
    ' tick the shipment row in the waves lines. The wave form is a details page: the lines section can be below the fold or collapsed,
    ' so scroll down and look for any element showing the shipment id (the wave can hold several shipments).
    Set oRow = Nothing
    For i = 1 To 6
        Set oFound = GetPage().WebElement("xpath:=(//*[@role='row' and contains(.,'" & shipmentId & "')]//*[@role='checkbox' or @data-dyn-role='RowCheckBox' or contains(@class,'row-checkbox')])[last()]")
        If Not oFound.Exist(3) Then Set oFound = GetPage().WebElement("xpath:=(//*[@role='row' and contains(.,'" & shipmentId & "')])[last()]")
        If oFound.Exist(2) Then Set oRow = oFound: Exit For
        D365_ScrollFormDown
    Next
    If oRow Is Nothing Then
        ' The shipment is not in the DOM: accept only when the wave opened is the very wave this shipment is assigned to
        ' (wave id read from the shipment details page / release message equals the wave in the page heading)
        waveOpened = D365_ReadFirstMatch("WAVE-\d+")
        VerifyTrue (waveId <> "" And StrComp(waveOpened, waveId, vbTextCompare) = 0), stepName & " - opened wave " & waveOpened & " is the wave of shipment " & shipmentId & " (" & waveId & ")"
        If waveId = "" Or StrComp(waveOpened, waveId, vbTextCompare) <> 0 Then
            LogFail stepName, "Shipment " & shipmentId & " not listed in the wave lines"
            Exit Function
        End If
    Else
        WebClickRobust oRow, stepName & " - select shipment " & shipmentId
    End If
    If Not D365FO_ClickLabel("Process", stepName & " - Process") Then Exit Function
    D365_ProcessWaveOfShipment = True
End Function

' Opens the lookup of the field with the label and picks the FIRST row (data not named by the case: found at run time).
' Used for the Pack station Location and Packing profile ID. True when the field has a value afterwards.
Function D365FO_PickFirstLookupByLabel(label, stepName)
    Dim obj, dr, oRow, k, rowX
    D365FO_PickFirstLookupByLabel = False
    Set obj = D365FO_InputByLabel(label)
    If Not obj.Exist(DefaultTimeout()) Then
        LogFail stepName, "Field with label " & label & " not found"
        Exit Function
    End If
    rowX = "(//*[@data-dyn-role='Lookup' or contains(@class,'lookup')]//*[(@role='row' or @data-dyn-role='Row') and not(.//*[@role='columnheader'])]//*[not(*) and normalize-space(text())!=''])[1]"
    For k = 1 To 3
        obj.Click
        Set dr = CreateObject("Mercury.DeviceReplay")
        dr.KeyDown 56 : dr.PressKey 208 : dr.KeyUp 56       ' Alt+Down opens the lookup list
        D365_WaitIdle 20
        Set oRow = GetPage().WebElement("xpath:=" & rowX)
        If oRow.Exist(5) Then
            WebClickRobust oRow, stepName & " - first lookup row"
            D365_WaitIdle 20
        Else
            ' the lookup DOM is not known: keyboard fallback, Down selects the first row and Enter takes it
            dr.PressKey 208
            dr.PressKey 28
            D365_WaitIdle 20
        End If
        If D365FO_GetByLabel(label) <> "" Then
            D365FO_PickFirstLookupByLabel = True
            LogPass stepName, label & " set at run time to " & D365FO_GetByLabel(label)
            Exit Function
        End If
    Next
    LogFail stepName, "No value could be picked in the " & label & " lookup. Lookup content: " & D365FO_LookupText()
End Function

' Pack: opened with the D365 navigation search, confirms the dialog (site / warehouse must be the allocated ones), enters the shipment id.
' TODO(C9652910): every label / control name below is unconfirmed - verify with Object Spy.
Function D365_PackShipment(shipmentId, expectedSite, expectedWarehouse, stepName)
    Dim oSearch, dr, v, w
    D365_PackShipment = False
    ' Step 23: search "Pack" in the NAVIGATION search (magnifier in the top bar). Alt+Q opened the page's own action search on the Work
    ' page ("We didn't find anything to show here"), so click the top bar icon (any search button in the shell header, y < 70).
    On Error Resume Next
    GetPage().RunScript "(function(){var b=document.querySelectorAll('button,a,[role=button],[role=search]');for(var i=0;i<b.length;i++){var x=b[i];var l=((x.getAttribute('aria-label')||'')+' '+(x.getAttribute('title')||'')+' '+(x.getAttribute('data-dyn-controlname')||'')+' '+(x.id||'')).toLowerCase();var r=x.getBoundingClientRect();if(l.indexOf('search')>=0&&r.top<70&&r.width>0){x.click();return 'ok';}}return 'none';})()"
    Err.Clear
    On Error GoTo 0
    D365_WaitIdle 10
    Set dr = CreateObject("Mercury.DeviceReplay")
    dr.SendString "Pack"
    D365_WaitIdle 10
    Set oSearch = GetPage().WebElement("xpath:=(//*[not(*) and normalize-space(text())='Pack' and ancestor::*[contains(@class,'search') or contains(@data-dyn-controlname,'earch')]])[1]")
    If Not oSearch.Exist(10) Then
        LogFail stepName, "Pack not found in the navigation search"
        Exit Function
    End If
    WebClickRobust oSearch, stepName & " - open Pack"
    D365_WaitIdle 30
    ' Step 24: site / warehouse / location / profile. The "Select packing station" dialog opens EMPTY here (screenshot), so the allocated
    ' site and warehouse are entered, then Location and Packing profile ID (not named by the case) are picked at run time (first lookup row).
    If D365FO_GetByLabel("Site") = "" And expectedSite <> "" Then
        If Not D365FO_PickLookupByLabel("Site", expectedSite, stepName & " - Site") Then Exit Function
    End If
    If D365FO_GetByLabel("Warehouse") = "" And expectedWarehouse <> "" Then
        If Not D365FO_PickLookupByLabel("Warehouse", expectedWarehouse, stepName & " - Warehouse") Then Exit Function
    End If
    If D365FO_GetByLabel("Location") = "" Then
        If Not D365FO_PickFirstLookupByLabel("Location", stepName & " - Location") Then Exit Function
    End If
    ' The profile is not named by the case. The first profile may be refused by the dialog ("The selected packing profile has a container
    ' close profile ..." - the dialog stays open and the Pack page never shows the Shipment id field), so profiles are tried one by one below.
    ' Reviewer / case step 23: the packing profile id is "FL-TAM Default".
    If StrComp(D365FO_GetByLabel("Packing profile ID"), "FL-TAM Default", vbTextCompare) <> 0 Then
        ' The lookup list of this field did not open / render rows (screenshot: field stays empty, lookup content empty). Type the id first
        ' (the field validates it against the profiles), and use the lookup list only when typing did not resolve it.
        D365FO_FocusTypeVerify D365FO_InputByLabel("Packing profile ID"), "FL-TAM Default"
        If StrComp(D365FO_GetByLabel("Packing profile ID"), "FL-TAM Default", vbTextCompare) <> 0 Then
            If Not D365FO_PickLookupByLabel("Packing profile ID", "FL-TAM Default", stepName & " - Packing profile ID") Then Exit Function
        End If
    End If
    v = D365FO_GetByLabel("Site")
    w = D365FO_GetByLabel("Warehouse")
    If expectedSite <> "" Then VerifyEquals v, expectedSite, stepName & " - Pack dialog site is the allocated site"
    If expectedWarehouse <> "" Then VerifyEquals w, expectedWarehouse, stepName & " - Pack dialog warehouse is the allocated warehouse"
    LogInfo stepName, "Pack dialog site=" & v & " warehouse=" & w & " location=" & D365FO_GetByLabel("Location") & " profile=" & D365FO_GetByLabel("Packing profile ID")
    If Not D365FO_ClickLabel("OK", stepName & " - OK on the Pack dialog") Then Exit Function
    ' The Pack page is open when the Shipment id field exists; else the dialog refused the profile: try the next profile of the lookup
    D365_WaitIdle 20
    If Not D365FO_InputByLabel("License plate or shipment").Exist(8) Then
        If Not D365_TryNextPackingProfile(stepName) Then Exit Function
    End If
    ' Step 26: shipment id
    ' The Pack page (screenshot) has the editable field "License plate or shipment"; "Shipment ID" is read-only and is filled once the id is entered.
    If Not D365FO_SetByLabel("License plate or shipment", shipmentId, stepName & " - shipment id") Then Exit Function
    D365_WaitIdle 20
    VerifyTrue D365_WaitFieldValue("Shipment ID", shipmentId, 20), stepName & " - shipment " & shipmentId & " details are loaded on the Pack page"
    D365_PackShipment = True
End Function

' Packing profile refused by the Select packing station dialog (it stays open with the warning "The selected packing profile has a container
' close profile that does not match the packing station location"): log the full message, then pick the other profiles of the lookup one by one WITH THE MOUSE (no Enter: Enter / a click on
' the blue profile text drills down into the Packing profiles setup form) and click OK until the Pack page (Shipment id field) opens.
' TODO(C9652910): which profile belongs to warehouse 420 is not known - verify with a person.
Function D365_TryNextPackingProfile(stepName)
    Dim n, dr, v, tried, msg, oRow, rowX, cur
    D365_TryNextPackingProfile = False
    tried = "|" & D365FO_GetByLabel("Packing profile ID") & "|"
    On Error Resume Next
    msg = GetPage().RunScript("(function(){var all=document.querySelectorAll('*');var best='';for(var i=0;i<all.length;i++){var t=(all[i].innerText||'');if(t.indexOf('container close profile')>=0&&(best===''||t.length<best.length))best=t;}return best.replace(/\s+/g,' ');})()")
    Err.Clear
    On Error GoTo 0
    LogInfo stepName, "Pack dialog message: " & msg
    For n = 1 To 40
        ' the dialog must still be open, else this is not the refused-profile case
        If D365FO_InputByLabel("License plate or shipment").Exist(1) Then
            D365_TryNextPackingProfile = True
            Exit Function
        End If
        If Not GetPage().WebElement("xpath:=(//*[not(*) and normalize-space(text())='Select packing station'])[1]").Exist(2) Then Exit For
        Set dr = CreateObject("Mercury.DeviceReplay")
        ' NEVER click the field: with a value its blue text is a drill-down link and a click opens the Packing profiles setup form.
        ' Focus it (no click), empty it (Ctrl+A, Delete) and open the lookup with Alt+Down.
        On Error Resume Next
        D365FO_InputByLabel("Packing profile ID").Object.focus
        Err.Clear
        On Error GoTo 0
        dr.KeyDown 29 : dr.PressKey 30 : dr.KeyUp 29        ' Ctrl+A
        dr.PressKey 211                                     ' Delete
        dr.KeyDown 56 : dr.PressKey 208 : dr.KeyUp 56       ' Alt+Down opens the lookup list
        D365_WaitIdle 20
        rowX = "(//*[@data-dyn-role='Lookup' or contains(@class,'lookup')]//*[(@role='row' or @data-dyn-role='Row') and not(.//*[@role='columnheader'])])[" & n & "]//*[not(*) and normalize-space(text())!=''][1]"
        Set oRow = GetPage().WebElement("xpath:=" & rowX)
        If Not oRow.Exist(5) Then
            dr.PressKey 1   ' Esc closes the lookup
            Exit For
        End If
        cur = Trim(Replace(CStr(oRow.GetROProperty("innertext")), Chr(160), " "))
        If InStr(tried, "|" & cur & "|") > 0 Then
            dr.PressKey 1
        Else
            WebClickRobust oRow, stepName & " - lookup row " & n
            D365_WaitIdle 20
            v = D365FO_GetByLabel("Packing profile ID")
            If v <> "" And InStr(tried, "|" & v & "|") = 0 Then
                tried = tried & v & "|"
                LogInfo stepName, "Packing profile refused, trying '" & v & "'"
                If D365FO_ClickLabel("OK", stepName & " - OK on the Pack dialog with profile " & v) Then
                    D365_WaitIdle 20
                    If D365FO_InputByLabel("License plate or shipment").Exist(8) Then
                        LogPass stepName, "Pack page opened with packing profile " & v
                        D365_TryNextPackingProfile = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    LogFail stepName, "No packing profile was accepted by the Pack dialog (tried " & tried & "), message: " & msg
End Function

' Pack page: New container, OK, select the open lines, Pack, Close container with weight / dimensions, OK
Function D365_PackContainer(stepName)
    D365_PackContainer = False
    If Not D365FO_ClickLabel("New container", stepName & " - New container") Then Exit Function
    If Not D365FO_ClickLabel("OK", stepName & " - OK (container information)") Then Exit Function
    ' Step 29: both lines of the Open lines section (header checkbox of that grid; the grid has no "Select all" button)
    D365_SelectAllOpenLines stepName & " - select both open lines"
    If Not D365FO_ClickLabel("Pack", stepName & " - Pack") Then Exit Function
    VerifyTrue D365FO_WaitText("All lines", 20), stepName & " - lines moved to the All lines section"
    If Not D365FO_ClickLabel("Close container", stepName & " - Close container") Then Exit Function
    D365FO_ClickLabel "Get system weight", stepName & " - Get system weight"
    ' The button is named "Get System Dimensions" in the dialog
    If Not D365FO_ClickLabel("Get System Dimensions", stepName & " - Get system dimension") Then D365FO_ClickLabel "Get system dimension", stepName & " - Get system dimension"
    ' Weight must be positive to close the container. TODO(C9652910): the system weight of these items is 0.00 on stage (screenshot), 1 is used - needs a person.
    If D365_IsZeroValue(D365FO_GetByLabel("Weight")) Then D365FO_FocusTypeVerify D365FO_InputByLabel("Weight"), "1"
    ' dimensions still blank (0.00): 10 x 10 x 10 as in the case
    If D365_IsZeroValue(D365FO_GetByLabel("Length")) Then D365FO_SetByLabel "Length", "10", stepName & " - length"
    If D365_IsZeroValue(D365FO_GetByLabel("Width")) Then D365FO_SetByLabel "Width", "10", stepName & " - width"
    If D365_IsZeroValue(D365FO_GetByLabel("Height")) Then D365FO_SetByLabel "Height", "10", stepName & " - height"
    D365_PackContainer = D365FO_ClickLabel("OK", stepName & " - OK (close container)")
End Function

' Lookup field by label: scroll it into view, type the value + Tab (F&O resolves the key / name), check the field shows it;
' only when that fails open the lookup list and click the row (D365FO_PickLookupByLabel). The dialog fields below the fold are not in view.
Function D365FO_SetLookupTyped(label, value, stepName)
    Dim obj, actual
    D365FO_SetLookupTyped = False
    Set obj = D365FO_InputByLabel(label)
    If Not obj.Exist(DefaultTimeout()) Then
        LogFail stepName, "Field with label " & label & " not found"
        Exit Function
    End If
    ScrollIntoView obj
    If D365FO_TypeAndVerify(obj, value) Then
        D365FO_SetLookupTyped = True
        LogPass stepName, value & " entered in field " & label
        Exit Function
    End If
    D365FO_SetLookupTyped = D365FO_PickLookupByLabel(label, value, stepName)
End Function

' Sales origin in the Create sales order dialog: the dialog labels the field "Source" (the case says "sales origin").
' Tries the AOT control name (SalesOrigin), then the labels "Sales origin" and "Source".
' TODO(C9652910): confirm control name with Object Spy
Function D365FO_SetSalesOrigin(value, stepName)
    Dim obj
    D365FO_SetSalesOrigin = False
    Set obj = GetPage().WebEdit("xpath:=(//*[contains(@data-dyn-controlname,'SalesOrigin')]//input[not(@type='hidden') and not(@type='checkbox')])[last()]")
    If obj.Exist(5) Then
        ScrollIntoView obj
        If D365FO_TypeAndVerify(obj, value) Then
            D365FO_SetSalesOrigin = True
            LogPass stepName, value & " entered in SalesOrigin"
            Exit Function
        End If
    End If
    If D365FO_InputByLabel("Sales origin").Exist(3) Then
        D365FO_SetSalesOrigin = D365FO_SetLookupTyped("Sales origin", value, stepName)
    Else
        D365FO_SetSalesOrigin = D365FO_SetLookupTyped("Source", value, stepName)
    End If
End Function

' Lookup field of the open Create sales order dialog, found by AOT control name (comma-separated parts) INSIDE the dialog.
' A label search is not used: the label "Mode of delivery" also matched a link-style control that navigates to the Modes of delivery page
' when clicked (the typed text then went into that page's search box). After typing, the page must still be the order page.
' TODO(C9652910): confirm the control names (DlvMode / SalesOrigin) with Object Spy.
Function D365FO_SetDialogLookup(ctlParts, label, value, stepName)
    Dim parts, i, obj, url, found
    D365FO_SetDialogLookup = False
    parts = Split(ctlParts, ",")
    found = False
    For i = 0 To UBound(parts)
        Set obj = GetPage().WebEdit("xpath:=(//*[(@data-dyn-role='Dialog' or @role='dialog')]//*[contains(@data-dyn-controlname,'" & Trim(parts(i)) & "')]//input[not(@type='hidden') and not(@type='checkbox')])[last()]")
        If obj.Exist(4) Then
            found = True
            Exit For
        End If
    Next
    If Not found Then
        ' control name differs (e.g. the dialog field "Source"): input of the control that holds the visible label, inside the dialog
        Set obj = GetPage().WebEdit("xpath:=(//*[(@data-dyn-role='Dialog' or @role='dialog')]//*[not(*) and normalize-space(text())='" & label & "']/ancestor::*[@data-dyn-controlname][1]//input[not(@type='hidden') and not(@type='checkbox')])[last()]")
        If obj.Exist(4) Then found = True
    End If
    If Not found Then
        LogFail stepName, "Field " & label & " (control " & ctlParts & ") not found in the Create sales order dialog. Control names: " & D365FO_ControlNamesLike("S")
        Exit Function
    End If
    ScrollIntoView obj
    If Not D365FO_FocusTypeVerify(obj, value) Then
        LogFail stepName, "Field " & label & " does not show " & value & " after typing"
        Exit Function
    End If
    url = ""
    On Error Resume Next
    url = CStr(GetPage().GetROProperty("url"))
    Err.Clear
    On Error GoTo 0
    If InStr(1, url, "mi=SalesTableListPage", vbTextCompare) = 0 And InStr(1, url, "mi=SalesTableDetails", vbTextCompare) = 0 Then
        LogFail stepName, "Page navigated away while setting " & label & " (URL " & url & ")"
        Exit Function
    End If
    D365FO_SetDialogLookup = True
    LogPass stepName, value & " entered in field " & label
End Function

' Types into a field WITHOUT clicking it. A lookup that already holds a value (here the default mode of delivery of the customer) is shown
' as a link: a click on it opens the main record (Modes of delivery page). Focus by script, select all, type, Tab, read back.
Function D365FO_FocusTypeVerify(obj, value)
    Dim dr, actual
    On Error Resume Next
    obj.Object.focus
    obj.Object.select
    Err.Clear
    On Error GoTo 0
    Set dr = CreateObject("Mercury.DeviceReplay")
    dr.KeyDown 29 : dr.PressKey 30 : dr.KeyUp 29       ' Ctrl+A
    dr.SendString CStr(value)
    Wait 0, 500
    dr.PressKey 15                                      ' Tab
    D365_WaitIdle 30
    actual = Trim(Replace(CStr(obj.GetROProperty("value")), Chr(160), " "))
    If Not D365FO_ValueMatches(actual, CStr(value)) Then
        ' key events did not arrive (e.g. the desktop session is not interactive): set the value through UFT / the DOM, then Tab
        On Error Resume Next
        obj.Set CStr(value)
        If Err.Number <> 0 Then
            Err.Clear
            WebSetDomValue obj, CStr(value)
        End If
        Err.Clear
        On Error GoTo 0
        Wait 0, 500
        dr.PressKey 15
        D365_WaitIdle 30
        actual = Trim(Replace(CStr(obj.GetROProperty("value")), Chr(160), " "))
    End If
    D365FO_FocusTypeVerify = D365FO_ValueMatches(actual, CStr(value))
End Function

' Allocation ids of the order lines (grid column SNPOUniteMAOAllocationId): the non-empty cell values, "|"-separated. Cells can be inputs or
' plain text, so both are read; header cells are skipped.
Function D365_GetAllocationIds()
    Dim mx, pos, best, r
    D365_GetAllocationIds = ""
    best = ""
    ' the F&O grid only renders the columns that are on screen: scroll the lines grid sideways and read the column below its header
    mx = D365_ScrollLinesGrid(0)
    pos = 0
    Do
        r = D365_ReadAllocationCells()
        If Len(r) > Len(best) Then best = r
        If best <> "" Then
            If UBound(Split(best, "|")) + 1 >= 2 Then Exit Do
        End If
        If pos >= mx Then Exit Do
        pos = pos + 400
        D365_ScrollLinesGrid pos
        Wait 1   ' TODO(C9652910): the grid renders newly visible columns asynchronously (short poll, no better signal on the grid)
    Loop
    D365_ScrollLinesGrid 0
    D365_GetAllocationIds = best
End Function

' Reads the values under the column header "SNPOUniteMAOAllocationId" (cells have no AOT control name containing AllocationId, so they are found by
' position under the header; confirmed on the run screenshot): one value per line that has an id, "|"-separated.
Function D365_ReadAllocationCells()
    Dim r
    D365_ReadAllocationCells = ""
    On Error Resume Next
    r = GetPage().RunScript("(function(){var root=document;var all=document.querySelectorAll('*');var hdr=null,ha=1e12;for(var i=0;i<all.length;i++){var e=all[i];if(e.children.length>2)continue;var t=(e.innerText||'').replace(/\s+/g,' ').trim();if(t.length>0&&t.length<45&&t.toLowerCase().indexOf('allocationid')>=0&&t.toLowerCase().indexOf('force')<0){var r=e.getBoundingClientRect();if(r.width>0&&r.height>0&&r.width*r.height<ha){ha=r.width*r.height;hdr=e;}}}if(!hdr)return '';var hr=hdr.getBoundingClientRect();var g=hdr.closest('[data-dyn-role=Grid]')||hdr.closest('[role=grid]')||document;var els=g.querySelectorAll('*');var rows={};for(var j=0;j<els.length;j++){var c=els[j];if(c===hdr||c.contains(hdr)||hdr.contains(c))continue;var rc=c.getBoundingClientRect();if(rc.width<=0||rc.height<=0||rc.top<hr.bottom-2)continue;var cx=rc.left+rc.width/2;if(cx<hr.left||cx>hr.right)continue;var v='';if(c.tagName==='INPUT')v=c.value;else if(c.children.length===0)v=c.innerText||c.textContent||'';v=(v||'').replace(/\s+/g,' ').trim();if(!v)continue;var k=Math.round(rc.top/10);if(!rows[k])rows[k]=v;}var o=[];Object.keys(rows).sort(function(x,y){return x-y;}).forEach(function(k){o.push(rows[k]);});return o.join('|');})()")
    If Err.Number = 0 Then D365_ReadAllocationCells = CStr(r)
    Err.Clear
    On Error GoTo 0
End Function

' Scrolls the sales order lines grid horizontally to pos (px) and returns the maximum scroll position. Grid scrollers are found inside the
' grid control (SalesLineGrid / any data-dyn-role Grid) as the elements wider than their box.
Function D365_ScrollLinesGrid(pos)
    Dim r
    D365_ScrollLinesGrid = 0
    On Error Resume Next
    r = GetPage().RunScript("(function(){var g=document.querySelectorAll('[data-dyn-controlname=""SalesLineGrid""],[data-dyn-controlname=""Grid""],[data-dyn-role=""Grid""]');var mx=0;for(var k=0;k<g.length;k++){var all=g[k].querySelectorAll('*');for(var i=0;i<all.length;i++){var e=all[i];if(e.scrollWidth>e.clientWidth+20&&e.clientWidth>200){var m=e.scrollWidth-e.clientWidth;if(" & CLng(pos) & "===0)e.scrollLeft=0;else e.scrollLeft=Math.min(" & CLng(pos) & ",m);if(m>mx)mx=m;}}}return mx;})()")
    If Err.Number = 0 Then
        If IsNumeric(r) Then D365_ScrollLinesGrid = CLng(r)
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Polls (refresh between rounds) until at least minCount allocation ids are filled in the lines grid; returns the ids ("" when not in time)
Function D365_WaitAllocationIds(minCount, timeoutSec)
    Dim t, v
    D365_WaitAllocationIds = ""
    t = Timer
    Do
        D365_WaitIdle 10
        v = D365_GetAllocationIds()
        If v <> "" Then
            If UBound(Split(v, "|")) + 1 >= minCount Then D365_WaitAllocationIds = v: Exit Do
        End If
        D365FO_ClickCtl "SystemDefinedRefreshButton", "Refresh", "Refresh the order (wait for the MAO allocation ids)"
    Loop While Timer - t < timeoutSec And Timer >= t
    If D365_WaitAllocationIds = "" Then
        LogInfo "Allocation id", "Ids read at the end: '" & v & "' / controls: " & D365FO_ControlNamesLike("llocation")
    End If
End Function


' True when the text is empty or a zero number (0, 0.00, 0,00)
Function D365_IsZeroValue(v)
    v = Trim(Replace(CStr(v), Chr(160), ""))
    D365_IsZeroValue = (v = "" Or v = "0" Or v = "0.00" Or v = "0,00")
End Function

' Polls the value of the input with this label (read-only / link fields are inputs: their value is not page text)
Function D365_WaitFieldValue(label, expected, timeoutSec)
    Dim t
    D365_WaitFieldValue = False
    t = Timer
    Do
        If StrComp(Trim(D365FO_GetByLabel(label)), expected, vbTextCompare) = 0 Then D365_WaitFieldValue = True: Exit Do
        D365_WaitIdle 5
    Loop While Timer - t < timeoutSec And Timer >= t
End Function

' Step 29: tick every row of the Open lines grid with the grid's header (select all) checkbox.
' TODO(C9652910): verify the header checkbox object with Object Spy.
Function D365_SelectAllOpenLines(stepName)
    Dim shapes, i, obj
    D365_SelectAllOpenLines = False
    shapes = Array( _
        "(//*[normalize-space(text())='Open lines']/ancestor::*[@data-dyn-controlname][1]//*[@data-dyn-role='Grid' or @role='grid']//*[@role='checkbox' and (ancestor::*[@role='columnheader'] or ancestor::*[contains(@class,'header')])])[1]", _
        "(//*[normalize-space(text())='Open lines']/ancestor::*[@data-dyn-controlname][1]//*[contains(@class,'dyn-headerCheckbox') or contains(@class,'headerCheckbox') or contains(@class,'selectAll') or contains(@class,'select-all')])[1]", _
        "(//*[normalize-space(text())='Open lines']/ancestor::*[@data-dyn-controlname][1]//*[@role='columnheader']//*[@role='checkbox' or contains(@class,'checkbox') or contains(@class,'RowMark') or contains(@class,'rowMark')])[1]")
    For i = 0 To UBound(shapes)
        Set obj = GetPage().WebElement("xpath:=" & shapes(i))
        If obj.Exist(4) Then
            D365_SelectAllOpenLines = WebClickRobust(obj, stepName)
            D365_WaitIdle 20
            Exit Function
        End If
    Next
    D365_SelectAllOpenLines = D365FO_ClickLabel("Select all", stepName)
End Function

' Step 36: the Shipment ID field of the Pack page is a link inside an input; click it (else the Shipment details button) to open the shipment.
Function D365_OpenShipmentFromPack(shipmentId, stepName)
    Dim obj
    D365_OpenShipmentFromPack = False
    Set obj = D365FO_InputByLabel("Shipment ID")
    If obj.Exist(DefaultTimeout()) Then
        WebClickRobust obj, stepName
        D365_WaitIdle 30
    End If
    If Not D365FO_LabelExists("Confirm shipment", 10) Then
        If Not D365FO_ClickLabel("Shipment details", stepName & " - Shipment details button") Then Exit Function
    End If
    D365_OpenShipmentFromPack = D365FO_LabelExists("Confirm shipment", 20)
    If Not D365_OpenShipmentFromPack Then LogFail stepName, "Shipment details page of " & shipmentId & " did not open"
End Function


' True while the Create sales order dialog is still open
Function D365_CreateOrderDialogOpen()
    D365_CreateOrderDialogOpen = GetPage().WebElement("xpath:=(//*[normalize-space(text())='Create sales order'])[1]").Exist(3)
End Function

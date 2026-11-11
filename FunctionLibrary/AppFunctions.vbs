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
        ' the lookup button next to the input
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
    Dim amount, nm, oSummary, t
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
        If D365FO_LabelExists("Close", 3) Then D365FO_ClickLabel "Close", stepName & " - Close the Sales order summary"
        D365_PayCash = WaitUntilGone(oSummary, 40)
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
    Dim t
    D365_WaitWorkRows = False
    t = Timer
    Do
        D365_WaitIdle 20
        If Not GetPage().WebElement("xpath:=//*[not(*) and contains(normalize-space(text()),'find anything to show here')]").Exist(3) Then D365_WaitWorkRows = True: Exit Do
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

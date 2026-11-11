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


'------------------------------------------------------------------------------
' TC_C9652593_CancelSalesOrder   (TestRail C9652593 - Cancel a Sales Order)
'   TestData:  salesId:<optional fixed order number>
'   Without salesId the script picks, at run time, the LAST order with status "Open order" in All sales orders
'   (the cancel is irreversible, so a fresh order is taken every run). Needs an Open order with no lines
'   processed for delivery, the legal entity (D365_COMPANY in .env) and a user allowed to cancel.
'------------------------------------------------------------------------------
Sub TC_C9652593_CancelSalesOrder(dictData)
    Dim salesId, status, msg, expectedMsg
    salesId = GetData(dictData, "salesId", "")
    expectedMsg = "Do you want to cancel all line quantities that are not processed for delivery in the selected sales orders?"

    status = D365_EnsureLoggedIn()
    RunStep "Login to Dynamics 365", status
    If Not status Then Exit Sub

    ' Step 1: Sales and marketing > All sales orders
    status = D365_OpenMenuItem("SalesTableListPage", "Step 1 - Navigate to Sales and marketing > All sales orders")
    If Not status Then Exit Sub

    ' Step 2: select the last sales order with status Open order (found at run time)
    If salesId = "" Or Left(salesId, 1) = "<" Then
        salesId = D365_FindLastOpenSalesOrder("Step 2 - Find last Open order")
        If salesId = "" Then Exit Sub
    End If
    status = D365_SelectSalesOrder(salesId, "Step 2 - Select sales order " & salesId)
    If Not status Then Exit Sub

    ' Step 3: Sales order > Maintain > Cancel, expect the Yes/No message
    status = D365_ClickMaintainCancel("Step 3 - Maintain > Cancel")
    If Not status Then Exit Sub
    msg = D365_GetDialogText()
    VerifyContains msg, expectedMsg, "Step 3 - Cancel confirmation message is displayed"

    ' Step 4: Yes, order is cancelled
    status = D365_ClickDialogButton("Yes", "Step 4 - Click Yes")
    If Not status Then Exit Sub
    VerifyContains D365_GetSalesStatus(), "Cancel", "Step 4 - Sales order status is Cancelled"
End Sub


'------------------------------------------------------------------------------
' TC_C9652592_CreateSalesOrder   (TestRail C9652592 - Creation of Sales Order, order to delivery)
'   TestData (all optional, defaults = the values of the case):
'     customer:S005471;item:RD-RS-LA;qty:2;payMethod:cash;workUser:rohita
'   Found at run time: sales order number (order header), payment amount (balance on the payment screen),
'   wave / work / shipment (through the new order). Needs stock for the item, the MAO allocation,
'   the legal entity (D365_COMPANY in .env) and a user allowed to sell, take payment, release and post.
'   TODO(C9652592): every control name / label is unconfirmed - verify with Object Spy.
'------------------------------------------------------------------------------
Sub TC_C9652592_CreateSalesOrder(dictData)
    Dim customer, item, qty, payMethod, workUser, status, salesId, alloc

    customer = GetData(dictData, "customer", "S005471")
    item = GetData(dictData, "item", "RD-RS-LA")
    qty = GetData(dictData, "qty", "2")
    payMethod = GetData(dictData, "payMethod", "cash")
    workUser = GetData(dictData, "workUser", "rohita")

    status = D365_EnsureLoggedIn()
    RunStep "Login to Dynamics 365", status
    If Not status Then Exit Sub

    ' Step 1: Sales and marketing > All sales orders
    status = D365_OpenMenuItem("SalesTableListPage", "Step 1 - Navigate to Sales and marketing > All sales orders")
    If Not status Then Exit Sub

    ' Step 2: New -> Create sales order dialog (the Sales order summary panel is cancelled when shown)
    status = D365FO_ClickCtl("SystemDefinedNewButton", "New", "Step 2 - Click New")
    If Not status Then Exit Sub

    ' Steps 3-4: customer account, OK -> New sales order page
    status = D365_CreateSalesOrder(customer, "Steps 3-4 - Create order for " & customer)
    If Not status Then Exit Sub
    salesId = D365FO_GetSalesIdFromHeader()
    VerifyTrue (salesId <> ""), "Step 4 - New sales order page opened (order '" & salesId & "')"
    If salesId = "" Then Exit Sub

    ' Steps 5-6: Sales order lines > Add line, item + quantity
    status = D365_AddSalesLine(item, qty, "Steps 5-6 - Add sales line")
    If Not status Then Exit Sub
    ' TODO(C9652592): verify the populated item field (site)
    VerifyTrue (D365FO_GetInput("InventSiteId") <> ""), "Step 6 - Item related fields (site) are populated"

    ' Step 7: the SNPOUniteMAOAllocationId column is in the grid below the Sales order lines section (quick check, no tab search)
    VerifyTrue D365_AllocationColumnVisible(15), "Step 7 - Column SNPOUniteMAOAllocationId is visible in the sales order lines grid"

    ' Steps 8-9: Sell > Confirm now
    status = D365FO_ClickLabel("Sell", "Step 8 - Click Sell")
    If Not status Then Exit Sub
    ' Step 9: the case says "Confirm now (if not enabled click refresh until it gets enabled)". The enabled state is not readable from the DOM
    ' (the aria-disabled check passed while the button was still greyed out, the click was a no-op and Complete never came), so use the
    ' visible effect: click, wait for Complete (only shown once the order is confirmed), else refresh and click again.
    status = D365_ConfirmNowUntilDone(150, "Step 9 - Click Confirm now")
    VerifyTrue status, "Step 9 - Confirm now clicked once enabled and the order is confirmed (Complete is available)"
    If Not status Then Exit Sub

    ' Step 10: Complete -> payment screen
    ' TODO(C9652592): the tab holding Complete is not named in the case; the likely tabs are tried
    ' Root cause of the failure: Complete only shows on the Sell tab after the confirmation has finished (it was found there in the earlier
    ' runs after a long wait + refresh). Poll: refresh the order, go back to Sell, look for Complete; other tabs only as last resort.
    D365_WaitCompleteButton 90
    If Not D365FO_LabelExists("Complete", 2) Then D365FO_OpenTabWith "Complete", Array("Commerce", "Sales order", "General", "Manage")
    status = D365FO_ClickLabel("Complete", "Step 10 - Click Complete")
    If Not status Then Exit Sub
    VerifyObjectExist GetPage().WebElement("xpath:=(//*[normalize-space(text())='Payments'])[1]"), 20, "Step 10 - Payment screen is displayed"

    ' Steps 11-13: Payments > Add, cash, balance, OK, Submit
    status = D365_PayCash(payMethod, "Steps 11-13 - Cash payment")
    If Not status Then Exit Sub

    ' Step 14: Warehouse > Release to warehouse -> wave
    status = D365FO_ClickRibbonTab("Warehouse", "Step 14 - Warehouse tab")
    If Not status Then Exit Sub
    status = D365FO_ClickLabel("Release to warehouse", "Step 14 - Release to warehouse")
    If Not status Then Exit Sub
    ' TODO(C9652592): verify the wave confirmation text
    VerifyTrue D365FO_WaitTextContains("has been submitted to run in a batch", 60), "Step 14 - Wave generated (message: Wave ... has been submitted to run in a batch)"

    ' Step 15: Warehouse > Work details -> Work page
    ' Root cause of the failure: the ribbon Warehouse tab only holds Loads + Release to warehouse. Work details is in the Warehouse dropdown
    ' of the Sales order lines grid (Shipment details / Load details / Work details), so use that button here (last match = grid button).
    status = D365FO_ClickLabel("Warehouse", "Step 15 - Warehouse dropdown of the Sales order lines")
    If Not status Then Exit Sub
    status = D365FO_ClickLabel("Work details", "Step 15 - Work details")
    If Not status Then Exit Sub
    VerifyObjectExist GetPage().WebElement("xpath:=(//*[@data-dyn-controlname='WHSWorkTable' or normalize-space(text())='Work'])[1]"), 20, "Step 15 - Work page is open"

    ' Root cause of the failure: the wave runs in a batch, the Work page was opened before the work existed (empty grid, Complete work disabled).
    VerifyTrue D365_WaitWorkRows(180, "Step 15"), "Step 15 - Work of the order is listed"

    ' Steps 16-19: complete the work as the work user
    status = D365_CompleteWork(workUser, "Steps 16-19 - Complete work")
    If Not status Then Exit Sub

    ' Step 20: Back
    VerifyTrue D365FO_BackToWorkPage("Step 20 - Back"), "Step 20 - Back to the Work page"

    ' Steps 21-22: Shipment > Complete shipment, then select the sales order again
    status = D365FO_ClickLabel("Shipment", "Step 21 - Shipment")
    If Not status Then Exit Sub
    ' Root cause of the failure: on this environment the Shipment tab of the Work page only holds "Confirm shipment" (no "Complete shipment"),
    ' so use the button the tab offers: "Complete shipment" when present, else "Confirm shipment".
    If D365FO_LabelExists("Complete shipment", 3) Then
        status = D365FO_ClickLabel("Complete shipment", "Step 21 - Complete shipment")
    Else
        status = D365FO_ClickLabel("Confirm shipment", "Step 21 - Complete shipment (button is named Confirm shipment)")
    End If
    If Not status Then Exit Sub
    ' TODO(C9652592): verify the shipment completed message
    VerifyTrue (D365FO_WaitTextContains("ompleted", 30) Or D365FO_WaitTextContains("onfirmed", 5)), "Step 21 - Shipment completed / confirmed"
    status = D365_OpenMenuItem("SalesTableListPage", "Step 22 - All sales orders")
    If Not status Then Exit Sub
    status = D365_SelectSalesOrder(salesId, "Step 22 - Select sales order " & salesId)
    If Not status Then Exit Sub

    ' Steps 23-25: Pick and pack > Post packing slip, OK, OK -> Delivered
    status = D365FO_ClickTab("Pick and pack", "Step 23 - Pick and pack")
    If Not status Then Exit Sub
    status = D365FO_ClickLabel("Post packing slip", "Step 23 - Post packing slip")
    If Not status Then Exit Sub
    VerifyObjectExist GetPage().WebElement("xpath:=(//*[contains(normalize-space(text()),'acking slip')])[1]"), 15, "Step 23 - Packing slip posting dialog is open"
    status = D365FO_ClickLabel("OK", "Step 24 - Click OK")
    If Not status Then Exit Sub
    D365FO_ClickLabel "OK", "Step 25 - Click OK"
    ' Root cause of the failure: the order did become Delivered (screenshot) but the status cell is not matched by text() alone.
    VerifyTrue D365_WaitStatusText("Delivered", 90), "Step 25 - Sales order " & salesId & " status is Delivered (was Open order)"
End Sub

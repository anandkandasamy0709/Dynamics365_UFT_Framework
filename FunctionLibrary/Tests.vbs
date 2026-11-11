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

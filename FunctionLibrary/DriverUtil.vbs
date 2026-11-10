'==============================================================================
' DriverUtil.vbs
'   Framework entry points, called from Driver\Action1:
'     InitFramework       -> loads .env, shared OR
'     MainDriver          -> runs every row flagged YES in the execution CSV
'     TearDownFramework   -> unloads OR, closes browser
'
'   Test case contract:
'     A row's Function_Name is a Sub in FunctionLibrary\Tests.vbs (or any
'     loaded library) with ONE parameter: the TestData dictionary.
'
'       Sub TC_MyScenario(dictData)
'           RunStep "Step description", SomeKeyword(dictData("key"))
'       End Sub
'
'   A test case is Passed only if at least one step passed, no step failed
'   and no runtime error occurred.
'==============================================================================

Sub InitFramework()
    LoadConfig
    Reporter.ReportEvent micDone, "Framework", _
        "Root: " & GetFrameworkRoot() & " | App: " & GetConfigOrDefault("APP_NAME", "(not set)") & _
        " | Run: " & GetRunId()
    LoadSharedOR
End Sub

Sub TearDownFramework()
    UnloadSharedOR
End Sub


'------------------------------------------------------------------------------
' MainDriver
'------------------------------------------------------------------------------
Sub MainDriver()
    Dim execData, results(), i, rowData
    Dim tcName, execStatus, funcName, testData, tcResult
    Dim cntTotal, cntRun, cntPass, cntFail, cntSkip, summary

    cntRun = 0 : cntPass = 0 : cntFail = 0 : cntSkip = 0

    execData = LoadExecutionData()
    If IsEmpty(execData) Then
        Reporter.ReportEvent micFail, "Driver", "No test cases loaded from: " & GetExecutionCsvPath()
        Exit Sub
    End If

    cntTotal = UBound(execData) + 1
    ReDim results(UBound(execData))
    Reporter.ReportEvent micDone, "Driver", cntTotal & " test case(s) found in " & GetExecutionCsvPath()

    For i = 0 To UBound(execData)
        Set rowData = execData(i)
        tcName     = rowData("TCName")
        execStatus = rowData("Execute_Status")
        funcName   = rowData("Function_Name")
        testData   = rowData("TestData")

        If UCase(Trim(execStatus)) = "YES" Then
            Set tcResult = RunTestCase(tcName, funcName, testData, i + 2)
            cntRun = cntRun + 1
            If tcResult("TC_Status") = "Passed" Then cntPass = cntPass + 1 Else cntFail = cntFail + 1
        Else
            Set tcResult = NewTCResult(tcName, funcName, execStatus)
            tcResult("TC_Status") = "Skipped"
            tcResult("Error") = "Execute_Status = '" & execStatus & "'"
            cntSkip = cntSkip + 1
            Reporter.ReportEvent micDone, tcName & " : Skipped", "Execute_Status is '" & execStatus & "' (only YES runs)"
        End If

        Set results(i) = tcResult
    Next

    summary = "Total: " & cntTotal & " | Executed: " & cntRun & " | Passed: " & cntPass & _
              " | Failed: " & cntFail & " | Skipped: " & cntSkip
    If cntRun = 0 Then
        Reporter.ReportEvent micWarning, "Execution Summary", "No test case is flagged YES. " & summary
    Else
        Reporter.ReportEvent micDone, "Execution Summary", summary
    End If

    WriteExecutionStatus results
End Sub


'------------------------------------------------------------------------------
' RunTestCase - runs one test function and decides Passed / Failed
'------------------------------------------------------------------------------
Function RunTestCase(tcName, funcName, testData, rowNum)
    Dim tcResult, dictData, errText, tcStatus, startTime, runStatusBefore, fn, perTestBrowser

    Set tcResult = NewTCResult(tcName, funcName, "Yes")
    startTime = Now
    tcResult("Start_Time") = CStr(startTime)
    errText = ""
    perTestBrowser = GetConfigBool("OPEN_BROWSER_PER_TEST", True)

    ResetTCStatus
    runStatusBefore = Reporter.RunStatus
    Reporter.ReportEvent micDone, tcName, "===== TEST STARTED ===== Function: " & funcName & " | Row: " & rowNum

    On Error Resume Next

    ' 1. Resolve the test function
    If funcName = "" Then
        errText = "Function_Name is blank in the execution CSV"
    Else
        Err.Clear
        Set fn = GetRef(funcName)
        If Err.Number <> 0 Then errText = "Function '" & funcName & "' not found in any loaded library"
    End If

    ' 2. Test data
    If errText = "" Then
        Err.Clear
        Set dictData = ConvertToDictionary(testData)
        If Err.Number <> 0 Then errText = "Invalid TestData '" & testData & "' - " & Err.Description
    End If

    ' 3. Browser
    If errText = "" And perTestBrowser Then
        Err.Clear
        OpenBrowser GetData(dictData, "url", "")
        If Err.Number <> 0 Then errText = "OpenBrowser failed - " & Err.Description
    End If

    ' 4. Execute - a runtime error stops the test function and fails the TC
    If errText = "" Then
        Err.Clear
        Call fn(dictData)
        If Err.Number <> 0 Then errText = "Runtime error in " & funcName & " - Err " & Err.Number & ": " & Err.Description
    End If

    ' 5. Cleanup (always)
    If perTestBrowser Then
        Err.Clear
        CloseBrowser
    End If
    Err.Clear
    On Error GoTo 0

    ' 6. Decide status
    If errText <> "" Then
        Reporter.ReportEvent micFail, tcName & " - Execution Error", errText
        tcStatus = "Failed"
    ElseIf gTC_StepsFailed > 0 Then
        tcStatus = "Failed"
        errText = gTC_StepsFailed & " of " & (gTC_StepsPassed + gTC_StepsFailed) & " step(s) failed: " & gTC_FailedSteps
    ElseIf runStatusBefore <> micFail And Reporter.RunStatus = micFail Then
        tcStatus = "Failed"
        errText = "A failure was reported directly to the run results during this test case"
    ElseIf gTC_StepsPassed = 0 Then
        tcStatus = "Failed"
        errText = "No verification steps were reported by " & funcName
    Else
        tcStatus = "Passed"
    End If

    ' 7. Report
    If tcStatus = "Passed" Then
        Reporter.ReportEvent micPass, tcName & " : Passed", _
            "PASS - " & tcName & " | Function: " & funcName & " | Steps passed: " & gTC_StepsPassed
    Else
        Reporter.ReportEvent micFail, tcName & " : Failed", "FAIL - " & tcName & " | Error: " & errText
    End If

    tcResult("TC_Status")    = tcStatus
    tcResult("Steps_Passed") = gTC_StepsPassed
    tcResult("Steps_Failed") = gTC_StepsFailed
    tcResult("End_Time")     = CStr(Now)
    tcResult("Duration_Sec") = DateDiff("s", startTime, Now)
    tcResult("Error")        = errText

    Reporter.ReportEvent micDone, tcName, "===== TEST ENDED ===== Status: " & tcStatus & _
        " | Duration: " & tcResult("Duration_Sec") & " sec"

    Set RunTestCase = tcResult
End Function


Function NewTCResult(tcName, funcName, execStatus)
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "TCName",         tcName
    d.Add "Function_Name",  funcName
    d.Add "Execute_Status", execStatus
    d.Add "TC_Status",      ""
    d.Add "Steps_Passed",   0
    d.Add "Steps_Failed",   0
    d.Add "Start_Time",     ""
    d.Add "End_Time",       ""
    d.Add "Duration_Sec",   ""
    d.Add "Error",          ""
    Set NewTCResult = d
End Function

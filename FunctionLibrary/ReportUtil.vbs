'==============================================================================
' ReportUtil.vbs
'   Central logging. Every pass/fail step goes through these helpers so that
'   MainDriver can decide a Passed / Failed status for each test case.
'
'   Screenshots: <root>\Screenshots\<RunId>\...
'     SCREENSHOT_ON_FAIL (default true)  - capture on failed steps
'     SCREENSHOT_ON_PASS (default false) - capture on passed steps
'==============================================================================

Dim gTC_StepsPassed, gTC_StepsFailed, gTC_FailedSteps, gRunId
gTC_StepsPassed = 0
gTC_StepsFailed = 0
gTC_FailedSteps = ""
gRunId = ""

Sub ResetTCStatus()
    gTC_StepsPassed = 0
    gTC_StepsFailed = 0
    gTC_FailedSteps = ""
End Sub

Sub TrackStep(stepName, isPassed)
    If isPassed Then
        gTC_StepsPassed = gTC_StepsPassed + 1
    Else
        gTC_StepsFailed = gTC_StepsFailed + 1
        If gTC_FailedSteps <> "" Then gTC_FailedSteps = gTC_FailedSteps & "; "
        gTC_FailedSteps = gTC_FailedSteps & stepName
    End If
End Sub


'------------------------------------------------------------------------------
' Public logging API
'------------------------------------------------------------------------------
Sub LogPass(stepName, msg)
    Capture micPass, stepName, msg
End Sub

Sub LogFail(stepName, msg)
    Capture micFail, stepName, msg
End Sub

Sub LogInfo(stepName, msg)
    Reporter.ReportEvent micDone, stepName, msg
End Sub

Sub LogWarning(stepName, msg)
    Reporter.ReportEvent micWarning, stepName, msg
End Sub

' errObj may be the VBScript Err object or a plain error number (0 = success)
Sub LogResult(stepName, errObj)
    Dim errNum, errDesc
    If IsObject(errObj) Then
        errNum = errObj.Number : errDesc = errObj.Description
        errObj.Clear
    Else
        errNum = CLng(errObj) : errDesc = "Error " & errNum
    End If

    If errNum = 0 Then
        LogPass stepName, "Step successful"
    Else
        LogFail stepName, errDesc
    End If
End Sub

' status: True = pass, False = fail
Sub LogBoolResult(stepName, status)
    If status = True Then
        LogPass stepName, "PASSED"
    Else
        LogFail stepName, "FAILED"
    End If
End Sub

Sub RunStep(stepName, status)
    LogBoolResult stepName, status
End Sub


'------------------------------------------------------------------------------
' Capture - reports the event, attaches a screenshot when configured,
' and counts the step for the current test case
'------------------------------------------------------------------------------
Sub Capture(status, stepName, msg)
    Dim img, wantShot, capturedOk, shotDir, oFso

    wantShot = False
    If status = micFail Then wantShot = GetConfigBool("SCREENSHOT_ON_FAIL", True)
    If status = micPass Then wantShot = GetConfigBool("SCREENSHOT_ON_PASS", False)

    capturedOk = False
    If wantShot Then
        ' A screenshot problem must never hide the step result
        On Error Resume Next
        Set oFso = CreateObject("Scripting.FileSystemObject")
        shotDir = GetFrameworkRoot() & "\Screenshots\" & GetRunId()
        If Not oFso.FolderExists(GetFrameworkRoot() & "\Screenshots") Then oFso.CreateFolder GetFrameworkRoot() & "\Screenshots"
        If Not oFso.FolderExists(shotDir) Then oFso.CreateFolder shotDir
        img = shotDir & "\" & SafeFileName(stepName) & "_" & TimeStamp() & ".png"
        Desktop.CaptureBitmap img, True
        capturedOk = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End If

    If capturedOk Then
        Reporter.ReportEvent status, stepName, msg, img
    Else
        Reporter.ReportEvent status, stepName, msg
    End If

    If status = micPass Then
        TrackStep stepName, True
    ElseIf status = micFail Then
        TrackStep stepName, False
    End If
End Sub


'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Function GetRunId()
    If gRunId = "" Then gRunId = "Run_" & TimeStamp()
    GetRunId = gRunId
End Function

' yyyyMMdd_HHmmss - locale independent
Function TimeStamp()
    Dim n
    n = Now
    TimeStamp = Year(n) & Pad2(Month(n)) & Pad2(Day(n)) & "_" & _
                Pad2(Hour(n)) & Pad2(Minute(n)) & Pad2(Second(n))
End Function

Function Pad2(v)
    Pad2 = Right("0" & v, 2)
End Function

Function SafeFileName(name)
    Dim s, ch
    s = CStr(name)
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbCr, vbLf, vbTab)
        s = Replace(s, ch, "_")
    Next
    If Len(s) > 80 Then s = Left(s, 80)
    SafeFileName = s
End Function

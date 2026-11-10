'==============================================================================
' RunTest.vbs - runs the Driver test from the command line / CI
'
'   cscript //nologo Scripts\RunTest.vbs [testPath] [resultsDir]
'
'   testPath    default: <root>\Driver
'   resultsDir  default: RESULTS_DIR from .env, else <root>\Results
'
'   Exit codes: 0 Passed/Warning, 1 Failed, 2 UFT not available,
'               3 test could not be opened, 4 run error, 5 UFT not ready
'==============================================================================
Option Explicit

Dim fso, root, logs, cfg, testPath, resultsBase, resultFolder
Dim qtApp, qtTest, startTime, status, exitCode

Set fso  = CreateObject("Scripting.FileSystemObject")
root     = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
Set logs = fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "RunTest.log"), 8, True)
Set cfg  = ReadEnv(root & "\.env")

Sub Log(msg)
    logs.WriteLine Now & " - " & msg
    WScript.Echo msg
End Sub

Sub Finish(code)
    Log "Exit code: " & code
    logs.Close
    WScript.Quit code
End Sub

' Minimal .env reader (same rules as FunctionLibrary\EnvUtil.vbs)
Function ReadEnv(path)
    Dim d, ts, line, p, k, v, q
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    If fso.FileExists(path) Then
        Set ts = fso.OpenTextFile(path, 1, False, -2)
        Do Until ts.AtEndOfStream
            line = Trim(Replace(Replace(ts.ReadLine, vbCr, ""), Chr(239) & Chr(187) & Chr(191), ""))
            p = InStr(line, "=")
            If line <> "" And Left(line, 1) <> "#" And p > 1 Then
                k = Trim(Left(line, p - 1)) : v = Trim(Mid(line, p + 1))
                q = Left(v, 1)
                If (q = """" Or q = "'") And Len(v) >= 2 And InStr(2, v, q) > 0 Then
                    v = Mid(v, 2, InStr(2, v, q) - 2)
                ElseIf InStr(v, " #") > 0 Then
                    v = Trim(Left(v, InStr(v, " #") - 1))
                End If
                d(k) = v
            End If
        Loop
        ts.Close
    End If
    Set ReadEnv = d
End Function

Function Cfg_(key, def)
    Cfg_ = def
    If cfg.Exists(key) Then
        If Trim(cfg(key)) <> "" Then Cfg_ = cfg(key)
    End If
End Function

Function Stamp()
    Dim n : n = Now
    Stamp = Year(n) & Right("0" & Month(n), 2) & Right("0" & Day(n), 2) & "_" & _
            Right("0" & Hour(n), 2) & Right("0" & Minute(n), 2) & Right("0" & Second(n), 2)
End Function

'---------------------------------------------------------------- arguments
testPath = root & "\Driver"
If WScript.Arguments.Count >= 1 Then testPath = WScript.Arguments(0)

resultsBase = Cfg_("RESULTS_DIR", root & "\Results")
If WScript.Arguments.Count >= 2 Then resultsBase = WScript.Arguments(1)
If Not (Mid(resultsBase, 2, 1) = ":" Or Left(resultsBase, 2) = "\\") Then resultsBase = root & "\" & resultsBase
If Not fso.FolderExists(resultsBase) Then fso.CreateFolder resultsBase
resultFolder = resultsBase & "\Run_" & Stamp()

Log "==== RunTest.vbs ===="
Log "Framework root : " & root
Log "Test path      : " & testPath
Log "Results folder : " & resultFolder
If Not fso.FileExists(root & "\.env") Then Log "WARNING: .env not found - copy .env.example to .env"

'---------------------------------------------------------------- launch UFT
On Error Resume Next
Set qtApp = CreateObject("QuickTest.Application")
If Err.Number <> 0 Then
    Log "ERROR: Cannot create QuickTest.Application - " & Err.Number & " : " & Err.Description
    Finish 2
End If
On Error GoTo 0

If Not qtApp.Launched Then qtApp.Launch
qtApp.Visible = (LCase(Cfg_("UFT_VISIBLE", "true")) = "true")

startTime = Timer
Do While qtApp.GetStatus <> "Ready"
    WScript.Sleep 1000
    If Timer - startTime > 120 Then
        Log "ERROR: UFT did not become Ready within 120 seconds"
        qtApp.Quit
        Finish 5
    End If
Loop

'---------------------------------------------------------------- open + run
On Error Resume Next
qtApp.Open testPath, False, False
If Err.Number <> 0 Then
    Log "ERROR opening test: " & Err.Number & " : " & Err.Description
    qtApp.Quit
    Finish 3
End If

Set qtTest = qtApp.Test
Log "Starting test run..."
Dim resOpts
Set resOpts = CreateObject("QuickTest.RunResultsOptions")
resOpts.ResultsLocation = resultFolder
qtTest.Run resOpts, True
If Err.Number <> 0 Then
    Log "ERROR during test run: " & Err.Number & " : " & Err.Description
    qtApp.Quit
    Finish 4
End If
status = qtTest.LastRunResults.Status
On Error GoTo 0

Log "Run status: " & status
qtTest.Close
qtApp.Quit

'---------------------------------------------------------------- HTML summary
CreateObject("WScript.Shell").Run "cscript //nologo """ & root & "\Scripts\GenerateHtmlReport.vbs""", 0, True
Log "HTML report: " & root & "\Reports\ExecutionReport.html"

If LCase(status) = "failed" Then exitCode = 1 Else exitCode = 0
Finish exitCode

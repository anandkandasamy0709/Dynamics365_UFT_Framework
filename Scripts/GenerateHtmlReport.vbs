'==============================================================================
' GenerateHtmlReport.vbs
'   Builds Reports\ExecutionReport.html from Reports\ExecutionStatus.csv
'   (written by MainDriver at the end of every run).
'
'   cscript //nologo Scripts\GenerateHtmlReport.vbs [statusCsv] [outHtml]
'==============================================================================
Option Explicit

Dim fso, root, csvPath, htmlPath, appName, ts, line, isHeader, f, rows, html
Dim total, passed, failed, skipped, cls

Set fso  = CreateObject("Scripting.FileSystemObject")
root     = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
csvPath  = root & "\Reports\ExecutionStatus.csv"
htmlPath = root & "\Reports\ExecutionReport.html"
If WScript.Arguments.Count >= 1 Then csvPath = WScript.Arguments(0)
If WScript.Arguments.Count >= 2 Then htmlPath = WScript.Arguments(1)

If Not fso.FileExists(csvPath) Then
    WScript.Echo "Status file not found: " & csvPath
    WScript.Quit 1
End If

appName = ReadEnvValue(root & "\.env", "APP_NAME")
If appName = "" Then appName = "Application"

total = 0 : passed = 0 : failed = 0 : skipped = 0 : rows = ""
Set ts = fso.OpenTextFile(csvPath, 1, False, -2)
isHeader = True
Do Until ts.AtEndOfStream
    line = ts.ReadLine
    If Trim(line) <> "" Then
        If isHeader Then
            isHeader = False
        Else
            ' TCName,Function_Name,Execute_Status,TC_Status,Steps_Passed,Steps_Failed,Start_Time,End_Time,Duration_Sec,Error
            f = ParseCsvLine(line)
            If UBound(f) >= 9 Then
                total = total + 1
                Select Case f(3)
                    Case "Passed":  passed = passed + 1 : cls = "pass"
                    Case "Failed":  failed = failed + 1 : cls = "fail"
                    Case Else:      skipped = skipped + 1 : cls = "skip"
                End Select
                rows = rows & "<tr class='" & cls & "'><td>" & H(f(0)) & "</td><td>" & H(f(1)) & _
                       "</td><td><b>" & H(f(3)) & "</b></td><td>" & H(f(4)) & "/" & H(f(5)) & _
                       "</td><td>" & H(f(6)) & "</td><td>" & H(f(8)) & "</td><td>" & H(f(9)) & "</td></tr>" & vbCrLf
            End If
        End If
    End If
Loop
ts.Close

html = "<!DOCTYPE html><html><head><meta charset='utf-8'><title>" & H(appName) & " - Execution Report</title>" & _
       "<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#222}" & _
       "table{border-collapse:collapse;width:100%;margin-top:12px}th,td{border:1px solid #ccc;padding:6px 8px;text-align:left;font-size:13px}" & _
       "th{background:#f0f0f0}tr.pass td:nth-child(3){color:#1a7f37}tr.fail td:nth-child(3){color:#cf222e}tr.skip td{color:#888}" & _
       ".cards span{display:inline-block;margin-right:12px;padding:8px 14px;border-radius:6px;background:#f6f8fa;border:1px solid #ddd}</style></head><body>" & _
       "<h2>" & H(appName) & " - Execution Report</h2><p>Generated: " & Now & "</p>" & _
       "<div class='cards'><span>Total: <b>" & total & "</b></span><span>Passed: <b>" & passed & "</b></span>" & _
       "<span>Failed: <b>" & failed & "</b></span><span>Skipped: <b>" & skipped & "</b></span></div>" & _
       "<table><tr><th>TC</th><th>Function</th><th>Status</th><th>Steps P/F</th><th>Start</th><th>Duration (s)</th><th>Error</th></tr>" & vbCrLf & _
       rows & "</table></body></html>"

Set ts = fso.CreateTextFile(htmlPath, True, False)
ts.Write html
ts.Close
WScript.Echo "Report written: " & htmlPath


Function H(s)
    H = Replace(Replace(Replace(CStr(s), "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
End Function

Function ReadEnvValue(path, key)
    Dim t, l, p
    ReadEnvValue = ""
    If Not fso.FileExists(path) Then Exit Function
    Set t = fso.OpenTextFile(path, 1, False, -2)
    Do Until t.AtEndOfStream
        l = Trim(t.ReadLine)
        p = InStr(l, "=")
        If p > 1 And Left(l, 1) <> "#" Then
            If UCase(Trim(Left(l, p - 1))) = UCase(key) Then ReadEnvValue = Replace(Trim(Mid(l, p + 1)), """", "")
        End If
    Loop
    t.Close
End Function

Function ParseCsvLine(l)
    Dim out(), n, i, ch, cur, inQ
    n = 0 : cur = "" : inQ = False
    ReDim out(0)
    For i = 1 To Len(l)
        ch = Mid(l, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid(l, i + 1, 1) = """" Then
                    cur = cur & """" : i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        ElseIf ch = """" Then
            inQ = True
        ElseIf ch = "," Then
            out(n) = cur : cur = "" : n = n + 1 : ReDim Preserve out(n)
        Else
            cur = cur & ch
        End If
    Next
    out(n) = cur
    ParseCsvLine = out
End Function

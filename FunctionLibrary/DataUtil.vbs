'==============================================================================
' DataUtil.vbs
'   CSV-based test data (no Excel / ADO dependency).
'
'   Execution sheet (path from EXECUTION_FILE in .env, default
'   TestData\Execution.csv) columns:
'     TCName, Execute_Status, Function_Name, TestData
'
'   TestData format :  key:value;key2:value2
'   Values may reference .env keys:  user:${APP_USERNAME}
'==============================================================================

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim gExecData, gExecDataLoaded
gExecDataLoaded = False


Function GetExecutionCsvPath()
    GetExecutionCsvPath = ResolvePath(GetConfigOrDefault("EXECUTION_FILE", "TestData\Execution.csv"))
End Function


'------------------------------------------------------------------------------
' LoadExecutionData
'   Returns an array of Dictionaries (TCName, Execute_Status, Function_Name,
'   TestData) or Empty when nothing could be loaded. Cached for the run.
'------------------------------------------------------------------------------
Function LoadExecutionData()
    If gExecDataLoaded Then
        LoadExecutionData = gExecData
        Exit Function
    End If

    Dim path, ts, line, isHeader, parts, list(), count, dict

    gExecData = Empty
    path = GetExecutionCsvPath()
    Reporter.ReportEvent micDone, "Test Data", "Loading CSV: " & path

    If Not fso.FileExists(path) Then
        Reporter.ReportEvent micFail, "CSV Error", "File not found: " & path
        LoadExecutionData = Empty
        Exit Function
    End If

    Set ts = fso.OpenTextFile(path, 1, False, -2)
    isHeader = True
    count = -1

    Do Until ts.AtEndOfStream
        line = Trim(Replace(ts.ReadLine, Chr(239) & Chr(187) & Chr(191), ""))

        If line <> "" And Left(line, 1) <> "#" Then
            If isHeader Then
                isHeader = False
            Else
                parts = ParseCsvLine(line)
                If UBound(parts) >= 3 Then
                    count = count + 1
                    ReDim Preserve list(count)
                    Set dict = CreateObject("Scripting.Dictionary")
                    dict.Add "TCName",         Trim(parts(0))
                    dict.Add "Execute_Status", Trim(parts(1))
                    dict.Add "Function_Name",  Trim(parts(2))
                    dict.Add "TestData",       Trim(parts(3))
                    Set list(count) = dict
                Else
                    Reporter.ReportEvent micWarning, "CSV Warning", "Ignored row (less than 4 columns): " & line
                End If
            End If
        End If
    Loop
    ts.Close

    If count >= 0 Then gExecData = list
    gExecDataLoaded = True
    LoadExecutionData = gExecData
End Function


'------------------------------------------------------------------------------
' ParseCsvLine - RFC-4180 style: handles "quoted, values" and "" escapes.
' Any columns beyond the 4th are merged back into the TestData column.
'------------------------------------------------------------------------------
Function ParseCsvLine(line)
    Dim fields(), n, i, ch, cur, inQ
    n = 0 : cur = "" : inQ = False
    ReDim fields(0)

    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid(line, i + 1, 1) = """" Then
                    cur = cur & """" : i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inQ = True
            ElseIf ch = "," Then
                fields(n) = cur : cur = ""
                n = n + 1 : ReDim Preserve fields(n)
            Else
                cur = cur & ch
            End If
        End If
    Next
    fields(n) = cur

    If n > 3 Then
        For i = 4 To n
            fields(3) = fields(3) & "," & fields(i)
        Next
        ReDim Preserve fields(3)
    End If

    ParseCsvLine = fields
End Function


'------------------------------------------------------------------------------
' ConvertToDictionary - "key:value;key2:value2" -> Dictionary
'   Only the first ':' splits key and value, so URLs are safe as values.
'   ${KEY} placeholders are resolved from .env.
'------------------------------------------------------------------------------
Function ConvertToDictionary(data)
    Dim dict, pairs, i, kv
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    If Trim(data) <> "" Then
        pairs = Split(data, ";")
        For i = 0 To UBound(pairs)
            kv = Split(pairs(i), ":", 2)
            If UBound(kv) >= 1 Then
                If Not dict.Exists(Trim(kv(0))) Then
                    dict.Add Trim(kv(0)), ExpandConfigRefs(Trim(kv(1)))
                End If
            End If
        Next
    End If

    Set ConvertToDictionary = dict
End Function

' Safe read from the test data dictionary
Function GetData(dictData, key, defaultValue)
    GetData = defaultValue
    If IsObject(dictData) Then
        If dictData.Exists(key) Then GetData = dictData(key)
    End If
End Function


'------------------------------------------------------------------------------
' WriteExecutionStatus
'   Reports\ExecutionStatus.csv            (latest run, used by the HTML report)
'   Reports\ExecutionStatus_<RunId>.csv    (history)
'------------------------------------------------------------------------------
Sub WriteExecutionStatus(resultList)
    Dim reportsDir, latest, history

    reportsDir = GetFrameworkRoot() & "\Reports"
    latest  = reportsDir & "\ExecutionStatus.csv"
    history = reportsDir & "\ExecutionStatus_" & GetRunId() & ".csv"

    On Error Resume Next
    If Not fso.FolderExists(reportsDir) Then fso.CreateFolder reportsDir
    On Error GoTo 0

    WriteStatusFile latest, resultList
    WriteStatusFile history, resultList
    Reporter.ReportEvent micDone, "Execution Status", "TC status written to: " & latest
End Sub

Sub WriteStatusFile(outPath, resultList)
    Dim ts, i, r

    On Error Resume Next
    Set ts = fso.CreateTextFile(outPath, True, False)
    If Err.Number <> 0 Then
        Reporter.ReportEvent micWarning, "Execution Status", _
            "Could not write " & outPath & " - " & Err.Description & " (is it open in Excel?)"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    ts.WriteLine "TCName,Function_Name,Execute_Status,TC_Status,Steps_Passed,Steps_Failed,Start_Time,End_Time,Duration_Sec,Error"
    For i = 0 To UBound(resultList)
        Set r = resultList(i)
        ts.WriteLine CsvQuote(r("TCName")) & "," & CsvQuote(r("Function_Name")) & "," & _
                     CsvQuote(r("Execute_Status")) & "," & CsvQuote(r("TC_Status")) & "," & _
                     r("Steps_Passed") & "," & r("Steps_Failed") & "," & _
                     CsvQuote(r("Start_Time")) & "," & CsvQuote(r("End_Time")) & "," & _
                     r("Duration_Sec") & "," & CsvQuote(r("Error"))
    Next
    ts.Close
End Sub

Function CsvQuote(value)
    Dim s
    s = CStr(value)
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then
        s = """" & Replace(s, """", """""") & """"
    End If
    CsvQuote = s
End Function

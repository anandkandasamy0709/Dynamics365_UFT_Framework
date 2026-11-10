'==============================================================================
' EnvUtil.vbs
'   Loads configuration (URLs, credentials, browser, timeouts...) from the
'   .env file in the framework root and exposes it to the rest of the
'   framework through GetConfig().
'
'   Load order (later wins):
'     1. <root>\.env
'     2. <root>\.env.<TEST_ENV>   (only if TEST_ENV is set, e.g. .env.qa)
'     3. Windows environment variables with the same name (CI/CD override)
'
'   .env syntax:
'     KEY=VALUE              plain value
'     KEY="value # text"     quotes keep '#' and surrounding spaces
'     # comment              full-line comment
'     KEY=value  # comment   inline comment (unquoted values only)
'     KEY=${OTHER_KEY}/path  reference to another key
'==============================================================================

Dim gConfig
Set gConfig = CreateObject("Scripting.Dictionary")
gConfig.CompareMode = 1          ' vbTextCompare - keys are case-insensitive

Dim gConfigLoaded
gConfigLoaded = False


'------------------------------------------------------------------------------
' GetFrameworkRoot
'   The Driver test lives in <root>\Driver, so the root is its parent folder.
'   No hard-coded paths: works for every clone location and every CI node.
'------------------------------------------------------------------------------
Function GetFrameworkRoot()
    Dim root, oFso
    Set oFso = CreateObject("Scripting.FileSystemObject")

    root = ""
    On Error Resume Next
    root = Environment.Value("FrameworkRoot")
    If Err.Number <> 0 Then root = ""
    Err.Clear
    On Error GoTo 0

    If Trim(root) = "" Then root = oFso.GetParentFolderName(Environment.Value("TestDir"))
    If Right(root, 1) = "\" Then root = Left(root, Len(root) - 1)

    GetFrameworkRoot = root
End Function


'------------------------------------------------------------------------------
' ResolvePath - relative paths are resolved against the framework root
'------------------------------------------------------------------------------
Function ResolvePath(path)
    Dim p
    p = Trim(path)
    If p = "" Then
        ResolvePath = ""
    ElseIf Mid(p, 2, 1) = ":" Or Left(p, 2) = "\\" Then
        ResolvePath = p
    Else
        If Left(p, 2) = ".\" Then p = Mid(p, 3)
        ResolvePath = GetFrameworkRoot() & "\" & p
    End If
End Function


'------------------------------------------------------------------------------
' LoadConfig - call once at the start of the run (done by InitFramework)
'------------------------------------------------------------------------------
Sub LoadConfig()
    Dim root, envName, k, keys, osVal, oShell

    gConfig.RemoveAll
    root = GetFrameworkRoot()

    ' 1. Base .env
    If Not LoadEnvFile(root & "\.env") Then
        Reporter.ReportEvent micWarning, "Config", _
            ".env not found at " & root & "\.env - copy .env.example to .env and fill it in. " & _
            "Falling back to Windows environment variables only."
    End If

    ' 2. Optional environment overlay (.env.qa, .env.uat ...)
    Set oShell = CreateObject("WScript.Shell")
    envName = ReadOsVariable(oShell, "TEST_ENV")
    If envName = "" And gConfig.Exists("TEST_ENV") Then envName = gConfig("TEST_ENV")
    If Trim(envName) <> "" Then
        If LoadEnvFile(root & "\.env." & Trim(envName)) Then
            Reporter.ReportEvent micDone, "Config", "Overlay loaded: .env." & envName
        End If
    End If

    ' 3. Windows environment variables override file values (Jenkins credentials etc.)
    keys = gConfig.Keys
    For Each k In keys
        osVal = ReadOsVariable(oShell, k)
        If osVal <> "" Then gConfig(k) = osVal
    Next

    ' 4. Expand ${KEY} references
    keys = gConfig.Keys
    For Each k In keys
        gConfig(k) = ExpandConfigRefs(gConfig(k))
    Next

    gConfigLoaded = True
    Reporter.ReportEvent micDone, "Config", "Configuration loaded:" & vbCrLf & DescribeConfig()
End Sub


'------------------------------------------------------------------------------
' LoadEnvFile - parses one .env file into gConfig. Returns False if missing.
'------------------------------------------------------------------------------
Function LoadEnvFile(path)
    Dim oFso, ts, line, pos, key, value, q

    Set oFso = CreateObject("Scripting.FileSystemObject")
    LoadEnvFile = False
    If Not oFso.FileExists(path) Then Exit Function

    Set ts = oFso.OpenTextFile(path, 1, False, -2)    ' -2 = system default encoding
    Do Until ts.AtEndOfStream
        line = Replace(ts.ReadLine, vbCr, "")
        line = Replace(line, Chr(239) & Chr(187) & Chr(191), "")   ' strip UTF-8 BOM if present
        line = Trim(Replace(line, vbTab, " "))

        If line <> "" And Left(line, 1) <> "#" Then
            If LCase(Left(line, 7)) = "export " Then line = Trim(Mid(line, 8))
            pos = InStr(line, "=")
            If pos > 1 Then
                key = Trim(Left(line, pos - 1))
                value = Trim(Mid(line, pos + 1))

                q = Left(value, 1)
                If (q = """" Or q = "'") And Len(value) >= 2 And InStr(2, value, q) > 0 Then
                    value = Mid(value, 2, InStr(2, value, q) - 2)
                Else
                    pos = InStr(value, " #")
                    If pos > 0 Then value = Trim(Left(value, pos - 1))
                End If

                gConfig(key) = value
            End If
        End If
    Loop
    ts.Close

    LoadEnvFile = True
End Function


Function ReadOsVariable(oShell, name)
    Dim v
    v = oShell.ExpandEnvironmentStrings("%" & name & "%")
    If v = "%" & name & "%" Then v = ""
    ReadOsVariable = v
End Function


Function ExpandConfigRefs(value)
    Dim s, p1, p2, refKey, guard
    s = CStr(value)
    guard = 0
    p1 = InStr(s, "${")
    Do While p1 > 0 And guard < 20
        p2 = InStr(p1, s, "}")
        If p2 = 0 Then Exit Do
        refKey = Mid(s, p1 + 2, p2 - p1 - 2)
        s = Left(s, p1 - 1) & GetConfig(refKey) & Mid(s, p2 + 1)
        p1 = InStr(p1, s, "${")
        guard = guard + 1
    Loop
    ExpandConfigRefs = s
End Function


'------------------------------------------------------------------------------
' Accessors
'------------------------------------------------------------------------------
Function GetConfig(key)
    If gConfig.Exists(key) Then
        GetConfig = gConfig(key)
    Else
        GetConfig = ""
    End If
End Function

Function GetConfigOrDefault(key, defaultValue)
    Dim v
    v = GetConfig(key)
    If Trim(v) = "" Then v = defaultValue
    GetConfigOrDefault = v
End Function

Function GetConfigInt(key, defaultValue)
    Dim v
    v = Trim(GetConfig(key))
    If IsNumeric(v) And v <> "" Then
        GetConfigInt = CLng(v)
    Else
        GetConfigInt = defaultValue
    End If
End Function

Function GetConfigBool(key, defaultValue)
    Select Case LCase(Trim(GetConfig(key)))
        Case "true", "yes", "y", "1", "on":  GetConfigBool = True
        Case "false", "no", "n", "0", "off": GetConfigBool = False
        Case Else:                           GetConfigBool = defaultValue
    End Select
End Function

' Fails the step (and returns "") when a mandatory key is missing
Function RequireConfig(key)
    Dim v
    v = GetConfig(key)
    If Trim(v) = "" Then
        Reporter.ReportEvent micFail, "Config", "Required key '" & key & "' is missing or empty in .env"
    End If
    RequireConfig = v
End Function

Sub SetConfig(key, value)
    gConfig(key) = value
End Sub


'------------------------------------------------------------------------------
' Secrets are masked whenever configuration is written to the results
'------------------------------------------------------------------------------
Function IsSecretKey(key)
    Dim k
    k = UCase(key)
    IsSecretKey = (InStr(k, "PASSWORD") > 0 Or InStr(k, "SECRET") > 0 Or _
                   InStr(k, "TOKEN") > 0 Or InStr(k, "APIKEY") > 0 Or InStr(k, "API_KEY") > 0)
End Function

Function DescribeConfig()
    Dim k, s, v
    s = ""
    For Each k In gConfig.Keys
        v = gConfig(k)
        If IsSecretKey(k) And v <> "" Then v = "********"
        s = s & k & " = " & v & vbCrLf
    Next
    DescribeConfig = s
End Function

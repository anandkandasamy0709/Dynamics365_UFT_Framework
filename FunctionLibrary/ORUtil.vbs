'==============================================================================
' ORUtil.vbs
'   Loads shared object repositories (.tsr) at runtime.
'
'   OBJECT_REPOSITORY in .env:
'     - one path            ObjectRepository\Application.tsr
'     - several paths       ObjectRepository\Login.tsr;ObjectRepository\Home.tsr
'     - empty               every *.tsr found in <root>\ObjectRepository
'   Relative paths are resolved against the framework root.
'
'   A missing repository is a warning, not a failure, so a fresh project can
'   run with descriptive programming before any .tsr exists.
'==============================================================================

Sub LoadSharedOR()
    Dim setting, paths, p, orDir, f, oFso, loaded
    Set oFso = CreateObject("Scripting.FileSystemObject")
    loaded = 0

    setting = Trim(GetConfig("OBJECT_REPOSITORY"))

    If setting <> "" Then
        paths = Split(setting, ";")
        For Each p In paths
            If Trim(p) <> "" Then
                If AddRepository(ResolvePath(p)) Then loaded = loaded + 1
            End If
        Next
    Else
        orDir = GetFrameworkRoot() & "\ObjectRepository"
        If oFso.FolderExists(orDir) Then
            For Each f In oFso.GetFolder(orDir).Files
                If LCase(oFso.GetExtensionName(f.Name)) = "tsr" Then
                    If AddRepository(f.Path) Then loaded = loaded + 1
                End If
            Next
        End If
    End If

    If loaded = 0 Then
        Reporter.ReportEvent micWarning, "Object Repository", _
            "No shared object repository loaded. Tests can still use descriptive programming."
    End If
End Sub

Function AddRepository(path)
    Dim oFso
    Set oFso = CreateObject("Scripting.FileSystemObject")
    AddRepository = False

    If Not oFso.FileExists(path) Then
        Reporter.ReportEvent micWarning, "Object Repository", "Not found: " & path
        Exit Function
    End If

    If RepositoriesCollection.Find(path) <> -1 Then
        AddRepository = True
        Exit Function
    End If

    On Error Resume Next
    RepositoriesCollection.Add path
    If Err.Number = 0 Then
        Reporter.ReportEvent micDone, "Object Repository", "Loaded: " & path
        AddRepository = True
    Else
        Reporter.ReportEvent micFail, "Object Repository", "Could not load " & path & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Function

Sub UnloadSharedOR()
    On Error Resume Next
    RepositoriesCollection.RemoveAll
    Err.Clear
    On Error GoTo 0
End Sub

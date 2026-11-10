'==============================================================================
' Driver - Action1 (framework bootstrap)
'
'   No hard-coded paths: the framework root is the parent of this test folder,
'   so the same code runs from any clone location and on any CI node.
'   URLs, credentials and settings come from <root>\.env (see EnvUtil.vbs).
'==============================================================================
Dim frameworkRoot, libraries, lib, libPath, oFso

Set oFso = CreateObject("Scripting.FileSystemObject")
frameworkRoot = oFso.GetParentFolderName(Environment.Value("TestDir"))
Environment.Value("FrameworkRoot") = frameworkRoot

'------------------------------------------------------------------
' Load function libraries (order matters: EnvUtil and ReportUtil first)
' Add new libraries to this list.
'------------------------------------------------------------------
libraries = Array( _
    "EnvUtil.vbs", _
    "ReportUtil.vbs", _
    "DataUtil.vbs", _
    "ORUtil.vbs", _
    "BrowserUtil.vbs", _
    "WebActions.vbs", _
    "Assertions.vbs", _
    "DriverUtil.vbs", _
    "AppFunctions.vbs", _
    "Tests.vbs")

For Each lib In libraries
    libPath = frameworkRoot & "\FunctionLibrary\" & lib
    If oFso.FileExists(libPath) Then
        LoadFunctionLibrary libPath
    Else
        Reporter.ReportEvent micFail, "Framework", "Function library not found: " & libPath
        ExitTest
    End If
Next
Reporter.ReportEvent micDone, "Framework", "Libraries loaded from " & frameworkRoot & "\FunctionLibrary"

'------------------------------------------------------------------
' Run
'------------------------------------------------------------------
InitFramework
MainDriver
TearDownFramework

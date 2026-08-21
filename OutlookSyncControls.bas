Attribute VB_Name = "OutlookSyncControls"
Option Explicit

' Set this locally after placing outlook_calendar_sync.py on the target PC.
' Keep it blank in the published template so local paths are not shared.
Private Const SYNC_SCRIPT As String = ""

' Leave blank to use settings.py next to outlook_calendar_sync.py.
Private Const SETTINGS_FILE As String = ""

' Leave blank to use the Windows Python launcher for the default Python 3.
' You may also set this to py -3.12, py -3.14, or a full python.exe path.
Private Const PYTHON_COMMAND As String = ""
Private Const DEFAULT_PYTHON_COMMAND As String = "py -3"

Private Function Quote(ByVal value As String) As String
    Quote = """" & Replace(value, """", """""") & """"
End Function

Private Function ExpandEnv(ByVal value As String) As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    ExpandEnv = shell.ExpandEnvironmentStrings(value)
End Function

Private Function CleanConfiguredPath(ByVal value As String) As String
    Dim cleaned As String
    cleaned = Trim$(value)

    If Len(cleaned) >= 2 Then
        If Left$(cleaned, 1) = """" And Right$(cleaned, 1) = """" Then
            cleaned = Mid$(cleaned, 2, Len(cleaned) - 2)
            cleaned = Trim$(cleaned)
        End If
    End If

    CleanConfiguredPath = cleaned
End Function

Private Function NormalizeSpaces(ByVal value As String) As String
    Dim normalized As String
    normalized = Trim$(value)

    Do While InStr(1, normalized, "  ", vbBinaryCompare) > 0
        normalized = Replace(normalized, "  ", " ")
    Loop

    NormalizeSpaces = normalized
End Function

Private Function IsDigitsOnly(ByVal value As String) As Boolean
    Dim index As Long
    Dim ch As String

    If Len(value) = 0 Then
        IsDigitsOnly = False
        Exit Function
    End If

    For index = 1 To Len(value)
        ch = Mid$(value, index, 1)
        If ch < "0" Or ch > "9" Then
            IsDigitsOnly = False
            Exit Function
        End If
    Next

    IsDigitsOnly = True
End Function

Private Function IsAllowedPyLauncherArgument(ByVal value As String) As Boolean
    Dim versionPart As String
    Dim architecturePart As String
    Dim dashAt As Long

    If value = "-3" Then
        IsAllowedPyLauncherArgument = True
        Exit Function
    End If

    If Left$(value, 3) <> "-3." Then
        IsAllowedPyLauncherArgument = False
        Exit Function
    End If

    versionPart = Mid$(value, 4)
    dashAt = InStr(1, versionPart, "-", vbBinaryCompare)
    If dashAt > 0 Then
        architecturePart = Mid$(versionPart, dashAt + 1)
        versionPart = Left$(versionPart, dashAt - 1)
        If architecturePart <> "32" And architecturePart <> "64" Then
            IsAllowedPyLauncherArgument = False
            Exit Function
        End If
    End If

    IsAllowedPyLauncherArgument = IsDigitsOnly(versionPart)
End Function

Private Function IsPyLauncherCommand(ByVal value As String) As Boolean
    Dim parts() As String
    Dim normalized As String

    normalized = NormalizeSpaces(value)
    parts = Split(normalized, " ")

    If LCase$(parts(0)) <> "py" Then
        IsPyLauncherCommand = False
        Exit Function
    End If

    If UBound(parts) = 0 Then
        IsPyLauncherCommand = True
        Exit Function
    End If

    If UBound(parts) = 1 Then
        IsPyLauncherCommand = IsAllowedPyLauncherArgument(parts(1))
        Exit Function
    End If

    IsPyLauncherCommand = False
End Function

Private Function PythonCommandPrefix(ByVal value As String) As String
    Dim cleaned As String

    cleaned = CleanConfiguredPath(value)
    If Len(cleaned) = 0 Then
        cleaned = DEFAULT_PYTHON_COMMAND
    End If

    RejectUnsafePath cleaned, "PYTHON_COMMAND"
    cleaned = NormalizeSpaces(cleaned)

    If LCase$(cleaned) = "py" Or Left$(LCase$(cleaned), 3) = "py " Then
        If Not IsPyLauncherCommand(cleaned) Then
            Err.Raise vbObjectError + 517, "OutlookSyncControls", _
                "PYTHON_COMMAND can be py, py -3, or py -3.12. For anything else, use the full path to python.exe."
        End If
        PythonCommandPrefix = cleaned
        Exit Function
    End If

    PythonCommandPrefix = Quote(ExpandEnv(cleaned))
End Function

Private Sub RejectUnsafePath(ByVal value As String, ByVal label As String)
    If InStr(1, value, """", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 513, "OutlookSyncControls", _
            label & " contains an invalid quote. Paste the path without quotes."
    End If

    If InStr(1, value, vbCr, vbBinaryCompare) > 0 Or _
       InStr(1, value, vbLf, vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 514, "OutlookSyncControls", _
            label & " contains an invalid line break. Paste a single full path."
    End If
End Sub

Private Function RequiredConfiguredPath(ByVal value As String, ByVal label As String) As String
    Dim cleaned As String
    Dim targetName As String

    cleaned = CleanConfiguredPath(value)
    If UCase$(label) = "PYTHON_COMMAND" Then
        targetName = "python.exe"
    Else
        targetName = "outlook_calendar_sync.py"
    End If

    If Len(cleaned) = 0 Then
        If UCase$(label) = "SYNC_SCRIPT" Then
            Err.Raise vbObjectError + 515, "OutlookSyncControls", _
                "Set SYNC_SCRIPT in OutlookSyncControls.bas to the full path of outlook_calendar_sync.py."
        Else
            Err.Raise vbObjectError + 515, "OutlookSyncControls", _
                "Set " & label & " in OutlookSyncControls.bas to the full path of " & targetName & "."
        End If
    End If

    RejectUnsafePath cleaned, label
    RequiredConfiguredPath = ExpandEnv(cleaned)
End Function

Private Function OptionalConfiguredPath(ByVal value As String, ByVal label As String) As String
    Dim cleaned As String
    cleaned = CleanConfiguredPath(value)

    If Len(cleaned) = 0 Then
        OptionalConfiguredPath = ""
        Exit Function
    End If

    RejectUnsafePath cleaned, label
    OptionalConfiguredPath = ExpandEnv(cleaned)
End Function

Private Function SafeAction(ByVal action As String) As String
    Select Case action
        Case "--start", "--stop", "--status", "--restart", "--sync-selected"
            SafeAction = action
        Case Else
            Err.Raise vbObjectError + 516, "OutlookSyncControls", _
                "Unknown Outlook Calendar Sync action: " & action
    End Select
End Function

Private Function SyncCommand(ByVal action As String) As String
    Dim settingsPath As String

    SyncCommand = PythonCommandPrefix(PYTHON_COMMAND) & _
        " " & Quote(RequiredConfiguredPath(SYNC_SCRIPT, "SYNC_SCRIPT")) & _
        " " & SafeAction(action)

    settingsPath = OptionalConfiguredPath(SETTINGS_FILE, "SETTINGS_FILE")
    If Len(settingsPath) > 0 Then
        SyncCommand = SyncCommand & " --settings " & Quote(settingsPath)
    End If
End Function

Private Sub RunHidden(ByVal action As String)
    Dim shell As Object
    On Error GoTo HandleError

    Set shell = CreateObject("WScript.Shell")
    shell.Run SyncCommand(action), 0, False
    Exit Sub

HandleError:
    MsgBox "Cannot run Outlook Calendar Sync." & vbCrLf & Err.Description, _
        vbExclamation, "Outlook Calendar Sync"
End Sub

Private Function RunAndCapture(ByVal action As String) As String
    Dim shell As Object
    Dim fso As Object
    Dim outputPath As String
    Dim commandPath As String
    Dim commandFile As Object
    Dim stream As Object
    Dim output As String

    On Error GoTo HandleError

    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    outputPath = fso.BuildPath(shell.ExpandEnvironmentStrings("%TEMP%"), fso.GetTempName)
    commandPath = outputPath & ".cmd"

    Set commandFile = fso.CreateTextFile(commandPath, True, False)
    commandFile.WriteLine "@echo off"
    commandFile.WriteLine SyncCommand(action) & " > " & Quote(outputPath) & " 2>&1"
    commandFile.Close

    shell.Run Quote(commandPath), 0, True

    If fso.FileExists(outputPath) Then
        Set stream = fso.OpenTextFile(outputPath, 1, False)
        output = stream.ReadAll
        stream.Close

        On Error Resume Next
        fso.DeleteFile outputPath, True
        On Error GoTo 0
    End If

    On Error Resume Next
    If Len(commandPath) > 0 Then
        fso.DeleteFile commandPath, True
    End If
    On Error GoTo 0

    If Len(Trim$(output)) = 0 Then
        output = "No response from Outlook Calendar Sync."
    End If

    RunAndCapture = output
    Exit Function

HandleError:
    RunAndCapture = "Cannot run Outlook Calendar Sync." & vbCrLf & Err.Description
End Function

Public Sub OutlookSync_Start()
    MsgBox RunAndCapture("--start"), vbInformation, "Outlook Calendar Sync"
End Sub

Public Sub OutlookSync_Stop()
    MsgBox RunAndCapture("--stop"), vbInformation, "Outlook Calendar Sync"
End Sub

Public Sub OutlookSync_Status()
    MsgBox RunAndCapture("--status"), vbInformation, "Outlook Calendar Sync"
End Sub

Public Sub OutlookSync_Restart()
    MsgBox RunAndCapture("--restart"), vbInformation, "Outlook Calendar Sync"
End Sub

Public Sub OutlookSync_SyncSelected()
    MsgBox RunAndCapture("--sync-selected"), vbInformation, "Outlook Calendar Sync"
End Sub

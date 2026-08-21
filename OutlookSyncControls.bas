Attribute VB_Name = "OutlookSyncControls"
Option Explicit

' Edit this path after placing outlook_calendar_sync.py on the target PC.
Private Const SYNC_SCRIPT As String = "C:\Path\To\outlook_calendar_sync.py"

' Leave blank to use settings.py next to outlook_calendar_sync.py.
Private Const SETTINGS_FILE As String = ""

' Use python.exe directly. Outlook may fail when launching the Windows "py" launcher.
Private Const PYTHON_COMMAND As String = "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"

Private Function Quote(ByVal value As String) As String
    Quote = """" & Replace(value, """", """""") & """"
End Function

Private Function ExpandEnv(ByVal value As String) As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    ExpandEnv = shell.ExpandEnvironmentStrings(value)
End Function

Private Function SyncCommand(ByVal action As String) As String
    SyncCommand = Quote(ExpandEnv(PYTHON_COMMAND)) & " " & Quote(SYNC_SCRIPT) & " " & action

    If Len(Trim$(SETTINGS_FILE)) > 0 Then
        SyncCommand = SyncCommand & " --settings " & Quote(SETTINGS_FILE)
    End If
End Function

Private Sub RunHidden(ByVal action As String)
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run SyncCommand(action), 0, False
End Sub

Private Function RunAndCapture(ByVal action As String) As String
    Dim shell As Object
    Dim fso As Object
    Dim outputPath As String
    Dim command As String
    Dim stream As Object
    Dim output As String

    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    outputPath = fso.BuildPath(shell.ExpandEnvironmentStrings("%TEMP%"), fso.GetTempName)
    command = Quote(shell.ExpandEnvironmentStrings("%ComSpec%")) & " /d /s /c " & _
        Quote(SyncCommand(action) & " > " & Quote(outputPath) & " 2>&1")

    shell.Run command, 0, True

    If fso.FileExists(outputPath) Then
        Set stream = fso.OpenTextFile(outputPath, 1, False)
        output = stream.ReadAll
        stream.Close

        On Error Resume Next
        fso.DeleteFile outputPath, True
        On Error GoTo 0
    End If

    If Len(Trim$(output)) = 0 Then
        output = "No response from Outlook Calendar Sync."
    End If

    RunAndCapture = output
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

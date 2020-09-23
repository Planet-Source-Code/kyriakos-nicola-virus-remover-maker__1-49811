Attribute VB_Name = "modMain"
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Const WM_CLOSE = &H10 'close window
Public Const RE_RUN As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"
Public Const RE_RUNONCE As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\"
Public Const RE_RUNSERVICE As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunService\"

Public Function GSF(sValue As Integer)

Dim fso, SpecialFolder As String
Set fso = CreateObject("Scripting.FileSystemObject")
Select Case sValue
    Case Is = 0
        SpecialFolder = fso.GetSpecialFolder(0) 'Window folder
    Case Is = 1
        SpecialFolder = fso.GetSpecialFolder(1) 'System folder
End Select

GSF = SpecialFolder
SpecialFolder = ""
End Function

Public Function KillWindow(sName As String)
Dim WinWnd As Long
WinWnd = FindWindow("ThunderRT6FormDC", sName)
PostMessage WinWnd, WM_CLOSE, 0&, 0&
End Function

Public Function DelRegKey(sLocation As String, sName As String)
On Error Resume Next
Dim WScript
Set WScript = CreateObject("WScript.Shell")

Select Case sLocation
    Case Is = "Run"
        WScript.RegDelete RE_RUN & sName
    Case Is = "RunOnce"
        WScript.RegDelete RE_RUNONCE & sName
    Case Is = "RunService"
        WScript.RegDelete RE_RUNSERVICE & sName
End Select
Exit Function

Err:
If Err.Number = -2147024894 Then
    MsgBox "Registry entry not found!", vbExclamation, "Error!"
Else
    MsgBox "Critical Error!", vbCritical, "Error!"
End If
End Function

Public Sub RemoveFromWinINI(AppName As String, KeyName As String)

WritePrivateProfileString AppName, KeyName, "", GSF(0) & "\win.ini"

End Sub


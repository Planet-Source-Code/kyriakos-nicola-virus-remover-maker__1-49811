VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Virus Remover Maker by Kyriakos Nicola"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox txtVirusName 
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   4575
   End
   Begin VB.CheckBox chkWinINI 
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   5400
      Width           =   255
   End
   Begin VB.TextBox txtExeFile 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Text            =   "Test.exe"
      Top             =   6330
      Width           =   3855
   End
   Begin VB.OptionButton optCustom 
      Caption         =   "Custom"
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   1080
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optSystem 
      Caption         =   "System"
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton optWindows 
      Caption         =   "Windows"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdWriteEXE 
      Caption         =   "Write EXE"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtKeyName 
      BackColor       =   &H00EAEAEA&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox txtAppName 
      BackColor       =   &H00EAEAEA&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox txtRegRunService 
      BackColor       =   &H00EAEAEA&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CheckBox chkRunService 
      Caption         =   "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunService\"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4200
      Width           =   6855
   End
   Begin VB.TextBox txtRegRunOnce 
      BackColor       =   &H00EAEAEA&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CheckBox chkRunOnce 
      Caption         =   "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   6615
   End
   Begin VB.TextBox txtRegRun 
      BackColor       =   &H00EAEAEA&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CheckBox chkRun 
      Caption         =   "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   6255
   End
   Begin VB.TextBox txtWindow 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label Label8 
      Caption         =   "Name of the virus:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Path:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   7080
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   240
      X2              =   7080
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label5 
      Caption         =   "Key Name:"
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Application Name:"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "win.ini entry to Delete:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Registry entry to Delete:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Virus' default path:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Window Caption to kill process:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Menu b 
      Caption         =   "b"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PropBag As New PropertyBag

Public Function GSF(sValue As Integer)

Dim fso, SpecialFolder As String
Set fso = CreateObject("Scripting.FileSystemObject")
Select Case sValue
    Case Is = 0
        SpecialFolder = fso.GetSpecialFolder(0) 'Windows folder
    Case Is = 1
        SpecialFolder = fso.GetSpecialFolder(1) 'System folder
End Select

GSF = SpecialFolder
End Function

Private Sub CheckValue(sCheckBox As CheckBox, sTextBox As TextBox)

If sCheckBox.Value = 1 Then
    sTextBox.Enabled = True
    sTextBox.BackColor = vbWhite
    sTextBox.SetFocus
Else
    sTextBox.Enabled = False
    sTextBox.BackColor = &HEAEAEA
End If

End Sub

Private Sub chkRun_Click()

CheckValue chkRun, txtRegRun

End Sub

Private Sub chkRunOnce_Click()

CheckValue chkRunOnce, txtRegRunOnce

End Sub

Private Sub chkRunService_Click()

CheckValue chkRunService, txtRegRunService

End Sub

Private Sub chkWinINI_Click()

CheckValue chkWinINI, txtKeyName
CheckValue chkWinINI, txtAppName

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdWriteEXE_Click()

On Local Error GoTo Err

Dim StartPos As Long
Dim Buff As Variant
Dim msg As String

With PropBag
    .WriteProperty "VirusName", txtVirusName.Text
    .WriteProperty "File", txtFile.Text
    .WriteProperty "Caption", txtWindow.Text
    If chkRun.Value = 1 Then
        .WriteProperty "Run", txtRegRun.Text
    Else
        .WriteProperty "Run", vbNullString
    End If
    If chkRunOnce.Value = 1 Then
        .WriteProperty "RunOnce", txtRegRunOnce.Text
    Else
        .WriteProperty "RunOnce", vbNullString
    End If
    If chkRunService.Value = 1 Then
        .WriteProperty "RunService", txtRegRunService.Text
    Else
        .WriteProperty "RunService", vbNullString
    End If
    If chkWinINI.Value = 1 Then
        .WriteProperty "AppName", txtAppName.Text
        .WriteProperty "KeyName", txtKeyName.Text
    Else
        .WriteProperty "AppName", vbNullString
        .WriteProperty "KeyName", vbNullString
    End If
End With

FileCopy App.Path & "\vr.ex_", App.Path & "\" & txtExeFile.Text

Open App.Path & "\" & txtExeFile.Text For Binary As #1
    StartPos = LOF(1)
            
    Buff = PropBag.Contents
            
    Seek #1, LOF(1)
    Put #1, , Buff
    Put #1, , StartPos

Close #1

MsgBox "Exe File created without a problem!", vbInformation, "Finished"
Exit Sub

Err:
msg = "There was an error during compilation" & vbNewLine
msg = msg & vbNewLine & Err.Description
MsgBox msg, vbCritical, "Error"
End Sub

Private Sub optCustom_Click()

txtFile.Text = ""

End Sub

Private Sub optSystem_Click()

txtFile.Text = GSF(1) & "\"

End Sub

Private Sub optWindows_Click()

txtFile.Text = GSF(0) & "\"

End Sub

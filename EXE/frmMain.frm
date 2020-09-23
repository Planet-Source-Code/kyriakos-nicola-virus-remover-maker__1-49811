VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Virus Remover by Kyriakos Nicola"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtWindow 
      BackColor       =   &H00EAEAEA&
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H00EAEAEA&
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   4935
   End
   Begin VB.TextBox txtRegRun 
      BackColor       =   &H00EAEAEA&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox txtRegRunOnce 
      BackColor       =   &H00EAEAEA&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox txtRegRunService 
      BackColor       =   &H00EAEAEA&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox txtAppName 
      BackColor       =   &H00EAEAEA&
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txtKeyName 
      BackColor       =   &H00EAEAEA&
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "[VirusName] Remover"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label9 
      Caption         =   "RunService:"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "RunOnce:"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Run:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Window to Kill (Caption):"
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
      Left            =   360
      TabIndex        =   13
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "File to Delete:"
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
      Left            =   360
      TabIndex        =   12
      Top             =   720
      Width           =   1935
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
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "WinINI entry to Delete:"
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
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Application Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Key Name:"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   3600
      Width           =   1935
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

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdRemove_Click()

On Error Resume Next

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

KillWindow txtWindow.Text
Debug.Print "Closing Window."
Kill txtFile.Text
Debug.Print "Deleting File."
DelRegKey RE_RUN, txtRegRun.Text
DelRegKey RE_RUNONCE, txtRegRunOnce.Text
DelRegKey RE_RUNSERVICE, txtRegRunService.Text
Debug.Print "Deleting Registry Entries."
RemoveFromWinINI txtAppName.Text, txtKeyName.Text
Debug.Print "Deleting Entry from Win.INI"

MsgBox "Virus removed successfully!", vbInformation, "FINISH"

End Sub

Private Sub Form_Load()

On Error GoTo Err

Dim StartPos As Long
Dim varTemp As Variant

Dim byteArr() As Byte

Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
    Get #1, LOF(1) - 3, StartPos

    Seek #1, StartPos
    Get #1, , varTemp
    
    byteArr = varTemp
    PropBag.Contents = byteArr

    PropBag.WriteProperty "LOF", LOF(1)
    PropBag.WriteProperty "StartPos", StartPos
Close #1

With PropBag
    lblTitle.Caption = .ReadProperty("VirusName") & " Remover"
    txtFile.Text = .ReadProperty("File")
    txtWindow.Text = .ReadProperty("Caption")
    txtRegRun.Text = .ReadProperty("Run")
    txtRegRunOnce.Text = .ReadProperty("RunOnce")
    txtRegRunService.Text = .ReadProperty("RunService")
    txtAppName.Text = .ReadProperty("AppName")
    txtKeyName.Text = .ReadProperty("KeyName")
End With
Exit Sub

Err:
MsgBox "An Error Has Occured:" & vbNewLine & Err.Description, vbCritical, "Error"
End

End Sub


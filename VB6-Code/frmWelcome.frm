VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Welcome to Voice Navigation System"
   ClientHeight    =   3468
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   4164
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   3468
   ScaleWidth      =   4164
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Do not show this window at startup"
      Height          =   192
      Left            =   480
      TabIndex        =   6
      Top             =   3000
      Width           =   3132
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   372
      Left            =   720
      TabIndex        =   0
      Top             =   2400
      Width           =   1212
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   372
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Terminates Voice Navigation System"
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Frame fraOptionFrame 
      Caption         =   "Select any option"
      Height          =   1932
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   3132
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   492
         Left            =   240
         Picture         =   "frmWelcome.frx":0000
         ScaleHeight     =   492
         ScaleWidth      =   492
         TabIndex        =   7
         Top             =   1320
         Width           =   492
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "Show Command List"
         Height          =   372
         Index           =   2
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Width           =   1812
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "Run in System Tray"
         Height          =   372
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   840
         Width           =   1812
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "Show Application"
         Height          =   252
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1812
      End
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentIndex As Byte

Private Sub Check1_Click()
    Dim longstatus As Long
    Dim chkValue As String
    chkValue = Check1.Value
    longstatus = WritePrivateProfileString("TTS", "Check Welcome", chkValue, App.Path + "\Vns.ini")
    frmOptions.chkShowWel = Check1.Value
End Sub

Private Sub cmdExit_Click()
    frmMain.g_iFlag = FORMMAIN
    Unload Me
    End
End Sub

Private Sub cmdOK_Click()
Select Case CurrentIndex
    Case 0: 'Show Application
            frmMain.g_iFlag = FORMMAIN
            Unload Me
            
    Case 1: 'Run in Sys Tray
            Unload Me
            frmMain.g_iFlag = FORMMAIN
            frmMain.mnuFileSysTray_Click
    Case 2: 'Show Command List
            Unload Me
            frmMain.g_iFlag = FORMCOMMANDLIST
            frmCmdList.Show vbModal
End Select

End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdOk.ToolTipText = optOptions(CurrentIndex).Caption
End Sub

'Private Sub Command1_Click()
'   Unload Me
'   frmMain.mnuFile_Click
'End Sub

Private Sub Form_Load()
    CurrentIndex = 0
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Welcome 'Call Sub Welcome and greet the user
End Sub

Private Sub optOptions_Click(Index As Integer)
    CurrentIndex = Index
End Sub
Sub Welcome()
Dim HR As Byte
Dim Msg As String
    HR = Val(Hour(Time))
    Select Case HR
        Case 6 To 11: Msg = "Good Morning"
        Case 12 To 16: Msg = "Good Afternoon"
        Case 17 To 21: Msg = "Good Evening"
        Case 22 To 23, 0 To 5: Msg = "Its Time to Sleep !"
    End Select
    Speech2User Msg + " and welcome to Voice Navigation System !", 4
End Sub

Private Sub optOptions_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        CurrentIndex = Index
        cmdOK_Click
    End If
End Sub

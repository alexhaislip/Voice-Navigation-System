VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEleSpk 
   Caption         =   "Electronic Speech"
   ClientHeight    =   5088
   ClientLeft      =   132
   ClientTop       =   396
   ClientWidth     =   5856
   LinkTopic       =   "Form1"
   ScaleHeight     =   5088
   ScaleWidth      =   5856
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOptions 
      Caption         =   " Options"
      Height          =   1812
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   5652
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   372
         Left            =   4080
         TabIndex        =   9
         Top             =   1200
         Width           =   1212
      End
      Begin VB.CommandButton cmdSpeak 
         Caption         =   "&Speak"
         Height          =   372
         Left            =   4080
         TabIndex        =   8
         Top             =   720
         Width           =   1212
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
         Height          =   372
         Left            =   4080
         TabIndex        =   7
         Top             =   240
         Width           =   1212
      End
      Begin VB.Frame fraVoiceType 
         Caption         =   "Voice Type"
         Height          =   1452
         Left            =   2160
         TabIndex        =   6
         Top             =   120
         Width           =   1692
         Begin VB.OptionButton optVoiceType 
            Caption         =   "Male"
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1452
         End
         Begin VB.OptionButton optVoiceType 
            Caption         =   "Female"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   1452
         End
         Begin VB.OptionButton optVoiceType 
            Caption         =   "All"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1452
         End
      End
      Begin VB.ListBox List1 
         Height          =   1392
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1932
      End
   End
   Begin VB.Frame fraSpeed 
      Caption         =   "Speed"
      Height          =   732
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   5652
      Begin MSComctlLib.Slider Slider1 
         Height          =   372
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5412
         _ExtentX        =   9546
         _ExtentY        =   656
         _Version        =   393216
         Min             =   30
         Max             =   300
         SelStart        =   200
         TickStyle       =   3
         Value           =   200
      End
   End
   Begin RichTextLib.RichTextBox rtfText 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   2172
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   3831
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmEleSpk.frx":0000
   End
   Begin VB.PictureBox Picture1 
      Height          =   2052
      Left            =   3960
      ScaleHeight     =   2004
      ScaleWidth      =   1764
      TabIndex        =   3
      Top             =   120
      Width           =   1812
   End
End
Attribute VB_Name = "frmEleSpk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Private Sub cmdOpen_Click()
Dim sFile As String

    With frmMain.dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Text File (*.txt)|*.txt"
        .ShowOpen
        If Len(.fileName) = 0 Then
            Exit Sub
        End If
        sFile = .fileName
    End With
    frmEleSpk.rtfText.LoadFile sFile
    frmEleSpk.Caption = "Voice Navigation System [" & sFile & "]"
    
End Sub

Private Sub cmdStop_Click()
    frmMain.TTS.StopSpeaking
End Sub

Private Sub Form_Load()
     frmMain.g_iFlag = FORMELESPK
   
    'Deactivate DSR when this form loaded
    'frmMain.DSR.Deactivate
    frmMain.TTS.Speed = Slider1.Value
    'frmmain.TTS.
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    List1.Clear
    frmOptions.CountEngines List1
    List1.ListIndex = 0
    'For i = 0 To cnt
    
    'rtfText.Height = frmEleSpk.ScaleHeight - 1500
    'rtfText.Width = frmEleSpk.ScaleWidth - 1800
    'cmdOpen.Left = rtfText.Width + 480
    'cmdSpeak.Left = rtfText.Width + 480
    'cmdStop.Left = rtfText.Width + 480

    
End Sub

Private Sub Form_Resize()
    'rtfText.Height = frmEleSpk.ScaleHeight - 1500
    'rtfText.Width = frmEleSpk.ScaleWidth - 1800
    'cmdOpen.Left = rtfText.Width + 480
    'cmdSpeak.Left = rtfText.Width + 480
    'cmdStop.Left = rtfText.Width + 480
End Sub

Private Sub cmdSpeak_Click()
    frmMain.TTS.Speed = Slider1.Value
    If rtfText.Text = "" Then
        'Speech2User "There is nothing to read in this file", 0
        frmain.TTS.Speak "There is nothing to read"
    Else
        frmMain.TTS.CurrentMode = frmOptions.FindMode(List1.Text)
        frmMain.TTS.Speak rtfText.Text
        'Speech2User rtfText.Text, 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
     frmMain.g_iFlag = FORMMAIN

End Sub

Private Sub optVoiceType_Click(Index As Integer)

Dim i As Integer
    cnt = frmMain.TTS.CountEngines
    Select Case Index
        Case 0: 'All Engines
            List1.Clear
            frmOptions.CountEngines List1
        Case 1: 'Engines with Male Voices
            List1.Clear
            For i = 1 To cnt
                If (frmMain.TTS.Gender(i) = 1) Then
                    List1.AddItem (frmMain.TTS.ModeName(i))
                End If
            Next i
        Case 2: 'Engines with Female Voices
            List1.Clear
            For i = 1 To cnt
                If (frmMain.TTS.Gender(i) = 2) Then
                    List1.AddItem (frmMain.TTS.ModeName(i))
                End If
            Next i
    End Select
    'List1.ListIndex = 0 'For Marking 1st Item

End Sub

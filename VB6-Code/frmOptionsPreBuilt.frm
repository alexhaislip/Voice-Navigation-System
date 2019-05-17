VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5460
   ClientLeft      =   2568
   ClientTop       =   1500
   ClientWidth     =   6900
   Icon            =   "frmOptionsPreBuilt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3900
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3900
      ScaleWidth      =   5892
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   5892
      Begin VB.Frame fraSample2 
         Caption         =   "Text To Speech"
         Height          =   3828
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   5652
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1332
            Left            =   4200
            Picture         =   "frmOptionsPreBuilt.frx":000C
            ScaleHeight     =   1332
            ScaleWidth      =   1332
            TabIndex        =   30
            Top             =   360
            Width           =   1332
         End
         Begin VB.CommandButton cmdTestVoice 
            Caption         =   "&Test Voice"
            Height          =   372
            Left            =   4200
            TabIndex        =   29
            Top             =   2400
            Width           =   1212
         End
         Begin VB.Frame fraCtrlType 
            Caption         =   "User Response Type"
            Height          =   1572
            Left            =   2040
            TabIndex        =   20
            Top             =   1200
            Width           =   1932
            Begin VB.OptionButton optCtrlType 
               Caption         =   "Miscellaneous"
               Height          =   312
               Index           =   4
               Left            =   120
               TabIndex        =   28
               Top             =   1200
               Width           =   1572
            End
            Begin VB.OptionButton optCtrlType 
               Caption         =   "Combo Box"
               Height          =   312
               Index           =   3
               Left            =   120
               TabIndex        =   27
               Top             =   960
               Width           =   1692
            End
            Begin VB.OptionButton optCtrlType 
               Caption         =   "Text Areas"
               Height          =   312
               Index           =   2
               Left            =   120
               TabIndex        =   26
               Top             =   720
               Width           =   1572
            End
            Begin VB.OptionButton optCtrlType 
               Caption         =   "Input Box"
               Height          =   312
               Index           =   1
               Left            =   120
               TabIndex        =   25
               Top             =   480
               Width           =   1692
            End
            Begin VB.OptionButton optCtrlType 
               Caption         =   "Message Box"
               Height          =   312
               Index           =   0
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Value           =   -1  'True
               Width           =   1572
            End
         End
         Begin VB.Frame fraVoice 
            Caption         =   "Voice Type"
            Height          =   972
            Left            =   2040
            TabIndex        =   19
            Top             =   120
            Width           =   1932
            Begin VB.OptionButton optVoiceType 
               Caption         =   "Female"
               Height          =   192
               Index           =   2
               Left            =   120
               TabIndex        =   23
               Top             =   720
               Width           =   1452
            End
            Begin VB.OptionButton optVoiceType 
               Caption         =   "Male"
               Height          =   192
               Index           =   1
               Left            =   120
               TabIndex        =   22
               Top             =   480
               Width           =   1572
            End
            Begin VB.OptionButton optVoiceType 
               Caption         =   "All"
               Height          =   192
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Value           =   -1  'True
               Width           =   1692
            End
         End
         Begin VB.Frame fraSpeed 
            Caption         =   "Speed"
            Height          =   612
            Left            =   240
            TabIndex        =   17
            Top             =   3000
            Width           =   5172
            Begin MSComctlLib.Slider sldSpeed 
               Height          =   252
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   4932
               _ExtentX        =   8700
               _ExtentY        =   445
               _Version        =   393216
               Min             =   30
               Max             =   300
               SelStart        =   200
               TickStyle       =   3
               Value           =   200
            End
         End
         Begin VB.ListBox lstVoiceType 
            Height          =   2544
            ItemData        =   "frmOptionsPreBuilt.frx":0E7A
            Left            =   120
            List            =   "frmOptionsPreBuilt.frx":0E7C
            TabIndex        =   16
            Top             =   240
            Width           =   1812
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3300
      Index           =   0
      Left            =   480
      ScaleHeight     =   3300
      ScaleWidth      =   5892
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   5892
      Begin VB.Frame fraSample1 
         Caption         =   "Settings"
         Height          =   2628
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5532
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1812
            Left            =   3960
            Picture         =   "frmOptionsPreBuilt.frx":0E7E
            ScaleHeight     =   1812
            ScaleWidth      =   1452
            TabIndex        =   31
            Top             =   360
            Width           =   1452
         End
         Begin VB.CheckBox chkShowWel 
            Caption         =   "Do not show ""Welcome Window"" at start up"
            Height          =   252
            Left            =   480
            TabIndex        =   15
            Top             =   840
            Width           =   3492
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Disable all speech response"
            Height          =   252
            Left            =   480
            TabIndex        =   14
            Top             =   1800
            Width           =   2772
         End
         Begin VB.CheckBox chkRandom 
            Caption         =   "Disable random computer response"
            Height          =   252
            Left            =   480
            TabIndex        =   13
            Top             =   1320
            Width           =   3012
         End
         Begin VB.CheckBox chkShowTips 
            Caption         =   "&Show ""Tip of The Day"" at start up"
            Height          =   252
            Left            =   480
            TabIndex        =   12
            Top             =   360
            Width           =   3252
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4692
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   8276
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            Object.ToolTipText     =   "General Properties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Text To Speech"
            Key             =   "TTSCtrl"
            Object.ToolTipText     =   "Text to Speech Controls"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5688
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   10
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5688
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdOption 
      Caption         =   "Apply"
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   3
      Top             =   4932
      Width           =   1095
   End
   Begin VB.CommandButton cmdOption 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   4932
      Width           =   1095
   End
   Begin VB.CommandButton cmdOption 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   4932
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ArrTempINI(0 To 4) As IOToINI
Dim cnt As Byte
Dim i As Integer
Dim CurrentIndex As Byte
Dim mode As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub General_Check()
Dim lonStatus As String
Dim checkval As String
    'checkval = chkShowIntro.Value
    'lonstatus = WritePrivateProfileString("TTS", "Check Introduction", checkval, App.Path + "\Vns.ini")
    checkval = chkShowWel.Value
    lonStatus = WritePrivateProfileString("TTS", "Check Welcome", checkval, App.Path + "\Vns.ini")
    checkval = chkRandom.Value
    lonStatus = WritePrivateProfileString("TTS", "Check Random Speech", checkval, App.Path + "\Vns.ini")
    checkval = chkAll.Value
    lonStatus = WritePrivateProfileString("TTS", "Check All Response", checkval, App.Path + "\Vns.ini")
    'Unload Me
End Sub

Private Sub chkShowTips_Click()
     SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkShowTips.Value
     frmTip.chkLoadTipsAtStartup.Value = chkShowTips.Value
End Sub

Private Sub cmdOption_Click(Index As Integer)
    Select Case Index
        Case 0: 'OK
                If cmdOption(2).Enabled Then
                    'If tbsOptions.SelectedItem.Key = "TTSCtrl" Then
                        ApplyTemp2Global 'APPLY and OK
                        WriteToVnsINI
                    'ElseIf tbsOptions.SelectedItem.Key = "General" Then
                        General_Check
                   ' End If
                End If
                Unload Me
        Case 1: 'Cancel
                Unload Me ' frmmain.tts.GeneralDlg hWind, InTitle'CANCEL
                'End
        Case 2: 'Apply
                If tbsOptions.SelectedItem.Key = "TTSCtrl" Then
                    ApplyTemp2Global 'APPLY
                    WriteToVnsINI
                    cmdOption(2).Enabled = False
                ElseIf tbsOptions.SelectedItem.Key = "General" Then
                    General_Check
                End If
                cmdOption(2).Enabled = False
                'frmTip.chkLoadTipsAtStartup_Click
    '***********SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkShowTips.Value
    End Select
End Sub

Private Sub cmdTestVoice_Click()
Dim SliderSpeed As Integer
Dim CtrlText As String * 100
Dim VoiceType As String * 100

    'frmMain.g_iFlag = FORMMAIN 'assign to main Window
    
    '**Setting Control Type as Msgbox, TextBox etc...
    For i = 0 To optCtrlType.UBound
        If optCtrlType(i).Value = True Then
            CtrlText = optCtrlType(i).Caption
        End If
    Next i

'If lstVoiceType.ListIndex <= 0 Then Exit Sub
'**Setting Voice Type
    VoiceType = Trim(lstVoiceType.Text)
    
    'Get the mode of VoiceType
    mode = FindMode(Me.lstVoiceType.Text)
    frmMain.TTS.Select mode
    '**Setting Speed
    
    SliderSpeed = sldSpeed.Value
    frmMain.TTS.Speed = SliderSpeed '****CHANGE IF APPLIED*****
    If cmdTestVoice.Caption = "&Test Voice" Then
        cmdTestVoice.Caption = "&Stop"
        frmMain.TTS.Speak CtrlText + " has voice of " + _
        VoiceType + " with speed " + Str(SliderSpeed) + _
        " Words Per Minute" 'TEST VOICE
    Else
        cmdTestVoice.Caption = "&Test Voice"
        frmMain.TTS.StopSpeaking
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Dim ShowAtStartup As Long
    'center the form
    'Me.Left = Screen.Width / 2 - Me.Width / 2
    'Me.Top = Screen.Height / 2 - Me.Height / 2

Dim str1 As String
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    CurrentIndex = 0
    CountEngines lstVoiceType 'Counts TTS Engines and adds to List
    '*************Set initial ctrl index
    'TempCtrlIndex = optCtrlType(0).Index
    'lstVoiceType.ListIndex = 0 'For Marking 1st Item
    ReadGlobal2Temp
    sldSpeed.Value = ArrTempINI(0).CtrlSpeed
    'For i = 0 To 4
    '    List1.AddItem ArrTempINI(i).CtrlVoice
    'Next i
    'Exit Sub
    Dim ValComp As Integer
    str1 = ArrTempINI(0).CtrlVoice
    For i = 0 To lstVoiceType.ListCount
    ValComp = StrComp(Trim(str1), Trim(lstVoiceType.List(i)), vbTextCompare)
            If ValComp = 0 Then
            'MsgBox Str1 + "=" + lstVoiceType.List(i)
            lstVoiceType.ListIndex = i
            Exit For
        End If
    Next i
    
    cmdOption(2).Enabled = True 'Disable Apply at Load
    
    
    ' Check value for frmTip
    'ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 0)
    'If ShowAtStartup = 0 Then
    '    chkShowTips.Value = 0
    'Else
    '    chkShowTips.Value = 1
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.g_iFlag = FORMMAIN
End Sub

Private Sub optCtrlType_Click(Index As Integer)
    CurrentIndex = Index 'MAKE IT LATER
        sldSpeed.Value = ArrTempINI(Index).CtrlSpeed
        For i = 0 To lstVoiceType.ListCount
            If StrComp(Trim(ArrTempINI(Index).CtrlVoice), Trim(lstVoiceType.List(i)), vbTextCompare) = 0 Then
                lstVoiceType.ListIndex = i
                Exit For
            End If
        Next i
End Sub

Private Sub optVoiceType_Click(Index As Integer)
    Select Case Index
        Case 0: 'All Engines
            lstVoiceType.Clear
            CountEngines lstVoiceType
        Case 1: 'Engines with Male Voices
            lstVoiceType.Clear
            For i = 1 To cnt
                If (frmMain.TTS.Gender(i) = 2) Then
                    lstVoiceType.AddItem (frmMain.TTS.ModeName(i))
                End If
            Next i
        Case 2: 'Engines with Female Voices
            lstVoiceType.Clear
            For i = 1 To cnt
                If (frmMain.TTS.Gender(i) = 1) Then
                    lstVoiceType.AddItem (frmMain.TTS.ModeName(i))
                End If
            Next i
    End Select
    lstVoiceType.ListIndex = 0 'For Marking 1st Item
End Sub

Private Sub sldSpeed_Click()
    'cmdOption(2).Enabled = True
End Sub

Private Sub sldSpeed_Scroll()

    With ArrTempINI(CurrentIndex) 'consider
        .CtrlSpeed = sldSpeed.Value
        .CtrlType = optCtrlType(CurrentIndex).Caption
        .CtrlVoice = Trim(lstVoiceType.Text)
    End With

End Sub

Private Sub tbsOptions_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    cmdOption(2).Enabled = True
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 300
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
End Sub
Sub WriteToVnsINI()
Dim lonStatus As Long
Dim StrSpeed As String
Dim CtrlVoice As String
    For i = 0 To 4
        With ArrGlobalINI(i)
            StrSpeed = Trim(Str(.CtrlSpeed))
            CtrlVoice = Trim(.CtrlVoice)
            lonStatus = WritePrivateProfileString("TTS", .CtrlType + " Speed", StrSpeed, App.Path + "\Vns.ini")
            lonStatus = WritePrivateProfileString("TTS", .CtrlType + " Speaker", CtrlVoice, App.Path + "\Vns.ini")
        End With
    Next i
'This is to ascertain if DATA properly written to INI file
'If lonStatus = 0 Then
'    MsgBox "Error"
'Else
'    MsgBox "Success"
'End If
End Sub
Public Sub ApplyTemp2Global()
    For i = 0 To 4
        With ArrGlobalINI(i)
            .CtrlSpeed = ArrTempINI(i).CtrlSpeed
            .CtrlType = ArrTempINI(i).CtrlType
            .CtrlVoice = ArrTempINI(i).CtrlVoice
        End With
    Next i
End Sub

Sub ReadGlobal2Temp()
    For i = 0 To 4
        With ArrTempINI(i)
            .CtrlSpeed = ArrGlobalINI(i).CtrlSpeed
            .CtrlType = ArrGlobalINI(i).CtrlType
            .CtrlVoice = ArrGlobalINI(i).CtrlVoice
        End With
    Next i
End Sub
Public Sub CountEngines(List1 As ListBox)
    cnt = frmMain.TTS.CountEngines
    For i = 1 To cnt
        List1.AddItem (frmMain.TTS.ModeName(i))
    Next i
End Sub


'Returns the mode of voice type
Public Function FindMode(VoiceType As String) As Integer
Dim mode As Integer

    Select Case Trim(VoiceType)
       Case "Mary":
           mode = 1
       Case "Mary (for Telephone)":
           mode = 2
       Case "Mike":
           mode = 3
       Case "Mike (for Telephone)":
           mode = 4
       Case "Sam":
           mode = 5
       Case "Mary in Space":
           mode = 6
       Case "Mary in Hall":
           mode = 7
       Case "Mary in Stadium":
           mode = 8
       Case "RoboSoft Six":
           mode = 9
       Case "RoboSoft Five":
           mode = 10
       Case "RoboSoft Four":
           mode = 11
       Case "RoboSoft One":
           mode = 12
       Case "RoboSoft Two":
           mode = 13
       Case "RoboSoft Three":
           mode = 14
       Case "Mike in Hall":
           mode = 15
       Case "Mike in Stadium":
           mode = 16
       Case "Mike in Space":
           mode = 17
    End Select
    FindMode = mode
End Function

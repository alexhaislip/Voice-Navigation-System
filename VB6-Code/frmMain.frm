VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Object = "{4E3D9D11-0C63-11D1-8BFB-0060081841DE}#1.0#0"; "Xlisten.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Voice Navigation System"
   ClientHeight    =   5592
   ClientLeft      =   48
   ClientTop       =   600
   ClientWidth     =   8904
   LinkTopic       =   "Form1"
   ScaleHeight     =   5592
   ScaleWidth      =   8904
   Begin MSComctlLib.ProgressBar pgrBar 
      Height          =   2892
      Left            =   6240
      TabIndex        =   3
      Top             =   360
      Width           =   612
      _ExtentX        =   1080
      _ExtentY        =   5101
      _Version        =   393216
      Appearance      =   1
      Max             =   65535
      Orientation     =   1
      Scrolling       =   1
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   3372
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   5532
      _ExtentX        =   9758
      _ExtentY        =   5948
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   6960
      Top             =   0
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      HelpFile        =   "Vns.hlp"
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   3000
   End
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR DSR 
      Height          =   372
      Left            =   7560
      OleObjectBlob   =   "frmMain.frx":00AE
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   372
   End
   Begin HTTSLibCtl.TextToSpeech TTS 
      Height          =   324
      Left            =   6360
      OleObjectBlob   =   "frmMain.frx":00D2
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   444
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSysTray 
         Caption         =   "&Run in System Tray"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuUtils 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuUtilsSystemNav 
         Caption         =   "&System Navigator"
      End
      Begin VB.Menu mnuUtilsUser 
         Caption         =   "&User Commands"
      End
      Begin VB.Menu mnuUtilsESpeech 
         Caption         =   "&Electronic Speech"
      End
      Begin VB.Menu mnuUtilsPFour 
         Caption         =   "&Game Plot Four"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewVoiceCommands 
         Caption         =   "&Command List"
      End
      Begin VB.Menu mnuViewSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpIntro 
         Caption         =   "&Introduction"
      End
      Begin VB.Menu mnuHelpContent 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help on..."
      End
      Begin VB.Menu mnuHelpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********* API Declares for No Multiple Instances of Application
Private Declare Function FindWindow Lib "user32" Alias "FindWindowW" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function ShowWindow Lib "user32" _
(ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'*********************End No Multiple Instances

Public gRestoreGrammar As Boolean   'Takes Grammar from file/Grammar from string
Public g_iFlag As Integer

Public gGrammarFile As String 'Contains the CFG file name

Public RunningInTray As Boolean 'To check weather the application is running in system tray or not
Public SpeakingOver As Boolean
Public flag As Boolean

Private Sub DSR_VUMeter(ByVal beginhi As Long, ByVal beginlo As Long, ByVal level As Long)
    pgrBar.Value = level
End Sub

'Public WelcomeCount As Integer
Private Sub Form_Initialize()
 
 'Procedure for running the Application Instance Only Once
 
    If App.PrevInstance Then
        Dim FWReturn As Long
        Dim SWReturn As Long
        Dim lpClassName As String
        FWReturn = FindWindow(lpClassName, "frmMain")
        SWReturn = ShowWindow(FWReturn, 3)
        
        End
    End If
    'g_iFlag = FORMMAIN
'End NoMultiple Instances, since more than one App instance can Hog up Computer Resources
End Sub

Private Sub Form_Resize()
'    rtfText.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

'Private Sub mnuFileOpen_Click()
Public Sub mnuFileOpen_Click()
    Dim sFile As String

    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.fileName) = 0 Then
            Exit Sub
        End If
        sFile = .fileName
    End With
    frmMain.rtfText.LoadFile sFile
    frmMain.Caption = "Voice Navigation System [" & sFile & "]"

End Sub

'Private Sub mnuFileRead_Click()
Public Sub mnuFileRead_Click()
    If mnuFileRead.Caption = "Rea&d" Then
        'DSR.Deactivate
        If rtfText.Text = "" Then
            Speech2User "There is nothing to read in this file", 0
        Else
            mnuFileRead.Caption = "&Stop"
            Speech2User rtfText.Text, 2
        End If
    ElseIf mnuFileRead.Caption = "&Stop" Then
            mnuFileRead.Caption = "Rea&d"
            frmMain.TTS.StopSpeaking
    End If
End Sub

Private Sub mnuHelpAbout_Click()
'    frmAddCmd.Show
End Sub

Private Sub mnuHelpContent_Click()
    'ShowHelp vns.hlp
   ' With dlgCommonDialog
        '.HelpContext = "100"
        '.HelpKey = ID_aboutVNS
    '    .HelpFile = App.Path + "\Vns.hlp"
        '.h
     '   .HelpCommand
      '  .ShowHelp
    With dlgCommonDialog
        .HelpFile = App.Path & "\Chessrm1.hlp"
        .HelpCommand = cdlHelpContents
        .ShowHelp
    End With
    'End With
End Sub


Private Sub mnuUtilsDiskNav_Click()
    gRestoreGrammar = False
    frmDiskNav.Show vbModal
End Sub

Private Sub mnuUtilsESpeech_Click()
    frmEleSpk.Show vbModal
End Sub

Private Sub mnuUtilsPFour_Click()
'Dim TitlePLUSMenuBar As Long
'Have grammer EngLetters for Plot Fout A-H
'DSR.GrammarFromFile App.Path & "\" & "EngLetters.Cfg"
'TitlePLUSMenuBar = Me.Height - Me.ScaleHeight

'    frmGame.Width = Me.ScaleWidth
'    frmGame.Height = Me.ScaleHeight

    frmGame.Show vbModal
'    frmGame.Move frmMain.Left, Me.Top + TitlePLUSMenuBar, frmMain.ScaleWidth, frmMain.ScaleHeight
End Sub

Private Sub mnuUtilsSystemNav_Click()
    frmSysNavigate.Show vbModal
    'frm.Show vbModal
End Sub

Private Sub mnuUtilsUser_Click()
If mnuUtilsUser.Caption = "&User Commands" Then
    frmUserCommands.Show ', Me.hWnd
    mnuUtilsUser.Caption = "Unload " & mnuUtilsUser.Caption
Else
    Unload frmUserCommands
    mnuUtilsUser.Caption = "&User Commands"
End If
End Sub

Private Sub mnuViewVoiceCommands_Click()
    g_iFlag = FORMCOMMANDLIST
    frmCmdList.Show vbModal
End Sub

Private Sub Timer1_Timer()
'On Error GoTo label
On Error Resume Next 'CHECK This line has a potential Error
'Rectify if problems exitsts with frmTip
frmSplash.Show ' vbModal
    frmSplash.Refresh
       
    Timer1.Enabled = False
    'MousePointer = vbHourglass

    engine = DSR.Find("MfgName=Microsoft;Grammars=1")

    DSR.Select engine
    
    '*****First load standard CFG
    DSR.GrammarFromFile App.Path + "\ExeStdFile.cfg"
    'Dim abc
    'abc = Dir(App.Path + "\EngLetters.cfg")
    'MsgBox abc
    'DSR.GrammarFromFile App.Path + "\EngLetters.cfg"
    
    DSR.initialized = 1
    DSR.Activate
    frmMain.SpeakingOver = True 'To enable Welcome Initially
    If Val(chkString) = 0 Then '0 To Show Welcome
     Unload frmSplash
       g_iFlag = FORMWELCOME
        frmWelcome.Show vbModal    'Temporarily Hidden
        frmTip.Show vbModal 'Show Tip of the day form
    Else
     Unload frmSplash
        g_iFlag = FORMMAIN
        'frmTip.Show vbModal
    End If
'    MousePointer = vbNormal
    Exit Sub
label:
    MsgBox "Error" + Err.Description + Err.Source
End Sub
Private Sub TTS_SpeakingDone()

    If g_iFlag = FORMELESPK Then Exit Sub
    'If frmEleSpk.mnuFileRead.Caption = "&Stop" Then
    '    frmEleSpk.mnuFileRead.Caption = "Rea&d"
    'End If
    If g_iFlag <> FORMGAME Then
        If (gRestoreGrammar) Then
            DSR.GrammarFromFile App.Path + "\" + gGrammarFile + ".def"
        Else
            DSR.GrammarFromString frmSysNavigate.strGrammar
        End If
    ElseIf g_iFlag = FORMGAME Then
        DSR.GrammarFromFile App.Path + "\EngLetters.cfg"
    
    End If
    'MsgBox Str(g_iFlag)
    'MsgBox "Speaking done"
    
    DSR.Activate
    'MsgBox "activated"
    SpeakingOver = True
    
End Sub

Private Sub DSR_PhraseFinish(ByVal flags As Long, ByVal beginhi As Long, ByVal beginlo As Long, ByVal endhi As Long, ByVal endlo As Long, ByVal Phrase As String, ByVal parsed As String, ByVal results As Long)

Dim IsGeneralCFG As Boolean
Dim i As Long
Dim result As Integer
Dim Index As Integer

    If (parsed = "") Then
        If frmOptions.chkRandom.Value = 1 Then Exit Sub '*****Disable Speech Responses
        'DSR.Deactivate
        If g_iFlag = FORMMAIN Then
            If (Rnd > 0.5) Then
                Speech2User "Please command me so that i can help you ", 4
            Else
                Speech2User "I didn't understand you", 4
            End If
        ElseIf g_iFlag = FORMDISKNAV Then
            If (Rnd > 0.5) Then
                Speech2User "Please say any one of the command from the list view", 4
            Else
                Speech2User "I didn't understand you", 4
            End If
            
        ElseIf g_iFlag = FORMWELCOME Then
            If (Rnd > 0.5) Then
                Speech2User "Please say any one of the following three options", 4
            Else
                Speech2User "I didn't understand you", 4
            End If
        
        ElseIf g_iFlag = FORMCOMMANDLIST Then
            If (Rnd > 0.5) Then
                Speech2User "Please say any one of the command from the command list", 4
            Else
                Speech2User "I didn't understand you", 4
            End If
        ElseIf g_iFlag = FORMGAME Then
            If (parsed = "") Then
                If (Rnd > 0.5) Then
                    Speech2User "Say coloumn numbers to play ", 4
                Else
                    Speech2User "I didn't understand you", 4
                End If
        End If
        
        End If
    Else 'This executes when SOMETHING is Understood by Computer
        If g_iFlag = FORMGAME Then
        
            Select Case parsed
                Case "A": frmGame.lbla_Click 0
                Case "B": frmGame.lblb_Click 0
                Case "C": frmGame.lblc_Click 0
                Case "D": frmGame.lbld_Click 0
                Case "E": frmGame.lble_Click 0
                Case "F": frmGame.lblf_Click 0
                Case "G": frmGame.lblg_Click 0
                Case "H": frmGame.lblh_Click 0
            End Select
                
           Exit Sub
        Else
            'MsgBox parsed
            IsGeneralCFG = General_CFG(parsed)
            If (IsGeneralCFG = True) Then ' EXE Files
                'DSR.Deactivate
                'TTS.Speak Phrase
                'MsgBox parsed
                Speech2User Phrase, 4
                'Dim FileFound As String
                'FileFound = Dir(parsed)
                'If FileFound <> "" Then
                    Shell parsed, vbNormalFocus
                'Else
                '    MsgBox "File not found", vbInformation, "Warning!"
                '    Exit Sub
                'End If
            Else 'this executes when NO EXE TO RUN
                'If Grammar form string
                'MsgBox "grestoregarmmar=" + Str(gRestoreGrammar)
                If gRestoreGrammar = False And g_iFlag = FORMDISKNAV Then
                    result = ListCompare(parsed)
                    'MsgBox result
                    'frmsysnavigate.dirListView.ListIndex = ListCompare(parsed)
                    If result <> 0 Then
                        frmSysNavigate.ListView1.ListItems.item(ListCompare(parsed)).Selected = True
                        frmSysNavigate.ListView1_Click
                        'MsgBox frmsysnavigate.ListView1.SelectedItem
                    Else
                        Index = FindIndex(parsed)
                        MsgBox Str(Index)
                        If Index Then
                            'MsgBox frmSysNavigate.drvListView.Drive
                            frmSysNavigate.dirListView.ListIndex = Index
                        Else
                            'Index = 0 'when user say C drive or D Drive so on
                            'For i = 0 To frmSysNavigate.drvListView.ListCount - 1
                            '    If parsed = frmSysNavigate.drvListView.List(i) Then
                        '            ''frmSysNavigate.drvListView.Drive = frmSysNavigate.drvListView.List(i)
                                      'frmsysnavigate.drvListView.ListIndex=i
                        '        End If
                        '    Next i
                        End If
                    End If
                    
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    dlgCommonDialog.HelpFile = App.HelpFile
    ReadFromVnsINI
    
    'rtfText.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    'rtfText.Enabled = False
    'RunningInTray = False
    gRestoreGrammar = True
'    Unload frmSplash
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Asks before terminating application
'FOR FINAL RELEASE
'    ans = MsgBox("Quit Voice Navigation System ?", vbYesNo + vbExclamation, "Confirm ...")
'        If ans = vbNo Then
'            Cancel = 1
'            Exit Sub
'        Else
            Unload Me
            End
'        End If
    
    'DoUnLoadPreCheck UnloadMode
    
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hWnd = Me.hWnd
    VBGTray.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, VBGTray)
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
    End
End Sub

Public Sub mnuFileSysTray_Click()
    If mnuFileSysTray.Caption = "&Run in System Tray" Then
        mnuFileSysTray.Caption = "&Show Application"
        RunningInTray = True  '**** adding to check
        Call GoSystemTray 'To Run in Tray
    Else
        frmMain.Show
        mnuFileSysTray.Caption = "&Run in System Tray"
        
        VBGTray.cbSize = Len(VBGTray)
        VBGTray.hWnd = Me.hWnd
        VBGTray.uId = vbNull
        RunningInTray = False '****adding to check
        Call Shell_NotifyIcon(NIM_DELETE, VBGTray)
    End If
End Sub

Private Sub mnuViewOptions_Click()
    g_iFlag = FORMOPTION
    frmOptions.Show vbModal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static lngMsg As Long
Static blnFlag As Boolean
    
    lngMsg = x / Screen.TwipsPerPixelX
    
    If blnFlag = False Then
        blnFlag = True
        Select Case lngMsg
          'doubleclick
            Case WM_LBUTTONDBLCLICK
                g_iFlag = FORMMAIN
                Me.Show
                Me.mnuFileSysTray.Caption = "&Run in System Tray"
                'right-click
            Case WM_RBUTTONUP
                Me.PopupMenu mnuFile
          End Select
          blnFlag = False
    End If
End Sub

Private Sub GoSystemTray()
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hWnd = Me.hWnd
    VBGTray.uId = vbNull
    VBGTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    VBGTray.ucallbackMessage = WM_MOUSEMOVE
    
    VBGTray.hIcon = Me.Icon
    'tooltiptext
    VBGTray.szTip = Me.Caption & vbNullChar
    
    Call Shell_NotifyIcon(NIM_ADD, VBGTray)
    
    App.TaskVisible = False   'remove application from taskbar
    Me.Hide
    
End Sub

Private Function General_CFG(InParsed As String) As Boolean
Dim AlertMessage As String
    Select Case InParsed
        
        Case "Command List":
            'AlertMessage = "Command List Window is Already open"
            'On Error GoTo Alert
            
            If g_iFlag = FORMCOMMANDLIST Then
                Speech2User "Command List Window is Already open", 0
            ElseIf g_iFlag = FORMWELCOME Then
                'frmWelcome.optOptions(2).Value = True
                g_iFlag = FORMCOMMANDLIST
                Unload frmWelcome
                frmCmdList.Show vbModal
            ElseIf g_iFlag = FORMMAIN Then
                g_iFlag = FORMCOMMANDLIST
                frmCmdList.Show vbModal
            End If
            General_CFG = False
        Case "Disk Navigator"
            ' Must do Validation
            g_iFlag = FORMDISKNAV
            frmDiskNav.Show vbModal
            General_CFG = False
        Case "System Tray":
            'Unload frmWelcome
            'If g_iFlag = 2 Then
            '    Unload frmWelcome
            If RunningInTray = False Then 'if not running in system tray
                'frmWelcome.optOptions(1).Value = True
                If g_iFlag = FORMWELCOME Then       'if welcome window is open
                    Unload frmWelcome     'unload welcome window
                End If
                frmMain.mnuFileSysTray_Click 'call run in system tray
            End If
            g_iFlag = FORMMAIN                   'assign flag to main window
            General_CFG = False
        Case "Show Application":
        
           'If Me.Caption = "Welcome to Voice Navigation System" Then
           
            If RunningInTray = False Then
               ' frmWelcome.optOptions(0).Value = True
                If g_iFlag = FORMWELCOME Then
                    Unload frmWelcome
                End If
            Else
                frmMain.mnuFileSysTray_Click
            End If
            'frmMain.Show
            g_iFlag = FORMMAIN
            General_CFG = False
        Case "Options":
                'Check for formmain if it is then display Option
                If g_iFlag = FORMMAIN Then
                g_iFlag = FORMOPTION
                frmOptions.Show vbModal
            Else
               ' DSR.Deactivate
                Speech2User "Cannot open Options Please activate Main window", 4
            End If
            General_CFG = False
        Case "Open"
            If g_iFlag = FORMMAIN Then
                frmMain.mnuFileOpen_Click
            Else
                'DSR.Deactivate
                Speech2User "Please activate main window cannot open file here", 4
            End If
        Case "Read"
            If g_iFlag = FORMMAIN Then
                frmMain.mnuFileRead_Click
            Else
                'DSR.Deactivate
                Speech2User "Please activate main window cannot perform read operation here", 4
            End If
        Case Else
            'General_CFG = True
            'If Me.Caption = "Command List" Or Me.Caption = "Voice Navigation System" Then 'me.name is retriving main form but has to give the frmcmdlist
            If gRestoreGrammar Then
                If g_iFlag = FORMCOMMANDLIST Or g_iFlag = FORMMAIN Then
                    General_CFG = True
                Else
                    General_CFG = False
                    'DSR.Deactivate
                    Speech2User "Cannot Run application here Please open Command list", 4
                End If
            Else
                General_CFG = False
            End If
        End Select
Exit Function
Alert:
        Speech2User AlertMessage, 0
End Function
Private Function ListCompare(InParsed As String) As Integer
    'For i = 0 To frmsysnavigate.dirListView.ListCount - 1
    '    'For j = 4 To Len(frmDiskNav.Dir1.List(i))
    '
    '    'MsgBox "Dir=" + frmDiskNav.Dir1.List(i) + "  Parsed=" + InParsed
    '    If frmsysnavigate.dirListView.List(i) = InParsed Then
    '        ListCompare = i
    '        Exit For
    '    End If
    'Next i
Dim flag As Boolean
Dim fileName As String
    flag = False
    
    For i = 1 To frmSysNavigate.ListView1.ListItems.Count
        fileName = frmCmdList.ExtractFileName(frmSysNavigate.ListView1.ListItems.item(i))
        If fileName = InParsed Then
            flag = True
            Exit For
        End If
    Next i
    If flag = True Then
        ListCompare = i
    Else
        ListCompare = 0
    End If
End Function
Private Function FindIndex(InParsed As String) As Integer
Dim Index As Integer
Dim subDir As String
Dim c As String * 1
Dim InPath As String
    Index = 0
    InPath = frmSysNavigate.dirListView.Path
    If Right(InPath, 1) = "\" Then
        InPath = Left(InPath, 2)
    End If
    InPath = InPath & "\"
    For i = Len(InPath) To 1 Step -1
        c = Mid(InPath, i, 1)
        If c <> "\" And c <> ":" Then
            subDir = subDir & c
        Else
            Index = Index - 1
            subDir = StrReverse(subDir)
            If UCase(subDir) = UCase(InParsed) Then
                'flag = True
                Exit For
            End If
            subDir = ""
        End If
    Next i
    If Index <> 0 Then
        Index = Index + 1
    End If
    FindIndex = Index
End Function



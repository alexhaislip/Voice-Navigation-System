Attribute VB_Name = "modGlobal"
'FOR FINAL RELEASE
'*************** Declarations for running in System Tray ***********************
Public Const WM_LBUTTONDBLCLICK = &H203
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public chkString As String

'*****Flag to keep track of activated windows
Public Const FORMMAIN = 1          'g_iFlag=1 for frmMain
Public Const FORMWELCOME = 2       'g_iFlag=2 for frmWelcome
Public Const FORMCOMMANDLIST = 3   'g_iFlag=3 for frmCommandList
Public Const FORMOPTION = 4        'g_iFlag=4 for frmOption
Public Const FORMDISKNAV = 5       'g_iFlag=5 for frmDiskNav
Public Const FORMGAME = 6          'g_iFlag=6 for frmGame
Public Const FORMELESPK = 7        '


Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uId As Long
  uFlags As Long
  ucallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Public VBGTray As NOTIFYICONDATA
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'*************** System Tray Declarations END ***********************

'*************** Declarations for Reading and Writing to Vns.ini ********
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'End of Vns.ini Declarations

'Global Type variables Written and Read from VNS.INI
Public Type IOToINI 'This User type is for TTS Options
    CtrlType As String
    CtrlSpeed As Integer
    CtrlVoice As String
End Type
Public gStrFixed As String * 100 'Fixed String for Reading from VNS.INI
Public ArrGlobalINI(0 To 4) As IOToINI

'***********************This is to be done NEXT
Public Sub ReadFromVnsINI() '******READ FROM INI FILE******
Dim lonStatus As Long
Dim StrSpeed As String
Dim CtrlVoice As String
Dim j As Integer

ArrGlobalINI(0).CtrlType = "Message Box"
ArrGlobalINI(1).CtrlType = "Input Box"
ArrGlobalINI(2).CtrlType = "Text Areas"
ArrGlobalINI(3).CtrlType = "Combo Box"
ArrGlobalINI(4).CtrlType = "Miscellaneous"

For j = 0 To 4 'Later Direct ie Entire Section
    lonStatus = GetPrivateProfileString("TTS", ArrGlobalINI(j).CtrlType + " Speed", "?!?", gStrFixed, 100, App.Path + "\Vns.ini")
    StrSpeed = Left$(gStrFixed, lonStatus)
    ArrGlobalINI(j).CtrlSpeed = Val(StrSpeed)
    
    lonStatus = GetPrivateProfileString("TTS", ArrGlobalINI(j).CtrlType + " Speaker", "?!?", gStrFixed, 100, App.Path + "\Vns.ini")
    CtrlVoice = Left$(gStrFixed, lonStatus)
    ArrGlobalINI(j).CtrlVoice = CtrlVoice
Next j

lonStatus = GetPrivateProfileString("TTS", "Check Welcome", "?!?", gStrFixed, 100, App.Path + "\Vns.ini")
chkString = Left$(gStrFixed, lonStatus) ' chkString assigned as a CHECK FOR FRMWELCOME
frmOptions.chkShowWel.Value = Val(chkString)

lonStatus = GetPrivateProfileString("TTS", "Check Random Speech", "?!?", gStrFixed, 100, App.Path + "\Vns.ini")
frmOptions.chkRandom.Value = Val(Left$(gStrFixed, lonStatus))

lonStatus = GetPrivateProfileString("TTS", "Check All Response", "?!?", gStrFixed, 100, App.Path + "\Vns.ini")
frmOptions.chkAll.Value = Val(Left$(gStrFixed, lonStatus))

lonStatus = GetPrivateProfileString("TTS", "Grammar File", "?!?", gStrFixed, 100, App.Path + "\Vns.ini")
frmMain.gGrammarFile = Left$(gStrFixed, lonStatus)


End Sub '**********END READING & ASSIGINING GLOBAL VARS


Sub Speech2User(ByVal Message As String, ByVal Index As Integer)
'If bIsTTS = True And Index = 0 Then GoTo THISMSG
'Else
'Exit Sub
'endif
'TODO Random speech and All speech
   If frmOptions.chkAll.Value = 1 Then GoTo THISMSG
   Dim mode As Byte
   If frmMain.SpeakingOver = True Then
        frmMain.DSR.Deactivate
        
'    Else
'        frmMain.TTS.StopSpeaking
    End If
   'If frmMain.SpeakingOver = True Then
   'frmMain.DSR.Deactivate ' DeActivate Speech Recognition for
                          ' Speaking a sentence to the User
   
   mode = 1
   Select Case ArrGlobalINI(Index).CtrlVoice
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
   frmMain.SpeakingOver = False
   
   'frmMain.TTS.CurrentModAe = mode
   frmMain.TTS.Select mode
   frmMain.TTS.Speed = ArrGlobalINI(Index).CtrlSpeed
   frmMain.TTS.Speak Message
THISMSG:
   If Index = 0 Then
        MsgBox Message, vbInformation
   End If
'frmMain.DSR.Activate ' Activate Speech Recognition

End Sub

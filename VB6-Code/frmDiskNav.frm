VERSION 5.00
Begin VB.Form frmDiskNav 
   Caption         =   "Disk Navigator"
   ClientHeight    =   5304
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7884
   LinkTopic       =   "Form1"
   ScaleHeight     =   5304
   ScaleWidth      =   7884
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Make CFG"
      Height          =   492
      Left            =   480
      TabIndex        =   5
      Top             =   4680
      Width           =   1452
   End
   Begin VB.DirListBox Dir1 
      Height          =   3096
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2172
   End
   Begin VB.ListBox List1 
      Height          =   3120
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   2172
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2172
   End
   Begin VB.FileListBox File1 
      Height          =   456
      Left            =   4680
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.ListBox List2 
      Height          =   3120
      Left            =   4680
      TabIndex        =   0
      Top             =   840
      Width           =   2292
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4560
      TabIndex        =   7
      Top             =   480
      Width           =   2412
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sub Directories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   2172
   End
End
Attribute VB_Name = "frmDiskNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const tagStart = vbNewLine + "<start>="
'Public strGrammar As String

Private Sub Command1_Click()
Dim cnt As Integer

    frmMain.gRestoreGrammar = False
    cnt = List1.ListCount

    strGrammar = "[Grammer]" & vbNewLine & "langid = 1033" & vbNewLine & "type=cfg" & vbNewLine & vbNewLine & "[<start>]"
    
    For i = 0 To cnt - 1
        strGrammar = strGrammar & tagStart & List1.List(i) & " " & """" & List1.List(i) & """"
    Next i
    
    'MsgBox strGrammar
    frmMain.DSR.GrammarFromString strGrammar
    'frmMain.DSR.Activate
End Sub

'Private Sub Dir1_Change()
Public Sub Dir1_Change()
Dim cnt As Integer
Dim pth As String
Dim lst As String
Dim newstr As String
    'MsgBox "Directory Changed"
    List2.Clear
    File1.Path = Dir1.Path
    List1.Clear
    cnt = Dir1.ListCount
    For i = 0 To cnt - 1
        pth = Trim$(Dir1.List(i))
        For j = 1 To Len(pth)
              lst = Mid(pth, j, 1)
            If lst <> "\" Then
                newstr = newstr + lst
            Else
                newstr = ""
            End If
        Next j
            pth = newstr
            List1.AddItem (pth)
    Next i
    
    List2.Clear
    cnt = File1.ListCount
    For i = 0 To cnt - 1
        pth = Trim$(File1.List(i))
        List2.AddItem (pth)
    Next i
End Sub

Private Sub Dir1_Click()
'todo
'MsgBox "Click called"
Dir1_Change
End Sub
Public Sub Dir1_GotFocus()
  Dir1_Change
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    frmMain.g_iFlag = FORMDISKNAV
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.g_iFlag = FORMMAIN
    frmMain.gRestoreGrammar = True
End Sub

Public Sub result()
    Dim cnt As Integer
Dim pth As String
Dim lst As String
Dim newstr As String
    'MsgBox "Directory Changed"
    List2.Clear
    File1.Path = Dir1.Path
    List1.Clear
    cnt = Dir1.ListCount
    For i = 0 To cnt - 1
        pth = Trim$(Dir1.List(i))
        For j = 1 To Len(pth)
              lst = Mid(pth, j, 1)
            If lst <> "\" Then
                newstr = newstr + lst
            Else
                newstr = ""
            End If
        Next j
            pth = newstr
            List1.AddItem (pth)
    Next i
    
    List2.Clear
    cnt = File1.ListCount
    For i = 0 To cnt - 1
        pth = Trim$(File1.List(i))
        List2.AddItem (pth)
    Next i
End Sub

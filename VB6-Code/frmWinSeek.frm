VERSION 5.00
Begin VB.Form WinSeek 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Files"
   ClientHeight    =   4020
   ClientLeft      =   1920
   ClientTop       =   1896
   ClientWidth     =   9048
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000080&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   9048
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   492
      Left            =   2040
      TabIndex        =   13
      Top             =   3240
      Width           =   1212
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   -20000
      ScaleHeight     =   2892
      ScaleWidth      =   8292
      TabIndex        =   2
      Top             =   120
      Width           =   8292
      Begin VB.DriveListBox drvList 
         Height          =   288
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1692
      End
      Begin VB.DirListBox dirList 
         Height          =   1800
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2532
      End
      Begin VB.FileListBox filList 
         Height          =   1800
         Left            =   2880
         TabIndex        =   5
         Top             =   960
         Width           =   3732
      End
      Begin VB.TextBox txtSearchSpec 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   4
         Text            =   "*.exe"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblCriteria 
         Caption         =   "Search &Criteria:"
         Height          =   252
         Left            =   2880
         TabIndex        =   3
         Top             =   600
         Width           =   1332
      End
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   480
      Left            =   600
      TabIndex        =   0
      Top             =   3240
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
      Height          =   468
      Left            =   3480
      TabIndex        =   1
      Top             =   3240
      Width           =   1200
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3132
      Left            =   480
      ScaleHeight     =   3132
      ScaleWidth      =   8772
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   8772
      Begin VB.ListBox lstCommand 
         Height          =   2160
         ItemData        =   "frmWinSeek.frx":0000
         Left            =   120
         List            =   "frmWinSeek.frx":0002
         TabIndex        =   12
         Top             =   480
         Width           =   1452
      End
      Begin VB.ListBox lstFoundFiles 
         Height          =   2160
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   6732
      End
      Begin VB.Label Label2 
         Caption         =   "File Path"
         Height          =   252
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "Commands"
         Height          =   252
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1332
      End
      Begin VB.Label lblCount 
         Caption         =   "0"
         Height          =   252
         Left            =   2760
         TabIndex        =   10
         Top             =   2760
         Width           =   1092
      End
      Begin VB.Label lblfound 
         Caption         =   "&Files Found:"
         Height          =   252
         Left            =   1680
         TabIndex        =   9
         Top             =   2760
         Width           =   1092
      End
   End
End
Attribute VB_Name = "WinSeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SearchFlag As Integer   ' Used as flag for cancel and other operations.
Dim i As Integer

Private Sub cmdExit_Click()
    If cmdExit.Caption = "E&xit" Then
        'End
        frmCmdList.FindFiles App.Path, frmUserCommands.cmbFileName
        For i = 0 To frmUserCommands.cmbFileName.ListCount
            If frmUserCommands.cmbFileName.List(i) = "ExeStdFile" Then
                frmUserCommands.cmbFileName.RemoveItem (i)
                Exit For
            End If
        Next i
        
        frmUserCommands.cmbFileName.ListIndex = 0
        Unload WinSeek
    Else                    ' If user chose Cancel, just end Search.
        SearchFlag = False
    End If
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
Dim item As ListItem
Dim flag As Boolean 'to check weather any commands are added or not
flag = False
    
    For i = 0 To lstFoundFiles.ListCount
        If lstCommand.List(i) <> "" Then
            flag = True
            Set item = frmUserCommands.ListView1.ListItems.Add(1, , lstCommand.List(i))
            item.SubItems(frmUserCommands.ListView1.ColumnHeaders("FilePath").SubItemIndex) = lstFoundFiles.List(i)
        End If
        
    Next i
    
    If flag Then
        'save flag=false because we are adding listitems
        'Otherwise don't change flag
        frmUserCommands.gSave_Flag = False
    'Else
        'Nothing is added into the listview
    '    frmUserCommands.gSave_Flag = True
    End If
    
    Unload Me 'WinSeek
End Sub

Private Sub cmdSearch_Click()
' Initialize for search, then perform recursive search.
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  ' Check what the user did last.
    'cmdOk.Enabled = True
    If cmdSearch.Caption = "&Reset" Then  ' If just a reset, initialize and exit.
        ResetSearch
        lstCommand.Clear '******To clear Command list
        'txtSearchSpec.SetFocus
        Exit Sub
    End If

    ' Update dirList.Path if it is different from the currently
    ' selected directory, otherwise perform the search.
    If dirList.Path <> dirList.List(dirList.ListIndex) Then
        dirList.Path = dirList.List(dirList.ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If

    ' Continue with the search.
    Picture2.Move 0, 0
    Picture1.Visible = False
    Picture2.Visible = True

    cmdExit.Caption = "Cancel"

    filList.Pattern = txtSearchSpec.Text
    FirstPath = dirList.Path
    DirCount = dirList.ListCount

    ' Start recursive direcory search.
    NumFiles = 0                       ' Reset found files indicator.
    result = DirDiver(FirstPath, DirCount, "")
    filList.Path = dirList.Path
    cmdSearch.Caption = "&Reset"
    cmdSearch.SetFocus
    cmdExit.Caption = "E&xit"
    Dim i As Integer
    For i = 0 To lstFoundFiles.ListCount - 1
        lstCommand.AddItem ""
    Next
End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiver = False            ' Set to True if there is an error.
    retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = dirList.ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = dirList.Path                      ' Save old path for next recursion.
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
            ' Get to the node bottom.
            dirList.Path = dirList.List(DirsToPeek - 1)
            AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = dirList.Path                  ' If at root level, leave as is...
        Else
            ThePath = dirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For ind = 0 To filList.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath + filList.List(ind)
            lstFoundFiles.AddItem entry
            lblCount.Caption = Str(Val(lblCount.Caption) + 1)
        Next ind
    End If
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        dirList.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiver = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else                        ' Otherwise display error message and quit.
        MsgBox Error
        End
    End If
End Function

Private Sub DirList_Change()
    ' Update the file list box to synchronize with the directory list box.
    filList.Path = dirList.Path
End Sub

Private Sub DirList_LostFocus()
    dirList.Path = dirList.List(dirList.ListIndex)
End Sub

Private Sub DrvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub


Private Sub Form_Load()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    
    Picture2.Move 0, 0
    Picture2.Width = WinSeek.ScaleWidth
    Picture2.BackColor = WinSeek.BackColor
    lblCount.BackColor = WinSeek.BackColor
    lblCriteria.BackColor = WinSeek.BackColor
    lblfound.BackColor = WinSeek.BackColor
    Picture1.Move 0, 0
    Picture1.Width = WinSeek.ScaleWidth
    Picture1.BackColor = WinSeek.BackColor
    cmdOk.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'End
End Sub

Private Sub ResetSearch()
    ' Reinitialize before starting a new search.
    lstFoundFiles.Clear
    lblCount.Caption = 0
    SearchFlag = False                  ' Flag indicating search in progress.
    Picture2.Visible = False
    cmdSearch.Caption = "&Search"
    cmdExit.Caption = "E&xit"
    Picture1.Visible = True
    dirList.Path = CurDir: drvList.Drive = dirList.Path ' Reset the path.
End Sub

Private Sub lstCommand_Click()
'Dim command As String

    'lstFoundFiles.ListIndex = lstCommand.ListIndex
    'command = InputBox("Enter the Voice Command for File:", "New Command")
    'If Trim(command) <> Trim("") Then
    '    lstCommand.List(lstCommand.ListIndex) = command
    '    cmdOk.Enabled = True
    'End If
    
End Sub

'Private Sub lstCommand_DblClick()
'Dim command As String
'    command = InputBox("Enter the Voice Command for File:", "New Command")
'    If Trim(command) <> Trim("") Then
'        lstCommand.List(lstCommand.ListIndex) = command
'    End If
'End Sub

Private Sub lstFoundFiles_Click()
Dim command As String
    lstCommand.ListIndex = lstFoundFiles.ListIndex
    command = InputBox("Enter the Voice Command for File:", "New Command")
    If Trim(command) <> "" Then
        lstCommand.List(lstFoundFiles.ListIndex) = command
        cmdOk.Enabled = True
    Else
        lstCommand.List(lstFoundFiles.ListIndex) = ""
    End If
End Sub

'Private Sub lstFoundFiles_DblClick()
'Dim command As String
'    command = InputBox("Enter the Voice Command for File:", "New Command")
'    If Trim(command) <> Trim("") Then
'        lstCommand.List(lstFoundFiles.ListIndex) = command
'    End If
'End Sub

Private Sub txtSearchSpec_Change()
    ' Update file list box if user changes pattern.
    filList.Pattern = txtSearchSpec.Text
End Sub

Private Sub txtSearchSpec_GotFocus()
    txtSearchSpec.SelStart = 0          ' Highlight the current entry.
    txtSearchSpec.SelLength = Len(txtSearchSpec.Text)
End Sub


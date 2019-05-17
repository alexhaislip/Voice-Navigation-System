VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCmdList 
   Caption         =   "Command List"
   ClientHeight    =   4656
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7944
   LinkTopic       =   "Form1"
   ScaleHeight     =   4656
   ScaleWidth      =   7944
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   240
      Picture         =   "frmCmdList1.frx":0000
      ScaleHeight     =   252
      ScaleWidth      =   252
      TabIndex        =   6
      Top             =   240
      Width           =   252
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Grammer"
      Height          =   372
      Left            =   6480
      TabIndex        =   5
      Top             =   2280
      Width           =   1332
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   372
      Left            =   6480
      TabIndex        =   4
      Top             =   1800
      Width           =   1332
   End
   Begin VB.CommandButton cmdDefAll 
      Caption         =   "Default All"
      Height          =   372
      Left            =   6480
      TabIndex        =   3
      Top             =   1320
      Width           =   1332
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default Name"
      Height          =   372
      Left            =   6480
      TabIndex        =   2
      Top             =   840
      Width           =   1332
   End
   Begin VB.ComboBox cmbFileNames 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   2172
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3492
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6012
      _ExtentX        =   10605
      _ExtentY        =   6160
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Commands"
         Text            =   "Commands"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "FilePath"
         Text            =   "File Name"
         Object.Width           =   9596
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Command List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   372
      Left            =   4440
      TabIndex        =   7
      Top             =   240
      Width           =   1812
   End
End
Attribute VB_Name = "frmCmdList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prevStr As String
Dim currStr As String
Dim m_inAttr As Integer
'Dim Not_Reverse As Boolean

Private Sub cmbFileNames_Click()
    frmUserCommands.LoadCFGFile cmbFileNames.Text + ".def", ListView1
End Sub

Private Sub cmdDefAll_Click()
    frmUserCommands.LoadCFGFile cmbFileNames.Text + ".cfg", ListView1
    
    'write the entire cfg file from list and appends the general cfg
    WriteGenCFG2DefCFG cmbFileNames.Text
End Sub

Private Sub cmdDefault_Click()
Dim FileNo As Integer
Dim cnt As Integer
Dim Index As Integer
Dim item As ListItem
Dim realCommand As String
Dim getrec() As String
    FileNo = FreeFile

    Open App.Path + "\" + cmbFileNames.Text + ".cfg" For Input As #FileNo
    cnt = 0
    'Retrive file into dynamic array "getrec"
    
    Do While Not EOF(FileNo)
        ReDim Preserve getrec(cnt)
        Line Input #FileNo, getrec(cnt)
        cnt = cnt + 1
    Loop
    Close FileNo
    
    'index is the index of selected item
    Index = ListView1.SelectedItem.Index
    
    'item contains the string from first column of the respective index
    Set item = ListView1.ListItems.item(Index)
    
    'realCommand is the string from second column of same index(File Name)
    realCommand = item.SubItems(ListView1.ColumnHeaders("FilePath").SubItemIndex)
    
    'Compare the realCommand with currString string, if same replace the item(index) in listview
    For i = 5 To UBound(getrec)
        findString (getrec(i))
        If currStr = realCommand Then
            ListView1.ListItems.item(Index) = prevStr
            WriteGenCFG2DefCFG cmbFileNames.Text
            Exit For
        End If
    Next i
End Sub

Private Sub cmdOK_Click()
    
    If frmMain.g_iFlag = FORMCOMMANDLIST Then
        frmMain.g_iFlag = FORMMAIN 'Main frame is activated
    End If
    
    WriteGenCFG2DefCFG cmbFileNames.Text
    cmdSelect_Click
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    
    frmMain.gGrammarFile = Trim(cmbFileNames.Text)
    MsgBox frmMain.gGrammarFile
    lonStatus = WritePrivateProfileString("TTS", "Grammar File", frmMain.gGrammarFile, App.Path + "\Vns.ini")
    frmMain.DSR.GrammarFromFile frmMain.gGrammarFile
End Sub

Private Sub Form_Load()
    g_iFlag = FORMCOMMANDLIST

    'cmbFileNames.Text = frmMain.gGrammarFile
    m_inAttr = vbNormal + vbDirectory
    
    '***** Finds all files in current directory having .def extension
    FindFiles App.Path, cmbFileNames
    
    'Remove Game CFG file from combo box
     For i = 0 To cmbFileNames.ListCount
        If (cmbFileNames.List(i) = "EngLetters") Then
            cmbFileNames.RemoveItem (i) 'It is removing the item but leaving the affect "Has to be Seen"
        End If
    Next i
    
    '*****Extract phrases and parsed from default cfg file
    '*****and add this into the respective list
    lonStatus = GetPrivateProfileString("TTS", "Grammar File", "?!?", gStrFixed, 100, App.Path + "\Vns.ini")
    frmMain.gGrammarFile = Left$(gStrFixed, lonStatus)
    
    frmUserCommands.LoadCFGFile frmMain.gGrammarFile + ".def", ListView1
    
End Sub
'finds the files with .def extension in current directory
Public Sub FindFiles(ByVal inPath As String, Combo As ComboBox)
Dim strFileName As String
Dim strSearchPath As String
Dim strFileSpec As String
Dim Index As Integer
Dim objectCount As Integer

    cmbFileNames.Clear
    strFileSpec = "*.def"
    strSearchPath = inPath
    If Right(strSearchPath, 1) = "\" Then
        strSearchPath = Left(strSearchPath, 2)
    End If
    strAttributePath = strSearchPath
    strSearchPath = strSearchPath + "\" + strFileSpec
    strFileName = Dir(strSearchPath, m_intAtts)
    'index = 0
    objectCount = 0
    Do While (strFileName <> "")
        objectCount = objectCount + 1
        If strFileName <> "." And strFileName <> ".." Then
                strFileName = ExtractFileName(strFileName)
                If frmMain.gGrammarFile = strFileName Then
                    Index = objectCount - 1
                End If
                Combo.AddItem (strFileName)
        End If
        strFileName = Dir
    Loop
    Combo.ListIndex = Index
    
End Sub

'Extracts the file names without extension
Public Function ExtractFileName(str1 As String) As String
Dim temp As String
Dim c As String * 1
    For i = 1 To Len(str1)
        c = Mid(str1, i, 1)
        If c <> "." Then
            temp = temp & c
        Else
            Exit For
        End If
    Next i
    ExtractFileName = temp
End Function

'Separate the phrase from parsed from the record of a file
Sub findString(str1 As String)
Dim temp1 As String
Dim temp2 As String
Dim c1 As String * 1

    For i = 1 To Len(str1)
        c1 = Mid(str1, i, 1)
        If c1 <> "=" Then
            temp1 = temp1 + c1
        Else
            c1 = ""
            temp1 = ""
        End If
    Next i
    
    For i = 1 To Len(temp1) - 1
       c1 = Mid(temp1, i, 1)
        If c1 = "." Then
            Exit For
        ElseIf c1 <> """" Then
            temp2 = temp2 + c1
        ElseIf c1 = """" Then
            c1 = ""
            prevStr = temp2
            temp2 = ""
        End If
   currStr = temp2
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.g_iFlag = FORMMAIN
    'MsgBox frmMain.g_iFlag
End Sub


Private Sub WriteGenCFG2DefCFG(fileName As String)
Dim FileNo As Integer
Dim getrec() As String
Dim item As ListItem
Dim str1 As String
    FileNo = FreeFile
    
    '***** Writing the default cfg from list
    Open App.Path + "\" + fileName + ".def" For Output As #FileNo
    Print #FileNo, "[Grammar]"
    Print #FileNo, "langid=1033"
    Print #FileNo, "type=cfg"
    Print #FileNo, ""
    Print #FileNo, "[<start>]"
    For i = 1 To ListView1.ListItems.Count  'List1.ListCount - 1
        Set item = ListView1.ListItems.item(i)
        str1 = item.SubItems(ListView1.ColumnHeaders("FilePath").SubItemIndex)
        'Print #FileNo, "<start>=" & Trim(List1.List(i)) & " " & """" & Trim(List2.List(i)) & ".exe" & """"
        Print #FileNo, "<start>=" & Trim(item) & " " & """" & Trim(str1) & ".exe" & """"
    Next i
    Close FileNo
    
    '*****Writing general command to the default cfg file
    FileNo = FreeFile
    Open App.Path + "\GeneralCFG.cfg" For Input As #FileNo
    
    cnt = 0
    Do While Not EOF(FileNo)
        ReDim Preserve getrec(cnt)
        Line Input #FileNo, getrec(cnt)
        cnt = cnt + 1
    Loop
    Close FileNo
    FileNo = FreeFile
    
    Open App.Path + "\" + fileName + ".def" For Append As #FileNo
        Print #FileNo, ""
        For i = 0 To cnt - 1
            Print #FileNo, getrec(i)
        Next i
    Close FileNo
    
End Sub
Private Function ListCompare(EditCommand As String) As Boolean
Dim flag As Boolean

    
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = EditCommand Then
            flag = True
            Exit For
        End If
    Next i
    
    If flag Then
        ListCompare = True
    Else
        ListCompare = False
    End If
End Function

'Working
Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    ListView1.ListItems.item(ListView1.SelectedItem.Index) = NewString
    WriteGenCFG2DefCFG cmbFileNames.Text

End Sub


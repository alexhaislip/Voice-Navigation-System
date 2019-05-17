VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserCommands 
   Caption         =   "User Commands"
   ClientHeight    =   4788
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   8208
   LinkTopic       =   "Form1"
   ScaleHeight     =   4788
   ScaleWidth      =   8208
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   3360
      Picture         =   "frmUserCommands.frx":0000
      ScaleHeight     =   612
      ScaleWidth      =   1212
      TabIndex        =   10
      Top             =   120
      Width           =   1212
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   372
      Left            =   6720
      TabIndex        =   9
      Top             =   2400
      Width           =   1212
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New Grammer"
      Height          =   372
      Left            =   6720
      TabIndex        =   8
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   372
      Left            =   6720
      TabIndex        =   7
      Top             =   1440
      Width           =   1212
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   372
      Left            =   6720
      TabIndex        =   6
      Top             =   1920
      Width           =   1212
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   372
      Left            =   6720
      TabIndex        =   5
      Top             =   2880
      Width           =   1212
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   372
      Left            =   6720
      TabIndex        =   4
      Top             =   3360
      Width           =   1212
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   372
      Left            =   6720
      TabIndex        =   3
      Top             =   3840
      Width           =   1212
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3612
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   6252
      _ExtentX        =   11028
      _ExtentY        =   6371
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Text            =   "File Path"
         Object.Width           =   9596
      EndProperty
   End
   Begin VB.ComboBox cmbFileName 
      Height          =   288
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label Label5 
      Caption         =   "Old Grammer:"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1332
   End
End
Attribute VB_Name = "frmUserCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gSave_Flag As Boolean 'if False Save the save.
Dim currStr As String
Dim prevStr As String

Private Sub cmbFileName_Click()
    LoadCFGFile cmbFileName.Text & ".def", ListView1
End Sub

Private Sub cmdClear_Click()
    ListView1.ListItems.Clear
End Sub

Private Sub cmdDelete_Click()
Dim ToDel As Integer
Dim Index As Integer

    Index = ListView1.SelectedItem.Index
    'MsgBox Str(Index)
    ans = MsgBox("Do You want to Delete " + ListView1.SelectedItem + "? ", vbYesNo + vbExclamation, "Alert")
    If ans = vbYes Then
        gSave_Flag = False
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    End If
'Exit Sub

'ToDel = ListView1.ListItems(ListView1.SelectedItem(Index))
'MsgBox ToDel
'    MsgBox ListView1.SelectedItem
'    ListView1.SelectedItem = "abc"
    'ListView1.l
End Sub

Private Sub cmdNew_Click()
    'gSave_Flag = False
    'cmbFileName.Enabled = False
    cmbFileName.Clear
    ListView1.ListItems.Clear
    WinSeek.Show vbModal
End Sub

Private Sub cmdOK_Click()
    If Not gSave_Flag Then 'if file is not saved
        ans = MsgBox("File is not saved.Do you want to save it", vbYesNo + vbExclamation, "Confirm...")
        If ans = vbNo Then
            gSave_Flag = True 'set the flag to save
            Unload Me
            Exit Sub
        Else
            If Trim(cmbFileName.Text) = "" Then 'if file name is not given, prompt
               ' MsgBox "File name is not specified", vbInformation, "Specify file name..."
               'if new file is created
               cmdSave_Click
            Else
                'gSave_Flag= True
                'Existing file is updated
                SaveCFGFile cmbFileName.Text
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim cfgName As String
Dim FileExist As Boolean
    If gSave_Flag = False Then
        'If Trim(txtName.Text) = "" Then
        '    MsgBox "File name is not specified", vbInformation, "Specify file name..."
        'Else
        '    'gSave_Flag = True
        '    SaveCFGFile txtName.Text
        'End If
        If Trim(cmbFileName.Text) <> "" Then
            SaveCFGFile cmbFileName.Text
        Else
label:
            cfgName = InputBox("Enter the File name:", "Save as")
            If Trim(cfgName) <> "" Then
                FileExist = ComboCompare(cfgName)
                If FileExist = True Then
                'MsgBox FileExist
                    ans = MsgBox("File already exist.Do you want to replace it?", vbYesNo + vbExclamation, "Alert")
                    If ans = vbNo Then
                        'Input box appears
                        GoTo label
                    Else
                    'If ans = vbsyes Then
                    '    'Overwrites the existing file
                        SaveCFGFile cfgName
                    '    MsgBox "File is Saved"
                    'Else
                    '    MsgBox "File Not Saved"
                    End If
                 Else
                    'File Not exists
                    SaveCFGFile cfgName
                End If
                ''cmbFileName.AddItem (cfgName)
            End If
        frmCmdList.FindFiles App.Path, cmbFileName
        cmbFileName.ListIndex = cmbFileName.ListCount - 1
        End If
    End If
End Sub

Private Sub cmdUpdate_Click()
    frmAddCmd.txtCFGName.Text = cmbFileName.Text
    frmAddCmd.Show
End Sub

Private Sub Form_Load()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    
    'save is true since items in listview is not altered
    gSave_Flag = True
   
    
    'search the files and add them into combo Box
    frmCmdList.FindFiles App.Path, cmbFileName
    
    For i = 0 To cmbFileName.ListCount
        If (cmbFileName.List(i) = "ExeStdFile") Then
            cmbFileName.RemoveItem (i) 'It is removing the item but leaving the affect "Has to be Seen"
        End If
        If (cmbFileName.List(i) = "EngLetters") Then
            cmbFileName.RemoveItem (i) 'It is removing the item but leaving the affect "Has to be Seen"
        End If
    Next i
    'cmbFileName Click event occur
     frmUserCommands.cmbFileName.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If gSave_Flag = False Then cmdOK_Click
    
    frmMain.mnuUtilsUser.Caption = "&User Commands"
    
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    gSave_Flag = False
End Sub

Public Sub SaveCFGFile(fileName As String)
Dim FileNo As String
Dim item As ListItem
Dim str1 As String
Dim i As Integer

    FileNo = FreeFile
    gSave_Flag = True
    Open App.Path & "\" & fileName & ".cfg" For Output As #FileNo
        Print #FileNo, "[Grammar]"
        Print #FileNo, "langid=1033"
        Print #FileNo, "type=cfg"
        Print #FileNo, ""
        Print #FileNo, "[<start>]"
        'Print #FileNo, ""
        For i = 1 To ListView1.ListItems.Count
        
            'MsgBox ListView1.ListItems.item(i)
            Set item = ListView1.ListItems.item(i)
            str1 = item.SubItems(frmUserCommands.ListView1.ColumnHeaders("FilePath").SubItemIndex) '= lstFoundFiles.List(i)
            str1 = ReplaceSlash(str1)
            Print #FileNo, "<start>=" & Trim(item) & " " & """" & Trim(str1) & """"
            
        Next i
        
        'FileCopy fileName & ".cfg", fileName & ".def"
    Close FileNo
    FileNo = FreeFile
    'Make the default file
    FileCopy App.Path & "\" & fileName & ".cfg", App.Path & "\" & fileName & ".def"
    
End Sub

Public Sub LoadCFGFile(fileName As String, ListView1 As ListView)
    
Dim FileNo As Integer
Dim cnt As Integer
Dim getrec() As String
Dim item As ListItem
    FileNo = FreeFile
    
    ListView1.ListItems.Clear
    Open App.Path + "\" + fileName For Input As #FileNo
    cnt = 0
    Do While Not EOF(FileNo)
        ReDim Preserve getrec(cnt)
        Line Input #FileNo, getrec(cnt)
        cnt = cnt + 1
    Loop
    Close FileNo
    'For i = UBound(getrec) - 9 To 5 Step -1
    For i = 5 To UBound(getrec)
        If getrec(i) = "" Then
            Exit For
        End If
        findString (getrec(i))
                
        Set item = ListView1.ListItems.Add(1, , prevStr)
        item.SubItems(ListView1.ColumnHeaders("FilePath").SubItemIndex) = currStr
    Next i
End Sub
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
'Checks for the distinct file names
Private Function ComboCompare(newFileName As String) As Boolean
Dim flag As Boolean
    
    For i = 0 To cmbFileName.ListCount - 1
        If cmbFileName.List(i) = newFileName Then
            flag = True
            Exit For
        End If
    Next i
    
    If flag Then
        ComboCompare = True
    Else
        ComboCompare = False
    End If
End Function

Private Function ReplaceSlash(InPath As String) As String
    Dim str1 As String
    Dim c As String * 1
    For i = 1 To Len(InPath)
        c = Mid(InPath, i, 1)
        If c <> "\" Then
            str1 = str1 & c
        Else
            str1 = str1 & "\" & c
        End If
    Next i
    ReplaceSlash = str1
End Function


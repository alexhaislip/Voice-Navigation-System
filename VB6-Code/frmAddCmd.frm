VERSION 5.00
Begin VB.Form frmAddCmd 
   Caption         =   "Add Commands"
   ClientHeight    =   2724
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   5544
   LinkTopic       =   "Form1"
   ScaleHeight     =   2724
   ScaleWidth      =   5544
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   3120
      TabIndex        =   8
      Top             =   2040
      Width           =   1212
   End
   Begin VB.TextBox txtCFGName 
      Enabled         =   0   'False
      Height          =   372
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   2052
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   372
      Left            =   1680
      TabIndex        =   6
      Top             =   2040
      Width           =   1212
   End
   Begin VB.TextBox txtPath 
      Height          =   372
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   2052
   End
   Begin VB.TextBox txtCommand 
      Height          =   372
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   2052
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   372
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "CFG File Name:"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "File Path:"
      Height          =   252
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Command:"
      Height          =   252
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   972
   End
End
Attribute VB_Name = "frmAddCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
'Dim fileno As Integer
'Dim Filename As String
    'fileno = FreeFile
    'Filename = "User_Office"
    'Open App.Path + "\" + Filename + ".cfg" For Append As #fileno
    '    Print #fileno, "<start>=" + txtCommand + " " + """" + txtPath + """"
    'Close fileno
    'fileno = FreeFile
    'Open App.Path + "\" + Filename + ".def" For Append As #fileno
    '    Print #fileno, "<start>=" + txtCommand + " " + """" + txtPath + """"
    'Close fileno
    'txtCommand.Text = ""
    'txtPath.Text = ""
Dim item As ListItem
    If Trim(txtCommand.Text) <> "" And Trim(txtPath.Text) <> "" Then
        Set item = frmUserCommands.ListView1.ListItems.Add(1, , txtCommand.Text)
        item.SubItems(frmUserCommands.ListView1.ColumnHeaders("FilePath").SubItemIndex) = txtPath.Text
        ' If added into the listview make flag false
        'otherwise don't change
        frmUserCommands.gSave_Flag = False
        
    'Else
        '
    '    frmUserCommands.gSave_Flag = True
    End If
    
    Unload Me
End Sub

Private Sub cmdBrowse_Click()
    Dim sFile As String

    With frmMain.dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "EXE Files (*.exe)|*.exe"
        .ShowOpen
        If Len(.fileName) = 0 Then
            Exit Sub
        End If
        sFile = .fileName
    End With
    'frmEleSpk.rtfText.LoadFile sFile
    'frmEleSpk.Caption = "Voice Navigation System [" & sFile & "]"
    txtPath.Text = sFile
End Sub

Private Sub Command1_Click()
Dim FileNo As Integer
Dim fileName As String
    FileNo = FreeFile
    fileName = "User_Office"
    Open App.Path + "\" + fileName + ".cfg" For Append As #FileNo
        Print #FileNo, "<start>=" + txtCommand + " " + """" + txtPath + """"
    Close FileNo
    FileNo = FreeFile
    Open App.Path + "\" + fileName + ".def" For Append As #FileNo
        Print #FileNo, "<start>=" + txtCommand + " " + """" + txtPath + """"
    Close FileNo
    txtCommand.Text = ""
    txtPath.Text = ""
    
End Sub

Private Sub cmdCancel_Click()
    'If frmUserCommands.gSave_Flag Then
    '    frmUserCommands.gSave_Flag = True
    Unload Me
End Sub

Private Sub Form_Load()

    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    
End Sub

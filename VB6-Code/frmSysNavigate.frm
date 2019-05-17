VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysNavigate 
   Caption         =   "System Navigator"
   ClientHeight    =   5184
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   6492
   LinkTopic       =   "Form1"
   ScaleHeight     =   5184
   ScaleWidth      =   6492
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "List Options"
      Height          =   1572
      Left            =   2520
      TabIndex        =   8
      Top             =   3480
      Width           =   3852
      Begin VB.Frame Frame3 
         Caption         =   "Arrangement"
         Height          =   1212
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1812
         Begin VB.OptionButton rbAlignTop 
            Caption         =   "Align Top"
            Height          =   252
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1572
         End
         Begin VB.OptionButton rbAlignLeft 
            Caption         =   "Align Left"
            Height          =   252
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1452
         End
         Begin VB.OptionButton rbNoArrange 
            Caption         =   "No Arrange"
            Height          =   252
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1572
         End
      End
      Begin MSComctlLib.ImageList imlSmallIcons 
         Left            =   2520
         Top             =   1200
         _ExtentX        =   804
         _ExtentY        =   804
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlIcons 
         Left            =   2040
         Top             =   1200
         _ExtentX        =   804
         _ExtentY        =   804
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show These Types of Files"
      Height          =   1572
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   2412
      Begin VB.TextBox txtFileSpec 
         Height          =   288
         Left            =   1080
         TabIndex        =   6
         Text            =   "*.*"
         Top             =   1080
         Width           =   1212
      End
      Begin VB.CheckBox cbArchive 
         Caption         =   "Archive Files"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1212
      End
      Begin VB.CheckBox cbHidden 
         Caption         =   "Hidden Files"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1212
      End
      Begin VB.CheckBox cbSystem 
         Caption         =   "System Files"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "File Spec"
         Height          =   252
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   852
      End
   End
   Begin VB.DirListBox dirListView 
      Height          =   2664
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2052
   End
   Begin VB.DriveListBox drvListView 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2052
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2652
      Left            =   2400
      TabIndex        =   13
      Top             =   480
      Width           =   3972
      _ExtentX        =   7006
      _ExtentY        =   4678
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblObjectCount 
      Height          =   252
      Left            =   2400
      TabIndex        =   17
      Top             =   3120
      Width           =   1932
   End
   Begin VB.Label lblContents 
      BorderStyle     =   1  'Fixed Single
      Height          =   372
      Left            =   2400
      TabIndex        =   16
      Top             =   0
      Width           =   3972
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   372
      Left            =   2760
      TabIndex        =   15
      Top             =   2280
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   372
      Left            =   2760
      TabIndex        =   14
      Top             =   2280
      Width           =   972
   End
End
Attribute VB_Name = "frmSysNavigate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const tagStart = vbNewLine + "<start>="
Public strGrammar As String

'Stores total attributes for types of files displayed
Dim m_intAtts As Integer
'
'Stores ListView mode
Dim m_intViewType As Integer
'
'Flag indicating whether combo box has been populated
Dim m_blComboPop As Boolean
'
'Flag indicating whether columns have been set for Report mode
Dim m_blnFlgColumnsSet As Boolean
'
'Constants for use with the ListImage control
Const SMALL_FOLDER = "C:\Program Files\Microsoft Visual Studio\Common\Graphics\Bitmaps\Outline\closed.bmp"
Const SMALL_FILE = "C:\Program Files\Microsoft Visual Studio\Common\Graphics\Bitmaps\Outline\leaf.bmp"
Const LARGE_FOLDER = "C:\Program Files\Microsoft Visual Studio\Common\Graphics\Icons\Win95\clsdfold.ico"
Const LARGE_FILE = "C:\Program Files\Microsoft Visual Studio\Common\Graphics\Icons\dragdrop\drag1pg.ico"

Private Sub cbArchive_Click()
    '
    ProcessFileAtts
    PopulateList ""
    '
End Sub

Private Sub cbHidden_Click()
    '
    ProcessFileAtts
    PopulateList ""
    '
End Sub

Private Sub cbSystem_Click()
    '
    ProcessFileAtts
    PopulateList ""
    '
End Sub

Private Sub dirListView_Change()
'MsgBox "ch"
    
    '
    'frmMain.DSR.GrammarFromString strGrammar
    'Initalize strGrammar at directory change
    strGrammar = "[Grammer]" & vbNewLine & "langid = 1033" & vbNewLine & "type=cfg" & vbNewLine & vbNewLine & "[<start>]"
    ExtractParentDir dirListView.Path
    PopulateList ""
    frmMain.DSR.GrammarFromString strGrammar
    'To activate DSR******not activating
    'Speech2User "Directory is changed", 4
    
    frmMain.DSR.Activate
    '
End Sub

Private Sub dirListView_Click()
    
    'strSearchPath = dirListView.List(dirListView.ListIndex)
    
    'Change the path to selected directory it invokes dir_change event
    dirListView.Path = dirListView.List(dirListView.ListIndex)
End Sub

Private Sub drvListView_Change()
    '
    'Synchronize drive and directory controls
    'Also, set keyboard focus to directory list immediately after drive change
    dirListView.Path = drvListView.Drive
    dirListView.SetFocus
    '
End Sub

Private Sub PopulateList(ByVal InPath As String)

    Dim lvListItem As ListItem
    Dim strFileName As String
    Dim strSearchPath As String
    Dim strFileSpec As String
    Dim strItemType As String
    Dim intObjectCount As Integer
    '
    'Empty items for the View control
    ListView1.ListItems.Clear
    '
    'Get the file spec value for the form
    strFileSpec = txtFileSpec.Text
    '
    If InPath <> "" Then
        strSearchPath = InPath
    Else
        strSearchPath = dirListView.Path
    End If
    

    '
    'Set the 'Content of' caption
    lblContents.Caption = "Contents of " & strSearchPath
    '
    'Remove trailing backslash in case current directory is
    'root directory (in which case the \ is added by the O/S)
    If Right(strSearchPath, 1) = "\" Then
        strSearchPath = Left(strSearchPath, 2)
    End If
    '
    'Store the path so we can locate the file later in order
    'to retrieve its attributes
    strAttributePath = strSearchPath
    '
    'Add the file spec to the search path
    strSearchPath = strSearchPath + "\" + strFileSpec
    '
    'Retrieve the first object from the directory using
    strFileName = Dir(strSearchPath, m_intAtts)
    '
    Do While (strFileName <> "")
        intObjectCount = intObjectCount + 1
        If strFileName <> "." And strFileName <> ".." Then
            Set lvListItem = ListView1.ListItems.Add(, , strFileName)
            '
            'This logic will trap all non-directory items found. This saves me for
            'testing for all of the attributes associated with a file
            If (GetAttr(strAttributePath + "\" + strFileName) <> vbDirectory) Then
                lvListItem.SmallIcon = "File"
                lvListItem.Icon = "File"
                strItemType = "File"
            Else
                lvListItem.SmallIcon = "Folder"
                lvListItem.Icon = "Folder"
                strItemType = "File Folder"
            End If
            
            'If ListView1.View = lvwReport Then
            '    lvListItem.SubItems(ListView1.ColumnHeaders("Type").SubItemIndex) = strItemType
            '   '
            '    If strItemType = "File" Then
            '        lvListItem.SubItems(ListView1.ColumnHeaders("Size").SubItemIndex) = Str$(FileLen(strAttributePath + "\" + strFileName))
            '    Else
            '        lvListItem.SubItems(ListView1.ColumnHeaders("Size").SubItemIndex) = " "
            '    End If
            '    lvListItem.SubItems(ListView1.ColumnHeaders("Date").SubItemIndex) = Str$(FileDateTime(strAttributePath + "\" + strFileName))
            'End If
            
        End If
        '
        If strFileName <> "." And strFileName <> ".." Then
                strFileName = frmCmdList.ExtractFileName(strFileName)
                StringGrammar strFileName '*****Changes
        End If
        strFileName = Dir
    '
    Loop
    'MsgBox strGrammar
    If intObjectCount = 1 Then
        lblObjectCount.Caption = Str(intObjectCount) + " object found."
    Else
        lblObjectCount.Caption = Str(intObjectCount) + " objects found."
    End If
    '
End Sub

Private Sub Form_Load()
     '
    frmMain.g_iFlag = FORMDISKNAV
    frmMain.gRestoreGrammar = False
    'MsgBox frmMain.g_iFlag
    'Speech2User "Exlorer is opened", 4
    
    'MsgBox Str(gRestoreGrammar)
    
    'frmMain.DSR.Activate
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    
    ListView1.View = lvwList
    
    InitImageList
    m_intAtts = vbNormal + vbDirectory
    
    strGrammar = "[Grammer]" & vbNewLine & "langid = 1033" & vbNewLine & "type=cfg" & vbNewLine & vbNewLine & "[<start>]"
    
    ExtractParentDir (dirListView.Path) 'Get the name of parent directories
    
    PopulateList ""
    
    
    'frmMain.DSR.Activate
    ListView1.SortOrder = lvwAscending
    '
End Sub
Private Sub InitImageList()
    '
    'Specifies images used for the 2 Imagelist controls used in the project
    Dim liListView As ListImage
    Set liListView = imlSmallIcons.ListImages.Add(, "File", LoadPicture(SMALL_FILE))
    Set liListView = imlSmallIcons.ListImages.Add(, "Folder", LoadPicture(SMALL_FOLDER))
    '
    Set liListView = imlIcons.ListImages.Add(, "File", LoadPicture(LARGE_FILE))
    Set liListView = imlIcons.ListImages.Add(, "Folder", LoadPicture(LARGE_FOLDER))
    '
    ListView1.Icons = imlIcons
    ListView1.SmallIcons = imlSmallIcons
    '
End Sub

Private Sub ProcessFileAtts()
    '
    m_intAtts = vbNormal + vbDirectory
    '
    If cbSystem.Value = 1 Then
        m_intAtts = m_intAtts + vbSystem
    End If
    '
    If cbArchive.Value = 1 Then
        m_intAtts = m_intAtts + vbArchive
    End If
    '
    If cbHidden.Value = 1 Then
        m_intAtts = m_intAtts + vbHidden
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.g_iFlag = FORMMAIN
    frmMain.gRestoreGrammar = True
End Sub

Public Sub ListView1_Click()
Dim InPath As String
Dim filename As String
    
    ' if C:\ or D:\ then remove "\"
    If Right(dirListView.Path, 1) = "\" Then
        InPath = Left(dirListView.Path, 2)
    Else
        InPath = dirListView.Path
    End If
    'MsgBox "Path=" & InPath & "\" & ListView1.ListItems.item(ListView1.SelectedItem.Index)
    InPath = InPath & "\" & ListView1.ListItems.item(ListView1.SelectedItem.index)
    If GetAttr(InPath) = vbDirectory Then
        dirListView.Path = InPath
    Else
         MsgBox "File"
    End If
    
End Sub

'Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'    '
'    If ListView1.SortOrder = lvwAscending Then
'        ListView1.SortOrder = lvwDescending
'    Else
'        ListView1.SortOrder = lvwAscending
'    End If
'    '
'    ListView1.SortKey = ColumnHeader.Index - 1
'    ListView1.Sorted = True
'    '
'End Sub

Private Sub rbNoArrange_Click()
    '
    ListView1.Arrange = 0
    PopulateList ""
    '
End Sub

Private Sub rbAlignLeft_Click()
    '
    ListView1.Arrange = 1
    PopulateList "" '(m_CurrentListTable)
    '
End Sub

Private Sub rbAlignTop_Click()
    '
    ListView1.Arrange = 2
    PopulateList "" '(m_CurrentListTable)
    '
End Sub

Private Sub txtFileSpec_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = vbKeyReturn Then
        PopulateList ""
    End If
    '
End Sub

Private Sub StringGrammar(str1 As String)
    
    strGrammar = strGrammar & tagStart & Trim(str1) & " " & """" & Trim(str1) & """"
    
End Sub

Private Sub ExtractParentDir(InPath As String)
'Get parent dir names to create cfg string
'inorder to move from sub dir to parent dir direction
    Dim c As String * 1
    Dim subDir As String
    
    If Right$(InPath, 1) <> "\" Then
        InPath = InPath & "\"
    End If
    For i = 0 To drvListView.ListCount - 1
        strGrammar = strGrammar & tagStart & Trim(Left$(drvListView.List(i), 1)) & """" & Trim(Left$(drvListView.List(i), 1)) & """"
    Next i
    'MsgBox strGrammar
    For i = 4 To Len(InPath) 'Leaving Main directory like C or D or A
        c = Mid(InPath, i, 1)
        If c <> "\" Then
            subDir = subDir & c
        Else
            strGrammar = strGrammar & tagStart & Trim(subDir) & " " & """" & Trim(subDir) & """"
            subDir = ""
        End If
    Next i
End Sub


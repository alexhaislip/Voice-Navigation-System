VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4584
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4584
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H8000000E&
      Height          =   4572
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7380
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   600
         Top             =   3240
      End
      Begin VB.PictureBox picLogo 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   2004
         Left            =   360
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   2004
         ScaleWidth      =   1692
         TabIndex        =   1
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Ruqsana Nazneen"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3360
         TabIndex        =   10
         Top             =   1560
         Width           =   2292
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Voice Navigation System"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   19.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Left            =   2040
         TabIndex        =   3
         Tag             =   "Product"
         Top             =   240
         Width           =   4188
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "T.Gayathri"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3360
         TabIndex        =   9
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   4452
         Left            =   -1200
         Shape           =   2  'Oval
         Top             =   120
         Width           =   2412
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H8000000E&
         Caption         =   "Copyright 2002"
         Height          =   255
         Left            =   4710
         TabIndex        =   8
         Tag             =   "Copyright"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H8000000E&
         Caption         =   "Company Aditya S/W Sols."
         Height          =   255
         Left            =   4710
         TabIndex        =   7
         Tag             =   "Company"
         Top             =   3330
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Version 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5892
         TabIndex        =   6
         Tag             =   "Version"
         Top             =   2760
         Width           =   1116
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Windows 9x / NT Platfrom"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2400
         TabIndex        =   5
         Tag             =   "Platform"
         Top             =   2400
         Width           =   4608
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Project By:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   4
         Tag             =   "CompanyProduct"
         Top             =   840
         Width           =   1272
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "LicenseTo : VEC"
         Height          =   252
         Left            =   240
         TabIndex        =   2
         Tag             =   "LicenseTo"
         Top             =   3720
         Width           =   6852
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub

Private Sub Timer1_Timer()
    'frmStartOptions.Show
    'Unload Me
End Sub

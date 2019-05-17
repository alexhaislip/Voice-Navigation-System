VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   Caption         =   "Plot Four Game"
   ClientHeight    =   6552
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8424
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6552
   ScaleWidth      =   8424
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6570
      Top             =   4395
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6570
      TabIndex        =   67
      Top             =   3675
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF0000&
      Caption         =   "&Blue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   66
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000000FF&
      Caption         =   "&Red"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   65
      Top             =   1320
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "&Restart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6570
      TabIndex        =   64
      Top             =   2955
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Height          =   5808
      Left            =   576
      Shape           =   1  'Square
      Top             =   120
      Width           =   5748
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   0
      Left            =   2736
      TabIndex        =   63
      Tag             =   "d1"
      Top             =   5160
      Width           =   696
   End
   Begin VB.Label lblc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   1
      Left            =   2016
      TabIndex        =   62
      Tag             =   "c2"
      Top             =   4440
      Width           =   696
   End
   Begin VB.Label lble 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   1
      Left            =   3456
      TabIndex        =   61
      Tag             =   "e2"
      Top             =   4440
      Width           =   696
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   2
      Left            =   2736
      TabIndex        =   60
      Tag             =   "d3"
      Top             =   3720
      Width           =   696
   End
   Begin VB.Label lblf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   2
      Left            =   4176
      TabIndex        =   59
      Tag             =   "f3"
      Top             =   3720
      Width           =   696
   End
   Begin VB.Label lble 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   3
      Left            =   3456
      TabIndex        =   58
      Tag             =   "e4"
      Top             =   3000
      Width           =   696
   End
   Begin VB.Label lblg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   3
      Left            =   4896
      TabIndex        =   57
      Tag             =   "g4"
      Top             =   3000
      Width           =   696
   End
   Begin VB.Label lblf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   4
      Left            =   4176
      TabIndex        =   56
      Tag             =   "f5"
      Top             =   2280
      Width           =   696
   End
   Begin VB.Label lblc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   0
      Left            =   2016
      TabIndex        =   55
      Tag             =   "c1"
      Top             =   5160
      Width           =   696
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   1
      Left            =   2736
      TabIndex        =   54
      Tag             =   "d2"
      Top             =   4440
      Width           =   696
   End
   Begin VB.Label lble 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   2
      Left            =   3456
      TabIndex        =   53
      Tag             =   "e3"
      Top             =   3720
      Width           =   696
   End
   Begin VB.Label lblf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   3
      Left            =   4176
      TabIndex        =   52
      Tag             =   "f4"
      Top             =   3000
      Width           =   696
   End
   Begin VB.Label lblg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   4
      Left            =   4896
      TabIndex        =   51
      Tag             =   "g5"
      Top             =   2280
      Width           =   696
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   0
      Left            =   576
      TabIndex        =   50
      Tag             =   "a1"
      Top             =   5160
      Width           =   696
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   0
      Left            =   1296
      TabIndex        =   49
      Tag             =   "b1"
      Top             =   5160
      Width           =   696
   End
   Begin VB.Label lble 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   0
      Left            =   3456
      TabIndex        =   48
      Tag             =   "e1"
      Top             =   5160
      Width           =   696
   End
   Begin VB.Label lblf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   0
      Left            =   4176
      TabIndex        =   47
      Tag             =   "f1"
      Top             =   5160
      Width           =   696
   End
   Begin VB.Label lblg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   0
      Left            =   4896
      TabIndex        =   46
      Tag             =   "g1"
      Top             =   5160
      Width           =   696
   End
   Begin VB.Label lblh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   0
      Left            =   5616
      TabIndex        =   45
      Tag             =   "h1"
      Top             =   5160
      Width           =   696
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   1
      Left            =   576
      TabIndex        =   44
      Tag             =   "a2"
      Top             =   4440
      Width           =   696
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   1
      Left            =   1296
      TabIndex        =   43
      Tag             =   "b2"
      Top             =   4440
      Width           =   696
   End
   Begin VB.Label lblf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   1
      Left            =   4176
      TabIndex        =   42
      Tag             =   "f2"
      Top             =   4440
      Width           =   696
   End
   Begin VB.Label lblg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   1
      Left            =   4896
      TabIndex        =   41
      Tag             =   "g2"
      Top             =   4440
      Width           =   696
   End
   Begin VB.Label lblh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   1
      Left            =   5616
      TabIndex        =   40
      Tag             =   "h2"
      Top             =   4440
      Width           =   696
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   2
      Left            =   576
      TabIndex        =   39
      Tag             =   "a3"
      Top             =   3720
      Width           =   696
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   2
      Left            =   1296
      TabIndex        =   38
      Tag             =   "b3"
      Top             =   3720
      Width           =   696
   End
   Begin VB.Label lblc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   2
      Left            =   2016
      TabIndex        =   37
      Tag             =   "c3"
      Top             =   3720
      Width           =   696
   End
   Begin VB.Label lblg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   2
      Left            =   4896
      TabIndex        =   36
      Tag             =   "g3"
      Top             =   3720
      Width           =   696
   End
   Begin VB.Label lblh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   2
      Left            =   5616
      TabIndex        =   35
      Tag             =   "h3"
      Top             =   3720
      Width           =   696
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   3
      Left            =   576
      TabIndex        =   34
      Tag             =   "a4"
      Top             =   3000
      Width           =   696
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   3
      Left            =   1296
      TabIndex        =   33
      Tag             =   "b4"
      Top             =   3000
      Width           =   696
   End
   Begin VB.Label lblc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   3
      Left            =   2016
      TabIndex        =   32
      Tag             =   "c4"
      Top             =   3000
      Width           =   696
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   3
      Left            =   2736
      TabIndex        =   31
      Tag             =   "d4"
      Top             =   3000
      Width           =   696
   End
   Begin VB.Label lblh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   3
      Left            =   5616
      TabIndex        =   30
      Tag             =   "h4"
      Top             =   3000
      Width           =   696
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   4
      Left            =   576
      TabIndex        =   29
      Tag             =   "a5"
      Top             =   2280
      Width           =   696
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   4
      Left            =   1296
      TabIndex        =   28
      Tag             =   "b5"
      Top             =   2280
      Width           =   696
   End
   Begin VB.Label lblc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   4
      Left            =   2016
      TabIndex        =   27
      Tag             =   "c5"
      Top             =   2280
      Width           =   696
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   4
      Left            =   2736
      TabIndex        =   26
      Tag             =   "d5"
      Top             =   2280
      Width           =   696
   End
   Begin VB.Label lble 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   4
      Left            =   3456
      TabIndex        =   25
      Tag             =   "e5"
      Top             =   2280
      Width           =   696
   End
   Begin VB.Label lblh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   4
      Left            =   5616
      TabIndex        =   24
      Tag             =   "h5"
      Top             =   2280
      Width           =   696
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   5
      Left            =   576
      TabIndex        =   23
      Tag             =   "a6"
      Top             =   1560
      Width           =   696
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   5
      Left            =   1296
      TabIndex        =   22
      Tag             =   "b6"
      Top             =   1560
      Width           =   696
   End
   Begin VB.Label lblc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   5
      Left            =   2016
      TabIndex        =   21
      Tag             =   "c6"
      Top             =   1560
      Width           =   696
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   5
      Left            =   2736
      TabIndex        =   20
      Tag             =   "d6"
      Top             =   1560
      Width           =   696
   End
   Begin VB.Label lble 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   5
      Left            =   3456
      TabIndex        =   19
      Tag             =   "e6"
      Top             =   1560
      Width           =   696
   End
   Begin VB.Label lblf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   5
      Left            =   4176
      TabIndex        =   18
      Tag             =   "f6"
      Top             =   1560
      Width           =   696
   End
   Begin VB.Label lblg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   5
      Left            =   4896
      TabIndex        =   17
      Tag             =   "g6"
      Top             =   1560
      Width           =   696
   End
   Begin VB.Label lblh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   5
      Left            =   5616
      TabIndex        =   16
      Tag             =   "h6"
      Top             =   1560
      Width           =   696
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   6
      Left            =   576
      TabIndex        =   15
      Tag             =   "a7"
      Top             =   840
      Width           =   696
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   6
      Left            =   1296
      TabIndex        =   14
      Tag             =   "b7"
      Top             =   840
      Width           =   696
   End
   Begin VB.Label lblc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   6
      Left            =   2016
      TabIndex        =   13
      Tag             =   "c7"
      Top             =   840
      Width           =   696
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   6
      Left            =   2736
      TabIndex        =   12
      Tag             =   "d7"
      Top             =   840
      Width           =   696
   End
   Begin VB.Label lble 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   6
      Left            =   3456
      TabIndex        =   11
      Tag             =   "e7"
      Top             =   840
      Width           =   696
   End
   Begin VB.Label lblf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   6
      Left            =   4176
      TabIndex        =   10
      Tag             =   "f7"
      Top             =   840
      Width           =   696
   End
   Begin VB.Label lblg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   6
      Left            =   4896
      TabIndex        =   9
      Tag             =   "g7"
      Top             =   840
      Width           =   696
   End
   Begin VB.Label lblh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   6
      Left            =   5616
      TabIndex        =   8
      Tag             =   "h7"
      Top             =   840
      Width           =   696
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   7
      Left            =   576
      TabIndex        =   7
      Tag             =   "a8"
      Top             =   120
      Width           =   696
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   7
      Left            =   1320
      TabIndex        =   6
      Tag             =   "b8"
      Top             =   120
      Width           =   696
   End
   Begin VB.Label lblc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   7
      Left            =   2040
      TabIndex        =   5
      Tag             =   "c8"
      Top             =   120
      Width           =   696
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   7
      Left            =   2736
      TabIndex        =   4
      Tag             =   "d8"
      Top             =   120
      Width           =   696
   End
   Begin VB.Label lble 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   7
      Left            =   3456
      TabIndex        =   3
      Tag             =   "e8"
      Top             =   120
      Width           =   696
   End
   Begin VB.Label lblf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   7
      Left            =   4200
      TabIndex        =   2
      Tag             =   "f8"
      Top             =   120
      Width           =   696
   End
   Begin VB.Label lblg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   7
      Left            =   4896
      TabIndex        =   1
      Tag             =   "g8"
      Top             =   120
      Width           =   696
   End
   Begin VB.Label lblh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   696
      Index           =   7
      Left            =   5616
      TabIndex        =   0
      Tag             =   "h8"
      Top             =   120
      Width           =   696
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Ctr As String * 1 'Taken for Marking Square Caption as "o"
Dim PlotFoured As Boolean
Dim i As Integer 'Taken as Counter
Dim ArrBlink4(1 To 4) As label

Private Sub cmdExit_Click()
    Unload Me
    'End
End Sub

Private Sub cmdRestart_Click()
PlotFoured = False
    For i = 0 To 7
        lbla(i).Caption = "": lbla(i).BackColor = vbWhite
        lblb(i).Caption = "": lblb(i).BackColor = vbWhite
        lblc(i).Caption = "": lblc(i).BackColor = vbWhite
        lbld(i).Caption = "": lbld(i).BackColor = vbWhite
        lble(i).Caption = "": lble(i).BackColor = vbWhite
        lblf(i).Caption = "": lblf(i).BackColor = vbWhite
        lblg(i).Caption = "": lblg(i).BackColor = vbWhite
        lblh(i).Caption = "": lblh(i).BackColor = vbWhite
    Next i
    For i = 0 To 7
        lbla(i).Caption = "": lbla(i).ForeColor = vbWhite
        lblb(i).Caption = "": lblb(i).ForeColor = vbWhite
        lblc(i).Caption = "": lblc(i).ForeColor = vbWhite
        lbld(i).Caption = "": lbld(i).ForeColor = vbWhite
        lble(i).Caption = "": lble(i).ForeColor = vbWhite
        lblf(i).Caption = "": lblf(i).ForeColor = vbWhite
        lblg(i).Caption = "": lblg(i).ForeColor = vbWhite
        lblh(i).Caption = "": lblh(i).ForeColor = vbWhite
    Next i
End Sub

Private Sub Form_Load()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    
    Ctr = "o"
    For i = 0 To 7
        lbla(i).Caption = "": lbla(i).BackColor = vbWhite
        lblb(i).Caption = "": lblb(i).BackColor = vbWhite
        lblc(i).Caption = "": lblc(i).BackColor = vbWhite
        lbld(i).Caption = "": lbld(i).BackColor = vbWhite
        lble(i).Caption = "": lble(i).BackColor = vbWhite
        lblf(i).Caption = "": lblf(i).BackColor = vbWhite
        lblg(i).Caption = "": lblg(i).BackColor = vbWhite
        lblh(i).Caption = "": lblh(i).BackColor = vbWhite
    Next i
PlotFoured = False
DisplayCoors
frmMain.g_iFlag = FORMGAME
End Sub

Private Sub Form_Resize()
'    Me.Height = frmMain.ScaleHeight
'    Me.Width = frmMain.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.g_iFlag = FORMMAIN
frmMain.DSR.GrammarFromFile App.Path + "\" + frmMain.gGrammarFile + ".def"
End Sub

Public Sub lbla_Click(Index As Integer)
For i = 0 To 7
    If lbla(i).Caption = "" Then
        lbla(i).Caption = Ctr
        Exit For
    End If
Next i
If i = 8 Then Exit Sub
If Option1.Value = True Then
    lbla(i).ForeColor = Option1.BackColor: lbla(i).BackColor = Option1.BackColor
Else
    lbla(i).ForeColor = Option2.BackColor: lbla(i).BackColor = Option2.BackColor
End If
Chk18 lbla(i)
Option1.Value = Not (Option1.Value)
Option2.Value = Not (Option1.Value)
End Sub

Public Sub lblb_Click(Index As Integer)
For i = 0 To 7
    If lblb(i).Caption = "" Then
        lblb(i).Caption = Ctr
        Exit For
    End If
Next i
If i = 8 Then Exit Sub
If Option1.Value = True Then
    lblb(i).ForeColor = Option1.BackColor: lblb(i).BackColor = Option1.BackColor
Else
    lblb(i).ForeColor = Option2.BackColor: lblb(i).BackColor = Option2.BackColor
End If
Chk18 lblb(i)
Option1.Value = Not (Option1.Value)
Option2.Value = Not (Option1.Value)
End Sub

Public Sub lblc_Click(Index As Integer)
For i = 0 To 7
    If lblc(i).Caption = "" Then
        lblc(i).Caption = Ctr
        Exit For
    End If
Next i
If i = 8 Then Exit Sub
If Option1.Value = True Then
    lblc(i).ForeColor = Option1.BackColor: lblc(i).BackColor = Option1.BackColor
Else
    lblc(i).ForeColor = Option2.BackColor: lblc(i).BackColor = Option2.BackColor
End If
Chk18 lblc(i)
Option1.Value = Not (Option1.Value)
Option2.Value = Not (Option1.Value)
End Sub

Public Sub lbld_Click(Index As Integer)
For i = 0 To 7
    If lbld(i).Caption = "" Then
        lbld(i).Caption = Ctr
        Exit For
    End If
Next i
If i = 8 Then Exit Sub
If Option1.Value = True Then
    lbld(i).ForeColor = Option1.BackColor: lbld(i).BackColor = Option1.BackColor
Else
    lbld(i).ForeColor = Option2.BackColor: lbld(i).BackColor = Option2.BackColor
End If
Chk18 lbld(i)
Option1.Value = Not (Option1.Value)
Option2.Value = Not (Option1.Value)
End Sub

Public Sub lble_Click(Index As Integer)
For i = 0 To 7
    If lble(i).Caption = "" Then
        lble(i).Caption = Ctr
        Exit For
    End If
Next i
If i = 8 Then Exit Sub
If Option1.Value = True Then
    lble(i).ForeColor = Option1.BackColor: lble(i).BackColor = Option1.BackColor
Else
    lble(i).ForeColor = Option2.BackColor: lble(i).BackColor = Option2.BackColor
End If
Chk18 lble(i)
Option1.Value = Not (Option1.Value)
Option2.Value = Not (Option1.Value)
End Sub

Public Sub lblf_Click(Index As Integer)
For i = 0 To 7
    If lblf(i).Caption = "" Then
        lblf(i).Caption = Ctr
        Exit For
    End If
Next i
If i = 8 Then Exit Sub
If Option1.Value = True Then
    lblf(i).ForeColor = Option1.BackColor: lblf(i).BackColor = Option1.BackColor
Else
    lblf(i).ForeColor = Option2.BackColor: lblf(i).BackColor = Option2.BackColor
End If
Chk18 lblf(i)
Option1.Value = Not (Option1.Value)
Option2.Value = Not (Option1.Value)
End Sub

Public Sub lblg_Click(Index As Integer)
For i = 0 To 7
    If lblg(i).Caption = "" Then
        lblg(i).Caption = Ctr
        Exit For
    End If
Next i
If i = 8 Then Exit Sub
If Option1.Value = True Then
    lblg(i).ForeColor = Option1.BackColor: lblg(i).BackColor = Option1.BackColor
Else
    lblg(i).ForeColor = Option2.BackColor: lblg(i).BackColor = Option2.BackColor
End If
Chk18 lblg(i)
Option1.Value = Not (Option1.Value)
Option2.Value = Not (Option1.Value)
End Sub

Public Sub lblh_Click(Index As Integer)
For i = 0 To 7
    If lblh(i).Caption = "" Then
        lblh(i).Caption = Ctr
        Exit For
    End If
Next i
If i = 8 Then Exit Sub
If Option1.Value = True Then
    lblh(i).ForeColor = Option1.BackColor: lblh(i).BackColor = Option1.BackColor
Else
    lblh(i).ForeColor = Option2.BackColor: lblh(i).BackColor = Option2.BackColor
End If
'Check4Win lblh(i)
Chk18 lblh(i)
Option1.Value = Not (Option1.Value)
Option2.Value = Not (Option1.Value)
End Sub

Public Sub Chk18(ByRef CurrentBox As label)
Set ArrBlink4(1) = CurrentBox

'****************** Vertical ***************************

Dim Vert(0 To 7) As label
Dim Horz(0 To 7) As label
Dim CurrentLabel As label
Dim Cntr, Id As Byte
Dim xCor As Integer
Set CurrentLabel = CurrentBox
Dim Tempid As Byte
Select Case CurrentBox.Name
    Case "lbla"
        For i = 0 To 7
            Set Vert(i) = lbla(i)
        Next i
    Case "lblb"
        For i = 0 To 7
            Set Vert(i) = lblb(i)
        Next i
    Case "lblc"
        For i = 0 To 7
            Set Vert(i) = lblc(i)
        Next i
    Case "lbld"
        For i = 0 To 7
            Set Vert(i) = lbld(i)
        Next i
    Case "lble"
        For i = 0 To 7
            Set Vert(i) = lble(i)
        Next i
    Case "lblf"
        For i = 0 To 7
            Set Vert(i) = lblf(i)
        Next i
    Case "lblg"
        For i = 0 To 7
            Set Vert(i) = lblg(i)
        Next i
    Case "lblh"
        For i = 0 To 7
            Set Vert(i) = lblh(i)
        Next i
End Select

Cntr = 1
For i = 0 To 6
If Vert(i).Caption = "" Then Exit For ' Nothing in Vert(i)
    If Vert(i).ForeColor = Vert(i + 1).ForeColor Then
        Set ArrBlink4(Cntr) = Vert(i) 'Set current To-Blink Square in Array
        Cntr = Cntr + 1 '****todod

        If Cntr = 4 Then
            Set ArrBlink4(Cntr) = Vert(i + 1)
            MsgBox "Plot Four"
            PlotFoured = True
            SelectAndBlinkFour
            Exit Sub
        End If
    Else
        Cntr = 1
    End If
Next i
'        Set ArrBlink4(Cntr) = Vert(i) 'Set current To-Blink Square in Array
'            Set ArrBlink4(Cntr) = Vert(i + 1)
'            SelectAndBlinkFour

'*********************************  Horizontal  **************************************
Id = CurrentBox.Index
Set Horz(0) = lbla(Id): Set Horz(1) = lblb(Id): Set Horz(2) = lblc(Id): Set Horz(3) = lbld(Id)
Set Horz(4) = lble(Id): Set Horz(5) = lblf(Id): Set Horz(6) = lblg(Id): Set Horz(7) = lblh(Id)
Cntr = 1
For i = 0 To 6
    If Horz(i).ForeColor = Horz(i + 1).ForeColor And Horz(i).Caption <> "" Then
        Set ArrBlink4(Cntr) = Horz(i) 'Set current To-Blink Square in Array
        Cntr = Cntr + 1
        If Cntr = 4 Then
                Set ArrBlink4(Cntr) = Horz(i + 1)
            MsgBox "Plot Four"
            PlotFoured = True
                SelectAndBlinkFour
            Exit Sub
        End If
    Else
        Cntr = 1
    End If
Next i

'******************  Diagonal UP  ****************************
Dim TDArray(0 To 7, 0 To 7) As label 'Rows,Cols
Dim StartLblUP As label
For i = 0 To 7
    Set TDArray(0, i) = lbla(i)
    Set TDArray(1, i) = lblb(i)
    Set TDArray(2, i) = lblc(i)
    Set TDArray(3, i) = lbld(i)
    Set TDArray(4, i) = lble(i)
    Set TDArray(5, i) = lblf(i)
    Set TDArray(6, i) = lblg(i)
    Set TDArray(7, i) = lblh(i)
Next i
'For asceraitaining StartLblUP
Select Case CurrentBox.Name
    Case "lbla"
        Select Case Id
            Case 0: Set StartLblUP = lbla(0)
            Case 1: Set StartLblUP = lbla(1)
            Case 2: Set StartLblUP = lbla(2)
            Case 3: Set StartLblUP = lbla(3)
            Case 4: Set StartLblUP = lbla(4)
            Case Else: GoTo SkipDiagonalUPCheck
        End Select
    Case "lblb"
        Select Case Id
            Case 0: Set StartLblUP = lblb(0)
            Case 1: Set StartLblUP = lbla(0)
            Case 2: Set StartLblUP = lbla(1)
            Case 3: Set StartLblUP = lbla(2)
            Case 4: Set StartLblUP = lbla(3)
            Case 5: Set StartLblUP = lbla(4)
            Case Else: GoTo SkipDiagonalUPCheck
        End Select
    Case "lblc"
        Select Case Id
            Case 0: Set StartLblUP = lblc(0)
            Case 1: Set StartLblUP = lblb(0)
            Case 2: Set StartLblUP = lbla(0)
            Case 3: Set StartLblUP = lbla(1)
            Case 4: Set StartLblUP = lbla(2)
            Case 5: Set StartLblUP = lbla(3)
            Case 6: Set StartLblUP = lbla(4)
            Case Else: GoTo SkipDiagonalUPCheck
        End Select
    Case "lbld"
        Select Case Id
            Case 0: Set StartLblUP = lbld(0)
            Case 1: Set StartLblUP = lblc(0)
            Case 2: Set StartLblUP = lblb(0)
            Case 3: Set StartLblUP = lbla(0)
            Case 4: Set StartLblUP = lbla(1)
            Case 5: Set StartLblUP = lbla(2)
            Case 6: Set StartLblUP = lbla(3)
            Case 7: Set StartLblUP = lbla(4)
            Case Else: GoTo SkipDiagonalUPCheck
        End Select
    Case "lble"
        Select Case Id
            Case 0: Set StartLblUP = lble(0)
            Case 1: Set StartLblUP = lbld(0)
            Case 2: Set StartLblUP = lblc(0)
            Case 3: Set StartLblUP = lblb(0)
            Case 4: Set StartLblUP = lbla(0)
            Case 5: Set StartLblUP = lbla(1)
            Case 6: Set StartLblUP = lbla(2)
            Case 7: Set StartLblUP = lbla(3)
            Case Else: GoTo SkipDiagonalUPCheck
        End Select
    Case "lblf"
        Select Case Id
            Case 1: Set StartLblUP = lble(0)
            Case 2: Set StartLblUP = lbld(0)
            Case 3: Set StartLblUP = lblc(0)
            Case 4: Set StartLblUP = lblb(0)
            Case 5: Set StartLblUP = lbla(0)
            Case 6: Set StartLblUP = lbla(1)
            Case 7: Set StartLblUP = lbla(2)
            Case Else: GoTo SkipDiagonalUPCheck
        End Select
    Case "lblg"
        Select Case Id
            Case 2: Set StartLblUP = lble(0)
            Case 3: Set StartLblUP = lbld(0)
            Case 4: Set StartLblUP = lblc(0)
            Case 5: Set StartLblUP = lblb(0)
            Case 6: Set StartLblUP = lbla(0)
            Case 7: Set StartLblUP = lbla(1)
            Case Else: GoTo SkipDiagonalUPCheck
        End Select
    Case "lblh"
        Select Case Id
            Case 3: Set StartLblUP = lble(0)
            Case 4: Set StartLblUP = lbld(0)
            Case 5: Set StartLblUP = lblc(0)
            Case 6: Set StartLblUP = lblb(0)
            Case 7: Set StartLblUP = lbla(0)
            Case Else: GoTo SkipDiagonalUPCheck
        End Select
End Select
'Now we know the Start UP label
'Start with check now
'If StartLblUP Then
    Select Case StartLblUP.Name
        Case "lbla"
            xCor = 1 'Taken for X-Axis
        Case "lblb"
            xCor = 2 'Since lbla,lblb .... are stored  as 0,1,2...7
        Case "lblc"
            xCor = 3
        Case "lbld"
            xCor = 4
        Case "lble"
            xCor = 5
        Case Else
            MsgBox "Not found"
    End Select

Dim TempLabel As label
Dim yCOR As Integer
Dim NextLabel As label 'Label 2 check for next

Set CurrentLabel = StartLblUP 'CurrentLabel SET
Cntr = 1
yCOR = StartLblUP.Index + 1
Set NextLabel = TDArray(xCor, yCOR) 'NextLabel SET
'JUST REPLACE TEMPID WITH YCOR AND CHECK
    'Do While CurrentLabel.Name <> "lblg" And yCOR < 7    'Tempid <= 6
    Do While xCor <= 7 And yCOR <= 7   'Tempid <= 6
        'If xCor >= 7 Or yCOR >= 7 Then Exit Do
        
        If CurrentLabel.ForeColor = NextLabel.ForeColor And CurrentLabel.Caption <> "" Then
        'Squares Matching
         Set ArrBlink4(Cntr) = CurrentLabel 'Set current To-Blink Square in Array
           
            Cntr = Cntr + 1
                If Cntr = 4 Then
                        Set ArrBlink4(Cntr) = NextLabel
                    MsgBox "Plot Four"
                    PlotFoured = True
                        SelectAndBlinkFour
                    Exit Sub
                End If
            xCor = xCor + 1
            yCOR = yCOR + 1
            If xCor > 7 Or yCOR > 7 Then Exit Do
            Set TempLabel = TDArray(xCor, yCOR)
            Set CurrentLabel = NextLabel
            Set NextLabel = TempLabel
        Else
            Cntr = 1
            xCor = xCor + 1
            yCOR = yCOR + 1
            If xCor > 7 Or yCOR > 7 Then Exit Do
            Set TempLabel = TDArray(xCor, yCOR)
            Set CurrentLabel = NextLabel
            Set NextLabel = TempLabel
            'Tempid = Tempid + 1
        End If
'        xCor = xCor + 1
'        yCOR = yCOR + 1
    Loop 'Wend
    
SkipDiagonalUPCheck:         'Control is sent here if Label is "f to g"

'******************  Diagonal DOWN  ****************************
'TDArray(0 To 7, 0 To 7) As Label 'Rows,Cols is set at Diagonal UP
Dim StartLblDOWN As label
'For asceraitaining StartLblDOWN
Select Case CurrentBox.Name
    Case "lbla"
        Select Case Id
            Case 7: Set StartLblDOWN = lbla(7)
            Case 6: Set StartLblDOWN = lbla(6)
            Case 5: Set StartLblDOWN = lbla(5)
            Case 4: Set StartLblDOWN = lbla(4)
            Case 3: Set StartLblDOWN = lbla(3)
            Case Else: GoTo SkipDiagonalDOWNCheck
        End Select
    Case "lblb"
        Select Case Id
            Case 7: Set StartLblDOWN = lblb(7)
            Case 6: Set StartLblDOWN = lbla(7)
            Case 5: Set StartLblDOWN = lbla(6)
            Case 4: Set StartLblDOWN = lbla(5)
            Case 3: Set StartLblDOWN = lbla(4)
            Case 2: Set StartLblDOWN = lbla(3)
            Case Else: GoTo SkipDiagonalDOWNCheck
        End Select
    Case "lblc"
        Select Case Id
            Case 7: Set StartLblDOWN = lblc(7)
            Case 6: Set StartLblDOWN = lblb(7)
            Case 5: Set StartLblDOWN = lbla(7)
            Case 4: Set StartLblDOWN = lbla(6)
            Case 3: Set StartLblDOWN = lbla(5)
            Case 2: Set StartLblDOWN = lbla(4)
            Case 2: Set StartLblDOWN = lbla(3)
            Case Else: GoTo SkipDiagonalDOWNCheck
        End Select
    Case "lbld"
        Select Case Id
            Case 7: Set StartLblDOWN = lbld(7)
            Case 6: Set StartLblDOWN = lblc(7)
            Case 5: Set StartLblDOWN = lblb(7)
            Case 4: Set StartLblDOWN = lbla(7)
            Case 3: Set StartLblDOWN = lbla(6)
            Case 2: Set StartLblDOWN = lbla(5)
            Case 1: Set StartLblDOWN = lbla(4)
            Case 0: Set StartLblDOWN = lbla(3)
            Case Else: GoTo SkipDiagonalDOWNCheck
        End Select
    Case "lble"
        Select Case Id
            Case 7: Set StartLblDOWN = lble(7)
            Case 6: Set StartLblDOWN = lbld(7)
            Case 5: Set StartLblDOWN = lblc(7)
            Case 4: Set StartLblDOWN = lblb(7)
            Case 3: Set StartLblDOWN = lbla(7)
            Case 2: Set StartLblDOWN = lbla(6)
            Case 1: Set StartLblDOWN = lbla(5)
            Case 0: Set StartLblDOWN = lbla(4)
            Case Else: GoTo SkipDiagonalDOWNCheck
        End Select
    Case "lblf"
        Select Case Id
            Case 6: Set StartLblDOWN = lble(7)
            Case 5: Set StartLblDOWN = lbld(7)
            Case 4: Set StartLblDOWN = lblc(7)
            Case 3: Set StartLblDOWN = lblb(7)
            Case 2: Set StartLblDOWN = lbla(7)
            Case 1: Set StartLblDOWN = lbla(6)
            Case 0: Set StartLblDOWN = lbla(5)
            Case Else: GoTo SkipDiagonalDOWNCheck
        End Select
    Case "lblg"
        Select Case Id
            Case 5: Set StartLblDOWN = lble(7)
            Case 4: Set StartLblDOWN = lbld(7)
            Case 3: Set StartLblDOWN = lblc(7)
            Case 2: Set StartLblDOWN = lblb(7)
            Case 1: Set StartLblDOWN = lbla(7)
            Case 0: Set StartLblDOWN = lbla(6)
            Case Else: GoTo SkipDiagonalDOWNCheck
        End Select
    Case "lblh"
        Select Case Id
            Case 4: Set StartLblDOWN = lble(7)
            Case 3: Set StartLblDOWN = lbld(7)
            Case 2: Set StartLblDOWN = lblc(7)
            Case 1: Set StartLblDOWN = lblb(7)
            Case 0: Set StartLblDOWN = lbla(7)
            Case Else: GoTo SkipDiagonalDOWNCheck
        End Select
End Select
'Now we know the Start DOWN label
'Start with check now
'If StartLblDOWN Then
    Select Case StartLblDOWN.Name
        Case "lbla"
            xCor = 1 'Taken for X-Axis
        Case "lblb"
            xCor = 2 'Since lbla,lblb .... are stored  as 0,1,2...7
        Case "lblc"
            xCor = 3
        Case "lbld"
            xCor = 4
        Case "lble"
            xCor = 5

    End Select

'Start checking for Diagnol Down
Set CurrentLabel = StartLblDOWN
yCOR = StartLblDOWN.Index - 1 'Taking next pt on Y-Axis
Cntr = 1 'To Count Matched Squares
'Tempid = StartLblDOWN.Index 'Setting for checking till y-axis =>6
'JUST REPLACE TEMPID WITH YCOR AND CHECK
Set NextLabel = TDArray(xCor, yCOR)
    
    'Do While CurrentLabel.Name <> "lblg" Or yCOR >= 0 ' Tempid > 1  'TODO TEMP LBL NAME
    Do While xCor <= 7 And yCOR >= 0  'Tempid <= 6
        'If xCor >= 7 Or yCOR <= 0 Then Exit Do
        If CurrentLabel.ForeColor = NextLabel.ForeColor And CurrentLabel.Caption <> "" Then
        Set ArrBlink4(Cntr) = CurrentLabel 'Set current To-Blink Square in Array
            Cntr = Cntr + 1
                If Cntr = 4 Then
                        Set ArrBlink4(Cntr) = NextLabel
                    MsgBox "Plot Four"
                    PlotFoured = True
                        SelectAndBlinkFour
                    Exit Sub
                End If
            xCor = xCor + 1
            yCOR = yCOR - 1
            If xCor > 7 Or yCOR < 0 Then Exit Do
            Set TempLabel = TDArray(xCor, yCOR)
            Set CurrentLabel = NextLabel
            Set NextLabel = TempLabel
        Else
            Cntr = 1
            xCor = xCor + 1
            yCOR = yCOR - 1
            If xCor > 7 Or yCOR < 0 Then Exit Do
            Set TempLabel = TDArray(xCor, yCOR)
            Set CurrentLabel = NextLabel
            Set NextLabel = TempLabel
        End If
    Loop 'Wend
Exit Sub
SkipDiagonalDOWNCheck:
End Sub

Sub SelectAndBlinkFour()
        Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    For i = 1 To 4
         ArrBlink4(i).BackColor = vbGreen
    Next i
    Timer1.Enabled = False
End Sub
Private Sub DisplayCoors()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
    FontSize = 16
    CurrentX = 844
    CurrentY = 6000
    Print "1   "; "    2   "; "    3   "; "    4   "; "     5    "; "    6    "; "    7   "; "     8   "
End Sub

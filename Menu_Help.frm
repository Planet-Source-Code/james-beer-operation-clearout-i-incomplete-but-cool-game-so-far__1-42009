VERSION 5.00
Begin VB.Form Menu_Help 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Help"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   Enabled         =   0   'False
   LinkTopic       =   "Form4"
   Picture         =   "Menu_Help.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Menu_Help_Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Done"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Pause/Resume Game           = P Key"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Activate Item                         = ENTER Key"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Item Select (not working)       = ""["" or ""]"""
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Weapon Select                     = 1 to 9 Keys"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Use / Swicth                            = E Key"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Cover / Fly Lower                     = DOWN Key"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Shoot / Detonate                     = CTRL Key"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Move Right / Speed Plane Up = RIGHT Key"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Move Left / Slow Plane Down = LEFT Key"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Jump / Fly Higher                     = UP Key"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00004000&
      Caption         =   "Controls:"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   $"Menu_Help.frx":CCAD
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Menu_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Menu_Help_Command1_Click()
Menu_Help.Enabled = False
Menu_Help.Hide
Unload Menu_Help
Form1.Show
Form1.Enabled = True
End Sub

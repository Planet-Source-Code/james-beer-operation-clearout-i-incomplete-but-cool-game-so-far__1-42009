VERSION 5.00
Begin VB.Form Menu_Objectives 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8940
   LinkTopic       =   "Form4"
   Picture         =   "Menu_Objectives.frx":0000
   ScaleHeight     =   4575
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label MBText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   8415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Review Berifing"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Menu_Objectives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Menu_Objectives
Form1.Enabled = True
Form1.Show
End Sub

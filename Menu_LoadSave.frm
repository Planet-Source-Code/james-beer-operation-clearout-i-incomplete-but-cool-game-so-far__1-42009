VERSION 5.00
Begin VB.Form Menu_LoadSave 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Load/Save Campaign"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   Enabled         =   0   'False
   LinkTopic       =   "Form4"
   Picture         =   "Menu_LoadSave.frx":0000
   ScaleHeight     =   4815
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Close"
      Height          =   375
      Index           =   3
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   3480
      Width           =   3975
   End
   Begin VB.ListBox List 
      Height          =   2205
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Delete"
      Height          =   375
      Index           =   2
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Load"
      Height          =   375
      Index           =   0
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Not Available for Trail Version!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Load Campaign"
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
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "Menu_LoadSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case Is = 3

If Form1.Pause_Screen.Visible = False Then
Form1.Menu_Screen.Enabled = True
Else
Command1(1).Enabled = False
End If

Menu_LoadSave.Hide
Menu_LoadSave.Enabled = False
Unload Menu_LoadSave
Form1.Show
Form1.Enabled = True


End Select
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To Command1.UBound
If Command1(i).BackColor <> &HC000& Then Command1(i).BackColor = &HC000&: Command1(i).FontBold = False
Next i
Command1(Index).BackColor = vbRed: Command1(Index).FontBold = True
End Sub


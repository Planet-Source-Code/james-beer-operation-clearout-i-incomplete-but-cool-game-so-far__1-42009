VERSION 5.00
Begin VB.Form Menu_Options 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   Enabled         =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Menu_Options.frx":0000
   ScaleHeight     =   4815
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Difficulty 
      BackColor       =   &H00000000&
      Caption         =   "Easy"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   13
      Top             =   1200
      Value           =   -1  'True
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cheats"
      ForeColor       =   &H0000FFFF&
      Height          =   2415
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   4695
      Begin VB.CheckBox Cheat 
         BackColor       =   &H00000000&
         Caption         =   "All Enemies are Osamas (Not Available on Trail)"
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   4455
      End
      Begin VB.CheckBox Cheat 
         BackColor       =   &H00000000&
         Caption         =   "Enable [ Zombie Mode ] (Not Available on Trail)"
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   4455
      End
      Begin VB.CheckBox Cheat 
         BackColor       =   &H00000000&
         Caption         =   "Enable [Nuclear Strike (Press X)]"
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   4455
      End
      Begin VB.CheckBox Cheat 
         BackColor       =   &H00000000&
         Caption         =   "Jetpack Mode"
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   4455
      End
      Begin VB.CheckBox Cheat 
         BackColor       =   &H00000000&
         Caption         =   "One Hit Kills"
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   4455
      End
      Begin VB.CheckBox Cheat 
         BackColor       =   &H00000000&
         Caption         =   "Infinite Ammuntion"
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4455
      End
      Begin VB.CheckBox Cheat 
         BackColor       =   &H00000000&
         Caption         =   "All Weapons/Items"
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   4455
      End
      Begin VB.CheckBox Cheat 
         BackColor       =   &H00000000&
         Caption         =   "Invurnerable"
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.OptionButton Difficulty 
      BackColor       =   &H00000000&
      Caption         =   "Medium"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.OptionButton Difficulty 
      BackColor       =   &H00000000&
      Caption         =   "Hard"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Accept"
      Height          =   375
      Index           =   0
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Difficulty:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Game Options"
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
      Index           =   1
      Left            =   720
      TabIndex        =   14
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Menu_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case Is = 0
If Difficulty(0).Value = True Then Game_Difficulty = "Easy"
If Difficulty(1).Value = True Then Game_Difficulty = "Normal"
If Difficulty(2).Value = True Then Game_Difficulty = "Hard"

For i = 0 To 7
If Cheat(i).Value = 1 Then Cheats(i) = True Else Cheats(i) = False
Next i

Menu_Options.Enabled = False
Menu_Options.Hide
Unload Menu_Options
Form1.Show
Form1.Enabled = True

Case Is = 1

If Game_Difficulty = "Easy" Then Difficulty(0).Value = True
If Game_Difficulty = "Normal" Then Difficulty(1).Value = True
If Game_Difficulty = "Hard" Then Difficulty(2).Value = True

For i = 0 To 7
If Cheats(i) = True Then Cheat(i).Value = 1 Else Cheat(i).Value = 0
Next i

Menu_Options.Enabled = False
Menu_Options.Hide
Unload Menu_Options
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

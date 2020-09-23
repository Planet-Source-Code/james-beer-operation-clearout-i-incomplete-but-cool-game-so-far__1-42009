VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Before Playing the Game"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000A&
   Icon            =   "Loading.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Loading.frx":0E42
   ScaleHeight     =   187
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Techincal Information"
      Height          =   855
      Left            =   1800
      Picture         =   "Loading.frx":76F2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System Requirements"
      Height          =   855
      Left            =   120
      MaskColor       =   &H0000C000&
      Picture         =   "Loading.frx":7CB8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Continue"
      Height          =   855
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Loading.frx":88FA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Please! Goto 'HELP' in the main menu for key commands!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Loading.frx":953C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Clearout I (tm)                  Created and Programmed by:         James Beer M1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()
MsgBox "System Requirements for Operation Clearout I:" & Chr(13) & _
       "OS: Win9x, Win2000 or WinXP" & Chr(13) & _
       "Processor: PIII, 200Mhz minimum" & Chr(13) & _
       "Memory: 32MB Mininum - 64MB Recommended" & Chr(13) & _
       "Disk Space: 5MB Recommended" & Chr(13) & _
       "Other: Keyboard, Mouse and General MIDI Compatible Mixer", vbInformation, "Game System Requirements"
End Sub

Private Sub Command3_Click()
MsgBox "If any game or techical problems occur contact me immediantly at 'Beerboy160@hotmail.com', and send as much information about the problem as possible. Or Send a comment to my code on Planet-Source-Code.com, I'll answer them there too.", vbExclamation, "Techincal Support"
End Sub

Private Sub Form_Load()
Unload Menu_Help
'Unload Menu_LevelSelect
Unload Menu_LoadSave
Unload Menu_Options

If App.PrevInstance = True Then
MsgBox "Operation Clearout 1 cannot run when the same application is already running", vbExclamation, "Program Terminated"
End
End If

End Sub

Private Sub Form_Terminate()
Unload Menu_Help
'Unload Menu_LevelSelect
Unload Menu_LoadSave
Unload Menu_Options
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Menu_Help
'Unload Menu_LevelSelect
Unload Menu_LoadSave
Unload Menu_Options
End
End Sub


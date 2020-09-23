VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Operation Clearout"
   ClientHeight    =   8145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9930
   Icon            =   "Game.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   543
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   662
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Credit_Screen 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   0
      Picture         =   "Game.frx":0E42
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   9975
      Begin VB.Timer Credit_OsamaDraw 
         Enabled         =   0   'False
         Interval        =   125
         Left            =   360
         Top             =   960
      End
      Begin VB.PictureBox Credits_Osama 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   4560
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   40
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Credits_Command1 
         BackColor       =   &H000000FF&
         Caption         =   "Return to Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Timer CreditTime 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   360
         Top             =   240
      End
      Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
         Height          =   615
         Left            =   7920
         TabIndex        =   48
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   0
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   -1  'True
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   0
         WindowlessVideo =   0   'False
      End
      Begin VB.Label CreditCheat 
         BackStyle       =   0  'Transparent
         Caption         =   "You've successfully completed the game without cheats!, Press O - B - L to goto final level instantly!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   1
         Left            =   1920
         TabIndex        =   47
         Top             =   5400
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label CreditCheat 
         BackStyle       =   0  'Transparent
         Caption         =   $"Game.frx":3DE8
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Index           =   0
         Left            =   1920
         TabIndex        =   41
         Top             =   4560
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label Credit_Taunt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ha Ha Ha, You can't get me yet Ryan and I'll be waiting for you prepared when you do..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   1800
         TabIndex        =   38
         Top             =   3120
         Width           =   6495
      End
      Begin VB.Label Credit_Text 
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Game.frx":3E7C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   10095
         Left            =   840
         TabIndex        =   37
         Top             =   -2400
         Width           =   8175
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   33
      Top             =   7560
      Visible         =   0   'False
      Width           =   9975
      Begin VB.Label VersionData 
         BackStyle       =   0  'Transparent
         Caption         =   "OC_THFO v1.01.0000 Trail"
         ForeColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.PictureBox QMsg 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3240
      ScaleHeight     =   1215
      ScaleWidth      =   3855
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "ERROR: No message was loaded"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   975
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.PictureBox Pause_Screen 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   3480
      Picture         =   "Game.frx":45B0
      ScaleHeight     =   3735
      ScaleWidth      =   3375
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton Pause_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Quit Game"
         Height          =   495
         Index           =   4
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CommandButton Pause_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Exit to Menu"
         Height          =   495
         Index           =   3
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CommandButton Pause_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Restart Level"
         Height          =   495
         Index           =   2
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CommandButton Pause_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Review Berifing"
         Height          =   495
         Index           =   5
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Pause_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Load/Save Campaign"
         Height          =   495
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton Pause_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Resume"
         Height          =   495
         Index           =   0
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Pause Menu"
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
         Left            =   360
         TabIndex        =   23
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   4350
      Index           =   1
      Left            =   0
      Picture         =   "Game.frx":AE60
      ScaleHeight     =   290
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   705
      TabIndex        =   11
      Top             =   4080
      Width           =   10575
      Begin VB.CommandButton Menu_Msg 
         BackColor       =   &H0000C000&
         Caption         =   "Continue"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Menu_Msg 
         BackColor       =   &H0000C000&
         Caption         =   "Quit"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton Menu_Msg 
         BackColor       =   &H0000C000&
         Caption         =   "Play"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label ScoreTab 
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
         ForeColor       =   &H00C0FFFF&
         Height          =   1575
         Left            =   1920
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.Label LoadStats 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label LoadStats 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label LoadStats 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label LoadStats 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Image Image5 
         Height          =   480
         Index           =   1
         Left            =   0
         Picture         =   "Game.frx":17B0D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10095
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4080
      Index           =   0
      Left            =   0
      Picture         =   "Game.frx":19892
      ScaleHeight     =   272
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   705
      TabIndex        =   10
      Top             =   0
      Width           =   10575
      Begin VB.Timer BT 
         Interval        =   1
         Left            =   8280
         Top             =   240
      End
      Begin VB.Timer Msg_Door 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7800
         Top             =   240
      End
      Begin VB.Timer QuickMsg 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   7320
         Top             =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   9360
         Top             =   240
      End
      Begin VB.PictureBox Menu_Info 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   120
         ScaleHeight     =   3015
         ScaleWidth      =   9615
         TabIndex        =   28
         Top             =   480
         Width           =   9615
         Begin VB.PictureBox PB1 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   641
            TabIndex        =   29
            Top             =   2760
            Visible         =   0   'False
            Width           =   9615
            Begin VB.Shape Shape1 
               BorderColor     =   &H00000000&
               FillColor       =   &H000080FF&
               FillStyle       =   0  'Solid
               Height          =   495
               Left            =   0
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.Label Msg 
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
            Height          =   2415
            Index           =   0
            Left            =   0
            TabIndex        =   31
            Tag             =   $"Game.frx":2653F
            Top             =   360
            Width           =   9615
         End
         Begin VB.Label Msg_LevelName 
            BackColor       =   &H000000C0&
            Caption         =   "Main Menu"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   9615
         End
      End
      Begin VB.Image Image5 
         Height          =   480
         Index           =   0
         Left            =   0
         Picture         =   "Game.frx":265D6
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   10095
      End
   End
   Begin VB.PictureBox Menu_Screen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8175
      Left            =   0
      Picture         =   "Game.frx":2835B
      ScaleHeight     =   8175
      ScaleWidth      =   9975
      TabIndex        =   2
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton Menu_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Contact Me!"
         Height          =   615
         Index           =   6
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   6480
         Width           =   3135
      End
      Begin VB.CommandButton Menu_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Exit"
         Height          =   615
         Index           =   4
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5520
         Width           =   3135
      End
      Begin VB.CommandButton Menu_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Options"
         Height          =   615
         Index           =   3
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4920
         Width           =   3135
      End
      Begin VB.CommandButton Menu_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Help"
         Height          =   615
         Index           =   5
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4320
         Width           =   3135
      End
      Begin VB.CommandButton Menu_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Load Saved Campaign"
         Height          =   615
         Index           =   2
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3720
         Width           =   3135
      End
      Begin VB.CommandButton Menu_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Continue"
         Height          =   615
         Index           =   1
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3120
         Width           =   3135
      End
      Begin VB.CommandButton Menu_Command1 
         BackColor       =   &H0000C000&
         Caption         =   "New Campaign"
         Height          =   615
         Index           =   0
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "A Year 10 programming project to demostrate my capabilities in programming in Visual Basic 6.0"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   3120
         TabIndex        =   44
         Top             =   6600
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "For any queries, comments, suggestions, ideas or problems e-mail at: Beerboy160@Hotmail.com"
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
         Height          =   495
         Left            =   240
         TabIndex        =   43
         Top             =   6120
         Width           =   6135
      End
      Begin VB.Image Image2 
         Height          =   780
         Left            =   240
         Picture         =   "Game.frx":2935E
         Top             =   6600
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         Caption         =   "BETA Freeware Version 1.2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   5640
         Width           =   6015
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   5450
         Picture         =   "Game.frx":30490
         Top             =   6600
         Width           =   930
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "'The Hunt for Osama' by James Beer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   5280
         Width           =   6015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Operation Clearout I:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   6015
      End
      Begin VB.Image Image3 
         Height          =   1950
         Left            =   6480
         Picture         =   "Game.frx":32F6A
         Stretch         =   -1  'True
         Top             =   480
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   4140
         Left            =   240
         Picture         =   "Game.frx":3390A
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   6015
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   6975
         Left            =   120
         Top             =   480
         Width           =   6255
      End
   End
   Begin VB.PictureBox Game_Screen 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   0
      ScaleHeight     =   569
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   8145
         Left            =   0
         ScaleHeight     =   543
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   664
         TabIndex        =   1
         Top             =   0
         Width           =   9960
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Note: I structured the code the way I did primarily on
'Speed Optimization, because of this it runs 6x faster
'than before. Set Compile Options to native code and enable
'Optimize for Speed, and Compile to EXE
'to get best performance.

'PIII or Greater Required!


Option Explicit

Dim Quake As Boolean 'Earthquake Trigger Boolean Value
Dim QuakeTime As Long 'Earthquake time
Dim Shake_X As Long 'Shake at Offset X
Dim Shake_Y As Long 'Shake at Offset Y

Dim Game_Delay As Long 'A Setting to synchonize timers to FPS
Dim KillScore As Long 'Score
Dim Held As Boolean 'If JetPack Button Held

Dim Set_FPS As Long
Dim FPS As Long 'Frames Per Second

Dim CreditFrame As Long 'Credit Animation Picture
Dim Msg_Type As Long 'For Custom-made Msgbox Interface

Dim WinTime As Long 'Delay before Win screen is shown
Dim DeadTime As Long 'Delay before Dead Screen is shown
Dim Run As Boolean 'Game Engine is running


Private Function DoLoop()
'The Core Loop for my Game
Dim i As Long
Dim j As Long
Dim a As Long
Dim c As Long
Dim n As Long
Dim w As Long

Dim FC As Long

If LevelSel <= 1 Then KillScore = 0

ScoreTab.Visible = False
MissionTime = 0

If Cheats(1) = True Then
MainChar.Weapons(0).Have = True
MainChar.Weapons(1).Have = True: MainChar.Weapons(1).Ammo = 200
MainChar.Weapons(2).Have = True: MainChar.Weapons(2).Ammo = 500
MainChar.Weapons(3).Have = True: MainChar.Weapons(3).Ammo = 500
MainChar.Weapons(4).Have = True: MainChar.Weapons(4).Ammo = 1000
MainChar.Weapons(5).Have = True: MainChar.Weapons(5).Ammo = 50
MainChar.Weapons(6).Have = True: MainChar.Weapons(6).Ammo = 50
MainChar.Weapons(7).Have = True: MainChar.Weapons(7).Ammo = 10
MainChar.Weapons(8).Have = True: MainChar.Weapons(8).Ammo = 1000
MainChar.Armor = 100
End If

If MainChar.Health > 0 Then
MainChar.AnimCount = 0
MainChar.AnimFrame = 0
End If

Randomize Timer
Do Until Run = False

FPS = FPS + 1

If Quake = True Then
QuakeTime = QuakeTime + 1
If QuakeTime > 200 Then Quake = False
If FPS Mod 2 = 0 Then
Shake_X = Int(Rnd * ((200 - QuakeTime) \ 4))
Shake_Y = Int(Rnd * ((200 - QuakeTime) \ 4))
End If

Game_Offset_X = MainChar.X - (Picture1.Width \ 2) - Shake_X
If MainChar.Y <= Max_Y_Depth - MainChar.Height Then Game_Offset_Y = MainChar.Y - (Picture1.Height \ 2) - Shake_Y
If Shake_X > 0 Then Shake_X = -Shake_X
If Shake_Y > 0 Then Shake_Y = -Shake_Y
Else
Game_Offset_X = MainChar.X - (Picture1.Width \ 2)
If MainChar.Y <= Max_Y_Depth Then Game_Offset_Y = MainChar.Y - (Picture1.Height \ 2)
End If

BitBlt Picture1.Hdc, 0, 0, 664, 543, Form2.BG(BGSet).Hdc, 0, 0, vbSrcCopy

Game_Delay = Game_Delay + 1

a = 0
If Cheats(0) = True And NukeSet = False Or _
   Cheats(0) = True And Cheats(5) = True Then
MainChar.Health = 100
MainChar.Armor = 100
MainChar.Squished = False
If MainChar.Y < Max_Y_Depth Then
MainChar.AnimCount = 0
MainChar.AnimFrame = 0
DeadTime = 0
End If
End If

    If Cheats(2) = True Then
        For i = 1 To UBound(MainChar.Weapons)
        MainChar.Weapons(i).Ammo = 32767
        Next i
    End If

If MainChar.Y > Max_Y_Depth - MainChar.Height Then MainChar.Health = 0

For i = 0 To 9
If MainChar.Weapon_TimeOut(i) > 0 Then
    MainChar.Weapon_TimeOut(i) = MainChar.Weapon_TimeOut(i) - 1
End If
Next i

If Nuking = False Then
GroundDraw
Else
Quake = True
QuakeTime = 0
Picture1.Cls
End If
WallProcess MainChar.X, MainChar.Y, MainChar.Width, MainChar.Height, MainChar.MoveSpeed, MainChar.GravityForce
UpdateBullet
Swicthes
Effects

If NukeSet = True And Nuking = True Then
NukeTime = NukeTime + 1
Picture1.BackColor = RGB(NukeTime * 15, NukeTime * 15, NukeTime * 15)
End If

If NukeSet = True Or AI(OsamaID).Health <= 0 Then OsamaTaunt = False
If OsamaTaunt = True Then CreateMsg "OSAMA_DEATH", OsamaID

Objectives

BitBlt Picture1.Hdc, MissionTime_HUD_X, MissionTime_HUD_Y, 24, 24, Form2.HUD(0).Hdc, 24 * 4, 96, vbSrcAnd
BitBlt Picture1.Hdc, MissionTime_HUD_X, MissionTime_HUD_Y, 24, 24, Form2.HUD(1).Hdc, 24 * 4, 96, vbSrcPaint
Picture1.CurrentX = MissionTime_HUD_X + 26
Picture1.CurrentY = MissionTime_HUD_Y
Picture1.ForeColor = vbGreen
Picture1.Print MissionTime \ 60 Mod 60 & ":" & IIf(MissionTime Mod 60 >= 10, MissionTime Mod 60, "0" & (MissionTime Mod 60))

If Game_Delay >= Set_Game_Delay Then
If Run = True Then MissionTime = MissionTime + 1
Game_Delay = 0
End If

If NukeTime = 20 Then Nuking = False

If MainChar.Armor < 0 Then MainChar.Armor = 0

If MainChar.Squished = False Then
MainChar.Y = MainChar.Y + MainChar.GravityForce
    If MainChar.GravityForce < 64 And MainChar.Ground = False Then
        MainChar.GravityForce = MainChar.GravityForce + 1
    End If
End If

If MainChar.Health <= 0 Then a = 0: Picture1.ForeColor = RGB(200, 200, 200)
If MainChar.Health > 0 Then a = 1: Picture1.ForeColor = vbRed
If MainChar.Health >= 30 Then a = 2: Picture1.ForeColor = &H80FF&
If MainChar.Health >= 50 Then a = 3: Picture1.ForeColor = vbYellow
If MainChar.Health >= 80 Then a = 4: Picture1.ForeColor = vbGreen

BitBlt Picture1.Hdc, Health_HUD_X, Health_HUD_Y, 24, 24, Form2.HUD(0).Hdc, 24 * a, 0, vbSrcAnd
BitBlt Picture1.Hdc, Health_HUD_X, Health_HUD_Y, 24, 24, Form2.HUD(1).Hdc, 24 * a, 0, vbSrcPaint
Picture1.CurrentX = Health_HUD_X + 26
Picture1.CurrentY = Health_HUD_Y
Picture1.Print CCur(MainChar.Health) & "%"

If MainChar.Armor <= 0 Then a = 0: Picture1.ForeColor = RGB(200, 200, 200)
If MainChar.Armor > 0 Then a = 1: Picture1.ForeColor = vbRed
If MainChar.Armor >= 30 Then a = 2: Picture1.ForeColor = &H80FF&
If MainChar.Armor >= 50 Then a = 3: Picture1.ForeColor = vbYellow
If MainChar.Armor >= 80 Then a = 4: Picture1.ForeColor = vbGreen

BitBlt Picture1.Hdc, Armor_HUD_X, Armor_HUD_Y, 24, 24, Form2.HUD(0).Hdc, 24 * a, 24, vbSrcAnd
BitBlt Picture1.Hdc, Armor_HUD_X, Armor_HUD_Y, 24, 24, Form2.HUD(1).Hdc, 24 * a, 24, vbSrcPaint
Picture1.CurrentX = Armor_HUD_X + 26
Picture1.CurrentY = Armor_HUD_Y
Picture1.Print CCur(MainChar.Armor) & "%"

Picture1.CurrentX = TalibanKills_HUD_X + 26
Picture1.CurrentY = TalibanKills_HUD_Y
Picture1.ForeColor = vbWhite

a = 0
If KilledCount >= 1 Then a = 1
If KilledCount >= (UBound(AI) * 0.25) Then a = 2
If KilledCount >= (UBound(AI) * 0.5) Then a = 3
If KilledCount >= (UBound(AI) * 0.75) Then a = 4

BitBlt Picture1.Hdc, TalibanKills_HUD_X, TalibanKills_HUD_Y, 24, 24, Form2.HUD(0).Hdc, 24 * a, 72, vbSrcAnd
BitBlt Picture1.Hdc, TalibanKills_HUD_X, TalibanKills_HUD_Y, 24, 24, Form2.HUD(1).Hdc, 24 * a, 72, vbSrcPaint
Picture1.Print KillScore

If MainChar.ItemSel = 0 Then
If MainChar.JetPack_Have = True Then
    Picture1.CurrentX = 582
    Picture1.CurrentY = 64
    BitBlt Picture1.Hdc, 548, 64, 32, 32, Form2.ItemIcons.Hdc, 0, 0, vbSrcCopy
    If MainChar.JetPack_Fuel < 25 Then Picture1.ForeColor = vbRed Else Picture1.ForeColor = vbGreen
    Picture1.Print CLng(MainChar.JetPack_Fuel) & "%"
End If
End If


Picture1.ForeColor = RGB(200, 200, 200)

Picture1.CurrentY = 35
Picture1.CurrentX = 548

If MainChar.Gun_Tag >= 1 And _
   MainChar.Gun_Tag <= 8 Then

        If MainChar.Gun_Tag > 4 Then w = 5
        If MainChar.Gun_Tag = 8 Then w = 3
        If MainChar.Gun_Tag = 4 Then w = 4
        If MainChar.Gun_Tag < 4 Then w = MainChar.Gun_Tag - 1
        BitBlt Picture1.Hdc, 548, 2, 64, 32, Form2.Items(0).Hdc, 0, 32 * w, vbSrcCopy
        
        n = 0
        If MainChar.Gun_Tag = 5 Then n = 0
        If MainChar.Gun_Tag = 6 Then n = 1
        If MainChar.Gun_Tag = 7 Then n = 2
        
        BitBlt Picture1.Hdc, 610, 2, 32, 32, Form2.Items(0).Hdc, 64, 32 * (w + n), vbSrcCopy
        
        Picture1.Print MainChar.Weapons(MainChar.Gun_Tag).Ammo

    If MainChar.Gun_Tag = 7 Then
        Picture1.Line (548, 34)-(640, 38), vbBlack, BF
        c = (93 / MainChar.Weapons(MainChar.Gun_Tag).TimeOut_Rate) * (MainChar.Weapons(MainChar.Gun_Tag).TimeOut_Rate - MainChar.Weapon_TimeOut(MainChar.Gun_Tag))
        a = vbRed
        If c >= (92 * 0.33) Then a = vbYellow
        If c >= (92 * 0.66) Then a = vbGreen
        Picture1.Line (548, 34)-(548 + c, 38), a, BF
    End If
    
End If
Picture1.ForeColor = vbBlue
Picture1.CurrentX = 2
Picture1.CurrentY = 520
Picture1.Print "FPS: " & Set_FPS
If MainChar.Health > 0 Then KeyCommands

DoEvents

If MainChar.AnimFrame >= 8 And DeadTime >= 75 Then
Msg_Door.Enabled = False
MainChar.Squished = False
SetMusicType "FAIL_LEVEL", True
Menu_Msg(0).Visible = True
Menu_Msg(1).Visible = True
Menu_Msg(2).Visible = False
DeadTime = 0
Run = False
StartGame = False
Msg_Type = 6
Msg_DoorOpen = True
Msg_LevelName = "Level Terminated"
Msg(0).Tag = "KILL IN ACTION - MISSION FAILED" & Chr(13) & "Play Again?"
Menu_Msg(0).Caption = "Yes": Menu_Msg(0).Visible = True
Menu_Msg(1).Caption = "No": Menu_Msg(1).Visible = True
Msg_Door.Enabled = True
End If

Loop

End Function

Private Function CheckCollision(ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal h As Long, ByVal MS As Long, ByVal G As Long, Target As String, Optional Selected As Long)
'by using byval varibles for the regions in IF statements
'it really boosts the game's speed by 50%!
'The AI Collision Detection

Dim i As Long
Dim j As Long
Dim k As Long
Dim Z As Long

For j = 1 To UBound(Wall)
If X >= Wall(j).X - w - MS And _
   X <= Wall(j).X + Wall(j).Width + MS Then
ReDim Preserve WallReg(UBound(WallReg) + 1)
WallReg(UBound(WallReg)) = j
End If
Next j

For k = 1 To UBound(WallReg)
j = WallReg(k)
If Y <= Wall(j).Y + Wall(j).Height And _
   Y >= Wall(j).Y - Wall(j).Height - h And _
   AI(Selected).Squished = False Then

If AI(Selected).Contact = False Then
AI(Selected).Ground = False
End If

If Y + h > Wall(j).Y Then
If Wall(j).Hazardous = True Then
AI(Selected).Health = AI(Selected).Health - (Wall(j).Damage / ((AI(Selected).Armor \ 33) + 1))
AI(Selected).Armor = AI(Selected).Armor - Wall(j).Damage
End If
End If

    If Wall(j).Water = False Then

If X + w >= (Wall(j).X) And _
   X + w <= (Wall(j).X + MS) And _
   Y + h > (Wall(j).Y) And _
   Y < (Wall(j).Y + Wall(j).Height) Then
   AI(Selected).Key_Lock(1) = True
Z = (AI(Selected).X + AI(Selected).Width) - Wall(j).X
AI(Selected).X = AI(Selected).X - Z
End If

If X >= (Wall(j).X + Wall(j).Width - MS) And _
   X <= (Wall(j).X + Wall(j).Width) And _
   Y + h > (Wall(j).Y) And _
   Y < (Wall(j).Y + Wall(j).Height) Then
   AI(Selected).Key_Lock(0) = True
Z = (AI(Selected).X) - (Wall(j).X + Wall(j).Width)
AI(Selected).X = AI(Selected).X - Z
End If

If (Y + h) >= Wall(j).Y - G And _
   (Y) <= (Wall(j).Y + (Wall(j).Height / 2)) And _
   (X + w) > Wall(j).X + MS And _
   (X) < (Wall(j).X + Wall(j).Width) - MS And _
   AI(Selected).Ground = False Then
Z = (AI(Selected).Y + AI(Selected).Height) - Wall(j).Y
AI(Selected).Y = AI(Selected).Y - Z
AI(Selected).Ground = True

If AI(Selected).GravityForce > 24 Then
AI(Selected).Health = AI(Selected).Health - ((AI(Selected).GravityForce - 24) * ((AI(Selected).GravityForce - 24) \ 5))
End If

AI(Selected).GravityForce = 0

AI(Selected).Contact = True
End If

If (Y) >= (Wall(j).Y - G) And _
   (Y) <= (Wall(j).Y + Wall(j).Height) - G And _
   (X + w) > Wall(j).X + MS And _
   (X) < (Wall(j).X + Wall(j).Width) - MS And _
   G < 1 Then
Z = Abs(AI(Selected).GravityForce) + 4
AI(Selected).Y = AI(Selected).Y + Z
AI(Selected).GravityForce = 1
AI(Selected).Ground = True
End If

    Else
    If G > 8 Then AI(Selected).GravityForce = 8
    End If
End If
Next k
End Function


Private Function WallProcess(X As Long, Y As Long, w As Long, h As Long, ByVal MS As Long, G As Single)
'Character Collision Detection
Dim i As Long
Dim j As Long
Dim k As Long
Dim Z As Long

MainChar.Key_Lock(0) = False
MainChar.Key_Lock(1) = False

ReDim WallReg(0)

For i = 1 To UBound(Wall)
If X >= Wall(i).X - w - MS And _
   X <= Wall(i).X + Wall(i).Width + MS And _
   MainChar.Squished = False Then
ReDim Preserve WallReg(UBound(WallReg) + 1)
WallReg(UBound(WallReg)) = i
End If
Next i

For k = 1 To UBound(WallReg)
i = WallReg(k)
If MainChar.Contact = False Then
MainChar.Ground = False
End If

If Y <= Wall(i).Y + Wall(i).Height And _
   Y >= Wall(i).Y - Wall(i).Height - h And _
   MainChar.Squished = False Then

If Y + h > Wall(i).Y Then
If Wall(i).Hazardous = True Then
MainChar.Health = MainChar.Health - (Wall(i).Damage / ((MainChar.Armor \ 33) + 1))
MainChar.Armor = MainChar.Armor - Wall(i).Damage
End If
End If

    If Wall(i).Water = False Then

If X + w >= (Wall(i).X) And _
   X + w <= (Wall(i).X + MS) And _
   (Y + h) > (Wall(i).Y) And _
   Y < (Wall(i).Y + Wall(i).Height) Then
   MainChar.Key_Lock(1) = True
Z = (MainChar.X + MainChar.Width) - Wall(i).X
MainChar.X = MainChar.X - Z
End If

If X >= (Wall(i).X + Wall(i).Width - MS) And _
   X <= (Wall(i).X + Wall(i).Width) And _
   (Y + h) > (Wall(i).Y) And _
   Y < (Wall(i).Y + Wall(i).Height) Then
   MainChar.Key_Lock(0) = True
Z = (MainChar.X) - (Wall(i).X + Wall(i).Width)
MainChar.X = MainChar.X - Z
End If

If (Y + h) >= Wall(i).Y - G And _
   (Y) <= (Wall(i).Y + (Wall(i).Height / 2)) And _
   (X + w) > Wall(i).X And _
   (X) < (Wall(i).X + Wall(i).Width) And _
   MainChar.GravityForce >= 0 Then
Z = (MainChar.Y + MainChar.Height) - Wall(i).Y
MainChar.Y = MainChar.Y - Z
MainChar.Ground = True

If MainChar.GravityForce > 24 Then
MainChar.Health = MainChar.Health - ((MainChar.GravityForce - 24) * ((MainChar.GravityForce - 24) \ 5))
End If

MainChar.GravityForce = 0

MainChar.Contact = True
End If

If (Y) >= (Wall(i).Y - G) And _
   (Y) <= (Wall(i).Y + Wall(i).Height) - G And _
   (X + w) > Wall(i).X + MS And _
   (X) < (Wall(i).X + Wall(i).Width) - MS And _
   MainChar.GravityForce < 1 Then
Z = Abs(MainChar.GravityForce) + 4
MainChar.Y = MainChar.Y + Z
MainChar.GravityForce = 1
End If

    Else
    If Y + h > Wall(i).Y Then
    If MainChar.GravityForce > 8 Then MainChar.GravityForce = 8
    End If
    End If
End If

Next k

MainChar.Contact = False

End Function

Private Function GroundDraw()
'A key process to draw the picture and process some
'of the collision detection rountines
Dim i As Long
Dim j As Long
Dim k As Long
Dim c As Long
Dim X As Long

If NukeSet = False Then
ReDim BGReg(0)

For i = 1 To UBound(Background)
Background(i).Count = Background(i).Count + 1
If Background(i).Count >= Background(i).MaxCount Then Background(i).AnimFrame = Background(i).AnimFrame + 1: Background(i).Count = 0
If Background(i).AnimFrame > Background(i).MaxFrame Then Background(i).AnimFrame = 0: Background(i).Count = 0

If Background(i).X - Game_Offset_X >= -Background(i).Width And _
   Background(i).X - Game_Offset_X <= Picture1.Width And _
   Background(i).Y - Game_Offset_Y >= -Background(i).Height And _
   Background(i).Y - Game_Offset_Y <= Picture1.Height And _
   Background(i).Use = True Then
BitBlt Picture1.Hdc, Background(i).X - Game_Offset_X, Background(i).Y - Game_Offset_Y, Background(i).Width, Background(i).Height, Form2.Picture2.Hdc, Background(i).PX + (Background(i).Width * Background(i).AnimFrame), Background(i).PY, vbSrcCopy
End If
Next i

For i = 1 To UBound(BackProp)
BackProp(i).Count = BackProp(i).Count + 1
If BackProp(i).Count >= BackProp(i).MaxCount Then BackProp(i).AnimFrame = BackProp(i).AnimFrame + 1: BackProp(i).Count = 0
If BackProp(i).AnimFrame > BackProp(i).MaxFrame Then BackProp(i).AnimFrame = 0: BackProp(i).Count = 0
If BackProp(i).X - Game_Offset_X >= -BackProp(i).Width And _
   BackProp(i).X - Game_Offset_X <= Picture1.Width And _
   BackProp(i).Y - Game_Offset_Y >= -BackProp(i).Height And _
   BackProp(i).Y - Game_Offset_Y <= Picture1.Height And _
   BackProp(i).Use = True Then
BitBlt Picture1.Hdc, BackProp(i).X - Game_Offset_X, BackProp(i).Y - Game_Offset_Y, BackProp(i).Width, BackProp(i).Height, Form2.Picture2.Hdc, BackProp(i).PX + (BackProp(i).Width * BackProp(i).AnimFrame), BackProp(i).PY, vbSrcAnd
BitBlt Picture1.Hdc, BackProp(i).X - Game_Offset_X, BackProp(i).Y - Game_Offset_Y, BackProp(i).Width, BackProp(i).Height, Form2.Picture2.Hdc, BackProp(i).PX + (BackProp(i).Width * BackProp(i).AnimFrame), BackProp(i).PY + BackProp(i).Height, vbSrcPaint
End If
Next i
End If

For i = 1 To UBound(Swicth)
If Swicth(i).EnterSet = False Then
If Swicth(i).On = False Then
BitBlt Picture1.Hdc, Swicth(i).X - Game_Offset_X, Swicth(i).Y - Game_Offset_Y, Swicth(i).Width, Swicth(i).Height, Form2.Picture2.Hdc, Swicth(i).PX(0), Swicth(i).PY(0), vbSrcCopy
Else
BitBlt Picture1.Hdc, Swicth(i).X - Game_Offset_X, Swicth(i).Y - Game_Offset_Y, Swicth(i).Width, Swicth(i).Height, Form2.Picture2.Hdc, Swicth(i).PX(1), Swicth(i).PY(1), vbSrcCopy
End If
End If
Next i

For j = 1 To UBound(Wall)
Wall(j).Count = Wall(j).Count + 1
If Wall(j).Count >= Wall(j).MaxCount Then Wall(j).AnimFrame = Wall(j).AnimFrame + 1: Wall(j).Count = 0
If Wall(j).AnimFrame > Wall(j).MaxFrame Then Wall(j).AnimFrame = 0: Wall(j).Count = 0
If Wall(j).X - Game_Offset_X >= -Wall(j).Width And _
   Wall(j).X - Game_Offset_X <= Picture1.Width And _
   Wall(j).Y - Game_Offset_Y >= -Wall(j).Height And _
   Wall(j).Y - Game_Offset_Y <= Picture1.Height And _
   Wall(j).Use = True Then
BitBlt Picture1.Hdc, Wall(j).X - Game_Offset_X, Wall(j).Y - Game_Offset_Y, Wall(j).Width, Wall(j).Height, Form2.Picture2.Hdc, Wall(j).PX + (Wall(j).Width * Wall(j).AnimFrame), Wall(j).PY, vbSrcCopy
End If
Next j


For i = 1 To UBound(AI)
ReDim WallReg(0)
CheckCollision AI(i).X, AI(i).Y, AI(i).Width, AI(i).Height, AI(i).MoveSpeed, AI(i).GravityForce, "AI", i

If AI(i).X - Game_Offset_X > -AI(i).Width And _
   AI(i).X - Game_Offset_X < Picture1.Width And _
   AI(i).Y - Game_Offset_Y > -AI(i).Height And _
   AI(i).Y - Game_Offset_Y < Picture1.Height Then

    If AI(i).AI_Type = "TALIBAN_1" Then
        If AI(i).Health > 0 Then
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 34, Form2.Baddie1(0).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (34 * 3) * AI(i).Gun_Draw_Tag + (34 * AI(i).Walk_Tag), vbSrcAnd
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 34, Form2.Baddie1(1).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (34 * 3) * AI(i).Gun_Draw_Tag + (34 * AI(i).Walk_Tag), vbSrcPaint
        Else
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 34, Form2.Baddie1(0).Hdc, 32 * AI(i).Direction_Tag, (34 * AI(i).AnimFrame), vbSrcAnd
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 34, Form2.Baddie1(1).Hdc, 32 * AI(i).Direction_Tag, (34 * AI(i).AnimFrame), vbSrcPaint
        End If
    ElseIf AI(i).AI_Type = "TALIBAN_2" Then
        If AI(i).Health > 0 Then
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 34, Form2.Baddie2(0).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (34 * 3) * AI(i).Gun_Draw_Tag + (34 * AI(i).Walk_Tag), vbSrcAnd
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 34, Form2.Baddie2(1).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (34 * 3) * AI(i).Gun_Draw_Tag + (34 * AI(i).Walk_Tag), vbSrcPaint
        Else
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 34, Form2.Baddie2(0).Hdc, 32 * AI(i).Direction_Tag, (34 * AI(i).AnimFrame), vbSrcAnd
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 34, Form2.Baddie2(1).Hdc, 32 * AI(i).Direction_Tag, (34 * AI(i).AnimFrame), vbSrcPaint
        End If
    ElseIf AI(i).AI_Type = "TALIBAN_3" Then
        If AI(i).Health > 0 Then
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 34, Form2.Baddie3(0).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (34 * 3) * AI(i).Gun_Draw_Tag + (34 * AI(i).Walk_Tag), vbSrcAnd
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 34, Form2.Baddie3(1).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (34 * 3) * AI(i).Gun_Draw_Tag + (34 * AI(i).Walk_Tag), vbSrcPaint
        Else
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 34, Form2.Baddie3(0).Hdc, 32 * AI(i).Direction_Tag, (34 * AI(i).AnimFrame), vbSrcAnd
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 34, Form2.Baddie3(1).Hdc, 32 * AI(i).Direction_Tag, (34 * AI(i).AnimFrame), vbSrcPaint
        End If
    ElseIf AI(i).AI_Type = "TALIBAN_4" Then
        If AI(i).Health > 0 Then
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 34, Form2.Baddie4(0).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (34 * 3) * AI(i).Gun_Draw_Tag + (34 * AI(i).Walk_Tag), vbSrcAnd
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 34, Form2.Baddie4(1).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (34 * 3) * AI(i).Gun_Draw_Tag + (34 * AI(i).Walk_Tag), vbSrcPaint
        Else
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 34, Form2.Baddie4(0).Hdc, 32 * AI(i).Direction_Tag, (34 * AI(i).AnimFrame), vbSrcAnd
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 34, Form2.Baddie4(1).Hdc, 32 * AI(i).Direction_Tag, (34 * AI(i).AnimFrame), vbSrcPaint
        End If
    ElseIf AI(i).AI_Type = "TALIBAN_5" Then
        If AI(i).Health > 0 Then
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 34, Form2.Baddie5(0).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (34 * 3) * AI(i).Gun_Draw_Tag + (34 * AI(i).Walk_Tag), vbSrcAnd
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 34, Form2.Baddie5(1).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (34 * 3) * AI(i).Gun_Draw_Tag + (34 * AI(i).Walk_Tag), vbSrcPaint
        Else
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 34, Form2.Baddie5(0).Hdc, 32 * AI(i).Direction_Tag, (34 * AI(i).AnimFrame), vbSrcAnd
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 34, Form2.Baddie5(1).Hdc, 32 * AI(i).Direction_Tag, (34 * AI(i).AnimFrame), vbSrcPaint
        End If
    ElseIf AI(i).AI_Type = "TALIBAN_BOSS1" Then
        If AI(i).Health > 0 Then
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 45, 45, Form2.BaddieBoss1(0).Hdc, 90 * AI(i).Direction_Tag + (45 * AI(i).Attack_Tag), (45 * 3) * AI(i).Gun_Draw_Tag + (45 * AI(i).Walk_Tag), vbSrcAnd
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 45, 45, Form2.BaddieBoss1(1).Hdc, 90 * AI(i).Direction_Tag + (45 * AI(i).Attack_Tag), (45 * 3) * AI(i).Gun_Draw_Tag + (45 * AI(i).Walk_Tag), vbSrcPaint
        Else
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 45, 45, Form2.BaddieBoss1(0).Hdc, 45 * AI(i).Direction_Tag, (45 * AI(i).AnimFrame), vbSrcAnd
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 45, 45, Form2.BaddieBoss1(1).Hdc, 45 * AI(i).Direction_Tag, (45 * AI(i).AnimFrame), vbSrcPaint
        End If
    ElseIf AI(i).AI_Type = "TALIBAN_BOSS_OSAMA" Then
        If AI(i).Health > 0 Then
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 53, Form2.Osama(0).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (53 * 3) * AI(i).Gun_Draw_Tag + (53 * AI(i).Walk_Tag), vbSrcAnd
            BitBlt Picture1.Hdc, (AI(i).X - Game_Offset_X), (AI(i).Y - Game_Offset_Y), 32, 53, Form2.Osama(1).Hdc, 64 * AI(i).Direction_Tag + (32 * AI(i).Attack_Tag), (53 * 3) * AI(i).Gun_Draw_Tag + (53 * AI(i).Walk_Tag), vbSrcPaint
        Else
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 53, Form2.Osama(0).Hdc, 32 * AI(i).Direction_Tag, (53 * AI(i).AnimFrame), vbSrcAnd
            BitBlt Picture1.Hdc, AI(i).X - Game_Offset_X, AI(i).Y - Game_Offset_Y, 32, 53, Form2.Osama(1).Hdc, 32 * AI(i).Direction_Tag, (53 * AI(i).AnimFrame), vbSrcPaint
        End If
End If

If AI(i).Health <= 0 And AI(i).AnimFrame < 11 Then
Picture1.CurrentX = AI(i).X - Game_Offset_X
Picture1.CurrentY = AI(i).Y - Game_Offset_Y - (AI(i).AnimFrame * 2)
Picture1.FontSize = 9
Picture1.ForeColor = vbWhite
Picture1.Print (AI(i).MaxHealth * 10) + ((AI(i).MaxArmor \ 2) * 10)
Picture1.FontSize = 14
End If

If AI(i).Armor < 0 Then AI(i).Armor = 0

'If AI(i).Boss = True Then
Picture1.Line (AI(i).X - Game_Offset_X, AI(i).Y + AI(i).Height + 4 - Game_Offset_Y)-(AI(i).X + AI(i).Width - Game_Offset_X, AI(i).Y + AI(i).Height + 8 - Game_Offset_Y), 0, BF
If AI(i).Health > 0 Then c = vbRed
If AI(i).Health > (AI(i).MaxHealth / 3) Then c = vbYellow
If AI(i).Health > (AI(i).MaxHealth / 3) * 2 Then c = vbGreen
If AI(i).Health > 0 Then Picture1.Line (AI(i).X - Game_Offset_X, AI(i).Y + AI(i).Height + 4 - Game_Offset_Y)-(AI(i).X + (AI(i).Health * (AI(i).Width / AI(i).MaxHealth)) - Game_Offset_X, AI(i).Y + AI(i).Height + 8 - Game_Offset_Y), c, BF

Picture1.Line (AI(i).X - Game_Offset_X, AI(i).Y + AI(i).Height + 8 - Game_Offset_Y)-(AI(i).X + AI(i).Width - Game_Offset_X, AI(i).Y + AI(i).Height + 12 - Game_Offset_Y), 0, BF
c = vbBlue
If AI(i).Armor > 0 Then Picture1.Line (AI(i).X - Game_Offset_X, AI(i).Y + AI(i).Height + 8 - Game_Offset_Y)-(AI(i).X + (AI(i).Armor * (AI(i).Width / AI(i).MaxArmor)) - Game_Offset_X, AI(i).Y + AI(i).Height + 12 - Game_Offset_Y), c, BF
'End If
End If

If AI(i).Health <= 0 And AI(i).AnimFrame < 10 And AI(i).AnimCount = 0 Then
KillScore = KillScore + CLng((AI(i).MaxHealth) + (AI(i).MaxArmor \ 2))
End If

    If AI(i).Weapon_TimeOut > 0 Then
       AI(i).Weapon_TimeOut = AI(i).Weapon_TimeOut - 1
   End If

    For k = 0 To 3
        AI(i).Keys(k) = False
    Next k

    If AI(i).Health > 0 Then
        AI(i).Attack_Tag = 0
    If AI(i).Active = True Then AI_System i, AI(i).AI_Type
    Else
        If AI(i).AnimFrame = 11 Then
        AI(i).AnimFrame = 11
            Else
            AI(i).AnimCount = AI(i).AnimCount + 1
            If AI(i).AnimCount >= 2 Then
            AI(i).AnimFrame = AI(i).AnimFrame + 1
            AI(i).AnimCount = 0
            End If
    End If
End If

    AI_Movement i

If AI(i).Squished = False Then AI(i).Y = AI(i).Y + AI(i).GravityForce
    If AI(i).GravityForce < 64 And AI(i).Ground = False Then
        AI(i).GravityForce = AI(i).GravityForce + 1
    End If

AI(i).Contact = False
Next i

For i = 1 To UBound(Item)
Items Item(i).X, Item(i).Y, Item(i).Width, Item(i).Height, Item(i).Item, i
Next i

If MainChar.Health > 0 Then
    BitBlt Picture1.Hdc, (MainChar.X - Game_Offset_X), (MainChar.Y - Game_Offset_Y), 32, 34, Form2.Picture1(0).Hdc, 64 * MainChar.Direction_Tag + (32 * MainChar.Attack_Tag), (34 * 3) * MainChar.Gun_Draw_Tag + (34 * MainChar.Walk_Tag), vbSrcAnd
    BitBlt Picture1.Hdc, (MainChar.X - Game_Offset_X), (MainChar.Y - Game_Offset_Y), 32, 34, Form2.Picture1(1).Hdc, 64 * MainChar.Direction_Tag + (32 * MainChar.Attack_Tag), (34 * 3) * MainChar.Gun_Draw_Tag + (34 * MainChar.Walk_Tag), vbSrcPaint
If MainChar.JetPack_Anim > 0 Then
    If MainChar.Direction_Tag = 0 Then
    BitBlt Picture1.Hdc, 4 + MainChar.X - Game_Offset_X, MainChar.Y - 4 + (MainChar.Height \ 2) - Game_Offset_Y, 8, 16, Form2.Picture3.Hdc, 8 * (MainChar.JetPack_Anim - 1), 0, vbSrcAnd
    BitBlt Picture1.Hdc, 4 + MainChar.X - Game_Offset_X, MainChar.Y - 4 + (MainChar.Height \ 2) - Game_Offset_Y, 8, 16, Form2.Picture3.Hdc, 24 + (8 * (MainChar.JetPack_Anim - 1)), 0, vbSrcPaint
    Else
    BitBlt Picture1.Hdc, 20 + MainChar.X - Game_Offset_X, MainChar.Y - 4 + (MainChar.Height \ 2) - Game_Offset_Y, 8, 16, Form2.Picture3.Hdc, 8 * (MainChar.JetPack_Anim - 1), 0, vbSrcAnd
    BitBlt Picture1.Hdc, 20 + MainChar.X - Game_Offset_X, MainChar.Y - 4 + (MainChar.Height \ 2) - Game_Offset_Y, 8, 16, Form2.Picture3.Hdc, 24 + (8 * (MainChar.JetPack_Anim - 1)), 0, vbSrcPaint
    End If
End If
Else
    DeadTime = DeadTime + 1
    If MainChar.AnimFrame = 8 Then
    MainChar.AnimFrame = 8
    BitBlt Picture1.Hdc, MainChar.X - Game_Offset_X, MainChar.Y - Game_Offset_Y, 32, 34, Form2.Picture1(0).Hdc, 32 * 4, (34 * 8), vbSrcAnd
    BitBlt Picture1.Hdc, MainChar.X - Game_Offset_X, MainChar.Y - Game_Offset_Y, 32, 34, Form2.Picture1(1).Hdc, 32 * 4, (34 * 8), vbSrcPaint
    Else
    MainChar.AnimCount = MainChar.AnimCount + 1
    If MainChar.AnimCount >= 3 Then
    MainChar.AnimFrame = MainChar.AnimFrame + 1
    MainChar.AnimCount = 0
    End If

    BitBlt Picture1.Hdc, MainChar.X - Game_Offset_X, MainChar.Y - Game_Offset_Y, 32, 34, Form2.Picture1(0).Hdc, 32 * (4 + MainChar.Direction_Tag), (34 * MainChar.AnimFrame), vbSrcAnd
    BitBlt Picture1.Hdc, MainChar.X - Game_Offset_X, MainChar.Y - Game_Offset_Y, 32, 34, Form2.Picture1(1).Hdc, 32 * (4 + MainChar.Direction_Tag), (34 * MainChar.AnimFrame), vbSrcPaint
    End If
End If

If NukeSet = False Then
For i = 1 To UBound(Foreground)
Foreground(i).Count = Foreground(i).Count + 1
If Foreground(i).Count >= Foreground(i).MaxCount Then Foreground(i).AnimFrame = Foreground(i).AnimFrame + 1: Foreground(i).Count = 0
If Foreground(i).AnimFrame > Foreground(i).MaxFrame Then Foreground(i).AnimFrame = 0: Foreground(i).Count = 0
If Foreground(i).X - Game_Offset_X >= -Foreground(i).Width And _
   Foreground(i).X - Game_Offset_X <= Picture1.Width And _
   Foreground(i).Y - Game_Offset_Y >= -Foreground(i).Height And _
   Foreground(i).Y - Game_Offset_Y <= Picture1.Height And _
   Foreground(i).Use = True Then
BitBlt Picture1.Hdc, Foreground(i).X - Game_Offset_X, Foreground(i).Y - Game_Offset_Y, Foreground(i).Width, Foreground(i).Height, Form2.Picture2.Hdc, Foreground(i).PX, Foreground(i).PY, vbSrcCopy
End If
Next i

For i = 1 To UBound(ForeProp)
ForeProp(i).Count = ForeProp(i).Count + 1
If ForeProp(i).Count >= ForeProp(i).MaxCount Then ForeProp(i).AnimFrame = ForeProp(i).AnimFrame + 1: ForeProp(i).Count = 0
If ForeProp(i).AnimFrame > ForeProp(i).MaxFrame Then ForeProp(i).AnimFrame = 0: ForeProp(i).Count = 0
If ForeProp(i).X - Game_Offset_X >= -ForeProp(i).Width And _
   ForeProp(i).X - Game_Offset_X <= Picture1.Width And _
   ForeProp(i).Y - Game_Offset_Y >= -ForeProp(i).Height And _
   ForeProp(i).Y - Game_Offset_Y <= Picture1.Height And _
   ForeProp(i).Use = True Then
BitBlt Picture1.Hdc, ForeProp(i).X - Game_Offset_X, ForeProp(i).Y - Game_Offset_Y, ForeProp(i).Width, ForeProp(i).Height, Form2.Picture2.Hdc, ForeProp(i).PX + (ForeProp(i).Width * ForeProp(i).AnimFrame), ForeProp(i).PY, vbSrcAnd
BitBlt Picture1.Hdc, ForeProp(i).X - Game_Offset_X, ForeProp(i).Y - Game_Offset_Y, ForeProp(i).Width, ForeProp(i).Height, Form2.Picture2.Hdc, ForeProp(i).PX + (ForeProp(i).Width * ForeProp(i).AnimFrame), ForeProp(i).PY + ForeProp(i).Height, vbSrcPaint
End If
Next i
End If
End Function

Private Function KeyCommands()
'The character Key controls
Dim i As Long

If UserInput = True Then

For i = 0 To 4
MainChar.Keys(i) = False
Next i

MainChar.Attack_Tag = 0
MainChar.Weapons(0).Ammo = 1

If GetAsyncKeyState(vbKeyLeft) < 0 And MainChar.Key_Lock(0) = False Then MainChar.Keys(0) = True
If GetAsyncKeyState(vbKeyRight) < 0 And MainChar.Key_Lock(1) = False Then MainChar.Keys(1) = True
If GetAsyncKeyState(vbKeyUp) < 0 Then MainChar.Keys(2) = True
If GetAsyncKeyState(vbKeyDown) < 0 Then MainChar.Keys(3) = True
If GetAsyncKeyState(vbKeyE) < 0 Then MainChar.Keys(4) = True

If GetAsyncKeyState(vbKeyX) < 0 And NukeSet = False And _
   Cheats(5) = True Then
Picture1.Cls
Nuking = True
NukeSet = True
SetDamage "MAINCHAR", 10000
For i = 0 To UBound(AI)
SetDamage "AI", 10000, i
Next i
If BGSet <> 5 Then BGSet = 2
End If

If GetAsyncKeyState(vbKeyReturn) = 0 Then Held = False
If GetAsyncKeyState(vbKeyReturn) < 0 And Held = False Then
Held = True
If MainChar.ItemSel = 0 Then
If MainChar.JetPack_Have = True Then
MainChar.JetPack_Mode = Not MainChar.JetPack_Mode
If MainChar.JetPack_Fuel <= 0 And MainChar.JetPack_Mode = True Then MainChar.JetPack_Mode = False
End If
End If
End If

If GetAsyncKeyState(vbKey1) < 0 And MainChar.Weapons(0).Have = True And MainChar.Gun_Tag <> 0 Then MainChar.Gun_Tag = 0: MainChar.Gun_Draw_Tag = 0
If GetAsyncKeyState(vbKey2) < 0 And MainChar.Weapons(1).Have = True And MainChar.Gun_Tag <> 1 Then MainChar.Gun_Tag = 1: MainChar.Gun_Draw_Tag = 1
If GetAsyncKeyState(vbKey3) < 0 And MainChar.Weapons(2).Have = True And MainChar.Gun_Tag <> 2 Then MainChar.Gun_Tag = 2: MainChar.Gun_Draw_Tag = 2
If GetAsyncKeyState(vbKey5) < 0 And MainChar.Weapons(3).Have = True And MainChar.Gun_Tag <> 3 Then MainChar.Gun_Tag = 3: MainChar.Gun_Draw_Tag = 3
If GetAsyncKeyState(vbKey6) < 0 And MainChar.Weapons(4).Have = True And MainChar.Gun_Tag <> 4 Then MainChar.Gun_Tag = 4: MainChar.Gun_Draw_Tag = 4
If GetAsyncKeyState(vbKey7) < 0 And MainChar.Weapons(5).Have = True And MainChar.Gun_Tag <> 5 Then MainChar.Gun_Tag = 5: MainChar.Gun_Draw_Tag = 5
If GetAsyncKeyState(vbKey8) < 0 And MainChar.Weapons(6).Have = True And MainChar.Gun_Tag <> 6 Then MainChar.Gun_Tag = 6: MainChar.Gun_Draw_Tag = 5
If GetAsyncKeyState(vbKey9) < 0 And MainChar.Weapons(7).Have = True And MainChar.Gun_Tag <> 7 Then MainChar.Gun_Tag = 7: MainChar.Gun_Draw_Tag = 5
If GetAsyncKeyState(vbKey4) < 0 And MainChar.Weapons(8).Have = True And MainChar.Gun_Tag <> 8 Then MainChar.Gun_Tag = 8: MainChar.Gun_Draw_Tag = 6

If GetAsyncKeyState(vbKeyControl) < 0 And MainChar.Weapon_TimeOut(MainChar.Gun_Tag) = 0 And _
   MainChar.Weapons(MainChar.Gun_Tag).Ammo > 0 Then MakeBullet

End If

Do Until MainChar.Weapons(MainChar.Gun_Tag).Ammo > 0
If MainChar.Gun_Tag = 0 Then Exit Do
If MainChar.Weapons(MainChar.Gun_Tag).Ammo <= 0 Then
MainChar.Gun_Tag = MainChar.Gun_Tag - 1

If MainChar.Gun_Tag > 7 Then
MainChar.Gun_Draw_Tag = 5
Else
MainChar.Gun_Draw_Tag = MainChar.Gun_Tag
End If

End If
Loop

If MainChar.Keys(0) = True And MainChar.Keys(1) = False Then
MainChar.X = MainChar.X - MainChar.MoveSpeed
MainChar.Direction_Tag = 1
End If

If MainChar.Keys(1) = True And MainChar.Keys(0) = False Then
MainChar.X = MainChar.X + MainChar.MoveSpeed
MainChar.Direction_Tag = 0
End If

If MainChar.Keys(0) = True And MainChar.Keys(1) = False Or _
   MainChar.Keys(0) = False And MainChar.Keys(1) = True Then
If MainChar.Ground = True Then
MainChar.Walk_Delay = MainChar.Walk_Delay - 1

If MainChar.Walk_Delay <= 0 Then

If MainChar.Walk_Tag = 1 Then
MainChar.Walk_Tag = 0
Else
MainChar.Walk_Tag = 1
End If

MainChar.Walk_Delay = 2

End If
End If

Else

MainChar.Walk_Tag = 2

End If

If Cheats(4) = True Then MainChar.JetPack_Fuel = 100: MainChar.JetPack_Have = True
If MainChar.JetPack_Mode = False Then
MainChar.JetPack_Anim = 0
        If MainChar.Keys(2) = True And MainChar.Ground = True Then
        MainChar.GravityForce = -15
        End If
Else

If MainChar.JetPack_Have = True Then
MainChar.JetPack_Anim = 1
    
If MainChar.GravityForce > 4 And MainChar.JetPack_Mode = True Then
    MainChar.GravityForce = MainChar.GravityForce - 2
    MainChar.JetPack_Anim = 2
    If MainChar.JetPack_Fuel > 0 Then MainChar.JetPack_Fuel = MainChar.JetPack_Fuel - 0.1
End If
    
    If MainChar.Keys(2) = True And MainChar.GravityForce > -14 And MainChar.JetPack_Fuel > 0 Then
    MainChar.GravityForce = MainChar.GravityForce - 2
    MainChar.JetPack_Anim = 3
    If MainChar.JetPack_Fuel > 0 Then MainChar.JetPack_Fuel = MainChar.JetPack_Fuel - 0.1
    End If
If MainChar.JetPack_Fuel <= 0 Then MainChar.JetPack_Mode = False
End If
End If

End Function

Private Function UpdateBullet()
Dim i As Long
Dim j As Long

For i = 0 To UBound(Bullet)

If Bullet(i).Used = True Then

If Bullet(i).BurnOut = True Then
Bullet(i).BurnOutTime = Bullet(i).BurnOutTime - 1
If Bullet(i).BurnOutTime <= 0 Then
Bullet(i).Used = False
End If
End If

For j = 1 To UBound(Wall)
If Wall(j).Water = False Then
If Bullet(i).X >= Wall(j).X - MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Width And _
   Bullet(i).X <= (Wall(j).X + Wall(j).Width) + MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Width And _
   Bullet(i).Y >= Wall(j).Y - 1 And _
   Bullet(i).Y <= (Wall(j).Y + Wall(j).Height) + 1 Then

If Bullet(i).Effect > 0 Then
SetEffect Bullet(i).Effect, Bullet(i).X, Bullet(i).Y, Bullet(i).From, i
Else
Bullet(i).Used = False
End If

End If
End If
Next j

For j = 0 To UBound(AI)
If AI(j).Health > 0 Then
If Bullet(i).X >= AI(j).X - Abs(Bullet(i).Speed * 2) And _
   Bullet(i).X <= AI(j).X + AI(j).Width + Abs(Bullet(i).Speed * 2) And _
   Bullet(i).Y >= AI(j).Y And _
   Bullet(i).Y <= AI(j).Y + AI(j).Height And _
   Bullet(i).From = 0 And Bullet(i).Used = True Then
   
If Bullet(i).Effect > 0 Then
SetEffect Bullet(i).Effect, AI(j).X + (AI(j).Width / 2), AI(j).Y + (AI(j).Height / 2), Bullet(i).From, i
Else
If AI(j).Armor < 0 Then AI(j).Armor = 0
SetDamage "AI", Bullet(i).Damage, j
Bullet(i).Used = False
End If
End If

End If
Next j

If Bullet(i).X >= MainChar.X - Abs(Bullet(i).Speed * 2) And _
   Bullet(i).X <= MainChar.X + MainChar.Width + Abs(Bullet(i).Speed * 2) And _
   Bullet(i).Y >= MainChar.Y And _
   Bullet(i).Y <= MainChar.Y + MainChar.Height And _
   Bullet(i).From = 1 And MainChar.Health > 0 And _
   Bullet(i).Used = True Then

If Bullet(i).Effect > 0 Then
SetEffect Bullet(i).Effect, MainChar.X + (MainChar.Width / 2), MainChar.Y + (MainChar.Height / 2), Bullet(i).From, i
Else
If MainChar.Armor < 0 Then MainChar.Armor = 0
SetDamage "MAINCHAR", Bullet(i).Damage
Bullet(i).Used = False
End If

End If

Bullet(i).X = Bullet(i).X + Bullet(i).Speed


If Bullet(i).X - Game_Offset_X > -16 And _
   Bullet(i).X - Game_Offset_X < Picture1.Width And _
   Bullet(i).Y - Game_Offset_Y > -16 And _
   Bullet(i).Y - Game_Offset_Y < Picture1.Height Then
If Bullet(i).Used = True Then
If Bullet(i).Speed < 0 Then
BitBlt Picture1.Hdc, Bullet(i).X - Game_Offset_X, Bullet(i).Y - Game_Offset_Y, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Width, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Height, Form2.Bullets.Hdc, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_X + 24, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Y, vbSrcAnd
BitBlt Picture1.Hdc, Bullet(i).X - Game_Offset_X, Bullet(i).Y - Game_Offset_Y, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Width, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Height, Form2.Bullets.Hdc, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_X + 36, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Y, vbSrcPaint
Else
BitBlt Picture1.Hdc, Bullet(i).X - Game_Offset_X, Bullet(i).Y - Game_Offset_Y, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Width, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Height, Form2.Bullets.Hdc, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_X, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Y, vbSrcAnd
BitBlt Picture1.Hdc, Bullet(i).X - Game_Offset_X, Bullet(i).Y - Game_Offset_Y, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Width, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Height, Form2.Bullets.Hdc, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_X + 12, MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Y, vbSrcPaint
End If
End If
End If

If Bullet(i).X < -256 Then Bullet(i).Used = False
If Bullet(i).X > 10000 Then Bullet(i).Used = False

End If

Next i
End Function

Private Function MakeBullet()
Dim i As Long

For i = 0 To UBound(Bullet)
If Bullet(i).Used = False Then
Bullet(i).Used = True
MainChar.Weapon_TimeOut(MainChar.Gun_Tag) = MainChar.Weapons(MainChar.Gun_Tag).TimeOut_Rate

If MainChar.Weapons(MainChar.Gun_Tag).SetBurnOutTime > 0 Then
Bullet(i).BurnOut = True
Bullet(i).BurnOutTime = MainChar.Weapons(MainChar.Gun_Tag).SetBurnOutTime
Else
Bullet(i).BurnOut = False
Bullet(i).BurnOutTime = 0
End If

Bullet(i).From = 0
Bullet(i).Weapon = MainChar.Gun_Tag
Bullet(i).Effect = MainChar.Weapons(Bullet(i).Weapon).Effect
Bullet(i).Damage = MainChar.Weapons(Bullet(i).Weapon).Damage
Bullet(i).Speed = MainChar.Weapons(Bullet(i).Weapon).Speed

If MainChar.Direction_Tag = 0 Then
Bullet(i).Speed = MainChar.Weapons(Bullet(i).Weapon).Speed
Bullet(i).X = MainChar.X + MainChar.Width
Bullet(i).Y = MainChar.Y + (MainChar.Height / 2)
Else
Bullet(i).Speed = -MainChar.Weapons(Bullet(i).Weapon).Speed
Bullet(i).X = MainChar.X
Bullet(i).Y = MainChar.Y + (MainChar.Height / 2)
End If
MainChar.Attack_Tag = 1
MainChar.Weapons(MainChar.Gun_Tag).Ammo = MainChar.Weapons(MainChar.Gun_Tag).Ammo - 1
Exit For
End If
Next i
End Function



Private Function MakeAIBullet(ByVal j As Long, ByVal From As Byte)
Dim i As Long
For i = 0 To UBound(Bullet)
If Bullet(i).Used = False Then
Bullet(i).Used = True
AI(j).Weapon_TimeOut = MainChar.Weapons(AI(j).Gun_Tag).TimeOut_Rate

If MainChar.Weapons(AI(j).Gun_Tag).SetBurnOutTime > 0 Then
Bullet(i).BurnOut = True
Bullet(i).BurnOutTime = MainChar.Weapons(AI(j).Gun_Tag).SetBurnOutTime
Else
Bullet(i).BurnOut = False
End If

Bullet(i).From = From
Bullet(i).Weapon = AI(j).Gun_Tag
Bullet(i).Effect = MainChar.Weapons(AI(j).Gun_Tag).Effect
Bullet(i).Damage = MainChar.Weapons(AI(j).Gun_Tag).Damage
Bullet(i).Speed = MainChar.Weapons(AI(j).Gun_Tag).Speed

If AI(j).Direction_Tag = 0 Then
Bullet(i).Speed = MainChar.Weapons(Bullet(i).Weapon).Speed
Bullet(i).X = AI(j).X + AI(j).Width - MainChar.Weapons(Bullet(i).Weapon).Bullet_Draw_Width
Bullet(i).Y = AI(j).Y + (AI(j).Height / 2)
Else
Bullet(i).Speed = -MainChar.Weapons(Bullet(i).Weapon).Speed
Bullet(i).X = AI(j).X
Bullet(i).Y = AI(j).Y + (AI(j).Height / 2)
End If
AI(j).Attack_Tag = 1
Exit For
End If
Next i
End Function


Private Function Effects()
'Generates Explosion, Fire and Chemical effects
Dim i As Long
Dim j As Long
Dim k As Long
Dim c As Long

For i = 0 To UBound(T_Effect)

If T_Effect(i).Used = True Then c = c + 1

If c = UBound(T_Effect) Then Err.Raise 1280, "Effects Procedure", "Array Full"

If T_Effect(i).Used = True Then

If T_Effect(i).Effect > 0 Then T_Effect(i).From = 2

If T_Effect(i).Frame = 0 Then

    If T_Effect(i).From = 0 Then

For k = 0 To UBound(AI)
If AI(k).Armor < 0 Then AI(k).Armor = 0
If AI(k).X + AI(k).Width >= T_Effect(i).X And _
   AI(k).X <= T_Effect(i).X + T_Effect(i).w And _
   AI(k).Y + AI(k).Height >= T_Effect(i).Y And _
   AI(k).Y <= T_Effect(i).Y + T_Effect(i).h And _
   AI(k).Health > 0 Then
SetDamage "AI", T_Effect(i).Damage, k
End If
Next k
    ElseIf T_Effect(i).From = 1 Then

If MainChar.Armor < 0 Then MainChar.Armor = 0
If MainChar.X + MainChar.Width >= T_Effect(i).X And _
   MainChar.X <= T_Effect(i).X + T_Effect(i).w And _
   MainChar.Y + MainChar.Height >= T_Effect(i).Y And _
   MainChar.Y <= T_Effect(i).Y + T_Effect(i).h Then
SetDamage "MAINCHAR", T_Effect(i).Damage
End If
    ElseIf T_Effect(i).From = 2 Then

If T_Effect(i).Frame = 0 And T_Effect(i).Count = 0 Or T_Effect(i).Effect = 3 Then
For k = 0 To UBound(AI)
If AI(k).Armor < 0 Then AI(k).Armor = 0
If AI(k).X + AI(k).Width >= T_Effect(i).X And _
   AI(k).X <= T_Effect(i).X + T_Effect(i).w And _
   AI(k).Y + AI(k).Height >= T_Effect(i).Y And _
   AI(k).Y <= T_Effect(i).Y + T_Effect(i).h Then
SetDamage "AI", T_Effect(i).Damage, k
End If
Next k
If MainChar.Armor < 0 Then MainChar.Armor = 0
If MainChar.X + MainChar.Width >= T_Effect(i).X And _
   MainChar.X <= T_Effect(i).X + T_Effect(i).w And _
   MainChar.Y + MainChar.Height >= T_Effect(i).Y And _
   MainChar.Y <= T_Effect(i).Y + T_Effect(i).h Then
SetDamage "MAINCHAR", T_Effect(i).Damage
End If
End If
    End If
End If

T_Effect(i).Count = T_Effect(i).Count + 1

    If T_Effect(i).Effect = 0 Then

T_Effect(i).Used = False
    
    ElseIf T_Effect(i).Effect = 1 Then

BitBlt Picture1.Hdc, T_Effect(i).X - Game_Offset_X, T_Effect(i).Y - Game_Offset_Y, T_Effect(i).w, T_Effect(i).h, Form2.Effect1(0).Hdc, T_Effect(i).w * T_Effect(i).Frame, 0, vbSrcAnd
BitBlt Picture1.Hdc, T_Effect(i).X - Game_Offset_X, T_Effect(i).Y - Game_Offset_Y, T_Effect(i).w, T_Effect(i).h, Form2.Effect1(1).Hdc, T_Effect(i).w * T_Effect(i).Frame, 0, vbSrcPaint
If T_Effect(i).Count > 1 Then
T_Effect(i).Frame = T_Effect(i).Frame + 1
T_Effect(i).Count = 0
End If
If T_Effect(i).Frame > 5 Then T_Effect(i).Used = False
    
    ElseIf T_Effect(i).Effect = 2 Then

BitBlt Picture1.Hdc, T_Effect(i).X - Game_Offset_X, T_Effect(i).Y - Game_Offset_Y, T_Effect(i).w, T_Effect(i).h, Form2.Effect2(0).Hdc, T_Effect(i).w * T_Effect(i).Frame, 0, vbSrcAnd
BitBlt Picture1.Hdc, T_Effect(i).X - Game_Offset_X, T_Effect(i).Y - Game_Offset_Y, T_Effect(i).w, T_Effect(i).h, Form2.Effect2(1).Hdc, T_Effect(i).w * T_Effect(i).Frame, 0, vbSrcPaint
If T_Effect(i).Count > 2 Then
T_Effect(i).Frame = T_Effect(i).Frame + 1
T_Effect(i).Count = 0
End If
If T_Effect(i).Frame > 5 Then T_Effect(i).Used = False
    
    ElseIf T_Effect(i).Effect = 3 Then

BitBlt Picture1.Hdc, T_Effect(i).X - Game_Offset_X, T_Effect(i).Y - Game_Offset_Y, T_Effect(i).w, T_Effect(i).h, Form2.Effect3(0).Hdc, T_Effect(i).w * T_Effect(i).Frame, 0, vbSrcAnd
BitBlt Picture1.Hdc, T_Effect(i).X - Game_Offset_X, T_Effect(i).Y - Game_Offset_Y, T_Effect(i).w, T_Effect(i).h, Form2.Effect3(1).Hdc, T_Effect(i).w * T_Effect(i).Frame, 0, vbSrcPaint
If T_Effect(i).Count > 2 Then
T_Effect(i).Frame = T_Effect(i).Frame + 1
T_Effect(i).Count = 0
End If
If T_Effect(i).Frame > 2 Then T_Effect(i).Frame = 0: T_Effect(i).Times = T_Effect(i).Times + 1
If T_Effect(i).Times >= 2 Then T_Effect(i).Used = False: T_Effect(i).Times = 0

    ElseIf T_Effect(i).Effect = 4 Then

BitBlt Picture1.Hdc, T_Effect(i).X - Game_Offset_X, T_Effect(i).Y - Game_Offset_Y, T_Effect(i).w, T_Effect(i).h, Form2.Effect4(0).Hdc, T_Effect(i).w * T_Effect(i).Frame, 0, vbSrcAnd
BitBlt Picture1.Hdc, T_Effect(i).X - Game_Offset_X, T_Effect(i).Y - Game_Offset_Y, T_Effect(i).w, T_Effect(i).h, Form2.Effect4(1).Hdc, T_Effect(i).w * T_Effect(i).Frame, 0, vbSrcPaint
If T_Effect(i).Count > 2 Then
T_Effect(i).Frame = T_Effect(i).Frame + 1
T_Effect(i).Count = 0
End If
If T_Effect(i).Frame > 2 Then T_Effect(i).Frame = 0: T_Effect(i).Times = T_Effect(i).Times + 1
If T_Effect(i).Times >= 4 Then T_Effect(i).Used = False: T_Effect(i).Times = 0

    End If

End If

Next i

End Function



Private Sub BT_Timer()
Dim i As Long
DoorDelay = DoorDelay + 1
Msg(0) = Mid(Msg(0).Tag, 1, DoorDelay)

If DoorDelay >= Len(Msg(0).Tag) Then
DoorDelay = 0
For i = 0 To 2
Menu_Msg(i).Enabled = True
Next i
BT.Enabled = False
End If

End Sub

Private Sub Credit_OsamaDraw_Timer()
Dim CheatOn As Boolean
Dim i As Long
If CreditFrame > 0 Then
CreditFrame = CreditFrame - 1
Credits_Osama.Cls
BitBlt Credits_Osama.Hdc, 0, 0, 32, 53, Form2.Osama(1).Hdc, 0, (53 * CreditFrame), vbSrcCopy
DoEvents
Else
For i = 0 To UBound(Cheats)
If Cheats(i) = True Then CheatOn = True: Exit For
Next i
CreditCheat(0).Visible = True
If CheatOn = False Then CreditCheat(1).Visible = True
Credit_OsamaDraw.Enabled = False
End If
End Sub

Private Sub Credits_Command1_Click()
Dim i As Long
LevelSel = 1
ScoreTab.Visible = False
Credit_Screen.Visible = False
Credit_Screen.Enabled = False
Menu_Screen.Enabled = True
Msg(0).Caption = ""
        For i = 0 To Menu_Msg.UBound
        Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
        Next i
    Msg_Door.Enabled = True
    Menu_Screen.Visible = True
SetMusicType "menu", True
End Sub

Private Sub CreditTime_Timer()
Credit_Text.Top = Credit_Text.Top - 1
If Credit_Text.Top <= -Credit_Text.Height Then
CreditFrame = 8
Credit_Taunt.Visible = True
Credits_Command1.Visible = True
Credits_Osama.Visible = True
CreditTime.Enabled = False
Credit_OsamaDraw.Enabled = True
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

DoEvents
If UserInput = True Then

If GetAsyncKeyState(vbKeyControl) < 0 And _
   GetAsyncKeyState(vbKeyU) < 0 And _
   GetAsyncKeyState(vbKeyS) < 0 Then
For i = 0 To 7
Cheat_Unlock(i) = True
Menu_Options.Cheat(i).Enabled = True
Label7 = "Cheats All Unlocked"
QMsg.Visible = True
QuickMsg.Enabled = True
Next i

End If

If GetAsyncKeyState(vbKeyA) < 0 Then
For i = 0 To UBound(Cheats)
Cheats(i) = True
Next i
Label7 = "Debug Bypass"
KillScore = 0
QMsg.Visible = True
QuickMsg.Enabled = True
LevelSel = 1
DoEvents
LoadLevelToEngine
End If

If GetAsyncKeyState(vbKeyO) < 0 And _
   GetAsyncKeyState(vbKeyB) < 0 And _
   GetAsyncKeyState(vbKeyL) < 0 Then
Label7 = "Initiating Final Level... PASSED"
KillScore = 0
QMsg.Visible = True
QuickMsg.Enabled = True
LevelSel = MaxLevels
DoEvents
LoadLevelToEngine

End If

If GetAsyncKeyState(vbKeyP) < 0 And StartGame = True And Msg_DoorOpen = True Then
If Pause_Screen.Visible = False Then
Run = False
Pause_Screen.Visible = True
Else
Pause_Screen.Visible = False
Run = True
DoLoop
End If
End If

End If

End Sub

Private Sub Form_Load()
Dim i As Long

ReDim Objective(1)

UserInput = True

For i = 0 To UBound(Cheats)
'If i <> 1 And i <> 2 Then Cheats(i) = True
Next i

Msg_Type = 1
Game_Difficulty = "Easy"
DoEvents
Form3.Hide

'StartGame = True
'For i = 0 To Menu_Msg.UBound
'Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
'Next i
'Msg(0) = "Loading... " & Chr(13) & "User Level"
'DoEvents
'Msg_Door.Enabled = True
'Menu_Screen.Visible = False
'Run = Not Run
'DoEvents
'SetMusicType "INGAME", False
LevelSel = 1

End Sub

Private Sub Form_Terminate()
Unload Form1
Unload Form2
Unload Form3
Unload Menu_Help
'Unload Menu_LevelSelect
Unload Menu_LoadSave
Unload Menu_Options
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Form2
Unload Form3
Unload Menu_Help
'Unload Menu_LevelSelect
Unload Menu_LoadSave
Unload Menu_Options
End
End Sub

Private Sub Image9_Click()
MsgBox "Cheats for Operation Clearout I: The Hunt for Osama" & Chr(13) & _
       "[Ctrl] + [U] + [S] = Get all cheats unlocked" & Chr(13) & _
       "[O] + [B] + [L] = Goto Final level instantly!" & Chr(13) & _
       "And Enjoy...", vbExclamation, "Surprize, you've found a Sercet Message!"
End Sub

Private Sub Menu_Command1_Click(Index As Integer)
Dim i As Long
If Msg_DoorOpen = True Then

Select Case Index

Case Is = 0
LevelSel = 1
LoadMapFile LevelSel
Menu_Screen.Enabled = False
Msg_Door.Enabled = True
StartGame = True
Menu_Msg(0).Visible = False
Menu_Msg(1).Visible = False
Menu_Msg(2).Visible = True
Msg_LevelName = Level.Name
Msg(0).Tag = Level.Mission_Berifing
SetMusicType "START_LEVEL", True

Case Is = 2
Form1.Enabled = False
Load Menu_LoadSave
Menu_LoadSave.Enabled = True
Menu_LoadSave.Show

Case Is = 3

If Game_Difficulty = "Easy" Then Menu_Options.Difficulty(0).Value = True
If Game_Difficulty = "Normal" Then Menu_Options.Difficulty(1).Value = True
If Game_Difficulty = "Hard" Then Menu_Options.Difficulty(2).Value = True

For i = 0 To 7
Menu_Options.Cheat(i).Enabled = Cheat_Unlock(i)
If Cheats(i) = True Then Menu_Options.Cheat(i).Value = 1 Else Menu_Options.Cheat(i).Value = 0
Next i

Form1.Enabled = False
Load Menu_Options
Menu_Options.Enabled = True
Menu_Options.Show

Case Is = 4
Msg_LevelName = "Main Menu"
Msg_Door.Enabled = True
Menu_Msg(2).Visible = False
Msg(0).Tag = "Are You Sure?"
Msg_Type = 3
Menu_Msg(0).Caption = "Yes": Menu_Msg(0).Visible = True
Menu_Msg(1).Caption = "No": Menu_Msg(1).Visible = True

Case Is = 5
Form1.Enabled = False
Load Menu_Help
Menu_Help.Enabled = True
Menu_Help.Show
Case Is = 6
Shell "explorer.exe mailto:beerboy160@hotmail.com"
End Select
End If


End Sub

Private Sub Menu_Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
For i = 0 To Menu_Command1.UBound
If Menu_Command1(i).BackColor <> &HC000& Then Menu_Command1(i).BackColor = &HC000&: Menu_Command1(i).FontBold = False
Next i
Menu_Command1(Index).BackColor = vbRed: Menu_Command1(Index).FontBold = True
End Sub


Private Sub Menu_Msg_Click(Index As Integer)
Dim i As Long

If Index = 0 Then

Select Case Msg_Type
Case Is = 1
'Open to Menu
    Msg(0).Caption = ""
    For i = 0 To Menu_Msg.UBound
    Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
    Next i
    Msg_Door.Enabled = True
    Menu_Screen.Enabled = True
SetMusicType "menu", True
Case Is = 2
'Load Level
LoadLevelToEngine
Case Is = 3
'Quit Game
End
Case Is = 4
'Return to Menu
    Run = False
    StartGame = False
    Menu_Screen.Enabled = True
    Msg(0).Caption = ""
        For i = 0 To Menu_Msg.UBound
        Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
        Next i
    Msg_Door.Enabled = True
    Menu_Screen.Visible = True
SetMusicType "menu", True
Case Is = 5
'Quit Game
Unload Form1
Unload Form2
Unload Form3
Unload Menu_Help
'Unload Menu_LevelSelect
Unload Menu_LoadSave
Unload Menu_Options
End
Case Is = 6
'Reshow Berifing
LoadMapFile LevelSel 'obtains starting information
'and pre-buffers the level data
ScoreTab.Visible = False
Menu_Msg(0).Visible = False: Menu_Msg(0).Enabled = False
Menu_Msg(1).Visible = False: Menu_Msg(1).Enabled = False
Menu_Msg(2).Visible = True: Menu_Msg(2).Enabled = False
Msg_LevelName = Level.Name
Msg(0).Tag = Level.Mission_Berifing
BT.Enabled = True
SetMusicType "START_LEVEL", True
Case Is = 7

For i = 0 To Menu_Msg.UBound
Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
Next i

SetMusicType "CREDITS", True
CreditCheat(0).Visible = False
CreditCheat(1).Visible = False
Credit_Taunt.Visible = False
Credits_Osama.Visible = False
Credits_Command1.Visible = False
Credit_Screen.Visible = True
Credit_Screen.Enabled = True
Credit_Text.Top = Form1.ScaleHeight
CreditTime.Enabled = True
End Select


ElseIf Index = 1 Then

Select Case Msg_Type
Case Is = 1
'Quit Game
Unload Form1
Unload Form2
Unload Form3
Unload Menu_Help
'Unload Menu_LevelSelect
Unload Menu_LoadSave
Unload Menu_Options
End
Case Is = 2
'Return to Pause Menu without Change
    Pause_Screen.Visible = True
    Msg(0).Caption = ""
        For i = 0 To Menu_Msg.UBound
        Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
        Next i
    Msg_Door.Enabled = True
Case Is = 3
'Return to Menu (Game Not Running)
    Msg(0).Caption = ""
        For i = 0 To Menu_Msg.UBound
        Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
        Next i
    Msg_Door.Enabled = True
SetMusicType "menu", True
Case Is = 4
'Return to Pause Menu Without Change
    Pause_Screen.Visible = True
    Msg(0).Caption = ""
        For i = 0 To Menu_Msg.UBound
        Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
        Next i
    Msg_Door.Enabled = True
Case Is = 5
'Return to Pause Menu without Change
    Pause_Screen.Visible = True
    Msg(0).Caption = ""
        For i = 0 To Menu_Msg.UBound
        Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
        Next i
    Msg_Door.Enabled = True
Case 6, 7
'Return to Menu (Game Not Running)
    ScoreTab.Visible = False
    Msg(0).Caption = ""
        For i = 0 To Menu_Msg.UBound
        Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
        Next i
    Msg_Door.Enabled = True
    Menu_Screen.Visible = True
    Menu_Screen.Enabled = True
    LevelSel = 1
    SetMusicType "menu", True
End Select

Else
LoadLevelToEngine
End If
End Sub

Private Sub Menu_Msg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
For i = 0 To Menu_Msg.UBound
If Menu_Msg(i).BackColor <> &HC000& Then Menu_Msg(i).BackColor = &HC000&: Menu_Msg(i).FontBold = False
Next i
Menu_Msg(Index).BackColor = vbRed: Menu_Msg(Index).FontBold = True
End Sub

Private Sub Msg_Door_Timer()
Dim i As Long
If Msg_DoorOpen = True Then
Picture2(0).Top = Picture2(0).Top + 10
Picture2(1).Top = Picture2(1).Top - 10

If Picture2(0).Top >= 0 Then
Msg_Door.Enabled = False
Msg_DoorOpen = False
BT.Enabled = True
End If

Else

Picture2(0).Top = Picture2(0).Top - 10
Picture2(1).Top = Picture2(1).Top + 10

If Picture2(0).Top <= -Picture2(0).Height Then
For i = 0 To Menu_Msg.UBound
Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
Next i

Msg_Door.Enabled = False
Msg_DoorOpen = True
If StartGame = True Then
'Shape1.Width = 0
'PB1.Visible = False
'Msg(0) = Empty
For i = 0 To 3
LoadStats(i).Visible = False
LoadStats(i) = Empty
Next i

End If

End If
End If

End Sub




Private Sub Pause_Command1_Click(Index As Integer)
Select Case Index
Case Is = 0
Pause_Screen.Visible = False
StartGame = True
Run = True
DoLoop
Case Is = 1
Form1.Enabled = False
Load Menu_LoadSave
Menu_LoadSave.Show
Menu_LoadSave.Enabled = True
Menu_LoadSave.Command1(1).Enabled = True
Case Is = 2
Menu_Msg(0).Visible = True
Menu_Msg(1).Visible = True
Menu_Msg(2).Visible = False

Pause_Screen.Visible = False
Msg_Type = 2
Msg_LevelName = "Menu Main"
Msg(0).Tag = "Are you sure to Restart Level?"
Menu_Msg(0).Caption = "Yes": Menu_Msg(0).Visible = True
Menu_Msg(1).Caption = "No": Menu_Msg(1).Visible = True
Msg_Door.Enabled = True
Case Is = 3
Menu_Msg(0).Visible = True
Menu_Msg(1).Visible = True
Menu_Msg(2).Visible = False

Pause_Screen.Visible = False
Msg_Type = 4
Msg_LevelName = "Menu Main"
Msg(0).Tag = "Are you sure to Quit to Menu?"
Menu_Msg(0).Caption = "Yes": Menu_Msg(0).Visible = True
Menu_Msg(1).Caption = "No": Menu_Msg(1).Visible = True
Msg_Door.Enabled = True

Case Is = 4
Pause_Screen.Visible = False
Msg_LevelName = "Main Menu"
Msg_Door.Enabled = True
Menu_Msg(2).Visible = False
Msg(0).Tag = "Are You Sure?"
Msg_Type = 5
Menu_Msg(0).Caption = "Yes": Menu_Msg(0).Visible = True
Menu_Msg(1).Caption = "No": Menu_Msg(1).Visible = True

Case Is = 5
Form1.Enabled = False
Load Menu_Objectives
Menu_Objectives.MBText = Level.Mission_Berifing
Menu_Objectives.Show
End Select

End Sub

Private Sub Pause_Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
For i = 0 To Pause_Command1.UBound
If Pause_Command1(i).BackColor <> &HC000& Then Pause_Command1(i).BackColor = &HC000&: Pause_Command1(i).FontBold = False
Next i
Pause_Command1(Index).BackColor = vbRed: Pause_Command1(Index).FontBold = True
End Sub

Private Sub QuickMsg_Timer()
Label7 = Empty
QMsg.Visible = False
End Sub



Private Function SetEffect(ByVal Effect As Byte, ByVal X As Long, ByVal Y As Long, ByVal From As Byte, ByVal b As Long)
Dim i As Long
For i = 0 To UBound(T_Effect)
If T_Effect(i).Used = False Then
T_Effect(i).Used = True
T_Effect(i).Effect = Effect
T_Effect(i).From = From
T_Effect(i).Damage = Bullet(b).Damage
T_Effect(i).Times = 0

    If T_Effect(i).Effect = 0 Then

T_Effect(i).X = X
T_Effect(i).Y = Y
T_Effect(i).Count = 0
T_Effect(i).Frame = 0

    ElseIf T_Effect(i).Effect = 1 Then

T_Effect(i).Count = 0
T_Effect(i).Frame = 0
T_Effect(i).w = 64
T_Effect(i).h = 64
T_Effect(i).X = X - (T_Effect(i).w / 2)
T_Effect(i).Y = Y - (T_Effect(i).h / 2)

    ElseIf T_Effect(i).Effect = 2 Then

T_Effect(i).Count = 0
T_Effect(i).Frame = 0
T_Effect(i).w = 256
T_Effect(i).h = 256
T_Effect(i).X = X - (T_Effect(i).w / 2)
T_Effect(i).Y = Y - (T_Effect(i).h / 2)

    ElseIf T_Effect(i).Effect = 3 Then

T_Effect(i).Count = 0
T_Effect(i).Frame = 0
T_Effect(i).w = 32
T_Effect(i).h = 32
T_Effect(i).X = X - (T_Effect(i).w / 2)
T_Effect(i).Y = Y - (T_Effect(i).h / 2)

    ElseIf T_Effect(i).Effect = 4 Then

T_Effect(i).Count = 0
T_Effect(i).Frame = 0
T_Effect(i).w = 32
T_Effect(i).h = 32
T_Effect(i).X = X - (T_Effect(i).w / 2)
T_Effect(i).Y = Y - (T_Effect(i).h / 2)

End If

Bullet(b).Used = False
Exit For
End If
Next i

If Bullet(b).Used = True Then Bullet(b).Used = False

End Function

Private Function AI_System(ByVal AI_Number As Long, ByVal AI_Type As String)
'The AI System which control the baddies action
Dim i As Long
Dim c As Long
If AI(AI_Number).SetPoint(AI(AI_Number).Point).Max_Wait > 0 Then
    If AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = False Then
    If AI(AI_Number).SetPoint(AI(AI_Number).Point).Wait >= AI(AI_Number).SetPoint(AI(AI_Number).Point).Max_Wait Then
    If AI(AI_Number).Point = 1 Then
        AI(AI_Number).SetPoint(0).Wait = 0
        AI(AI_Number).Point = 0
    Else
        AI(AI_Number).SetPoint(1).Wait = 0
        AI(AI_Number).Point = 1
        End If
    End If
End If

If AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = False Then
If AI(AI_Number).X < AI(AI_Number).SetPoint(1).X And AI(AI_Number).Point = 0 Or _
   AI(AI_Number).X > AI(AI_Number).SetPoint(0).X And AI(AI_Number).Point = 1 Then

If AI(AI_Number).Point = 0 Then
AI(AI_Number).Direction_Tag = 0
AI(AI_Number).Keys(1) = True
ElseIf AI(AI_Number).Point = 1 Then
AI(AI_Number).Direction_Tag = 1
AI(AI_Number).Keys(0) = True
End If

Else

AI(AI_Number).SetPoint(AI(AI_Number).Point).Wait = AI(AI_Number).SetPoint(AI(AI_Number).Point).Wait + 1

End If
End If
End If

    If AI_Type = "TALIBAN_1" Then 'CASE = "TALIBAN_1"

If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) < 320 And _
   AI(AI_Number).Y + AI(AI_Number).Height >= MainChar.Y And _
   AI(AI_Number).Y <= MainChar.Y + MainChar.Height And _
   MainChar.Health > 0 Then

If MainChar.X < AI(AI_Number).X Then
AI(AI_Number).Direction_Tag = 1
Else
AI(AI_Number).Direction_Tag = 0
End If
   
   
If AI(AI_Number).Weapon_TimeOut = 0 And MainChar.Health > 0 And _
   AI(AI_Number).Y + AI(AI_Number).Height >= MainChar.Y - 8 And _
   AI(AI_Number).Y <= MainChar.Y + MainChar.Height + 8 Then
 
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True
MakeAIBullet AI_Number, 1
End If
Else
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = False
End If


    ElseIf AI_Type = "TALIBAN_2" Then 'CASE = "TALIBAN_2"


'Jump Dodge
If MainChar.Attack_Tag = 1 And MainChar.Gun_Tag > 0 Then
AI(AI_Number).Keys(2) = True
End If

'Primary Range detector
If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) < 190 And _
   AI(AI_Number).Y + AI(AI_Number).Height >= MainChar.Y And _
   AI(AI_Number).Y <= MainChar.Y + MainChar.Height And _
   MainChar.Health > 0 Then

If MainChar.X < AI(AI_Number).X Then
AI(AI_Number).Direction_Tag = 1
Else
AI(AI_Number).Direction_Tag = 0
End If
      
   
If AI(AI_Number).Weapon_TimeOut = 0 And MainChar.Health > 0 And _
   AI(AI_Number).Y + AI(AI_Number).Height >= MainChar.Y - 8 And _
   AI(AI_Number).Y <= MainChar.Y + MainChar.Height + 8 Then

   
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True
MakeAIBullet AI_Number, 1
End If
Else
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = False
End If

    ElseIf AI_Type = "TALIBAN_3" Then 'CASE = "TALIBAN_3"


'Primary Range detector
If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) < 132 And _
   AI(AI_Number).Y + AI(AI_Number).Height >= MainChar.Y And _
   AI(AI_Number).Y <= MainChar.Y + MainChar.Height And _
   MainChar.Health > 0 Then


If MainChar.X < AI(AI_Number).X Then
AI(AI_Number).Direction_Tag = 1
Else
AI(AI_Number).Direction_Tag = 0
End If
   
   
If AI(AI_Number).Weapon_TimeOut = 0 And MainChar.Health > 0 And _
   AI(AI_Number).Y + AI(AI_Number).Height >= MainChar.Y - 8 And _
   AI(AI_Number).Y <= MainChar.Y + MainChar.Height + 8 Then

AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True
MakeAIBullet AI_Number, 1
End If
Else
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = False
End If

    ElseIf AI_Type = "TALIBAN_4" Then 'CASE = "TALIBAN_4"


If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) < 640 And _
   AI(AI_Number).Y + AI(AI_Number).Height >= MainChar.Y And _
   AI(AI_Number).Y <= MainChar.Y + MainChar.Height And _
   MainChar.Health > 0 Then

If MainChar.X < AI(AI_Number).X Then
AI(AI_Number).Direction_Tag = 1
Else
AI(AI_Number).Direction_Tag = 0
End If
   
If AI(AI_Number).Weapon_TimeOut = 0 And MainChar.Health > 0 And _
   AI(AI_Number).Y + AI(AI_Number).Height >= MainChar.Y - 8 And _
   AI(AI_Number).Y <= MainChar.Y + MainChar.Height + 8 Then

AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True
MakeAIBullet AI_Number, 1
End If
Else
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = False
End If

    ElseIf AI_Type = "TALIBAN_5" Then 'CASE = "TALIBAN_5"


'Primary Range detector
If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) < 160 And _
   Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) > 48 And _
   AI(AI_Number).Y + AI(AI_Number).Height >= MainChar.Y And _
   AI(AI_Number).Y <= MainChar.Y + MainChar.Height And _
   MainChar.Health > 0 Then


If MainChar.X < AI(AI_Number).X Then
AI(AI_Number).Direction_Tag = 1
Else
AI(AI_Number).Direction_Tag = 0
End If
   
   
If AI(AI_Number).Weapon_TimeOut = 0 And MainChar.Health > 0 And _
   AI(AI_Number).Y + AI(AI_Number).Height >= MainChar.Y - 8 And _
   AI(AI_Number).Y <= MainChar.Y + MainChar.Height + 8 Then
   
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True
MakeAIBullet AI_Number, 1
End If
Else
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = False
End If

    ElseIf AI_Type = "TALIBAN_BOSS1" Then 'CASE = "TALIBAN_BOSS1"

SetMusicType "INGAME_BOSS", True
'Be warned these guys can run you over, killing you instantly
'they take a hell of a beating without a good weapon.
'Note: They don't stop moving at all when they're attacking
'so move out of their way!

If MainChar.X + MainChar.Width > AI(AI_Number).X And _
   MainChar.X < AI(AI_Number).X + AI(AI_Number).Width And _
   MainChar.Y >= AI(AI_Number).Y And _
   MainChar.Y <= AI(AI_Number).Y + (AI(AI_Number).Height / 2) And _
   AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = False Then
MainChar.Health = 0
MainChar.Squished = True
End If

If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) >= 32 And _
   Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) <= 300 And _
   AI(AI_Number).Y >= MainChar.Y - 64 And _
   AI(AI_Number).Y <= MainChar.Y + 8 And _
   MainChar.Health > 0 Then
If MainChar.X < AI(AI_Number).X Then
AI(AI_Number).Direction_Tag = 1
Else
AI(AI_Number).Direction_Tag = 0
End If
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True
If AI(AI_Number).Weapon_TimeOut = 0 Then
AI(AI_Number).Gun_Tag = 2
MakeAIBullet AI_Number, 1
AI(AI_Number).Weapon_TimeOut = 3
End If
End If

If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) > 300 And _
   Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) < 600 And _
   AI(AI_Number).Y >= MainChar.Y - 64 And _
   AI(AI_Number).Y <= MainChar.Y + 8 And _
   MainChar.Health > 0 Then

If MainChar.X < AI(AI_Number).X Then
AI(AI_Number).Direction_Tag = 1
AI(AI_Number).Point = 1
Else
AI(AI_Number).Direction_Tag = 0
AI(AI_Number).Point = 0
End If
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True
c = Int(Rnd * 16)
If AI(AI_Number).Weapon_TimeOut = 0 And MainChar.Health > 0 And _
   AI(AI_Number).Y >= MainChar.Y - MainChar.Height - 8 And _
   AI(AI_Number).Y <= MainChar.Y + 8 Then
   If c <> 1 Then
AI(AI_Number).Gun_Tag = 6

   Else
AI(AI_Number).Gun_Tag = 7
   End If
MakeAIBullet AI_Number, 1
AI(AI_Number).Weapon_TimeOut = 8
End If
End If

If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) <= 240 Or _
   Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) >= 600 Or _
   AI(AI_Number).Y <= (MainChar.Y - 64) Or _
   AI(AI_Number).Y >= (MainChar.Y + 8) Then
    If AI(AI_Number).Y >= MainChar.Y - 64 And _
        AI(AI_Number).Y <= MainChar.Y + 8 And _
        MainChar.Health > 0 Then
            If MainChar.X > AI(AI_Number).X Then
            AI(AI_Number).Direction_Tag = 0
            AI(AI_Number).Point = 0
            AI(AI_Number).SetPoint(0).Wait = 0
            Else
            AI(AI_Number).Direction_Tag = 1
            AI(AI_Number).Point = 1
            AI(AI_Number).SetPoint(1).Wait = 0
            End If
    End If
AI(AI_Number).SetPoint(0).Stop = False
AI(AI_Number).SetPoint(1).Stop = False
End If

    ElseIf AI_Type = "TALIBAN_BOSS_OSAMA" Then 'CASE = "TALIBAN_BOSS_OSAMA"

If AI(AI_Number).Health > Osama_Low_Health Then
SetMusicType "INGAME_BOSS_OSAMA", True
Else
SetMusicType "INGAME_BOSS_OSAMA_LOWHEALTH", True
End If

OsamaTaunt = False
If MainChar.Attack_Tag = 1 And MainChar.Gun_Tag > 0 And _
   AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True Then
AI(AI_Number).Keys(2) = True
End If
If MainChar.Y < AI(AI_Number).Y - MainChar.Height And MainChar.Attack_Tag = 0 And _
   AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True Then
AI(AI_Number).Keys(2) = True
End If


If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) < (AI(AI_Number).Width / 1.5) And _
   MainChar.Y > AI(AI_Number).Y - MainChar.Height And _
   MainChar.Y < AI(AI_Number).Y + AI(AI_Number).Height And _
   MainChar.Health > 0 Then
If AI(AI_Number).X < (MainChar.X + MainChar.Width / 2) Then
AI(AI_Number).X = AI(AI_Number).SetPoint(1).X
AI(AI_Number).Point = 1
Else
AI(AI_Number).X = AI(AI_Number).SetPoint(0).X
AI(AI_Number).Point = 0
End If
End If


If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) < 360 And _
   AI(AI_Number).Y >= MainChar.Y - AI(AI_Number).Height - MainChar.Height And _
   AI(AI_Number).Y <= MainChar.Y + AI(AI_Number).Height + MainChar.Height And _
   MainChar.Health > 0 Then
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True
Else
AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = False
End If

If AI(AI_Number).SetPoint(AI(AI_Number).Point).Stop = True Then

If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) >= 0 And _
   Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) <= 150 And _
   MainChar.Health > 0 Then
If MainChar.X < AI(AI_Number).X Then
AI(AI_Number).Direction_Tag = 1
Else
AI(AI_Number).Direction_Tag = 0
End If
If AI(AI_Number).Weapon_TimeOut = 0 Then
AI(AI_Number).Gun_Tag = 3
MakeAIBullet AI_Number, 1
AI(AI_Number).Weapon_TimeOut = 0
End If
End If

If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) > 150 And _
   Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) <= 272 And _
   MainChar.Health > 0 Then
If MainChar.X < AI(AI_Number).X Then
AI(AI_Number).Direction_Tag = 1
Else
AI(AI_Number).Direction_Tag = 0
End If
If AI(AI_Number).Weapon_TimeOut = 0 Then
AI(AI_Number).Gun_Tag = 2
MakeAIBullet AI_Number, 1
AI(AI_Number).Weapon_TimeOut = 3
End If
End If

If Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) > 272 And _
   Abs((AI(AI_Number).X + (AI(AI_Number).Width / 2)) - (MainChar.X + (MainChar.Width / 2))) < 480 And _
   AI(AI_Number).Y >= MainChar.Y - 128 And _
   AI(AI_Number).Y <= MainChar.Y + 64 And _
   MainChar.Health > 0 Then

If MainChar.X < AI(AI_Number).X Then
AI(AI_Number).Direction_Tag = 1
Else
AI(AI_Number).Direction_Tag = 0
End If
If AI(AI_Number).Weapon_TimeOut = 0 And MainChar.Health > 0 And _
   AI(AI_Number).Y >= MainChar.Y - MainChar.Height - 8 And _
   AI(AI_Number).Y <= MainChar.Y + 8 Then
AI(AI_Number).Gun_Tag = 7
MakeAIBullet AI_Number, 1
AI(AI_Number).Weapon_TimeOut = 15
End If
End If
End If

If AI(AI_Number).Health <= Osama_Low_Health And AI(AI_Number).Weapon_TimeOut >= 0 Then
AI(AI_Number).Weapon_TimeOut = -Set_Game_Delay * 16
'to prevent osama from self destructing, you must deliver
'a powerful blow when the health is about 1/4
End If

If AI(AI_Number).Health <= Osama_Low_Health Then
AI(AI_Number).Weapon_TimeOut = AI(AI_Number).Weapon_TimeOut + 1
OsamaTaunt = True
End If

If AI(AI_Number).Health <= Osama_Low_Health And AI(AI_Number).Weapon_TimeOut >= -1 And NukeSet = False Then
Nuking = True
NukeSet = True
SetDamage "MAINCHAR", 10000
For i = 0 To UBound(AI)
SetDamage "AI", 10000, i
Next i
If BGSet <> 5 Then BGSet = 2
End If

End If 'CASE


End Function

Private Function AI_Movement(ByVal AI_Number As Long)
'Moves the AI like the character does
If AI(AI_Number).Keys(0) = True And AI(AI_Number).Keys(1) = False Then
AI(AI_Number).X = AI(AI_Number).X - AI(AI_Number).MoveSpeed
AI(AI_Number).Direction_Tag = 1
End If

If AI(AI_Number).Keys(1) = True And AI(AI_Number).Keys(0) = False Then
AI(AI_Number).X = AI(AI_Number).X + AI(AI_Number).MoveSpeed
AI(AI_Number).Direction_Tag = 0
End If

If AI(AI_Number).Keys(0) = True And AI(AI_Number).Keys(1) = False Or _
   AI(AI_Number).Keys(0) = False And AI(AI_Number).Keys(1) = True Then
If AI(AI_Number).Ground = True Then
AI(AI_Number).Walk_Delay = AI(AI_Number).Walk_Delay - 1

If AI(AI_Number).Walk_Delay <= 0 Then

If AI(AI_Number).Walk_Tag = 1 Then
AI(AI_Number).Walk_Tag = 0
Else
AI(AI_Number).Walk_Tag = 1
End If

AI(AI_Number).Walk_Delay = 2

End If
End If

Else

AI(AI_Number).Walk_Tag = 2

End If

If AI(AI_Number).Keys(2) = True And AI(AI_Number).Ground = True Then
If AI(AI_Number).AI_Type = "TALIBAN_BOSS_OSAMA" Then
AI(AI_Number).GravityForce = -16
Else
AI(AI_Number).GravityForce = -8
End If
End If

End Function

Private Function Items(ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal h As Long, ByVal Item_Name As String, ByVal i As Long)
'Generates and Check Collision of Game Items
On Error Resume Next

If Item(i).Used = False Then

If X - Game_Offset_X > -Item(i).Width And _
   X - Game_Offset_X < Picture1.Width And _
   Y - Game_Offset_Y > -Item(i).Height And _
   Y - Game_Offset_Y < Picture1.Height Then
BitBlt Picture1.Hdc, X - Game_Offset_X, Y - Game_Offset_Y, w, h, Form2.Items(0).Hdc, Item(i).Draw_X, Item(i).Draw_Y, vbSrcAnd
BitBlt Picture1.Hdc, X - Game_Offset_X, Y - Game_Offset_Y, w, h, Form2.Items(1).Hdc, Item(i).Draw_X, Item(i).Draw_Y, vbSrcPaint

If MainChar.X + MainChar.Width > X And _
   MainChar.X < Item(i).X + w And _
   MainChar.Y + MainChar.Height > Y And _
   MainChar.Y < Item(i).Y + h Then
   
    If Item_Name = "PISTOL" Then 'ITEM CASE

MainChar.Weapons(1).Have = True
If MainChar.Weapons(1).Ammo <= 190 Then
MainChar.Weapons(1).Ammo = MainChar.Weapons(1).Ammo + 10
Else
MainChar.Weapons(1).Ammo = MainChar.Weapons(1).Ammo + (MainChar.Weapons(1).Ammo - 190)
End If

Item(i).Used = True
MainChar.Gun_Tag = 1: MainChar.Gun_Draw_Tag = 1: MainChar.Weapon_TimeOut(1) = 0

    ElseIf Item_Name = "PISTOL_AMMO" Then

Select Case MainChar.Weapons(1).Ammo
Case Is <= 180
MainChar.Weapons(1).Ammo = MainChar.Weapons(1).Ammo + 20
Item(i).Used = True
Case 181 To 199
MainChar.Weapons(1).Ammo = MainChar.Weapons(1).Ammo + (MainChar.Weapons(1).Ammo - 180)
Item(i).Used = True
Case Is >= 200
End Select

    ElseIf Item_Name = "AK33" Then

MainChar.Weapons(2).Have = True
If MainChar.Weapons(2).Ammo <= 475 Then
MainChar.Weapons(2).Ammo = MainChar.Weapons(2).Ammo + 25
Else
MainChar.Weapons(2).Ammo = MainChar.Weapons(2).Ammo + (MainChar.Weapons(2).Ammo - 475)
End If
Item(i).Used = True
MainChar.Gun_Tag = 2: MainChar.Gun_Draw_Tag = 2: MainChar.Weapon_TimeOut(2) = 0

    ElseIf Item_Name = "AK33_AMMO" Then

Select Case MainChar.Weapons(2).Ammo
Case Is <= 450
MainChar.Weapons(2).Ammo = MainChar.Weapons(2).Ammo + 50
Item(i).Used = True
Case 451 To 499
MainChar.Weapons(2).Ammo = MainChar.Weapons(2).Ammo + (MainChar.Weapons(2).Ammo - 450)
Item(i).Used = True
Case Is >= 500
End Select

    ElseIf Item_Name = "PULSAR" Then

MainChar.Weapons(3).Have = True
If MainChar.Weapons(3).Ammo <= 475 Then
MainChar.Weapons(3).Ammo = MainChar.Weapons(3).Ammo + 25
Else
MainChar.Weapons(3).Ammo = MainChar.Weapons(3).Ammo + (MainChar.Weapons(3).Ammo - 475)
End If
Item(i).Used = True
MainChar.Gun_Tag = 3: MainChar.Gun_Draw_Tag = 3: MainChar.Weapon_TimeOut(3) = 0

    ElseIf Item_Name = "PULSAR_AMMO" Then

Select Case MainChar.Weapons(3).Ammo
Case Is <= 450
MainChar.Weapons(3).Ammo = MainChar.Weapons(3).Ammo + 50
Item(i).Used = True
Case 451 To 499
MainChar.Weapons(3).Ammo = MainChar.Weapons(3).Ammo + (MainChar.Weapons(3).Ammo - 450)
Item(i).Used = True
Case Is >= 500
End Select

    ElseIf Item_Name = "SLIME" Then
    
MainChar.Weapons(8).Have = True
If MainChar.Weapons(8).Ammo <= 900 Then
MainChar.Weapons(8).Ammo = MainChar.Weapons(8).Ammo + 100
Else
MainChar.Weapons(8).Ammo = MainChar.Weapons(8).Ammo + (MainChar.Weapons(8).Ammo - 900)
End If
Item(i).Used = True
MainChar.Gun_Tag = 8: MainChar.Gun_Draw_Tag = 6: MainChar.Weapon_TimeOut(8) = 0

    ElseIf Item_Name = "SLIME_AMMO" Then

Select Case MainChar.Weapons(8).Ammo
Case Is <= 800
MainChar.Weapons(8).Ammo = MainChar.Weapons(8).Ammo + 200
Item(i).Used = True
Case 801 To 899
MainChar.Weapons(8).Ammo = MainChar.Weapons(8).Ammo + (MainChar.Weapons(8).Ammo - 800)
Item(i).Used = True
Case Is >= 900
End Select

    ElseIf Item_Name = "FLAME" Then

MainChar.Weapons(4).Have = True
If MainChar.Weapons(4).Ammo <= 900 Then
MainChar.Weapons(4).Ammo = MainChar.Weapons(4).Ammo + 100
Else
MainChar.Weapons(4).Ammo = MainChar.Weapons(4).Ammo + (MainChar.Weapons(4).Ammo - 900)
End If
Item(i).Used = True
MainChar.Gun_Tag = 4: MainChar.Gun_Draw_Tag = 4: MainChar.Weapon_TimeOut(4) = 0

    ElseIf Item_Name = "FLAME_AMMO" Then

Select Case MainChar.Weapons(4).Ammo
Case Is <= 800
MainChar.Weapons(4).Ammo = MainChar.Weapons(4).Ammo + 200
Item(i).Used = True
Case 801 To 899
MainChar.Weapons(4).Ammo = MainChar.Weapons(4).Ammo + (MainChar.Weapons(4).Ammo - 800)
Item(i).Used = True
Case Is >= 900
End Select

    ElseIf Item_Name = "RGB_DEVICE" Then

MainChar.Weapons(5).Have = True
MainChar.Weapons(6).Have = True
MainChar.Weapons(7).Have = True
Item(i).Used = True

    ElseIf Item_Name = "GRENADE" Then

Select Case MainChar.Weapons(5).Ammo
Case Is <= 45
MainChar.Weapons(5).Ammo = MainChar.Weapons(5).Ammo + 5
Item(i).Used = True
Case 46 To 49
MainChar.Weapons(5).Ammo = MainChar.Weapons(5).Ammo + (MainChar.Weapons(5).Ammo - 45)
Item(i).Used = True
Case Is >= 50
End Select

    ElseIf Item_Name = "ROCKET" Then

Select Case MainChar.Weapons(6).Ammo
Case Is <= 45
MainChar.Weapons(6).Ammo = MainChar.Weapons(6).Ammo + 5
Item(i).Used = True
Case 46 To 49
MainChar.Weapons(6).Ammo = MainChar.Weapons(6).Ammo + (MainChar.Weapons(6).Ammo - 45)
Item(i).Used = True
Case Is >= 50
End Select

    ElseIf Item_Name = "NUKEBLAST" Then

Select Case MainChar.Weapons(7).Ammo
Case Is <= 10
MainChar.Weapons(7).Ammo = MainChar.Weapons(7).Ammo + 1
Item(i).Used = True
Case Is >= 10
End Select

    ElseIf Item_Name = "10HEALTH" Then

If MainChar.Health < 100 Then
  If MainChar.Health < 90 Then
MainChar.Health = MainChar.Health + 10
  Else
MainChar.Health = MainChar.Health + (100 - MainChar.Health)
  End If
Item(i).Used = True
End If

    ElseIf Item_Name = "25HEALTH" Then
    
If MainChar.Health < 100 Then
  If MainChar.Health < 75 Then
MainChar.Health = MainChar.Health + 25
  Else
MainChar.Health = MainChar.Health + (100 - MainChar.Health)
  End If
Item(i).Used = True
End If

    ElseIf Item_Name = "JETPACK" Then

If MainChar.JetPack_Fuel < 100 Or MainChar.JetPack_Have = False Then
MainChar.ItemSel = 0
MainChar.JetPack_Have = True
MainChar.JetPack_Fuel = 100
Item(i).Used = True
End If

    ElseIf Item_Name = "ARMOR" Then
If MainChar.Armor < 100 Then
MainChar.Armor = 100
Item(i).Used = True
End If

End If

End If



End If
End If


End Function

Private Function Swicthes()
'The main swicth process, which activates commands
'from swicthes
Dim i As Long

For i = 1 To UBound(Swicth)

If Swicth(i).EnterSet = True And _
   MainChar.X > Swicth(i).X - MainChar.Width And _
   MainChar.X < Swicth(i).X + Swicth(i).Width And _
   MainChar.Y > Swicth(i).Y - MainChar.Height And _
   MainChar.Y < Swicth(i).Y + Swicth(i).Height Then
Swicth(i).Enter = True
End If
If Swicth(i).PressedSet = True And _
   MainChar.X > Swicth(i).X - MainChar.Width And _
   MainChar.X < Swicth(i).X + Swicth(i).Width And _
   MainChar.Y > Swicth(i).Y - MainChar.Height And _
   MainChar.Y < Swicth(i).Y + Swicth(i).Height And _
   MainChar.Keys(4) = True Then
Swicth(i).Pressed = True
End If


If Swicth(i).Linked = True Then
Swicth(i).On = Swicth(Swicth(i).LinkSwicthTo).On
End If

If Swicth(i).Enter = True Or Swicth(i).Pressed = True Then

If Swicth(i).Times >= 0 Then
If Swicth(i).Times = Swicth(i).Max_Times Then Swicth(i).On = Not Swicth(i).On

Swicth(i).Times = Swicth(i).Times - 1
If Swicth(i).Command = "MOVE" Then
Console.Move Swicth(i).Target, Swicth(i).Selected, Swicth(i).By_X, Swicth(i).By_Y
ElseIf Swicth(i).Command = "SET ON" Then
Console.SetOn Swicth(i).Target, Swicth(i).Selected
ElseIf Swicth(i).Command = "SET OFF" Then
If Swicth(i).Times Mod GetEffectLength(Swicth(i).By_X) = 0 Then
Console.SetOff Swicth(i).Target, Swicth(i).Selected, Swicth(i).By_X
End If
ElseIf Swicth(i).Command = "CHANGE MUSIC" Then
SetMusicType Swicth(i).Target, CBool(Swicth(i).By_X)
End If
End If
End If

If Swicth(i).Pressed = True Or Swicth(i).Enter = True Then
If Swicth(i).Linked = True Then
Swicth(Swicth(i).LinkSwicthTo).Enter = True
Swicth(Swicth(i).LinkSwicthTo).Pressed = True
Swicth(i).Enter = False
Swicth(i).Pressed = False
End If
End If

If Swicth(i).Reverse = True And Swicth(i).Times = 0 Then
Swicth(i).Enter = False: Swicth(i).Pressed = False
Swicth(i).Times = Swicth(i).Max_Times
If Swicth(i).By_X > 0 Then Swicth(i).By_X = -Swicth(i).By_X Else Swicth(i).By_X = Abs(Swicth(i).By_X)
If Swicth(i).By_Y > 0 Then Swicth(i).By_Y = -Swicth(i).By_Y Else Swicth(i).By_Y = Abs(Swicth(i).By_Y)
ElseIf Swicth(i).Reverse = False And Swicth(i).Times = 0 Then
Swicth(i).Pressed = False
Swicth(i).Enter = False
End If

Next i
End Function

Private Function Objectives()
'The Objectives Process
Dim i As Long
Dim j As Long
Dim a As Long
Dim c As Long
Dim d As Single
Dim OKC As Long
Dim OCC As Long
OCC = 0
OKC = 0
For i = 1 To UBound(Objective)

If Objective(i).Condition = "SWICTH ENABLED" Then
 If Swicth(Objective(i).Selected).Enabled = True Then
 Objective(i).Completed = True
 Else
 Objective(i).Failed = True
 End If
ElseIf Objective(i).Condition = "KILL AI" Then
 If AI(Objective(i).Selected).Health <= 0 Then Objective(i).Completed = True
ElseIf Objective(i).Condition = "KILL ALL AI" Then
a = 0
 For j = 1 To UBound(AI)
 If AI(j).Health <= 0 And AI(j).AnimFrame >= 8 Then a = a + 1
 If a >= UBound(AI) Then Objective(i).Completed = True
 Next j
ElseIf Objective(i).Condition = "SWICTH ON" Then
 If Swicth(Objective(i).Selected).On = True Then Objective(i).Completed = True Else Objective(i).Completed = False
ElseIf Objective(i).Condition = "SWICTH OFF" Then
 If Swicth(Objective(i).Selected).On = False Then Objective(i).Completed = True Else Objective(i).Completed = False
ElseIf Objective(i).Condition = "ITEM CLAIMED" Then
 If Item(Objective(i).Selected).Used = True Then Objective(i).Completed = True Else Objective(i).Completed = False
ElseIf Objective(i).Condition = "LOCATION REACHED" Then
 If MainChar.X + MainChar.Width >= Objective(i).X And _
 MainChar.X <= Objective(i).X + Objective(i).Width And _
 MainChar.Y + MainChar.Height >= Objective(i).Y And _
 MainChar.Y <= Objective(i).Y + Objective(i).Height Then
 Objective(i).Completed = True
 Else
 Objective(i).Completed = False
 End If
ElseIf Objective(i).Condition = "LOCATION REACHED STANDARD" Then
 If MainChar.X > Objective(i).X Then Objective(i).Completed = True Else Objective(i).Completed = False
Else
End If


If Objective(i).TimeLimit = True And Objective(i).Completed = False Then
Picture1.CurrentX = Time_HUD_X + 26
Picture1.CurrentY = Time_HUD_Y

If Game_Delay >= Set_Game_Delay Then
    Objective(i).TimeLeft = Objective(i).TimeLeft - 1
End If

If Objective(i).TimeLeft < 300 Then a = 4
If Objective(i).TimeLeft < 120 Then a = 3
If Objective(i).TimeLeft < 60 Then a = 2
If Objective(i).TimeLeft < 30 Then a = 1
If Objective(i).TimeLeft = 0 Then a = 0

BitBlt Picture1.Hdc, Time_HUD_X, Time_HUD_Y, 24, 24, Form2.HUD(0).Hdc, 24 * a, 96, vbSrcAnd
BitBlt Picture1.Hdc, Time_HUD_X, Time_HUD_Y, 24, 24, Form2.HUD(1).Hdc, 24 * a, 96, vbSrcPaint
If Objective(i).TimeLeft >= 30 Then Picture1.ForeColor = vbGreen Else Picture1.ForeColor = vbRed


If MediaPlayer1.Tag = "INGAME" Or MediaPlayer1.Tag = "INGAME_LOWTIME" Then
If Objective(i).TimeLeft < 30 Then SetMusicType "INGAME_LOWTIME", True Else SetMusicType "INGAME", True
End If

Picture1.Print Objective(i).TimeLeft \ 60 Mod 60 & ":" & IIf(Objective(i).TimeLeft Mod 60 >= 10, Objective(i).TimeLeft Mod 60, "0" & (Objective(i).TimeLeft Mod 60))

 If Objective(i).TimeLeft <= 0 Then
   Objective(i).Failed = True
 End If
End If


If Objective(i).Key = True And Objective(i).Failed = True Then
SetMusicType "FAIL_LEVEL", True
Menu_Msg(0).Visible = True
Menu_Msg(1).Visible = True
Menu_Msg(2).Visible = False
Run = False
DeadTime = 0
StartGame = False
Msg_Type = 6
Msg_LevelName = "Level Failed"
Msg(0).Tag = "Primary Objective Comprimised - MISSION FAILED" & Chr(13) & Level.Mission_Fail_Text & Chr(13) & "Play Again?"
Menu_Msg(0).Caption = "Yes": Menu_Msg(0).Visible = True
Menu_Msg(1).Caption = "No": Menu_Msg(1).Visible = True
Msg_Door.Enabled = True
End If

If Objective(i).Key = True Then OKC = OKC + 1
If Objective(i).Key = True And Objective(i).Completed = True Then OCC = OCC + 1


Next i

Picture1.CurrentX = Missions_HUD_X + 26
Picture1.CurrentY = Missions_HUD_Y
Picture1.ForeColor = vbWhite
BitBlt Picture1.Hdc, Missions_HUD_X, Missions_HUD_Y, 24, 24, Form2.HUD(0).Hdc, 24 * CLng(((OCC / OKC) * 4)), 48, vbSrcAnd
BitBlt Picture1.Hdc, Missions_HUD_X, Missions_HUD_Y, 24, 24, Form2.HUD(1).Hdc, 24 * CLng(((OCC / OKC) * 4)), 48, vbSrcPaint
Picture1.Print OCC

If OCC >= OKC And OKC > 0 And OCC > 0 And MainChar.Health > 0 Then

c = 0
For j = 1 To UBound(AI)
If AI(j).Health <= 0 And AI(j).AnimFrame >= 8 Then c = c + 1
Next j

For i = 1 To 9
PL_Weapons(i).Ammo = MainChar.Weapons(i).Ammo
PL_Weapons(i).Have = MainChar.Weapons(i).Have
Next i

SetMusicType "PASS_LEVEL", True
Menu_Msg(0).Visible = True
Menu_Msg(1).Visible = True
Menu_Msg(2).Visible = False
Msg_DoorOpen = True
Run = False
DeadTime = 0
StartGame = False
Msg_Type = 6
Msg_LevelName = "Level Complete - Mission Time: " & MissionTime & " Sec - Par Time: " & Level.Mission_Par_Time & " Sec"
Msg(0).Tag = "Mission Accomplished" & Chr(13) & Level.Mission_Pass_Text
ScoreTab = Empty
ScoreTab.Visible = True
ScoreTab = "Taliban Scumbags Killed: " & c & " of " & UBound(AI) & Chr(13) & _
           "Key Objectives Completed: " & (OCC) & " of " & UBound(Objective) & Chr(13) & _
           CLng((c + OCC) / ((UBound(AI)) + UBound(Objective)) * 100) & "% - Overall Completed"
If CLng((c + OCC) / ((UBound(AI)) + UBound(Objective)) * 100) = 100 Then ScoreTab = ScoreTab & Chr(13) & Level.Mission_RevealCheat
Menu_Msg(0).Caption = "Continue": Menu_Msg(0).Visible = True
Menu_Msg(1).Caption = "Exit": Menu_Msg(1).Visible = True
Msg_Door.Enabled = True
LevelSel = LevelSel + 1

If LevelSel > MaxLevels Then
Msg_Type = 7
Menu_Msg(0).Caption = "Credits": Menu_Msg(0).Visible = True
Menu_Msg(1).Caption = "Exit": Menu_Msg(1).Visible = True
LevelSel = 1
End If
End If

End Function

Private Sub Timer1_Timer()
Set_FPS = FPS
FPS = 0
End Sub


Function SetMusicType(ByVal Interface_Type As String, ByVal Play As Boolean)
On Error Resume Next
Dim MN As String
If MediaPlayer1.Tag <> Interface_Type Then
MediaPlayer1.Tag = Interface_Type

If Play = False Then MediaPlayer1.Stop
Select Case UCase(Interface_Type)
Case Is = "INGAME"
MN = "\ci-offical.mid"
Case Is = "INGAME_LOWTIME"
MN = "\ci-fastmix.mid"
Case Is = "START_LEVEL"
MN = "\startlevel.mid"
Case Is = "PASS_LEVEL"
MN = "\endlevel.mid"
Case Is = "FAIL_LEVEL"
MN = "\failed.mid"
Case Is = "MENU"
MN = "\cactus.mid"
Case Is = "INGAME_BOSS"
MN = "\metallica_fourhorsemen.mid"
Case Is = "INGAME_BOSS_OSAMA"
MN = "\cradle.mid"
Case Is = "INGAME_BOSS_OSAMA_LOWHEALTH"
MN = "\cradleX.mid"
Case Is = "CREDITS"
MN = "\metallica_unforgiven.mid"
End Select
MediaPlayer1.FileName = App.Path & MN

If Play = True Then MediaPlayer1.Play
End If

End Function

Private Function LoadLevelToEngine()
'Loads the map file to the game
Dim i As Long
Msg_Door.Enabled = False
Msg_DoorOpen = False
BT.Enabled = True
DoEvents
LoadMapFile LevelSel
DoorDelay = Len(Level.Name) * 2
Msg(0).Tag = Empty
Msg(0).Tag = "Loading... " & Chr(13) & Level.Name
DoEvents
For i = 0 To Menu_Msg.UBound
Menu_Msg(i).Visible = False: Menu_Msg(i).Enabled = False
Next i
Msg_DoorOpen = False
Msg_Door.Enabled = True
Menu_Screen.Enabled = False
Menu_Screen.Visible = False
DoEvents
SetMusicType "INGAME", False
StartGame = True
Run = True
Quake = False
QuakeTime = 0
DeadTime = 0
Game_Delay = 0
DoLoop
End Function

Private Function SetDamage(ByVal Target As String, ByVal Damage As Single, Optional Selected As Long)
If Target = "AI" Then
AI(Selected).Health = AI(Selected).Health - IIf(AI(Selected).Armor > 0, (Damage / (AI(Selected).Armor / 33 + 1)), Damage)
If AI(Selected).Health < 0 Then AI(Selected).Health = 0
AI(Selected).Armor = AI(Selected).Armor - Damage
If AI(Selected).Armor < 0 Then AI(Selected).Armor = 0
ElseIf Target = "MAINCHAR" Then
MainChar.Health = MainChar.Health - IIf(MainChar.Armor > 0, (Damage / (MainChar.Armor / 33 + 1)), Damage)
If MainChar.Health < 0 Then MainChar.Health = 0
MainChar.Armor = MainChar.Armor - Damage
If MainChar.Armor < 0 Then MainChar.Armor = 0
End If
End Function

Private Function CreateMsg(ByVal Showmsg As String, Optional ByVal ID As Long, Optional ByVal Y As Long, Optional ByVal ExtraText As String)
'The Message Creator
Select Case Showmsg
Case Is = "OSAMA_DEATH"
Picture1.FontSize = 9
Picture1.ForeColor = vbBlack
Picture1.CurrentX = Picture1.Width / 6
Picture1.CurrentY = 32
Picture1.Print "Ha Ha Ha, You should of defeated me more effectively, now I'll"
Picture1.CurrentX = Picture1.Width / 6
Picture1.CurrentY = 48
Picture1.Print "send you straight to hell!, Ha Ha Ha Ha"
Picture1.CurrentX = Picture1.Width / 6
Picture1.CurrentY = 64
Picture1.Print "Osama has Armed a Nuclear device, Kill him quickly!"
Picture1.CurrentX = Picture1.Width / 6
Picture1.CurrentY = 80
Picture1.ForeColor = vbRed
Picture1.Print "Time Remaining: " & Abs(AI(ID).Weapon_TimeOut \ Set_Game_Delay)
End Select
Picture1.FontSize = 14
End Function

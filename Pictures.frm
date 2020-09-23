VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form2"
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   707
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox HUD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1800
      Index           =   1
      Left            =   4080
      Picture         =   "Pictures.frx":0000
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   38
      Top             =   5880
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox HUD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1800
      Index           =   0
      Left            =   3960
      Picture         =   "Pictures.frx":0A67
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox BG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   8145
      Index           =   5
      Left            =   7080
      Picture         =   "Pictures.frx":1537
      ScaleHeight     =   543
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   664
      TabIndex        =   37
      Top             =   5640
      Visible         =   0   'False
      Width           =   9960
   End
   Begin VB.PictureBox BG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   8145
      Index           =   4
      Left            =   6960
      Picture         =   "Pictures.frx":44DD
      ScaleHeight     =   543
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   664
      TabIndex        =   36
      Top             =   5520
      Visible         =   0   'False
      Width           =   9960
   End
   Begin VB.PictureBox BG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   8145
      Index           =   3
      Left            =   6840
      Picture         =   "Pictures.frx":6FCD
      ScaleHeight     =   543
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   664
      TabIndex        =   35
      Top             =   5400
      Visible         =   0   'False
      Width           =   9960
   End
   Begin VB.PictureBox BG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   8145
      Index           =   2
      Left            =   6720
      Picture         =   "Pictures.frx":A7A2
      ScaleHeight     =   543
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   664
      TabIndex        =   34
      Top             =   5280
      Visible         =   0   'False
      Width           =   9960
   End
   Begin VB.PictureBox BG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   8145
      Index           =   1
      Left            =   6600
      Picture         =   "Pictures.frx":E130
      ScaleHeight     =   543
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   664
      TabIndex        =   33
      Top             =   5160
      Visible         =   0   'False
      Width           =   9960
   End
   Begin VB.PictureBox Osama 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   7890
      Index           =   1
      Left            =   5160
      Picture         =   "Pictures.frx":11782
      ScaleHeight     =   526
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   32
      Top             =   240
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Osama 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   7890
      Index           =   0
      Left            =   5040
      Picture         =   "Pictures.frx":1484C
      ScaleHeight     =   526
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Effect4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Index           =   1
      Left            =   5040
      Picture         =   "Pictures.frx":179C7
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   30
      Top             =   3240
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.PictureBox Effect4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Index           =   0
      Left            =   4920
      Picture         =   "Pictures.frx":17F5F
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   29
      Top             =   3120
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.PictureBox Baddie5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   1
      Left            =   1920
      Picture         =   "Pictures.frx":184F7
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   28
      Top             =   1920
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Baddie5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   0
      Left            =   1800
      Picture         =   "Pictures.frx":198AC
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   27
      Top             =   1800
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox BaddieBoss1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6750
      Index           =   1
      Left            =   1680
      Picture         =   "Pictures.frx":1AD22
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox BaddieBoss1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6750
      Index           =   0
      Left            =   1560
      Picture         =   "Pictures.frx":1DD0E
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   25
      Top             =   1560
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox BG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   8145
      Index           =   0
      Left            =   6480
      Picture         =   "Pictures.frx":20F1A
      ScaleHeight     =   543
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   664
      TabIndex        =   22
      Top             =   5040
      Visible         =   0   'False
      Width           =   9960
   End
   Begin VB.PictureBox ItemIcons 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   9120
      Picture         =   "Pictures.frx":21F1D
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   24
      Top             =   5160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   240
      Left            =   9120
      Picture         =   "Pictures.frx":22389
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Items 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   5130
      Index           =   1
      Left            =   3120
      Picture         =   "Pictures.frx":227A6
      ScaleHeight     =   342
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox Items 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   5130
      Index           =   0
      Left            =   3000
      Picture         =   "Pictures.frx":232C8
      ScaleHeight     =   342
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox Baddie4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   1
      Left            =   1440
      Picture         =   "Pictures.frx":23EB1
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Baddie4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   0
      Left            =   1320
      Picture         =   "Pictures.frx":25109
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Baddie3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   1
      Left            =   1200
      Picture         =   "Pictures.frx":25CF4
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Baddie3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   0
      Left            =   1080
      Picture         =   "Pictures.frx":2705B
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Effect3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Index           =   1
      Left            =   3480
      Picture         =   "Pictures.frx":27C21
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox Effect3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Index           =   0
      Left            =   3360
      Picture         =   "Pictures.frx":281CA
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox Baddie2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   1
      Left            =   960
      Picture         =   "Pictures.frx":28773
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Baddie2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   0
      Left            =   840
      Picture         =   "Pictures.frx":297BB
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Effect2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   3840
      Index           =   1
      Left            =   2040
      Picture         =   "Pictures.frx":2A393
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1536
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   23040
   End
   Begin VB.PictureBox Effect2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   3840
      Index           =   0
      Left            =   960
      Picture         =   "Pictures.frx":2BD10
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1536
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   23040
   End
   Begin VB.PictureBox Baddie1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   1
      Left            =   720
      Picture         =   "Pictures.frx":2D68D
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Baddie1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   0
      Left            =   600
      Picture         =   "Pictures.frx":2E8F4
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Effect1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   960
      Index           =   1
      Left            =   3120
      Picture         =   "Pictures.frx":2F495
      ScaleHeight     =   960
      ScaleWidth      =   5760
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.PictureBox Effect1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   960
      Index           =   0
      Left            =   3000
      Picture         =   "Pictures.frx":2FD2F
      ScaleHeight     =   960
      ScaleWidth      =   5760
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.PictureBox Bullets 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   9120
      Picture         =   "Pictures.frx":305C9
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   10710
      Index           =   1
      Left            =   240
      Picture         =   "Pictures.frx":30B13
      ScaleHeight     =   714
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.PictureBox Map 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   720
      Left            =   5160
      Picture         =   "Pictures.frx":33EC8
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   235
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   8520
      Left            =   2160
      Picture         =   "Pictures.frx":3C3CA
      ScaleHeight     =   568
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   888
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   13320
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   10710
      Index           =   0
      Left            =   120
      Picture         =   "Pictures.frx":8FFE7
      ScaleHeight     =   714
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2940
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

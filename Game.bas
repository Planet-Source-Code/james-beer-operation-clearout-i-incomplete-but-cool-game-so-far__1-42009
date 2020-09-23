Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'IMPORTANT OPTIMIZATION TIPS:
'* Use global declarations on any variable or array
'if it's used by several forms as it boosts the program's
'speed but takes more memory.
'* Use DIM instead of Private to make private varible to also
'aid the program's speed
'* Have Option Explicit written.
'* Use Long Integers for as many mathematical operations
'requiring integers as this boosts speed.

Type file_info
Wall_Size As Integer
Backprop_Size As Integer
Background_Size As Integer
Foreprop_Size As Integer
Foreground_Size As Integer
AI_Size As Integer
Item_Size As Integer
Swicth_Size As Integer
Objective_Size As Integer
End Type

Type file_info2
Wall_Size As Integer
Backprop_Size As Integer
Background_Size As Integer
Foreprop_Size As Integer
Foreground_Size As Integer
AI_Size As Integer
Item_Size As Integer
Swicth_Size As Integer
Objective_Size As Integer
End Type

Global FileInfo As file_info
Global NewFileInfo As file_info2

Global UserInput As Boolean

Global MissionTime As Long
Global BGSet As Long

Global Game_Offset_X As Long
Global Game_Offset_Y As Long

Global Game_Offset_X2 As Long
Global Game_Offset_Y2 As Long

Global Max_Y_Depth As Long

Global Game_Difficulty As String
Global Cheats(7) As Boolean
Global Cheat_Unlock(7) As Boolean

Type LevelInfo
Name As String
Artist As String
Mission_Berifing As String
Mission_Fail_Text As String
Mission_Pass_Text As String
Mission_RevealCheat As String
Mission_Par_Time As Long
Offical As Boolean
LevelID As String
End Type

Global LevelSel As Long
Global Level As LevelInfo

Private Type Block
Use As Boolean 'Use will now be to set clipping, not general use
X As Long
Y As Long
Width As Long
Height As Long
PX As Long
PY As Long
AnimFrame As Byte
MaxFrame As Byte
Count As Byte
MaxCount As Byte
Water As Boolean
Hazardous As Boolean
Damage As Single
End Type

Private Type Back_Block
Use As Boolean
X As Long
Y As Long
Width As Long
Height As Long
PX As Long
PY As Long
AnimFrame As Byte
MaxFrame As Byte
Count As Byte
MaxCount As Byte
End Type

Type Weapon_Info
Name As String
Bullet_Draw_X As Long
Bullet_Draw_Y As Long
Bullet_Draw_Width As Byte
Bullet_Draw_Height As Byte
Bullet_Speed As Byte
Ammo As Long
Speed As Long
Damage As Long
Effect As Byte
TimeOut_Rate As Long
SetBurnOutTime As Byte
Have As Boolean
End Type

Type Points
X As Long
Wait As Long
Max_Wait As Long
Stop As Boolean
End Type

Type Char_Info
X As Long
Y As Long
Width As Long
Height As Long
Squished As Boolean
GravityForce As Single
Ground As Boolean

Health As Single
Armor As Single

Keys(4) As Boolean
Key_Lock(1) As Boolean
Weapons(9) As Weapon_Info
Weapon_TimeOut(9) As Long
Walk_Tag As Byte
Attack_Tag As Byte
Gun_Tag As Byte
Gun_Draw_Tag As Byte
Direction_Tag As Byte
Walk_Delay As Long

MoveSpeed As Long

Contact As Boolean

AnimFrame As Long
AnimCount As Long

ItemSel As Byte

JetPack_Mode As Boolean
JetPack_Fuel As Single
JetPack_Have As Boolean
JetPack_Anim As Byte
End Type

Type AI_Char_Info
Active As Boolean
X As Long
Y As Long
Width As Long
Height As Long
Squished As Boolean
GravityForce As Single
Ground As Boolean

Boss As Boolean
Health As Single
MaxHealth As Single
MaxArmor As Single
Armor As Single

Keys(4) As Boolean
Key_Lock(1) As Boolean
Weapon_TimeOut As Long
Walk_Tag As Byte
Attack_Tag As Byte
Gun_Tag As Byte
Gun_Draw_Tag As Byte
Direction_Tag As Byte
Walk_Delay As Long
MoveSpeed As Long
Contact As Boolean

AnimFrame As Long
AnimCount As Long
AI_Type  As String

Point As Long
SetPoint(1) As Points

End Type


Type Bullet_Info
From As Byte
Used As Boolean
Speed As Long
Weapon As Byte
BurnOut As Boolean
BurnOutTime As Byte
X As Long
Y As Long
Damage As Single
Effect As Byte
End Type

Type Effect_Info
X As Long
Y As Long
w As Long
h As Long
Damage As Single
Range As Single
Used As Boolean
Effect As Byte
Frame As Long
Times As Long
Count As Long
From As Byte
End Type

Type Item_Info
X As Long
Y As Long
Draw_X As Long
Draw_Y As Long
Width As Long
Height As Long
Item As String
Used As Boolean
End Type

Type Swicth_Info
Enabled As Boolean

Linked As Boolean 'If it uses another swicth's information
LinkSwicthTo As Long 'If this swicth requires to get/set data
'from an existing swicth then this is needed

X As Long
Y As Long
Width As Long
Height As Long
PX(1) As Long
PY(1) As Long

EnterSet As Boolean 'If by Entering in the Block it triggers
PressedSet As Boolean 'If by pressing ALT in the block it triggers
Enter As Boolean 'If entered
Pressed As Boolean 'If shot at

ResetOnReverseTime As Boolean

On As Boolean 'If the swicth is activated
Times As Long 'Number of times to play command
Max_Times As Long
Reverse As Boolean 'Opposite the command second time (Only for MOVE)

'Sometimes a command need to be sent as well,  EG. A airstrike
'beacon should Set On an AI to bomb an certain area.
Command As String
Target As String
Selected As Long
By_X As Long
By_Y As Long
End Type

Type BasicWeaponInfo
Have As Boolean
Ammo As Long
End Type

Global PL_Weapons(1 To 9) As BasicWeaponInfo

Type Objective_Info

TimeLimit As Boolean
TimeLeft As Long 'every game second a tick is deducted
'so on different machines it might vary.

Condition As String
X As Long
Y As Long
Width As Long
Height As Long

Selected As Long
SetKillCount As Long
Key As Boolean
Completed As Boolean
Failed As Boolean
'In Order to pass the level all objectives with Key set as true
'need to be completed, and failing one of them will end the mission too.
End Type




Type SaveScoreTabs
AI_Killed As Long
Time_Elapsed As Long
Objectives As Long
Percentage As Long
End Type

Type SavedCampaignData
Scores() As SaveScoreTabs
Level As Long
InProgress As Boolean 'Determins if the level was incomplete that
'was saved, therefore it skips the level select screen and
'straight to the game itself.
End Type


Global Const MaxLevels As Long = 3
Global Const Health_HUD_X As Long = 2
Global Const Health_HUD_Y As Long = 2
Global Const Armor_HUD_X As Long = 2
Global Const Armor_HUD_Y As Long = 26
Global Const Missions_HUD_X As Long = 2
Global Const Missions_HUD_Y As Long = 96
Global Const TalibanKills_HUD_X As Long = 2
Global Const TalibanKills_HUD_Y As Long = 72
Global Const Time_HUD_X As Long = 300
Global Const Time_HUD_Y As Long = 2
Global Const MissionTime_HUD_X As Long = 2
Global Const MissionTime_HUD_Y As Long = 50
Global Const Set_Game_Delay As Long = 48
Global Const Moderator As Long = 15
Global Const Osama_Low_Health As Long = 250

Global SelectedCampaign As Long

Global OsamaID As Long
Global StartGame As Boolean
Global OsamaTaunt As Boolean
Global DoorDelay As Long
Global Msg_DoorOpen As Boolean
Global Nuke_Countdown As Long
Global NukeSet As Boolean
Global Nuking As Boolean
Global NukeTime As Long

Global Objective() As Objective_Info
Global T_Effect(500) As Effect_Info
Global Wall() As Block

Global WallReg() As Long 'gets wall regions of clipping
'area and uses it for collision detection

Global BackProp() As Back_Block 'NOTE: Transparent sprite walls cannot be solid
Global Background() As Back_Block
Global Foreground() As Back_Block
Global ForeProp() As Back_Block
Global Swicth() As Swicth_Info
Global AI() As AI_Char_Info
Global Item() As Item_Info
Global MainChar As Char_Info
Global Bullet(500) As Bullet_Info

Function LoadMapFile(LevelSel As Long)
Dim i As Long
On Error GoTo MainErr
Erase Bullet
Erase Wall
Erase Background
Erase Foreground
Erase BackProp
Erase ForeProp
Erase Item
Erase AI
Erase Swicth
Erase Objective
Erase WallReg

Level.Name = Empty
Level.Artist = Empty
Level.Mission_Berifing = Empty
Level.Mission_Par_Time = 0
Level.Mission_Fail_Text = Empty
Level.Mission_Pass_Text = Empty
Level.Mission_RevealCheat = Empty

NukeSet = False
Nuking = False
NukeTime = 0

Open App.Path & "\level" & LevelSel & ".map" For Input As #1
Close #1

Open App.Path & "\level" & LevelSel & ".map" For Binary Access Read As #1
Get #1, , NewFileInfo
ReDim Wall(NewFileInfo.Wall_Size)
ReDim BackProp(NewFileInfo.Backprop_Size)
ReDim Background(NewFileInfo.Background_Size)
ReDim ForeProp(NewFileInfo.Foreprop_Size)
ReDim Foreground(NewFileInfo.Foreground_Size)
ReDim AI(NewFileInfo.AI_Size)
ReDim Item(NewFileInfo.Item_Size)
ReDim Swicth(NewFileInfo.Swicth_Size)
ReDim Objective(NewFileInfo.Objective_Size)

Get #1, , Wall
Get #1, , BackProp
Get #1, , Background
Get #1, , ForeProp
Get #1, , Foreground
Get #1, , MainChar
Get #1, , AI
Get #1, , Item
Get #1, , Swicth
Get #1, , Objective
Get #1, , Level
Get #1, , Max_Y_Depth
Get #1, , BGSet
Close #1


If Level.Offical = True Then
If Level.LevelID = "OC1_LEVEL_" & LevelSel Then
Else
   MsgBox "The Game detects a corrupted Offical Level ID and will not continue", vbExclamation, "Offical Map Tampered"
   End
End If
End If

If LevelSel > 1 And Level.Offical = True Then
'after the first offical map, you only get the weapons
'you already got from the last level
For i = 1 To 9
MainChar.Weapons(i).Ammo = PL_Weapons(i).Ammo
MainChar.Weapons(i).Have = PL_Weapons(i).Have
Next i
End If

SetWeapons

For i = 1 To UBound(AI)
CreateAI CLng(i), AI(i).AI_Type, AI(i).X, AI(i).Y, CLng(AI(i).Point), AI(i).SetPoint(0).X, AI(i).SetPoint(1).X, AI(i).SetPoint(0).Max_Wait, AI(i).Active
Next i
For i = 1 To UBound(Item)
CreateItem CLng(i), Item(i).Item, Item(i).X, Item(i).Y, Item(i).Used
Next i

For i = 1 To UBound(Swicth)
Swicth(i).Times = Swicth(i).Max_Times
If Swicth(i).Command = "SET OFF" Then Swicth(i).By_Y = -1
Next i

MainChar.GravityForce = 0
MainChar.Direction_Tag = 0
Game_Offset_X = -((Form1.Picture1.Width / 2) - MainChar.X)
Game_Offset_Y = -((Form1.Picture1.Height / 2) - MainChar.Y)
MainChar.Width = 32
MainChar.Height = 34
MainChar.MoveSpeed = 6
MainChar.Attack_Tag = 0
MainChar.Gun_Tag = 0
MainChar.Gun_Draw_Tag = 0
MainChar.AnimFrame = 0
MainChar.AnimCount = 0

Exit Function
MainErr:
MsgBox "Level" & LevelSel & ".map" & " was not found or failed to load", vbCritical + vbDefaultButton1, "Loading..."
End
End Function

Function LoadChar()
On Error Resume Next

MainChar.JetPack_Have = True
MainChar.JetPack_Mode = True
MainChar.JetPack_Fuel = 15
MainChar.GravityForce = 0
MainChar.Health = 100
MainChar.Armor = 0
MainChar.X = 0
MainChar.Y = 664
MainChar.Direction_Tag = 0
Game_Offset_X = -((Form1.Picture1.Width / 2) - MainChar.X)
Game_Offset_Y = -((Form1.Picture1.Height / 2) - MainChar.Y)
MainChar.Width = 32
MainChar.Height = 34
MainChar.MoveSpeed = 6
MainChar.Attack_Tag = 0
MainChar.Gun_Tag = 0
MainChar.Gun_Draw_Tag = 0
MainChar.AnimFrame = 0
MainChar.AnimCount = 0
MainChar.Weapons(1).Ammo = 100
MainChar.Weapons(1).Have = True


CreateAI 1, "TALIBAN_1", 850, 862, 1, 850 - 480, 850 + 120, 100, True
CreateAI 2, "TALIBAN_1", 800, 862, 1, 800 - 480, 800 + 120, 100, True
CreateAI 3, "TALIBAN_1", 900, 862, 1, 900 - 480, 900 + 120, 100, True
CreateAI 4, "TALIBAN_2", 950, 862, 1, 950 - 480, 950 + 120, 100, True
CreateAI 5, "TALIBAN_2", 1000, 862, 1, 1000 - 480, 1000 + 120, 100, True
CreateAI 6, "TALIBAN_5", 1700, 1016, 1, 1600, 1800, 100, True
CreateAI 7, "TALIBAN_BOSS1", 7000, 1210, 1, 6500, 7200, 50, True

End Function

Function SetWeapons()

MainChar.Weapons(0).Damage = 4
MainChar.Weapons(0).Effect = 0
MainChar.Weapons(0).SetBurnOutTime = 3
MainChar.Weapons(0).Speed = 0
MainChar.Weapons(0).TimeOut_Rate = 10
MainChar.Weapons(0).Bullet_Draw_X = 0
MainChar.Weapons(0).Bullet_Draw_Y = 0
MainChar.Weapons(0).Bullet_Draw_Width = 0
MainChar.Weapons(0).Bullet_Draw_Height = 0
MainChar.Weapons(0).Name = "Combat Knife"

MainChar.Weapons(1).Damage = 5
MainChar.Weapons(1).Effect = 0
MainChar.Weapons(1).SetBurnOutTime = 0
MainChar.Weapons(1).Speed = 10
MainChar.Weapons(1).TimeOut_Rate = 16
MainChar.Weapons(1).Bullet_Draw_X = 0
MainChar.Weapons(1).Bullet_Draw_Y = 0
MainChar.Weapons(1).Bullet_Draw_Width = 6
MainChar.Weapons(1).Bullet_Draw_Height = 2
MainChar.Weapons(1).Name = "9mm Pistol"

MainChar.Weapons(2).Damage = 6
MainChar.Weapons(2).Effect = 0
MainChar.Weapons(2).SetBurnOutTime = 0
MainChar.Weapons(2).Speed = 12
MainChar.Weapons(2).TimeOut_Rate = 7
MainChar.Weapons(2).Bullet_Draw_X = 0
MainChar.Weapons(2).Bullet_Draw_Y = 2
MainChar.Weapons(2).Bullet_Draw_Width = 6
MainChar.Weapons(2).Bullet_Draw_Height = 2
MainChar.Weapons(2).Name = "AK33 Assault Rifle"


MainChar.Weapons(3).Damage = 4
MainChar.Weapons(3).Effect = 0
MainChar.Weapons(3).SetBurnOutTime = 12
MainChar.Weapons(3).Speed = 15
MainChar.Weapons(3).TimeOut_Rate = 3
MainChar.Weapons(3).Bullet_Draw_X = 0
MainChar.Weapons(3).Bullet_Draw_Y = 4
MainChar.Weapons(3).Bullet_Draw_Width = 6
MainChar.Weapons(3).Bullet_Draw_Height = 2
MainChar.Weapons(3).Name = "Pulsar Rifle"


MainChar.Weapons(4).Damage = 2
MainChar.Weapons(4).Effect = 3
MainChar.Weapons(4).SetBurnOutTime = 12
MainChar.Weapons(4).Speed = 8
MainChar.Weapons(4).TimeOut_Rate = 3
MainChar.Weapons(4).Bullet_Draw_X = 0
MainChar.Weapons(4).Bullet_Draw_Y = 16
MainChar.Weapons(4).Bullet_Draw_Width = 7
MainChar.Weapons(4).Bullet_Draw_Height = 4
MainChar.Weapons(4).Name = "Flamethrower"

MainChar.Weapons(5).Damage = 120
MainChar.Weapons(5).Effect = 1
MainChar.Weapons(5).SetBurnOutTime = 0
MainChar.Weapons(5).Speed = 16
MainChar.Weapons(5).TimeOut_Rate = 25
MainChar.Weapons(5).Bullet_Draw_X = 0
MainChar.Weapons(5).Bullet_Draw_Y = 6
MainChar.Weapons(5).Bullet_Draw_Width = 12
MainChar.Weapons(5).Bullet_Draw_Height = 3
MainChar.Weapons(5).Name = "RPG Standard Launcher"

MainChar.Weapons(6).Damage = 160
MainChar.Weapons(6).Effect = 1
MainChar.Weapons(6).SetBurnOutTime = 0
MainChar.Weapons(6).Speed = 21
MainChar.Weapons(6).TimeOut_Rate = 40
MainChar.Weapons(6).Bullet_Draw_X = 0
MainChar.Weapons(6).Bullet_Draw_Y = 9
MainChar.Weapons(6).Bullet_Draw_Width = 12
MainChar.Weapons(6).Bullet_Draw_Height = 3
MainChar.Weapons(6).Name = "Anti-Tank Rocket Launcher"

MainChar.Weapons(7).Damage = 500
MainChar.Weapons(7).Effect = 2
MainChar.Weapons(7).SetBurnOutTime = 0
MainChar.Weapons(7).Speed = 7
MainChar.Weapons(7).TimeOut_Rate = 100
MainChar.Weapons(7).Bullet_Draw_X = 0
MainChar.Weapons(7).Bullet_Draw_Y = 12
MainChar.Weapons(7).Bullet_Draw_Width = 6
MainChar.Weapons(7).Bullet_Draw_Height = 4
MainChar.Weapons(7).Name = "Pulse Cannon"

MainChar.Weapons(8).Damage = 1
MainChar.Weapons(8).Effect = 4
MainChar.Weapons(8).SetBurnOutTime = 16
MainChar.Weapons(8).Speed = 8
MainChar.Weapons(8).TimeOut_Rate = 5
MainChar.Weapons(8).Bullet_Draw_X = 0
MainChar.Weapons(8).Bullet_Draw_Y = 20
MainChar.Weapons(8).Bullet_Draw_Width = 8
MainChar.Weapons(8).Bullet_Draw_Height = 12
MainChar.Weapons(8).Name = "Chemical Launcher"

End Function


Function CreateAI(i As Long, AI_Type As String, X As Long, Y As Long, Point As Byte, X1 As Long, X2 As Long, WaitTime As Long, Active As Boolean)
OsamaID = 0
If Cheats(3) = True Then Game_Difficulty = "NONE": MainChar.Health = 1: MainChar.Armor = 0

AI(i).Active = Active
AI(i).AI_Type = AI_Type
AI(i).X = X
AI(i).Y = Y
AI(i).GravityForce = 0
AI(i).AnimCount = 0
AI(i).AnimFrame = 0

Select Case AI_Type
Case Is = "TALIBAN_1"
    Select Case UCase(Game_Difficulty)
        Case Is = "EASY"
        AI(i).Health = 10
        AI(i).Armor = 0
        Case Is = "NORMAL"
        AI(i).Health = 20
        AI(i).Armor = 0
        Case Is = "HARD"
        AI(i).Health = 25
        AI(i).Armor = 20
        Case Is = "NONE"
        AI(i).Health = 1
    End Select
AI(i).Width = 32
AI(i).Height = 34
AI(i).MoveSpeed = 6
AI(i).Gun_Tag = 1
AI(i).Gun_Draw_Tag = 0
Case Is = "TALIBAN_2"
    Select Case UCase(Game_Difficulty)
        Case Is = "EASY"
        AI(i).Health = 20
        AI(i).Armor = 0
        Case Is = "NORMAL"
        AI(i).Health = 30
        AI(i).Armor = 0
        Case Is = "HARD"
        AI(i).Health = 35
        AI(i).Armor = 20
        Case Is = "NONE"
        AI(i).Health = 1
    End Select
AI(i).Width = 32
AI(i).Height = 34
AI(i).MoveSpeed = 6
AI(i).Gun_Tag = 3
AI(i).Gun_Draw_Tag = 0
Case Is = "TALIBAN_3"
    Select Case UCase(Game_Difficulty)
        Case Is = "EASY"
        AI(i).Health = 30
        AI(i).Armor = 0
        Case Is = "NORMAL"
        AI(i).Health = 40
        AI(i).Armor = 0
        Case Is = "HARD"
        AI(i).Health = 45
        AI(i).Armor = 20
        Case Is = "NONE"
        AI(i).Health = 1
    End Select
AI(i).Width = 32
AI(i).Height = 34
AI(i).MoveSpeed = 6
AI(i).Gun_Tag = 4
AI(i).Gun_Draw_Tag = 0
Case Is = "TALIBAN_4"
    Select Case UCase(Game_Difficulty)
        Case Is = "EASY"
        AI(i).Health = 40
        AI(i).Armor = 0
        Case Is = "NORMAL"
        AI(i).Health = 50
        AI(i).Armor = 0
        Case Is = "HARD"
        AI(i).Health = 55
        AI(i).Armor = 20
        Case Is = "NONE"
        AI(i).Health = 1
    End Select
AI(i).Width = 32
AI(i).Height = 34
AI(i).MoveSpeed = 6
AI(i).Gun_Tag = 5
AI(i).Gun_Draw_Tag = 0
Case Is = "TALIBAN_5"
    Select Case UCase(Game_Difficulty)
        Case Is = "EASY"
        AI(i).Health = 30
        AI(i).Armor = 0
        Case Is = "NORMAL"
        AI(i).Health = 40
        AI(i).Armor = 0
        Case Is = "HARD"
        AI(i).Health = 45
        AI(i).Armor = 20
        Case Is = "NONE"
        AI(i).Health = 1
    End Select
AI(i).Width = 32
AI(i).Height = 34
AI(i).MoveSpeed = 6
AI(i).Gun_Tag = 8
AI(i).Gun_Draw_Tag = 0
Case Is = "TALIBAN_BOSS1"

    Select Case UCase(Game_Difficulty)
        Case Is = "EASY"
        AI(i).Health = 100
        AI(i).Armor = 1000
        Case Is = "NORMAL"
        AI(i).Health = 200
        AI(i).Armor = 1500
        Case Is = "HARD"
        AI(i).Health = 300
        AI(i).Armor = 2000
        Case Is = "NONE"
        AI(i).Health = 100
    End Select

AI(i).Boss = True
AI(i).Width = 45
AI(i).Height = 45
AI(i).MoveSpeed = 4
AI(i).Gun_Tag = 6
AI(i).Gun_Draw_Tag = 0
Case Is = "TALIBAN_BOSS_OSAMA"
OsamaID = i
    Select Case UCase(Game_Difficulty)
        Case Is = "EASY"
        AI(i).Health = 300
        AI(i).Armor = 2000
        Case Is = "NORMAL"
        AI(i).Health = 400
        AI(i).Armor = 3500
        Case Is = "HARD"
        AI(i).Health = 500
        AI(i).Armor = 5000
        Case Is = "NONE"
        AI(i).Health = 300
    End Select

AI(i).Boss = True
AI(i).Width = 32
AI(i).Height = 53
AI(i).MoveSpeed = 4
AI(i).Gun_Tag = 7
AI(i).Gun_Draw_Tag = 0
End Select

AI(i).Point = Point
AI(i).SetPoint(0).X = X1
AI(i).SetPoint(0).Wait = 0
AI(i).SetPoint(0).Max_Wait = WaitTime

AI(i).SetPoint(1).X = X2
AI(i).SetPoint(1).Wait = 0
AI(i).SetPoint(1).Max_Wait = WaitTime

AI(i).MaxHealth = AI(i).Health
AI(i).MaxArmor = AI(i).Armor
End Function

Function CreateItem(i As Long, StrItem As String, X As Long, Y As Long, IsUsed As Boolean)
Item(i).Item = StrItem
Item(i).Used = IsUsed
Item(i).X = X
Item(i).Y = Y
    Select Case StrItem
    Case Is = "PISTOL"
        Item(i).Draw_X = 0
        Item(i).Draw_Y = 32 * 0
        Item(i).Width = 64
        Item(i).Height = 32
    Case Is = "PISTOL_AMMO"
        Item(i).Draw_X = 64
        Item(i).Draw_Y = 32 * 0
        Item(i).Width = 32
        Item(i).Height = 32
    Case Is = "AK33"
        Item(i).Draw_X = 0
        Item(i).Draw_Y = 32 * 1
        Item(i).Width = 64
        Item(i).Height = 32
    Case Is = "AK33_AMMO"
        Item(i).Draw_X = 64
        Item(i).Draw_Y = 32 * 1
        Item(i).Width = 32
        Item(i).Height = 32
    Case Is = "PULSAR"
        Item(i).Draw_X = 0
        Item(i).Draw_Y = 32 * 2
        Item(i).Width = 64
        Item(i).Height = 32
    Case Is = "PULSAR_AMMO"
        Item(i).Draw_X = 64
        Item(i).Draw_Y = 32 * 2
        Item(i).Width = 32
        Item(i).Height = 32
    Case Is = "SLIME"
        Item(i).Draw_X = 0
        Item(i).Draw_Y = 32 * 3
        Item(i).Width = 64
        Item(i).Height = 32
    Case Is = "SLIME_AMMO"
        Item(i).Draw_X = 64
        Item(i).Draw_Y = 32 * 3
        Item(i).Width = 32
        Item(i).Height = 32
    Case Is = "FLAME"
        Item(i).Draw_X = 0
        Item(i).Draw_Y = 32 * 4
        Item(i).Width = 64
        Item(i).Height = 32
    Case Is = "FLAME_AMMO"
        Item(i).Draw_X = 64
        Item(i).Draw_Y = 32 * 4
        Item(i).Width = 32
        Item(i).Height = 32
    Case Is = "RGB_DEVICE"
        Item(i).Draw_X = 0
        Item(i).Draw_Y = 32 * 5
        Item(i).Width = 64
        Item(i).Height = 32
    Case Is = "GRENADE"
        Item(i).Draw_X = 64
        Item(i).Draw_Y = 32 * 5
        Item(i).Width = 32
        Item(i).Height = 32
    Case Is = "ROCKET"
        Item(i).Draw_X = 64
        Item(i).Draw_Y = 32 * 6
        Item(i).Width = 32
        Item(i).Height = 32
    Case Is = "NUKEBLAST"
        Item(i).Draw_X = 64
        Item(i).Draw_Y = 32 * 7
        Item(i).Width = 32
        Item(i).Height = 32
    Case Is = "10HEALTH"
        Item(i).Draw_X = 0
        Item(i).Draw_Y = 32 * 6
        Item(i).Width = 64
        Item(i).Height = 32
    Case Is = "25HEALTH"
        Item(i).Draw_X = 0
        Item(i).Draw_Y = 32 * 7
        Item(i).Width = 64
        Item(i).Height = 32
    Case Is = "JETPACK"
        Item(i).Draw_X = 0
        Item(i).Draw_Y = 32 * 8
        Item(i).Width = 32
        Item(i).Height = 32
    Case Is = "ARMOR"
        Item(i).Draw_X = 0
        Item(i).Draw_Y = 32 * 9
        Item(i).Width = 32
        Item(i).Height = 32
End Select
End Function


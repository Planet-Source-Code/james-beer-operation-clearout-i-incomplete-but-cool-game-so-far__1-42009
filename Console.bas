Attribute VB_Name = "Console"
'These are commands reponsed by a swicth's command name
'this creates the moving and disappearing effects


Public Function Move(Target As String, Selected As Long, By_X As Long, By_Y As Long)
Select Case Target
Case Is = "BACKGROUND"
Background(Selected).X = Background(Selected).X + By_X
Background(Selected).Y = Background(Selected).Y + By_Y
Case Is = "WALL"

Wall(Selected).X = Wall(Selected).X + By_X
Wall(Selected).Y = Wall(Selected).Y + By_Y

For i = 0 To UBound(AI)
If AI(i).Y + AI(i).Height >= Wall(Selected).Y - (By_Y + 1) And _
   AI(i).X >= Wall(Selected).X - AI(i).Width And _
   AI(i).X <= Wall(Selected).X + Wall(Selected).Width And _
   AI(i).Y <= Wall(Selected).Y + Wall(Selected).Height + By_Y Then
   
AI(i).X = AI(i).X + By_X

If AI(i).X >= Wall(Selected).X - (Wall(Selected).Width / 2) And _
   AI(i).X <= Wall(Selected).X + (Wall(Selected).Width / 2) Then
If By_Y < 0 Then
AI(i).Y = AI(i).Y + By_Y
Else

    If AI(i).Y > Wall(Selected).Y Then
    AI(i).GravityForce = (By_Y * 2)
    Else
    AI(i).GravityForce = By_Y
    End If

End If
End If

End If
Next i


If MainChar.Y + MainChar.Height >= Wall(Selected).Y - (By_Y) And _
   MainChar.X >= Wall(Selected).X - MainChar.Width And _
   MainChar.X <= Wall(Selected).X + Wall(Selected).Width And _
   MainChar.Y <= Wall(Selected).Y + Wall(Selected).Height + By_Y Then

MainChar.X = MainChar.X + By_X

If MainChar.X >= Wall(Selected).X - (Wall(Selected).Width / 2) And _
   MainChar.X <= Wall(Selected).X + (Wall(Selected).Width / 2) Then

MainChar.Y = MainChar.Y + By_Y
End If
End If

If MainChar.X + MainChar.Width > Wall(Selected).X And _
   MainChar.X < Wall(Selected).X + Wall(Selected).Width And _
   MainChar.Y >= Wall(Selected).Y + 4 And _
   MainChar.Y <= Wall(Selected).Y + Wall(Selected).Height + By_Y And _
   MainChar.Ground = True And _
   By_Y < 0 Then
MainChar.Health = 0
MainChar.Squished = True
End If

If MainChar.X + MainChar.Width > Wall(Selected).X And _
   MainChar.X < Wall(Selected).X + Wall(Selected).Width And _
   MainChar.Y >= Wall(Selected).Y + Wall(Selected).Height And _
   MainChar.Y <= Wall(Selected).Y + Wall(Selected).Height + By_Y And _
   MainChar.Ground = True And _
   By_Y > 0 Then
MainChar.Health = 0
MainChar.Squished = True
End If

Case Is = "MAINCHAR"
MainChar.X = By_X
MainChar.Y = By_Y
Case Is = "AI"
AI(Selected).X = By_X
AI(Selected).Y = By_Y
Case Is = "ITEM"
Item(Selected).X = Item(Selected).X + By_X
Item(Selected).Y = Item(Selected).Y + By_Y
End Select
End Function

Public Function SetOff(Target As String, Selected As Long, Effect As Long)
Dim a As Long
Dim X As Long
Dim Y As Long

Select Case Target
Case Is = "BACKGROUND"
Background(Selected).Use = False
X = Background(Selected).X
Y = Background(Selected).Y
Case Is = "BACKPROP"
BackProp(Selected).Use = False
X = BackProp(Selected).X
Y = BackProp(Selected).Y
Case Is = "FOREGROUND"
Foreground(Selected).Use = False
X = Foreground(Selected).X
Y = Foreground(Selected).Y
Case Is = "FOREPROP"
ForeProp(Selected).Use = False
X = ForeProp(Selected).X
Y = ForeProp(Selected).Y
Case Is = "WALL"
Wall(Selected).Use = False
X = Wall(Selected).X
Y = Wall(Selected).Y
Case Is = "ITEM"
Item(Selected).Used = True
X = Item(Selected).X
Y = Item(Selected).Y
Case Is = "AI"
AI(Selected).Active = False
X = AI(Selected).X
Y = AI(Selected).Y
Case Is = "MAINCHAR"
MainChar.Health = 0
X = MainChar.X
Y = MainChar.Y
Case Is = "SWICTH"
Swicth(Selected).Enabled = False
X = Swicth(Selected).X
Y = Swicth(Selected).Y
End Select

If Effect > 0 Then
For i = 0 To UBound(T_Effect)
If T_Effect(i).Used = False Then
    Select Case Effect
        Case Is = 1
    T_Effect(i).w = 64
    T_Effect(i).h = 64
    T_Effect(i).Damage = MainChar.Weapons(6).Damage
        Case Is = 2
    T_Effect(i).w = 128
    T_Effect(i).h = 128
    T_Effect(i).Damage = MainChar.Weapons(7).Damage
        Case Is = 3
    T_Effect(i).w = 32
    T_Effect(i).h = 32
    T_Effect(i).Damage = MainChar.Weapons(4).Damage
        Case Is = 4
    T_Effect(i).w = 32
    T_Effect(i).h = 32
    T_Effect(i).Damage = MainChar.Weapons(8).Damage
    End Select
    T_Effect(i).Used = True
    T_Effect(i).Effect = Effect
    T_Effect(i).From = 2
    T_Effect(i).Times = 0
    T_Effect(i).Count = 0
    T_Effect(i).Frame = 0
    T_Effect(i).X = X - (T_Effect(i).w / 2)
    T_Effect(i).Y = Y - (T_Effect(i).h / 2)
Exit For
End If
Next i
End If

End Function

Public Function SetOn(Target As String, Selected As Long)
Select Case Target
Case Is = "BACKGROUND"
Background(Selected).Use = True
Case Is = "BACKPROP"
BackProp(Selected).Use = True
Case Is = "FOREGROUND"
Foreground(Selected).Use = True
Case Is = "FOREPROP"
ForeProp(Selected).Use = True
Case Is = "WALL"
Wall(Selected).Use = True
Case Is = "ITEM"
Item(Selected).Used = False
Case Is = "MAINCHAR"
MainChar.Health = 100
Case Is = "AI"
AI(Selected).Active = True
Case Is = "SWICTH"
Swicth(Selected).Enabled = True
End Select
End Function

Function KilledCount() As Long
For i = 0 To UBound(AI)
If AI(i).Health <= 0 Then KilledCount = KilledCount + 1
Next i
End Function

Function GetEffectLength(Effect As Long) As Long
Select Case Effect
Case Is = 1
GetEffectLength = 5
Case Is = 2
GetEffectLength = 10
Case Is = 3
GetEffectLength = 4
Case Is = 4
GetEffectLength = 8
Case Else
GetEffectLength = 1
End Select
End Function


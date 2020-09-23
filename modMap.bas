Attribute VB_Name = "modMap"
Option Explicit

Declare Function GetTickCount Lib "kernel32" () As Long

Public Const TileWidth = 18
Public Const TileHeight = 18

Enum nDir
up = 0
down = 2
Left = 3
Right = 1
End Enum

Enum nAI_kinds
Still = 0
Pace = 1
Spin = 2
Random = 3
End Enum

Type PackItem
Caption As String
Value As Long
FuncID As String 'the function is preforms (ie magic dust = +10 magic, or candy = +5 health)
Act As Long
End Type

Type Pack
Item(1 To 20) As PackItem

mDust As Long
money As Long
End Type

Type BattleInfo
Health As Long
Attack As Long
Defense As Long
MaxHealth As Long
Level As Long

Exp As Long
TExp As Long
End Type

Type Player
X As Long
Y As Long

Walking As Boolean
DestX As Long
DestY As Long
WalkC As Long

dir As nDir

B As BattleInfo
Pack As Pack

Step As Long
StepC As Long

SpellCut As Boolean
End Type

Type Tile
X As Long
Y As Long

Type As Long

Walkable As Long
DoorPath As String
DoorWay As Long
DoorDir As Long

SignText As String
isBattle As Long
isMarket As Long
plID As Long

Ani As Long
AniC As Long
End Type

Type Map
Name As String
Midi As String
Mapwidth As Long
Mapheight As Long

MarketPath As String

MaxHP As Long 'the maxium stats of a wild monster in the high grass
MaxAttack As Long
MaxDefense As Long
End Type

Type obj
AI As Long
T As Long
Name As String

Speek() As String
SpeekI As Long

AI_Type As nAI_kinds
dir As Long
Person As Long

Walking As Long
WalkC As Long

X As Long
Y As Long
DestX As Long
DestY As Long
Step As Long
StepC As Long

plID As Long

Act As Long
End Type

Type enemy
HP As Long
MaxHP As Long

Attack As Long
Defense As Long
End Type

Public M As Map
Public T(5000) As Tile
Public P As Player
Public o(1 To 20) As obj
Public MessageTime As Long
Public Enem As enemy
Public E As New ClsEngine
Public FPS As Long

Function LoadMap(path As String, door As Long, dir)
Dim lineread() As String, C() As String, free, tfree, X, Y, func() As String, I
Dim funct As String, u() As String

ClearStuff

Open App.path + "\maps\" & path For Input As #1

Input #1, M.Name
Input #1, M.MarketPath
Input #1, free, tfree

frmGame.market.Picture = LoadPicture(App.path + "\other data\" & free)
frmGame.battle.Picture = LoadPicture(App.path + "\other data\" & tfree)

Input #1, M.MaxHP, M.MaxAttack, M.MaxDefense
Input #1, M.Midi
Input #1, M.Mapwidth, M.Mapheight
Input #1, P.X, P.Y

P.X = P.X * TileWidth
P.Y = P.Y * TileHeight

tfree = ""
free = ""

Do Until EOF(1)

    Line Input #1, free

    If free = "END TILES" Then GoTo DoneTiles

    tfree = tfree & free & vbCrLf

Loop

DoneTiles:

I = 1

Do

    Line Input #1, funct

    If funct = "END OBJECTS" Then GoTo DoneObjs

        u() = Split(funct, ",")
        o(I).Act = True
        o(I).dir = nDir.down
        o(I).Name = u(0)
        o(I).Speek() = Split(u(1), "|")
        o(I).SpeekI = 0
        o(I).dir = Val(u(5))
        o(I).T = Val(u(2))
        o(I).AI = 0
        o(I).AI_Type = Val(u(4))
        o(I).plID = Val(u(3))
        o(I).Walking = False
        o(I).Person = False
        If isValue(o(I).T, 1, 2) = True Then o(I).Person = True

    DoEvents
    I = I + 1

Loop

DoneObjs:

lineread() = Split(tfree, vbCrLf)
I = 0

frmGame.buffer.Width = M.Mapwidth * TileWidth
frmGame.buffer.Height = M.Mapheight * TileHeight

For Y = 1 To M.Mapheight
    
    C() = Split(lineread(Y - 1), "|")
    
    For X = 1 To M.Mapwidth
    
        func() = Split(C(X - 1), ",")
    
        T(I).X = X * TileWidth
        T(I).Y = Y * TileHeight
        T(I).Type = Val(func(0))
        T(I).DoorPath = func(1)
        If InStr(1, func(1), ":") > 1 Then T(I).DoorPath = Split(func(1), ":")(0)
        If Len(T(I).DoorPath) > 1 Then T(I).DoorWay = Val(Split(func(1), ":")(1))
        If Len(T(I).DoorPath) > 1 Then T(I).DoorDir = Val(Split(func(1), ":")(2))
        T(I).SignText = func(2)
        T(I).isMarket = Val(func(3))
        T(I).plID = Val(func(4))
        If T(I).plID = door Then P.X = T(I).X: P.Y = T(I).Y: P.DestX = P.X: P.DestY = P.Y
        T(I).Walkable = True
        If isValue(T(I).Type, 0, 2, 3, 7, 9, 10, 11, 12, 13) = True Then T(I).Walkable = False
    
        I = I + 1
    Next X
    
Next Y

Close #1

SetObjs

P.dir = dir
End Function

Function WalkOn(TileI As Long)
If T(TileI).DoorPath <> "" Then LoadMap T(TileI).DoorPath, T(TileI).DoorWay, T(TileI).DoorDir
If T(TileI).Type = 6 And Int(Rnd * 10) = 5 Then LoadBattle
End Function

Function GetTile() As Long
Dim I As Long

For I = 0 To M.Mapheight * M.Mapwidth
    
    If T(I).X = P.DestX And T(I).Y = P.DestY Then GetTile = I: Exit For

Next I
End Function

Function GetObjTile(Ind) As Long
Dim I As Long

For I = 0 To M.Mapheight * M.Mapwidth

    If T(I).X = o(Ind).DestX And T(I).Y = o(Ind).DestY Then GetObjTile = I: Exit For

Next I
End Function

Function SetObjs() As Long
Dim I As Long, i2

For I = 0 To M.Mapheight * M.Mapwidth

     For i2 = 1 To 20
     
        If o(i2).plID = T(I).plID Then
            o(i2).X = T(I).X
            o(i2).Y = T(I).Y
            o(i2).DestX = o(i2).X
            o(i2).DestY = o(i2).Y
        End If
     
     Next i2
     
Next I
End Function

Function TalkTo()
Dim TileI As Long

Select Case P.dir

    Case nDir.up: TileI = GetTile - M.Mapheight
    Case nDir.down: TileI = GetTile + M.Mapheight
    Case nDir.Left: TileI = GetTile - 1
    Case nDir.Right: TileI = GetTile + 1

End Select

If T(TileI).SignText <> "" Then frmGame.lblM.Caption = T(TileI).SignText: MessageTime = 10 + (Len(T(TileI).SignText) * 1)

If T(TileI).isMarket = -1 Then LoadMarket

If isTileTaken(TileI) = True And MessageTime < 11 Then
    
    Dim H As Long, odir

    H = GetObj(TileI)

    If o(H).Person = True Then
        
        Select Case P.dir
        
            Case nDir.down: odir = nDir.up
            Case nDir.Left: odir = nDir.Right
            Case nDir.Right: odir = nDir.Left
            Case nDir.up: odir = nDir.down
            
        End Select

    o(H).dir = odir
    
End If

    frmGame.lblM.Caption = o(H).Speek(o(H).SpeekI)

    If frmGame.lblM.Caption = "Here I'll teach you the cut spell" Then P.SpellCut = True

    MessageTime = 10 + Len(frmGame.lblM.Caption)
    o(H).SpeekI = o(H).SpeekI + 1: If o(H).SpeekI > UBound(o(H).Speek()) Then o(H).SpeekI = 0

End If
End Function

Function ClearStuff()
Dim I As Long

For I = 0 To 5000

    T(I).X = 0
    T(I).Y = 0
    T(I).Type = 0
    
Next I

For I = 1 To 20

    o(I).Act = False
    
Next I
End Function

Function isValue(Value As Long, ParamArray Nums() As Variant) As Boolean
Dim I As Long

isValue = False

For I = LBound(Nums()) To UBound(Nums())

    If Value = Nums(I) Then isValue = True
    
Next I
End Function

Function CheckInput()
If E.IsPressed(vbKeyLeft) Then

    If P.Walking = False Then P.DestX = P.X - TileWidth: P.Walking = True
    
End If

If E.IsPressed(vbKeyRight) Then

    If P.Walking = False Then P.DestX = P.X + TileWidth: P.Walking = True
    
End If

If E.IsPressed(vbKeyUp) Then

    If P.Walking = False Then P.DestY = P.Y - TileHeight: P.Walking = True
    
End If

If E.IsPressed(vbKeyDown) Then

    If P.Walking = False Then P.DestY = P.Y - TileHeight: P.Walking = True
    
End If
End Function

Function LoadMenu()
Dim I As Long

frmGame.board.Visible = False
frmGame.menu.Visible = True
frmGame.battle.Visible = False
frmGame.main.Visible = False
frmGame.market.Visible = False

frmGame.lstItems.Clear

For I = 1 To 20

    If P.Pack.Item(I).Act = True Then
    
        frmGame.lstItems.AddItem P.Pack.Item(I).Caption
        
    End If

Next I
End Function

Function UseItem(cap As String, I)
Select Case UCase(cap)
    Case "CANDY" 'candy
        P.B.Health = P.B.Health + 5
    Case "MAGIC DUST" 'magic dust
        P.Pack.mDust = P.Pack.mDust + 10
    Case "ENERGY JUICE" 'energy juice
        P.B.Attack = P.B.Attack + 5
    Case "IRON JUICE" 'iron juice
        P.B.Defense = P.B.Defense + 5
End Select

P.Pack.Item(I + 1).Act = False
End Function

Function LoadBattle()
frmGame.tmrAttack.Enabled = False
frmGame.tmrAttack2 = False
frmGame.tmrMagic.Enabled = False
frmGame.imgMagic.Visible = False
frmGame.imgAttack.Visible = False
frmGame.imgAttack2.Visible = False
frmGame.imgEnemy.Picture = frmGame.enemy(Int(Rnd * (frmGame.enemy.UBound + 1))).Picture
frmGame.board.Visible = False
frmGame.menu.Visible = False
frmGame.battle.Visible = True
frmGame.main.Visible = False
frmGame.market.Visible = False

Enem.HP = Int(Rnd * M.MaxHP) + 1
Enem.MaxHP = Enem.HP
Enem.Attack = Int(Rnd * M.MaxAttack) + 1
Enem.Defense = Int(Rnd * M.MaxDefense) + 1
End Function

Function NewGame()
LoadMap "house.dat", 2, 2 'load the house map

P.dir = down

P.Pack.Item(1).Act = True
P.Pack.Item(1).Value = 50
P.Pack.Item(1).Caption = "Candy"
P.Pack.Item(2).Act = True
P.Pack.Item(2).Value = 50
P.Pack.Item(2).Caption = "Candy"
P.Pack.Item(3).Act = True
P.Pack.Item(3).Value = 100
P.Pack.Item(3).Caption = "Magic Dust"

P.SpellCut = False

P.B.Level = 1
P.B.Attack = Int(Rnd * 10) + 1
P.B.Defense = Int(Rnd * 7) + 1
P.B.Health = 50
P.B.MaxHealth = 50
P.B.Exp = 0
P.B.TExp = 10

P.Pack.money = 100
P.Pack.mDust = 10
P.Walking = False
P.DestX = P.X
P.DestY = P.Y
End Function

Function MovePlayer()
If P.Walking = True And P.WalkC > 3 Then
    P.WalkC = 0

    Select Case P.X
        Case Is < P.DestX
          P.X = P.X + 6
        Case Is > P.DestX
            P.X = P.X - 6
        Case Is = P.DestX
            If isValue(P.dir, nDir.Left, nDir.Right) And P.Walking = True Then P.Walking = False: WalkOn GetTile
    End Select

    Select Case P.Y
        Case Is < P.DestY
            P.Y = P.Y + 6
        Case Is > P.DestY
            P.Y = P.Y - 6
        Case Is = P.DestY
            If isValue(P.dir, nDir.up, nDir.down) And P.Walking = True Then P.Walking = False: WalkOn GetTile
    End Select

Else

    P.WalkC = P.WalkC + 1

End If

If P.StepC > 5 Then

    P.StepC = 0
    P.Step = P.Step + 1: If P.Step > 1 Then P.Step = 0

Else

    P.StepC = P.StepC + 1

End If

If E.IsPressed(vbKeyUp) Then

    If P.Walking = False Then P.DestY = P.Y - 18: P.Walking = True: P.dir = up
    If T(GetTile).Walkable = False Then P.DestY = P.Y
    If isTileTaken(GetTile) = True Then P.DestY = P.Y
    
End If

If E.IsPressed(vbKeyDown) Then

    If P.Walking = False Then P.DestY = P.Y + 18: P.Walking = True: P.dir = down
    If T(GetTile).Walkable = False Then P.DestY = P.Y
    If isTileTaken(GetTile) = True Then P.DestY = P.Y
    
End If

If E.IsPressed(vbKeyLeft) Then

    If P.Walking = False Then P.DestX = P.X - 18: P.Walking = True: P.dir = Left
    If T(GetTile).Walkable = False Then P.DestX = P.X
    If isTileTaken(GetTile) = True Then P.DestX = P.X
    
End If

If E.IsPressed(vbKeyRight) Then

    If P.Walking = False Then P.DestX = P.X + 18: P.Walking = True: P.dir = Right
    If T(GetTile).Walkable = False Then P.DestX = P.X
    If isTileTaken(GetTile) = True Then P.DestX = P.X
    
End If

If E.IsPressed(vbKeySpace) Then TalkTo

If E.IsPressed(vbKeyZ) Then DoSpell
End Function

Function isTileTaken(I As Long)
Dim n
isTileTaken = False

        For n = 1 To 20
        
            If o(n).Act = True And o(n).DestX = T(I).X And o(n).DestY = T(I).Y Then isTileTaken = True: Exit Function
        
        Next n

End Function

Function LoadMarket()
frmGame.board.Visible = False
frmGame.menu.Visible = False
frmGame.battle.Visible = False
frmGame.main.Visible = False
frmGame.market.Visible = True

frmGame.lblMoney.Caption = P.Pack.money & "$"

frmGame.lstStuff.Clear
frmGame.lstICost.Clear

For I = 1 To 20

    If P.Pack.Item(I).Act = True Then
    
        frmGame.lstStuff.AddItem P.Pack.Item(I).Caption
        frmGame.lstICost.AddItem P.Pack.Item(I).Value
        
    End If

Next I

Dim n As String, n2

Open App.path + "\other data\" & M.MarketPath For Input As #1

Input #1, n

frmGame.market.Picture = LoadPicture(App.path + n)

frmGame.lstBuy.Clear
frmGame.lstMCost.Clear

Do Until EOF(1)

    Input #1, n, n2

    frmGame.lstBuy.AddItem n
    frmGame.lstMCost.AddItem n2

    DoEvents
    
Loop
End Function

Function AddPackItem(Caption As String, Cost As Long) As Boolean
Dim I As Long

AddPackItem = True

For I = 1 To 20

    If P.Pack.Item(I).Act = False Then

        P.Pack.Item(I).Act = True
        P.Pack.Item(I).Caption = Caption
        P.Pack.Item(I).Value = Cost
        GoTo Done

    End If

Next I

AddPackItem = False
MsgBox ("Not enough room in your pack for the item!"), , "MagicQuest"

Done:
End Function

Function MoveObjs()
Dim I As Long

DoAI

For I = 1 To 20

    If o(I).Walking = True And o(I).WalkC > 3 Then

        o(I).WalkC = 0

        Select Case o(I).X
            Case Is < o(I).DestX
                o(I).X = o(I).X + 6
            Case Is > o(I).DestX
                o(I).X = o(I).X - 6
            Case Is = o(I).DestX
                If isValue(o(I).dir, nDir.Left, nDir.Right) And o(I).Walking = True Then o(I).Walking = False: WalkOn GetTile
        End Select

        Select Case o(I).Y
            Case Is < o(I).DestY
                o(I).Y = o(I).Y + 6
            Case Is > o(I).DestY
                o(I).Y = o(I).Y - 6
            Case Is = o(I).DestY
        If isValue(o(I).dir, nDir.up, nDir.down) And o(I).Walking = True Then o(I).Walking = False: WalkOn GetTile
        End Select

    Else

        o(I).WalkC = o(I).WalkC + 1

    End If

    If o(I).Person = True Then

        If o(I).StepC > 5 Then

            o(I).StepC = 0
            o(I).Step = o(I).Step + 1: If o(I).Step > 1 Then o(I).Step = 0

        Else
            
            o(I).StepC = o(I).StepC + 1
            
        End If

    End If

Next I
End Function

Function GetObj(Ind) As Long
Dim I As Long

For I = 1 To 20

    If o(I).Act = True And o(I).DestX = T(Ind).X And o(I).DestY = T(Ind).Y Then GetObj = I: Exit For

Next I
End Function

Function isPlayerOnTile(TileI As Long) As Boolean
isPlayerOnTile = False
If P.DestX = T(TileI).X And P.DestY = T(TileI).Y Then isPlayerOnTile = True
End Function

Function DoSpell()
Dim TileI As Long

Select Case P.dir

    Case nDir.up: TileI = GetTile - M.Mapheight
    Case nDir.down: TileI = GetTile + M.Mapheight
    Case nDir.Left: TileI = GetTile - 1
    Case nDir.Right: TileI = GetTile + 1

End Select

If T(TileI).Type = 3 And P.SpellCut = True And P.Pack.mDust > 0 Then

    P.Pack.mDust = P.Pack.mDust - 5: If P.Pack.mDust < 0 Then P.Pack.mDust = 0
    T(TileI).Type = 1
    T(TileI).Walkable = True
    
End If
End Function

Function DoAI()
Dim I As Long

For I = 1 To 20

    If o(I).Act = True Then

        o(I).AI = o(I).AI + 1

        If o(I).AI > 200 And o(I).AI_Type = 1 Then
        
        Select Case o(I).dir
            Case nDir.up: o(I).dir = nDir.Right
            Case nDir.Right: o(I).dir = nDir.down
            Case nDir.down: o(I).dir = nDir.Left
            Case nDir.Left: o(I).dir = nDir.up
        End Select
        
        o(I).AI = 0
    
    Else
    
        o(I).AI = o(I).AI + 1

    End If

    If o(I).AI > 70 And o(I).AI_Type = 3 Then
    
        o(I).AI = 0
            
        Dim n As Long
        n = Int(Rnd * 8)

        Select Case n
            Case 4
                If o(I).Walking = False Then o(I).DestY = o(I).Y - 18: o(I).Walking = True: o(I).dir = nDir.up
                If T(GetObjTile(I)).Walkable = False Then o(I).DestY = o(I).Y: o(I).Walking = False
                If isTileTaken(GetObjTile(I)) = True Then o(I).DestY = o(I).Y: o(I).Walking = False
                If isPlayerOnTile(GetObjTile(I)) = True Then o(I).DestY = o(I).Y: o(I).Walking = False
            Case 5
                If o(I).Walking = False Then o(I).DestY = o(I).Y + 18: o(I).Walking = True: o(I).dir = nDir.down
                If T(GetObjTile(I)).Walkable = False Then o(I).DestY = o(I).Y: o(I).Walking = False
                If isTileTaken(GetObjTile(I)) = True Then o(I).DestY = o(I).Y: o(I).Walking = False
                If isPlayerOnTile(GetObjTile(I)) = True Then o(I).DestY = o(I).Y: o(I).Walking = False
            Case 6
                If o(I).Walking = False Then o(I).DestX = o(I).X - 18: o(I).Walking = True: o(I).dir = nDir.Left
                If T(GetObjTile(I)).Walkable = False Then o(I).DestX = o(I).X: o(I).Walking = False
                If isTileTaken(GetObjTile(I)) = True Then o(I).DestX = o(I).X: o(I).Walking = False
                If isPlayerOnTile(GetObjTile(I)) = True Then o(I).DestX = o(I).X: o(I).Walking = False
            Case 7
                If o(I).Walking = False Then o(I).DestX = o(I).X + 18: o(I).Walking = True: o(I).dir = nDir.Right
                If T(GetObjTile(I)).Walkable = False Then o(I).DestX = o(I).X: o(I).Walking = False
                If isTileTaken(GetObjTile(I)) = True Then o(I).DestX = o(I).X: o(I).Walking = False
                If isPlayerOnTile(GetObjTile(I)) = True Then o(I).DestX = o(I).X: o(I).Walking = False
            Case Else
                o(I).dir = n
            End Select

        Else
        
            o(I).AI = o(I).AI + 1
            
        End If
        
    End If

Next I
End Function

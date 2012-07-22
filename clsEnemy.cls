VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Note:
'   Property Let assigns a value (passed as a parameter) to a property
'   Property Get retrieves the value of a property
'If a property doesn't need to be readable/writeable from outside the class, the Get/Let can be ommited for the property

'Properties store information about the enemy

'pIndex stores the index of the object (since I can't find a way for an object to find out its own index otherwise)
Private pIndex As Byte

Private pHP As Long
Private pMP As Integer
Private pLvl As Byte
Private pStrength As Byte
Private pDefense As Byte
Private pSpeed As Byte
Private pHax As Byte
Private pLuck As Byte

Private pTurnTimer As Integer
'pInternetsGained is the number of Internets the player gets for beating the enemy
Private pInternetsGained As Long
'pEXPGained is amount of EXP the player gets for beating the enemy
Private pEXPGained As Long

'Doese damage to the player
Public Sub AttackPlayer()
    
    'Declare variables
    '   damage is the damage done to the enemy if the attack hits
    '   target determines which Anon is being attacked
    '   special determines if the enemy wil use their special attack or not
    '   StatUsed stores Strength if the enemy is using an attack or Hax if its using a special ability
    '   AbilityPower is a value from 0 to 1 that is a multiplier for the base damage calculations
    Dim damage As Integer, target As Byte, special As Boolean, StatUsed As Byte, AbilityPower As Single
    
    'Since the player is taking their turn, the turn order is shifted
    Call frmBattle.ShiftTurn
    
    'Determines which Anon will be attacked (loops until a living Anon is found)
    Do
        target = CByte(Int(Rnd * 3))
    Loop While frmMenu.lblKO(target).Visible
    
    'Determines if the enemy will use their special attack
    special = CBool(Int(Rnd * 2))
    
    'Determines whether Strength or Hax will be used in the damage formula and sets the power of the ability
    If special Then
        StatUsed = pHax
        'If the enemy has enough MP to use the ability
        If pMP >= 20 Then
            AbilityPower = 1
            'MP is reduced
            pMP = pMP - 20
        Else
            'If the enemy doesn't have the necessary MP, the ability power is set to 0 which renders the attack useless
            AbilityPower = 0
        End If
    Else
        StatUsed = pStrength
        AbilityPower = 0.5
    End If
    
    
    'If the attack will successfully hit the enemy
    If HitSuccess() Then
       
        'Damage is calculated
        '   The division by 1.5 is because people were complaining that their characters were dying to suddenly
        damage = ((DamageCalc(StatUsed, pLvl, CByte(frmMenu.lblDefense(target).Caption), AbilityPower)) \ 1.5)
        
        'Does damage to the Anon
        frmMenu.lblCurrHP(target).Caption = CStr(CInt(frmMenu.lblCurrHP(target).Caption) - damage)
        
        'Checks if the attack killed the Anon
        If CInt(frmMenu.lblCurrHP(target).Caption) <= 0 Then
            'Declares the Anon as dead
            frmMenu.lblCurrHP(target).Caption = 0
            frmMenu.lblKO(target).Visible = True
            'Resets the turn counter
            frmBattle.lblAnon(target).Tag = "0"
        End If
    End If
End Sub

'Increases the ememy's turn timer
Public Sub IncreaseTime()
    
    'Declare variables
    '   increase is the value that an enemy's turn timer goes up by every 0.1 seconds
    Dim increase As Integer
    
    'The turn counter won't increase for dead enemies or for enemies that have reached their turn and are waiting in the queue
    If pHP > 0 And pTurnTimer < 30000 Then
        'Gets the turn timer increase
        increase = TurnIncrease(pSpeed)
            
        'If the enemy reaches the max for the turn timer
        If pTurnTimer + increase >= 30000 Then
            pTurnTimer = 30000
            'They are added to the turn queue
            Call frmBattle.QueueTurn(CByte(pIndex + 3))
        'Otherwise their turn counter is simply increased
        Else
            pTurnTimer = pTurnTimer + increase
        End If
    End If
End Sub

'Resets the enemy's turn timer
Public Sub ResetTime()
    pTurnTimer = 0
End Sub


Public Property Let Index(ByVal newData As Byte)
    pIndex = newData
End Property

Public Property Let HP(ByVal newData As Long)
    
    'The change cannot result in a negative HP
    If newData <= 0 Then
        
        newData = 0
        
        'Since the enemy has died, their picturebox/image is made invisible
        If frmBattle.BossBattle Then
            frmBattle.imgBoss(frmBattle.Area - 2).Visible = False
        Else
            frmBattle.picTroll(pIndex).Visible = False
        End If
    End If
    
    pHP = newData
End Property

Public Property Get HP() As Long
    HP = pHP
End Property

Public Property Get Lvl() As Byte
    Lvl = pLvl
End Property

Public Property Get Defense() As Byte
    Defense = pDefense
End Property

Public Property Get InternetsGained() As Long
    InternetsGained = pInternetsGained
End Property

Public Property Get EXPGained() As Long
    EXPGained = pEXPGained
End Property

'When the object is created, its values are set
Private Sub Class_Initialize()
    
    'If it is a boss being created
    If frmBattle.BossBattle Then
        
        'Sets the Boss' stats based on which area the player is moving to
        If frmBattle.Area = 2 Then
            
            pHP = 1
            pMP = 1
            pLvl = 1
            pStrength = 1
            pDefense = 1
            pSpeed = 1
            pHax = 1
            pLuck = 1
        
        ElseIf frmBattle.Area = 3 Then
            
            pHP = 6000
            pMP = 30
            pLvl = 20
            pStrength = 16
            pDefense = 2
            pSpeed = 5
            pHax = 2
            pLuck = 2
        
        ElseIf frmBattle.Area = 4 Then
            
            pHP = 10000
            pMP = 40
            pLvl = 30
            pStrength = 25
            pDefense = 25
            pSpeed = 12
            pHax = 20
            pLuck = 15
        
        ElseIf frmBattle.Area = 5 Then
            
            pHP = 12000
            pMP = 100
            pLvl = 50
            pStrength = 27
            pDefense = 100
            pSpeed = 20
            pHax = 20
            pLuck = 20
        
        ElseIf frmBattle.Area = 6 Then
            
            pHP = 18000
            pMP = 130
            pLvl = 60
            pStrength = 40
            pDefense = 5
            pSpeed = 19
            pHax = 1
            pLuck = 15
        
        ElseIf frmBattle.Area = 7 Then
            
            pHP = 23555
            pMP = 55
            pLvl = 70
            pStrength = 55
            pDefense = 55
            pSpeed = 55
            pHax = 5
            pLuck = 5
        
        ElseIf frmBattle.Area = 8 Then
            
            pHP = 32000
            pMP = 500
            pLvl = 80
            pStrength = 45
            pDefense = 30
            pSpeed = 30
            pHax = 70
            pLuck = 20
        
        ElseIf frmBattle.Area = 9 Then
            
            pHP = 70000
            pMP = 590
            pLvl = 90
            pStrength = 90
            pDefense = 55
            pSpeed = 50
            pHax = 60
            pLuck = 28
        
        Else
            
            pHP = 5000000
            pMP = 999
            pLvl = 99
            pStrength = 255
            pDefense = 240
            pSpeed = 150
            pHax = 255
            pLuck = 50
            
        End If
    'If a troll is being created
    Else
        
        'Sets stats based on the area the player is in
        pHP = CLng((frmBattle.Area * 25) ^ 1.2 * RandomMultiplier)
        pMP = Int(((frmBattle.Area * 2) + 5) ^ 1.2 * RandomMultiplier)
        pLvl = CByte(10 * Rnd + ((frmBattle.Area - 1) * 10) + 1)
        pStrength = CByte(((frmBattle.Area + 5) ^ 1.2) * RandomMultiplier)
        pDefense = CByte(((frmBattle.Area + 5) ^ 1.1) * RandomMultiplier)
        pSpeed = CByte(((frmBattle.Area + 5) ^ 1.2) * RandomMultiplier)
        pHax = CByte(((frmBattle.Area + 3) ^ 1.1) * RandomMultiplier)
        pLuck = CByte(((frmBattle.Area + 15) ^ 1.3) * RandomMultiplier)
    
    End If
    
    
    'Sets the number of Internets the player gets for beating the enemy
    pInternetsGained = CLng((150 * ((frmBattle.Area + 1) ^ 1.5)) ^ (CInt(frmBattle.BossBattle) + 1))
    'Sets the amount of EXP the player gets for beating the enemy
    pEXPGained = CLng((10 * ((frmBattle.Area + 1) ^ 1.5)) ^ (CInt(frmBattle.BossBattle) + 1))
    
    'Sets the turn counter to the enemy's starting bonus
    pTurnTimer = StartingTurnCounterBonus(pLuck)
    
End Sub

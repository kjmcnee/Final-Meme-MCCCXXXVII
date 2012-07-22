Attribute VB_Name = "ModuleBattleSystem"
Option Explicit

'Changes a character's stats based on which Item is used
'   ItemID identifies the item being used (the values are the same as the index values for the item controls in the Menu form (for example: Potion = 0, Phoenix Down = 5))
'   Target identifies which character the item is being used on
Sub UseItem(ItemID As Byte, target As Byte)
    'If the player has 1 or more of the item, it can be used
    If CInt(frmMenu.lblItemAmount(ItemID).Caption) > 0 Then
        'The number of that item the user has left is reduced, since one will have been used
        frmMenu.lblItemAmount(ItemID).Caption = CStr(CInt(frmMenu.lblItemAmount(ItemID).Caption) - 1)
        
        'If the item is a Phoenix Down
        If ItemID = 5 Then
            'The character is revived
            frmMenu.lblKO(target).Visible = False
            'Character is given a fraction of their HP back (if they have actually been been revived)
            If frmMenu.lblCurrHP(target).Caption = "0" Then
                frmMenu.lblCurrHP(target).Caption = CStr(Int(CInt(frmMenu.lblMaxHP(target).Caption) / 4))
            End If
        'If the item is one that changes HP or MP
        ElseIf ItemID <= 4 Or ItemID = 6 Then
            
            'The character must be alive for the item to have an effect (this prevents a character from being dead while having HP that is greater than 0)
            If Not frmMenu.lblKO(target).Visible Then
                
                'Adjusts current HP/MP based on which item has been selected
                'Potion:
                'HP +200
                If ItemID = 0 Then
                    frmMenu.lblCurrHP(target).Caption = CStr(CInt(frmMenu.lblCurrHP(target).Caption) + 200)
                'Hi-Potion:
                'HP +1000
                ElseIf ItemID = 1 Then
                    frmMenu.lblCurrHP(target).Caption = CStr(CInt(frmMenu.lblCurrHP(target).Caption) + 1000)
                'X-Potion
                'HP fully restored
                ElseIf ItemID = 2 Then
                    frmMenu.lblCurrHP(target).Caption = frmMenu.lblMaxHP(target).Caption
                'Ether
                'MP +100
                ElseIf ItemID = 3 Then
                    frmMenu.lblCurrMP(target).Caption = CStr(CInt(frmMenu.lblCurrMP(target).Caption) + 100)
                'Turbo Ether
                'MP +500
                ElseIf ItemID = 4 Then
                    frmMenu.lblCurrMP(target).Caption = CStr(CInt(frmMenu.lblCurrMP(target).Caption) + 500)
                'Elixir
                'HP fully restored
                'MP fully restored
                ElseIf ItemID = 6 Then
                    frmMenu.lblCurrHP(target).Caption = frmMenu.lblMaxHP(target).Caption
                    frmMenu.lblCurrMP(target).Caption = frmMenu.lblMaxMP(target).Caption
                End If
               
                'Makes sure that the increase doesn't cause the stat to exceed the cap
                Call EnforceStatCaps
            End If
        'If the item is one that increases stats such as Strength or Speed
        ElseIf ItemID >= 7 Then
            
            'Increases stats based on which item has been selected
            'HP+
            'Increase by 200
            If ItemID = 7 Then
                frmMenu.lblMaxHP(target).Caption = CStr(CInt(frmMenu.lblMaxHP(target).Caption) + 200)
            'MP+
            'Increase by 20
            ElseIf ItemID = 8 Then
                frmMenu.lblMaxMP(target).Caption = CStr(CInt(frmMenu.lblMaxMP(target).Caption) + 20)
            'Strength+
            ElseIf ItemID = 9 Then
                frmMenu.lblStrength(target).Caption = CStr(CInt(frmMenu.lblStrength(target).Caption) + 1)
            'Defense+
            ElseIf ItemID = 10 Then
                frmMenu.lblDefense(target).Caption = CStr(CInt(frmMenu.lblDefense(target).Caption) + 1)
            'Speed+
            ElseIf ItemID = 11 Then
                frmMenu.lblSpeed(target).Caption = CStr(CInt(frmMenu.lblSpeed(target).Caption) + 1)
            'Hax+
            ElseIf ItemID = 12 Then
                frmMenu.lblHax(target).Caption = CStr(CInt(frmMenu.lblHax(target).Caption) + 1)
            'Luck+
            ElseIf ItemID = 13 Then
                frmMenu.lblLuck(target).Caption = CStr(CInt(frmMenu.lblLuck(target).Caption) + 1)
            End If
             
            'Makes sure that the increase doesn't cause the stat to exceed the cap
            Call EnforceStatCaps
        End If
    End If
End Sub

'Prevents stat increases from exceeding the stats' max values
Sub EnforceStatCaps()
    
    'Declare variables
    '   count is a loop counter used to check stats for each Anon and each Item
    Dim count As Integer
    
    'If any stat is equal to or greater than its max, its value is set to its max
    
    'Checks stats of each anon
    For count = frmMenu.lblAnon.LBound To frmMenu.lblAnon.UBound
    
        If CInt(frmMenu.lblMaxHP(count).Caption) >= 9999 Then
            frmMenu.lblMaxHP(count).Caption = "9999"
        End If
        If CInt(frmMenu.lblMaxMP(count).Caption) >= 999 Then
            frmMenu.lblMaxMP(count).Caption = "999"
        End If
        
        'A character's Current HP/MP can't be greater than the character's Max HP/MP
        If CInt(frmMenu.lblCurrHP(count).Caption) >= CInt(frmMenu.lblMaxHP(count).Caption) Then
            frmMenu.lblCurrHP(count).Caption = frmMenu.lblMaxHP(count).Caption
        End If
        If CInt(frmMenu.lblCurrMP(count).Caption) >= CInt(frmMenu.lblMaxMP(count).Caption) Then
            frmMenu.lblCurrMP(count).Caption = frmMenu.lblMaxMP(count).Caption
        End If
        
        If CByte(frmMenu.lblLvl(count).Caption) >= 99 Then
            frmMenu.lblLvl(count).Caption = "99"
        End If
        
        If CByte(frmMenu.lblStrength(count).Caption) >= 255 Then
            frmMenu.lblStrength(count).Caption = "255"
        End If
        If CByte(frmMenu.lblDefense(count).Caption) >= 255 Then
            frmMenu.lblDefense(count).Caption = "255"
        End If
        If CByte(frmMenu.lblSpeed(count).Caption) >= 255 Then
            frmMenu.lblSpeed(count).Caption = "255"
        End If
        If CByte(frmMenu.lblHax(count).Caption) >= 255 Then
            frmMenu.lblHax(count).Caption = "255"
        End If
        If CByte(frmMenu.lblLuck(count).Caption) >= 255 Then
            frmMenu.lblLuck(count).Caption = "255"
        End If
    Next count
    
    'Checks the number of items in the player's inventory
    For count = frmMenu.lblItemAmount.LBound To frmMenu.lblItemAmount.UBound
        If CByte(frmMenu.lblItemAmount(count).Caption) >= 99 Then
            frmMenu.lblItemAmount(count).Caption = "99"
        End If
    Next count
    
    'Caps the player's currency
    If CLng(frmMenu.lblInternets.Caption) >= 999999999 Then
        frmMenu.lblInternets.Caption = "999999999"
    End If
    
End Sub

'Returns a value from 0.85 to 1.15
'Used in damage formula, troll stat generation, etc. to introduce some randomness in the result
Function RandomMultiplier() As Single
    RandomMultiplier = 0.3 * Rnd + 0.85
End Function

'Gives EXP to each Anon
'   EXPIncrease is th amount of EXP earned
Sub GiveEXP(EXPIncrease As Long)
    
    'Declare variables
    '   count is the loop counter for giving EXP to each Anon
    Dim count As Integer
    
    'For each Anon
    For count = frmMenu.lblExp().LBound To frmMenu.lblExp().UBound
        'Does not give EXP when the Anon is level 99 since there are no more levels after that
        If CInt(frmMenu.lblLvl(count).Caption) < 99 Then
            'If giving the EXP causes the required EXP to the next level to reach 0, the Anon is leveled up
            If CLng(frmMenu.lblExp(count).Caption) - EXPIncrease <= 0 Then
                Call LevelUp(CByte(count))
            'Otherwise the EXP is simply subtracted from the amount required
            Else
                frmMenu.lblExp(count).Caption = CStr(CLng(frmMenu.lblExp(count).Caption) - EXPIncrease)
            End If
        End If
    Next
    
End Sub

'When an Anon earns enough EXP to cause them to level up, their stats are increased
'   PlayerID identifies which Anon is leveling up
Sub LevelUp(PlayerID As Byte)

    'Increments the Anon's level
    frmMenu.lblLvl(PlayerID).Caption = CStr(CInt(frmMenu.lblLvl(PlayerID).Caption) + 1)
    
    'Sets the new EXP required to make it to the next level
    If CInt(frmMenu.lblLvl(PlayerID).Caption) < 99 Then
        frmMenu.lblExp(PlayerID).Caption = CStr((CInt(frmMenu.lblLvl(PlayerID).Caption) + 5) ^ 3)
    'This does not happen at level 99 since there are no more levels after that
    Else
        'Instead, the EXP label for that Anon is set to 0
        frmMenu.lblExp(PlayerID).Caption = "0"
    End If
    
    'Increases stats
    frmMenu.lblMaxHP(PlayerID).Caption = CStr(Int(CInt(frmMenu.lblMaxHP(PlayerID).Caption) + Sqr(CInt(frmMenu.lblLvl(PlayerID).Caption)) * 10))
    frmMenu.lblMaxMP(PlayerID).Caption = CStr(Int(CInt(frmMenu.lblMaxMP(PlayerID).Caption) + Sqr(CInt(frmMenu.lblLvl(PlayerID).Caption)) * 1.8))
    frmMenu.lblStrength(PlayerID).Caption = CStr(CInt(frmMenu.lblStrength(PlayerID).Caption) + RandomStatIncrease)
    frmMenu.lblDefense(PlayerID).Caption = CStr(CInt(frmMenu.lblDefense(PlayerID).Caption) + RandomStatIncrease)
    frmMenu.lblSpeed(PlayerID).Caption = CStr(CInt(frmMenu.lblSpeed(PlayerID).Caption) + RandomStatIncrease)
    frmMenu.lblHax(PlayerID).Caption = CStr(CInt(frmMenu.lblHax(PlayerID).Caption) + RandomStatIncrease)
    frmMenu.lblLuck(PlayerID).Caption = CStr(CInt(frmMenu.lblLuck(PlayerID).Caption) + RandomStatIncrease)

End Sub

'Returns a number from 0-3 which is added to an Anon's stats
Function RandomStatIncrease() As Byte
    
    'Declare variables
    '   increase is the random number that determines the increase (then it stores the increase itself)
    Dim increase As Byte
    
    'Generates a value from 0-7
    increase = CByte(Rnd * 8)
    
    'Determines the stat increase from the generated value (these multiple steps are needed because different stat increases have different chances of occuring)
    If increase <= 2 Then
        increase = 1
    ElseIf increase >= 3 And increase <= 5 Then
        increase = 2
    ElseIf increase = 6 Then
        increase = 3
    ElseIf increase = 7 Then
        increase = 0
    End If
    
    RandomStatIncrease = increase
    
End Function

'Returns the increase to the turn timer of a character
'   Speed is the character's speed
Function TurnIncrease(Speed As Byte) As Integer
    TurnIncrease = Int((((Speed ^ 2 + 50) * 128) / Speed) / 2)
End Function

'Determines if an attack will hit the target (returns True) or miss (returns False)
'   I used to have it so that the speed of the attacker and defender determined the chance of hitting, but there were so many misses that I went with a constant 80% chance of hitting
Function HitSuccess() As Boolean
    
    'Generates a random number between 0 and 4 so the attack has an 80% chance of hitting (0 = miss and 1,2,3,4 = hit)
    HitSuccess = CBool((Int(Rnd * 5) > 0))
    
End Function

'Determines the amount of damage done by an attack
'   The passed values include the stats of the character doing the attack and the target of the attack that affect the damage done
'   AttackerPower will either be the Strength or Hax stat depending on which ability is being used
'   AbilityPower is a value from 0 to 1 that is a multiplier for the base damage calculations (for example: a normal attack has a value of 0.5, and the B& Hammer has a value of 0.9, therefore the B& Hammer will deal more damage)
Function DamageCalc(AttackerPower As Byte, AttackerLvl As Byte, DefenderDefense As Byte, AbilityPower As Single) As Integer
    
    'Declare variables
    '   damage is the calculated damage to be done to the defender
    Dim damage As Integer
    
    'Calculates damage
    damage = Int(RandomMultiplier * (AbilityPower * (512 - DefenderDefense) * (6 * (AttackerPower + AttackerLvl)) / 50))
    
    'The damage is capped at 9999
    If damage >= 9999 Then
        damage = 9999
    End If
    
    DamageCalc = damage
End Function

'A bonus given to a character's turn timer at the beginning of battle
'   The character's luck determines the bonus
Function StartingTurnCounterBonus(Luck As Byte) As Integer
    StartingTurnCounterBonus = CInt(Luck * 50 * RandomMultiplier)
End Function

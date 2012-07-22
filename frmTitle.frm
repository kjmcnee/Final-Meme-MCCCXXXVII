VERSION 5.00
Begin VB.Form frmTitle 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Final Meme MCCCXXXVII"
   ClientHeight    =   6645
   ClientLeft      =   4740
   ClientTop       =   3810
   ClientWidth     =   9030
   Icon            =   "frmTitle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCursor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2520
      Picture         =   "frmTitle.frx":030A
      ScaleHeight     =   375
      ScaleWidth      =   675
      TabIndex        =   5
      Tag             =   "0"
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Explaination"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6615
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   8775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblProtip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Protip: Use the Arrow Keys to navigate the menu and the Spacebar or Enter to make a selection."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   5880
      Width           =   6855
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblOption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblOption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Load Game"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblOption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Final Meme MCCCXXXVII"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   1380
      TabIndex        =   0
      Top             =   600
      Width           =   6090
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Note: the cursor's tag stores the index of the control it is currently next to in a control array
'This is the same for all forms

'This is Mr. Hudson's function:
'This function receives a filename and based on the folder
'   where the program is being run from it returns the
'   proper formatted path and filename combined.
Function GetPath(FileName As String) As String
    'Declaring variables:
    'Path is a temporary variable used to hold the
    '   application's path.
    Dim Path As String
    'Get the "app"lication's path.
    Path = App.Path
    'Check that there is a slash on the path's
    '   folder name -- if not then add it.
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    'Return the correct path and tack on the file's name.
    GetPath = Path & FileName
End Function
'The rest is my own code

'Moves the cursor through the label array
'   Direction is the direction the cursor moves
'   1 will increase the tag's value (moves down corresponding with the control array's index)
'   -1 will decrease the tag's value (moves up corresponding with the control array's index)
Sub MoveCursor(Direction As Integer)
    
    'Cursor tag is adjusted in the desired direction
    picCursor.Tag = CStr(CInt(picCursor.Tag) + Direction)
    
    'If the tag exceeds the bounds of the array it is cycled
    If CInt(picCursor.Tag) < lblOption.LBound Then
        picCursor.Tag = CStr(lblOption.ubound)
    ElseIf CInt(picCursor.Tag) > lblOption.ubound Then
        picCursor.Tag = CStr(lblOption.LBound)
    End If
    
    'Changes the position of the cursor so that it is next to the new control
    picCursor.Top = lblOption(CInt(picCursor.Tag)).Top + (lblOption(0).Height / 4)
    picCursor.Left = lblOption(CInt(picCursor.Tag)).Left - picCursor.Width
    
End Sub
'Takes the relevant action based on which label the cursor is next to
Sub UserSelection()
    'If the cursor is next to the New Game option, the game instructions are shown
    If CInt(picCursor.Tag) = 0 Then
        lblInstructions.Visible = True
        picCursor.Visible = False
    'If the cursor is next to the Load Game option, the game is loaded from the save file
    ElseIf CInt(picCursor.Tag) = 1 Then
        Call LoadGame
    'If the cursor is next to the Quit option, the program closes
    ElseIf CInt(picCursor.Tag) = 2 Then
        End
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If the user previously selected "New Game" and the instructions are being shown, another key press will actually start the new game
    If lblInstructions.Visible Then
        Call NewGame
    Else
        'If the player presses up, the cursor moves up
        If KeyCode = vbKeyUp Then
            Call MoveCursor(-1)
        'If the player presses down, the cursor moves down
        ElseIf KeyCode = vbKeyDown Then
            Call MoveCursor(1)
        'If the player presses enter, the relevant action is taken
        ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
            Call UserSelection
        End If
    End If
End Sub

Private Sub Form_Load()
    
    Randomize
    
    'Sets the cursor next to the first label
    picCursor.Tag = "0"
    picCursor.Top = lblOption(0).Top + (lblOption(0).Height / 4)
    picCursor.Left = lblOption(0).Left - picCursor.Width
    'Writes the instructions in the Instructions label's caption
    lblInstructions.Caption = vbCrLf & "Internet memes have infected this computer!" & vbCrLf & vbCrLf & "We need your help, Anonymous! The Final Boss of the Internet has gathered OVER 9000 internet memes. We need you to kill all of the memes and then the Final Boss of the Internet in order to restore peace to the kernel." & vbCrLf & vbCrLf & "Don't worry, we are fully aware that you are absolutely useless right now. Fortunately, the memes brought an infinite number of trolls with them. So, you can kill them to gain experience and level up. That way the Bosses won't kill you in one turn." & vbCrLf & vbCrLf & "Controls:" & vbCrLf & "Movement/Menu Navigation: Arrow Keys" & vbCrLf & "Select/Enter Store: Enter or Spacebar" & vbCrLf & "Back/Exit Store/Exit Menu: B or Del(Numpad)" & vbCrLf & "Open Menu: M or +(Numpad)" & vbCrLf & "Escape from Battle: Esc" & vbCrLf & vbCrLf & "tl;dr: Go kill things." & vbCrLf & vbCrLf & "Now go, anon! Press any key to continue!"
    lblInstructions.Visible = False
End Sub

'Starts a new game with starting stats
Sub NewGame()
    
    'Declare variables
    '   count is a loop counter for the initialization of items and stats
    Dim count As Integer
    
    'Loads the Menu form
    Load frmMenu
    
    'Sets the starting values for all stats
    
    'Sets the starting number of items
    For count = frmMenu.lblItemAmount.LBound To frmMenu.lblItemAmount.ubound
        'The player starts with 5 Potions
        If count = 0 Then
            frmMenu.lblItemAmount(count).Caption = "5"
        '1 Hi-Potion
        ElseIf count = 1 Then
            frmMenu.lblItemAmount(count).Caption = "1"
        '2 Phoenix Downs
        ElseIf count = 5 Then
            frmMenu.lblItemAmount(count).Caption = "2"
        '0 for every other item
        Else
            frmMenu.lblItemAmount(count).Caption = "0"
        End If
    Next count
    
    'Sets the starting amount of Internets (the in-game currency)
    frmMenu.lblInternets.Caption = "500"
    
    'Sets player starting stats
    For count = frmMenu.lblAnon.LBound To frmMenu.lblAnon.ubound
        frmMenu.lblLvl(count).Caption = "1"
        frmMenu.lblExp(count).Caption = "343"
        frmMenu.lblCurrHP(count).Caption = "500"
        frmMenu.lblMaxHP(count).Caption = "500"
        frmMenu.lblCurrMP(count).Caption = "30"
        frmMenu.lblMaxMP(count).Caption = "30"
        frmMenu.lblStrength(count).Caption = "17"
        frmMenu.lblDefense(count).Caption = "15"
        frmMenu.lblSpeed(count).Caption = "13"
        frmMenu.lblHax(count).Caption = "8"
        frmMenu.lblLuck(count).Caption = "16"
        frmMenu.lblKO(count).Visible = False
    Next count
    
    'Loads the Field form
    Load frmField
    
    'Sets the area that the player can go to without triggering a boss battle
    frmField.PlayerProgress = 1
    
    'Sets up the field form and shows it
    Call SetupPlayer
    
End Sub

'Continues with previous game with the stats loaded from a save file
Sub LoadGame()
    
    'Declare Variables
    '   SaveFileLoc is the path to the save file
    '   SaveData is used to store the data from the save file
    '   count is a loop counter used when extracting data from the save file
    Dim SaveFileLoc As String, SaveData As String, count As Integer
    
    'Sets the location of the save file
    SaveFileLoc = GetPath("Final_Meme_Save_Data.save")
    
    'Checks for the existance of a save file
    If Dir(SaveFileLoc) <> "" Then
        
        'Opens the file
        Open SaveFileLoc For Binary Access Read As #1
            'Reads all of the data from the file
            SaveData = Input(LOF(1), #1)
        Close #1
        
        'Decrypts the data
        SaveData = EncryptDecrypt(SaveData, False)
        
        'Compares the save data to how it should look
        'If there is a difference then the file has been tampered with
        If SaveData Like ":Potion*:Hi-Potion*:X-Potion*:Ether*:Turbo Ether*:Phoenix Down*:Elixir*:HP+*:MP+*:Strength+*:Defense+*:Speed+*:Hax+*:Luck+*:Anon1Lvl*:Anon1Exp*:Anon1CurrHP*:Anon1MaxHP*:Anon1CurrMP*:Anon1MaxMP*:Anon1Strength*:Anon1Defense*:Anon1Speed*:Anon1Hax*:Anon1Luck*:Anon1KO*:Anon2Lvl*:Anon2Exp*:Anon2CurrHP*:Anon2MaxHP*:Anon2CurrMP*:Anon2MaxMP*:Anon2Strength*:Anon2Defense*:Anon2Speed*:Anon2Hax*:Anon2Luck*:Anon2KO*:Anon3Lvl*:Anon3Exp*:Anon3CurrHP*:Anon3MaxHP*:Anon3CurrMP*:Anon3MaxMP*:Anon3Strength*:Anon3Defense*:Anon3Speed*:Anon3Hax*:Anon3Luck*:Anon3KO*:Internets*:PlayerProgress#:" Then
            Load frmMenu
            Load frmField
            
            'Extracts each stat from the file and stores it in the relevent label caption, variable, etc.
            'Items:
            For count = frmMenu.lblItemAmount.LBound To frmMenu.lblItemAmount.ubound
                frmMenu.lblItemAmount(count).Caption = ExtractStat(SaveData, frmMenu.lblItemName(count))
            Next count
            
            'Player stats:
            For count = frmMenu.lblAnon.LBound To frmMenu.lblAnon.ubound
                frmMenu.lblLvl(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "Lvl")
                frmMenu.lblExp(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "Exp")
                frmMenu.lblCurrHP(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "CurrHP")
                frmMenu.lblMaxHP(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "MaxHP")
                frmMenu.lblCurrMP(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "CurrMP")
                frmMenu.lblMaxMP(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "MaxMP")
                frmMenu.lblStrength(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "Strength")
                frmMenu.lblDefense(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "Defense")
                frmMenu.lblSpeed(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "Speed")
                frmMenu.lblHax(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "Hax")
                frmMenu.lblLuck(count).Caption = ExtractStat(SaveData, "Anon" & CStr(count + 1) & "Luck")
                frmMenu.lblKO(count).Visible = CBool(ExtractStat(SaveData, "Anon" & CStr(count + 1) & "KO"))
            Next count
            
            'Internets (currency):
            frmMenu.lblInternets.Caption = ExtractStat(SaveData, "Internets")
            
            'The area that the player can go to without triggering a boss battle:
            frmField.PlayerProgress = CByte(ExtractStat(SaveData, "PlayerProgress"))
            'End of stat extraction
            
            'Sets up the field form and shows it
            Call SetupPlayer
        Else
            'Gives the user a talkin' to
            MsgBox "Even though the file was encrypted you still thought you could get away with making changes? We'll just see about that...", vbCritical, "Stop Right There, Criminal Scum!"
            'Closes the program... that'll teach 'em
            End
        End If
    Else
        'If the save file isn't found, the user is alerted
        MsgBox "The save file was not found." & vbCrLf & vbCrLf & "If you haven't saved your game before, select the New Game option. If you have, make sure that 'Final_Meme_Save_Data.save' is in the same folder as the application and try again.", vbCritical, "Save File Not Found"
    End If
End Sub

'Prepares the Field form and shows it
Sub SetupPlayer()
    
    'Makes all area frames invisible
    frmField.AllAreasInvisible
    'Makes the "Town" frame visible
    frmField.fraArea(0).Visible = True
    
    'Set player position to the middle of the town frame
    frmField.picAnon.Top = frmField.fraArea(0).Top + (frmField.fraArea(0).Height / 2) - (frmField.picAnon.Height / 2)
    frmField.picAnon.Left = frmField.fraArea(0).Left + (frmField.fraArea(0).Width / 2) - (frmField.picAnon.Width / 2)
    
    'Sets the label that identifies the current area to the starting area (the town)
    frmField.lblAreaIndicator.Caption = "Town"
    
    'Shows the Field form in order to start the game
    frmField.Show
    Unload frmTitle
    
End Sub

'Finds the stat in the save data string and returns its value
'   Data is the save data that holds the stat and its value
'   StatName is the name of the stat
Function ExtractStat(Data As String, StatName As String) As String

    'Declare variables
    '   StatPos is the position of the first character of the VALUE of the stat
    '   EndStat is the position of the last character of the value of the stat
    Dim StatPos As Integer, EndStat As Integer
    
    'Finds the position of the first character of the stat value
    StatPos = InStr(Data, StatName) + Len(StatName)
    
    'Finds the position of the last character of the stat value
    EndStat = InStr(StatPos, Data, ":")
    
    'Finds and returns the stat value
    ExtractStat = Mid(Data, StatPos, EndStat - StatPos)

End Function

VERSION 5.00
Begin VB.Form frmStore 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Store"
   ClientHeight    =   6900
   ClientLeft      =   4185
   ClientTop       =   2010
   ClientWidth     =   11340
   Icon            =   "frmStore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCursor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      Picture         =   "frmStore.frx":030A
      ScaleHeight     =   375
      ScaleWidth      =   675
      TabIndex        =   1
      Tag             =   "0"
      Top             =   2280
      Width           =   675
   End
   Begin VB.Label lblAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   7800
      TabIndex        =   28
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label lblTotalCost 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "9000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   8280
      TabIndex        =   27
      Top             =   5640
      Width           =   1020
   End
   Begin VB.Label lblCostPerUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   8880
      TabIndex        =   26
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label lblCurrStock 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   9000
      TabIndex        =   25
      Top             =   2760
      Width           =   510
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Protip: Use the number keys and the backspace to change the amount of the item you want to buy."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   7
      Left            =   5880
      TabIndex        =   24
      Top             =   4920
      Width           =   5205
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   6
      Left            =   5880
      TabIndex        =   23
      Top             =   4320
      Width           =   1755
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cost:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   8
      Left            =   5880
      TabIndex        =   22
      Top             =   5640
      Width           =   2235
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Per Unit:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   5
      Left            =   5880
      TabIndex        =   21
      Top             =   3600
      Width           =   2865
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Stock:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   4
      Left            =   5880
      TabIndex        =   20
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lblInternets 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "999999999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   8040
      TabIndex        =   19
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Internets:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   3
      Left            =   5880
      TabIndex        =   18
      Top             =   1800
      Width           =   2085
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   2
      Left            =   4200
      TabIndex        =   17
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Potion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   16
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi-Potion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   15
      Top             =   2760
      Width           =   1995
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "X-Potion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   14
      Top             =   3360
      Width           =   1995
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Ether"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   960
      TabIndex        =   13
      Top             =   3960
      Width           =   1995
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Turbo Ether"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   960
      TabIndex        =   12
      Top             =   4560
      Width           =   2595
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Phoenix Down"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   5
      Left            =   960
      TabIndex        =   11
      Top             =   5160
      Width           =   2835
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Elixir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   6
      Left            =   960
      TabIndex        =   10
      Top             =   5760
      Width           =   1995
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   4020
      TabIndex        =   9
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   960
      TabIndex        =   8
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   4020
      TabIndex        =   7
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   4020
      TabIndex        =   6
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   4020
      TabIndex        =   5
      Top             =   3960
      Width           =   1035
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2500"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   4020
      TabIndex        =   4
      Top             =   4560
      Width           =   1035
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   5
      Left            =   4020
      TabIndex        =   3
      Top             =   5160
      Width           =   1035
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   6
      Left            =   4020
      TabIndex        =   2
      Top             =   5760
      Width           =   1035
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Index           =   0
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If the player presses up, the cursor moves up
    If KeyCode = vbKeyUp Then
        Call MoveCursor(-1)
    'If the player presses down, the cursor moves down
    ElseIf KeyCode = vbKeyDown Then
        Call MoveCursor(1)
    'If the player presses enter or space
    ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        'The item is bought
        Call BuyItem
    'If the player presses B or Decimal(Numpad)
    ElseIf KeyCode = vbKeyB Or KeyCode = vbKeyDecimal Then
        'The store is closed and the player is returned to the field
        Unload Me
   'If the player presses a number key
    ElseIf (KeyCode >= vbKey0 And KeyCode <= vbKey9) Or (KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9) Then
        'Finds the value of the number based on the KeyCode
        If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
            KeyCode = KeyCode - vbKey0
        Else
            KeyCode = KeyCode - vbKeyNumpad0
        End If
        'Uses the value to change the item amount
        lblAmount.Caption = ChangeItemAmount(CByte(KeyCode))
    'If the player presses backspace
    ElseIf KeyCode = vbKeyBack Then
        'The item amount is cleared
        lblAmount.Caption = "0"
    End If
    'Updates the labels' values
    Call LabelRefresh
End Sub

Private Sub Form_Load()
    'Sets the cursor next to the first item label
    picCursor.Tag = "0"
    picCursor.Top = lblItem(0).Top + (lblItem(0).Height / 4)
    picCursor.Left = lblItem(0).Left - picCursor.Width
    
    'Sets the starting amount for the number of each item to buy to 1
    lblAmount.Caption = "1"
    
    'Updates the labels' values
    Call LabelRefresh
    
End Sub

'Updates the information in the labels to reflect the change done in the calling routine
Sub LabelRefresh()
    
    'Updates the amount of Internets the player has left
    lblInternets.Caption = frmMenu.lblInternets.Caption
    
    'Updates the amount of the currently selected item that the player has in their inventory
    lblCurrStock.Caption = frmMenu.lblItemAmount(CInt(picCursor.Tag)).Caption
    
    'Updates the cost per unit of the currently selected item
    lblCostPerUnit.Caption = lblCost(CInt(picCursor.Tag)).Caption
    
    'Calculates and shows the total cost for buying the item
    lblTotalCost.Caption = CStr(CLng(lblCostPerUnit.Caption) * CLng(lblAmount.Caption))
    
    'If the player can afford to buy the item
    If CLng(lblTotalCost.Caption) <= CLng(lblInternets.Caption) Then
        'Make the total cost label transparent (so that the red back color will be invisible and the blue form back color will be shown)
        lblTotalCost.BackStyle = 0
    'If the player can't afford it
    Else
        'Make the total cost label opaque (so that the red back color will be visible)
        lblTotalCost.BackStyle = 1
    End If
End Sub

'Moves the cursor through the label array
'   Direction is the direction the cursor moves
'   1 will increase the tag's value (moves down corresponding with the control array's index)
'   -1 will decrease the tag's value (moves up corresponding with the control array's index)
Sub MoveCursor(Direction As Integer)
    
    'Cursor tag is adjusted in the desired direction
    picCursor.Tag = CStr(CInt(picCursor.Tag) + Direction)
    
    'If the tag exceeds the bounds of the array it is cycled
    If CInt(picCursor.Tag) < lblItem.LBound Then
        picCursor.Tag = CStr(lblItem.uBound)
    ElseIf CInt(picCursor.Tag) > lblItem.uBound Then
        picCursor.Tag = CStr(lblItem.LBound)
    End If
    
    'Changes the position of the cursor so that it is next to the new control
    picCursor.Top = lblItem(CInt(picCursor.Tag)).Top + (lblItem(0).Height / 4)
    picCursor.Left = lblItem(CInt(picCursor.Tag)).Left - picCursor.Width
    
End Sub

'Adds the item to the player's inventory and subtracts the cost
Sub BuyItem()
    'If the total cost label is transparent (the user can afford the item)
    If lblTotalCost.BackStyle = 0 Then
        
        'The payment is made
        frmMenu.lblInternets.Caption = CStr(CLng(frmMenu.lblInternets.Caption) - CLng(lblTotalCost.Caption))
        
        'The item is added to the player's inventory
        frmMenu.lblItemAmount(CInt(picCursor.Tag)).Caption = CStr(CByte(frmMenu.lblItemAmount(CInt(picCursor.Tag)).Caption) + CByte(lblAmount.Caption))
        
        'Prevents the number of items the player has from exceeding the limit
        Call EnforceStatCaps
        
    End If
End Sub

'Changes the amount of each item that the player wants to buy
'   Number is the number that is being concatenated
Function ChangeItemAmount(Number As Byte) As String
    
    'Concatenates the new number to the end, takes only 2 digits, and returns the result
    '   The Cint() is to remove leading zeros
    ChangeItemAmount = Right(CStr(CInt(lblAmount.Caption & CStr(Number))), 2)
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    'The player is returned to the field
    frmField.Show
End Sub

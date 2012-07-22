VERSION 5.00
Begin VB.Form frmBattle 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battle"
   ClientHeight    =   9240
   ClientLeft      =   2985
   ClientTop       =   2550
   ClientWidth     =   13320
   Icon            =   "frmBattle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   13320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTurnClock 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picCursor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9480
      Picture         =   "frmBattle.frx":030A
      ScaleHeight     =   375
      ScaleWidth      =   675
      TabIndex        =   34
      Tag             =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox picTroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Index           =   2
      Left            =   6240
      Picture         =   "frmBattle.frx":1503
      ScaleHeight     =   2250
      ScaleWidth      =   2655
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox picTroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Index           =   1
      Left            =   3240
      Picture         =   "frmBattle.frx":5720
      ScaleHeight     =   2250
      ScaleWidth      =   2655
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox picTroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Index           =   0
      Left            =   240
      Picture         =   "frmBattle.frx":993D
      ScaleHeight     =   2250
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame fraStatus 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   9255
      Begin VB.Label lblAnon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Anon3"
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
         Index           =   2
         Left            =   6855
         TabIndex        =   33
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
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
         Height          =   435
         Index           =   11
         Left            =   6135
         TabIndex        =   32
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP:"
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
         Height          =   435
         Index           =   10
         Left            =   6120
         TabIndex        =   31
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Height          =   435
         Index           =   9
         Left            =   7815
         TabIndex        =   30
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Height          =   435
         Index           =   8
         Left            =   7620
         TabIndex        =   29
         Top             =   1920
         Width           =   195
      End
      Begin VB.Label lblCurrHP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Height          =   435
         Index           =   2
         Left            =   6855
         TabIndex        =   28
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblMaxHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Height          =   435
         Index           =   2
         Left            =   8055
         TabIndex        =   27
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblCurrMP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Height          =   435
         Index           =   2
         Left            =   6855
         TabIndex        =   26
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lblMaxMP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Height          =   435
         Index           =   2
         Left            =   7905
         TabIndex        =   25
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lblAnon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Anon2"
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
         Index           =   1
         Left            =   3855
         TabIndex        =   24
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
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
         Height          =   435
         Index           =   7
         Left            =   3135
         TabIndex        =   23
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP:"
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
         Height          =   435
         Index           =   6
         Left            =   3120
         TabIndex        =   22
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Height          =   435
         Index           =   5
         Left            =   4815
         TabIndex        =   21
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Height          =   435
         Index           =   4
         Left            =   4620
         TabIndex        =   20
         Top             =   1920
         Width           =   195
      End
      Begin VB.Label lblCurrHP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Height          =   435
         Index           =   1
         Left            =   3855
         TabIndex        =   19
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblMaxHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Height          =   435
         Index           =   1
         Left            =   5055
         TabIndex        =   18
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblCurrMP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Height          =   435
         Index           =   1
         Left            =   3855
         TabIndex        =   17
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lblMaxMP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Height          =   435
         Index           =   1
         Left            =   4905
         TabIndex        =   16
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lblAnon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Anon1"
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
         Index           =   0
         Left            =   960
         TabIndex        =   15
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
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
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP:"
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
         Height          =   435
         Index           =   2
         Left            =   225
         TabIndex        =   13
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Height          =   435
         Index           =   1
         Left            =   1920
         TabIndex        =   12
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Height          =   435
         Index           =   3
         Left            =   1725
         TabIndex        =   11
         Top             =   1920
         Width           =   195
      End
      Begin VB.Label lblCurrHP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Height          =   435
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblMaxHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Height          =   435
         Index           =   0
         Left            =   2160
         TabIndex        =   9
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblCurrMP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Height          =   435
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lblMaxMP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Height          =   435
         Index           =   0
         Left            =   2010
         TabIndex        =   7
         Top             =   1920
         Width           =   675
      End
   End
   Begin VB.Frame fraOption 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   9255
      Index           =   1
      Left            =   9240
      TabIndex        =   2
      Top             =   0
      Width           =   4095
      Begin VB.Label lblHaxLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " MP Cost"
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
         Height          =   855
         Index           =   1
         Left            =   3000
         TabIndex        =   63
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblHaxLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hax"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   0
         Left            =   1080
         TabIndex        =   62
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblMPNeeded 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "90"
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
         Left            =   2880
         TabIndex        =   61
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lblMPNeeded 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "70"
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
         Left            =   2880
         TabIndex        =   60
         Top             =   4800
         Width           =   795
      End
      Begin VB.Label lblMPNeeded 
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
         Index           =   2
         Left            =   2880
         TabIndex        =   59
         Top             =   3480
         Width           =   795
      End
      Begin VB.Label lblMPNeeded 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Left            =   2880
         TabIndex        =   58
         Top             =   2520
         Width           =   795
      End
      Begin VB.Label lblMPNeeded 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10"
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
         Height          =   615
         Index           =   0
         Left            =   2880
         TabIndex        =   57
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lblHax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "404"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Index           =   4
         Left            =   720
         TabIndex        =   42
         Top             =   6000
         Width           =   2025
      End
      Begin VB.Label lblHax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B& Hammer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1200
         Index           =   3
         Left            =   720
         TabIndex        =   41
         Top             =   4440
         UseMnemonic     =   0   'False
         Width           =   2025
      End
      Begin VB.Label lblHax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DDoS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Index           =   2
         Left            =   720
         TabIndex        =   40
         Top             =   3480
         Width           =   2025
      End
      Begin VB.Label lblHax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Spam"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Index           =   1
         Left            =   720
         TabIndex        =   39
         Top             =   2400
         Width           =   2025
      End
      Begin VB.Label lblHax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Index           =   0
         Left            =   720
         TabIndex        =   38
         Top             =   1440
         Width           =   2025
      End
   End
   Begin VB.Frame fraOption 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   9255
      Index           =   0
      Left            =   9240
      TabIndex        =   1
      Top             =   0
      Width           =   4095
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Index           =   2
         Left            =   1320
         TabIndex        =   37
         Top             =   3600
         Width           =   1665
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hax"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Index           =   1
         Left            =   1320
         TabIndex        =   36
         Top             =   2280
         Width           =   1665
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Attack"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Index           =   0
         Left            =   1320
         TabIndex        =   35
         Top             =   960
         Width           =   1665
      End
   End
   Begin VB.Frame fraOption 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   9255
      Index           =   2
      Left            =   9240
      TabIndex        =   3
      Top             =   0
      Width           =   4095
      Begin VB.Label lblItemName 
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
         Left            =   720
         TabIndex        =   56
         Top             =   840
         Width           =   1995
      End
      Begin VB.Label lblItemName 
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
         Left            =   720
         TabIndex        =   55
         Top             =   1560
         Width           =   1995
      End
      Begin VB.Label lblItemName 
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
         Left            =   720
         TabIndex        =   54
         Top             =   2280
         Width           =   1995
      End
      Begin VB.Label lblItemName 
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
         Left            =   720
         TabIndex        =   53
         Top             =   3000
         Width           =   1995
      End
      Begin VB.Label lblItemName 
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
         Left            =   720
         TabIndex        =   52
         Top             =   3720
         Width           =   2595
      End
      Begin VB.Label lblItemName 
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
         Left            =   720
         TabIndex        =   51
         Top             =   4440
         Width           =   2835
      End
      Begin VB.Label lblItemName 
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
         Left            =   720
         TabIndex        =   50
         Top             =   5160
         Width           =   1995
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "99"
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
         Left            =   3060
         TabIndex        =   49
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "99"
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
         Left            =   3060
         TabIndex        =   48
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "99"
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
         Left            =   3060
         TabIndex        =   47
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "99"
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
         Left            =   3060
         TabIndex        =   46
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "99"
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
         Left            =   3060
         TabIndex        =   45
         Top             =   3720
         Width           =   795
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "99"
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
         Left            =   3060
         TabIndex        =   44
         Top             =   4440
         Width           =   795
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "99"
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
         Left            =   3060
         TabIndex        =   43
         Top             =   5160
         Width           =   795
      End
   End
   Begin VB.Image imgBoss 
      Height          =   4815
      Index           =   8
      Left            =   2160
      Picture         =   "frmBattle.frx":DB5A
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Image imgBoss 
      Height          =   4575
      Index           =   7
      Left            =   2040
      Picture         =   "frmBattle.frx":2CA96
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image imgBoss 
      Height          =   5655
      Index           =   6
      Left            =   3000
      Picture         =   "frmBattle.frx":2E5E0
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image imgBoss 
      Height          =   4095
      Index           =   5
      Left            =   2160
      Picture         =   "frmBattle.frx":30E3A
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image imgBoss 
      Height          =   4335
      Index           =   4
      Left            =   2280
      Picture         =   "frmBattle.frx":3A49C
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Image imgBoss 
      Height          =   4575
      Index           =   3
      Left            =   2280
      Picture         =   "frmBattle.frx":5FDB0
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Image imgBoss 
      Height          =   4455
      Index           =   2
      Left            =   2280
      Picture         =   "frmBattle.frx":7630D
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgBoss 
      Height          =   5295
      Index           =   0
      Left            =   2040
      Picture         =   "frmBattle.frx":AE74D
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image imgBoss 
      Height          =   4455
      Index           =   1
      Left            =   2760
      Picture         =   "frmBattle.frx":B2735
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Causes the program to pause for however many milliseconds
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Declares the references to the enemy objects
Dim objEnemy() As clsEnemy

'BossBattle determines if the battle is against a boss (True) or just a random battle (False)
'Area is the current area that the battle is taking place in (it determines the strength of the Trolls or the Boss that is being fought)
Public BossBattle As Boolean, Area As Byte

'CurrCursorLvl determines which label array the cursor is cycling through
'TurnOrder is an array that acts as a queue for character turns (it is a character's turn if their identifier is in index 0). Since 0 is a possible value, the number used to identify that an element is empty is 255 which is an impossible value
Dim CurrCursorLvl As String, TurnOrder() As Byte
'Note: The turn counters for the character are stored in their respective tags (lblAnon().Tag)
'The turn counters for the enemies are stored as a property in their object

'Since the routine that loads this form also has to set this form's BossBattle and Area, and the following code depends on those set variables
'I have to wait until after Form_Load and the routine that sets the variables to run the following code
'Since no other forms of this program are not accessable by the user while this form is loaded, I don't have to worry about Form_Activate running more than once
Private Sub Form_Activate()
    
    'Declare variables
    '   count is the loop counter for initializing TurnOrder()
    Dim count As Integer
    
    'If it is a boss battle
    If BossBattle Then
        'The boss is created
        Call CreateBoss
    'If it is a random battle
    Else
        'Trolls are created
        Call CreateTrolls
    End If
    
    'Initializes turn order to 255 (the "empty" value)
    For count = LBound(TurnOrder()) To UBound(TurnOrder())
        TurnOrder(count) = 255
    Next
    
    'Updates the onscreen elements
    Call RefreshBattleScreen
    
    'Starts the turn clock
    tmrTurnClock.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Declare variables
    '   VerticalCursorMove determines if the cursor moves vertically (True) or horizontally (False) through the array
    Dim VerticalCursorMove As Boolean
    
    'If it is currently the player's turn
    If CheckPlayerTurn Then
        'Determines which direction the cursor is cycling
        VerticalCursorMove = (CurrCursorLvl = "Top") Or (CurrCursorLvl = "Hax") Or (CurrCursorLvl = "Item")
        
        'If the player presses up when the cursor is cycling vertically, the cursor moves up
        If KeyCode = vbKeyUp And VerticalCursorMove Then
            Call MoveCursor(-1)
        'If the player presses down when the cursor is cycling vertically, the cursor moves down
        ElseIf KeyCode = vbKeyDown And VerticalCursorMove Then
            Call MoveCursor(1)
        'If the player presses left when the cursor is cycling horizontally, the cursor moves left
        ElseIf KeyCode = vbKeyLeft And Not VerticalCursorMove Then
            Call MoveCursor(-1)
        'If the player presses right when the cursor is cycling horizontally, the cursor moves right
        ElseIf KeyCode = vbKeyRight And Not VerticalCursorMove Then
            Call MoveCursor(1)
        'If the player presses enter or space, the relevant action is taken
        ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
            Call UserSelection
        'If the player presses B or Decimal(Numpad), the cursor goes up one level
        ElseIf KeyCode = vbKeyB Or KeyCode = vbKeyDecimal Then
            Call CursorBack
        'If the player presses Escape, the player is returned to the field
        ElseIf KeyCode = vbKeyEscape Then
            
            'If the player escapes from a boss battle
            If BossBattle Then
                'Move one area to the left
                Call frmField.ChangeToArea(frmField.NextArea(-1))
            End If
            Unload Me
        
        End If
    End If
End Sub

Private Sub Form_Load()
    
    'Declare variables
    '   count is a loop counter for setting the turn counter for each Anon and setting their Max HP/MP
    Dim count As Integer
    
    'Sets the cursor next to the first label
    Call CursorBack
    
    'Sets the turn counter for each Anon to their starting bonus and sets the Max HP/MP of each Anon
    For count = lblAnon.LBound To lblAnon.ubound
        lblAnon(count).Tag = CStr(StartingTurnCounterBonus(CByte(frmMenu.lblLuck(count).Caption)))
        lblMaxHP(count).Caption = frmMenu.lblMaxHP(count).Caption
        lblMaxMP(count).Caption = frmMenu.lblMaxMP(count).Caption
    Next count
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If the user closes the battle form themselves
    If UnloadMode = vbFormControlMenu Then
        'If it is currently not the player's turn
        If Not CheckPlayerTurn Then
            'Cancels the unloading of the form (they can only escape from battle when it's their turn or they are in a random battle)
            Cancel = True
        'If the player escapes from a boss battle
        ElseIf BossBattle Then
            'Move one area to the left
            Call frmField.ChangeToArea(frmField.NextArea(-1))
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Declare variables
    '   count is the loop counter for dereferencing the enemy object variables
    Dim count As Integer
    
    'Dereferences the enemy object variables causing the object to be removed from memeory
    For count = LBound(objEnemy()) To UBound(objEnemy())
        Set objEnemy(count) = Nothing
    Next count
    
    'The player is returned to the field
    frmField.Show
End Sub

'Checks if the current turn belongs to one of the player's characters
Function CheckPlayerTurn() As Boolean
    'If the current turn is one of the player values it returns True
    'Otherwise. it is the enemy's turn and it returns False
    CheckPlayerTurn = (TurnOrder(0) <= 2)
End Function

'Creates the boss
Sub CreateBoss()
    
    'Declares 1 reference to the enemy objects
    ReDim objEnemy(0) As clsEnemy
    'Creates the boss object
    Set objEnemy(0) = New clsEnemy
    'Allows the object to identify itself
    objEnemy(0).Index = 0
    
    'Makes the picturebox with the correct boss visible
    imgBoss(Area - 2).Visible = True
    
    'Sets the size of the turn order array for 3 players + 1 boss
    ReDim TurnOrder(0 To 3) As Byte
    
End Sub

'Creates a random number of Trolls
Sub CreateTrolls()
    
    'Declare variables
    '   NumberOfTrolls is a random number that determines the number of trolls to be created for this battle
    '   count is a loop counter for the creation of trolls
    Dim NumberOfTrolls As Integer, count As Integer

    'Generates a number between 0 and 2 for the number of trolls to be created
    NumberOfTrolls = Int(3 * Rnd)
    
    'Declares the neccesary number of references to the enemy objects
    ReDim objEnemy(0 To NumberOfTrolls) As clsEnemy
    
    'Creates the troll objects
    For count = 0 To NumberOfTrolls
        Set objEnemy(count) = New clsEnemy
        'Allows the object to identify itself
        objEnemy(count).Index = CByte(count)
        
        'Make the troll's picturebox visible
        picTroll(count).Visible = True
    Next count
    
    'Sets the size of the turn order array for 3 players + the trolls
    ReDim TurnOrder(0 To (NumberOfTrolls + 3)) As Byte
    
End Sub

'Updates the on-screen components for changes made in the calling routine
Sub RefreshBattleScreen()
    
    'Declare variables
    '   count is the loop counter for updating each player/item amount
    Dim count As Integer
    
    'Checks each player
    For count = lblAnon.LBound To lblAnon.ubound
        
        'If the player is dead
        If frmMenu.lblCurrHP(count).Caption = "0" Then
            'Make the total cost label opaque (so that the red back color will be visible)
            lblAnon(count).BackStyle = 1
        'If the player is alive
        Else
            'Make the anon label transparent (so that the red back color will be invisible and the blue frame back color will be shown)
            lblAnon(count).BackStyle = 0
        End If
        
        'Updates HP/MP labels with current values
        lblCurrHP(count).Caption = frmMenu.lblCurrHP(count).Caption
        lblCurrMP(count).Caption = frmMenu.lblCurrMP(count).Caption
        
    Next count
    
    'Updates the number of items
    For count = lblItemAmount.LBound To lblItemAmount.ubound
        lblItemAmount(count).Caption = frmMenu.lblItemAmount(count).Caption
    Next count
    
End Sub

'Moves the cursor through the control array
'   Direction is the direction the cursor moves
'   1 will increase the tag's value (moves down/right corresponding with the control array's index)
'   -1 will decrease the tag's value (moves up/left corresponding with the control array's index)
Sub MoveCursor(Direction As Integer)

    'Declare variables
    '   AcceptablePosition is used to identify when a living troll has been cycled to
    Dim AcceptablePosition As Boolean

    'Cursor tag is adjusted in the desired direction
    picCursor.Tag = CStr(CInt(picCursor.Tag) + Direction)
    
    'Depending on which array the cursor is cycling through, differnt array bounds checks and cursor placements must be made
    'If it is cycling through the top level array
    If CurrCursorLvl = "Top" Then
        'If the tag exceeds the bounds of the array it is cycled
        If CInt(picCursor.Tag) < lblTop.LBound Then
            picCursor.Tag = CStr(lblTop.ubound)
        ElseIf CInt(picCursor.Tag) > lblTop.ubound Then
            picCursor.Tag = CStr(lblTop.LBound)
        End If
        
        'Changes the position of the cursor so that it is next to the new control
        picCursor.Top = lblTop(CInt(picCursor.Tag)).Top + (lblTop(0).Height / 4)
        picCursor.Left = lblTop(CInt(picCursor.Tag)).Left - picCursor.Width + fraOption(0).Left
    
    'If it is cycling through the special abilities (Hax) array
    ElseIf CurrCursorLvl = "Hax" Then
        'If the tag exceeds the bounds of the array it is cycled
        If CInt(picCursor.Tag) < lblHax.LBound Then
            picCursor.Tag = CStr(lblHax.ubound)
        ElseIf CInt(picCursor.Tag) > lblHax.ubound Then
            picCursor.Tag = CStr(lblHax.LBound)
        End If
        
        'Changes the position of the cursor so that it is next to the new control
        picCursor.Top = lblHax(CInt(picCursor.Tag)).Top + (lblHax(0).Height / 4)
        picCursor.Left = lblHax(CInt(picCursor.Tag)).Left - picCursor.Width + fraOption(1).Left
    
    'If it is cycling through the Item array
    ElseIf CurrCursorLvl = "Item" Then
        'If the tag exceeds the bounds of the array it is cycled
        If CInt(picCursor.Tag) < lblItemName.LBound Then
            picCursor.Tag = CStr(lblItemName.ubound)
        ElseIf CInt(picCursor.Tag) > lblItemName.ubound Then
            picCursor.Tag = CStr(lblItemName.LBound)
        End If
        
        'Changes the position of the cursor so that it is next to the new control
        picCursor.Top = lblItemName(CInt(picCursor.Tag)).Top + (lblItemName(0).Height / 4)
        picCursor.Left = lblItemName(CInt(picCursor.Tag)).Left - picCursor.Width + fraOption(2).Left
    
    'If it is cycling through the enemy picturebox array
    ElseIf CurrCursorLvl = "TargetEnemy" Then
        
        'If it is a boss battle or there is only one troll nothing is done since there is only one picture to cycle through
        If UBound(objEnemy) = 0 Then
            'The change to the cursor tag is undone
            picCursor.Tag = CStr(CInt(picCursor.Tag) - Direction)
        Else
            
            'Loop until a living troll has been cycled to
            Do
                'If the tag exceeds the bounds of the array it is cycled
                If CInt(picCursor.Tag) < picTroll.LBound Then
                    picCursor.Tag = CStr(picTroll.ubound)
                ElseIf CInt(picCursor.Tag) > picTroll.ubound Then
                    picCursor.Tag = CStr(picTroll.LBound)
                End If
                
                'If the troll is alive
                If picTroll(CInt(picCursor.Tag)).Visible Then
                    'Causes the loop to end
                    AcceptablePosition = True
                'If the troll is dead
                Else
                    'The cursor is adjusted again
                    picCursor.Tag = CStr(CInt(picCursor.Tag) + Direction)
                End If
            Loop Until AcceptablePosition
            
            'Changes the position of the cursor so that it is next to the new control
            picCursor.Top = picTroll(CInt(picCursor.Tag)).Top + (picTroll(0).Height / 4)
            picCursor.Left = picTroll(CInt(picCursor.Tag)).Left - 50
        End If
    
    'If it is cycling through the Anon label array
    ElseIf CurrCursorLvl = "TargetAlly" Then
        'If the tag exceeds the bounds of the array it is cycled
        If CInt(picCursor.Tag) < lblAnon.LBound Then
            picCursor.Tag = CStr(lblAnon.ubound)
        ElseIf CInt(picCursor.Tag) > lblAnon.ubound Then
            picCursor.Tag = CStr(lblAnon.LBound)
        End If
        
        'Changes the position of the cursor so that it is next to the new control
        picCursor.Top = lblAnon(CInt(picCursor.Tag)).Top + (lblAnon(0).Height / 4) + fraStatus.Top
        picCursor.Left = lblAnon(CInt(picCursor.Tag)).Left - picCursor.Width
    End If
End Sub

'Takes the relevant action based on which label the cursor is next to
Sub UserSelection()
    
    'Declare static variables
    '   AbilityType determines if the player has selected the Attack, Hax, or Item option from the top label array
    '   SpecificAbility determines which Hax or Item the player has selected
    Static AbilityType As Byte, SpecificAbility As Byte
    
    'If the cursor is cycling through the top level array
    If CurrCursorLvl = "Top" Then
        
        'Sets whether the player has selected Attack (0), Hax (1), or Item (2)
        AbilityType = CByte(picCursor.Tag)
        
        'Makes the top level frame invisible and the relevent frame visible
        fraOption(0).Visible = False
        fraOption(AbilityType).Visible = True
        
        'If Attack was selected
        If AbilityType = 0 Then
            'Moves the cursor to the enemy
            Call MoveCursorToEnemy
        'If Hax was selected
        ElseIf AbilityType = 1 Then
            'Adjusts the cursor
            CurrCursorLvl = "Hax"
            picCursor.Tag = "0"
            picCursor.Top = lblHax(0).Top + (lblHax(0).Height / 4)
            picCursor.Left = lblHax(0).Left - picCursor.Width + fraOption(1).Left
        'If Item was selected
        ElseIf AbilityType = 2 Then
            'Adjusts th cursor
            CurrCursorLvl = "Item"
            picCursor.Tag = "0"
            picCursor.Top = lblItemName(0).Top + (lblItemName(0).Height / 4)
            picCursor.Left = lblItemName(0).Left - picCursor.Width + fraOption(2).Left
        End If
    'If the cursor is cycling through the Hax array
    ElseIf CurrCursorLvl = "Hax" Then
        'Sets the currently selected Hax to SpecificAbility
        SpecificAbility = CByte(picCursor.Tag)
        'Moves the cursor to the enemy
        Call MoveCursorToEnemy
        
    'If the cursor is cycling through the Item array
    ElseIf CurrCursorLvl = "Item" Then
        'The player can't select the item if they don't have any of the item
        If lblItemAmount(CInt(picCursor.Tag)).Caption <> "0" Then
            'Sets the currently selected item to SpecificAbility
            SpecificAbility = CByte(picCursor.Tag)
            
            'Sets cursor next to the first Anon label (so the player can choose which Anon to use the item on)
            CurrCursorLvl = "TargetAlly"
            picCursor.Tag = "0"
            picCursor.Top = lblAnon(0).Top + (lblAnon(0).Height / 4) + fraStatus.Top
            picCursor.Left = lblAnon(0).Left - picCursor.Width
        End If
        
    'If the cursor is cycling through the trolls/bosses
    ElseIf CurrCursorLvl = "TargetEnemy" Then
        'Attacks the currently selected enemy
        Call AttackEnemy(AbilityType, SpecificAbility, CByte(picCursor.Tag))
        'Moves the cursor back to the top array
        Call CursorBack
        'Updates on-screen elements
        Call RefreshBattleScreen
        'Checks if the battle has been won
        Call CheckWinLoseConditions

    'If the cursor is cycling through which Anon to use an item on
    ElseIf CurrCursorLvl = "TargetAlly" Then
        'Uses the item on the selected Anon
        Call UseItem(SpecificAbility, CByte(picCursor.Tag))
        'Since a player just took their turn, the turn order array is shifted
        Call ShiftTurn
        'Moves the cursor back to the top array
        Call CursorBack
        'Updates on-screen elements
        Call RefreshBattleScreen
    End If
End Sub

'Returns the cursor to the top level label array
Sub CursorBack()
    'Sets the cursor next to the first label
    CurrCursorLvl = "Top"
    picCursor.Tag = "0"
    picCursor.Top = lblTop(0).Top + (lblTop(0).Height / 4) + fraOption(0).Top
    picCursor.Left = lblTop(0).Left - picCursor.Width + fraOption(0).Left
    'Makes only the top level frame visible
    fraOption(0).Visible = True
    fraOption(1).Visible = False
    fraOption(2).Visible = False
End Sub

'Moves the cursor to the first enemy that is alive
Sub MoveCursorToEnemy()
    
    'Declare variables
    '   TrollID indicates the troll that is being checked for if it is alive
    Dim TrollID As Byte
    
    'Adjusts the array being cycled
    CurrCursorLvl = "TargetEnemy"
    
    'If it's a boss battle
    If BossBattle Then
        'Moves cursor next to the bosses
        picCursor.Left = 1920
        picCursor.Top = 1440
        
    'If it's a random encounter
    Else
        'Checks the trolls until one that is alive is found
        Do Until picTroll(TrollID).Visible Or TrollID > UBound(objEnemy())
            TrollID = TrollID + 1
        Loop
        'Moves cursor next to the troll
        picCursor.Tag = CStr(TrollID)
        picCursor.Top = picTroll(TrollID).Top + (picTroll(TrollID).Height / 4)
        picCursor.Left = picTroll(TrollID).Left - 50
    End If
End Sub

'Deals damage to an enemy
'   AbilityType determines if the player has selected the Attack, Hax, or Item option from the top label array
'   SpecificAbility determines which Hax or Item the player has selected
'   Target is the index of the Troll being attacked
Sub AttackEnemy(AbilityType As Byte, SpecificAbility As Byte, ByVal target As Byte)
    
    'I'm not sure why the program sometimes crashes here
    'Sometimes when an enemy is selected for an attack, VB throws an error in this routine
    'I've spent about an hour and a half going through the code step-by-step, but I still can't find the problem (mostly because it occurs so rarely)
    'This statement will just tell the program to keep going. It doesn't really matter if this error occurs or not, since it will just appear to be a missed attack to the player
    On Error Resume Next
    
    
    'Declare variables
    '   AbilityPower is a value from 0 to 1 that is a multiplier for the base damage calculations (for example: a normal attack has a value of 0.5, and the B& Hammer has a value of 0.9, therefore the B& Hammer will deal more damage)
    '   StatUsed stores the attacker's Strength stat if the Attack option is selected or the Hax stat if a Hax ability is selected (one of these will be used in damage calculations)
    '   damage is the damage done to the enemy if the attack hits
    Dim AbilityPower As Single, StatUsed As Byte, damage As Integer
    
    'Since the player is taking their turn, the turn order is shifted
    Call ShiftTurn
    
    'Determines AbilityPower based on the ability selected
    'If the player selected the attack option
    If AbilityType = 0 Then
        AbilityPower = 0.5
        'Sets Strength as the stat used in the damage calculations
        StatUsed = CByte(frmMenu.lblStrength(TurnOrder(0)).Caption)
    'If they selected a Hax ability
    ElseIf AbilityType = 1 Then
        
        'If the Anon using the ability does not have enough MP
        If CInt(frmMenu.lblCurrMP((TurnOrder(0))).Caption) < CInt(lblMPNeeded(SpecificAbility).Caption) Then
            'The ability power is set to 0 which renders the attack useless
            AbilityPower = 0
        'If he does have enough
        Else
            'The Anon's MP is reduced
            frmMenu.lblCurrMP((TurnOrder(0))).Caption = CStr(CInt(frmMenu.lblCurrMP((TurnOrder(0))).Caption) - CInt(lblMPNeeded(SpecificAbility).Caption))
            
            'The attack power is set
            AbilityPower = (0.1 * SpecificAbility) + 0.6
        End If
        
        'Sets Hax as the stat used in the damage calculations
        StatUsed = CByte(frmMenu.lblHax(TurnOrder(0)).Caption)
        
    End If
    
    'If it is a boss battle the target is changed to 0 (this is always the boss' index), since the value obtained from the cursor tag that was passed to this routine is unusable (the tag was not changed when the cursor was positioned next to the boss)
    If BossBattle Then
        target = 0
    End If
    
    
    'If the attack will successfully hit the enemy
    If HitSuccess() Then
        
        'Damage is calculated
        damage = DamageCalc(StatUsed, CByte(frmMenu.lblLvl(TurnOrder(0)).Caption), objEnemy(target).Defense, AbilityPower)
        
        'Does damage to the enemy (checking if the result is less than 0 is done by the class module)
        objEnemy(target).HP = objEnemy(target).HP - damage
    End If
End Sub

'Checks if the battle has been won or lost and takes action accordingly
Sub CheckWinLoseConditions()
    
    'Declare variables
    '   AllDead is used to check if all of the Anons or all of the enemies are dead
    '   count is a loop counter used to check each Anon/enemy
    Dim AllDead As Boolean, count As Integer
    
    'Assumes everything is dead until a living Anon is found
    AllDead = True
    
    'For each Anon
    For count = frmMenu.lblKO().LBound To frmMenu.lblKO().ubound
        'Checks if they are dead
        AllDead = AllDead And frmMenu.lblKO(count).Visible
    Next count
    
    'If all players are dead
    If AllDead Then
        'Alerts the user that they lost
        MsgBox "All of your dudes are dead." & vbCrLf & vbCrLf & "You lost THE GAME.", vbCritical, "You Lost"
        End
    End If
    
    'Resets Dead in order to check the enemies
    AllDead = True
    
    'For each enemy
    For count = LBound(objEnemy()) To UBound(objEnemy())
        'Checks if they are dead
        AllDead = AllDead And (objEnemy(count).HP = 0)
    Next count
    
    'If all enemies are dead
    If AllDead Then
        Call FightWon
    End If
    
End Sub

'Shifts the elements in the turn order array over 1 position
Sub ShiftTurn()
    'Declare variables
    '   count is the loop counter for shifting the elements in TurnOrder()
    Dim count As Integer
    
    'Sets the turn counter of the character that is currently taking their turn to 0
    'Resets the counter for a player
    If CheckPlayerTurn Then
        lblAnon(TurnOrder(0)).Tag = "0"
    'Resets the counter for aan enemy
    Else
        objEnemy(TurnOrder(0) - 3).ResetTime
    End If
    
    'Sets each element to the next element value (except for the last element)
    For count = LBound(TurnOrder()) To UBound(TurnOrder()) - 1
        TurnOrder(count) = TurnOrder(count + 1)
    Next count
    
    'Sets the last element to 255 (the "empty" value)
    TurnOrder(UBound(TurnOrder())) = 255
    
    'Makes the cursor visible (if it is the player's turn) or invisible
    Call SetCursorVisibility
    
End Sub

'Adds the character to the first available position in the turn order array when their turn counter reachs its max
'   CharacterID identifies the character to be added
'Note that enemies have an ID of 3+ their actual index (since I need to differentiate between player and enemy)
Sub QueueTurn(CharacterID As Byte)
    
    'Declare variables
    '   count is a loop counter used to check each element of the array to see if it is free
    '   CompletedInsertion is used to end the loop when the CharacterID has been inserted into the array
    Dim count As Integer, CompletedInsertion As Boolean
    
    'Loop until the CharacterID has been inserted into the array
    Do
        'If the element is free
        If TurnOrder(count) = 255 Then
            'CharacterID is inserted into the array
            TurnOrder(count) = CharacterID
            'Allows the loop to end
            CompletedInsertion = True
        Else
            'Increments the counter to check the next array element
            count = count + 1
        End If
    Loop Until CompletedInsertion Or count >= UBound(TurnOrder())
    
    'Makes the cursor visible (if it is the player's turn) or invisible
    Call SetCursorVisibility
    
End Sub

'Makes the cursor visible (if it is the player's turn) or invisible (if it is the enemy's turn)
Sub SetCursorVisibility()
    If CheckPlayerTurn Then
        picCursor.Visible = True
    Else
        picCursor.Visible = False
    End If
End Sub

'When all enemies are dead, EXP and items are awarded, and the player is returned to the field
Sub FightWon()
    
    'Declare variables
    '   count is the loop counter used to get EXP and Internets from each enemy
    '   RandomItem identifies a random stat increase item that they player is given if they are in the final area
    Dim count As Integer, RandomItem As Byte
    
    'For each enemy that was defeated
    For count = LBound(objEnemy()) To UBound(objEnemy())
        'Gives EXP earned from this battle to each anon
        Call GiveEXP(objEnemy(count).EXPGained)
        'Gives Internets earned from this battle to the player
        frmMenu.lblInternets.Caption = CStr(CLng(frmMenu.lblInternets.Caption) + objEnemy(count).InternetsGained)
    Next count
    
    'If the player is in the final area
    If Area = 9 Then
        'They get a random stat increase item
        'Generates item
        RandomItem = CByte(Int((7 * Rnd) + 7))
        'Gves the item
        frmMenu.lblItemAmount(RandomItem).Caption = CStr(CByte(frmMenu.lblItemAmount(RandomItem).Caption) + 1)
    End If
    
    'Makes sure stat increases don't exceed their limits
    Call EnforceStatCaps
    
    
    'If the player defeated a boss
    If BossBattle Then
        'The player is allowed to freely move to that area from now on
        frmField.PlayerProgress = frmField.PlayerProgress + 1
    End If
    
    'If the player defeated the final boss
    If frmField.PlayerProgress >= 10 Then
        'Congratulates the player
        MsgBox "Congratulations, you won the game!" & vbCrLf & vbCrLf & "You killed The Final Boss of the Internet! This computer is now meme-free. By the way, even though you won this game, you just lost THE GAME.", vbExclamation, "A Winner Is You!"
        End
    End If
    
    'Waits for a couple seconds before returning to the field
    DoEvents
    Sleep (1000)
    
    'Returns player to the field form
    Unload Me
End Sub

Private Sub tmrTurnClock_Timer()
    
    'Stops the timer so that code in the timer won't run again if it is still executing a previous call to the timer
    tmrTurnClock.Enabled = False
    
    'Declare variables
    '   increase is the value that a character's turn timer goes up by every 0.1 seconds
    '   count is the loop counter for increasing the turn counters of each character and checking that no dead character is ever at the start of the queue
    Dim increase As Integer, count As Integer
    
    'For each Anon (and enemy)
    For count = lblAnon().LBound To lblAnon().ubound
        
        'The turn counter won't increase for dead players or for players that have reached their turn and are waiting in the queue
        If Not frmMenu.lblKO(count).Visible And CInt(lblAnon(count).Tag) < 30000 Then
            'Gets the turn timer increase
            increase = TurnIncrease(CByte(frmMenu.lblSpeed(count).Caption))
            
            'If the Anon reaches the max for the turn timer
            If CInt(lblAnon(count).Tag) + increase >= 30000 Then
                lblAnon(count).Tag = "30000"
                'They are added to the turn queue
                Call QueueTurn(CByte(count))
            'Otherwise their turn counter is simply increased
            Else
                lblAnon(count).Tag = CStr(CInt(lblAnon(count).Tag) + increase)
            End If
        End If
        
        'Makes sure that no dead anon is in the start of the turn queue
        If frmMenu.lblKO(count).Visible And TurnOrder(0) = CByte(count) Then
            Call ShiftTurn
        End If
        
        'If an enemy has been created at this index value
        If count <= UBound(objEnemy()) Then
            'The turn counter won't increase for dead enemies
            If objEnemy(count).HP > 0 Then
                'Increases the turn timer of the enemy
                objEnemy(count).IncreaseTime
            Else
                'Makes sure that no dead enemy is in the start of the turn queue
                If TurnOrder(0) = CByte(count + 3) Then
                    Call ShiftTurn
                End If
            End If
        End If
        
    Next count
    
    'If it is currently an enemy's turn (the less than 255 is because 255 is the "empty" value)
    If TurnOrder(0) >= 3 And TurnOrder(0) < 255 Then
        'Pauses for 1.5 second before and after the attack so that the player won't be startled by several attacks in a fraction of a second
        Sleep (1500)
        
        'The enemy attacks the player
        objEnemy(TurnOrder(0) - 3).AttackPlayer
        'Refresh the label values
        Call RefreshBattleScreen
        Sleep (1500)
        'Checks if that attack resulted in all the players being killed
        Call CheckWinLoseConditions
    End If
    
    'Re-enables the timer
    tmrTurnClock.Enabled = True
    
End Sub

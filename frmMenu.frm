VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   11310
   ClientLeft      =   4440
   ClientTop       =   1050
   ClientWidth     =   10935
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11310
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCursor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      Picture         =   "frmMenu.frx":030A
      ScaleHeight     =   375
      ScaleWidth      =   675
      TabIndex        =   110
      Tag             =   "0"
      Top             =   1200
      Width           =   675
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
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
         Height          =   630
         Index           =   2
         Left            =   1080
         TabIndex        =   3
         Top             =   7080
         Width           =   1755
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Height          =   630
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   4200
         Width           =   1755
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
         Height          =   630
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   1080
         Width           =   1635
      End
   End
   Begin VB.Frame fraStatus 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   11175
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   6975
      Begin VB.Label lblInternets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "123456789"
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
         Left            =   1800
         TabIndex        =   109
         Top             =   10440
         Width           =   1755
      End
      Begin VB.Label lblStatusLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Internets:"
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
         Index           =   36
         Left            =   120
         TabIndex        =   108
         Top             =   10440
         Width           =   1590
      End
      Begin VB.Label lblLuck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Left            =   5040
         TabIndex        =   107
         Top             =   9960
         Width           =   585
      End
      Begin VB.Label lblLuck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Index           =   1
         Left            =   5040
         TabIndex        =   106
         Top             =   6840
         Width           =   585
      End
      Begin VB.Label lblLuck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Index           =   0
         Left            =   5040
         TabIndex        =   105
         Top             =   3720
         Width           =   585
      End
      Begin VB.Label lblHax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Left            =   5040
         TabIndex        =   104
         Top             =   9480
         Width           =   585
      End
      Begin VB.Label lblHax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Index           =   1
         Left            =   5040
         TabIndex        =   103
         Top             =   6360
         Width           =   585
      End
      Begin VB.Label lblHax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Index           =   0
         Left            =   5040
         TabIndex        =   102
         Top             =   3240
         Width           =   585
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Left            =   5040
         TabIndex        =   101
         Top             =   9000
         Width           =   585
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Index           =   1
         Left            =   5040
         TabIndex        =   100
         Top             =   5880
         Width           =   585
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Index           =   0
         Left            =   5040
         TabIndex        =   99
         Top             =   2760
         Width           =   585
      End
      Begin VB.Label lblDefense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Left            =   5040
         TabIndex        =   98
         Top             =   8520
         Width           =   585
      End
      Begin VB.Label lblDefense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Index           =   1
         Left            =   5040
         TabIndex        =   97
         Top             =   5400
         Width           =   585
      End
      Begin VB.Label lblDefense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Index           =   0
         Left            =   5040
         TabIndex        =   96
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label lblStrength 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Left            =   5040
         TabIndex        =   95
         Top             =   8040
         Width           =   585
      End
      Begin VB.Label lblStrength 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Index           =   1
         Left            =   5040
         TabIndex        =   94
         Top             =   4920
         Width           =   585
      End
      Begin VB.Label lblStrength 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Index           =   0
         Left            =   5040
         TabIndex        =   93
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label lblMaxMP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Left            =   2280
         TabIndex        =   92
         Top             =   8640
         Width           =   585
      End
      Begin VB.Label lblMaxMP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Index           =   1
         Left            =   2280
         TabIndex        =   91
         Top             =   5520
         Width           =   585
      End
      Begin VB.Label lblMaxMP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Index           =   0
         Left            =   2280
         TabIndex        =   90
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label lblCurrMP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Left            =   1320
         TabIndex        =   89
         Top             =   8640
         Width           =   585
      End
      Begin VB.Label lblCurrMP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Index           =   1
         Left            =   1320
         TabIndex        =   88
         Top             =   5520
         Width           =   585
      End
      Begin VB.Label lblCurrMP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999"
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
         Index           =   0
         Left            =   1320
         TabIndex        =   87
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label lblMaxHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Left            =   2160
         TabIndex        =   86
         Top             =   8040
         Width           =   780
      End
      Begin VB.Label lblMaxHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Index           =   1
         Left            =   2160
         TabIndex        =   85
         Top             =   4920
         Width           =   780
      End
      Begin VB.Label lblMaxHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Index           =   0
         Left            =   2160
         TabIndex        =   84
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lblCurrHP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Left            =   1200
         TabIndex        =   83
         Top             =   8040
         Width           =   780
      End
      Begin VB.Label lblCurrHP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Index           =   1
         Left            =   1200
         TabIndex        =   82
         Top             =   4920
         Width           =   780
      End
      Begin VB.Label lblCurrHP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
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
         Index           =   0
         Left            =   1200
         TabIndex        =   81
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9000"
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
         Left            =   5400
         TabIndex        =   80
         Top             =   7440
         Width           =   780
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9000"
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
         Index           =   1
         Left            =   5400
         TabIndex        =   79
         Top             =   4320
         Width           =   780
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9000"
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
         Index           =   0
         Left            =   5400
         TabIndex        =   78
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "99"
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
         Left            =   3720
         TabIndex        =   77
         Top             =   7440
         Width           =   390
      End
      Begin VB.Label lblLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "99"
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
         Index           =   1
         Left            =   3720
         TabIndex        =   76
         Top             =   4320
         Width           =   390
      End
      Begin VB.Label lblLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "99"
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
         Index           =   0
         Left            =   3720
         TabIndex        =   75
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luck:"
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
         Index           =   35
         Left            =   4065
         TabIndex        =   74
         Top             =   9960
         Width           =   870
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hax:"
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
         Index           =   34
         Left            =   4170
         TabIndex        =   73
         Top             =   9480
         Width           =   765
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
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
         Index           =   33
         Left            =   3795
         TabIndex        =   72
         Top             =   9000
         Width           =   1140
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defense:"
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
         Index           =   32
         Left            =   3435
         TabIndex        =   71
         Top             =   8520
         Width           =   1545
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strength:"
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
         Index           =   31
         Left            =   3435
         TabIndex        =   70
         Top             =   8040
         Width           =   1545
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
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
         Index           =   30
         Left            =   4620
         TabIndex        =   69
         Top             =   7440
         Width           =   735
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl:"
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
         Index           =   29
         Left            =   2985
         TabIndex        =   68
         Top             =   7440
         Width           =   615
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Index           =   28
         Left            =   2010
         TabIndex        =   67
         Top             =   8640
         Width           =   165
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Index           =   27
         Left            =   2010
         TabIndex        =   66
         Top             =   8040
         Width           =   165
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP:"
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
         Index           =   26
         Left            =   480
         TabIndex        =   65
         Top             =   8640
         Width           =   645
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
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
         Index           =   25
         Left            =   495
         TabIndex        =   64
         Top             =   8040
         Width           =   615
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luck:"
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
         Index           =   24
         Left            =   4065
         TabIndex        =   63
         Top             =   6840
         Width           =   870
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hax:"
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
         Index           =   23
         Left            =   4170
         TabIndex        =   62
         Top             =   6360
         Width           =   765
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
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
         Index           =   22
         Left            =   3795
         TabIndex        =   61
         Top             =   5880
         Width           =   1140
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defense:"
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
         Index           =   21
         Left            =   3435
         TabIndex        =   60
         Top             =   5400
         Width           =   1545
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strength:"
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
         Index           =   20
         Left            =   3435
         TabIndex        =   59
         Top             =   4920
         Width           =   1545
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
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
         Index           =   19
         Left            =   4620
         TabIndex        =   58
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl:"
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
         Index           =   18
         Left            =   2985
         TabIndex        =   57
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Index           =   17
         Left            =   2010
         TabIndex        =   56
         Top             =   5520
         Width           =   165
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Index           =   16
         Left            =   2010
         TabIndex        =   55
         Top             =   4920
         Width           =   165
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP:"
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
         Index           =   15
         Left            =   480
         TabIndex        =   54
         Top             =   5520
         Width           =   645
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
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
         Index           =   14
         Left            =   495
         TabIndex        =   53
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luck:"
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
         Index           =   13
         Left            =   3990
         TabIndex        =   52
         Top             =   3720
         Width           =   870
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hax:"
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
         Index           =   12
         Left            =   4095
         TabIndex        =   51
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
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
         Index           =   11
         Left            =   3720
         TabIndex        =   50
         Top             =   2760
         Width           =   1140
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defense:"
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
         Index           =   10
         Left            =   3360
         TabIndex        =   49
         Top             =   2280
         Width           =   1545
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strength:"
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
         Index           =   9
         Left            =   3360
         TabIndex        =   48
         Top             =   1800
         Width           =   1545
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
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
         Index           =   8
         Left            =   4545
         TabIndex        =   47
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl:"
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
         Index           =   7
         Left            =   2910
         TabIndex        =   46
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Index           =   6
         Left            =   2010
         TabIndex        =   45
         Top             =   2400
         Width           =   165
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Index           =   5
         Left            =   2010
         TabIndex        =   44
         Top             =   1800
         Width           =   165
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP:"
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
         Index           =   4
         Left            =   525
         TabIndex        =   43
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label lblStatusLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
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
         Index           =   3
         Left            =   540
         TabIndex        =   42
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblAnon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anon3"
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
         Left            =   510
         TabIndex        =   41
         Top             =   7440
         Width           =   1005
      End
      Begin VB.Label lblAnon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anon2"
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
         Index           =   1
         Left            =   510
         TabIndex        =   40
         Top             =   4320
         Width           =   1005
      End
      Begin VB.Label lblAnon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anon1"
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
         Index           =   0
         Left            =   510
         TabIndex        =   39
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblKO 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   570
         Index           =   2
         Left            =   1680
         TabIndex        =   38
         Top             =   7320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblKO 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   570
         Index           =   1
         Left            =   1680
         TabIndex        =   37
         Top             =   4200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblKO 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   570
         Index           =   0
         Left            =   1680
         TabIndex        =   36
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblStatusTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   2160
         TabIndex        =   35
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame fraItem 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   9855
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   6375
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
         Index           =   13
         Left            =   4500
         TabIndex        =   34
         Top             =   9120
         Width           =   1035
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
         Index           =   12
         Left            =   4500
         TabIndex        =   33
         Top             =   8520
         Width           =   1035
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
         Index           =   11
         Left            =   4500
         TabIndex        =   32
         Top             =   7920
         Width           =   1035
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
         Index           =   10
         Left            =   4500
         TabIndex        =   31
         Top             =   7320
         Width           =   1035
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
         Index           =   9
         Left            =   4500
         TabIndex        =   30
         Top             =   6720
         Width           =   1035
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
         Index           =   8
         Left            =   4500
         TabIndex        =   29
         Top             =   6120
         Width           =   1035
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
         Index           =   7
         Left            =   4500
         TabIndex        =   28
         Top             =   5520
         Width           =   1035
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
         Left            =   4500
         TabIndex        =   27
         Top             =   4920
         Width           =   1035
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
         Left            =   4500
         TabIndex        =   26
         Top             =   4320
         Width           =   1035
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
         Left            =   4500
         TabIndex        =   25
         Top             =   3720
         Width           =   1035
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
         Left            =   4500
         TabIndex        =   24
         Top             =   3120
         Width           =   1035
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
         Left            =   4500
         TabIndex        =   23
         Top             =   2520
         Width           =   1035
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
         Left            =   4500
         TabIndex        =   22
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label lblItemTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   2880
         TabIndex        =   21
         Top             =   120
         Width           =   1365
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
         Left            =   4500
         TabIndex        =   20
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label lblItemName 
         BackStyle       =   0  'Transparent
         Caption         =   "Luck+"
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
         Index           =   13
         Left            =   1440
         TabIndex        =   19
         Top             =   9120
         Width           =   1995
      End
      Begin VB.Label lblItemName 
         BackStyle       =   0  'Transparent
         Caption         =   "Hax+"
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
         Index           =   12
         Left            =   1440
         TabIndex        =   18
         Top             =   8520
         Width           =   1995
      End
      Begin VB.Label lblItemName 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed+"
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
         Index           =   11
         Left            =   1440
         TabIndex        =   17
         Top             =   7920
         Width           =   1995
      End
      Begin VB.Label lblItemName 
         BackStyle       =   0  'Transparent
         Caption         =   "Defense+"
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
         Index           =   10
         Left            =   1440
         TabIndex        =   16
         Top             =   7320
         Width           =   1995
      End
      Begin VB.Label lblItemName 
         BackStyle       =   0  'Transparent
         Caption         =   "Strength+"
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
         Index           =   9
         Left            =   1440
         TabIndex        =   15
         Top             =   6720
         Width           =   1995
      End
      Begin VB.Label lblItemName 
         BackStyle       =   0  'Transparent
         Caption         =   "MP+"
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
         Index           =   8
         Left            =   1440
         TabIndex        =   14
         Top             =   6120
         Width           =   1995
      End
      Begin VB.Label lblItemName 
         BackStyle       =   0  'Transparent
         Caption         =   "HP+"
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
         Index           =   7
         Left            =   1440
         TabIndex        =   13
         Top             =   5520
         Width           =   1995
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
         Left            =   1440
         TabIndex        =   12
         Top             =   4920
         Width           =   1995
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
         Left            =   1440
         TabIndex        =   11
         Top             =   4320
         Width           =   2835
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
         Left            =   1440
         TabIndex        =   10
         Top             =   3720
         Width           =   2595
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
         Left            =   1440
         TabIndex        =   9
         Top             =   3120
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
         Left            =   1440
         TabIndex        =   8
         Top             =   2520
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
         Left            =   1440
         TabIndex        =   7
         Top             =   1920
         Width           =   1995
      End
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
         Left            =   1440
         TabIndex        =   6
         Top             =   1320
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CurrCursorLvl determines which label array the cursor is cycling through
'ItemSelected stores the index of the item has been selected by the user
Dim CurrCursorLvl As String, ItemSelected As Byte

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If the player presses up, the cursor moves up
    If KeyCode = vbKeyUp Then
        Call MoveCursor(-1)
    'If the player presses down, the cursor moves down
    ElseIf KeyCode = vbKeyDown Then
        Call MoveCursor(1)
    'If the player presses enter or space, the relevant action is taken
    ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        Call UserSelection
    'If the player presses B or Decimal(Numpad), the cursor goes up one level
    ElseIf KeyCode = vbKeyB Or KeyCode = vbKeyDecimal Then
        Call CursorBack
    End If
End Sub

'Moves the cursor through the label array
'   Direction is the direction the cursor moves
'   1 will increase the tag's value (moves down corresponding with the control array's index)
'   -1 will decrease the tag's value (moves up corresponding with the control array's index)
Sub MoveCursor(Direction As Integer)
    
    'Cursor tag is adjusted in the desired direction
    picCursor.Tag = CStr(CInt(picCursor.Tag) + Direction)
    
    'Depending on which array the cursor is cycling through, differnt array bounds checks and cursor placements must be made
    'If it is cycling through the top level array:
    If CurrCursorLvl = "Top" Then
        'If the tag exceeds the bounds of the array it is cycled
        If CInt(picCursor.Tag) < lblTop.lbound Then
            picCursor.Tag = CStr(lblTop.ubound)
        ElseIf CInt(picCursor.Tag) > lblTop.ubound Then
            picCursor.Tag = CStr(lblTop.lbound)
        End If
        
        'Changes the position of the cursor so that it is next to the new control
        picCursor.Top = lblTop(CInt(picCursor.Tag)).Top + (lblTop(CInt(picCursor.Tag)).Height / 4) + fraTop.Top
        picCursor.Left = lblTop(CInt(picCursor.Tag)).Left - picCursor.Width + fraTop.Left
        
        'Shows the item frame and hides the status frame, if the cursor is on the item option
        If picCursor.Tag = "0" Then
            fraItem.Visible = True
            fraStatus.Visible = False
        'Shows the status frame and hides the item frame if the cursor is on the status option
        ElseIf picCursor.Tag = "1" Then
            fraStatus.Visible = True
            fraItem.Visible = False
        'Hides both frames if the cursor is on the save option
        Else
            fraItem.Visible = False
            fraStatus.Visible = False
        End If
        
    'If it is cycling through the item array:
    ElseIf CurrCursorLvl = "Item" Then
        'If the tag exceeds the bounds of the array it is cycled
        If CInt(picCursor.Tag) < lblItemName.lbound Then
            picCursor.Tag = CStr(lblItemName.ubound)
        ElseIf CInt(picCursor.Tag) > lblItemName.ubound Then
            picCursor.Tag = CStr(lblItemName.lbound)
        End If
        
        'Changes the position of the cursor so that it is next to the new control
        picCursor.Top = lblItemName(CInt(picCursor.Tag)).Top + (lblItemName(CInt(picCursor.Tag)).Height / 4) + fraItem.Top
        picCursor.Left = lblItemName(CInt(picCursor.Tag)).Left - picCursor.Width + fraItem.Left
    'If it is cycling through the array for choosing which character to use an item on:
    ElseIf CurrCursorLvl = "UseOn" Then
        'If the tag exceeds the bounds of the array it is cycled
        If CInt(picCursor.Tag) < lblAnon.lbound Then
            picCursor.Tag = CStr(lblAnon.ubound)
        ElseIf CInt(picCursor.Tag) > lblAnon.ubound Then
            picCursor.Tag = CStr(lblAnon.lbound)
        End If

        'Changes the position of the cursor so that it is next to the new control
        picCursor.Top = lblAnon(CInt(picCursor.Tag)).Top + (lblAnon(0).Height / 4) + fraStatus.Top
        picCursor.Left = lblAnon(CInt(picCursor.Tag)).Left - picCursor.Width + fraStatus.Left
    End If

End Sub

'Takes the relevant action based on which label the cursor is next to
Sub UserSelection()
    
    'If the cursor is cycling through the top level array
    If CurrCursorLvl = "Top" Then
        'If it is next to the item option
        If picCursor.Tag = "0" Then
            'Changes which control array the cursor is cycling through to the item array
            CurrCursorLvl = "Item"
            
            'Changes the position of the cursor so that it is next to the first control
            picCursor.Top = lblItemName(0).Top + (lblItemName(0).Height / 4) + fraItem.Top
            picCursor.Left = lblItemName(0).Left - picCursor.Width + fraItem.Left
        'If it is next to the save option
        ElseIf picCursor.Tag = "2" Then
            'The player's stats are saved
            Call SaveGame
        End If
    'If it is cycling through the item array
    ElseIf CurrCursorLvl = "Item" Then
        
        'The user will only be given the option to select the character to use the item on if the player actually has the item
        If CByte(lblItemAmount(CInt(picCursor.Tag)).Caption) > 0 Then
            'Stores the index of the selected item
            ItemSelected = CByte(picCursor.Tag)
            
            'Changes which control array the cursor is cycling through for selecting which Anon to use the item on
            CurrCursorLvl = "UseOn"
            
            'Shows the status frame so the user can make a selection
            fraItem.Visible = False
            fraStatus.Visible = True
            
            'Changes the position of the cursor so that it is next to the first control
            picCursor.Tag = "0"
            picCursor.Top = lblAnon(0).Top + (lblAnon(0).Height / 4) + fraStatus.Top
            picCursor.Left = lblAnon(0).Left - picCursor.Width + fraStatus.Left
        End If
        
    'If the user is choosing which character to use an item on
    ElseIf CurrCursorLvl = "UseOn" Then
        'Uses the item on the selected character
        Call UseItem(ItemSelected, CByte(picCursor.Tag))
    End If
    
End Sub

'Returns the cursor to the label array one level up
Sub CursorBack()

    'If the cursor is cycling through the top level array
    If CurrCursorLvl = "Top" Then
        'The player is returned to the field form
        frmField.Show
        frmMenu.Hide
    
    'If it is cycling through the item array
    ElseIf CurrCursorLvl = "Item" Then
        'Returns cursor to the top level array
        CurrCursorLvl = "Top"
        
        'Positions cursor next to the Item option
        picCursor.Tag = "0"
        picCursor.Top = lblTop(0).Top + (lblTop(0).Height / 4) + fraTop.Top
        picCursor.Left = lblTop(0).Left - picCursor.Width + fraTop.Left
        
    'If the user is choosing which character to use an item on
    ElseIf CurrCursorLvl = "UseOn" Then
        
        'Returns cursor to item selection
        CurrCursorLvl = "Item"
        
        'Shows the item frame
        fraItem.Visible = True
        fraStatus.Visible = False
        
        'Sets the cursor next to the item that the user had selected
        picCursor.Tag = CStr(ItemSelected)
        picCursor.Top = lblItemName(CInt(picCursor.Tag)).Top + (lblItemName(CInt(picCursor.Tag)).Height / 4) + fraItem.Top
        picCursor.Left = lblItemName(CInt(picCursor.Tag)).Left - picCursor.Width + fraItem.Left

    End If

End Sub

'Attempts to save the player's progress in a file
Sub SaveGame()
    
    'Declare Variables
    '   SaveFileLoc is the path to the save file
    '   SaveData is used to store the data being saved to the save file
    '   count is a loop counter used when adding data to the SaveData string
    Dim SaveFileLoc As String, SaveData As String, count As Integer
    
    'Builds SaveData string with the player's current stats and progress
    '   This is the format of the file:
    '   Colons determine where individual stats start and end
    '   Each individual stat has the name of the stat followed by its value
    '   For example... if Anon1 was at level 99 it would look like this in the save file
    '   ...:Anon1Lvl99:...
    
    'Items:
    For count = frmMenu.lblItemAmount.lbound To frmMenu.lblItemAmount.ubound
        SaveData = SaveData & ":" & frmMenu.lblItemName(count) & frmMenu.lblItemAmount(count).Caption
    Next count
    
    'Player stats:
    For count = frmMenu.lblAnon.lbound To frmMenu.lblAnon.ubound
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "Lvl" & frmMenu.lblLvl(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "Exp" & frmMenu.lblExp(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "CurrHP" & frmMenu.lblCurrHP(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "MaxHP" & frmMenu.lblMaxHP(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "CurrMP" & frmMenu.lblCurrMP(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "MaxMP" & frmMenu.lblMaxMP(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "Strength" & frmMenu.lblStrength(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "Defense" & frmMenu.lblDefense(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "Speed" & frmMenu.lblSpeed(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "Hax" & frmMenu.lblHax(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "Luck" & frmMenu.lblLuck(count).Caption
        SaveData = SaveData & ":" & "Anon" & CStr(count + 1) & "KO" & CStr(frmMenu.lblKO(count).Visible)
    Next count
    
    'Internets (currency):
    SaveData = SaveData & ":" & "Internets" & frmMenu.lblInternets.Caption
    
    'The area that the player can go to without triggering a boss battle:
    SaveData = SaveData & ":" & "PlayerProgress" & CStr(frmField.PlayerProgress)
    
    'Adds the final colon seperator
    SaveData = SaveData & ":"
    'End of save data building
    
    'Encrypts the data
    SaveData = EncryptDecrypt(SaveData, True)
    
    
    'If writing in the directory will cause a permissions error, the On Error statment will prevent the error from becoming fatal
    On Error GoTo ErrorHandler
    
    'Sets the location of the save file
    SaveFileLoc = frmTitle.GetPath("Final_Meme_Save_Data.save")
        
    'If a previous save file exists, delete it
    'With binary files, the data file is overwritten only up to the length of the string
    'For example:
    '   Write "Hello" to a file
    '   The data in the file is: "Hello"
    '   Write "Hi" to the file
    '   The data in the file is: "Hillo"
    'In this case, since the second write was only two characters long, only two characters were overwritten
    'It is easier to just delete the file and make a new one than having a bloated save file with garbage at the end
    If Dir(SaveFileLoc) <> "" Then
        Kill SaveFileLoc
    End If
    
    'Creates and opens the save file
    Open SaveFileLoc For Binary Access Write As #1
        'Writes the save data to the file
        Put #1, 1, SaveData
    Close #1

    'Prevents the error handling code from excecuting if the routine ran without any problems
    Exit Sub
'Error handling code for when there is a permissions error due to trying to write to a read-only directory
ErrorHandler:
    'Alerts the user that the folder is read-only
    MsgBox "The directory that you are trying to save to is read-only. You cannot save your game unless you move the application to a folder that is not read-only.", vbCritical, "Cannot Save"

End Sub

Private Sub Form_Load()
    
    'Sets CurrCursorLvl so the cursor cycles though the lblTop array
    CurrCursorLvl = "Top"
    
    'Sets the cursor next to the first label
    picCursor.Tag = "0"
    picCursor.Top = lblTop(0).Top + (lblTop(0).Height / 4) + fraTop.Top
    picCursor.Left = lblTop(0).Left - picCursor.Width + fraTop.Left
    
    'Makes the item frame visible (since the cursor starts on the item option) and the status frame invisible
    fraItem.Visible = True
    fraStatus.Visible = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If the user accidentally closes the menu form themselves, this prevents the form from being unloaded, since I still need access to information stored in this form
    If UnloadMode = vbFormControlMenu Then
        'Cancels the unloading of the form
        Cancel = True
        'Returns the player to the field from
        frmField.Show
        frmMenu.Hide
    End If
End Sub

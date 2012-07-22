VERSION 5.00
Begin VB.Form frmField 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Field"
   ClientHeight    =   11070
   ClientLeft      =   1740
   ClientTop       =   1755
   ClientWidth     =   14415
   Icon            =   "frmField.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11070
   ScaleWidth      =   14415
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAnon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1710
      Left            =   360
      Picture         =   "frmField.frx":030A
      ScaleHeight     =   1710
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Frame fraArea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8000
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   14415
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   14
         Left            =   6000
         Picture         =   "frmField.frx":109F
         Top             =   5520
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   13
         Left            =   6000
         Picture         =   "frmField.frx":18E8
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   12
         Left            =   10200
         Picture         =   "frmField.frx":2131
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   11
         Left            =   9120
         Picture         =   "frmField.frx":297A
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   10
         Left            =   10560
         Picture         =   "frmField.frx":31C3
         Top             =   5640
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   9
         Left            =   11400
         Picture         =   "frmField.frx":3A0C
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   8
         Left            =   11640
         Picture         =   "frmField.frx":4255
         Top             =   120
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   7
         Left            =   9000
         Picture         =   "frmField.frx":4A9E
         Top             =   120
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   6
         Left            =   10080
         Picture         =   "frmField.frx":52E7
         Top             =   240
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   5
         Left            =   240
         Picture         =   "frmField.frx":5B30
         Top             =   4080
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   4
         Left            =   2640
         Picture         =   "frmField.frx":6379
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   3
         Left            =   1320
         Picture         =   "frmField.frx":6BC2
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   2
         Left            =   1320
         Picture         =   "frmField.frx":740B
         Top             =   5400
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   1
         Left            =   120
         Picture         =   "frmField.frx":7C54
         Top             =   5400
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   0
         Left            =   120
         Picture         =   "frmField.frx":849D
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgStore 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   480
         Picture         =   "frmField.frx":8CE6
         Top             =   240
         Width           =   5625
      End
   End
   Begin VB.Frame fraArea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8000
      Index           =   9
      Left            =   0
      TabIndex        =   11
      Top             =   1680
      Width           =   14415
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   149
         Left            =   8760
         Picture         =   "frmField.frx":C1ED
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   148
         Left            =   7440
         Picture         =   "frmField.frx":CA36
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   147
         Left            =   6240
         Picture         =   "frmField.frx":D27F
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   146
         Left            =   4800
         Picture         =   "frmField.frx":DAC8
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   145
         Left            =   3480
         Picture         =   "frmField.frx":E311
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   144
         Left            =   4440
         Picture         =   "frmField.frx":EB5A
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   143
         Left            =   5640
         Picture         =   "frmField.frx":F3A3
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   142
         Left            =   6960
         Picture         =   "frmField.frx":FBEC
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   141
         Left            =   8160
         Picture         =   "frmField.frx":10435
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   140
         Left            =   8880
         Picture         =   "frmField.frx":10C7E
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   139
         Left            =   7680
         Picture         =   "frmField.frx":114C7
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   138
         Left            =   6600
         Picture         =   "frmField.frx":11D10
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   137
         Left            =   5400
         Picture         =   "frmField.frx":12559
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   136
         Left            =   4200
         Picture         =   "frmField.frx":12DA2
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   135
         Left            =   3120
         Picture         =   "frmField.frx":135EB
         Top             =   6720
         Width           =   1050
      End
   End
   Begin VB.Frame fraArea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8000
      Index           =   8
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   14415
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   134
         Left            =   2760
         Picture         =   "frmField.frx":13E34
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   133
         Left            =   2760
         Picture         =   "frmField.frx":1467D
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   132
         Left            =   2880
         Picture         =   "frmField.frx":14EC6
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   131
         Left            =   2760
         Picture         =   "frmField.frx":1570F
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   130
         Left            =   5760
         Picture         =   "frmField.frx":15F58
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   129
         Left            =   5760
         Picture         =   "frmField.frx":167A1
         Top             =   5520
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   128
         Left            =   5760
         Picture         =   "frmField.frx":16FEA
         Top             =   4320
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   127
         Left            =   5760
         Picture         =   "frmField.frx":17833
         Top             =   3240
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   126
         Left            =   5760
         Picture         =   "frmField.frx":1807C
         Top             =   2040
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   125
         Left            =   8760
         Picture         =   "frmField.frx":188C5
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   124
         Left            =   8760
         Picture         =   "frmField.frx":1910E
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   123
         Left            =   8880
         Picture         =   "frmField.frx":19957
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   122
         Left            =   9000
         Picture         =   "frmField.frx":1A1A0
         Top             =   3360
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   121
         Left            =   9000
         Picture         =   "frmField.frx":1A9E9
         Top             =   4560
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   120
         Left            =   9840
         Picture         =   "frmField.frx":1B232
         Top             =   120
         Width           =   1050
      End
   End
   Begin VB.Frame fraArea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8000
      Index           =   7
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   14415
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   119
         Left            =   5880
         Picture         =   "frmField.frx":1BA7B
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   118
         Left            =   3360
         Picture         =   "frmField.frx":1C2C4
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   117
         Left            =   3720
         Picture         =   "frmField.frx":1CB0D
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   116
         Left            =   4320
         Picture         =   "frmField.frx":1D356
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   115
         Left            =   4800
         Picture         =   "frmField.frx":1DB9F
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   114
         Left            =   5280
         Picture         =   "frmField.frx":1E3E8
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   113
         Left            =   6360
         Picture         =   "frmField.frx":1EC31
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   112
         Left            =   5760
         Picture         =   "frmField.frx":1F47A
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   111
         Left            =   4560
         Picture         =   "frmField.frx":1FCC3
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   110
         Left            =   5640
         Picture         =   "frmField.frx":2050C
         Top             =   3720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   109
         Left            =   3480
         Picture         =   "frmField.frx":20D55
         Top             =   2520
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   108
         Left            =   6960
         Picture         =   "frmField.frx":2159E
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   107
         Left            =   7440
         Picture         =   "frmField.frx":21DE7
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   106
         Left            =   8040
         Picture         =   "frmField.frx":22630
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   105
         Left            =   6960
         Picture         =   "frmField.frx":22E79
         Top             =   2400
         Width           =   1050
      End
   End
   Begin VB.Frame fraArea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8000
      Index           =   6
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   14415
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   104
         Left            =   3240
         Picture         =   "frmField.frx":236C2
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   103
         Left            =   3960
         Picture         =   "frmField.frx":23F0B
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   102
         Left            =   4800
         Picture         =   "frmField.frx":24754
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   101
         Left            =   5760
         Picture         =   "frmField.frx":24F9D
         Top             =   3480
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   100
         Left            =   6720
         Picture         =   "frmField.frx":257E6
         Top             =   4560
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   99
         Left            =   10320
         Picture         =   "frmField.frx":2602F
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   98
         Left            =   10320
         Picture         =   "frmField.frx":26878
         Top             =   5760
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   97
         Left            =   2280
         Picture         =   "frmField.frx":270C1
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   96
         Left            =   2280
         Picture         =   "frmField.frx":2790A
         Top             =   5280
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   95
         Left            =   2280
         Picture         =   "frmField.frx":28153
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   94
         Left            =   10320
         Picture         =   "frmField.frx":2899C
         Top             =   4440
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   93
         Left            =   9960
         Picture         =   "frmField.frx":291E5
         Top             =   3240
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   92
         Left            =   11040
         Picture         =   "frmField.frx":29A2E
         Top             =   3360
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   91
         Left            =   10320
         Picture         =   "frmField.frx":2A277
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   90
         Left            =   3360
         Picture         =   "frmField.frx":2AAC0
         Top             =   6840
         Width           =   1050
      End
   End
   Begin VB.Frame fraArea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8000
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   14415
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   89
         Left            =   10800
         Picture         =   "frmField.frx":2B309
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   88
         Left            =   8160
         Picture         =   "frmField.frx":2BB52
         Top             =   120
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   87
         Left            =   3960
         Picture         =   "frmField.frx":2C39B
         Top             =   6600
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   86
         Left            =   3360
         Picture         =   "frmField.frx":2CBE4
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   85
         Left            =   7440
         Picture         =   "frmField.frx":2D42D
         Top             =   6600
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   84
         Left            =   5880
         Picture         =   "frmField.frx":2DC76
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   83
         Left            =   2760
         Picture         =   "frmField.frx":2E4BF
         Top             =   6600
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   82
         Left            =   4680
         Picture         =   "frmField.frx":2ED08
         Top             =   5160
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   81
         Left            =   9120
         Picture         =   "frmField.frx":2F551
         Top             =   5400
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   80
         Left            =   6840
         Picture         =   "frmField.frx":2FD9A
         Top             =   240
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   79
         Left            =   4560
         Picture         =   "frmField.frx":305E3
         Top             =   720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   78
         Left            =   2160
         Picture         =   "frmField.frx":30E2C
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   77
         Left            =   9720
         Picture         =   "frmField.frx":31675
         Top             =   120
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   76
         Left            =   9000
         Picture         =   "frmField.frx":31EBE
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   75
         Left            =   6000
         Picture         =   "frmField.frx":32707
         Top             =   6600
         Width           =   1050
      End
   End
   Begin VB.Frame fraArea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8000
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   14415
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   74
         Left            =   4560
         Picture         =   "frmField.frx":32F50
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   73
         Left            =   4560
         Picture         =   "frmField.frx":33799
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   72
         Left            =   4320
         Picture         =   "frmField.frx":33FE2
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   71
         Left            =   4320
         Picture         =   "frmField.frx":3482B
         Top             =   5760
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   70
         Left            =   4560
         Picture         =   "frmField.frx":35074
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   69
         Left            =   7440
         Picture         =   "frmField.frx":358BD
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   68
         Left            =   7440
         Picture         =   "frmField.frx":36106
         Top             =   3480
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   67
         Left            =   7440
         Picture         =   "frmField.frx":3694F
         Top             =   4680
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   66
         Left            =   10440
         Picture         =   "frmField.frx":37198
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   65
         Left            =   10680
         Picture         =   "frmField.frx":379E1
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   64
         Left            =   10440
         Picture         =   "frmField.frx":3822A
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   63
         Left            =   10680
         Picture         =   "frmField.frx":38A73
         Top             =   4320
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   62
         Left            =   10680
         Picture         =   "frmField.frx":392BC
         Top             =   5520
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   61
         Left            =   3360
         Picture         =   "frmField.frx":39B05
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   60
         Left            =   3240
         Picture         =   "frmField.frx":3A34E
         Top             =   6840
         Width           =   1050
      End
   End
   Begin VB.Frame fraArea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8000
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   14415
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   59
         Left            =   9360
         Picture         =   "frmField.frx":3AB97
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   58
         Left            =   8280
         Picture         =   "frmField.frx":3B3E0
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   57
         Left            =   3240
         Picture         =   "frmField.frx":3BC29
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   56
         Left            =   6480
         Picture         =   "frmField.frx":3C472
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   55
         Left            =   5400
         Picture         =   "frmField.frx":3CCBB
         Top             =   5280
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   54
         Left            =   11520
         Picture         =   "frmField.frx":3D504
         Top             =   6600
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   53
         Left            =   3000
         Picture         =   "frmField.frx":3DD4D
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   52
         Left            =   8520
         Picture         =   "frmField.frx":3E596
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   51
         Left            =   3360
         Picture         =   "frmField.frx":3EDDF
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   50
         Left            =   9240
         Picture         =   "frmField.frx":3F628
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   49
         Left            =   7680
         Picture         =   "frmField.frx":3FE71
         Top             =   6600
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   48
         Left            =   10440
         Picture         =   "frmField.frx":406BA
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   47
         Left            =   4560
         Picture         =   "frmField.frx":40F03
         Top             =   6480
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   46
         Left            =   6600
         Picture         =   "frmField.frx":4174C
         Top             =   6600
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   45
         Left            =   7200
         Picture         =   "frmField.frx":41F95
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.Frame fraArea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8000
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   14415
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   44
         Left            =   5640
         Picture         =   "frmField.frx":427DE
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   43
         Left            =   6600
         Picture         =   "frmField.frx":43027
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   42
         Left            =   5640
         Picture         =   "frmField.frx":43870
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   41
         Left            =   6600
         Picture         =   "frmField.frx":440B9
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   40
         Left            =   5640
         Picture         =   "frmField.frx":44902
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   39
         Left            =   6600
         Picture         =   "frmField.frx":4514B
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   38
         Left            =   5640
         Picture         =   "frmField.frx":45994
         Top             =   3360
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   37
         Left            =   6120
         Picture         =   "frmField.frx":461DD
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   36
         Left            =   6600
         Picture         =   "frmField.frx":46A26
         Top             =   3360
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   35
         Left            =   7200
         Picture         =   "frmField.frx":4726F
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   34
         Left            =   5040
         Picture         =   "frmField.frx":47AB8
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   33
         Left            =   8160
         Picture         =   "frmField.frx":48301
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   32
         Left            =   3960
         Picture         =   "frmField.frx":48B4A
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   31
         Left            =   9240
         Picture         =   "frmField.frx":49393
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   30
         Left            =   2880
         Picture         =   "frmField.frx":49BDC
         Top             =   6840
         Width           =   1050
      End
   End
   Begin VB.Frame fraArea 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8000
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   14415
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   29
         Left            =   5280
         Picture         =   "frmField.frx":4A425
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   28
         Left            =   6360
         Picture         =   "frmField.frx":4AC6E
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   27
         Left            =   7320
         Picture         =   "frmField.frx":4B4B7
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   26
         Left            =   4200
         Picture         =   "frmField.frx":4BD00
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   25
         Left            =   3120
         Picture         =   "frmField.frx":4C549
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   24
         Left            =   8280
         Picture         =   "frmField.frx":4CD92
         Top             =   0
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   23
         Left            =   9720
         Picture         =   "frmField.frx":4D5DB
         Top             =   6600
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   22
         Left            =   8640
         Picture         =   "frmField.frx":4DE24
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   21
         Left            =   5520
         Picture         =   "frmField.frx":4E66D
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   20
         Left            =   4440
         Picture         =   "frmField.frx":4EEB6
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   19
         Left            =   5640
         Picture         =   "frmField.frx":4F6FF
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   18
         Left            =   6480
         Picture         =   "frmField.frx":4FF48
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   17
         Left            =   7440
         Picture         =   "frmField.frx":50791
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   16
         Left            =   8520
         Picture         =   "frmField.frx":50FDA
         Top             =   5520
         Width           =   1050
      End
      Begin VB.Image imgBlockingTree 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   15
         Left            =   8520
         Picture         =   "frmField.frx":51823
         Top             =   4440
         Width           =   1050
      End
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   19
      Left            =   13320
      Picture         =   "frmField.frx":5206C
      Top             =   9720
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   18
      Left            =   12120
      Picture         =   "frmField.frx":528B5
      Top             =   9720
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   17
      Left            =   10800
      Picture         =   "frmField.frx":530FE
      Top             =   9720
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   16
      Left            =   9480
      Picture         =   "frmField.frx":53947
      Top             =   9720
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   15
      Left            =   8400
      Picture         =   "frmField.frx":54190
      Top             =   9720
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   14
      Left            =   6840
      Picture         =   "frmField.frx":549D9
      Top             =   9720
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   13
      Left            =   5040
      Picture         =   "frmField.frx":55222
      Top             =   9720
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   12
      Left            =   3600
      Picture         =   "frmField.frx":55A6B
      Top             =   9720
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   11
      Left            =   1800
      Picture         =   "frmField.frx":562B4
      Top             =   9720
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   10
      Left            =   360
      Picture         =   "frmField.frx":56AFD
      Top             =   9720
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   9
      Left            =   12960
      Picture         =   "frmField.frx":57346
      Top             =   480
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   8
      Left            =   11640
      Picture         =   "frmField.frx":57B8F
      Top             =   480
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   7
      Left            =   10080
      Picture         =   "frmField.frx":583D8
      Top             =   480
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   6
      Left            =   8280
      Picture         =   "frmField.frx":58C21
      Top             =   480
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   5
      Left            =   6840
      Picture         =   "frmField.frx":5946A
      Top             =   480
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   4
      Left            =   5640
      Picture         =   "frmField.frx":59CB3
      Top             =   480
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   3
      Left            =   3960
      Picture         =   "frmField.frx":5A4FC
      Top             =   480
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   2
      Left            =   2520
      Picture         =   "frmField.frx":5AD45
      Top             =   480
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   1
      Left            =   1200
      Picture         =   "frmField.frx":5B58E
      Top             =   480
      Width           =   1050
   End
   Begin VB.Image imgTree 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   0
      Left            =   120
      Picture         =   "frmField.frx":5BDD7
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label lblAreaIndicator 
      Alignment       =   2  'Center
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'PlayerProgress stores the max area frame index value that the user can move to without triggering a Boss Battle
'When the player defeats a boss, the counter is incrimented allowing them to freely move to and from the next area
Public PlayerProgress As Byte

'The "RECT" type that the IntersectRect API call uses
Private Type RECT
    Left As Long 'The left edge of the rectangle
    Top As Long 'The top edge of the rectangle
    Right As Long 'The right edge of the rectangle
    Bottom As Long 'The bottom edge of the rectangle
End Type

'API call used to check if two rectangles intersect. If they do intersect, the function returns 1, if they don't it returns 0
'   lpDestRect stores the overlapping area of the two rectangles
'   lpSrc1Rect and lpSrc2Rect are the two rectangles that are being checked for intersection
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

'Makes all area forms invisible
Sub AllAreasInvisible()
    'Declare variables
    '   count is the loop counter
    Dim count As Integer
    
    'Makes each area frame invisible
    For count = fraArea.LBound To fraArea.UBound
        fraArea(count).Visible = False
    Next count
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'If the player presses up, the player moves up
    If KeyCode = vbKeyUp Then
        Call MoveGuy(0, -300)
    'If the player presses down, the player moves down
    ElseIf KeyCode = vbKeyDown Then
        Call MoveGuy(0, 300)
    'If the player presses left, the player moves left
    ElseIf KeyCode = vbKeyLeft Then
        Call MoveGuy(-300, 0)
    'If the player presses right, the player moves right
    ElseIf KeyCode = vbKeyRight Then
        Call MoveGuy(300, 0)
    'If the player presses enter or space
    ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        'The store form is opened
        Call EnterStore
    'If the player presses M or +(Numpad)
    ElseIf KeyCode = vbKeyM Or KeyCode = vbKeyAdd Then
        'The menu form is opened
        frmMenu.Show
        frmField.Hide
    End If
End Sub

'Moves the picture of the player
'   X is the horizontal change
'   Y is the vertical change
Sub MoveGuy(X As Integer, Y As Integer)
    
    'Declare static variable
    '   StepsToEncounter counts down everytime the player moves and triggers a random battle when it reaches 0
    Static StepsToEncounter As Integer
    
    'Initialize StepsToEncounter
    'Since StepsToEncounter is reset in the same instance of the routine where it reached 0, this will catch it after it has declared and only then
    If StepsToEncounter = 0 Then
        StepsToEncounter = ResetStepsToEncounter
    End If
    
    'Changes the position of the player
    picAnon.Left = picAnon.Left + X
    picAnon.Top = picAnon.Top + Y
    
    'If the player hits one of the blocking trees or if the move caused the player to exceed the area frame
    If TreeCollision Or picAnon.Top < fraArea(0).Top Or picAnon.Top + picAnon.Height > fraArea(0).Top + fraArea(0).Height Then
        'The player is moved back
        picAnon.Left = picAnon.Left - X
        picAnon.Top = picAnon.Top - Y
    End If
    
    
    'If the player exceeds the left edge of the area frame
    If picAnon.Left <= fraArea(0).Left Then
        'If the player is in the town
        If lblAreaIndicator.Caption = "Town" Then
            'Since there are no more frames to the left (in the game world), the player is simply stopped at the edge
            picAnon.Left = fraArea(0).Left
        'If the player is in any other area
        Else
            'Move one area to the left
            Call ChangeToArea(NextArea(-1))
        End If
    'If the player exceeds the right edge of the area frame
    ElseIf picAnon.Left > fraArea(0).Left + fraArea(0).Width - picAnon.Width Then
        'If the player is in the last area
        If lblAreaIndicator.Caption = "Area 9" Then
            'Since there are no more frames to the right (in the game world), the player is simply stopped at the edge
            picAnon.Left = fraArea(0).Left + fraArea(0).Width - picAnon.Width
            
            'The final boss battle is triggered
            Call Battle(True, 10)
        'If the player is in any other area
        Else
            'Move one area to the right
            Call ChangeToArea(NextArea(1))
        End If
    End If
    
    'If the player is not in the town
    If lblAreaIndicator.Caption <> "Town" Then
        'The number of movements before a random battle is reduced
        StepsToEncounter = StepsToEncounter - 1
        
        'If it reaches 0
        If StepsToEncounter = 0 Then
            'Gets a new value for StepsToEncounter
            StepsToEncounter = ResetStepsToEncounter
            'Triggers a random battle (the rightmost character in the area indicator identifies the current area)
            Call Battle(False, CByte(Right(lblAreaIndicator.Caption, 1)))
        End If
    End If
End Sub

'Opens the store form if the player is close enough to the picture of the store
Sub EnterStore()
    
    'Declare variables
    '   count is a loop counter for restoring the HP/MP of each character
    Dim count As Integer
    
    'If the player is in the town and somewhere in the upper left corner of the frame where the picture of the store is
    If lblAreaIndicator.Caption = "Town" And picAnon.Left <= fraArea(0).Left + imgStore.Left + imgStore.Width And picAnon.Top <= fraArea(0).Top + imgStore.Top + imgStore.Height Then
        'Opens the store
        frmStore.Show
        Me.Hide
        
        'Accessing the store also restores the HP and MP of the party
        For count = frmMenu.lblAnon.LBound To frmMenu.lblAnon.UBound
            frmMenu.lblKO(count).Visible = False
            frmMenu.lblCurrHP(count).Caption = frmMenu.lblMaxHP(count).Caption
            frmMenu.lblCurrMP(count).Caption = frmMenu.lblMaxMP(count).Caption
        Next count
        
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If the user closes the field form themselves
    If UnloadMode = vbFormControlMenu Then
        'The user is asked if they really want to quit
        If MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo + vbDefaultButton2, "Quit?") = vbYes Then
            'Asks the user if they want to save
            If MsgBox("Any unsaved progress will be lost. Do you want to save before quitting?", vbQuestion + vbYesNo + vbDefaultButton1, "Save?") = vbYes Then
                'Saves the player's progress
                frmMenu.SaveGame
            End If
            'Unloads the menu form in order to completely close the program
            Unload frmMenu
        'If they don't want to quit
        Else
            'Cancels the unloading of the form
            Cancel = True
        End If
    End If
End Sub

'Generates a random number for the number of times the player can move before triggering a random battle
Function ResetStepsToEncounter() As Integer
    'Generates and returns a number between 20 and 35
    ResetStepsToEncounter = Int((16) * Rnd + 20)
End Function

'Checks if the player's picturebox has intersected one of the blocking trees
Function TreeCollision() As Boolean
    
    'Declare variables
    '   rectPlayer stores the dimensions of the player's picturebox
    '   rectTree stores the dimensions of the blocking trees
    '   rectOverlap stores the area where the two controls overlap (it's neccessary for the API call to have it, but since the API call returns a 1 if there is an overlap, that's what is used to determine if there has been a collision)
    '   result stores the result of the API call (True = collision, False = no collision)
    '   count is a loop counter used to check the trees for collisions (it starts at the lowest index value for the trees that are currently displayed)
    '   LastTree is the highest index value of the trees that are currently displayed
    Dim rectPlayer As RECT, rectTree As RECT, rectOverlap As RECT, result As Boolean, count As Integer, LastTree As Integer
    
    'Sets the dimensions of the player's picturebox
    rectPlayer.Top = picAnon.Top
    rectPlayer.Left = picAnon.Left
    rectPlayer.Bottom = picAnon.Top + picAnon.Height
    rectPlayer.Right = picAnon.Left + picAnon.Width
    
    'Since each area has 15 blocking trees, the starting and ending values can be found by using the area value
    'The only exception is area 0 which does not have the area value in the area indicator
    If lblAreaIndicator.Caption = "Town" Then
        'Starts the counter at the first tree that is displayed in the current area
        count = 0
        'Sets LastTree to the last tree that is displayed in the current area
        LastTree = 14
    'If the player is in any other area, a formula is used to find the values
    Else
        count = CInt(Right(lblAreaIndicator.Caption, 1)) * 15
        LastTree = count + 14
    End If
     
    'Ckecks each blocking tree
    Do
        'Sets the dimensions of the tree's picturebox
        rectTree.Top = imgBlockingTree(count).Top + fraArea(0).Top
        rectTree.Left = imgBlockingTree(count).Left + fraArea(0).Left
        rectTree.Bottom = rectTree.Top + imgBlockingTree(count).Height
        rectTree.Right = rectTree.Left + imgBlockingTree(count).Width
        
        'Checks if the player intercects the tree
        result = CBool(IntersectRect(rectOverlap, rectPlayer, rectTree))
        
        'Increments the loop counter
        count = count + 1
        
    'Loops until a tree is hit or all trees have been checked
    Loop Until result Or count > LastTree
    
    'Returns the result
    TreeCollision = result

End Function

'Finds the index value of the area the player is going to and repositions the player
'   Direction is which way the player is moving
'   -1 is to the left
'   1 is to the right
Function NextArea(Direction As Integer) As Byte
    
    'Declare variables
    '   AreaValue is the index value of the frame that the player is moving to
    Dim AreaValue As Byte
    
    'If the player is in the town, the only option is to move to area 1
    If lblAreaIndicator.Caption = "Town" Then
        AreaValue = 1
    'If they are in any other area
    Else
        'The index is found by changing the current area (found in the area indicator label) with the direction that the player is moving
        AreaValue = CByte(CInt(Right(lblAreaIndicator.Caption, 1)) + Direction)
    End If
    
    'If the player is moving from the left side of the area to the right...
    If Direction = -1 Then
        '...the player is positioned on the right side of the current area
        picAnon.Left = fraArea(0).Left + fraArea(0).Width - picAnon.Width
    'If the player is moving from the right side of the area to the left...
    Else
        '...the player is positioned on the left side of the current area
        picAnon.Left = fraArea(0).Left
        
        'If the player has not yet defeated the Boss in order to move to that area
        If AreaValue > PlayerProgress Then
            'The Boss battle is triggered
            Call Battle(True, AreaValue)
        End If
    End If
    
    'Returns the index value
    NextArea = AreaValue
    
End Function

'Makes the frame that the player is moving to visible and adjusts the area indicator label
Sub ChangeToArea(AreaIndex As Byte)
    
    'Makes all area frames invisible except for the one being moved to
    Call AllAreasInvisible
    fraArea(AreaIndex).Visible = True
    
    'If the player is moving to the town
    If AreaIndex = 0 Then
        'The indicator shows "Town" instead of "Area 0"
        lblAreaIndicator.Caption = "Town"
    'If they are moving to any other area
    Else
        'The indicator shows the area value
        lblAreaIndicator.Caption = "Area " & CStr(AreaIndex)
    End If
End Sub

'Triggers a battle
'   BossBattle determines if the battle is against a boss (True) or just a random battle (False)
'   Area is the current area that the battle is taking place in (it determines the strength of the Trolls or the Boss that is being fought)
Sub Battle(BossBattle As Boolean, Area As Byte)
    
    'Opens the Battle form
    frmBattle.Show
    
    'Sets the battle form's global variables with information about what kind of battle was triggered
    frmBattle.BossBattle = BossBattle
    frmBattle.Area = Area
    
    'If it is the final Boss battle then the field form is no longer needed
    If Area = 10 And BossBattle Then
        Unload Me
    'Otherwise the player will need to be returned to the field form after they win the battle
    Else
        Me.Hide
    End If
End Sub

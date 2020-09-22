VERSION 5.00
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Messenger Macro Maker"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   47
      Left            =   2640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   178
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picCURRENT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1560
      Picture         =   "frmMAIN.frx":058A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   177
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   119
      Left            =   7080
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   176
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   118
      Left            =   6720
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   175
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   117
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   174
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   116
      Left            =   6000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   173
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   115
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   172
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   114
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   171
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   113
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   170
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   112
      Left            =   4560
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   169
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   111
      Left            =   4200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   168
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   110
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   167
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   109
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   166
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   108
      Left            =   3120
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   165
      Top             =   3360
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   107
      Left            =   7080
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   164
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   106
      Left            =   6720
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   163
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   105
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   162
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   104
      Left            =   6000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   161
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   103
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   160
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   102
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   159
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   101
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   158
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   100
      Left            =   4560
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   157
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   99
      Left            =   4200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   156
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   98
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   155
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   97
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   154
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   96
      Left            =   3120
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   153
      Top             =   3000
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   95
      Left            =   7080
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   152
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   94
      Left            =   6720
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   151
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   93
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   150
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   92
      Left            =   6000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   149
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   91
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   148
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   90
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   147
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   89
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   146
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   88
      Left            =   4560
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   145
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   87
      Left            =   4200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   144
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   86
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   143
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   85
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   142
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   84
      Left            =   3120
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   141
      Top             =   2640
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   83
      Left            =   7080
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   140
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   82
      Left            =   6720
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   139
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   81
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   138
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   80
      Left            =   6000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   137
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   79
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   136
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   78
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   135
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   77
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   134
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   76
      Left            =   4560
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   133
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   75
      Left            =   4200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   132
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   74
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   131
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   73
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   130
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   72
      Left            =   3120
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   129
      Top             =   2280
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   71
      Left            =   7080
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   128
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   70
      Left            =   6720
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   127
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   69
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   126
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   68
      Left            =   6000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   125
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   67
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   124
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   66
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   123
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   65
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   122
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   64
      Left            =   4560
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   121
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   63
      Left            =   4200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   120
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   62
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   119
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   61
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   118
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   60
      Left            =   3120
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   117
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   59
      Left            =   7080
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   116
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   58
      Left            =   6720
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   115
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   57
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   114
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   56
      Left            =   6000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   113
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   55
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   112
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   54
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   111
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   53
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   110
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   52
      Left            =   4560
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   109
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   51
      Left            =   4200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   108
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   50
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   107
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   49
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   106
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   48
      Left            =   3120
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   105
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   47
      Left            =   7080
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   104
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   46
      Left            =   6720
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   103
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   45
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   102
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   44
      Left            =   6000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   101
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   43
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   100
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   42
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   99
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   41
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   98
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   40
      Left            =   4560
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   97
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   39
      Left            =   4200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   96
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   38
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   95
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   37
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   94
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   36
      Left            =   3120
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   93
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   35
      Left            =   7080
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   92
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   34
      Left            =   6720
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   91
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   33
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   90
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   32
      Left            =   6000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   89
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   31
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   88
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   30
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   87
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   29
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   86
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   28
      Left            =   4560
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   85
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   27
      Left            =   4200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   84
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   26
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   83
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   25
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   82
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   24
      Left            =   3120
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   81
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   23
      Left            =   7080
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   80
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   22
      Left            =   6720
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   79
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   21
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   78
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   20
      Left            =   6000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   77
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   19
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   76
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   18
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   75
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   17
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   74
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   16
      Left            =   4560
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   73
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   15
      Left            =   4200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   72
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   14
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   71
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   13
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   70
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   12
      Left            =   3120
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   69
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   11
      Left            =   7080
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   68
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   10
      Left            =   6720
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   67
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   9
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   66
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   8
      Left            =   6000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   65
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   7
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   64
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   6
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   63
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   62
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   4560
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   61
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   4200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   60
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   59
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   58
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picVAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   3120
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   57
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   46
      Left            =   120
      Picture         =   "frmMAIN.frx":0B4A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   56
      Tag             =   ";)"
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   45
      Left            =   2280
      Picture         =   "frmMAIN.frx":110A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   55
      Tag             =   "(W)"
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   44
      Left            =   1920
      Picture         =   "frmMAIN.frx":16CA
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   54
      Tag             =   ":|"
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   43
      Left            =   1560
      Picture         =   "frmMAIN.frx":1C8A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   53
      Tag             =   ":P"
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   42
      Left            =   1560
      Picture         =   "frmMAIN.frx":224A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   52
      Tag             =   "(Y)"
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   41
      Left            =   1200
      Picture         =   "frmMAIN.frx":280A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   51
      Tag             =   "(N)"
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   40
      Left            =   1200
      Picture         =   "frmMAIN.frx":2DCA
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   50
      Tag             =   ":D"
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   39
      Left            =   120
      Picture         =   "frmMAIN.frx":338A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   49
      Tag             =   "(#)"
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   38
      Left            =   1560
      Picture         =   "frmMAIN.frx":3842
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   48
      Tag             =   "(*)"
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   37
      Left            =   840
      Picture         =   "frmMAIN.frx":3E02
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   47
      Tag             =   "(H)"
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   36
      Left            =   480
      Picture         =   "frmMAIN.frx":43C2
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   46
      Tag             =   ":("
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   35
      Left            =   1920
      Picture         =   "frmMAIN.frx":4982
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   45
      Tag             =   "(F)"
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   34
      Left            =   120
      Picture         =   "frmMAIN.frx":4F42
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   44
      Tag             =   ":)"
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   33
      Left            =   840
      Picture         =   "frmMAIN.frx":5502
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   43
      Tag             =   "(R)"
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   32
      Left            =   2280
      Picture         =   "frmMAIN.frx":59CA
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   42
      Tag             =   "(G)"
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   31
      Left            =   2280
      Picture         =   "frmMAIN.frx":5F8A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   41
      Tag             =   "(T)"
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   30
      Left            =   2280
      Picture         =   "frmMAIN.frx":654A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   40
      Tag             =   ":O"
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   29
      Left            =   1200
      Picture         =   "frmMAIN.frx":6B0A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   39
      Tag             =   "(8)"
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   28
      Left            =   480
      Picture         =   "frmMAIN.frx":70CA
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   38
      Tag             =   "(S)"
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   27
      Left            =   1920
      Picture         =   "frmMAIN.frx":768A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   37
      Tag             =   "(M)"
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   26
      Left            =   480
      Picture         =   "frmMAIN.frx":7C4A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   36
      Tag             =   "(D)"
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   25
      Left            =   2640
      Picture         =   "frmMAIN.frx":820A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   35
      Tag             =   "(I)"
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   24
      Left            =   1920
      Picture         =   "frmMAIN.frx":87CA
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   34
      Tag             =   "(@)"
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   23
      Left            =   840
      Picture         =   "frmMAIN.frx":8D8A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   33
      Tag             =   "(K)"
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   22
      Left            =   120
      Picture         =   "frmMAIN.frx":934A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   32
      Tag             =   "(L)"
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   21
      Left            =   2640
      Picture         =   "frmMAIN.frx":990A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   31
      Tag             =   "(%)"
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   20
      Left            =   1200
      Picture         =   "frmMAIN.frx":9CC2
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   30
      Tag             =   "(Z)"
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   19
      Left            =   480
      Picture         =   "frmMAIN.frx":A282
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   29
      Tag             =   "(})"
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   18
      Left            =   840
      Picture         =   "frmMAIN.frx":A842
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   28
      Tag             =   "(X)"
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   17
      Left            =   840
      Picture         =   "frmMAIN.frx":AE02
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   27
      Tag             =   "(~)"
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   16
      Left            =   1920
      Picture         =   "frmMAIN.frx":B3C2
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   26
      Tag             =   "(E)"
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   15
      Left            =   840
      Picture         =   "frmMAIN.frx":B982
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   25
      Tag             =   ":$"
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   14
      Left            =   120
      Picture         =   "frmMAIN.frx":BF42
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   24
      Tag             =   "({)"
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   13
      Left            =   2640
      Picture         =   "frmMAIN.frx":C502
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   23
      Tag             =   "(6)"
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   12
      Left            =   480
      Picture         =   "frmMAIN.frx":CAC2
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   22
      Tag             =   ":'("
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   11
      Left            =   2640
      Picture         =   "frmMAIN.frx":D082
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   21
      Tag             =   ":S"
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   10
      Left            =   2640
      Picture         =   "frmMAIN.frx":D642
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   20
      Tag             =   "(C)"
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   9
      Left            =   2280
      Picture         =   "frmMAIN.frx":DC02
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   19
      Tag             =   "(O)"
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   8
      Left            =   1200
      Picture         =   "frmMAIN.frx":E1C2
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   18
      Tag             =   "(P)"
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   7
      Left            =   1920
      Picture         =   "frmMAIN.frx":E782
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   17
      Tag             =   "(^)"
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   6
      Left            =   480
      Picture         =   "frmMAIN.frx":ED42
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   16
      Tag             =   "(U)"
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   1560
      Picture         =   "frmMAIN.frx":F302
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   15
      Tag             =   "(&)"
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   120
      Picture         =   "frmMAIN.frx":F8C2
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   14
      Tag             =   "(B)"
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   1560
      Picture         =   "frmMAIN.frx":FE82
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   13
      Tag             =   ":["
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   2280
      Picture         =   "frmMAIN.frx":10442
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   12
      Tag             =   "(?)"
      Top             =   840
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   1560
      Picture         =   "frmMAIN.frx":1082E
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   11
      Tag             =   ":@"
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   1200
      Picture         =   "frmMAIN.frx":10DEE
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   10
      Tag             =   "(A)"
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdTOGGLE 
      Caption         =   "Toggle Grid Size."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdLOAD 
      Caption         =   "Load a saved macro."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "Save this macro."
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton cmdCOPY 
      Caption         =   "Copy to Clipboard."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton cmdCLEARMATCH 
      Caption         =   "Clear items matching current."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton cmdFILL 
      Caption         =   "Fill blanks with current."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton cmdCLEAR 
      Caption         =   "&Clear All"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtCURRENT 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   ":)"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtOUT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label lblLABEL 
      Caption         =   "Current:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum sSizeMode
    smTWELVE = 0
    smFIFTEEN = 1
End Enum
Public SizeMode As sSizeMode
Private Sub cmdCLEAR_Click()
    On Error Resume Next
    Dim X As Long
    For X = 0 To Me.picVAL.UBound
        Set Me.picVAL(X).Picture = Nothing
        Me.picVAL(X).Tag = ""
    Next X
    SetText
End Sub
Private Sub cmdCLEARMATCH_Click()
    On Error Resume Next
    Dim X As Long
    For X = 0 To Me.picVAL.UBound
        If Me.picVAL(X).Tag = Me.txtCURRENT.Text Then
            Me.picVAL(X).Tag = ""
            Set picVAL(X).Picture = Nothing
        End If
    Next X
    SetText
End Sub
Private Sub cmdFILL_Click()
    On Error Resume Next
    Dim X As Long
    For X = 0 To Me.picVAL.UBound
        If "" & Me.picVAL(X).Tag = "" Then
            Me.picVAL(X).Tag = Me.txtCURRENT
            Set picVAL(X).Picture = picCURRENT.Picture
        End If
    Next X
    SetText
End Sub

Private Sub cmdLOAD_Click()
    On Error Resume Next
    Dim fNAME As String, obj As CDLG
    Me.SizeMode = smTWELVE
    Set obj = New CDLG
    obj.VBGetOpenFileName fNAME, , , , , , "Macro Maker Files (*.mkf)|*.mkf|All Files (*.*)|*.*", , CurDir, "Load Macro", "*.mkf", Me.hwnd
    If fNAME <> "" Then
        LoadFile fNAME
    End If
End Sub

Private Sub cmdSAVE_Click()
    On Error Resume Next
    Dim fNAME As String
    Dim obj As CDLG
    Set obj = New CDLG
    obj.VBGetSaveFileName fNAME, , , "Macro Maker Files (*.mkf)|*.mkf|All Files (*.*)|*.*", , CurDir, "Save Macro", "*.mmm", Me.hwnd
    If fNAME <> "" Then
        SaveFile fNAME
    End If
End Sub
Private Sub SaveFile(fNAME As String)
    On Error GoTo ErrorSaveFile
    Dim f As Long, X As Long, v As String
    f = FreeFile
    Open fNAME For Output As #f
    If Me.SizeMode = smTWELVE Then
        Print #f, "0"
    Else
        Print #f, "1"
    End If
    For X = 0 To Me.picVAL.UBound
        v = Me.picVAL(X).Tag
        If v = "" Then v = ":)"
        Print #f, v
    Next X
    Close #f
    Exit Sub
ErrorSaveFile:
    MsgBox Err & ":Error in SaveFile.  Error Message: " _
    & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub LoadFile(fNAME As String)
    On Error GoTo ErrorLoadFile
    Dim l As String, X As Long, f As Long
    X = 0
    f = FreeFile
    Open fNAME For Input As #f
    Do While Not EOF(f)
        Line Input #f, l
        If l = "0" Or l = "1" Then
            ToggleSize CLng(l)
        Else
            Me.picVAL(X).Tag = l
            X = X + 1
        End If
    Loop
    Close #f
    FillPicsFromTags
    SetText
    Exit Sub
ErrorLoadFile:
    MsgBox Err & ":Error in LoadFile.  Error Message: " _
    & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub cmdCOPY_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Me.txtOUT.Text
End Sub

Private Sub ToggleSize(sm As sSizeMode)
    On Error Resume Next
    Dim X As Long
    If sm = smFIFTEEN Then
        For X = 1 To Me.picVAL.UBound
            If X Mod 15 <> 0 Then
                Me.picVAL(X).Move Me.picVAL(X - 1).Left + Me.picVAL(0).Width, Me.picVAL(X - 1).Top, Me.picVAL(0).Width, Me.picVAL(0).Height
            Else
                Me.picVAL(X).Move Me.picVAL(0).Left, Me.picVAL(X - 1).Top + Me.picVAL(0).Height, Me.picVAL(0).Width, Me.picVAL(0).Height
            End If
        Next X
        Me.SizeMode = smFIFTEEN
        Me.txtOUT.Width = Me.picVAL(0).Width * 15
    Else
        For X = 1 To Me.picVAL.UBound
            If X Mod 12 <> 0 Then
                Me.picVAL(X).Move Me.picVAL(X - 1).Left + Me.picVAL(0).Width, Me.picVAL(X - 1).Top, Me.picVAL(0).Width, Me.picVAL(0).Height
            Else
                Me.picVAL(X).Move Me.picVAL(0).Left, Me.picVAL(X - 1).Top + Me.picVAL(0).Height, Me.picVAL(0).Width, Me.picVAL(0).Height
            End If
        Next X
        Me.SizeMode = smTWELVE
        Me.txtOUT.Width = Me.picVAL(0).Width * 12
    End If
    Me.Width = Me.txtOUT.Left + Me.txtOUT.Width + 240
    SetText
End Sub
Private Sub cmdTOGGLE_Click()
    On Error Resume Next
    If Me.SizeMode = smFIFTEEN Then
        ToggleSize smTWELVE
    Else
        ToggleSize smFIFTEEN
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    cmdCLEAR_Click
    ToggleSize smTWELVE
End Sub
Private Sub SetText()
    On Error GoTo ErrorSetText
    Dim X As Long, v As String, oSTR As String
    oSTR = ""
    For X = 0 To Me.picVAL.UBound
        v = "" & Me.picVAL(X).Tag
        If Me.SizeMode = smTWELVE Then
            If v = "" Or (X + 1) Mod 12 = 0 Then
                If v = "" Then v = ":)"
                oSTR = oSTR & v
                If (X + 1) Mod 12 = 0 Then oSTR = oSTR & vbCrLf
            Else
                oSTR = oSTR & v
            End If
        Else
            If v = "" Or (X + 1) Mod 15 = 0 Then
                If v = "" Then v = ":)"
                oSTR = oSTR & v
                If (X + 1) Mod 15 = 0 Then oSTR = oSTR & vbCrLf
            Else
                oSTR = oSTR & v
            End If
        End If
    Next X
    Me.txtOUT.Text = oSTR
    Exit Sub
ErrorSetText:
    MsgBox Err & ":Error in SetText.  Error Message: " _
    & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub picICON_Click(Index As Integer)
    On Error Resume Next
    Me.txtCURRENT.Text = picICON(Index).Tag
    Set Me.picCURRENT.Picture = picICON(Index).Picture
End Sub
Private Sub picVAL_Click(Index As Integer)
    On Error Resume Next
    Set picVAL(Index).Picture = Me.picCURRENT.Picture
    picVAL(Index).Tag = Me.txtCURRENT.Text
    SetText
End Sub
Private Sub FillPicsFromTags()
    On Error GoTo ErrorFillPicsFromTags
    Dim X As Long, Y As Long
    For X = 0 To picVAL.UBound
        For Y = 0 To picICON.UBound
            If picICON(Y).Tag = picVAL(X).Tag Then
                Set picVAL(X).Picture = picICON(Y).Picture
                Exit For
            End If
        Next Y
    Next X
    Exit Sub
ErrorFillPicsFromTags:
    MsgBox Err & ":Error in FillPicsFromTags.  Error Message: " _
    & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dim Colour As Integer"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRave 
      Caption         =   "GET READY TO RAVE!!"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   14160
      Width           =   18015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   18600
      Top             =   14040
   End
   Begin VB.Image Image1 
      Height          =   5070
      Left            =   5760
      Picture         =   "rave.frx":0000
      Top             =   4080
      Width           =   6750
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   220
      Left            =   17520
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   219
      Left            =   17520
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   218
      Left            =   17520
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   217
      Left            =   17520
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   216
      Left            =   17520
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   215
      Left            =   17520
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   214
      Left            =   17520
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   213
      Left            =   17520
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   212
      Left            =   17520
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   211
      Left            =   17520
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   210
      Left            =   17520
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   209
      Left            =   17520
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   208
      Left            =   17520
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   207
      Left            =   16440
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   206
      Left            =   16440
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   205
      Left            =   16440
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   204
      Left            =   16440
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   203
      Left            =   16440
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   202
      Left            =   16440
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   201
      Left            =   16440
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   200
      Left            =   16440
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   199
      Left            =   16440
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   198
      Left            =   16440
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   197
      Left            =   16440
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   196
      Left            =   16440
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   195
      Left            =   16440
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   194
      Left            =   15360
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   193
      Left            =   15360
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   192
      Left            =   15360
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   191
      Left            =   15360
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   190
      Left            =   15360
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   189
      Left            =   15360
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   188
      Left            =   15360
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   187
      Left            =   15360
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   186
      Left            =   15360
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   185
      Left            =   15360
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   184
      Left            =   15360
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   183
      Left            =   15360
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   182
      Left            =   15360
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   181
      Left            =   14280
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   180
      Left            =   14280
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   179
      Left            =   14280
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   178
      Left            =   14280
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   177
      Left            =   14280
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   176
      Left            =   14280
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   175
      Left            =   14280
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   174
      Left            =   14280
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   173
      Left            =   14280
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   172
      Left            =   14280
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   171
      Left            =   14280
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   170
      Left            =   14280
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   169
      Left            =   14280
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   168
      Left            =   13200
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   167
      Left            =   13200
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   166
      Left            =   13200
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   165
      Left            =   13200
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   164
      Left            =   13200
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   163
      Left            =   13200
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   162
      Left            =   13200
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   161
      Left            =   13200
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   160
      Left            =   13200
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   159
      Left            =   13200
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   158
      Left            =   13200
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   157
      Left            =   13200
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   156
      Left            =   13200
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   155
      Left            =   12120
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   154
      Left            =   12120
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   153
      Left            =   12120
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   151
      Left            =   12120
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   149
      Left            =   12120
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   147
      Left            =   12120
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   145
      Left            =   12120
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   143
      Left            =   12120
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   141
      Left            =   12120
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   139
      Left            =   12120
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   137
      Left            =   12120
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   136
      Left            =   12120
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   135
      Left            =   12120
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   152
      Left            =   11040
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   150
      Left            =   9960
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   148
      Left            =   8880
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   146
      Left            =   7800
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   144
      Left            =   6720
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   142
      Left            =   5640
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   140
      Left            =   4560
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   138
      Left            =   3480
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   134
      Left            =   2400
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   133
      Left            =   1320
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   132
      Left            =   240
      Top             =   13080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   131
      Left            =   11040
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   130
      Left            =   11040
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   129
      Left            =   9960
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   128
      Left            =   9960
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   127
      Left            =   8880
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   126
      Left            =   8880
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   125
      Left            =   7800
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   124
      Left            =   7800
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   123
      Left            =   6720
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   122
      Left            =   6720
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   121
      Left            =   5640
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   120
      Left            =   5640
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   119
      Left            =   4560
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   118
      Left            =   4560
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   117
      Left            =   3480
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   116
      Left            =   3480
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   115
      Left            =   2400
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   114
      Left            =   1320
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   113
      Left            =   240
      Top             =   12000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   112
      Left            =   2400
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   111
      Left            =   1320
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   110
      Left            =   240
      Top             =   10920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   109
      Left            =   240
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   108
      Left            =   1320
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   107
      Left            =   2400
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   106
      Left            =   240
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   105
      Left            =   1320
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   104
      Left            =   2400
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   103
      Left            =   3480
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   102
      Left            =   3480
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   101
      Left            =   4560
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   100
      Left            =   4560
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   99
      Left            =   5640
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   98
      Left            =   5640
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   97
      Left            =   6720
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   96
      Left            =   6720
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   95
      Left            =   7800
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   94
      Left            =   7800
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   93
      Left            =   8880
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   92
      Left            =   8880
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   91
      Left            =   9960
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   90
      Left            =   9960
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   89
      Left            =   11040
      Top             =   8760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   88
      Left            =   11040
      Top             =   9840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   87
      Left            =   240
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   86
      Left            =   1320
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   85
      Left            =   2400
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   84
      Left            =   240
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   83
      Left            =   1320
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   82
      Left            =   2400
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   81
      Left            =   3480
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   80
      Left            =   3480
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   79
      Left            =   4560
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   78
      Left            =   4560
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   77
      Left            =   5640
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   76
      Left            =   5640
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   75
      Left            =   6720
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   74
      Left            =   6720
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   73
      Left            =   7800
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   72
      Left            =   7800
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   71
      Left            =   8880
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   70
      Left            =   8880
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   69
      Left            =   9960
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   68
      Left            =   9960
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   67
      Left            =   11040
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   66
      Left            =   11040
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   65
      Left            =   240
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   64
      Left            =   1320
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   63
      Left            =   2400
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   62
      Left            =   240
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   61
      Left            =   1320
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   60
      Left            =   2400
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   59
      Left            =   3480
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   58
      Left            =   3480
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   57
      Left            =   4560
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   56
      Left            =   4560
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   55
      Left            =   5640
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   54
      Left            =   5640
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   53
      Left            =   6720
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   52
      Left            =   6720
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   51
      Left            =   7800
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   50
      Left            =   7800
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   49
      Left            =   8880
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   48
      Left            =   8880
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   47
      Left            =   9960
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   46
      Left            =   9960
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   45
      Left            =   11040
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   44
      Left            =   11040
      Top             =   5520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   43
      Left            =   240
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   42
      Left            =   1320
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   41
      Left            =   2400
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   40
      Left            =   240
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   39
      Left            =   1320
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   38
      Left            =   2400
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   37
      Left            =   3480
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   36
      Left            =   3480
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   35
      Left            =   4560
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   34
      Left            =   4560
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   33
      Left            =   5640
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   32
      Left            =   5640
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   31
      Left            =   6720
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   30
      Left            =   6720
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   29
      Left            =   7800
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   28
      Left            =   7800
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   27
      Left            =   8880
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   26
      Left            =   8880
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   25
      Left            =   9960
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   24
      Left            =   9960
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   23
      Left            =   11040
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   22
      Left            =   11040
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   21
      Left            =   11040
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   20
      Left            =   11040
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   19
      Left            =   9960
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   18
      Left            =   9960
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   17
      Left            =   8880
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   16
      Left            =   8880
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   15
      Left            =   7800
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   14
      Left            =   7800
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   13
      Left            =   6720
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   12
      Left            =   6720
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   11
      Left            =   5640
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   10
      Left            =   5640
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   9
      Left            =   4560
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   8
      Left            =   4560
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   7
      Left            =   3480
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   6
      Left            =   3480
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   5
      Left            =   2400
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   4
      Left            =   1320
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   3
      Left            =   240
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   2
      Left            =   2400
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   1
      Left            =   1320
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRave_Click()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim Colour As Integer
Dim Count As Integer

For Count = 0 To 220

Randomize
Colour = Int(Rnd * 220) + 1

Select Case Colour
    Case Is = 1
        Shape1(Count).BackColor = &HC0C0FF
    Case Is = 2
        Shape1(Count).BackColor = &HFFFF80
    Case Is = 3
        Shape1(Count).BackColor = vbRed
    Case Is = 4
        Shape1(Count).BackColor = &HC0C000
    Case Is = 5
        Shape1(Count).BackColor = &H80
    Case Is = 6
        Shape1(Count).BackColor = vbRed
    Case Is = 7
        Shape1(Count).BackColor = vbBlue
    Case Is = 8
        Shape1(Count).BackColor = vbYellow
    Case Is = 9
        Shape1(Count).BackColor = vbGreen
    Case Is = 10
        Shape1(Count).BackColor = vbMagenta
End Select




Next
End Sub

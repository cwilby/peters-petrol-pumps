VERSION 5.00
Begin VB.Form frmTruthTable 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Peters Petrol Pumps"
   ClientHeight    =   9825
   ClientLeft      =   9000
   ClientTop       =   0
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   655
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReplaceHolster 
      Caption         =   "Replace Holster"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemoveHolster 
      Caption         =   "Remove Holster"
      Height          =   495
      Left            =   8520
      TabIndex        =   1
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Timer TimerPump 
      Interval        =   1
      Left            =   7320
      Top             =   9240
   End
   Begin VB.CommandButton cmdPump 
      BackColor       =   &H8000000B&
      Caption         =   "-------"
      Enabled         =   0   'False
      Height          =   555
      Left            =   8520
      MaskColor       =   &H0080FFFF&
      Picture         =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Timer PetrolRefresher 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   7800
      Top             =   9240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum of 0.5 litres per sale             See that pump is zeroed before use"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   1800
      TabIndex        =   3
      Top             =   8880
      Width           =   4455
   End
   Begin VB.Image State6 
      Height          =   540
      Left            =   1920
      Picture         =   "Form1.frx":2C055
      Top             =   8160
      Width           =   4155
   End
   Begin VB.Image State5 
      Height          =   540
      Left            =   1800
      Picture         =   "Form1.frx":30BA7
      Top             =   8160
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.Image State4 
      Height          =   540
      Left            =   1800
      Picture         =   "Form1.frx":366B7
      Top             =   8160
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.Image State3 
      Height          =   540
      Left            =   1920
      Picture         =   "Form1.frx":3BD33
      Top             =   8160
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.Image State2 
      Height          =   540
      Left            =   1800
      Picture         =   "Form1.frx":418BA
      Top             =   8160
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.Image State1 
      Height          =   540
      Left            =   1920
      Picture         =   "Form1.frx":45A9B
      Top             =   8160
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   5670
      Top             =   7170
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   6045
      Top             =   5100
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   6045
      Top             =   2580
      Width           =   135
   End
   Begin VB.Line pplFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   6
      Visible         =   0   'False
      X1              =   248
      X2              =   264
      Y1              =   448
      Y2              =   448
   End
   Begin VB.Line pplFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   5
      X1              =   240
      X2              =   240
      Y1              =   440
      Y2              =   424
   End
   Begin VB.Line pplFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   4
      X1              =   240
      X2              =   240
      Y1              =   472
      Y2              =   456
   End
   Begin VB.Line pplFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   3
      X1              =   248
      X2              =   264
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line pplFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   2
      X1              =   272
      X2              =   272
      Y1              =   456
      Y2              =   472
   End
   Begin VB.Line pplFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   1
      X1              =   272
      X2              =   272
      Y1              =   424
      Y2              =   440
   End
   Begin VB.Line pplFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   0
      X1              =   248
      X2              =   264
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line pplThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   6
      Visible         =   0   'False
      X1              =   312
      X2              =   296
      Y1              =   448
      Y2              =   448
   End
   Begin VB.Line pplThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   5
      X1              =   288
      X2              =   288
      Y1              =   424
      Y2              =   440
   End
   Begin VB.Line pplThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   4
      X1              =   288
      X2              =   288
      Y1              =   456
      Y2              =   472
   End
   Begin VB.Line pplThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   3
      X1              =   296
      X2              =   312
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line pplThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   2
      X1              =   320
      X2              =   320
      Y1              =   456
      Y2              =   472
   End
   Begin VB.Line pplThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   1
      X1              =   320
      X2              =   320
      Y1              =   424
      Y2              =   440
   End
   Begin VB.Line pplThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   0
      X1              =   296
      X2              =   312
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line pplSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   6
      Visible         =   0   'False
      X1              =   344
      X2              =   360
      Y1              =   448
      Y2              =   448
   End
   Begin VB.Line pplSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   5
      X1              =   336
      X2              =   336
      Y1              =   440
      Y2              =   424
   End
   Begin VB.Line pplSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   4
      X1              =   336
      X2              =   336
      Y1              =   472
      Y2              =   456
   End
   Begin VB.Line pplSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   3
      X1              =   344
      X2              =   360
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line pplSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   2
      X1              =   368
      X2              =   368
      Y1              =   456
      Y2              =   472
   End
   Begin VB.Line pplSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   1
      X1              =   368
      X2              =   368
      Y1              =   424
      Y2              =   440
   End
   Begin VB.Line pplSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   0
      X1              =   344
      X2              =   360
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line pplFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   6
      Visible         =   0   'False
      X1              =   408
      X2              =   424
      Y1              =   448
      Y2              =   448
   End
   Begin VB.Line pplFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   5
      X1              =   400
      X2              =   400
      Y1              =   440
      Y2              =   424
   End
   Begin VB.Line pplFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   4
      X1              =   400
      X2              =   400
      Y1              =   472
      Y2              =   456
   End
   Begin VB.Line pplFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   3
      X1              =   408
      X2              =   424
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line pplFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   2
      X1              =   432
      X2              =   432
      Y1              =   456
      Y2              =   472
   End
   Begin VB.Line pplFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   1
      X1              =   432
      X2              =   432
      Y1              =   424
      Y2              =   440
   End
   Begin VB.Line pplFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Index           =   0
      X1              =   408
      X2              =   424
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line PriceFifth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   6
      Visible         =   0   'False
      X1              =   248
      X2              =   272
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line PriceFifth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   5
      X1              =   240
      X2              =   240
      Y1              =   128
      Y2              =   104
   End
   Begin VB.Line PriceFifth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   4
      X1              =   240
      X2              =   240
      Y1              =   168
      Y2              =   144
   End
   Begin VB.Line PriceFifth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   3
      X1              =   248
      X2              =   272
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Line PriceFifth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   2
      X1              =   280
      X2              =   280
      Y1              =   144
      Y2              =   168
   End
   Begin VB.Line PriceFifth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   1
      X1              =   280
      X2              =   280
      Y1              =   104
      Y2              =   128
   End
   Begin VB.Line PriceFifth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   0
      X1              =   248
      X2              =   272
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line PriceFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   6
      Visible         =   0   'False
      X1              =   304
      X2              =   328
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line PriceFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   5
      X1              =   296
      X2              =   296
      Y1              =   128
      Y2              =   104
   End
   Begin VB.Line PriceFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   4
      X1              =   296
      X2              =   296
      Y1              =   168
      Y2              =   144
   End
   Begin VB.Line PriceFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   3
      X1              =   304
      X2              =   328
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Line PriceFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   2
      X1              =   336
      X2              =   336
      Y1              =   144
      Y2              =   168
   End
   Begin VB.Line PriceFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   1
      X1              =   336
      X2              =   336
      Y1              =   104
      Y2              =   128
   End
   Begin VB.Line PriceFourth 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   0
      X1              =   304
      X2              =   328
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line FourthSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   6
      Visible         =   0   'False
      X1              =   304
      X2              =   328
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line FourthSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   5
      X1              =   296
      X2              =   296
      Y1              =   296
      Y2              =   272
   End
   Begin VB.Line FourthSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   4
      X1              =   296
      X2              =   296
      Y1              =   336
      Y2              =   312
   End
   Begin VB.Line FourthSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   3
      X1              =   304
      X2              =   328
      Y1              =   344
      Y2              =   344
   End
   Begin VB.Line FourthSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   2
      X1              =   336
      X2              =   336
      Y1              =   312
      Y2              =   336
   End
   Begin VB.Line FourthSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   1
      X1              =   336
      X2              =   336
      Y1              =   272
      Y2              =   296
   End
   Begin VB.Line FourthSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   0
      X1              =   304
      X2              =   328
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Line PriceFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   6
      Visible         =   0   'False
      X1              =   488
      X2              =   512
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line PriceFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   5
      X1              =   480
      X2              =   480
      Y1              =   104
      Y2              =   128
   End
   Begin VB.Line PriceFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   4
      X1              =   480
      X2              =   480
      Y1              =   144
      Y2              =   168
   End
   Begin VB.Line PriceFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   3
      X1              =   488
      X2              =   512
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Line PriceFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   2
      X1              =   520
      X2              =   520
      Y1              =   144
      Y2              =   168
   End
   Begin VB.Line PriceFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   1
      X1              =   520
      X2              =   520
      Y1              =   104
      Y2              =   128
   End
   Begin VB.Line PriceFirst 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   0
      X1              =   488
      X2              =   512
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line PriceSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   6
      Visible         =   0   'False
      X1              =   432
      X2              =   456
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line PriceSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   5
      X1              =   424
      X2              =   424
      Y1              =   128
      Y2              =   104
   End
   Begin VB.Line PriceSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   4
      X1              =   424
      X2              =   424
      Y1              =   168
      Y2              =   144
   End
   Begin VB.Line PriceSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   3
      X1              =   432
      X2              =   456
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Line PriceSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   2
      X1              =   464
      X2              =   464
      Y1              =   144
      Y2              =   168
   End
   Begin VB.Line PriceSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   1
      X1              =   464
      X2              =   464
      Y1              =   104
      Y2              =   128
   End
   Begin VB.Line PriceSecond 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   0
      X1              =   432
      X2              =   456
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line PriceThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   6
      Visible         =   0   'False
      X1              =   360
      X2              =   384
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line PriceThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   5
      X1              =   352
      X2              =   352
      Y1              =   104
      Y2              =   128
   End
   Begin VB.Line PriceThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   4
      X1              =   352
      X2              =   352
      Y1              =   144
      Y2              =   168
   End
   Begin VB.Line PriceThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   3
      X1              =   360
      X2              =   384
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Line PriceThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   2
      X1              =   392
      X2              =   392
      Y1              =   144
      Y2              =   168
   End
   Begin VB.Line PriceThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   1
      X1              =   392
      X2              =   392
      Y1              =   104
      Y2              =   128
   End
   Begin VB.Line PriceThird 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   0
      X1              =   360
      X2              =   384
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line ThirdSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   6
      Visible         =   0   'False
      X1              =   360
      X2              =   384
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line ThirdSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   5
      X1              =   352
      X2              =   352
      Y1              =   296
      Y2              =   272
   End
   Begin VB.Line ThirdSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   4
      X1              =   352
      X2              =   352
      Y1              =   336
      Y2              =   312
   End
   Begin VB.Line ThirdSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   3
      X1              =   360
      X2              =   384
      Y1              =   344
      Y2              =   344
   End
   Begin VB.Line ThirdSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   2
      X1              =   392
      X2              =   392
      Y1              =   312
      Y2              =   336
   End
   Begin VB.Line ThirdSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   1
      X1              =   392
      X2              =   392
      Y1              =   272
      Y2              =   296
   End
   Begin VB.Line ThirdSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   0
      X1              =   360
      X2              =   384
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Line SecondSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   6
      Visible         =   0   'False
      X1              =   456
      X2              =   432
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line SecondSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   5
      X1              =   424
      X2              =   424
      Y1              =   272
      Y2              =   296
   End
   Begin VB.Line SecondSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   4
      X1              =   424
      X2              =   424
      Y1              =   312
      Y2              =   336
   End
   Begin VB.Line SecondSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   3
      X1              =   456
      X2              =   432
      Y1              =   344
      Y2              =   344
   End
   Begin VB.Line SecondSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   2
      X1              =   464
      X2              =   464
      Y1              =   312
      Y2              =   336
   End
   Begin VB.Line SecondSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   1
      X1              =   464
      X2              =   464
      Y1              =   296
      Y2              =   272
   End
   Begin VB.Line SecondSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   0
      X1              =   456
      X2              =   432
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Line FirstSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   6
      Visible         =   0   'False
      X1              =   512
      X2              =   488
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line FirstSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   5
      X1              =   480
      X2              =   480
      Y1              =   272
      Y2              =   296
   End
   Begin VB.Line FirstSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   4
      X1              =   480
      X2              =   480
      Y1              =   312
      Y2              =   336
   End
   Begin VB.Line FirstSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   3
      X1              =   488
      X2              =   512
      Y1              =   344
      Y2              =   344
   End
   Begin VB.Line FirstSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   2
      X1              =   520
      X2              =   520
      Y1              =   312
      Y2              =   336
   End
   Begin VB.Line FirstSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   1
      X1              =   520
      X2              =   520
      Y1              =   272
      Y2              =   296
   End
   Begin VB.Line FirstSeg 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      Index           =   0
      X1              =   512
      X2              =   488
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Image Image1 
      Height          =   9825
      Left            =   -240
      Picture         =   "Form1.frx":4AB95
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "frmTruthTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare Initial Variables
    Option Explicit
    Dim litres As Integer
    Dim price As Double
    Dim TruthTable(9, 6) As Boolean
    Dim TempVal As Integer
    Dim pri As Double
    Dim pricetwo As Double
    Dim CountSafe As Integer

Private Sub cmdRemoveHolster_Click()
    
    cmdPump.Caption = "Hold Down to Start Pump"
    SetTF (True)
    cmdRemoveHolster.Enabled = False
    cmdReplaceHolster.Enabled = True
    SetState (2)
        State1.Visible = False
        State2.Visible = True
        State3.Visible = False
        State4.Visible = False
        State5.Visible = False
        State6.Visible = False

End Sub

Private Sub cmdReplaceHolster_Click()
If GetLitres() = 0 Then
    SetTF (False)
    cmdReplaceHolster.Enabled = False
    cmdRemoveHolster.Enabled = True
    SetState (1)
        State1.Visible = True
        State2.Visible = False
        State3.Visible = False
        State4.Visible = False
        State5.Visible = False
        State6.Visible = False

    
Else
    If GetLitres() <= 50 Then
        SetState (4)
            State1.Visible = False
            State2.Visible = False
            State3.Visible = False
            State4.Visible = True
            State5.Visible = False
            State6.Visible = False
        cmdRemoveHolster.Enabled = False
        cmdReplaceHolster.Enabled = True
        SetTF (True)
    Else
        cmdRemoveHolster.Enabled = False
        cmdReplaceHolster.Enabled = False
        SetTF (False)
        SetState (5)
            State1.Visible = False
            State2.Visible = False
            State3.Visible = False
            State4.Visible = False
            State5.Visible = True
            State6.Visible = False
    End If
End If
End Sub

Private Sub Form_Load()
'Sets state of pump
    SetState (1)
    State1.Visible = True
    State2.Visible = False
    State3.Visible = False
    State4.Visible = False
    State5.Visible = False
    State6.Visible = False

'Sets availability to use pump
    SetTF (False)
    litres = 0
    price = 0
    
'Declare variables
    Dim FirstPpl As Integer
    Dim SecondPpl As Integer
    Dim ThirdPpl As Integer
    Dim FourthPpl As Integer
    Dim CountSafe As Integer
    Dim ppl As Double
    
'Finds price per litre
    ppl = GetPrice() * 10
        
'Deconcatenates Price per Litre
    FirstPpl = Int(ppl / 10 ^ 0) Mod 10
    SecondPpl = Int(ppl / 10 ^ 1) Mod 10
    ThirdPpl = Int(ppl / 10 ^ 2) Mod 10
    FourthPpl = Int(ppl / 10 ^ 3) Mod 10
    
'Call Truth Table procedure
    Call Truth_Table
    
'Change 7 seven segment display to show PPL
    Do Until CountSafe = 7
    pplFirst(CountSafe).Visible = TruthTable(FirstPpl, CountSafe)
    pplSecond(CountSafe).Visible = TruthTable(SecondPpl, CountSafe)
    pplThird(CountSafe).Visible = TruthTable(ThirdPpl, CountSafe)
    pplFourth(CountSafe).Visible = TruthTable(FourthPpl, CountSafe)
  
    CountSafe = CountSafe + 1
    Loop

End Sub

Private Sub Truth_Table()
'Truth Table
    TruthTable(0, 0) = True
    TruthTable(0, 1) = True
    TruthTable(0, 2) = True
    TruthTable(0, 3) = True
    TruthTable(0, 4) = True
    TruthTable(0, 5) = True
    TruthTable(0, 6) = False

'1
    TruthTable(1, 0) = False
    TruthTable(1, 1) = True
    TruthTable(1, 2) = True
    TruthTable(1, 3) = False
    TruthTable(1, 4) = False
    TruthTable(1, 5) = False
    TruthTable(1, 6) = False

'2
    TruthTable(2, 0) = True
    TruthTable(2, 1) = True
    TruthTable(2, 2) = False
    TruthTable(2, 3) = True
    TruthTable(2, 4) = True
    TruthTable(2, 5) = False
    TruthTable(2, 6) = True

'3
    TruthTable(3, 0) = True
    TruthTable(3, 1) = True
    TruthTable(3, 2) = True
    TruthTable(3, 3) = True
    TruthTable(3, 4) = False
    TruthTable(3, 5) = False
    TruthTable(3, 6) = True

'4
    TruthTable(4, 0) = False
    TruthTable(4, 1) = True
    TruthTable(4, 2) = True
    TruthTable(4, 3) = False
    TruthTable(4, 4) = False
    TruthTable(4, 5) = True
    TruthTable(4, 6) = True

'5
    TruthTable(5, 0) = True
    TruthTable(5, 1) = False
    TruthTable(5, 2) = True
    TruthTable(5, 3) = True
    TruthTable(5, 4) = False
    TruthTable(5, 5) = True
    TruthTable(5, 6) = True

'6
    TruthTable(6, 0) = True
    TruthTable(6, 1) = False
    TruthTable(6, 2) = True
    TruthTable(6, 3) = True
    TruthTable(6, 4) = True
    TruthTable(6, 5) = True
    TruthTable(6, 6) = True

'7
    TruthTable(7, 0) = True
    TruthTable(7, 1) = True
    TruthTable(7, 2) = True
    TruthTable(7, 3) = False
    TruthTable(7, 4) = False
    TruthTable(7, 5) = False
    TruthTable(7, 6) = False

'8
    TruthTable(8, 0) = True
    TruthTable(8, 1) = True
    TruthTable(8, 2) = True
    TruthTable(8, 3) = True
    TruthTable(8, 4) = True
    TruthTable(8, 5) = True
    TruthTable(8, 6) = True

'9
    TruthTable(9, 0) = True
    TruthTable(9, 1) = True
    TruthTable(9, 2) = True
    TruthTable(9, 3) = False
    TruthTable(9, 4) = False
    TruthTable(9, 5) = True
    TruthTable(9, 6) = True

End Sub

Private Sub cmdPump_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'When mouse is down, Start pump
    cmdPump.Caption = "Release To Stop Pump"
    cmdPump.BackColor = &H8000000A
    PetrolRefresher.Enabled = True

End Sub


Private Sub cmdPump_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'When mouse is lifted, Stop pump
    cmdPump.Caption = "Hold Down To Start Pump"
    cmdPump.BackColor = &H8000000B
    PetrolRefresher.Enabled = False

End Sub

Sub cmdReset_Click()

'Disables Timer
    PetrolRefresher.Enabled = False

'Resets all numbers to 0
    CountSafe = 0
    Do Until CountSafe = 7
    FirstSeg(CountSafe).Visible = TruthTable(0, CountSafe)
    SecondSeg(CountSafe).Visible = TruthTable(0, CountSafe)
    ThirdSeg(CountSafe).Visible = TruthTable(0, CountSafe)
    FourthSeg(CountSafe).Visible = TruthTable(0, CountSafe)
    
    PriceFirst(CountSafe).Visible = TruthTable(0, CountSafe)
    PriceSecond(CountSafe).Visible = TruthTable(0, CountSafe)
    PriceThird(CountSafe).Visible = TruthTable(0, CountSafe)
    PriceFourth(CountSafe).Visible = TruthTable(0, CountSafe)
    
    CountSafe = CountSafe + 1
    Loop
    
    'Resets Values
    litres = 0
    SetLitres (0)
    price = 0

End Sub

Sub PetrolRefresher_Timer()
'LITRES
    'Increment Litres
        If litres < 9999 Then
            litres = litres + 1
            SetLitres (litres)
        Else
            PetrolRefresher.Enabled = False
                SetState (3)
                State1.Visible = False
                State2.Visible = False
                State3.Visible = True
                State4.Visible = False
                State5.Visible = False
                State6.Visible = False
        End If
    
    'Declare Variables
        Dim FirstLitres As Integer
        Dim SecondLitres As Integer
        Dim ThirdLitres As Integer
        Dim FourthLitres As Integer
        Dim FirstPrice As Integer
        Dim SecondPrice As Integer
        Dim ThirdPrice As Integer
        Dim FourthPrice As Integer
        Dim FifthPrice As Integer
        Dim price As Double
        Dim CountSafe As Integer
    
    'Split Litres
        FirstLitres = Int(litres / 10 ^ 0) Mod 10
        SecondLitres = Int(litres / 10 ^ 1) Mod 10
        ThirdLitres = Int(litres / 10 ^ 2) Mod 10
        FourthLitres = Int(litres / 10 ^ 3) Mod 10

'PRICE
    'Declare Variables
        price = litres * GetPrice() / 100
    
    'Split Price
        FirstPrice = Int(price / 10 ^ 0) Mod 10
        SecondPrice = Int(price / 10 ^ 1) Mod 10
        ThirdPrice = Int(price / 10 ^ 2) Mod 10
        FourthPrice = Int(price / 10 ^ 3) Mod 10
        FifthPrice = Int(price / 10 ^ 4) Mod 10
        
        Dim PriceConc As String
            PriceConc = FifthPrice & FourthPrice & ThirdPrice & SecondPrice & FirstPrice
        SetTaken (PriceConc)

'PRICE PER LITRE
    'Call Truth Table procedure
        Call Truth_Table
        
    'Adjusts Seven Segments to display number
        CountSafe = 0
        Do Until CountSafe = 7
        
    'LITRES
        FirstSeg(CountSafe).Visible = TruthTable(FirstLitres, CountSafe)
        SecondSeg(CountSafe).Visible = TruthTable(SecondLitres, CountSafe)
        ThirdSeg(CountSafe).Visible = TruthTable(ThirdLitres, CountSafe)
        FourthSeg(CountSafe).Visible = TruthTable(FourthLitres, CountSafe)
        
    'PRICE
        PriceFirst(CountSafe).Visible = TruthTable(FirstPrice, CountSafe)
        PriceSecond(CountSafe).Visible = TruthTable(SecondPrice, CountSafe)
        PriceThird(CountSafe).Visible = TruthTable(ThirdPrice, CountSafe)
        PriceFourth(CountSafe).Visible = TruthTable(FourthPrice, CountSafe)
        PriceFifth(CountSafe).Visible = TruthTable(FifthPrice, CountSafe)
        
        CountSafe = CountSafe + 1
        Loop

End Sub

Private Sub TimerPump_Timer()
    cmdPump.Enabled = GetTF()
End Sub

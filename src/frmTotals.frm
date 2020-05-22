VERSION 5.00
Begin VB.Form frmTotals 
   Caption         =   "Totals"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   2445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblAmountTaken 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*amount*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Amount Taken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblLitresSold 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*litres*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Litres Sold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload frmTotals
End Sub

Private Sub Form_Load()
lblLitresSold.Caption = GetTotalLitres() / 100
lblAmountTaken.Caption = FormatCurrency(GetTotalPrice())
End Sub

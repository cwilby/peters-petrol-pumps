VERSION 5.00
Begin VB.Form frmConsole2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Peter's Console"
   ClientHeight    =   11355
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   16575
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11355
   ScaleWidth      =   16575
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameChangePrice 
      Caption         =   "Change Price of Petrol"
      Height          =   3255
      Left            =   9480
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton cmdSubmitChanges 
         Caption         =   "Submit Changes"
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox lblNewPrice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Enter New Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblCurrentPrice 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*price*"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Current Price of Petrol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdChangePrice 
      Caption         =   "Change Price of Petrol"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill Tanks"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmdEndShift 
      Caption         =   "End shift"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton cmdTotals 
      Caption         =   "Totals for the day"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdShutdown 
      Caption         =   "Shutdown"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton cmdMainScreen 
      Caption         =   "Main Screen"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Frame framePay 
      Caption         =   "Pay"
      Height          =   3015
      Left            =   3840
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmdPay 
         Caption         =   "Collect Payment"
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label lblTotal 
         Caption         =   "*total*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblPpl 
         Caption         =   "*priceperlitre*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblLitres 
         Caption         =   "*litres*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblPumpNumber 
         Caption         =   "*pumpnumber*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Total: £"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Price per litre:"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Amount of Litres:"
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Pump to be paid for:"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "STOP ALL PUMPS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11880
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer TimerConsole 
      Interval        =   500
      Left            =   12360
      Top             =   8040
   End
   Begin VB.CommandButton PumpOne 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblTime 
      Caption         =   "Time not available"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label lblCopyright 
      Caption         =   $"frmConsole.frx":0000
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   8520
      Width           =   12855
   End
   Begin VB.Label lblLitresOne 
      BackStyle       =   0  'Transparent
      Caption         =   "00.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "00.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   12840
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frmConsole2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdControlPanel_Click()
frmCp.Visible = True
End Sub

Private Sub cmdChangePrice_Click()
frameChangePrice.Visible = True
lblCurrentPrice.Caption = GetPrice()
End Sub

Private Sub cmdNewPriceGo_Click()

If IsNumeric(lblNewPrice.Text) = False Then
    MsgBox ("Please only enter numbers")
Else
    If lblNewPrice.Text = "" Then
        MsgBox ("Please enter data")
    Else
        SetPrice (lblNewPrice.Text)
        MsgBox ("Price successfully changed")
        frameChangePrice.Visible = False
        Unload frmTruthTable
        frmTruthTable.Visible = True
    End If
End If

End Sub

Private Sub cmdPay_Click()
If GetLitres() = 0 Then
    MsgBox ("Cannot pay if nothing is taken!")
Else
    MsgBox ("Thank customer for paying, Transaction is complete.")
    SetTotalLitres (GetTotalLitres() + GetLitres())
    SetTotalPrice (GetTotalPrice() + FormatCurrency(GetLitres() * GetPrice() / 10000))
    SetLitres (0)
    Unload frmTruthTable
    frmTruthTable.Visible = True
    framePay.Visible = False
End If
End Sub

Private Sub cmdShutDown_Click()
End
End Sub


Private Sub cmdTotals_Click()
frmTotals.Visible = True
End Sub



Private Sub Form_Load()

Dim PriceInput As String
PriceInput = InputBox("System booted, Please enter price of petrol", "System Startup")
If PriceInput = "" Then
    End
Else
    SetPrice (Val(PriceInput))
    frmTruthTable.Visible = True
End If

End Sub

Private Sub PumpOne_Click()

If GetLitres() = 0 Then
    MsgBox ("Cannot pay if nothing is taken!")
Else
    framePay.Visible = True
    lblPumpNumber.Caption = "1"
    lblLitres.Caption = FormatNumber(GetLitres() / 100)
    lblPpl.Caption = "£" & GetPrice() / 100
    lblTotal.Caption = FormatCurrency(GetLitres() * GetPrice() / 10000)
End If

End Sub

Private Sub TimerConsole_Timer()
Dim LitresDisplay As Currency
LitresDisplay = GetLitres() / 100
lblLitresOne.Caption = FormatNumber(LitresDisplay)

If GetLitres() > 0 Then
    PumpOne.BackColor = &HFFFF&
    PumpOne.Caption = "In Use Click to Pay"
Else
    PumpOne.BackColor = &HFF00&
    PumpOne.Caption = "1"
End If
    
'Time
lblTime.Caption = Time()
    
End Sub

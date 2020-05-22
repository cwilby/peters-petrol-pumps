VERSION 5.00
Begin VB.Form frmConsole 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "About"
      Height          =   375
      Left            =   7800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdZeroTotals 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Zero Totals"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2640
      Width           =   1750
   End
   Begin VB.Frame frameTotals 
      BackColor       =   &H8000000D&
      Height          =   1095
      Left            =   2880
      TabIndex        =   24
      Top             =   3960
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdXtotals 
         Caption         =   "X"
         Height          =   255
         Left            =   5040
         TabIndex        =   29
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblAmountTaken 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*amount*"
         Height          =   375
         Left            =   2160
         TabIndex        =   28
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount taken"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblLitresSold 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*litres*"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres Sold"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame framePay 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   2655
      Left            =   4920
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdPay 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Collect Payment"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "*total*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblPpl 
         BackStyle       =   0  'Transparent
         Caption         =   "*priceperlitre*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblLitres 
         BackStyle       =   0  'Transparent
         Caption         =   "*litres*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblPumpNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "*pumpnumber*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Price per litre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount of Litres:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Pump to be paid for:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame frameChangePrice 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton cmdXprice 
         Caption         =   "X"
         Height          =   255
         Left            =   1680
         TabIndex        =   30
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdSubmitChanges 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Submit Changes"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox lblNewPrice 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter new price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblCurrentPrice 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*price*"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Price Of Petrol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Timer TimerConsole 
      Interval        =   250
      Left            =   8400
      Top             =   4920
   End
   Begin VB.CommandButton cmdShutDown 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Shutdown"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1750
   End
   Begin VB.CommandButton cmdChangePrice 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Change price of Petrol"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1750
   End
   Begin VB.CommandButton cmdTotals 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Totals for the Day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1750
   End
   Begin VB.CommandButton cmdMainScreen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Main Screen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1750
   End
   Begin VB.CommandButton PumpOne 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Loading ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading ..."
      Height          =   255
      Left            =   7800
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Written and designed by Cameron Wilby"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5700
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Peter's Petrol Pumps v1.21"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "frmConsoleGui.frx":0000
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChangePrice_Click()
'Opens change price frame
    frameChangePrice.Visible = True
    lblCurrentPrice.Caption = GetPrice()
End Sub

Private Sub cmdMainScreen_Click()
'Closes all frames
    frameChangePrice.Visible = False
    framePay.Visible = False
    frameTotals.Visible = False
End Sub


Private Sub cmdPay_Click()
'Commands for when Pay button is clicked
    If GetLitres() = 0 Then
        MsgBox "Cannot pay if nothing is taken!", , "Notice"
    Else
        MsgBox "Thank customer for paying, Transaction is complete.", , "Notice"
        SetTotalLitres (GetTotalLitres() + GetLitres())
        SetTotalPrice (GetTotalPrice() + GetTaken() / 100)
        SetLitres (0)
        SetTaken (0)
        Unload frmTruthTable
        frmTruthTable.Visible = True
        framePay.Visible = False
    End If
End Sub

Private Sub cmdShutDown_Click()
'Shuts down program
    End
End Sub


Private Sub cmdTotals_Click()
'Shows totals
    frameTotals.Visible = True
    lblLitresSold.Caption = GetTotalLitres() / 100
    lblAmountTaken.Caption = FormatCurrency(GetTotalPrice())
End Sub

Private Sub cmdXprice_Click()
'Closes change price frame
    frameChangePrice.Visible = False
End Sub

Private Sub cmdXtotals_Click()
'Closes totals frame
    frameTotals.Visible = False
End Sub

Private Sub cmdZeroTotals_Click()
'Zeros everything for beginning a new day
    Notice = MsgBox("This will end one shift and begin another, Proceed?", VbMsgBoxStyle.vbYesNo, "Notice")
    If Notice = 6 Then
        SetTotalLitres (0)
        SetTotalPrice (0)
        lblLitresSold.Caption = "0"
        lblAmountTaken.Caption = "£0.00"
    End If
End Sub

Private Sub Command1_Click()
MsgBox "Written and designed by Cameron Wilby for AS Computing Project 2007-2008, V1.2", , "About Program"
End Sub

Private Sub Form_Load()
'System startup procedures
Dim PriceInput As String
PriceInput = InputBox("System booted, Please enter price of petrol (Eg £1.019 is 101.9)", "System Startup")
    If PriceInput = "" Then
        End
    Else
        If IsNumeric(PriceInput) = False Then
           MsgBox "Price must be a numerical value", , "Notice"
           Unload frmConsole
           frmConsole.Visible = True
        Else
            If PriceInput > 199.9 Then
                MsgBox "Price cannot be above £1.99", , "Notice"
                Unload frmConsole
                frmConsole.Visible = True
            Else
                If PriceInput < 0.1 Then
                    MsgBox "Price cannot be below 1p", , "Notice"
                    Unload frmConsole
                    frmConsole.Visible = True
                Else
                    SetPrice (Val(PriceInput))
                    frmTruthTable.Visible = True
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdSubmitChanges_Click()
'Submits changes, and validates entry
If GetState = 1 Then
    If IsNumeric(lblNewPrice.Text) = False Then
        MsgBox "Please only enter numbers", , "Notice"
    Else
        If lblNewPrice.Text = "" Then
            MsgBox ("Please enter data")
        Else
            If lblNewPrice.Text > 199.9 Then
                MsgBox "Price cannot be higher than £1.999 a litre!", , "Notice"
            Else
                If lblNewPrice.Text < 0.1 Then
                    MsgBox "Price cannot be lower than 1p a litre!", , "Notice"
                Else
                    SetPrice (lblNewPrice.Text)
                    MsgBox "Price successfully changed", , "Notice"
                    frameChangePrice.Visible = False
                    Unload frmTruthTable
                    frmTruthTable.Visible = True
                End If
            End If
        End If
    End If
Else
    MsgBox "Cannot change price in current state", , "Notice"
End If
End Sub


Private Sub PumpOne_Click()
'When pump button is clicked, open pay frame, also validates
    If GetLitres() = 0 Then
        MsgBox "Cannot pay if nothing is taken!", , "Notice"
    Else
        If GetLitres() < 50 Then
            MsgBox "Minimum of 0.5 litres must be dispensed!", , "Notice"
        Else
            If GetState() = 2 Then
                MsgBox "Cannot pay while pump is in use!", , "Notice"
            Else
                framePay.Visible = True
                SetTF (False)
                lblPumpNumber.Caption = "1"
                lblLitres.Caption = FormatNumber(GetLitres() / 100)
                lblPpl.Caption = "£" & GetPrice() / 100
                lblTotal.Caption = FormatCurrency(GetTaken() / 100)
            End If
        End If
    End If
End Sub

Private Sub TimerConsole_Timer()

'Every second, update the time, and status of pay button
    lblLitresSold.Caption = GetTotalLitres() / 100
    lblAmountTaken.Caption = FormatCurrency(GetTotalPrice())
        
    Dim LitresDisplay As Currency
    LitresDisplay = GetLitres() / 100
    
    If GetState() = 1 Then
        PumpOne.BackColor = &HFF00&
        PumpOne.Caption = "1"
    End If
    
    If GetState() = 2 Then
        PumpOne.BackColor = &HFFFF&
        PumpOne.Caption = "In Use  " & "L:" & FormatNumber(GetLitres() / 100) & " P:" & FormatCurrency(GetTaken() / 100)
    End If
    
    If GetState() = 3 Then
        PumpOne.BackColor = &HFF&
        PumpOne.Caption = "Click to Pay"
    End If
    
    If GetState() = 5 Then
        PumpOne.BackColor = &HFFF&
        PumpOne.Caption = "Click to Pay"
    End If
    
    If GetState() = 6 Then
        PumpOne.BackColor = &HFFF&
        PumpOne.Caption = "Click to Zero"
    End If
        
    
        
'Time
    lblTime.Caption = "Time: " & Time()
    
End Sub


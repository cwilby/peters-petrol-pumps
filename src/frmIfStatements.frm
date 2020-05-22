VERSION 5.00
Begin VB.Form frmIfStatements 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Line FirstDigit 
      BorderWidth     =   3
      Index           =   6
      X1              =   1560
      X2              =   1560
      Y1              =   1320
      Y2              =   1920
   End
   Begin VB.Line FirstDigit 
      BorderWidth     =   3
      Index           =   5
      X1              =   1560
      X2              =   1560
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Line FirstDigit 
      BorderWidth     =   3
      Index           =   4
      X1              =   2400
      X2              =   2400
      Y1              =   1320
      Y2              =   1920
   End
   Begin VB.Line FirstDigit 
      BorderWidth     =   3
      Index           =   3
      X1              =   2400
      X2              =   2400
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Line FirstDigit 
      BorderWidth     =   3
      Index           =   2
      X1              =   1560
      X2              =   2400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line FirstDigit 
      BorderWidth     =   3
      Index           =   1
      X1              =   1560
      X2              =   2400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line FirstDigit 
      BorderWidth     =   3
      Index           =   0
      X1              =   1560
      X2              =   2400
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line SecondDigit 
      BorderWidth     =   3
      Index           =   6
      X1              =   480
      X2              =   480
      Y1              =   1320
      Y2              =   1920
   End
   Begin VB.Line SecondDigit 
      BorderWidth     =   3
      Index           =   5
      X1              =   480
      X2              =   480
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Line SecondDigit 
      BorderWidth     =   3
      Index           =   4
      X1              =   1320
      X2              =   1320
      Y1              =   1320
      Y2              =   1920
   End
   Begin VB.Line SecondDigit 
      BorderWidth     =   3
      Index           =   3
      X1              =   1320
      X2              =   1320
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Line SecondDigit 
      BorderWidth     =   3
      Index           =   2
      X1              =   480
      X2              =   1320
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line SecondDigit 
      BorderWidth     =   3
      Index           =   1
      X1              =   480
      X2              =   1320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line SecondDigit 
      BorderWidth     =   3
      Index           =   0
      X1              =   480
      X2              =   1320
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmIfStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
'Variables
Dim Right As Integer
Dim Left As Integer
Dim TempVal As Integer

'Deconcatenate number
Right = TempVal Mod 10
Left = TempVal / 10

'Set Visibilities
'0
If Right = 0 Then
  FirstDigit(0).Visible = True
  FirstDigit(1).Visible = False
  FirstDigit(2).Visible = True
  FirstDigit(3).Visible = True
  FirstDigit(4).Visible = True
  FirstDigit(5).Visible = True
  FirstDigit(6).Visible = True
End If


'1
If Right = 1 Then
  FirstDigit(0).Visible = False
  FirstDigit(1).Visible = False
  FirstDigit(2).Visible = False
  FirstDigit(3).Visible = True
  FirstDigit(4).Visible = True
  FirstDigit(5).Visible = False
  FirstDigit(6).Visible = False
End If


'2
If Right = 2 Then
  FirstDigit(0).Visible = True
  FirstDigit(1).Visible = True
  FirstDigit(2).Visible = True
  FirstDigit(3).Visible = True
  FirstDigit(4).Visible = False
  FirstDigit(5).Visible = False
  FirstDigit(6).Visible = True
End If

'3

End Sub

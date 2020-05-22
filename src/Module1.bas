Attribute VB_Name = "Module1"
Dim price As Double
Dim litres As Integer
Dim Total As Double
Dim TotalPrice As Double
Dim TrueFalse As String
Dim State As Integer
Dim Taken As Double


Function SetState(sta As Double)
    State = sta
End Function

Function GetState()
    GetState = State
End Function

Function SetPrice(pri As Double)
    price = pri
End Function

Function GetPrice() As Double
    GetPrice = price
End Function

Function SetLitres(Lit As Double)
    litres = Lit
End Function

Function GetLitres()
    GetLitres = litres
End Function

Function SetTaken(Tak As Double)
    Taken = Tak
End Function

Function GetTaken()
    GetTaken = Taken
End Function
'Functions to record Total Litres
Function SetTotalLitres(Tot As Double)
    Total = Tot
End Function

Function GetTotalLitres() As Double
    GetTotalLitres = Total
End Function

'Functions to record Total Price
Function SetTotalPrice(Totp As Double)
    TotalPrice = Totp
End Function

Function GetTotalPrice() As Double
    GetTotalPrice = TotalPrice
End Function

Function SetTF(Tf As String)
    TrueFalse = Tf
End Function

Function GetTF()
    GetTF = TrueFalse
End Function

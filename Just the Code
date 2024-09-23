Option Explicit

Sub CalcTotal()

'label variables
Dim Q As Integer 'quantity
Dim P As Currency 'price
Dim EP As Currency 'extended price (Q*P)
Dim M As String 'membership
Dim CBP As Single 'cashback%
Dim CBA As Currency 'cashback amount
Dim TACB As Currency 'total after cashback
Dim MPCB As Currency 'minimum purchase for cashback
Dim ET As Currency 'encouragement threshold

'input Q, P, M, MPCB, ET
Q = Range("B3").Value
P = Range("B4").Value
M = Range("B7").Value
MPCB = Range("B22").Value
ET = Range("B24").Value

'calculate EP
EP = Q * P

'determine CBP
If M = Range("A16").Value Then
    CBP = Range("B16").Value
  ElseIf M = Range("A17").Value Then
    CBP = Range("B17").Value
  ElseIf M = Range("A18").Value Then
    CBP = Range("B18").Value
  ElseIf M = Range("A19").Value Then
    CBP = Range("B19").Value
End If

'calculate CBA
CBA = EP * CBP

  'CBA only if EP >= MPCB
  If EP >= MPCB Then
      CBA = CBA
    ElseIf EP < MPCB Then
      CBA = 0
  End If

'calculate TACB
TACB = EP - CBA

'output EP, CBP, CBA, TACB
Range("B5").Value = EP
Range("B8").Value = CBP
Range("B10").Value = CBA
Range("B11").Value = TACB

'CBA >= ET then MsgBox "make the purchase"
If CBA >= ET Then
  MsgBox ("Shoulder Devil: Make the purchase!")
End If

'answer
Range("B11").Activate
Beep

End Sub

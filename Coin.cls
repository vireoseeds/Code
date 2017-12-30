VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_Nominal As Double
Private p_Conv As Double
Private p_DebtTenor As Integer
Private p_RateCP As Double
Private p_RateDebt As Double
Private p_EqPayPortion As Double
Private p_Freq As Integer
Private p_UF As Double

Public Sub Init(CoinData As Range)

p_Nominal = CoinData(1, 1)
p_Conv = CoinData(1, 2) / 100
p_DebtTenor = WorksheetFunction.RoundUp(CoinData(1, 3) * 4, 0)
p_RateDebt = CoinData(1, 5) / 100
p_EqPayPortion = CoinData(1, 6) / 100
p_Freq = CoinData(1, 7)
p_RateCP = WorksheetFunction.RoundUp(CoinData(1, 4) / 100 * ((1 - p_Conv) + p_EqPayPortion * p_Conv), 4)

If p_Freq <> 3 And p_Freq <> 6 And p_Freq <> 12 Then
    p_Freq = 6
    Range("CoinData")(1, 7) = 6
End If

p_UF = CoinData(1, 8) / 100

End Sub

Public Property Get Nominal() As Double
Nominal = p_Nominal
End Property
Public Function ChangeNominal(ByVal NewNominal As Double)
p_Nominal = NewNominal
End Function
Public Property Get Conv() As Double
Conv = p_Conv
End Property
Public Property Get DebtTenor() As Double
DebtTenor = p_DebtTenor
End Property
Public Property Get RateCP() As Double
RateCP = p_RateCP
End Property
Public Property Get RateDebt() As Double
RateDebt = p_RateDebt
End Property
Public Property Get EqPayPortion() As Double
EqPayPortion = p_EqPayPortion
End Property
Public Property Get Freq() As Double
Freq = p_Freq
End Property
Public Property Get UF() As Double
UF = p_UF
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Equity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_Amount As Double
Private p_ROE As Double
Private p_Appr As Double

Public Sub Init(EquityData As Range, Row As Integer)

p_Amount = EquityData(Row, 1)
p_ROE = EquityData(Row, 2) / 100
p_Appr = EquityData(Row, 3) / 100

End Sub

Public Property Get Amount() As Double
Amount = p_Amount
End Property
Public Property Get ROE() As Double
ROE = p_ROE
End Property
Public Property Get Appr() As Double
Appr = p_Appr
End Property


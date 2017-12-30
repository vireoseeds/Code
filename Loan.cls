VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_Name As String
Private p_Nominal As Double
Private p_StartDate As Integer
Private p_Tenor As Integer
Private p_GP As Integer
Private p_Rate As Double
Private p_UF As Double
Private p_CF As Double
Private p_Freq As Integer
Private p_Bullet As Boolean
Private p_AllAtGP As Boolean

Public Sub Init(Data As Range, Row As Integer)

p_Name = Data(Row, 1)
If p_Name = "" Then Exit Sub
p_Nominal = Data(Row, 2)
p_StartDate = Data(Row, 3)
p_Tenor = WorksheetFunction.RoundUp(Data(Row, 4) * 4, 0)
p_GP = Data(Row, 5)
p_Rate = Data(Row, 6) / 100
p_UF = Data(Row, 7) / 100
p_CF = Data(Row, 8) / 100
p_Freq = Data(Row, 9)

If p_Freq <> 3 And p_Freq <> 6 And p_Freq <> 12 Then
    p_Freq = 6
    Range("LoanData")(Row, 9) = 6
End If

p_Bullet = Data(Row, 10)
p_AllAtGP = False

End Sub

Public Function IsAvailable(ByVal Period) As Boolean

IsAvailable = False
If Period >= StartDate Then IsAvailable = True
If p_Nominal = 0 Then IsAvailable = False

End Function
Public Property Get AllAtGP() As Boolean
AllAtGP = p_AllAtGP
End Property
Public Function AllIsAtGP()
p_AllAtGP = True
End Function
Public Property Get Name() As String
    Name = p_Name
End Property
Public Property Get Nominal() As Double
    Nominal = p_Nominal
End Property
Public Property Get StartDate() As Integer
    StartDate = p_StartDate
End Property
Public Property Get Tenor() As Integer
    Tenor = p_Tenor
End Property
Public Property Get GP() As Integer
    GP = p_GP
End Property
Public Property Get Rate() As Double
    Rate = p_Rate
End Property
Public Property Get UF() As Double
    UF = p_UF
End Property
Public Property Get CF() As Double
    CF = p_CF
End Property
Public Property Get Freq() As Integer
    Freq = p_Freq
End Property
Public Property Get Bullet() As Boolean
    Bullet = p_Bullet
End Property

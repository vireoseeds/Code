Attribute VB_Name = "Functions"
Option Explicit
Option Base 1

Public Function RankInstr(ByRef Instru() As String, ByRef Cash() As Double, ByRef Disposable() As Boolean)

Dim tmpInstr(8) As String: Call Copy_SMatrix(Instru, tmpInstr, 8)
Dim tmpCash(8, 2) As Double: Call Copy_DMatrix(Cash, tmpCash, 8, 2)
Dim tmpDisp(8) As Boolean: Call Copy_BMatrix(Disposable, tmpDisp, 8)
Dim i, count As Integer

count = 1
Dim LowerRateRnk As Integer: LowerRateRnk = 1

While count <= 8
    For i = 1 To 8
        If tmpCash(i, 2) < tmpCash(LowerRateRnk, 2) Then LowerRateRnk = i
    Next i
    
    Cash(count, 1) = tmpCash(LowerRateRnk, 1)
    Cash(count, 2) = tmpCash(LowerRateRnk, 2)
    Instru(count) = tmpInstr(LowerRateRnk)
    Disposable(count) = tmpDisp(LowerRateRnk)
    tmpCash(LowerRateRnk, 2) = 100
    
    count = count + 1
    LowerRateRnk = 1
Wend

End Function

Private Function Copy_SMatrix(ByRef MInput() As String, ByRef MOutput() As String, Row As Integer, Optional Col As Integer = 1)

Dim i, j As Integer

If Col = 1 Then
    For i = 1 To Row
        MOutput(i) = MInput(i)
    Next i
Else
    For i = 1 To Row
        For j = 1 To Col
            MOutput(i, j) = MInput(i, j)
        Next j
    Next i
End If

End Function

Private Function Copy_DMatrix(ByRef MInput() As Double, ByRef MOutput() As Double, Row As Integer, Optional Col As Integer = 1)

Dim i, j As Integer

If Col = 1 Then
    For i = 1 To Row
        MOutput(i) = MInput(i)
    Next i
Else
    For i = 1 To Row
        For j = 1 To Col
            MOutput(i, j) = MInput(i, j)
        Next j
    Next i
End If

End Function
Private Function Copy_IMatrix(ByRef MInput() As Integer, ByRef MOutput() As Integer, Row As Integer, Optional Col As Integer = 1)

Dim i, j As Integer

If Col = 1 Then
    For i = 1 To Row
        MOutput(i) = MInput(i)
    Next i
Else
    For i = 1 To Row
        For j = 1 To Col
            MOutput(i, j) = MInput(i, j)
        Next j
    Next i
End If

End Function
Private Function Copy_BMatrix(ByRef MInput() As Boolean, ByRef MOutput() As Boolean, Row As Integer, Optional Col As Integer = 1)

Dim i, j As Integer

If Col = 1 Then
    For i = 1 To Row
        MOutput(i) = MInput(i)
    Next i
Else
    For i = 1 To Row
        For j = 1 To Col
            MOutput(i, j) = MInput(i, j)
        Next j
    Next i
End If

End Function

Public Function IsPair(ByVal Num As Integer) As Integer

IsPair = 1
If Num * 5 - WorksheetFunction.RoundDown(Num / 2, 0) * 10 <> 0 Then IsPair = 0

End Function

Public Function IsArrayInitalized(myArray As Variant) As Boolean

On Error GoTo NotInit

Dim count As Integer: count = UBound(myArray)

IsArrayInitalized = True

Exit Function

NotInit:

IsArrayInitalized = False

End Function

Public Function GetWACC(ByVal Equity As Double, ByVal CoE As Double, ByVal Debt As Double, ByVal CoD As Double, ByVal TaxRate As Double)

GetWACC = Round(Equity / (Debt + Equity) * CoE + Debt / (Debt + Equity) * CoD * (1 - TaxRate), 4)

End Function

Public Function GetTaxRate(ByVal TaxType, Optional ByVal CsP As Double = 0#) As Double

Dim i As Integer

Select Case TaxType
    Case "Constant":
        GetTaxRate = Range("TaxesCste").Cells(1, 1).Value / 100
    Case "Multi":
        Dim tmpRate As Double: tmpRate = 0#
        
        For i = 1 To Range("TaxesMulti").Rows.count
            If CsP < Range("TaxesMulti").Cells(i, 2).Value Then
                tmpRate = (CsP - Range("TaxesMulti").Cells(i, 1).Value + 1) * Range("TaxesMulti").Cells(i, 3).Value / 100
            Else
                tmpRate = (Range("TaxesMulti").Cells(i, 2).Value - Range("TaxesMulti").Cells(i, 1).Value + 1) * Range("TaxesMulti").Cells(i, 3).Value / 100
            End If
        Next i
        
        tmpRate = tmpRate / CsP
End Select

End Function

Public Function GetIRR(ByRef CF() As Variant) As Double

Dim i, countPos, countNeg, count As Integer
Dim tmpRate As Double: tmpRate = 0.1
Dim Sum As Double
Dim WasNeg As Boolean: WasNeg = True

Sum = 0#
countPos = 0
countNeg = 0
count = 0

For i = 1 To UBound(CF)
    Sum = Sum + CF(i) / WorksheetFunction.Power(1 + tmpRate / 4, i)
Next i

If Sum > 100 Then
    tmpRate = tmpRate + 0.015
    WasNeg = False
ElseIf Sum < 100 Then
    tmpRate = tmpRate - 0.07
    WasNeg = True
End If

Recomp:
count = count + 1
If count = 20 Then GoTo EndProcess
Sum = 0#

For i = 1 To UBound(CF)
    Sum = Sum + CF(i) / WorksheetFunction.Power(1 + tmpRate / 4, i)
Next i

If Sum > 2000 Then
    tmpRate = tmpRate + 0.015 / (1 + 2 * countPos)
    If WasNeg Then countNeg = countNeg + 1
    WasNeg = False
    GoTo Recomp
ElseIf Sum < -2000 Then
    tmpRate = tmpRate - 0.007 / (1 + 2 * countNeg)
    If Not WasNeg Then countPos = countPos + 1
    WasNeg = True
    GoTo Recomp
End If

EndProcess:

GetIRR = Round(tmpRate, 4)

End Function

Public Function wsExists(ws As Worksheet) As Boolean

On Error GoTo NotExist

Dim res As String: res = ws.Name
wsExists = True

Exit Function

NotExist:

wsExists = False

End Function

Public Function Average(myRange As Range) As Double

Dim i, j As Integer
Dim Avg As Double: Avg = 0#

For i = 1 To myRange.Columns.count
    For j = 1 To myRange.Rows.count
        Avg = Avg + myRange.Cells(j, i).Value
    Next j
Next i

Avg = Avg / (myRange.Columns.count * myRange.Rows.count)

Average = Avg

End Function

Public Function Lang() As String

Dim CurrentLanguage As Long

CurrentLanguage = Application.LanguageSettings.LanguageID(msoLanguageIDUI)

Select Case CurrentLanguage
Case 1036, 2060, 11276, 3084, 9228, 12300, 15372, 5132, 13324, 6156, 14348, 58380, 8204, 10252, 4108, 7180:
    Lang = "French"
Case 1033, 2057, 3081, 10249, 4105, 9225, 15369, 16393, 14345, 6153, 8201, 17417, 5129, 13321, 18441, 7177, 11273, 12297:
    Lang = "English"
End Select

End Function


Option Explicit
Option Base 1

Public Sub designConstructionTable()

Dim i, CP As Integer: CP = WorksheetFunction.RoundUp(Range("ConstrPeriod").Value * 4, 0)

On Error GoTo IssueDesignCT

Worksheets("Constr CF").Range("B4:D90").ClearContents

If CP = 0 Then Exit Sub

For i = 1 To CP
    Worksheets("Constr CF").Cells(i + 3, 2).Value = "Q " & i
Next i

Exit Sub

IssueDesignCT:
Debug.Print "Error in designConstructionTable"

End Sub
Public Function designRiskTable(Optional ws As Worksheet = Empty, Optional ByVal Delay As Integer = -1) As Integer

On Error GoTo RiskTableError

Dim i, CP As Integer: CP = WorksheetFunction.RoundUp(Range("ConstrPeriod").Value * 4, 0)
Dim ConcP As Integer: ConcP = WorksheetFunction.RoundUp(Range("ConcPeriod").Value * 4, 0)
Dim OMImpact As Double: OMImpact = 1#
If Delay = -1 Then Delay = Range("Delay").Value

If Not wsExists(ws) Then
    Set ws = ActiveSheet
End If

If ws.Name = "Deg Risk" Then OMImpact = 1 - Range("OMImpact").Cells(1, 1).Value / 100

Dim Rate As Double
Dim Params() As Double

Worksheets("Graph Data").Range("A1:BB7").ClearContents

For i = 2 To WorksheetFunction.RoundDown((ConcP - CP - Delay) / 4, 0) + 1
    Worksheets("Graph Data").Cells(1, i).Value = "Year " & (i - 1)
Next i
    
If (ConcP - CP - Delay) / 4 * 10 - WorksheetFunction.RoundDown((ConcP - CP - Delay) / 4, 0) * 10 > 0 Then
    Worksheets("Graph Data").Cells(1, i).Value = "Year " & (i - 1)
    i = i + 1
Else

End If

designRiskTable = i - 2

Select Case ws.Name
    Case "Deg Risk":
        Dim Model As String: Model = Worksheets("Deg Risk").DegradationBox.Value
        If Model = "" Then
            MsgBox ("No degradation model has been selected")
            Exit Function
        End If
    Case "Clim Risk":
        Model = Worksheets("Clim Risk").ClimateBox.Value
        If Model = "" Then
            MsgBox ("No climate model has been selected")
            Exit Function
        End If
    Case Else:
        Exit Function
End Select

ws.Range("H4:I200").ClearContents

For i = 1 To CP + Delay
    ws.Cells(3 + i, 8).Value = "Q " & i
    ws.Cells(3 + i, 9).Value = 0
Next i
        
Select Case Model
    Case "Linear":
        If ws.Name = "Deg Risk" Then
            Rate = ws.Range("LinearDeg")(1).Value / 100
        Else
            Rate = ws.Range("LinearClim")(1).Value / 100
        End If
        
        For i = CP + Delay + 1 To ConcP
            ws.Cells(3 + i, 8).Value = "Q " & i
            ws.Cells(3 + i, 9).Value = WorksheetFunction.Max(1 - OMImpact * Rate * (i - CP + Delay - 1) / 4, 0)
        Next i
    Case "Multi-Linear":
        Dim PrevDF As Double: PrevDF = 1#
        For i = CP + Delay + 1 To ConcP
            Rate = OMImpact * GetRate(i - (CP + Delay + 1), ws, "Multi-Linear")
            ws.Cells(3 + i, 8).Value = "Q " & i
            PrevDF = WorksheetFunction.Max(PrevDF - Rate / 4, 0)
            ws.Cells(3 + i, 9).Value = PrevDF
        Next i
    Case "Stepped":
        PrevDF = 1#
        Dim PrevP As Integer: PrevP = 0
        Dim CurP As Integer
        
        For i = CP + Delay + 1 To ConcP
            Rate = OMImpact * GetRate(i - (CP + Delay + 1), ws, "Stepped", CurP)
            
            If CurP = PrevP Then
                Rate = 0#
            Else
                PrevP = CurP
            End If
            
            ws.Cells(3 + i, 8).Value = "Q " & i
            PrevDF = WorksheetFunction.Max(PrevDF - Rate, 0)
            ws.Cells(3 + i, 9).Value = PrevDF
        Next i
    Case "Cyclic Collapse":
        Dim Param As Range: Set Param = ws.Range("CycColClim")
        ReDim Params(2)
        Dim LastPeriod As Integer: LastPeriod = CP + Delay + 1
        
        If Param(2, 1) = 0 Then
            MsgBox ("Period for cyclic collapse model is equal to 0")
            Exit Function
        End If
        
        Params(1) = Param(2, 1)
        Params(2) = Param(4, 1)
        
        For i = CP + Delay + 1 To ConcP
            Dim TrendEf As Double: TrendEf = Abs((i - (CP + Delay + 1)) / 4 * Param(1, 1) / 100)
            Dim SubEf As Double: SubEf = Abs(Sin(WorksheetFunction.Pi() * (i - LastPeriod) / Params(1)) * Params(2) / 100)
            
            ws.Cells(3 + i, 8).Value = "Q " & i
            ws.Cells(3 + i, 9).Value = Round(1 - (TrendEf + SubEf), 6)
            
            If i - LastPeriod = Params(1) Then
                Params(1) = WorksheetFunction.Max(Params(1) + Param(3, 1), 1)
                Params(2) = WorksheetFunction.Max(Params(2) - Param(5, 1), 0)
                LastPeriod = i
            End If
            
        Next i
        
    Case "Cyclic Expansion":
        Set Param = ws.Range("CycExpClim")
        LastPeriod = CP + Delay + 1
        
        ReDim Params(2)
        
        If Param(2, 1) = 0 Then
            MsgBox ("Period for cyclic expansion model is equal to 0")
            Exit Function
        End If
        
        Params(1) = Param(2, 1)
        Params(2) = Param(4, 1)
        
        For i = CP + Delay + 1 To ConcP
            TrendEf = Abs((i - (CP + Delay + 1)) / 4 * Param(1, 1) / 100)
            SubEf = Abs(Sin(WorksheetFunction.Pi() * (i - LastPeriod) / Params(1)) * Params(2) / 100)
            
            ws.Cells(3 + i, 8).Value = "Q " & i
            ws.Cells(3 + i, 9).Value = Round(1 - (TrendEf + SubEf), 6)
            
            If i - LastPeriod = Params(1) Then
                Params(1) = WorksheetFunction.Max(Params(1) + Param(3, 1), 1)
                Params(2) = WorksheetFunction.Max(Params(2) + Param(5, 1), 0)
                LastPeriod = i
            End If
            
        Next i
        
    Case "Cyclic Curv":
        Set Param = ws.Range("CycCurClim")
            LastPeriod = CP + Delay + 1
        
            If Param(2, 1) = 0 Then
                MsgBox ("Period for cyclic curv model is equal to 0")
                Exit Function
            End If
            
            ReDim Params(4)
            
            Params(1) = Param(2, 1)
            Params(2) = Param(3, 1)
            Params(3) = Param(5, 1)
            Params(4) = Param(6, 1)
        
            For i = CP + Delay + 1 To ConcP
                TrendEf = Abs((i - (CP + Delay + 1)) / 4 * Param(1, 1) / 100)
                SubEf = Abs(Sin(WorksheetFunction.Pi() * (i - LastPeriod) / Params(1)) * Params(2) / 100)
                Dim SubSubEf As Double: SubSubEf = Abs(Sin(WorksheetFunction.Pi() * (i - (CP + Delay + 1)) / Params(3)) * Params(4) / 100)
            
                ws.Cells(3 + i, 8).Value = "Q " & i
                ws.Cells(3 + i, 9).Value = Round(1 - (TrendEf + SubEf + SubSubEf), 6)
            
                If i - LastPeriod = Params(1) Then
                    Params(2) = WorksheetFunction.Max(Params(2) + Param(4, 1), 0)
                    Params(4) = WorksheetFunction.Max(Params(4) + Param(7, 1), 0)
                    LastPeriod = i
                End If
            
        Next i
End Select

Exit Function

RiskTableError:
Debug.Print "Error in designRiskTable"

End Function

Public Function GetProdIncRate(ByVal Period As Integer) As Double

On Error GoTo GetRateError

Dim PrevPIPeriod, CurPIPeriod As Integer
PrevPIPeriod = 0
CurPIPeriod = 0
Dim i As Integer

CurPIPeriod = Range("PowerProdInc").Cells(1, 1).Value

If CurPIPeriod = 0 Then
    GetProdIncRate = 1#
    Exit Function
End If

If Period < CurPIPeriod Then
    GetProdIncRate = Range("PowerProdInc").Cells(1, 2).Value / 100
    Exit Function
End If

PrevPIPeriod = CurPIPeriod

For i = 2 To Range("PowerProdInc").Rows.count
    If Range("PowerProdInc").Cells(i, 1).Value = 0 Then
        GetProdIncRate = 1#
        Exit Function
    End If
    
    CurPIPeriod = CurPIPeriod + Range("PowerProdInc").Cells(i, 1).Value
    
    If Period > PrevPIPeriod And Period < CurPIPeriod Then
        GetProdIncRate = Range("PowerProdInc").Cells(i, 2).Value / 100
        Exit Function
    End If
    
    PrevPIPeriod = PrevPIPeriod + CurPIPeriod
Next i

Exit Function

GetRateError:
Debug.Print "Error in GetProdIncRate function"
GetProdIncRate = 0

End Function

Private Function GetRate(ByVal Period As Integer, ByRef ws As Worksheet, ByVal Model As String, Optional ByRef P As Integer = 0) As Double

On Error GoTo GetRateError

Dim table As Range
Dim RangeName As String

Select Case Model
    Case "Multi-Linear":
        RangeName = "Multi"
    Case "Stepped":
        RangeName = "Stepped"
    Case Else
        RangeName = Model
End Select

If ws.Name = "Deg Risk" Then
    RangeName = RangeName & "Deg"
ElseIf ws.Name = "Clim Risk" Then
    RangeName = RangeName & "Clim"
End If

Set table = ws.Range(RangeName)
Dim i As Integer

For i = 1 To 12
    If table(i, 1) <= Period And table(i, 2) >= Period Then
        GetRate = table(i, 3) / 100
        P = i
        Exit Function
    End If
Next i

MsgBox ("No rate for period " & Period)
GetRate = 1

Exit Function

GetRateError:
Debug.Print "Error in GetRate function"
GetRate = 0

End Function

Public Sub GetPPA(Optional ByVal Delay As Integer = -1)

On Error GoTo GetPPAError

Dim i, CP As Integer: CP = WorksheetFunction.RoundUp(Range("ConstrPeriod").Value * 4, 0)
Dim ConcP As Integer: ConcP = WorksheetFunction.RoundUp(Range("ConcPeriod").Value * 4, 0)

If Delay = -1 Then Delay = Range("Delay").Value

Dim PPAVal As Range: Set PPAVal = Range("PPAData")
Dim ws As Worksheet: Set ws = Worksheets("PPA")

ws.Range("B4:C200").ClearContents

Dim StartP As Integer: StartP = CP + Delay + 1

If PPAVal(1, 1) = 0 Or PPAVal(2, 1) = 0 Then
    MsgBox ("Issue with PPA parameters definition")
    Exit Sub
End If

Dim PPAa As Double: PPAa = Round((PPAVal(1, 1) * (ConcP - StartP) - PPAVal(3, 1) * PPAVal(2, 1)) / (ConcP - StartP - PPAVal(3, 1)), 6)

For i = StartP To ConcP
    ws.Cells(3 + i - StartP + 1, 2).Value = "Q " & i
    If (i - StartP + 1) <= PPAVal(3, 1) Then
        ws.Cells(3 + i - StartP + 1, 3).Value = PPAVal(2, 1)
    Else
        ws.Cells(3 + i - StartP + 1, 3).Value = PPAa
    End If
Next i

Exit Sub

GetPPAError:
Debug.Print "Error in GetPPA function"

End Sub

Public Sub SumFinancialCosts()

Dim CP As Integer: CP = WorksheetFunction.RoundUp(Range("ConstrPeriod").Value * 4, 0)
Dim Delay As Integer: Delay = Range("Delay").Value

Dim count As Integer: count = 5
Dim Period As String: Period = Worksheets("CF").Cells(2, count).Value

While Period <> ""
    Worksheets("CF").Cells(33, count).Value = Worksheets("CF").Cells(24, count).Value + Worksheets("CF").Cells(25, count).Value + Worksheets("CF").Cells(30, count).Value + Worksheets("CF").Cells(31, count).Value + Worksheets("CF").Cells(32, count).Value
    Dim LoanRepay As Double
    If IsNumeric(Worksheets("CF").Cells(27, count - 1).Value) Then
        LoanRepay = Worksheets("CF").Cells(27, count - 1).Value - Worksheets("CF").Cells(27, count).Value
        If LoanRepay < 0 Then LoanRepay = 0
    End If
    Worksheets("CF").Cells(33, count).Value = Worksheets("CF").Cells(33, count).Value + LoanRepay
    
    If IsNumeric(Worksheets("CF").Cells(28, count - 1).Value) Then
        LoanRepay = Worksheets("CF").Cells(28, count - 1).Value - Worksheets("CF").Cells(28, count).Value
        If LoanRepay < 0 Then LoanRepay = 0
    End If
    If count - 4 > CP + Delay + 1 Then Worksheets("CF").Cells(33, count).Value = Worksheets("CF").Cells(33, count).Value + LoanRepay
    
    count = count + 1
    Period = Worksheets("CF").Cells(2, count).Value
Wend

End Sub

Public Function GetCost(CostType As String, Revenues() As Double, Optional Rebate As Double = 0#, Optional Delay As Integer = -1) As Double()

Dim Res() As Double
Dim i, CP As Integer: CP = WorksheetFunction.RoundUp(Range("ConstrPeriod").Value * 4, 0)
If Delay = -1 Then Delay = Range("Delay").Value

ReDim Res(UBound(Revenues))

Select Case CostType
    Case "O&M":
        Select Case Worksheets("All Costs").OMBox.Value
            Case "Constant":
                For i = 1 To UBound(Res)
                    If Revenues(i) > 0 Then Res(i) = Round(WorksheetFunction.Max(Range("OMFloor").Cells(1, 1).Value / 4, Range("OMCste").Cells(1, 1).Value), 2)
                Next i
            Case "Multi":
                For i = 1 To UBound(Res)
                    If Revenues(i) > 0 Then
                        Res(i) = Round(WorksheetFunction.Max(Revenues(i) * GetRate(i - CP - Delay, Worksheets("All Costs"), "OMMulti"), Range("OMFloor").Cells(1, 1).Value / 4), 2)
                    End If
                Next i
            Case Else:
                For i = 1 To UBound(Res)
                    Res(i) = 0
                Next i
        End Select
    Case "SG&A":
        Select Case Worksheets("All Costs").SGABox.Value
            Case "Constant":
                For i = 1 To UBound(Res)
                    If Revenues(i) > 0 Then Res(i) = Round(WorksheetFunction.Max(Range("SGAFloor").Cells(1, 1).Value / 4, Range("SGACste").Cells(1, 1).Value), 2)
                Next i
            Case "Multi":
                For i = 1 To UBound(Res)
                    If Revenues(i) > 0 Then
                        Res(i) = Round(WorksheetFunction.Max(Revenues(i) * GetRate(i - CP - Delay, Worksheets("All Costs"), "SGAMulti"), Range("SGAFloor").Cells(1, 1).Value / 4), 2)
                    End If
                Next i
            Case Else:
                For i = 1 To UBound(Res)
                    Res(i) = 0
                Next i
        End Select
    Case "Royalties":
        Select Case Worksheets("All Costs").RoyaltiesBox.Value
            Case "Constant":
                For i = 1 To UBound(Res)
                    If Revenues(i) > 0 Then Res(i) = WorksheetFunction.Max(Round((Revenues(i) - Rebate) * Range("RoyaltiesCste").Cells(1, 1).Value / 100, 2), 0)
                Next i
            Case "Multi":
                For i = 1 To UBound(Res)
                    If Revenues(i) > 0 Then
                        Res(i) = WorksheetFunction.Max(Round((Revenues(i) - Rebate) * GetRate(i, Worksheets("All Costs"), "RoyaltiesMulti"), 2), 0)
                    End If
                Next i
            Case Else:
                For i = 1 To UBound(Res)
                    Res(i) = 0
                Next i
        End Select
    Case "Taxes":
        Select Case Worksheets("All Costs").TaxesBox.Value
            Case "Constant":
                For i = 1 To UBound(Res)
                    If Revenues(i) > 0 Then Res(i) = WorksheetFunction.Max(Round((Revenues(i) - Rebate) * Range("TaxesCste").Cells(1, 1).Value / 100, 2), 0)
                Next i
            Case "Multi":
                For i = 1 To UBound(Res)
                    If Revenues(i) > 0 Then
                        Res(i) = WorksheetFunction.Max(Round((Revenues(i) - Rebate) * GetRate(i, Worksheets("All Costs"), "TaxesMulti"), 2), 0)
                    End If
                Next i
            Case Else:
                For i = 1 To UBound(Res)
                    Res(i) = 0
                Next i
        End Select
End Select

GetCost = Res

End Function


Public Function GetAvgCashYield(ByRef ProjAvgCashYield As Variant, ByRef CoinsNotional As Variant) As Double

Dim i As Integer
Dim AvgCY, tmpSum As Double
AvgCY = 0#
tmpSum = 0#

For i = 1 To ProjAvgCashYield.count
    tmpSum = tmpSum + CoinsNotional(i)
    AvgCY = AvgCY + ProjAvgCashYield(i) * CoinsNotional(i)
Next i

GetAvgCashYield = AvgCY / tmpSum

End Function
Public Function GetMPYield(ByRef ProjectPower As Variant, ByRef ProjectPPAs As Variant, ByRef ProjectCosts As Variant) As Double()

Dim i As Integer
Dim AvgPrice1W, AvgPrice1kWh, tmpSum As Double
Dim Res(3) As Double
AvgPrice1W = 0#
AvgPrice1kWh = 0#
tmpSum = 0#

For i = 1 To ProjectPower.count
    tmpSum = tmpSum + ProjectPower(i)
    AvgPrice1W = AvgPrice1W + ProjectCosts(i)
Next i

Res(1) = Math.Round(AvgPrice1W / tmpSum, 6)

tmpSum = 0#

For i = 1 To ProjectPPAs.count
    tmpSum = tmpSum + ProjectPower(i)
    AvgPrice1kWh = AvgPrice1kWh + ProjectPPAs(i) * ProjectPower(i)
Next i

Res(2) = Math.Round(AvgPrice1kWh / tmpSum, 6)

Res(3) = Round((1 + Range("SecurityReturn1Y").Value) * WorksheetFunction.Power(Res(1), Res(2) / Res(1)) - 1, 4)

GetMPYield = Res

End Function
Public Function ComputeVireoShares(ByVal CoinEquityAmount As Double, ByVal Equities As Dictionary) As Double

Dim TotalEquity, SubTotal, tmpCoinEq As Double
TotalEquity = 0#
SubTotal = 0#
Dim Equity As Variant
Dim count As Integer: count = 1

For Each Equity In Equities.Items
    SubTotal = SubTotal + Equity.Amount * (1 + Equity.Appr)
Next

TotalEquity = SubTotal + CoinEquityAmount

tmpCoinEq = WorksheetFunction.Min(Round(CoinEquityAmount / TotalEquity, 4), Range("CapEquity").Cells(1, 1).Value)

Worksheets("Summary").Cells(10, 4).Value = tmpCoinEq

For Each Equity In Equities.Items
    Worksheets("Summary").Cells(7 + count, 4).Value = Round((1 - tmpCoinEq) * Equity.Amount * (1 + Equity.Appr) / SubTotal, 4)
    count = count + 1
Next

ComputeVireoShares = tmpCoinEq


End Function

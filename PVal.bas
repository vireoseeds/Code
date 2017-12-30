Option Explicit
Option Base 1

Private Loans As Dictionary
Private Equities As Dictionary
Private Coins As Coin
Private CsP, CP, Delay As Integer
Private CapexInc As Double
Private AvgOM, AvgSGA, AvgRevenues As Double

Public Sub LaunchNoRisk()

Call Launch(True)

End Sub

Public Sub Launch(Optional IsRisk As Boolean = False)

Dim IsOk, IsFirstRound As Boolean: IsOk = True
Dim InterestCoverage() As Double
Dim ErrMsg As String
Dim i, MaxYear As Integer
CP = WorksheetFunction.RoundUp(Range("ConstrPeriod").Value * 4, 0)
CsP = WorksheetFunction.RoundUp(Range("ConcPeriod").Value * 4, 0)
Delay = Range("Delay").Value

If Delay > 3 Then
    Delay = 3
    Range("Delay").Value = 3
End If

CapexInc = Range("CapexInc").Value / 100

IsFirstRound = True

On Error GoTo EndAnalysis

ErrMsg = "InitLoan"
For i = 1 To Range("LoanMsg").Rows.count
    Range("LoanMsg").Cells(i, 1).Value = ""
Next i
IsOk = InitLoans()
If Not IsOk Then GoTo EndAnalysis
ErrMsg = "InitEquity"
IsOk = InitEquity()
If Not IsOk Then GoTo EndAnalysis

Set Coins = New Coin
Coins.Init CoinData:=Range("CoinData")

SecondRound:
MaxYear = designRiskTable(Worksheets("Deg Risk"), Delay)
Call designRiskTable(Worksheets("Clim Risk"), Delay)
Call GetPPA(Delay)

ErrMsg = "InitCFTable"
Worksheets("CF").Range("E3:EH100").ClearContents
If Not InitCFTable() Then Exit Sub
ErrMsg = "WithDawCash"
Call WithDrawCash(InterestCoverage)
ErrMsg = "CheckLoanGP"
Call CheckLoanGP
ErrMsg = "GetUF"
Call GetUF
ErrMsg = "GetOutstandingNomAndCF"
Call GetOutstandingNomAndCF
ErrMsg = "DebtRepaymentAndInterest"
Call DebtRepaymentAndInterest
Call Tools.SumFinancialCosts
ErrMsg = "GetEBIT"
Call GetEBIT
ErrMsg = "CoverInterests"
Dim Revenues() As Double
Revenues = CoverInterests()
ErrMsg = "RoyaltiesAndTaxes"
Call RoyaltiesAndTaxes(Revenues)

If IsFirstRound Then
    IsFirstRound = False
    InterestCoverage = AdjustCoinNominal()
    If Not Worksheets("Param").IncludeConstRisk.Value Then
        Delay = 0
        CapexInc = 0#
        Call designRiskTable(Worksheets("Deg Risk"), 0)
        Call designRiskTable(Worksheets("Clim Risk"), 0)
    End If
    GoTo SecondRound
End If

Call GetKPI

ErrMsg = "GetDataForInvestors"
Call GetDataForInvestors

ErrMsg = "DoGraph"
Call DoGraph

ErrMsg = "YieldCurveSmoother"
Dim rngMax As String: rngMax = Worksheets("Graph Data").Cells(3, 1 + MaxYear).Address
Call YieldCurveSmoother(Worksheets("Graph Data").Range("$B$3:" & rngMax))

ErrMsg = "GetCropsYield"
Call GetCropsYield

If Not IsRisk Then Call GetRiskIndicator

Exit Sub
EndAnalysis:
MsgBox ("Error while proceeding to project analysis: issue with " & ErrMsg)

End Sub

Private Sub GetKPI()

Dim tmpInst As Variant
Dim TotalPrice As Double

TotalPrice = Coins.Nominal

For Each tmpInst In Loans.Items
    TotalPrice = TotalPrice + tmpInst.Nominal
Next tmpInst

For Each tmpInst In Equities.Items
    TotalPrice = TotalPrice + tmpInst.Amount
Next tmpInst

Range("ProjectKPI").Cells(1, 1).Value = Round(AvgRevenues / (Range("PowerProd").Value * 1000), 2)
Range("ProjectKPI").Cells(2, 1).Value = TotalPrice / (Range("PowerProd").Value * 1000)
Range("ProjectKPI").Cells(3, 1).Value = Round(AvgOM / (Range("PowerProd").Value * 1000 * (Range("ConcPeriod").Value - Range("ConstrPeriod").Value)), 4) * 100
Range("ProjectKPI").Cells(4, 1).Value = Round(AvgSGA / (Range("PowerProd").Value * 1000 * (Range("ConcPeriod").Value - Range("ConstrPeriod").Value)), 4) * 100

End Sub

Private Sub GetCropsYield()

Dim year As String: year = Worksheets("Graph Data").Cells(1, 2).Value
Dim count As Integer: count = 1
Dim AvgCropYield As Double: AvgCropYield = 0#

While year <> ""
    AvgCropYield = AvgCropYield + Worksheets("Graph Data").Cells(3, 1 + count).Value
    Worksheets("Graph Data").Cells(7, 1 + count).Value = AvgCropYield / count
    count = count + 1
    year = Worksheets("Graph Data").Cells(1, 1 + count).Value
Wend

End Sub
Private Sub DoGraph()

Dim i, j As Integer
Dim FirstPeriod, AvgCropYield, AvgCashYield As Double

AvgCropYield = 0#
AvgCashYield = 0#
FirstPeriod = 4 - ((CP + Delay) / 4 - WorksheetFunction.RoundDown((CP + Delay) / 4, 0)) * 4

Worksheets("Graph Data").Cells(2, 1).Value = "Cash Forwards"
Worksheets("Graph Data").Cells(2, 2).Value = 0#
Worksheets("Graph Data").Cells(3, 1).Value = "Crops Yield"
Worksheets("Graph Data").Cells(3, 2).Value = 0#
Worksheets("Graph Data").Cells(4, 1).Value = "Carbon Emission Reduction"
Worksheets("Graph Data").Cells(4, 2).Value = 0#
Worksheets("Graph Data").Cells(6, 1).Value = "Cash Yield Curve"
Worksheets("Graph Data").Cells(6, 2).Value = 0#
Worksheets("Graph Data").Cells(7, 1).Value = "Average Crops Yield Curve"
Worksheets("Graph Data").Cells(7, 2).Value = 0#
Worksheets("Graph Data").Cells(10, 1).Value = "Revenues"
Worksheets("Graph Data").Cells(10, 2).Value = 0#
Worksheets("Graph Data").Cells(11, 1).Value = "Expenses"
Worksheets("Graph Data").Cells(11, 2).Value = 0#
Worksheets("Graph Data").Cells(12, 1).Value = "Accumlated Project Cash"
Worksheets("Graph Data").Cells(12, 2).Value = 0#

For i = 4 + CP + Delay To 4 + CP + Delay + FirstPeriod
    Worksheets("Graph Data").Cells(2, 2).Value = Worksheets("Graph Data").Cells(2, 2).Value + Round(Worksheets("CF").Cells(55, i).Value / FirstPeriod, 5)   'Crops
    Worksheets("Graph Data").Cells(3, 2).Value = Worksheets("Graph Data").Cells(3, 2).Value + Round(Worksheets("CF").Cells(56, i).Value / FirstPeriod, 5)   'Cash
    Worksheets("Graph Data").Cells(4, 2).Value = Worksheets("Graph Data").Cells(4, 2).Value + Worksheets("CF").Cells(57, i).Value                           'tCO2 Emissions reduction
    Worksheets("Graph Data").Cells(10, 2).Value = Worksheets("Graph Data").Cells(10, 2).Value + Worksheets("CF").Cells(17, i).Value                           'Revenue
    Worksheets("Graph Data").Cells(11, 2).Value = Worksheets("Graph Data").Cells(11, 2).Value + Worksheets("CF").Cells(18, i).Value + Worksheets("CF").Cells(19, i).Value + Worksheets("CF").Cells(20, i).Value + Worksheets("CF").Cells(33, i).Value + Worksheets("CF").Cells(46, i).Value 'Expenses
Next i

Worksheets("Graph Data").Cells(12, 2).Value = Worksheets("Graph Data").Cells(12, 2).Value + Worksheets("Graph Data").Cells(10, 2).Value - Worksheets("Graph Data").Cells(11, 2).Value       'Project Accumulated Cash

AvgCashYield = Worksheets("Graph Data").Cells(2, 2).Value
'AvgCropYield = Worksheets("Graph Data").Cells(3, 2).Value

Worksheets("Graph Data").Cells(6, 2).Value = AvgCashYield
'Worksheets("Graph Data").Cells(7, 2).Value = AvgCropYield

FirstPeriod = 5 + CP + Delay + FirstPeriod

For i = FirstPeriod To 4 + CsP Step 4
    Dim Index As Integer: Index = WorksheetFunction.RoundDown((i - FirstPeriod) / 4 + 1, 0) + IIf(FirstPeriod = 5 + CP + Delay, 1, 2)
    Worksheets("Graph Data").Cells(2, Index).Value = 0#
    Worksheets("Graph Data").Cells(3, Index).Value = 0#
    Worksheets("Graph Data").Cells(4, Index).Value = 0#
    Worksheets("Graph Data").Cells(10, Index).Value = 0#
    Worksheets("Graph Data").Cells(11, Index).Value = 0#
    Worksheets("Graph Data").Cells(12, Index).Value = 0#
    For j = 0 To WorksheetFunction.Min(3, 4 + CsP - i)
        Worksheets("Graph Data").Cells(2, Index).Value = Worksheets("Graph Data").Cells(2, Index).Value + Worksheets("CF").Cells(55, i + j).Value                           'Crops
        Worksheets("Graph Data").Cells(3, Index).Value = Worksheets("Graph Data").Cells(3, Index).Value + Worksheets("CF").Cells(56, i + j).Value                           'Cash
        Worksheets("Graph Data").Cells(4, Index).Value = Worksheets("Graph Data").Cells(4, Index).Value + Worksheets("CF").Cells(57, i + j).Value                           'tCO2 Emissions
        Worksheets("Graph Data").Cells(10, Index).Value = Worksheets("Graph Data").Cells(10, Index).Value + Worksheets("CF").Cells(17, i + j).Value                           'Revenue
        Worksheets("Graph Data").Cells(11, Index).Value = Worksheets("Graph Data").Cells(11, Index).Value + Worksheets("CF").Cells(18, i + j).Value + Worksheets("CF").Cells(19, i + j).Value + Worksheets("CF").Cells(20, i + j).Value + Worksheets("CF").Cells(33, i + j).Value + Worksheets("CF").Cells(46, i + j).Value 'Expenses
    Next j
    
    Worksheets("Graph Data").Cells(12, Index).Value = Worksheets("Graph Data").Cells(12, Index).Value + Worksheets("Graph Data").Cells(12, Index - 1).Value + Worksheets("Graph Data").Cells(10, Index - 1).Value - Worksheets("Graph Data").Cells(11, Index - 1).Value 'Project Accumulated Cash
    Worksheets("Graph Data").Cells(2, Index).Value = Math.Round(Worksheets("Graph Data").Cells(2, Index).Value / j, 5)
    Worksheets("Graph Data").Cells(3, Index).Value = Math.Round(Worksheets("Graph Data").Cells(3, Index).Value / j, 5)
    
    AvgCashYield = AvgCashYield + Worksheets("Graph Data").Cells(2, Index).Value
    Worksheets("Graph Data").Cells(6, Index).Value = AvgCashYield / (Index - 1)
    'AvgCropYield = AvgCropYield + Worksheets("Graph Data").Cells(3, Index).Value
    'Worksheets("Graph Data").Cells(7, Index).Value = AvgCropYield / (Index - 1)
Next i

End Sub

Private Function GetCash0(ByVal IR1Y As Double) As Double

Dim i As Integer
Dim Res As Double: Res = 0#
Dim tmpNom1, tmpNom2 As Double: tmpNom1 = 0#

For i = 5 To 4 + CP + Delay
    tmpNom1 = tmpNom1 + Coins.Nominal - Worksheets("CF").Cells(28, i).Value
    tmpNom2 = tmpNom2 + Worksheets("CF").Cells(14, i).Value
    Res = Res + Worksheets("CF").Cells(50, i).Value
Next i

tmpNom1 = tmpNom1 / (CP + Delay) * 0.75 * IR1Y * (CP + Delay) / 12

Res = Res * (1 + 0.5 * (CP + Delay) / 24 * IR1Y) + (Coins.Nominal - tmpNom2) + tmpNom1

GetCash0 = Res

End Function

Private Sub GetDataForInvestors()

Dim i, idx As Integer
Dim PortionCashOut, VireoShares As Double
Dim tmpVal As Double: tmpVal = 0#
Dim IRR(), EIRR() As Variant
Dim BenefCash0 As Double: BenefCash0 = Round(GetCash0(Range("SecurityReturn1Y").Cells(1, 1).Value) / (CsP - CP - Delay), 2)
Dim AvgCashYield, AvgCropYield As Double
Dim ProjectCost, ProjectPPA, ProjectCapacity, ProjectAvgCshYlds, ProjectCoinsNotional, ProjectNames As Range
Dim MPYldRes, AvgCshYld As Double

ReDim IRR(CsP)
ReDim EIRR(CsP)
AvgCashYield = 0#
AvgCropYield = 0#

VireoShares = Tools.ComputeVireoShares(Coins.Conv * Coins.Nominal, Equities)

Dim CoinQRepay As Double: CoinQRepay = Round(Coins.Nominal * Coins.Conv / (CsP - CP - Delay), 2)
Dim AccNom As Double: AccNom = 0#
Dim VireoDBFileName As String: VireoDBFileName = "Vireo_DB.xlsm"
Dim VireoDBFilePathName As String: VireoDBFilePathName = ActiveWorkbook.Path & "\"
Dim ProjectName As String: ProjectName = Range("ProjectName").Value

For i = 5 To 4 + CsP
    tmpVal = Round(Worksheets("CF").Cells(48, i).Value * VireoShares * Range("CashDistrib").Cells(1, 1).Value / 100 + Worksheets("CF").Cells(32, i).Value, 2)
    
    If i > 4 + CP + Delay + 1 Then
        Worksheets("CF").Cells(50, i).Value = tmpVal + Worksheets("CF").Cells(28, i - 1).Value - Worksheets("CF").Cells(28, i).Value
        Worksheets("CF").Cells(54, i).Value = Round((BenefCash0 + Worksheets("CF").Cells(50, i).Value - CoinQRepay + AccNom * Range("SecurityReturn1Y").Cells(1, 1).Value / 400) / Worksheets("PPA").Cells(i - CP - Delay - 1, 3).Value, 2)
        Worksheets("CF").Cells(55, i).Value = Round(4 * (1 / Range("IssuanceRatio").Value) * (BenefCash0 + Worksheets("CF").Cells(50, i).Value - CoinQRepay + AccNom * Range("SecurityReturn1Y").Cells(1, 1).Value / 400) / Coins.Nominal, 4)
        AvgCashYield = AvgCashYield + Worksheets("CF").Cells(55, i).Value
        'Worksheets("CF").Cells(56, i).Value = Round(4 * Worksheets("CF").Cells(54, i).Value / Coins.Nominal * Range("CoinMV").Cells(1, 1).Value * Range("MPYield").Cells(1, 1).Value, 4)
        'AvgCropYield = AvgCropYield + Worksheets("CF").Cells(56, i).Value
        Worksheets("CF").Cells(53, i).Value = Round((Worksheets("CF").Cells(17, i).Value - Worksheets("CF").Cells(18, i).Value) * VireoShares / Worksheets("PPA").Cells(i - CP - Delay - 1, 3).Value, 2)
        Worksheets("CF").Cells(57, i).Value = Range("CO2Reduction").Cells(1, 1).Value * Worksheets("CF").Cells(53, i).Value / 1000
        Worksheets("CF").Cells(58, i).Value = Worksheets("CF").Cells(57, i).Value * Range("CarbonCredit").Cells(1, 1).Value
        IRR(i - 4) = Round(Worksheets("CF").Cells(21, i).Value - Worksheets("CF").Cells(3, i).Value)
        EIRR(i - 4) = Round(Worksheets("CF").Cells(35, i).Value - (Worksheets("CF").Cells(5, i).Value + Worksheets("CF").Cells(6, i).Value))
    Else
        Worksheets("CF").Cells(50, i).Value = tmpVal
        If i = 5 + CP + Delay Then
            Worksheets("CF").Cells(53, i).Value = Round((Worksheets("CF").Cells(17, i).Value - Worksheets("CF").Cells(18, i).Value) / Worksheets("PPA").Cells(4, 3).Value, 2)
            Worksheets("CF").Cells(54, i).Value = Round((BenefCash0 + Worksheets("CF").Cells(50, i).Value - CoinQRepay + AccNom * Range("SecurityReturn1Y").Cells(1, 1).Value / 400) / Worksheets("PPA").Cells(4, 3).Value, 2)
            Worksheets("CF").Cells(55, i).Value = Round(4 * (1 / Range("IssuanceRatio").Value) * (BenefCash0 + Worksheets("CF").Cells(50, i).Value - CoinQRepay + AccNom * Range("SecurityReturn1Y").Cells(1, 1).Value / 400) / Coins.Nominal, 4)
            AvgCashYield = AvgCashYield + Worksheets("CF").Cells(55, i).Value
            'Worksheets("CF").Cells(56, i).Value = Round(4 * Worksheets("CF").Cells(54, i).Value / Coins.Nominal * Range("CoinMV").Cells(1, 1).Value * Range("MPYield").Cells(1, 1).Value, 4)
            'AvgCropYield = AvgCropYield + Worksheets("CF").Cells(56, i).Value
            Worksheets("CF").Cells(57, i).Value = Range("CO2Reduction").Cells(1, 1).Value * Worksheets("CF").Cells(53, i).Value / 1000
            Worksheets("CF").Cells(58, i).Value = Worksheets("CF").Cells(57, i).Value * Range("CarbonCredit").Cells(1, 1).Value
        End If
        IRR(i - 4) = Round(Worksheets("CF").Cells(21, i).Value - Worksheets("CF").Cells(3, i).Value)
        EIRR(i - 4) = Round(Worksheets("CF").Cells(35, i).Value - (Worksheets("CF").Cells(5, i).Value + Worksheets("CF").Cells(6, i).Value))
        
        If i = 5 + CP + Delay Then
            EIRR(i - 4) = EIRR(i - 4) - Coins.Nominal * Coins.Conv
        End If
    End If
    
    AccNom = AccNom + CoinQRepay
Next i

Range("IRR").Cells(1, 1).Value = GetIRR(IRR)
Range("EIRR").Cells(1, 1).Value = GetIRR(EIRR)

Worksheets("Summary").Range("VireoRatios").Cells(2, 1).Value = Round(AvgCashYield / (CsP - CP - Delay), 4)

Dim OpenedDB As Object: Set OpenedDB = DB.OpenDB(VireoDBFilePathName + VireoDBFileName)
Dim OpenedDBSource As Worksheet: Set OpenedDBSource = DB.GetDBSource(OpenedDB, "VireoDB")
Set ProjectNames = OpenedDBSource.Range("ProjNames")

idx = Application.Match(ProjectName, ProjectNames, 1)
OpenedDBSource.Range("VDBRef").Offset(idx, 1).Value = Round(AvgCashYield / (CsP - CP - Delay), 4)
OpenedDBSource.Range("VDBRef").Offset(idx, 2).Value = Coins.Nominal / 1000000

Set ProjectCost = OpenedDBSource.Range("ProjCosts")
Set ProjectPPA = OpenedDBSource.Range("ProjPPAs")
Set ProjectCapacity = OpenedDBSource.Range("ProjCapacities")
Set ProjectAvgCshYlds = OpenedDBSource.Range("CashYld")
Set ProjectCoinsNotional = OpenedDBSource.Range("CoinNotional")

MPYldRes = Tools.GetMPYield(ProjectCapacity, ProjectPPA, ProjectCost)(3)
AvgCshYld = Tools.GetAvgCashYield(ProjectAvgCshYlds, ProjectCoinsNotional)

Call DB.CloseDB(OpenedDB, VireoDBFileName)

Range("MPYield").Cells(1, 1).Value = MPYldRes * WorksheetFunction.Max(0.9, WorksheetFunction.Min(1.1, Range("VireoRatios").Cells(2, 1).Value / AvgCshYld))

For i = 5 To 4 + CsP
    If i > 4 + CP + Delay + 1 Then
        Worksheets("CF").Cells(56, i).Value = Round(4 * Worksheets("CF").Cells(54, i).Value / Coins.Nominal * (1 / Range("IssuanceRatio").Cells(1, 1).Value) * Range("MPYield").Cells(1, 1).Value, 4)
        AvgCropYield = AvgCropYield + Worksheets("CF").Cells(56, i).Value
    Else
        If i = 5 + CP + Delay Then
            Worksheets("CF").Cells(56, i).Value = Round(4 * Worksheets("CF").Cells(54, i).Value / Coins.Nominal * (1 / Range("IssuanceRatio").Cells(1, 1).Value) * Range("MPYield").Cells(1, 1).Value, 4)
            AvgCropYield = AvgCropYield + Worksheets("CF").Cells(56, i).Value
        End If
    End If
Next i

Worksheets("Summary").Range("VireoRatios").Cells(1, 1).Value = Round(AvgCropYield / (CsP - CP - Delay), 4)

'Computation of WACC
Dim tmpInst As Variant
Dim TotEq, TotDebt As Double
TotEq = 0#
TotDebt = 0#

For Each tmpInst In Equities.Items
    TotEq = TotEq + tmpInst.Amount
Next

For Each tmpInst In Loans.Items
    TotDebt = TotDebt + tmpInst.Nominal
Next

Range("WACC").Cells(1, 1).Value = Functions.GetWACC(Coins.Nominal * Coins.Conv + TotEq, Range("WACCParam").Cells(1, 1).Value / 100, Coins.Nominal * (1 - Coins.Conv) + TotDebt, Range("WACCParam").Cells(2, 1).Value / 100, GetTaxRate(Worksheets("All Costs").TaxesBox.Value, CsP))

End Sub

Private Function AdjustCoinNominal() As Double()

Dim i As Integer
Dim Res() As Double: ReDim Res(CP + Delay)
Dim CoinIncrease As Double: CoinIncrease = 0#

For i = 5 To 4 + CP + Delay
    CoinIncrease = CoinIncrease + Worksheets("CF").Cells(38, i).Value
    Res(i - 4) = Worksheets("CF").Cells(38, i).Value
Next i

Coins.ChangeNominal (Coins.Nominal + CoinIncrease)
Range("TotalCoinNotional").Value = Worksheets("Financing").Cells(16, 3).Value + CoinIncrease

AdjustCoinNominal = Res

End Function
Private Sub RoyaltiesAndTaxes(Revenues() As Double)

Dim Royalties() As Double: Royalties = Tools.GetCost("Royalties", Revenues)
Dim Depreciation As Double: Depreciation = 0#
Dim i As Integer

If Worksheets("Param").IncludeContCost.Value Then
    For i = 5 To 4 + CP + Delay
        Depreciation = Depreciation + Worksheets("Constr CF").Cells(i, 3).Value * (1 + Range("CapexInc").Cells(1, 1).Value / 100) + Worksheets("Constr CF").Cells(i, 4).Value
    Next i
Else
    For i = 5 To 4 + CP + Delay
        Depreciation = Depreciation + Worksheets("Constr CF").Cells(i, 3).Value * (1 + Range("CapexInc").Cells(1, 1).Value / 100)
    Next i
End If


Depreciation = Math.Round(Depreciation / (4 * WorksheetFunction.Max(Range("Depreciation").Cells(1, 1).Value, Range("ConcPeriod").Cells(1, 1).Value)), 2)

Dim Taxes() As Double: Taxes = Tools.GetCost("Taxes", Revenues, Depreciation)

For i = 5 To 4 + CsP
    If i > 4 + Delay + CP Then
        Worksheets("CF").Cells(43, i).Value = Depreciation
        Worksheets("CF").Cells(44, i).Value = Royalties(i - (4 + CP + Delay))
        Worksheets("CF").Cells(45, i).Value = Taxes(i - (4 + CP + Delay))
    Else
        Worksheets("CF").Cells(43, i).Value = 0#
        Worksheets("CF").Cells(44, i).Value = 0#
        Worksheets("CF").Cells(45, i).Value = 0#
    End If
    
    Worksheets("CF").Cells(46, i).Value = Worksheets("CF").Cells(44, i).Value + Worksheets("CF").Cells(45, i).Value
    Worksheets("CF").Cells(48, i).Value = Worksheets("CF").Cells(40, i).Value - Worksheets("CF").Cells(46, i).Value
Next i

End Sub

Private Function CoverInterests() As Double()

Dim WK As Double: WK = WorksheetFunction.RoundUp(Range("WorkK").Value / 4, 2)
Dim i, IRCovPer As Integer: IRCovPer = Range("InterestCovPeriod").Cells(1, 1).Value
Dim CoinAddNom As Double: CoinAddNom = 0#
Dim AccInt As Double: AccInt = 0#
Dim Res() As Double

ReDim Res(CsP - Delay - CP)

For i = 5 To 5 + Delay + CP + IRCovPer
    If Worksheets("CF").Cells(35, i).Value < 0 Then
        Dim RT As Double: RT = 1 + 1 / 2 ^ 6 * (i - 5) / 4
        Dim LocalInt As Double: LocalInt = WorksheetFunction.RoundUp((WK + Math.Abs(Worksheets("CF").Cells(35, i).Value)) * (Exp(Coins.RateCP / 4 * 1.01 * RT) - 1), 2)
        Dim LocalNomInc As Double: LocalNomInc = WorksheetFunction.RoundUp((Math.Abs(Worksheets("CF").Cells(35, i).Value) + WK + AccInt) * Exp(Coins.RateCP / 4 * 1.01 * RT), 2)
        CoinAddNom = CoinAddNom + LocalNomInc
        AccInt = AccInt + LocalInt
        Worksheets("CF").Cells(38, i).Value = LocalNomInc
        Worksheets("CF").Cells(39, i).Value = WorksheetFunction.RoundUp(CoinAddNom * Coins.RateCP / 4, 2)
        Worksheets("CF").Cells(40, i).Value = Worksheets("CF").Cells(35, i).Value + Worksheets("CF").Cells(38, i).Value - Worksheets("CF").Cells(39, i).Value
    End If
    If i > 5 + Delay + CP - 1 Then Res(i - (5 + Delay + CP - 1)) = Worksheets("CF").Cells(40, i).Value
Next i

For i = 5 + Delay + CP + IRCovPer To 4 + CsP
    Worksheets("CF").Cells(40, i).Value = Worksheets("CF").Cells(35, i).Value + Worksheets("CF").Cells(38, i).Value - Worksheets("CF").Cells(39, i).Value
    Res(i - (5 + Delay + CP - 1)) = Worksheets("CF").Cells(40, i).Value
Next i

CoverInterests = Res

End Function

Private Sub GetEBIT()

Dim i, j As Integer
Dim Revenues() As Double: Revenues = GetProduction()
Dim OM() As Double: OM = Tools.GetCost("O&M", Revenues, Delay:=Delay)
Dim SGA() As Double: SGA = Tools.GetCost("SG&A", Revenues, Delay:=Delay)
Dim ProjectSize As Double: ProjectSize = 0
AvgOM = 0#
AvgSGA = 0#
AvgRevenues = 0#

If CP <> 0 Then
    For i = 1 To CP
        ProjectSize = ProjectSize + Worksheets("Constr CF").Cells(i + 3, 3).Value
    Next i
End If

Dim AuditFee As Double: AuditFee = Round(Range("AuditFee").Cells(1, 1).Value / 400 * ProjectSize, 2)

For i = 1 To UBound(Revenues)
    Worksheets("CF").Cells(17, 4 + i).Value = Revenues(i)
    AvgRevenues = AvgRevenues + Revenues(i)
    Worksheets("CF").Cells(18, 4 + i).Value = OM(i)
    AvgOM = AvgOM + OM(i)
    Worksheets("CF").Cells(19, 4 + i).Value = SGA(i)
    AvgSGA = AvgSGA + SGA(i)
    Worksheets("CF").Cells(20, 4 + i).Value = AuditFee
    Worksheets("CF").Cells(21, 4 + i).Value = Revenues(i) - OM(i) - SGA(i) - AuditFee
    
    If i < CP + Delay + 1 Then
        Dim CashIn As Double: CashIn = -Worksheets("CF").Cells(3, 4 + i).Value
        
        For j = 1 To 14
            CashIn = CashIn + Worksheets("CF").Cells(4 + j, 4 + i).Value
        Next j
    Else
        CashIn = 0#
    End If
    
    Worksheets("CF").Cells(35, 4 + i).Value = Worksheets("CF").Cells(21, 4 + i).Value + CashIn - Worksheets("CF").Cells(33, 4 + i).Value
Next i

End Sub

Private Function GetProduction() As Double()

Dim DegRisk(), ClimRisk(), PPA(), Res() As Double
Dim ws1 As Worksheet: Set ws1 = Worksheets("Deg Risk")
Dim ws2 As Worksheet: Set ws2 = Worksheets("Clim Risk")
Dim ws3 As Worksheet: Set ws3 = Worksheets("PPA")
Dim count As Integer: count = 1
Dim Period As String: Period = ws3.Cells(3 + count, 2).Value
'Dim OMImpact As Double: OMImpact = Range("OMImpact").Cells(1, 1).Value / 100
Dim i As Integer
ReDim DegRisk(800)
ReDim ClimRisk(800)
ReDim PPA(800)

While Period <> ""
    PPA(count) = ws3.Cells(3 + count, 3).Value
    count = count + 1
    Period = ws3.Cells(3 + count, 2).Value
Wend

count = count - 1
ReDim Preserve PPA(count)

count = 1
Period = ws1.Cells(3 + count, 8).Value

While Period <> ""
    DegRisk(count) = 1 - (1 - ws1.Cells(3 + count, 9).Value) ' * (1 - OMImpact)
    ClimRisk(count) = ws2.Cells(3 + count, 9).Value
    count = count + 1
    Period = ws1.Cells(3 + count, 8).Value
Wend

count = count - 1
ReDim Preserve DegRisk(count)
ReDim Preserve ClimRisk(count)
ReDim Res(count)

Dim MaxProd As Double: MaxProd = Range("PowerProd")(1, 1).Value * Range("PlantF")(1, 1).Value / 100 * 365 * 24 * (1 - (Range("Losses").Cells(1, 1).Value + Range("Losses").Cells(2, 1).Value) / 100)
MaxProd = MaxProd / 4

For i = 1 To count
    Res(i) = Round(DegRisk(i) * ClimRisk(i) * MaxProd, 2)
    If i > CP + Delay Then Res(i) = Round(Res(i) * Tools.GetProdIncRate(i - CP - Delay) * PPA(i - CP - Delay), 2)
Next i

GetProduction = Res

End Function
Private Sub DebtRepaymentAndInterest()

Dim tmpLoan As Variant
Dim i, j, k, l As Integer

'Equities
Dim AccOwnerEquity, AccOtherEquity As Double
AccOwnerEquity = 0#
AccOtherEquity = 0#

For i = 5 To 4 + CP + Delay
    AccOwnerEquity = AccOwnerEquity + Worksheets("CF").Cells(5, i).Value
    AccOtherEquity = AccOtherEquity + Worksheets("CF").Cells(6, i).Value
    Worksheets("CF").Cells(30, i).Value = Round(AccOwnerEquity * Equities("Project Owner").ROE() / 4 + AccOtherEquity * Equities("Others").ROE() / 4, 2)
Next i

'Loans
For Each tmpLoan In Loans.Items
    Dim RepaymentAmount, RepaidNominal As Double
    
    For i = 8 To 12
        If Worksheets("CF").Cells(i, 4).Value = tmpLoan.Name() Then
            For j = 5 To 4 + CP + Delay
                If Worksheets("CF").Cells(i, j).Value <> 0 Then
                    Dim AccNom As Double: AccNom = 0#
                    Dim Interval As Integer: Interval = 4 * tmpLoan.Freq() / 12
                    
                    'For k = j To j + tmpLoan.GP() - 1 + Interval + IIf(tmpLoan.AllAtGP(), -1, 0)
                    For k = j To j + tmpLoan.GP() - 1 + IIf(tmpLoan.AllAtGP(), 1, 0)
                        AccNom = AccNom + Worksheets("CF").Cells(i, k).Value
                        'If Worksheets("CF").Cells(i, k).Value = 0 Then Worksheets("CF").Cells(27, k).Value = Worksheets("CF").Cells(27, k).Value + AccNom
                        If k < j + tmpLoan.GP() Then Worksheets("CF").Cells(31, k).Value = Worksheets("CF").Cells(31, k).Value + WorksheetFunction.RoundUp(AccNom * tmpLoan.Rate() / 4, 2)
                    Next k
                    
                    RepaymentAmount = Interval * GetRepaymentNom(AccNom, tmpLoan.Tenor(), tmpLoan.GP())
                    RepaidNominal = 0
                    
                    Dim Adj As Integer: Adj = Functions.IsPair(j + tmpLoan.GP + IIf(tmpLoan.AllAtGP(), 1, 0))
                    If j + tmpLoan.GP >= 5 + CP + Delay Then Adj = 0
                    
                    For k = j + tmpLoan.GP() + Interval - 1 To j + tmpLoan.Tenor() Step Interval
                        If k > j + tmpLoan.Tenor() - Interval Then
                            For l = j + tmpLoan.Tenor() - Interval + Adj To j + tmpLoan.Tenor() - 1
                                Worksheets("CF").Cells(31, l).Value = Worksheets("CF").Cells(31, l).Value + WorksheetFunction.RoundUp(AccNom * tmpLoan.Rate() * Interval / 4, 2)
                                AccNom = AccNom - RepaymentAmount / Interval
                                
                                Worksheets("CF").Cells(27, l).Value = Worksheets("CF").Cells(27, l).Value - RepaidNominal
                                RepaidNominal = RepaidNominal + RepaymentAmount / Interval
                            Next l
                        Else
                            'Compute interest
                            Worksheets("CF").Cells(31, k).Value = Worksheets("CF").Cells(31, k).Value + WorksheetFunction.RoundUp(AccNom * tmpLoan.Rate() * Interval / 4, 2)
                            AccNom = AccNom - RepaymentAmount
                        
                            Worksheets("CF").Cells(27, k).Value = Worksheets("CF").Cells(27, k).Value - RepaidNominal
                            For l = k - 1 To k - Interval + 1 Step -1
                                Worksheets("CF").Cells(27, l).Value = Worksheets("CF").Cells(27, l).Value - RepaidNominal
                            Next l
                            
                            RepaidNominal = RepaidNominal + RepaymentAmount
                        End If
                        
                        'Round to 0 if remaining nominal is smaller than 10 and move to next loan
                        If Worksheets("CF").Cells(27, k).Value <= 10 Then
                            Worksheets("CF").Cells(27, k).Value = 0
                            GoTo NextLoan
                        End If
                    Next k
                    
                    GoTo NextLoan
                End If
            Next j
        End If
    Next i
NextLoan:
Next

'Coins
AccNom = 0#
Interval = 4 * Coins.Freq() / 12

For i = 5 To 4 + CP + Delay + 1
    AccNom = AccNom + Worksheets("CF").Cells(14, i).Value
    Worksheets("CF").Cells(32, i).Value = Round(AccNom * Coins.RateCP() / 4, 2)
Next i

AccNom = Round(Coins.Nominal() * (1 - Coins.Conv()), 2)
RepaymentAmount = Interval * GetRepaymentNom(AccNom, Coins.DebtTenor(), CP + Delay)

For i = 4 + CP + Delay + Interval To 4 + Coins.DebtTenor() Step Interval
    Worksheets("CF").Cells(28, i).Value = AccNom
    
    For j = i To i - Interval + 1 Step -1
        Worksheets("CF").Cells(28, j).Value = AccNom
    Next j
    
    Worksheets("CF").Cells(32, i).Value = Round(AccNom * Coins.RateDebt() / 4, 2)
    AccNom = AccNom - RepaymentAmount
Next i

End Sub

Private Function GetRepaymentNom(ByVal Nominal As Double, ByVal Tenor As Integer, ByVal GP As Integer) As Double

GetRepaymentNom = WorksheetFunction.RoundUp(Nominal / (Tenor - GP), 2)

End Function

Private Sub GetOutstandingNomAndCF()

Dim tmpLoan As Variant
Dim i, j, count As Integer

count = 1

For Each tmpLoan In Loans.Items
    For i = 8 To 12
        If Worksheets("CF").Cells(i, 4).Value = tmpLoan.Name() Then
            Dim NonWithNom As Double: NonWithNom = tmpLoan.Nominal()
            Dim OutStanding As Double: OutStanding = 0#
            Dim FirstWithdraw As Integer: FirstWithdraw = 4 + tmpLoan.StartDate() - 1
            
            If Worksheets("CF").Cells(i, FirstWithdraw).Value = tmpLoan.Name() Then
                FirstWithdraw = FirstWithdraw + 1
            Else
                While Worksheets("CF").Cells(i, FirstWithdraw).Value = ""
                    FirstWithdraw = FirstWithdraw + 1
                    If FirstWithdraw > 4 + CP + Delay Then GoTo NextLoan
                Wend
            End If
            
            Dim LastWithdraw As Integer: LastWithdraw = WorksheetFunction.Max(4 + CP + Delay, FirstWithdraw + tmpLoan.GP())
            
            For j = FirstWithdraw To LastWithdraw
                NonWithNom = NonWithNom - Worksheets("CF").Cells(i, j).Value
                OutStanding = OutStanding + Worksheets("CF").Cells(i, j).Value
                If OutStanding = 0 Then FirstWithdraw = FirstWithdraw + 1
                Worksheets("CF").Cells(25, j).Value = Worksheets("CF").Cells(25, j).Value + Round(tmpLoan.CF() / 4 * NonWithNom, 2)
                Worksheets("CF").Cells(27, j).Value = Worksheets("CF").Cells(27, j).Value + OutStanding
            Next j
            
            For j = LastWithdraw + 1 To FirstWithdraw + tmpLoan.Tenor() - 1
                Worksheets("CF").Cells(27, j).Value = Worksheets("CF").Cells(27, j).Value + OutStanding
            Next j
            
        End If
    Next i
NextLoan:
count = count + 1
Next

OutStanding = 0#

For i = 5 To 4 + CP + Delay
    OutStanding = OutStanding + Worksheets("CF").Cells(14, i).Value
    
    Worksheets("CF").Cells(28, i).Value = Worksheets("CF").Cells(28, i).Value + OutStanding
Next i

End Sub

Private Sub GetUF()

Dim i, j As Integer
Dim tmpLoan As Variant

For Each tmpLoan In Loans.Items
    For i = 8 To 12
        If Worksheets("CF").Cells(i, 4).Value = tmpLoan.Name() Then
            For j = 5 To 4 + CP + Delay
                If Worksheets("CF").Cells(i, j).Value <> 0 Then
                    Worksheets("CF").Cells(24, j).Value = Worksheets("CF").Cells(24, j).Value + Round(tmpLoan.UF() * tmpLoan.Nominal(), 2)
                    GoTo NextLoan
                End If
            Next j
        End If
    Next i
NextLoan:
Next

For i = 5 To 4 + CP + Delay
    If Worksheets("CF").Cells(14, i).Value <> 0 Then
        Worksheets("CF").Cells(24, i).Value = Worksheets("CF").Cells(24, i).Value + Round(Coins.UF() * Coins.Nominal(), 2)
        Exit Sub
    End If
Next i

End Sub

Private Sub CheckLoanGP()

Dim tmpLoan As Variant
Dim i, j, k As Integer: i = 8

For Each tmpLoan In Loans.Items
    Dim GP As Integer: GP = tmpLoan.GP()
    Dim TotalRemaining As Double: TotalRemaining = 0#
    Dim TotalLoan As Double: TotalLoan = 0#
    
    For j = 5 To 4 + CP + Delay
        If Worksheets("CF").Cells(i, j).Value <> 0 Then
            For k = j To j + GP - 1
                TotalLoan = TotalLoan + Worksheets("CF").Cells(i, k).Value
            Next k
            
            For k = j + GP To 4 + CP + Delay
                TotalLoan = TotalLoan + Worksheets("CF").Cells(i, k).Value
                TotalRemaining = TotalRemaining + Worksheets("CF").Cells(i, k).Value
                Worksheets("CF").Cells(i, k).Value = 0#
            Next k
            
            If TotalRemaining <> 0 Then Call tmpLoan.AllIsAtGP
            
            Worksheets("CF").Cells(i, WorksheetFunction.Min(4 + CP + Delay, j + GP)).Value = Worksheets("CF").Cells(i, WorksheetFunction.Min(4 + CP + Delay, j + GP)).Value + TotalRemaining
            If TotalLoan <> tmpLoan.Nominal() Then Range("LoanMsg")(i - 7, 1).Value = "Reduce Nominal"
            GoTo NextLoan
        End If
    Next j
    Range("LoanMsg")(i - 7, 1).Value = "Loan not used"
NextLoan:
    i = i + 1
Next

End Sub

Private Sub WithDrawCash(InterestCoverage() As Double)

Dim i, j As Integer
Dim Instruments(8) As String
Dim CashAtDisp(8, 2) As Double
Dim Disposable(8) As Boolean
Dim tmpLoan As Variant
Dim IsIntCov As Boolean: IsIntCov = Functions.IsArrayInitalized(InterestCoverage)

'Init the vectors
For i = 1 To 8
    Disposable(i) = False
Next i

Instruments(1) = "Project Owner"
CashAtDisp(1, 1) = Range("EquityData").Cells(1, 1).Value
CashAtDisp(1, 2) = Range("EquityData").Cells(1, 2).Value / 100
If CashAtDisp(1, 1) <> 0 Then Disposable(1) = True
Instruments(2) = "Others"
CashAtDisp(2, 1) = Range("EquityData").Cells(2, 1).Value
CashAtDisp(2, 2) = Range("EquityData").Cells(2, 2).Value / 100
If CashAtDisp(2, 1) <> 0 Then Disposable(2) = True

i = 3
For Each tmpLoan In Loans.Items
    Instruments(i) = tmpLoan.Name()
    CashAtDisp(i, 1) = tmpLoan.Nominal()
    CashAtDisp(i, 2) = tmpLoan.Rate()
    Disposable(i) = tmpLoan.IsAvailable(1)
    i = i + 1
Next

Instruments(8) = "Coins"
CashAtDisp(8, 1) = Coins.Nominal()
CashAtDisp(8, 2) = Coins.RateCP()
Disposable(8) = True

Call Functions.RankInstr(Instruments, CashAtDisp, Disposable)

For i = 1 To CP + Delay
    Dim tmpCF As Double
    If Not IsIntCov Then
        tmpCF = Worksheets("CF").Cells(3, 4 + i)
    Else
        tmpCF = Worksheets("CF").Cells(3, 4 + i) + InterestCoverage(i)
    End If
    
    For Each tmpLoan In Loans.Items
        If tmpLoan.IsAvailable(i) And Not tmpLoan.IsAvailable(i - 1) Then
            For j = 1 To 8
                If Instruments(j) = tmpLoan.Name() Then
                    Disposable(j) = True
                End If
            Next j
        End If
    Next
    
    For j = 1 To 8
        If Disposable(j) Then
            If CashAtDisp(j, 1) >= tmpCF Then
                CashAtDisp(j, 1) = CashAtDisp(j, 1) - tmpCF
                Call WriteInCF(Instruments(j), i, tmpCF)
                GoTo NextDate
            Else
                Call WriteInCF(Instruments(j), i, CashAtDisp(j, 1))
                tmpCF = tmpCF - CashAtDisp(j, 1)
                CashAtDisp(j, 1) = 0
            End If
        End If
    Next j
NextDate:
Next i

End Sub

Private Sub WriteInCF(ByVal Instru As String, ByVal Col As Integer, ByVal Amount As Double)

Dim i As Integer

For i = 5 To 14
    If Worksheets("CF").Cells(i, 4).Value = Instru Then
        Worksheets("CF").Cells(i, 4 + Col).Value = Amount
        Exit Sub
    End If
Next i

End Sub

Private Function InitCFTable() As Boolean

Dim i As Integer
Dim tmpLoan As Variant

If CsP > 16380 Then
    MsgBox ("Concession Period is too long")
    InitCFTable = False
    Exit Function
End If

For i = 1 To CsP
    Worksheets("CF").Cells(2, 4 + i).Value = "Q " & i
Next i

i = 1
For Each tmpLoan In Loans.Items
    Worksheets("CF").Cells(7 + i, 4).Value = tmpLoan.Name()
    i = i + 1
Next

For i = 1 To CP + Delay
    Worksheets("CF").Cells(3, 4 + i).Value = GetConstrCF(i, CP, CapexInc, Delay)
Next i
    
InitCFTable = True

End Function

Private Function GetConstrCF(ByVal Period As Integer, ByVal CP As Integer, Optional ByVal CapexIncrease As Double, Optional ByVal Delay As Integer) As Double

Dim TotalConstCost As Double: TotalConstCost = 0
Dim TotalConstCostDelay As Double: TotalConstCostDelay = 0
Dim Row, i As Integer: Row = 4
Dim tmpP As String: tmpP = Worksheets("Constr CF").Cells(Row, 2).Value

While tmpP <> ""
    TotalConstCost = TotalConstCost + Worksheets("Constr CF").Cells(Row, 3).Value
    Row = Row + 1
    tmpP = Worksheets("Constr CF").Cells(Row, 2).Value
Wend

Row = Row - 1

Dim tmpRes As Double

If Period + 3 < Row - 3 + WorksheetFunction.RoundDown(Delay / 2, 0) Then
    tmpRes = Worksheets("Constr CF").Cells(3 + Period, 3).Value + Worksheets("Constr CF").Cells(3 + Period, 4).Value
Else
    For i = Row - 3 + WorksheetFunction.RoundDown(Delay / 2, 0) To Row
        TotalConstCostDelay = TotalConstCostDelay + Worksheets("Constr CF").Cells(i, 3).Value + Worksheets("Constr CF").Cells(i, 4).Value
    Next i

    TotalConstCostDelay = TotalConstCostDelay + TotalConstCost * CapexIncrease

    tmpRes = Round(TotalConstCostDelay / (Delay + 3 - WorksheetFunction.RoundDown(Delay / 2, 0) + 1), 2)
End If

GetConstrCF = tmpRes

End Function
Private Function InitEquity() As Boolean

Dim tmpEquity As Equity
Dim eqData As Range: Set eqData = Range("EquityData")

On Error GoTo InitEquityError

Set tmpEquity = New Equity
Set Equities = New Dictionary

tmpEquity.Init EquityData:=eqData, Row:=1
Equities.Add Key:="Project Owner", Item:=tmpEquity

Set tmpEquity = New Equity
tmpEquity.Init EquityData:=eqData, Row:=2
Equities.Add Key:="Others", Item:=tmpEquity

InitEquity = True

Exit Function

InitEquityError:
Debug.Print "Error in InitEquity function"
InitEquity = False

End Function

Private Function InitLoans() As Boolean

Dim i As Integer
Dim tmpLoan As Loan

On Error GoTo InitLoansError

Dim loanData As Range: Set loanData = Range("LoanData")
Set Loans = New Dictionary

For i = 1 To loanData.Rows.count
    Set tmpLoan = New Loan
    tmpLoan.Init Data:=loanData, Row:=i
    If tmpLoan.Name <> "" Then Loans.Add Key:=tmpLoan.Name, Item:=tmpLoan
Next i

InitLoans = True

Exit Function

InitLoansError:
Debug.Print "Error in InitLoans function"
InitLoans = False

End Function

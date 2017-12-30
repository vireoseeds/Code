Option Explicit
Option Base 1


Public Sub YieldCurveSmoother(IValue As Range)
'
' Crops Yield Curve Smoothing Macro
'
    Dim yc, Result As Variant
    Dim Grad, Convex As Range
    Dim minGrad, maxGrad, Conv As Double
    Dim i, nYrs As Integer
    
    Application.ScreenUpdating = False
    Worksheets("Graph Data").Activate
    
    'yc = Range(Range("YCurveRef"), Range("YCurveRef").End(xlDown))
    minGrad = Range("SmoothingParam").Cells(1, 1).Value / 100
    maxGrad = Range("SmoothingParam").Cells(2, 1).Value / 100
    Conv = Range("SmoothingParam").Cells(3, 1).Value / 100
    
    nYrs = IValue.Columns.count
    
    Set Grad = Worksheets("Graph Data").Range(Worksheets("Graph Data").Cells(15, 2), Worksheets("Graph Data").Cells(15 + nYrs - 2, 2))
    Set Convex = Worksheets("Graph Data").Range(Worksheets("Graph Data").Cells(16, 3), Worksheets("Graph Data").Cells(16 + nYrs - 2, 3))
    
    For i = 1 To nYrs - 1
        Grad.Cells(i, 1).Formula = "=" & IValue.Cells(1, i + 1).Address & "-" & IValue.Cells(1, i).Address
        If i > 1 Then Convex.Cells(i - 1, 1).Formula = "=" & Grad.Cells(i, 1).Address & "-" & Grad.Cells(i - 1, 1).Address
    Next i
    
    Application.Run "SolverReset"
    For i = 1 To nYrs - 1
        'Gradient Constraints
        Application.Run "SolverAdd", Grad.Cells(i, 1).Address, 3, Application.WorksheetFunction.Text(minGrad * 1, "0.000%")
        Application.Run "SolverAdd", Grad.Cells(i, 1).Address, 1, Application.WorksheetFunction.Text(maxGrad * 1, "0.000%")
        If i < nYrs - 1 Then
            'Convexity Constraints ensure -ve convexity
            Application.Run "SolverAdd", Convex.Cells(i, 1).Address, 1, Application.WorksheetFunction.Text(Conv * 1, "0.000%")
        End If
    Next i
    
    Worksheets("Graph Data").Range("$A$100").Formula = "=AVERAGE($B$3:" & IValue.Cells(1, nYrs).Address & ")"
    
    'Global condition
    Dim AvgObj As Double: AvgObj = Functions.Average(IValue)
    Application.Run "SolverAdd", "$A$100", 2, Application.WorksheetFunction.Text(AvgObj, "0.000%")
    Application.Run "SolverOk", IValue.Cells(1, 1).Address, 1, 0, IValue.Cells(1, 1).Address & ":" & IValue.Cells(1, nYrs).Address, 1, "GRG Nonlinear"
    Result = Application.Run("Solver.xlam!SolverSolve", True)

    ' finish the analysis
    Application.Run "Solver.xlam!SolverFinish"

    Worksheets("Summary").Activate
    Application.ScreenUpdating = True
    
    Grad.ClearContents
    Convex.ClearContents
    Worksheets("Graph Data").Range("$A$100").ClearContents
    
    ' report on success of analysis
    If Result <= 3 Then
        ' Result = 0, Solution found, optimality and constraints satisfied
        ' Result = 1, Converged, constraints satisfied
        ' Result = 2, Cannot improve, constraints satisfied
        ' Result = 3, Stopped at maximum iterations
        ActiveSheet.Range("VireoRatios").Cells(1, 1).Interior.ColorIndex = 4
    Else
        ' Result = 4, Solver did not converge
        ' Result = 5, No feasible solution
        Beep
        ActiveSheet.Range("VireoRatios").Cells(1, 1).Interior.ColorIndex = 3
    End If
    
    
End Sub
 
Private Function ComputeFlatYield(DF As Variant, InitGuess As Double, BondPrice As Double) As Double

    Dim PVCF() As Variant
    Dim PVSum, sumDF, BPrice, Guess, tol As Double
    Dim i, iter, nYrs As Integer
    DF = Application.WorksheetFunction.Transpose(DF)
    nYrs = UBound(DF)
    Guess = InitGuess
    sumDF = Application.WorksheetFunction.Sum(DF)
    tol = 0.000001
    iter = 0
    
    PVSum = sumPV(DF, Guess * 1)
    
    Do While Abs(BondPrice - PVSum) > tol
        'Update PVCF for new guess
        Guess = Guess + ((BondPrice - PVSum) / sumDF)
        PVSum = sumPV(DF, Guess * 1)
        iter = iter + 1
        If iter = 100 Then
            Exit Do
        End If
        
    Loop
    
    ComputeFlatYield = Guess
End Function

Function sumPV(DF As Variant, Cpn As Double) As Double
    
    Dim i As Integer
    For i = 1 To UBound(DF) - 1
        sumPV = sumPV + Cpn * DF(i)
    Next i
    sumPV = sumPV + (1 + Cpn) * DF(UBound(DF))

End Function

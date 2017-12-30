Option Explicit
Option Base 1

Public Sub GetRiskIndicator()

Dim savedData As Double
Dim i, idx As Integer
Dim RiskType As String
Dim AvgCropsYield0, AvgCropsYield2 As Double
Dim SD As Double
Dim VireoDBFileName As String: VireoDBFileName = "Vireo_DB.xlsm"
Dim VireoDBFilePathName As String: VireoDBFilePathName = ActiveWorkbook.Path & "\"
Dim ProjectName As String: ProjectName = Range("ProjectName").Value
Dim ProjectNames As Variant

Application.ScreenUpdating = False

AvgCropsYield0 = Range("VireoRatios").Cells(1, 1).Value
AvgCropsYield2 = WorksheetFunction.Power(AvgCropsYield0, 2)
SD = AvgCropsYield2

Dim OpenedDB As Object: Set OpenedDB = DB.OpenDB(VireoDBFilePathName + VireoDBFileName)
Dim OpenedDBSource As Worksheet: Set OpenedDBSource = DB.GetDBSource(OpenedDB, "VireoDB")
Set ProjectNames = OpenedDBSource.Range("ProjNames")

idx = Application.Match(ProjectName, ProjectNames, 1)

'First Risk is the plant factor: move of +/-X%
Dim PFStressUp As Double: PFStressUp = OpenedDBSource.Range("VDBRef").Offset(idx, 11).Value
savedData = Range("PlantF").Value
Range("PlantF").Value = savedData * PFStressUp

Call PVal.Launch(True)
SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2

Dim PFStressDn As Double: PFStressDn = OpenedDBSource.Range("VDBRef").Offset(idx, 12).Value
Range("PlantF").Value = savedData * PFStressDn

Call PVal.Launch(True)
SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2

Range("PlantF").Value = savedData

'Include Delay and Capex
Worksheets("Param").IncludeConstRisk.Value = True

Call PVal.Launch(True)
SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2

Worksheets("Param").IncludeConstRisk.Value = False

'Impact of Degradation: minimum degradation of 0.25% per year and x2 Degradation Rate
RiskType = Worksheets("Deg Risk").DegradationBox.Value

savedData = Range("LinearDeg").Value
Dim MinDeg As Double: MinDeg = OpenedDBSource.Range("VDBRef").Offset(idx, 13).Value
Dim MaxDeg As Double: MaxDeg = OpenedDBSource.Range("VDBRef").Offset(idx, 14).Value
Range("LinearDeg").Value = MinDeg
Worksheets("Deg Risk").DegradationBox.Value = "Linear"
Call PVal.Launch(True)
SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2

Range("LinearDeg").Value = savedData
Worksheets("Deg Risk").DegradationBox.Value = RiskType
        
Select Case RiskType
    Case "Linear":
        Range("LinearDeg").Value = MaxDeg * savedData
        Call PVal.Launch(True)
        SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2
        
        Range("LinearDeg").Value = savedData
    Case "Multi-Linear"
        For i = 1 To Range("MultiDeg").Rows.count
            Range("MultiDeg").Cells(i, 3).Value = MaxDeg * Range("MultiDeg").Cells(i, 3).Value
        Next i
        
        Call PVal.Launch(True)
        SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2
        
        For i = 1 To Range("MultiDeg").Rows.count
            If Range("MultiDeg").Cells(i, 3).Value = 0 Then
                Range("MultiDeg").Cells(i, 3).Value = ""
            Else
                Range("MultiDeg").Cells(i, 3).Value = Range("MultiDeg").Cells(i, 3).Value / 2
            End If
        Next i
    Case "Stepped"
        For i = 1 To Range("SteppedDeg").Rows.count
            Range("SteppedDeg").Cells(i, 3).Value = MaxDeg * Range("SteppedDeg").Cells(i, 3).Value
        Next i
        
        Call PVal.Launch(True)
        SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2
        
        For i = 1 To Range("SteppedDeg").Rows.count
            If Range("SteppedDeg").Cells(i, 3).Value = 0 Then
                Range("SteppedDeg").Cells(i, 3).Value = ""
            Else
                Range("SteppedDeg").Cells(i, 3).Value = Range("SteppedDeg").Cells(i, 3).Value / 2
            End If
        Next i
End Select

'Impact of Climate Risk: minimum risk to 0% per year and max to x2 Climate Rate
RiskType = Worksheets("Clim Risk").ClimateBox.Value

savedData = Range("LinearClim").Value
Dim ClimMin As Double: ClimMin = OpenedDBSource.Range("VDBRef").Offset(idx, 15).Value
Dim ClimMax As Double: ClimMax = OpenedDBSource.Range("VDBRef").Offset(idx, 16).Value

Range("LinearClim").Value = ClimMin
Worksheets("Clim Risk").ClimateBox.Value = "Linear"
Call PVal.Launch(True)
SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2

Range("LinearClim").Value = savedData
Worksheets("Clim Risk").ClimateBox.Value = RiskType
        
Select Case RiskType
    Case "Linear":
        Range("LinearClim").Value = ClimMax * savedData
        Call PVal.Launch(True)
        SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2
        
        Range("LinearClim").Value = savedData
    Case "Multi-Linear"
        For i = 1 To Range("MultiClim").Rows.count
            Range("MultiClim").Cells(i, 3).Value = ClimMax * Range("MultiClim").Cells(i, 3).Value
        Next i
        
        Call PVal.Launch(True)
        SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2
        
        For i = 1 To Range("MultiClim").Rows.count
            If Range("MultiClim").Cells(i, 3).Value = 0 Then
                Range("MultiClim").Cells(i, 3).Value = ""
            Else
                Range("MultiClim").Cells(i, 3).Value = Range("MultiClim").Cells(i, 3).Value / 2
            End If
        Next i
    Case "Stepped"
        For i = 1 To Range("SteppedClim").Rows.count
            Range("SteppedClim").Cells(i, 3).Value = ClimMax * Range("SteppedClim").Cells(i, 3).Value
        Next i
        
        Call PVal.Launch(True)
        SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2
        
        For i = 1 To Range("SteppedClim").Rows.count
            If Range("SteppedClim").Cells(i, 3).Value = 0 Then
                Range("SteppedClim").Cells(i, 3).Value = ""
            Else
                Range("SteppedClim").Cells(i, 3).Value = Range("SteppedClim").Cells(i, 3).Value / 2
            End If
        Next i
    Case "Cyclic Collapse":
        Range("CycColClim").Cells(1, 1).Value = ClimMax * Range("CycColClim").Cells(1, 1).Value
        Range("CycColClim").Cells(4, 1).Value = ClimMax * Range("CycColClim").Cells(4, 1).Value
        
        Call PVal.Launch(True)
        SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2
        
        Range("CycColClim").Cells(1, 1).Value = Range("CycColClim").Cells(1, 1).Value / 2
        Range("CycColClim").Cells(4, 1).Value = Range("CycColClim").Cells(4, 1).Value / 2
    Case "Cyclic Expansion":
        Range("CycExpClim").Cells(1, 1).Value = ClimMax * Range("CycExpClim").Cells(1, 1).Value
        Range("CycExpClim").Cells(4, 1).Value = ClimMax * Range("CycExpClim").Cells(4, 1).Value
        
        Call PVal.Launch(True)
        SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2
        
        Range("CycExpClim").Cells(1, 1).Value = Range("CycExpClim").Cells(1, 1).Value / 2
        Range("CycExpClim").Cells(4, 1).Value = Range("CycExpClim").Cells(4, 1).Value / 2
    Case "Cyclic Curv":
        Range("CycCurClim").Cells(1, 1).Value = ClimMax * Range("CycCurClim").Cells(1, 1).Value
        Range("CycCurClim").Cells(3, 1).Value = ClimMax * Range("CycCurClim").Cells(3, 1).Value
        Range("CycCurClim").Cells(6, 1).Value = ClimMax * Range("CycCurClim").Cells(6, 1).Value
        
        Call PVal.Launch(True)
        SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2
        
        Range("CycCurClim").Cells(1, 1).Value = Range("CycCurClim").Cells(1, 1).Value / 2
        Range("CycCurClim").Cells(3, 1).Value = Range("CycCurClim").Cells(3, 1).Value / 2
        Range("CycCurClim").Cells(6, 1).Value = Range("CycCurClim").Cells(6, 1).Value / 2
End Select

'Impact of O&M: Increase by X% of O&M floor
savedData = Range("OMFloor").Value
Dim OMFlrStress As Double: OMFlrStress = OpenedDBSource.Range("VDBRef").Offset(idx, 17).Value
Range("OMFloor").Value = Range("OMFloor").Value * OMFlrStress

Call PVal.Launch(True)
SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2

Range("OMFloor").Value = savedData

'Impact of SG&A: Increase by X% of SG&A floor
savedData = Range("SGAFloor").Value
Dim SGAFlrStress As Double: SGAFlrStress = OpenedDBSource.Range("VDBRef").Offset(idx, 18).Value
Range("SGAFloor").Value = Range("SGAFloor").Value * SGAFlrStress

Call PVal.Launch(True)
SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2

Range("SGAFloor").Value = savedData

'Impact of energy losses: Increase by X% of estimated losses
Dim Loss1, Loss2 As Double
Loss1 = Range("Losses").Cells(1, 1).Value
Loss2 = Range("Losses").Cells(2, 1).Value

Dim Loss1Stress As Double: Loss1Stress = OpenedDBSource.Range("VDBRef").Offset(idx, 19).Value
Dim Loss2Stress As Double: Loss2Stress = OpenedDBSource.Range("VDBRef").Offset(idx, 20).Value

Range("Losses").Cells(1, 1).Value = Loss1Stress * Loss1
Range("Losses").Cells(2, 1).Value = Loss2Stress * Loss2

Call PVal.Launch(True)
SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2

'Range("Losses").Cells(1, 1).Value = 0.8 * Loss1
'Range("Losses").Cells(2, 1).Value = 0.8 * Loss2

'Call PVal.Launch(True)
'SD = SD + WorksheetFunction.Power(Range("VireoRatios").Cells(1, 1).Value, 2) - AvgCropsYield2

Range("Losses").Cells(1, 1).Value = Loss1
Range("Losses").Cells(2, 1).Value = Loss2

Call DB.CloseDBNoSave(OpenedDB, VireoDBFileName)

SD = Math.Sqr(SD / 11)

Range("RiskIndicator").Value = Round(1 / (1 + AvgCropsYield0 / SD), 2)

Application.ScreenUpdating = True

Call PVal.Launch(True)

End Sub

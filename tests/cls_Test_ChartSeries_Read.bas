Attribute VB_Name = "cls_Test_ChartSeries_Read"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider


'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    'otherwise currently most of the tests fail
    ThisWorkbook.Activate
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'==============================================================================
'unit tests for 'ChartSeries' -- read stuff
'==============================================================================

'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_SurfacePlot_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaSurface")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_SurfacePlot_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaSurface")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_SurfacePlot_ReturnsEmpty()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaSurface")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eEmpty
    Const aExpectedValue As String = vbNullString
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.FormulaPart
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_SurfacePlot_ReturnsEmpty()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaSurface")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eEmpty
    Const aExpectedValue As String = vbNullString
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.FormulaPart
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("ChartSeriesBubbleSizes")
Public Sub ChartSeriesBubbleSizes_NoBubbleChart_ReturnsError()
    Const ExpectedError As Long = vbObjectError + 102
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 1
    '==========================================================================
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .BubbleSizes.EntryType
        ActualValue = .BubbleSizes.FormulaPart
    End With
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_NoSpaceWithNameAllRanges_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_NoSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C4:C7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_NoSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "A4:A7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_NoSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_NoSpaceWithArrayValues_ReturnsArray()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 3
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eArray
    Const aExpectedValue As String = "{1.5,2.5,3.5,4.5}"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.FormulaPart
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_NoSpaceWithNoXValues_ReturnsEmptyString()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 3
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eEmpty
    Const aExpectedValue As String = vbNullString
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.FormulaPart
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_NoSpaceWithString_ReturnsString()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 3
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eString
    Const aExpectedValue As String = "just a test, with a comma"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.CleanFormulaPart
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_NoSpaceWithNameFourAreas_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_NoSpaceWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C4,C5,C6,C7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_NoSpaceWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "A4,A5,A6,A7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_NoSpaceWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_NoSpaceWithNameTwoAreas_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_NoSpaceWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C4,C5:C6,C7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_NoSpaceWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "A4,A5:A6,A7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_NoSpaceWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_WithSpaceWithNameAllRanges_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblWithSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_WithSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblWithSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C4:C7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_WithSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblWithSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "A4:A7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_WithSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblWithSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_SpaceCommaWithNameAllRanges_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_SpaceCommaWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C4:C7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_SpaceCommaWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "A4:A7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_SpaceCommaWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_NoSpaceWithStringTitleContainingComma_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 4
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "B4:B7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_RoundBracketsWithNameAllRanges_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_RoundBracketsWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C4:C7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_RoundBracketsWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "A4:A7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_RoundBracketsWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_RoundBracketsWithArrayValues_ReturnsArray()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 3
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eArray
    Const aExpectedValue As String = "{1.5,2.5,3.5,4.5}"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.FormulaPart
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_RoundBracketsWithString_ReturnsString()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 3
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eString
    Const aExpectedValue As String = "just a test, with a comma"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.CleanFormulaPart
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_RoundBracketsWithNameFourAreas_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_RoundBracketsWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C4,C5,C6,C7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_RoundBracketsWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "A4,A5,A6,A7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_RoundBracketsWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_RoundBracketsWithNameTwoAreas_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_RoundBracketsWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C4,C5:C6,C7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_RoundBracketsWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "A4,A5:A6,A7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod("MaxName")
Public Sub ChartSeriesYValues_MaxNameWithLongArrayLong_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblMaxName
    Set cha = wks.ChartObjects("chaMultipleAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = _
            "AAC1000001:AAC1000002,AAC1000005:AAC1000006,AAC1000009:AAC1000010," & _
            "AAC1000013:AAC1000014,AAC1000017:AAC1000018,AAC1000021:AAC1000022," & _
            "AAC1000025:AAC1000026,AAC1000029:AAC1000030,AAC1000033:AAC1000034," & _
            "AAC1000037:AAC1000038,AAC1000041:AAC1000042,AAC1000045:AAC1000046," & _
            "AAC1000049:AAC1000050"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("MaxName")
Public Sub ChartSeriesYValues_MaxNameWithLongArrayShort_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblMaxName
    Set cha = wks.ChartObjects("chaMultipleAreas")
    Const ciSeries As Long = 3
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = _
            "AAC1000001:AAC1000002,AAC1000005:AAC1000006,AAC1000009:AAC1000010," & _
            "AAC1000013:AAC1000014,AAC1000017:AAC1000018,AAC1000021:AAC1000022," & _
            "AAC1000025:AAC1000026,AAC1000029:AAC1000030,AAC1000033:AAC1000034," & _
            "AAC1000037:AAC1000038,AAC1000041:AAC1000042"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("ChartSeriesYValues")
Public Sub ChartSeriesYValues_NamedRange_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eEntryType
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblWithSpace
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 3
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eDefinedName
    Const aExpectedValue As String = "'Space, Comma'!wks_y2"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .Values.EntryType
        ActualValue = .Values.FormulaPart
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesXValues")
Public Sub ChartSeriesXValues_NamedRange_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eEntryType
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblWithSpace
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 3
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eDefinedName
    Const aExpectedValue As String = "cls_Test_ChartSeries.xlsm!wkb_y1"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .XValues.EntryType
        ActualValue = .XValues.FormulaPart
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("ChartSeriesBubbleSizes")
Public Sub ChartSeriesBubbleSizes_NoSpaceWithNameAllRangesBubblePlot_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneAreaBubble")
    Const ciSeries As Long = 1
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C4:C7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .BubbleSizes.EntryType
        ActualValue = .BubbleSizes.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_NoSpaceWithNameAllRangesBubblePlot_ReturnsOne()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneAreaBubble")
    Const ciSeries As Long = 1
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 1
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod("ChartSeriesBubbleSizes")
Public Sub ChartSeriesBubbleSizes_SpaceCommaWithNameAllRangesBubblePlot_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaTwoAreasBubble")
    Const ciSeries As Long = 1
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "B4,B5:B6,B7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .BubbleSizes.EntryType
        ActualValue = .BubbleSizes.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesPlotOrder")
Public Sub ChartSeriesPlotOrder_SpaceCommaWithNameAllRangesBubblePlot_ReturnsOne()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim ActualValue As Long
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaTwoAreasBubble")
    Const ciSeries As Long = 1
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eInteger
    Const aExpectedValue As Long = 1
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .PlotOrder.EntryType
        ActualValue = .PlotOrder.Value
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("ChartSeriesSeriesName")
Public Sub ChartSeriesSeriesName_RoundBracketsWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim ActualType As eElement
    Dim rng As Range
    Dim ActualValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpectedType As eEntryType
    aExpectedType = eRange
    Const aExpectedValue As String = "C3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        ActualType = .SeriesName.EntryType
        ActualValue = .SeriesName.RangeString
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, ActualType
        .AreEqual aExpectedValue, ActualValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("ChartSeriesNoOfPointsY")
Public Sub ChartSeriesNoOfPointsY_NoSpaceWithNameAllRanges_ReturnsFour()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim iPoints As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    Const cElement As Long = 3
    '==========================================================================
    Const aExpected As Long = 4
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        iPoints = .NoOfPoints
    End With
    
    'Assert:
    Assert.AreEqual aExpected, iPoints
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesNoOfPointsX")
Public Sub ChartSeriesNoOfPointsX_NoSpaceWithNameAllRanges_ReturnsFour()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim iPoints As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    Const cElement As Long = 2
    '==========================================================================
    Const aExpected As Long = 4
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        iPoints = .NoOfPoints
    End With
    
    'Assert:
    Assert.AreEqual aExpected, iPoints
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesDataSheetY")
Public Sub ChartSeriesDataSheetY_NoSpaceWithNameAllRanges_ReturnsWks()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim Actual As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpected As String
        aExpected = wks.Name
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    Actual = MySeries.XValues.RangeSheet
    
    'Assert:
    Assert.AreEqual aExpected, Actual
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesDataSheetX")
Public Sub ChartSeriesDataSheetX_NoSpaceWithNameAllRanges_ReturnsWks()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim Actual As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpected As String
        aExpected = wks.Name
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    Actual = MySeries.XValues.RangeSheet
    
    'Assert:
    Assert.AreEqual aExpected, Actual
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesDataSheetX")
Public Sub ChartSeriesDataSheetX_RoundBracketsWithNameAllRanges_ReturnsWks()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim Actual As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Dim aExpected As String
        aExpected = wks.Name
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    Actual = MySeries.XValues.RangeSheet
    
    'Assert:
    Assert.AreEqual aExpected, Actual
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesPointXSourceRange")
Public Sub ChartSeriesPointXSourceRange_NoSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim sPointAddress As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    Const cElement As Long = 2
    Const ciPoint As Long = 2
    '==========================================================================
    Dim aExpected As String
        aExpected = "NoSpace!A5"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        sPointAddress = .PointSourceRange(cElement, ciPoint)
    End With
    
    'Assert:
    Assert.AreEqual aExpected, sPointAddress
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChartSeriesPointYSourceRange")
Public Sub ChartSeriesPointYSourceRange_NoSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As IChartSeries
    
    Dim sPointAddress As String
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    Const cElement As Long = 3
    Const ciPoint As Long = 2
    '==========================================================================
    Dim aExpected As String
'NOTE: (totally) correct would be the variant *with* single quotes
'      aExpected = "'Space, Comma'!C5"
        aExpected = "Space, Comma!C5"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = ChartSeries.Create( _
        cha.Chart.SeriesCollection(ciSeries) _
    )
    
    'Act:
    With MySeries
        sPointAddress = .PointSourceRange(cElement, ciPoint)
    End With
    
    'Assert:
    Assert.AreEqual aExpected, sPointAddress
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

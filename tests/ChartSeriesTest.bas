Attribute VB_Name = "ChartSeriesTest"

'@TestModule
'@Folder("ChartSeries.Tests")

Option Explicit
Option Private Module

Private cha As Chart
Private srs As Series
Private sut As IChartSeries

Private Assert As Rubberduck.PermissiveAssertClass
'Private Fakes As Rubberduck.FakesProvider


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
'    Set Fakes = New Rubberduck.FakesProvider
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
'    Set Fakes = Nothing
End Sub


'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'==============================================================================
'@TestMethod("CreateClass")
Private Sub ChartSeries_SeriesIsNothing_Throws()
    Const ExpectedError As Long = eChartSeriesError.ErrIsNothing
    On Error GoTo TestFail
    
    Set srs = Nothing
    Set sut = ChartSeries.Create(srs)
    
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


'==============================================================================
'@TestMethod("FullFormula")
Private Sub FullFormula_SeriesFormulaInaccessible_Throws()
    Const ExpectedError As Long = eChartSeriesError.ErrNotAccessible
    On Error GoTo TestFail
    
    Set cha = tblNoSpace.ChartObjects("chaOneAreaBubble").Chart
    Set srs = cha.FullSeriesCollection(2)
    Set sut = ChartSeries.Create(srs)
    
    Dim Actual As String
    Actual = sut.FullFormula
    
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


'@TestMethod("FullFormula")
Private Sub FullFormula_HiddenSeriesFormulaInaccessible_Throws()
    Const ExpectedError As Long = eChartSeriesError.ErrNotAccessible
    On Error GoTo TestFail
    
    Set cha = tblNoSpace.ChartObjects("chaTwoAreas").Chart
    Set srs = cha.FullSeriesCollection(3)
    Set sut = ChartSeries.Create(srs)
    
    Dim Actual As String
    Actual = sut.FullFormula
    
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


'@TestMethod("FullFormula")
Private Sub FullFormula_OfBubbleChart_ReturnsString()
    On Error GoTo TestFail
    Const Expected As String = _
            "=SERIES('Space, Comma'!$C$3," & _
            "('Space, Comma'!$A$4,'Space, Comma'!$A$5:$A$6,'Space, Comma'!$A$7)," & _
            "('Space, Comma'!$C$4,'Space, Comma'!$C$5:$C$6,'Space, Comma'!$C$7)," & _
            "1," & _
            "('Space, Comma'!$B$4,'Space, Comma'!$B$5:$B$6,'Space, Comma'!$B$7))"
    
    Set cha = tblSpaceComma.ChartObjects("chaTwoAreasBubble").Chart
    Set srs = cha.FullSeriesCollection(1)
    Set sut = ChartSeries.Create(srs)
    
    Dim Actual As String
    Actual = sut.FullFormula
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("FullFormula")
Private Sub FullFormula_NotOfBubbleChart_ReturnsString()
    On Error GoTo TestFail
    Dim Expected As String
    Expected = _
            "=SERIES(""named ranges""," & ThisWorkbook.Name & "!wkb_y1," & _
            "'Space, Comma'!wks_y2,3)"
    
    Set cha = tblWithSpace.ChartObjects("chaFourAreas").Chart
    Set srs = cha.FullSeriesCollection(3)
    Set sut = ChartSeries.Create(srs)
    
    Dim Actual As String
    Actual = sut.FullFormula
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("SeriesName")
Private Sub SeriesName_SeriesFormulaInaccessible_Throws()
    Const ExpectedError As Long = eChartSeriesError.ErrNotAccessible
    On Error GoTo TestFail
    
    Set cha = tblNoSpace.ChartObjects("chaOneAreaBubble").Chart
    Set srs = cha.FullSeriesCollection(2)
    Set sut = ChartSeries.Create(srs)
    
    Dim Actual As ISeriesPart
    Set Actual = sut.SeriesName
    
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


'@TestMethod("SeriesName")
Private Sub SeriesName_SeriesFormulaAccessible_ReturnsObject()
    On Error GoTo TestFail
    
    Set cha = tblSpaceComma.ChartObjects("chaTwoAreasBubble").Chart
    Set srs = cha.FullSeriesCollection(1)
    Set sut = ChartSeries.Create(srs)
    
    Assert.IsNotNothing sut.SeriesName
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("XValues")
Private Sub XValues_SeriesFormulaInaccessible_Throws()
    Const ExpectedError As Long = eChartSeriesError.ErrNotAccessible
    On Error GoTo TestFail
    
    Set cha = chaBubbleChart
    Set srs = cha.FullSeriesCollection(2)
    Set sut = ChartSeries.Create(srs)
    
    Dim Actual As ISeriesPart
    Set Actual = sut.XValues
    
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


'@TestMethod("XValues")
Private Sub XValues_SeriesFormulaAccessible_ReturnsObject()
    On Error GoTo TestFail
    
    Set cha = tblSpaceComma.ChartObjects("chaFourAreas").Chart
    Set srs = cha.FullSeriesCollection(2)
    Set sut = ChartSeries.Create(srs)
    
    Assert.IsNotNothing sut.XValues
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("Values")
Private Sub Values_SeriesFormulaInaccessible_Throws()
    Const ExpectedError As Long = eChartSeriesError.ErrNotAccessible
    On Error GoTo TestFail
    
    Set cha = tblNoSpace.ChartObjects("chaOneAreaBubble").Chart
    Set srs = cha.FullSeriesCollection(2)
    Set sut = ChartSeries.Create(srs)
    
    Dim Actual As ISeriesPart
    Set Actual = sut.Values
    
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


'@TestMethod("Values")
Private Sub Values_SeriesFormulaAccessible_ReturnsObject()
    On Error GoTo TestFail
    
    Set cha = tblRoundBrackets.ChartObjects("chaOneArea").Chart
    Set srs = cha.FullSeriesCollection(3)
    Set sut = ChartSeries.Create(srs)
    
    Assert.IsNotNothing sut.Values
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("PlotOrder")
Private Sub PlotOrder_SeriesFormulaInaccessible_Throws()
    Const ExpectedError As Long = eChartSeriesError.ErrNotAccessible
    On Error GoTo TestFail
    
    Set cha = tblNoSpace.ChartObjects("chaOneAreaBubble").Chart
    Set srs = cha.FullSeriesCollection(2)
    Set sut = ChartSeries.Create(srs)
    
    Dim Actual As ISeriesPlotOrder
    Set Actual = sut.PlotOrder
    
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


'@TestMethod("PlotOrder")
Private Sub PlotOrder_SeriesFormulaAccessible_ReturnsObject()
    On Error GoTo TestFail
    
    Set cha = tblMaxName.ChartObjects("chaMultipleAreas").Chart
    Set srs = cha.FullSeriesCollection(1)
    Set sut = ChartSeries.Create(srs)
    
    Assert.IsNotNothing sut.PlotOrder
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("BubbleSizes")
Private Sub BubbleSizes_SeriesFormulaInaccessible_Throws()
    Const ExpectedError As Long = eChartSeriesError.ErrNotAccessible
    On Error GoTo TestFail
    
    Set cha = tblNoSpace.ChartObjects("chaOneAreaBubble").Chart
    Set srs = cha.FullSeriesCollection(2)
    Set sut = ChartSeries.Create(srs)
    
    Dim Actual As ISeriesPart
    Set Actual = sut.BubbleSizes
    
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


'@TestMethod("BubbleSizes")
Private Sub BubbleSizes_SeriesNotInBubbleChart_Throws()
    Const ExpectedError As Long = eChartSeriesError.ErrNotInBubbleChart
    On Error GoTo TestFail
    
    Set cha = tblWithSpace.ChartObjects("chaFourAreas").Chart
    Set srs = cha.FullSeriesCollection(3)
    Set sut = ChartSeries.Create(srs)
    
    Dim Actual As ISeriesPart
    Set Actual = sut.BubbleSizes
    
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


'@TestMethod("BubbleSizes")
Private Sub BubbleSizes_SeriesFormulaAccessible_ReturnsObject()
    On Error GoTo TestFail
    
    Set cha = tblRoundBrackets.ChartObjects("chaOneAreaBubble").Chart
    Set srs = cha.FullSeriesCollection(1)
    Set sut = ChartSeries.Create(srs)
    
    Assert.IsNotNothing sut.BubbleSizes
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("IsSeriesAccessible")
Private Sub IsSeriesAccessible_SeriesFormulaAccessible_ReturnsTrue()
    On Error GoTo TestFail
    
    Set cha = tblNoSpace.ChartObjects("chaOneAreaBubble").Chart
    Set srs = cha.FullSeriesCollection(1)
    Set sut = ChartSeries.Create(srs)
    
    Assert.IsTrue sut.IsSeriesAccessible
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsSeriesAccessible")
Private Sub IsSeriesAccessible_SeriesFormulaAccessible_ReturnsFalse()
    On Error GoTo TestFail
    
    Set cha = tblNoSpace.ChartObjects("chaOneAreaBubble").Chart
    Set srs = cha.FullSeriesCollection(2)
    Set sut = ChartSeries.Create(srs)
    
    Assert.IsFalse sut.IsSeriesAccessible
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("IsSeriesInBubbleChart")
Private Sub IsSeriesInBubbleChart_SeriesInBubbleChart_ReturnsTrue()
    On Error GoTo TestFail
    
    Set cha = chaBubbleChart
    Set srs = cha.FullSeriesCollection(1)
    Set sut = ChartSeries.Create(srs)
    
    Assert.IsTrue sut.IsSeriesInBubbleChart
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsSeriesInBubbleChart")
Private Sub IsSeriesInBubbleChart_SeriesNotInBubbleChart_ReturnsFalse()
    On Error GoTo TestFail
    
    Set cha = tblWithSpace.ChartObjects("chaOneArea").Chart
    Set srs = cha.FullSeriesCollection(1)
    Set sut = ChartSeries.Create(srs)
    
    Assert.IsFalse sut.IsSeriesInBubbleChart
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

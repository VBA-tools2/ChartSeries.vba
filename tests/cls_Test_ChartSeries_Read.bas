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
'unit tests for 'clsChartSeries' -- read stuff
'==============================================================================

'@TestMethod
Public Sub clsChartSeriesPlotOrder_SurfacePlot_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaSurface")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesYValues_SurfacePlot_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaSurface")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        sValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesXValues_SurfacePlot_ReturnsEmpty()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaSurface")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Empty"
    Const aExpectedValue As String = vbNullString
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_SurfacePlot_ReturnsEmpty()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaSurface")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Empty"
    Const aExpectedValue As String = vbNullString
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod
Public Sub clsChartSeriesBubbleSizes_NoBubbleChart_ReturnsError()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 1
    '==========================================================================
    Const aExpectedType As String = "Error - No Bubble Chart"
    Const aExpectedValue As String = "Error - No Bubble Chart"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .BubbleSizesType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .BubbleSizes
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_InvalidSeriesNumber_ReturnsError()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 5
    '==========================================================================
    Const aExpectedType As String = "ERROR - BAD SERIES NUMBER"
    Const aExpectedValue As String = "ERROR - BAD SERIES NUMBER"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesPlotOrder_NoSpaceWithNameAllRanges_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesYValues_NoSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$4:$C$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesXValues_NoSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$A$4:$A$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_NoSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod
Public Sub clsChartSeriesYValues_NoSpaceWithArrayValues_ReturnsArray()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 3
    '==========================================================================
    Const aExpectedType As String = "Array"
    Const aExpectedValue As String = "1.5,2.5,3.5,4.5"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesXValues_NoSpaceWithNoXValues_ReturnsEmptyString()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 3
    '==========================================================================
    Const aExpectedType As String = "Empty"
    Const aExpectedValue As String = vbNullString
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_NoSpaceWithString_ReturnsString()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 3
    '==========================================================================
    Const aExpectedType As String = "String"
    Const aExpectedValue As String = "just a test, with a comma"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod
Public Sub clsChartSeriesPlotOrder_NoSpaceWithNameFourAreas_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesYValues_NoSpaceWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$4,$C$5,$C$6,$C$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesXValues_NoSpaceWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$A$4,$A$5,$A$6,$A$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_NoSpaceWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod
Public Sub clsChartSeriesPlotOrder_NoSpaceWithNameTwoAreas_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesYValues_NoSpaceWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$4,$C$5:$C$6,$C$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesXValues_NoSpaceWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$A$4,$A$5:$A$6,$A$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_NoSpaceWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod
Public Sub clsChartSeriesPlotOrder_WithSpaceWithNameAllRanges_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblWithSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesYValues_WithSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblWithSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$4:$C$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesXValues_WithSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblWithSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$A$4:$A$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_WithSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblWithSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod
Public Sub clsChartSeriesPlotOrder_SpaceCommaWithNameAllRanges_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesYValues_SpaceCommaWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$4:$C$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesXValues_SpaceCommaWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$A$4:$A$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_SpaceCommaWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod
Public Sub clsChartSeriesYValues_NoSpaceWithStringTitleContainingComma_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 4
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$B$4:$B$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod
Public Sub clsChartSeriesPlotOrder_RoundBracketsWithNameAllRanges_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesYValues_RoundBracketsWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$4:$C$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesXValues_RoundBracketsWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$A$4:$A$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_RoundBracketsWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod
Public Sub clsChartSeriesYValues_RoundBracketsWithArrayValues_ReturnsArray()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 3
    '==========================================================================
    Const aExpectedType As String = "Array"
    Const aExpectedValue As String = "1.5,2.5,3.5,4.5"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_RoundBracketsWithString_ReturnsString()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 3
    '==========================================================================
    Const aExpectedType As String = "String"
    Const aExpectedValue As String = "just a test, with a comma"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod
Public Sub clsChartSeriesPlotOrder_RoundBracketsWithNameFourAreas_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesYValues_RoundBracketsWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$4,$C$5,$C$6,$C$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesXValues_RoundBracketsWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$A$4,$A$5,$A$6,$A$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesSeriesName_RoundBracketsWithNameFourAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaFourAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod
Public Sub clsChartSeriesPlotOrder_RoundBracketsWithNameTwoAreas_ReturnsTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 2
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesYValues_RoundBracketsWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$4,$C$5:$C$6,$C$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesXValues_RoundBracketsWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$A$4,$A$5:$A$6,$A$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod
Public Sub clsChartSeriesBubbleSizes_NoSpaceWithNameAllRangesBubblePlot_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneAreaBubble")
    Const ciSeries As Long = 1
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$4:$C$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .BubbleSizesType
        If sType = "Range" Then
            Set rng = .BubbleSizes
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesPlotOrder_NoSpaceWithNameAllRangesBubblePlot_ReturnsOne()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneAreaBubble")
    Const ciSeries As Long = 1
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 1
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'------------------------------------------------------------------------------
'@TestMethod
Public Sub clsChartSeriesBubbleSizes_SpaceCommaWithNameAllRangesBubblePlot_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaTwoAreasBubble")
    Const ciSeries As Long = 1
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$B$4,$B$5:$B$6,$B$7"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .BubbleSizesType
        If sType = "Range" Then
            Set rng = .BubbleSizes
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesPlotOrder_SpaceCommaWithNameAllRangesBubblePlot_ReturnsOne()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaTwoAreasBubble")
    Const ciSeries As Long = 1
    '==========================================================================
    Const aExpectedType As String = "Integer"
    Const aExpectedValue As Long = 1
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .PlotOrderType
        iValue = .PlotOrder
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, iValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod
Public Sub clsChartSeriesSeriesName_RoundBracketsWithNameTwoAreas_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedType As String = "Range"
    Const aExpectedValue As String = "$C$3"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sType = .SeriesNameType
        If sType = "Range" Then
            Set rng = .SeriesName
            sValue = rng.Address(External:=False)
        Else
            sValue = .SeriesName
        End If
    End With
    
    'Assert:
    With Assert
        .AreEqual aExpectedType, sType
        .AreEqual aExpectedValue, sValue
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod
Public Sub clsChartSeriesNoOfPointsY_NoSpaceWithNameAllRanges_ReturnsFour()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
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
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
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


'@TestMethod
Public Sub clsChartSeriesNoOfPointsX_NoSpaceWithNameAllRanges_ReturnsFour()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
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
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
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


'@TestMethod
Public Sub clsChartSeriesDataSheetY_NoSpaceWithNameAllRanges_ReturnsWks()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sDataSheet As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    Const cElement As Long = 3
    '==========================================================================
    Dim aExpected As String
        aExpected = wks.Name
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sDataSheet = .DataSheet(cElement)
    End With
    
    'Assert:
    Assert.AreEqual aExpected, sDataSheet
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesDataSheetX_NoSpaceWithNameAllRanges_ReturnsWks()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sDataSheet As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    Const cElement As Long = 2
    '==========================================================================
    Dim aExpected As String
        aExpected = wks.Name
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sDataSheet = .DataSheet(cElement)
    End With
    
    'Assert:
    Assert.AreEqual aExpected, sDataSheet
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesDataSheetX_RoundBracketsWithNameAllRanges_ReturnsWks()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sDataSheet As String
    
    '==========================================================================
    Set wks = tblRoundBrackets
    Set cha = wks.ChartObjects("chaOneArea")
    Const ciSeries As Long = 2
    Const cElement As Long = 2
    '==========================================================================
    Dim aExpected As String
        aExpected = wks.Name
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    'Act:
    With MySeries
        sDataSheet = .DataSheet(cElement)
    End With
    
    'Assert:
    Assert.AreEqual aExpected, sDataSheet
    
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesPointXSourceRange_NoSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
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
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
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


'@TestMethod
Public Sub clsChartSeriesPointYSourceRange_NoSpaceWithNameAllRanges_ReturnsAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sPointAddress As String
    
    '==========================================================================
    Set wks = tblSpaceComma
    Set cha = wks.ChartObjects("chaTwoAreas")
    Const ciSeries As Long = 2
    Const cElement As Long = 3
    Const ciPoint As Long = 2
    '==========================================================================
    Dim aExpected As String
        aExpected = "'Space, Comma'!C5"
    '==========================================================================
    
    
    'Arrange:
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
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

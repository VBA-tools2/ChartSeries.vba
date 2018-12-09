Attribute VB_Name = "cls_Test_ChartSeries_Write"

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
'unit tests for 'clsChartSeries' -- write stuff
'==============================================================================

'@TestMethod
Public Sub clsChartSeriesLetPlotOrder_NoSpaceWithNameTwoAreas_ReturnsSetTwo()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim iValue As Long
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreasWrite")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedValueStart As Long = 2
    Const aExpectedValueChange As Long = 1
    '==========================================================================
    
    
    'Arrange (and check):
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    iValue = MySeries.PlotOrder
    Assert.AreEqual aExpectedValueStart, iValue
    
    
    'Act:
    MySeries.PlotOrder = aExpectedValueChange
    
    'Assert:
    iValue = MySeries.PlotOrder
    Assert.AreEqual aExpectedValueChange, iValue
    
    
    'Revert (and check):
    With MySeries
        .PlotOrder = aExpectedValueStart
        iValue = .PlotOrder
    End With
    Assert.AreEqual aExpectedValueStart, iValue
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartLetSeriesYValues_NoSpaceWithNameTwoAreas_ReturnsSetAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim rng As Range
    Dim sType As String
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreasWrite")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedValueStart As String = "$C$4,$C$5:$C$6,$C$7"
    Dim arrChange As Variant
    arrChange = Array(1, 2, 3, 4)
    Const aExpectedValueChange As String = "1,2,3,4"
    '==========================================================================
    
    
    'Arrange (and check):
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    With MySeries
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    Assert.AreEqual aExpectedValueStart, sValue
    
    
    'Act:
    With MySeries
        .Values = arrChange
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    
    'Assert:
    Assert.AreEqual aExpectedValueChange, sValue
    
    
    'Revert (and check):
    With MySeries
        .Values = wks.Range(aExpectedValueStart)
        sType = .ValuesType
        If sType = "Range" Then
            Set rng = .Values
            sValue = rng.Address(External:=False)
        Else
            sValue = .Values
        End If
    End With
    Assert.AreEqual aExpectedValueStart, sValue
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartLetSeriesXValues_NoSpaceWithNameTwoAreas_ReturnsSetAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreasWrite")
    Const ciSeries As Long = 2
    '==========================================================================
    Const aExpectedValueStart As String = "$A$4,$A$5:$A$6,$A$7"
    Dim arrChange As Variant
    arrChange = Array(1, 2, 3, 4)
    Const aExpectedValueChange As String = "1,2,3,4"
    '==========================================================================
    
    
    'Arrange (and check):
    Set MySeries = New clsChartSeries
    With MySeries
        .Chart = cha.Chart
        .ChartSeries = ciSeries
    End With
    
    With MySeries
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    Assert.AreEqual aExpectedValueStart, sValue
    
    
    'Act:
    With MySeries
        .XValues = arrChange
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    
    'Assert:
    Assert.AreEqual aExpectedValueChange, sValue
    
    
    'Revert (and check):
    With MySeries
        .XValues = wks.Range(aExpectedValueStart)
        sType = .XValuesType
        If sType = "Range" Then
            Set rng = .XValues
            sValue = rng.Address(External:=False)
        Else
            sValue = .XValues
        End If
    End With
    Assert.AreEqual aExpectedValueStart, sValue
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub clsChartSeriesLetSeriesName_NoSpaceWithNameTwoAreas_ReturnsSetAddress()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    Dim cha As ChartObject
    Dim MySeries As clsChartSeries
    
    Dim sType As String
    Dim rng As Range
    Dim sValue As String
    
    '==========================================================================
    Set wks = tblNoSpace
    Set cha = wks.ChartObjects("chaTwoAreasWrite")
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

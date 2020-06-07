Attribute VB_Name = "SeriesFormulaPartsTest"

'@TestModule
'@Folder("ChartSeries.Tests")

Option Explicit
Option Private Module

Private sut As ISeriesFormulaParts

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
Private Sub PartSeriesFormula_EmptyFullSeriesFormula_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrChartWorkbookNameEmpty
    On Error GoTo TestFail
    
    Set sut = SeriesFormulaParts.Create(vbNullString)
    
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
'@TestMethod("LiteralArray")
Private Sub PartSeriesFormula_NoBubbleLiteralArray_ReturnsString()
    On Error GoTo TestFail
    Dim Expected(eElement.[_First] To eElement.[_Last]) As String
        Expected(eElement.eName) = vbNullString
        Expected(eElement.eXValues) = "{1.5,2.5,3.5,4.5}"
        Expected(eElement.eYValues) = "{1}"
        Expected(eElement.ePlotOrder) = "1"
        Expected(eElement.eBubbleSizes) = vbNullString
    
    Set sut = SeriesFormulaParts.Create( _
            "=SERIES(,{1.5,2.5,3.5,4.5},{1},1)", _
            False _
    )
    
    If sut.PartSeriesFormula(eElement.eName) <> _
            Expected(eElement.eName) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.ePlotOrder) <> _
            Expected(eElement.ePlotOrder) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.eBubbleSizes) <> _
            Expected(eElement.eBubbleSizes) Then Assert.Inconclusive
    
    With Assert
        .AreEqual Expected(eElement.eYValues), sut.PartSeriesFormula(eElement.eYValues)
        .AreEqual Expected(eElement.eXValues), sut.PartSeriesFormula(eElement.eXValues)
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("LiteralArray")
Private Sub PartSeriesFormula_BubbleSeriesLiteralArray_ReturnsString()
    On Error GoTo TestFail
    Dim Expected(eElement.[_First] To eElement.[_Last]) As String
        Expected(eElement.eName) = vbNullString
        Expected(eElement.eXValues) = "{1.5,2.5,3.5,4.5}"
        Expected(eElement.eYValues) = "{1}"
        Expected(eElement.ePlotOrder) = "1"
        Expected(eElement.eBubbleSizes) = "{1.5,2.5,3.5,4.5}"
    
    Set sut = SeriesFormulaParts.Create( _
            "=SERIES(,{1.5,2.5,3.5,4.5},{1},1,{1.5,2.5,3.5,4.5})", _
            True _
    )
    
    If sut.PartSeriesFormula(eElement.eName) <> _
            Expected(eElement.eName) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.ePlotOrder) <> _
            Expected(eElement.ePlotOrder) Then Assert.Inconclusive
    
    With Assert
        .AreEqual Expected(eElement.eBubbleSizes), sut.PartSeriesFormula(eElement.eBubbleSizes)
        .AreEqual Expected(eElement.eYValues), sut.PartSeriesFormula(eElement.eYValues)
        .AreEqual Expected(eElement.eXValues), sut.PartSeriesFormula(eElement.eXValues)
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("Empty")
Private Sub PartSeriesFormula_NoBubbleEmptyEntries_ReturnEmptyStrings()
    On Error GoTo TestFail
    Dim Expected(eElement.[_First] To eElement.[_Last]) As String
        Expected(eElement.eName) = vbNullString
        Expected(eElement.eXValues) = vbNullString
        Expected(eElement.eYValues) = "{1}"
        Expected(eElement.ePlotOrder) = "1"
        Expected(eElement.eBubbleSizes) = vbNullString
    
    Set sut = SeriesFormulaParts.Create( _
            "=SERIES(,,{1},1)", _
            False _
    )
    
    If sut.PartSeriesFormula(eElement.eName) <> _
            Expected(eElement.eName) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.ePlotOrder) <> _
            Expected(eElement.ePlotOrder) Then Assert.Inconclusive
    
    With Assert
        .AreEqual Expected(eElement.eBubbleSizes), sut.PartSeriesFormula(eElement.eBubbleSizes)
        .AreEqual Expected(eElement.eYValues), sut.PartSeriesFormula(eElement.eYValues)
        .AreEqual Expected(eElement.eXValues), sut.PartSeriesFormula(eElement.eXValues)
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Empty")
Private Sub PartSeriesFormula_BubbleSeriesEmptyEntries_ReturnEmptyStrings()
    On Error GoTo TestFail
    Dim Expected(eElement.[_First] To eElement.[_Last]) As String
        Expected(eElement.eName) = vbNullString
        Expected(eElement.eXValues) = vbNullString
        Expected(eElement.eYValues) = "{1}"
        Expected(eElement.ePlotOrder) = "1"
        Expected(eElement.eBubbleSizes) = "{1}"
    
    Set sut = SeriesFormulaParts.Create( _
            "=SERIES(,,{1},1,{1})", _
            True _
    )
    
    If sut.PartSeriesFormula(eElement.eName) <> _
            Expected(eElement.eName) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.ePlotOrder) <> _
            Expected(eElement.ePlotOrder) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.eBubbleSizes) <> _
            Expected(eElement.eBubbleSizes) Then Assert.Inconclusive
    
    With Assert
        .AreEqual Expected(eElement.eYValues), sut.PartSeriesFormula(eElement.eYValues)
        .AreEqual Expected(eElement.eXValues), sut.PartSeriesFormula(eElement.eXValues)
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("MultiAreaRange")
Private Sub PartSeriesFormula_NoBubbleSeriesSimpleMultiAreaRangesWithoutSingleQuotes_ReturnStrings()
    On Error GoTo TestFail
    Dim Expected(eElement.[_First] To eElement.[_Last]) As String
        Expected(eElement.eName) = "NoSpace!$C$3"
        Expected(eElement.eXValues) = "(NoSpace!$A$4,NoSpace!$A$5,NoSpace!$A$6,NoSpace!$A$7)"
        Expected(eElement.eYValues) = "(NoSpace!$C$4,NoSpace!$C$5,NoSpace!$C$6,NoSpace!$C$7)"
        Expected(eElement.ePlotOrder) = "2"
        Expected(eElement.eBubbleSizes) = vbNullString
    
    Set sut = SeriesFormulaParts.Create( _
            "=SERIES(NoSpace!$C$3," & _
                    "(NoSpace!$A$4,NoSpace!$A$5,NoSpace!$A$6,NoSpace!$A$7)," & _
                    "(NoSpace!$C$4,NoSpace!$C$5,NoSpace!$C$6,NoSpace!$C$7)," & _
                    "2)", _
            False _
    )
    
    If sut.PartSeriesFormula(eElement.eName) <> _
            Expected(eElement.eName) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.ePlotOrder) <> _
            Expected(eElement.ePlotOrder) Then Assert.Inconclusive
    
    With Assert
        .AreEqual Expected(eElement.eBubbleSizes), sut.PartSeriesFormula(eElement.eBubbleSizes)
        .AreEqual Expected(eElement.eYValues), sut.PartSeriesFormula(eElement.eYValues)
        .AreEqual Expected(eElement.eXValues), sut.PartSeriesFormula(eElement.eXValues)
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("MultiAreaRange")
Private Sub PartSeriesFormula_BubbleSeriesSimpleMultiAreaRangesWithSingleQuotes_ReturnStrings()
    On Error GoTo TestFail
    Dim Expected(eElement.[_First] To eElement.[_Last]) As String
        Expected(eElement.eName) = "',(''a""''!'!$C$4"
        Expected(eElement.eXValues) = "(',(''a""''!'!$A$4,',(''a""''!'!$A$5:$A$6)"
        Expected(eElement.eYValues) = "(',(''a""''!'!$B$4,',(''a""''!'!$B$5:$B$6)"
        Expected(eElement.ePlotOrder) = "4"
        Expected(eElement.eBubbleSizes) = "(',(''a""''!'!$C$4,',(''a""''!'!$C$5:$C$6)"
    
    Set sut = SeriesFormulaParts.Create( _
            "=SERIES(',(''a""''!'!$C$4," & _
                    "(',(''a""''!'!$A$4,',(''a""''!'!$A$5:$A$6)," & _
                    "(',(''a""''!'!$B$4,',(''a""''!'!$B$5:$B$6)," & _
                    "4," & _
                    "(',(''a""''!'!$C$4,',(''a""''!'!$C$5:$C$6))", _
            True _
    )
    
    If sut.PartSeriesFormula(eElement.eName) <> _
            Expected(eElement.eName) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.ePlotOrder) <> _
            Expected(eElement.ePlotOrder) Then Assert.Inconclusive
    
    With Assert
        .AreEqual Expected(eElement.eBubbleSizes), sut.PartSeriesFormula(eElement.eBubbleSizes)
        .AreEqual Expected(eElement.eYValues), sut.PartSeriesFormula(eElement.eYValues)
        .AreEqual Expected(eElement.eXValues), sut.PartSeriesFormula(eElement.eXValues)
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("SingleAreaRange")
Private Sub PartSeriesFormula_NoBubbleSeriesSingleAreaRangesWithoutSingleQuote_ReturnStrings()
    On Error GoTo TestFail
    Dim Expected(eElement.[_First] To eElement.[_Last]) As String
        Expected(eElement.eName) = vbNullString
        Expected(eElement.eXValues) = "NoSpace!$A$4:$A$7"
        Expected(eElement.eYValues) = "NoSpace!$C$4:$C$7"
        Expected(eElement.ePlotOrder) = "99"
        Expected(eElement.eBubbleSizes) = vbNullString
    
    Set sut = SeriesFormulaParts.Create( _
            "=SERIES(,NoSpace!$A$4:$A$7,NoSpace!$C$4:$C$7,99)", _
            False _
    )
    
    If sut.PartSeriesFormula(eElement.eName) <> _
            Expected(eElement.eName) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.ePlotOrder) <> _
            Expected(eElement.ePlotOrder) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.eBubbleSizes) <> _
            Expected(eElement.eBubbleSizes) Then Assert.Inconclusive
    
    With Assert
        .AreEqual Expected(eElement.eYValues), sut.PartSeriesFormula(eElement.eYValues)
        .AreEqual Expected(eElement.eXValues), sut.PartSeriesFormula(eElement.eXValues)
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SingleAreaRange")
Private Sub PartSeriesFormula_NoBubbleSeriesSingleAreaRangesWithCommata_ReturnStrings()
    On Error GoTo TestFail
    Dim Expected(eElement.[_First] To eElement.[_Last]) As String
        Expected(eElement.eName) = vbNullString
        Expected(eElement.eXValues) = "'Space, Comma'!$A$4:$A$7"
        Expected(eElement.eYValues) = "'Space, Comma'!$B$4:$B$7"
        Expected(eElement.ePlotOrder) = "143"
        Expected(eElement.eBubbleSizes) = vbNullString
    
    Set sut = SeriesFormulaParts.Create( _
            "=SERIES(,'Space, Comma'!$A$4:$A$7,'Space, Comma'!$B$4:$B$7,143)", _
            False _
    )
    
    If sut.PartSeriesFormula(eElement.eName) <> _
            Expected(eElement.eName) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.ePlotOrder) <> _
            Expected(eElement.ePlotOrder) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.eBubbleSizes) <> _
            Expected(eElement.eBubbleSizes) Then Assert.Inconclusive
    
    With Assert
        .AreEqual Expected(eElement.eYValues), sut.PartSeriesFormula(eElement.eYValues)
        .AreEqual Expected(eElement.eXValues), sut.PartSeriesFormula(eElement.eXValues)
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SingleAreaRange")
Private Sub PartSeriesFormula_BubbleSeriesSingleAreaRangesWithSingleQuotes_ReturnStrings()
    On Error GoTo TestFail
    Dim Expected(eElement.[_First] To eElement.[_Last]) As String
        Expected(eElement.eName) = vbNullString
        Expected(eElement.eXValues) = "'With Space'!$A$4:$A$7"
        Expected(eElement.eYValues) = "'With Space'!$B$4:$B$7"
        Expected(eElement.ePlotOrder) = "200"
        Expected(eElement.eBubbleSizes) = "'With Space'!$C$4:$C$7"
    
    Set sut = SeriesFormulaParts.Create( _
            "=SERIES(,'With Space'!$A$4:$A$7,'With Space'!$B$4:$B$7,200,'With Space'!$C$4:$C$7)", _
            True _
    )
    
    If sut.PartSeriesFormula(eElement.eName) <> _
            Expected(eElement.eName) Then Assert.Inconclusive
    If sut.PartSeriesFormula(eElement.ePlotOrder) <> _
            Expected(eElement.ePlotOrder) Then Assert.Inconclusive
    
    With Assert
        .AreEqual Expected(eElement.eBubbleSizes), sut.PartSeriesFormula(eElement.eBubbleSizes)
        .AreEqual Expected(eElement.eYValues), sut.PartSeriesFormula(eElement.eYValues)
        .AreEqual Expected(eElement.eXValues), sut.PartSeriesFormula(eElement.eXValues)
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("Dummy")
Private Sub Dummy_Empty_ReturnsEmptyEntryType()
    On Error GoTo TestFail
    Const Expected As String = "1"
    
    Set sut = SeriesFormulaParts.Create( _
            "=SERIES(,,{1},1)", _
            False _
    )
    
    Dim Actual As String
    Actual = sut.PartSeriesFormula(eElement.ePlotOrder)
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

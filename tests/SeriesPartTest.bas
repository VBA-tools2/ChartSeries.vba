Attribute VB_Name = "SeriesPartTest"

'@TestModule
'@Folder("ChartSeries.Tests")

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'NOTE: no tests for `SetRange` are done
'      (setting the range *should* work when the strings are right)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit
Option Private Module

Private Const ExternalOpenFilename As String = "ChartSeriesTest_Client.xlsx"
Private WorkbookName As String
Private sut As ISeriesPart

Private Assert As Rubberduck.PermissiveAssertClass
'Private Fakes As Rubberduck.FakesProvider


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
'    Set Fakes = New Rubberduck.FakesProvider
    
    'this file is needed for external open workbook tests
    Dim wkb As Workbook
    Set wkb = OpenWorkbook(ThisWorkbook.Path, ExternalOpenFilename)
    
    If wkb Is Nothing Then MsgBox ("'Client' file could not be opened.")
    
    WorkbookName = ThisWorkbook.Name
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
'    Set Fakes = Nothing
    
    WorkbookName = vbNullString
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
Private Sub SeriesPart_EmptyWorkbookName_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrChartWorkbookNameEmpty
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create(vbNullString, vbNullString)
    
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
'@TestMethod("EntryType")
Private Sub EntryType_Empty_ReturnsEmptyEntryType()
    On Error GoTo TestFail
    Dim Expected As eEntryType
    Expected = eEmpty
    
    Set sut = SeriesPart.Create(WorkbookName, vbNullString)
    
    Dim Actual As eEntryType
    Actual = sut.EntryType
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("EntryType")
Private Sub EntryType_String_ReturnsStringEntryType()
    On Error GoTo TestFail
    Dim Expected As eEntryType
    Expected = eString
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            """" & "just a test" & """" _
    )
    
    Dim Actual As eEntryType
    Actual = sut.EntryType
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("EntryType")
Private Sub EntryType_Array_ReturnsStringEntryType()
    On Error GoTo TestFail
    Dim Expected As eEntryType
    Expected = eArray
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "{7.245}" _
    )
    
    Dim Actual As eEntryType
    Actual = sut.EntryType
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("EntryType")
Private Sub EntryType_Integer_ReturnsIntegerEntryType()
    On Error GoTo TestFail
    Dim Expected As eEntryType
    Expected = eInteger
    
    Set sut = SeriesPart.Create(WorkbookName, "8")
    
    Dim Actual As eEntryType
    Actual = sut.EntryType
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("EntryType")
Private Sub EntryType_MultiAreaRange_ReturnsRangeEntryType()
    On Error GoTo TestFail
    Dim Expected As eEntryType
    Expected = eRange
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "(NoSpace!$A$4,NoSpace!$A$5:$A$6,NoSpace!$A$7)" _
    )
    
    Dim Actual As eEntryType
    Actual = sut.EntryType
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("EntryType")
Private Sub EntryType_InternalGlobalDefinedName_ReturnsDefinedNameEntryType()
    On Error GoTo TestFail
    Dim Expected As eEntryType
    Expected = eDefinedName
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            WorkbookName & "!wkb_y1" _
    )
    
    Dim Actual As eEntryType
    Actual = sut.EntryType
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("EntryType")
Private Sub EntryType_SingleAreaRange_ReturnsRangeEntryType()
    On Error GoTo TestFail
    Dim Expected As eEntryType
    Expected = eRange
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "NoSpace!$A$4:$A$7" _
    )
    
    Dim Actual As eEntryType
    Actual = sut.EntryType
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("FormulaPart")
Private Sub FormulaPart_Empty_ReturnsFormulaPartInput()
    On Error GoTo TestFail
    Const Expected As String = vbNullString
    
    Set sut = SeriesPart.Create(WorkbookName, Expected)
    If sut.EntryType <> eEmpty Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.FormulaPart
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("FormulaPart")
Private Sub FormulaPart_NotEmpty_ReturnsFormulaPartInput()
    On Error GoTo TestFail
    Const Expected As String = """" & "just a test" & """"
    
    Set sut = SeriesPart.Create(WorkbookName, Expected)
    If sut.EntryType <> eString Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.FormulaPart
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("CleanFormulaPart")
Private Sub CleanFormulaPart_WithoutSurroundingCharacters_ReturnsCleanFormulaPart()
    On Error GoTo TestFail
    Const Expected As String = "NoSpace!$A$4:$A$7"
    
    Set sut = SeriesPart.Create(WorkbookName, Expected)
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.CleanFormulaPart
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CleanFormulaPart")
Private Sub CleanFormulaPart_WithSurroundingQuotes_ReturnsCleanFormulaPart()
    On Error GoTo TestFail
    Const Expected As String = "just a test"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            """" & "just a test" & """" _
    )
    If sut.EntryType <> eString Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.CleanFormulaPart
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CleanFormulaPart")
Private Sub CleanFormulaPart_WithSurroundingCurlyBrackets_ReturnsCleanFormulaPart()
    On Error GoTo TestFail
    Const Expected As String = _
            "1.5,2.5,3.5,4.5"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "{1.5,2.5,3.5,4.5}" _
    )
    If sut.EntryType <> eArray Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.CleanFormulaPart
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CleanFormulaPart")
Private Sub CleanFormulaPart_WithSurroundingRoundBrackets_ReturnsCleanFormulaPart()
    On Error GoTo TestFail
    Const Expected As String = _
            "'With Space'!$A$4,'With Space'!$A$5,'With Space'!$A$6,'With Space'!$A$7"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "('With Space'!$A$4,'With Space'!$A$5,'With Space'!$A$6,'With Space'!$A$7)" _
    )
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.CleanFormulaPart
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("IsRange")
Private Sub IsRange_Array_ReturnsFalse()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "{1.5,2.5,3.5,4.5}" _
    )
    If sut.EntryType <> eArray Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.IsRange
    
    Assert.IsFalse Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsRange")
Private Sub IsRange_SingleCell_ReturnsTrue()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'With Space'!$A$4" _
    )
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.IsRange
    
    Assert.IsTrue Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsRange")
Private Sub IsRange_SingleArea_ReturnsTrue()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAA$1000001:$AAA$1000050" _
    )
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.IsRange
    
    Assert.IsTrue Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsRange")
Private Sub IsRange_MultiArea_ReturnsTrue()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "(',(''a""''!'!$A$4:$A$7,',(''a""''!'!$A$12:$A$24,',(''a""''!'!$A$45:$A$57)" _
    )
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.IsRange
    
    Assert.IsTrue Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsRange")
Private Sub IsRange_InternalLocalDefinedNameNotRange_ReturnsFalse()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "NoSpace!MyDefinedName" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.IsRange
    
    Assert.IsFalse Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsRange")
Private Sub IsRange_InternalGlobalDefinedNameRange_ReturnsTrue()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            WorkbookName & "!wkb_y1" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.IsRange
    
    Assert.IsTrue Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("IsDefinedNameRange")
Private Sub IsDefinedNameRange_Array_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotADefinedName
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "{1.25,3.75}" _
    )
    
    Dim Actual As String
    Actual = sut.IsDefinedNameRange
    
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


'@TestMethod("IsDefinedNameRange")
Private Sub IsDefinedNameRange_OpenExternalRange_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotADefinedName
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'[" & ExternalOpenFilename & "]"",)(''""'!$BQ$88:$BQ$137" _
    )
    
    Dim Actual As String
    Actual = sut.IsDefinedNameRange
    
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


'@TestMethod("IsDefinedNameRange")
Private Sub IsDefinedNameRange_InternalLocalDefinedNameNotRange_ReturnsFalse()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "NoSpace!MyDefinedName" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Assert.IsFalse sut.IsDefinedNameRange
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsDefinedNameRange")
Private Sub IsDefinedNameRange_InternalLocalDefinedNameRange_ReturnsTrue()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'Space, Comma'!wks_y2" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Assert.IsTrue sut.IsDefinedNameRange
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsDefinedNameRange")
Private Sub IsDefinedNameRange_InternalGlobalDefinedNameRange_ReturnsTrue()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            WorkbookName & "!wkb_y1" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Assert.IsTrue sut.IsDefinedNameRange
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsDefinedNameRange")
Private Sub IsDefinedNameRange_OpenExternalLocalDefinedNameNotRange_ReturnsFalse()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'[" & ExternalOpenFilename & "]"",)(''""'!LocalNotARange" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Assert.IsFalse sut.IsDefinedNameRange
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsDefinedNameRange")
Private Sub IsDefinedNameRange_OpenExternalGlobalDefinedNameNotRange_ReturnsFalse()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'" & ExternalOpenFilename & "'!GlobalNotARange" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Assert.IsFalse sut.IsDefinedNameRange
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsDefinedNameRange")
Private Sub IsDefinedNameRange_OpenExternalLocalDefinedNameRange_ReturnsTrue()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'[" & ExternalOpenFilename & "]"",)(''""'!LocalMyRange" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Assert.IsTrue sut.IsDefinedNameRange
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsDefinedNameRange")
Private Sub IsDefinedNameRange_OpenExternalGlobalDefinedNameRange_ReturnsTrue()
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            ExternalOpenFilename & "!GlobalMyRange" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Assert.IsTrue sut.IsDefinedNameRange
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("RangeString")
Private Sub RangeString_Array_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "{1.5,2.5,3.5,4.5}" _
    )
    
    Dim Actual As String
    Actual = sut.RangeString
    
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


'@TestMethod("RangeString")
Private Sub RangeString_SingleCell_ReturnsRangeCell()
    On Error GoTo TestFail
    Const Expected As String = "A4"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'With Space'!$A$4" _
    )
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeString
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeString")
Private Sub RangeString_SingleArea_ReturnsRangeArea()
    On Error GoTo TestFail
    Const Expected As String = "AAA1000001:AAA1000050"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAA$1000001:$AAA$1000050" _
    )
    
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeString
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeString")
Private Sub RangeString_MultiArea_ReturnsRangeAreas()
    On Error GoTo TestFail
    Const Expected As String = "A4:A7,A12:A24,A45:A57"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "(',(''a""''!'!$A$4:$A$7,',(''a""''!'!$A$12:$A$24,',(''a""''!'!$A$45:$A$57)" _
    )
    
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeString
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeString")
Private Sub RangeString_LongRange_ReturnsLongRangeString()
    On Error GoTo TestFail
    Const Expected As String = _
            "AAC1000001:AAC1000002,AAC1000005:AAC1000006,AAC1000009:AAC1000010," & _
            "AAC1000013:AAC1000014,AAC1000017:AAC1000018,AAC1000021:AAC1000022," & _
            "AAC1000025:AAC1000026,AAC1000029:AAC1000030,AAC1000033:AAC1000034," & _
            "AAC1000037:AAC1000038,AAC1000041:AAC1000042"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "('abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000001:$AAC$1000002," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000005:$AAC$1000006," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000009:$AAC$1000010," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000013:$AAC$1000014," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000017:$AAC$1000018," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000021:$AAC$1000022," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000025:$AAC$1000026," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000029:$AAC$1000030," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000033:$AAC$1000034," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000037:$AAC$1000038," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000041:$AAC$1000042)" _
    )
    
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeString
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeString")
Private Sub RangeString_VeryLongRange_ReturnsLongRangeString()
    On Error GoTo TestFail
    Const Expected As String = _
            "AAC1000001:AAC1000002,AAC1000005:AAC1000006,AAC1000009:AAC1000010," & _
            "AAC1000013:AAC1000014,AAC1000017:AAC1000018,AAC1000021:AAC1000022," & _
            "AAC1000025:AAC1000026,AAC1000029:AAC1000030,AAC1000033:AAC1000034," & _
            "AAC1000037:AAC1000038,AAC1000041:AAC1000042,AAC1000045:AAC1000046," & _
            "AAC1000049:AAC1000050"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "('abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000001:$AAC$1000002," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000005:$AAC$1000006," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000009:$AAC$1000010," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000013:$AAC$1000014," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000017:$AAC$1000018," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000021:$AAC$1000022," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000025:$AAC$1000026," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000029:$AAC$1000030," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000033:$AAC$1000034," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000037:$AAC$1000038," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000041:$AAC$1000042," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000045:$AAC$1000046," & _
            "'abcdefghijklmnopqrstuvwxyz ABCD'!$AAC$1000049:$AAC$1000050)" _
    )
    
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeString
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeString")
Private Sub RangeString_InternalLocalDefinedNameNotRange_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "NoSpace!MyDefinedName" _
    )
    
    Dim Actual As String
    Actual = sut.RangeString
    
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


'@TestMethod("RangeString")
Private Sub RangeString_InternalGlobalDefinedNameRange_ReturnsDefinedName()
    On Error GoTo TestFail
    Const Expected As String = "wkb_y1"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            WorkbookName & "!wkb_y1" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeString
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("RangeSheet")
Private Sub RangeSheet_Integer_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "125" _
    )
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
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


'@TestMethod("RangeSheet")
Private Sub RangeSheet_NotSingleQuotedSheetNameRange_ReturnsSheetName()
    On Error GoTo TestFail
    Const Expected As String = "NoSpace"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "(NoSpace!$A$4,NoSpace!$A$5:$A$6,NoSpace!$A$7)" _
    )
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeSheet")
Private Sub RangeSheet_SingleQuotedSheetNameRange_ReturnsSheetName()
    On Error GoTo TestFail
    Const Expected As String = "Space, Comma"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'Space, Comma'!$A$4:$A$7" _
    )
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeSheet")
Private Sub RangeSheet_SheetNameContainingSingleQuoteRange_ReturnsSheetName()
    On Error GoTo TestFail
    Const Expected As String = ",('a""'!"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "',(''a""''!'!$A$4:$A$7" _
    )
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeSheet")
Private Sub RangeSheet_OpenExternalRange_ReturnsSheetName()
    On Error GoTo TestFail
    Const Expected As String = """,)('"""
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'[" & ExternalOpenFilename & "]"",)(''""'!$BQ$88:$BQ$137" _
    )
    
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeSheet")
Private Sub RangeSheet_ClosedExternalRange_ReturnsSheetName()
    On Error GoTo TestFail
    Const Expected As String = "Test Sheet (666)"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'C:\Users\Random Guy\Desktop\[ClientClosed.xlsx]Test Sheet (666)'!$BK$190:$BK$221" _
    )
    
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeSheet")
Private Sub RangeSheet_InternalLocalDefinedNameNotRange_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "NoSpace!MyDefinedName" _
    )
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
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


'@TestMethod("RangeSheet")
Private Sub RangeSheet_InternalLocalDefinedNameRange_ReturnsDefinedName()
    On Error GoTo TestFail
    Const Expected As String = "Space, Comma"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'Space, Comma'!wks_y2" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeSheet")
Private Sub RangeSheet_InternalGlobalDefinedNameRange_ReturnsEmptyString()
    On Error GoTo TestFail
    Const Expected As String = vbNullString
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            WorkbookName & "!wkb_y1" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeSheet")
Private Sub RangeSheet_OpenExternalLocalDefinedNameNotRange_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'[" & ExternalOpenFilename & "]"",)(''""'!LocalNotARange" _
    )
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
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


'@TestMethod("RangeSheet")
Private Sub RangeSheet_OpenExternalGlobalDefinedNameNotRange_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'" & ExternalOpenFilename & "'!GlobalNotARange" _
    )
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
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


'@TestMethod("RangeSheet")
Private Sub RangeSheet_OpenExternalLocalDefinedNameRange_ReturnsDefinedName()
    On Error GoTo TestFail
    Const Expected As String = """,)('"""
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'[" & ExternalOpenFilename & "]"",)(''""'!LocalMyRange" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeSheet")
Private Sub RangeSheet_OpenExternalGlobalDefinedNameRange_ReturnsEmptyString()
    On Error GoTo TestFail
    Const Expected As String = vbNullString
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            ExternalOpenFilename & "!GlobalMyRange" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeSheet
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("RangeBook")
Private Sub RangeBook_Array_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "{7.123}" _
    )
    
    Dim Actual As String
    Actual = sut.RangeBook
    
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


'@TestMethod("RangeBook")
Private Sub RangeBook_InternalRange_ReturnsBookName()
    On Error GoTo TestFail
    Dim Expected As String
    Expected = WorkbookName
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "(NoSpace!$A$4,NoSpace!$A$5:$A$6,NoSpace!$A$7)" _
    )
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeBook
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeBook")
Private Sub RangeBook_OpenExternalRange_ReturnsBookName()
    On Error GoTo TestFail
    Const Expected As String = ExternalOpenFilename
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'[" & ExternalOpenFilename & "]"",)(''""'!$BQ$88:$BQ$137" _
    )
    
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeBook
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeBook")
Private Sub RangeBook_ClosedExternalRange_ReturnsBookName()
    On Error GoTo TestFail
    Const Expected As String = "ClientClosed.xlsx"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'C:\Users\Random Guy\Desktop\[ClientClosed.xlsx]Test Sheet (666)'!$BK$190:$BK$221" _
    )
    
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeBook
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeBook")
Private Sub RangeBook_InternalLocalDefinedNameNotRange_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "NoSpace!MyDefinedName" _
    )
    
    Dim Actual As String
    Actual = sut.RangeBook
    
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


'@TestMethod("RangeBook")
Private Sub RangeBook_InternalLocalDefinedNameRange_ReturnsBookName()
    On Error GoTo TestFail
    Dim Expected As String
    Expected = WorkbookName
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'Space, Comma'!wks_y2" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeBook
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeBook")
Private Sub RangeBook_InternalGlobalDefinedNameRange_ReturnsBookName()
    On Error GoTo TestFail
    Dim Expected As String
    Expected = WorkbookName
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            WorkbookName & "!wkb_y1" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeBook
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeBook")
Private Sub RangeBook_OpenExternalLocalDefinedNameNotRange_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'[" & ExternalOpenFilename & "]"",)(''""'!LocalNotARange" _
    )
    
    Dim Actual As String
    Actual = sut.RangeBook
    
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


'@TestMethod("RangeBook")
Private Sub RangeBook_OpenExternalGlobalDefinedNameNotRange_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'" & ExternalOpenFilename & "'!GlobalNotARange" _
    )
    
    Dim Actual As String
    Actual = sut.RangeBook
    
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


'@TestMethod("RangeBook")
Private Sub RangeBook_OpenExternalLocalDefinedNameRange_ReturnsBookName()
    On Error GoTo TestFail
    Dim Expected As String
    Expected = ExternalOpenFilename
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'[" & ExternalOpenFilename & "]"",)(''""'!LocalMyRange" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeBook
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangeBook")
Private Sub RangeBook_OpenExternalGlobalDefinedNameRange_ReturnsBookName()
    On Error GoTo TestFail
    Dim Expected As String
    Expected = ExternalOpenFilename
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            ExternalOpenFilename & "!GlobalMyRange" _
    )
    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangeBook
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("RangePath")
Private Sub RangePath_Array_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "{7.123}" _
    )
    
    Dim Actual As String
    Actual = sut.RangePath
    
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


'@TestMethod("RangePath")
Private Sub RangePath_InternalRange_ReturnsEmptyString()
    On Error GoTo TestFail
    Const Expected As String = vbNullString
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "(NoSpace!$A$4,NoSpace!$A$5:$A$6,NoSpace!$A$7)" _
    )
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangePath
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangePath")
Private Sub RangePath_OpenExternalRange_ReturnsEmptyString()
    On Error GoTo TestFail
    Const Expected As String = vbNullString
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'[" & ExternalOpenFilename & "]"",)(''""'!$BQ$88:$BQ$137" _
    )
    
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangePath
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangePath")
Private Sub RangePath_ClosedExternalRange_ReturnsPath()
    On Error GoTo TestFail
    Const Expected As String = "C:\Users\Random Guy\Desktop"
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'C:\Users\Random Guy\Desktop\[ClientClosed.xlsx]Test Sheet (666)'!$BK$190:$BK$221" _
    )
    
    If sut.EntryType <> eRange Then Assert.Inconclusive
    
    Dim Actual As String
    Actual = sut.RangePath
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("RangePath")
'In a closed file it can't be tested if the defined name is a range
Private Sub RangePath_ClosedExternalLocalDefinedName_Throws()
    Const ExpectedError As Long = eSeriesPartError.ErrNotARange
    On Error GoTo TestFail
    
    Set sut = SeriesPart.Create( _
            WorkbookName, _
            "'C:\Users\Random Guy\Desktop\[ClientClosed.xlsx]asdf'!TestName" _
    )
    
    Dim Actual As String
    Actual = sut.RangePath
    
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


''@TestMethod("NotTestable_ClosedExternalGlobalScopeDefinedName")
''currently (Excel 2010) not testable, because 'Series.Formula' is inaccessible
''in this case
'Private Sub RangeParts_ClosedExternalGlobalDefinedName_ReturnsRangeParts()
'    On Error GoTo TestFail
'    Const Expected As String = "C:\Users\Random Guy\Desktop"
'
'    Set sut = SeriesPart.Create( _
'            WorkbookName, _
'            "'C:\Users\Random Guy\Desktop\...!GlobalDefinedNameRange" _
'    )
'
'    If sut.EntryType <> eDefinedName Then Assert.Inconclusive
'
'    Dim Actual As String
'    Actual = sut.RangePath
'
'    Assert.AreEqual Expected, Actual
'    Assert.AreEqual "", sut.RangeBook
'    Assert.AreEqual "", sut.RangeSheet
'    Assert.AreEqual "", sut.RangeString
'    Assert.IsFalse sut.IsDefinedNameRange
'    Assert.IsFalse sut.IsRange
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub

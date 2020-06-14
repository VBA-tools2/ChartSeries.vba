Attribute VB_Name = "SeriesPlotOrderTest"

'@TestModule
'@Folder("ChartSeries.Tests")

Option Explicit
Option Private Module

Private sut As ISeriesPlotOrder

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
Private Sub SeriesPlotOrder_EmptyFormulaPart_Throws()
    Const ExpectedError As Long = eSeriesPlotOrderError.ErrNotNumericFormulaPart
    On Error GoTo TestFail
    
    Set sut = SeriesPlotOrder.Create(vbNullString)
    
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


'@TestMethod("CreateClass")
Private Sub SeriesPlotOrder_NotNumericFormulaPart_Throws()
    Const ExpectedError As Long = eSeriesPlotOrderError.ErrNotNumericFormulaPart
    On Error GoTo TestFail
    
    Set sut = SeriesPlotOrder.Create("a,1")
    
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
Private Sub EntryType_Integer_ReturnsIntegerEntryType()
    Dim Expected As eEntryType
    Expected = eInteger
    On Error GoTo TestFail
    
    Set sut = SeriesPlotOrder.Create("128")
    
    Dim Actual As eEntryType
    Actual = sut.EntryType
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'@TestMethod("Value")
Private Sub Value_Integer_ReturnsInteger()
    Const Expected As Byte = 255
    On Error GoTo TestFail
    
    Set sut = SeriesPlotOrder.Create("255")
    
    Dim Actual As Byte
    Actual = sut.Value
    
    Assert.AreEqual Expected, Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

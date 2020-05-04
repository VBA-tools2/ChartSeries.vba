Attribute VB_Name = "modUsefulFunctions"

Option Explicit
Option Private Module
Option Base 1


'==============================================================================
'adapted from <www.cpearson.com/excel/BetterUnion.aspx>
'A Union operation that accepts parameters that are 'Nothing'.
Public Function Union2( _
    ParamArray Ranges() As Variant _
        ) As Range
    
    Dim i As Long
    For i = LBound(Ranges) To UBound(Ranges)
        If IsObject(Ranges(i)) Then
            If Not Ranges(i) Is Nothing Then
                If TypeOf Ranges(i) Is Excel.Range Then
                    Dim rngUnion As Range
                    If Not rngUnion Is Nothing Then
                        Set rngUnion = Application.Union(rngUnion, Ranges(i))
                    Else
                        Set rngUnion = Ranges(i)
                    End If
                End If
            End If
        End If
    Next
    
    Set Union2 = rngUnion
    
End Function


Public Function ColumnLetterToNumber(ByVal strLetter As String) As Long
    ColumnLetterToNumber = ThisWorkbook.Worksheets(1).Columns(strLetter).Column
End Function


Public Function ColumnNumberToLetter( _
    ByVal lngNumber As Long, _
    Optional ByVal bAbsolute As Boolean = False _
        ) As String
    
    Dim sDummy As String
    sDummy = Split(ThisWorkbook.Worksheets(1).Columns(lngNumber).Address, ":")(0)
    
    If Not bAbsolute Then sDummy = Right$(sDummy, Len(sDummy) - 1)
    ColumnNumberToLetter = sDummy
    
End Function


'==============================================================================
'inspired by: <https://stackoverflow.com/a/21633724/5776000>
'1. to find out if name exists use 'Workbook' as 'Object'
'2. to find out if name is global scope do 1. and in addition
'   with 'Name' containing the prefix (i.e. "test.xlsm!") --> returns 'False'
'3. to find out if name is local scope do 1. and in addition
'   with 'Name' containing the prefix (i.e. "test.xlsm!") --> returns 'True'
Public Function DefinedNameExists( _
    ByVal Name As String, _
    Optional ByVal WorkbookOrWorksheet As Object _
        ) As Boolean
    
    If WorkbookOrWorksheet Is Nothing Then
        Dim Container As Object
        Set Container = ActiveWorkbook
    ElseIf (TypeOf WorkbookOrWorksheet Is Workbook) Or _
            (TypeOf WorkbookOrWorksheet Is Worksheet) _
    Then
        Set Container = WorkbookOrWorksheet
    Else
        GoTo errHandler
    End If
    
    On Error GoTo CleanExit:
    Dim Value As Variant
    Value = Container.Names(Name)
    On Error GoTo 0
    
'NOTE: this is an additional check (when the defined name exists)
'      --> add another optional argument?
    If Not InStr(1, CStr(Value), "#REF!") > 0 Then
        DefinedNameExists = True
        Exit Function
    End If
    
CleanExit:
    DefinedNameExists = False
    Exit Function
    
    
errHandler:
    MsgBox _
            "The only valid objects to search for defined " & _
            "names in are 'Workbooks' and 'Worksheets'."
    GoTo CleanExit
    
End Function

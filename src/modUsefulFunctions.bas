Attribute VB_Name = "modUsefulFunctions"

Option Explicit
Option Private Module
Option Base 1


'test, if a (normal or named) range exists
Public Function RangeExists( _
    ByVal wkb As Workbook, _
    ByVal RangeName As String _
        ) As Boolean
    
    RangeExists = False
    
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In wkb.Worksheets
        Dim rng As Range
        Set rng = ws.Range(RangeName)
        If Not rng Is Nothing Then
            RangeExists = True
            Exit Function
        End If
    Next
    On Error GoTo 0
    
End Function

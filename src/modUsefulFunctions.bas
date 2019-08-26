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

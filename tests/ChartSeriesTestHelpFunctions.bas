Attribute VB_Name = "ChartSeriesTestHelpFunctions"

'@Folder("ChartSeries.Tests")

Option Explicit
Option Private Module


Public Function OpenWorkbook( _
    ByVal FilePath As String, _
    ByVal FileName As String _
        ) As Workbook
    
    Dim FullFileName As String
    FullFileName = FilePath & "\" & FileName
    
    If Not DoesFileExist(FullFileName) Then Exit Function
    
    On Error Resume Next
    Dim wkb As Workbook
    Set wkb = Workbooks(FileName)
    
    If Not wkb Is Nothing Then
        If wkb.Path <> FilePath Then
            MsgBox "A file with the name '" & FileName & "' is already open, " & _
                    "but with another path."
        End If
    Else
        Set wkb = Workbooks.Open( _
                FileName:=FullFileName, _
                IgnoreReadOnlyRecommended:=True, _
                AddToMru:=False _
        )
    End If
    On Error GoTo 0
    
    Set OpenWorkbook = wkb
    
End Function


Private Function DoesFileExist( _
    ByVal FullFileName As String _
        ) As Boolean
    
    Dim DirFileName As String
    DirFileName = Dir(FullFileName)
    
    DoesFileExist = (Len(DirFileName) > 0)
    
End Function

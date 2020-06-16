VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("ChartSeries")
'@ModuleDescription("An object extending the 'Series' with the main purpose to access the 'Series.Formula' parts.")
'@PredeclaredId


'This Class module provides an easy way to access the items in a
'chart's SERIES formula. It can be exported and then imported into
'any project

'Developed by John Walkenbach, JWALK AND ASSOCIATES
'Copyright 1999. All rights reserved.
'May be used and distributed freely, but may not be sold.
'http://www.j-walk.com/ss/
'new website:
'<http://spreadsheetpage.com/index.php/file/a_class_module_to_manipulate_a_chart_series/>

'===================================
' PROPERTIES FOR CHARTSERIES OBJECT
'===================================

'Sub ExampleUsage()
'    '(with 'srs' being a 'Series' object)
'    Dim MySeries As IChartSeries
'    Set MySeries = ChartSeries.Create(srs)
'    With MySeries
'        If Not .IsSeriesAccessible Then Exit Sub
'
'        If .XValues.EntryType = eEntryType.eRange Then
'            MsgBox .XValues.RangeString & " with " & .NoOfPoints & " points"
'        Else
'            MsgBox .XValues.FormulaPart
'        End If
'    End With
'End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'2do:
'- "Modify an existing VBA class" --> Series
'  <https://stackoverflow.com/q/42590434>
'- add a 'Parent' property? (maybe other interesting stuff in this post)
'  <https://stackoverflow.com/a/9688292>
'- add a 'ChartSeriesCollection' class?
'  <https://stackoverflow.com/a/46936029>
'- idea for error numbers
'  <https://codereview.stackexchange.com/a/157444>
'  without having a look at it:
'  - make an enum with error numbers and create a dedicated 'RaiseError'
'    function in each class
'---
'- "Cloning Class Instances"
'  <http://dailydoseofexcel.com/archives/2012/10/09/cloning-class-instances/>
'---
'? is it possible to use a "Factory Method" when Chart.Type is known?
'  <https://stackoverflow.com/a/50968991>
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'- get data labels via '.DataLabel.Formula'
'  (see e.g. <https://www.get-digital-help.com/custom-data-labels-in-x-y-scatter-chart/>
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit

Implements IChartSeries

'==============================================================================
Private Const pcsSep As String = ","
'==============================================================================

Private Type TChartSeries
    Series As Series
    Chart As Chart
    IsSeriesAccessible As Boolean
    IsSeriesInBubbleChart As Boolean
    
    OldFullSeriesFormula As String
    FullSeriesFormula As String
    
    SeriesParts(eElement.[_First] To eElement.[_Last]) As ISeriesPart
    PlotOrder As ISeriesPlotOrder
    
    'helper stuff
    ChartWorkbookName As String
End Type
Private This As TChartSeries


'@Description("Factory method to create the extended 'Series' object (itself)")
Public Function Create(ByVal Value As Series) As IChartSeries
    With New ChartSeries
        Set .Series = Value
        .RefreshCachedProperties
        Set Create = .Self
    End With
End Function


'@Description("Sets the 'Series' object (itself)")
Friend Property Get Self() As IChartSeries
    Set Self = Me
End Property


'@Description("Sets the 'Series' object (itself)")
Friend Property Set Series(ByVal Value As Series)
    If Value Is Nothing Then
        RaiseErrorIsNothing
    Else
        Set This.Series = Value
    End If
End Property


'@Description("Return the full series formula ('ChartSeries')")
Private Property Get IChartSeries_FullFormula() As String
    With This
        If Not .IsSeriesAccessible Then
            RaiseErrorNotAccessible
        Else
            IChartSeries_FullFormula = This.FullSeriesFormula
        End If
    End With
End Property


Private Property Get IChartSeries_SeriesName() As ISeriesPart
    With This
        If .IsSeriesAccessible Then
            Set IChartSeries_SeriesName = This.SeriesParts(eElement.eName)
        Else
            RaiseErrorNotAccessible
        End If
    End With
End Property


Private Property Get IChartSeries_XValues() As ISeriesPart
    With This
        If .IsSeriesAccessible Then
            Set IChartSeries_XValues = .SeriesParts(eElement.eXValues)
        Else
            RaiseErrorNotAccessible
        End If
    End With
End Property


Private Property Get IChartSeries_Values() As ISeriesPart
    With This
        If .IsSeriesAccessible Then
            Set IChartSeries_Values = .SeriesParts(eElement.eYValues)
        Else
            RaiseErrorNotAccessible
        End If
    End With
End Property


Private Property Get IChartSeries_PlotOrder() As ISeriesPlotOrder
    With This
        If .IsSeriesAccessible Then
            Set IChartSeries_PlotOrder = _
                    SeriesPlotOrder.Create( _
                            .SeriesParts(eElement.ePlotOrder).FormulaPart _
                    )
        Else
            RaiseErrorNotAccessible
        End If
    End With
End Property


Private Property Get IChartSeries_BubbleSizes() As ISeriesPart
    With This
        If Not .IsSeriesAccessible Then
            RaiseErrorNotAccessible
        ElseIf .IsSeriesInBubbleChart Then
            Set IChartSeries_BubbleSizes = .SeriesParts(eElement.eBubbleSizes)
        Else
            RaiseErrorNotBubbleChart
        End If
    End With
End Property


'@Description("Returns the number of data points in the 'Series'")
Private Property Get IChartSeries_NoOfPoints() As Long
    IChartSeries_NoOfPoints = SERIESFormulaDataPoints()
End Property


'TODO: still needs to be refactored
'determine the source range of Point 'PointNumber' of SERIES
'the order of the points is as stated in the SERIES, not as shown in the graph
Private Property Get IChartSeries_PointSourceRange( _
    ByVal Element As Long, _
    ByVal PointNumber As Long _
        ) As Variant
    IChartSeries_PointSourceRange = PointRange(Element, PointNumber)
End Property


Private Function IChartSeries_IsSeriesAccessible() As Boolean
    IChartSeries_IsSeriesAccessible = This.IsSeriesAccessible
End Function


Private Function IChartSeries_IsSeriesInBubbleChart() As Boolean
    IChartSeries_IsSeriesInBubbleChart = This.IsSeriesInBubbleChart
End Function


'==============================================================================
Private Sub Class_Terminate()
    With This
        Dim Element As Long
        For Element = eElement.[_First] To eElement.[_Last]
            Set .SeriesParts(Element) = Nothing
        Next
        Set .PlotOrder = Nothing
        Set .Chart = Nothing
    End With
End Sub


'==============================================================================
'get meta data from Series
Friend Sub RefreshCachedProperties()
    
    GetChart
    GetChartWorkbookName
    GetIsSeriesInBubbleChart
    GetFullSeriesFormula
    GetIsSeriesAccessible
    
    If Not This.IsSeriesAccessible Then Exit Sub
    If Not HasFormulaChanged Then Exit Sub
    
    GetSeriesFormulaParts
    
End Sub


Private Sub GetChart()
    With This
        Set .Chart = .Series.Parent.Parent
    End With
End Sub


Private Sub GetChartWorkbookName()
    Dim ChartParent As Object
    Set ChartParent = This.Chart.Parent
    
    'if Chart sheet
    If TypeOf ChartParent Is Workbook Then
        This.ChartWorkbookName = ChartParent.Name
    'if embedded Chart (on Worksheet)
    Else
        Dim ChartWorksheet As Worksheet
        Set ChartWorksheet = ChartParent.Parent
        
        Dim ChartWorkbook As Workbook
        Set ChartWorkbook = ChartWorksheet.Parent
        
        This.ChartWorkbookName = ChartWorkbook.Name
    End If
End Sub


Private Sub GetIsSeriesInBubbleChart()
    With This
        Select Case .Series.ChartType
            Case xlBubble, xlBubble3DEffect
                .IsSeriesInBubbleChart = True
            Case Else
                .IsSeriesInBubbleChart = False
        End Select
    End With
End Sub


Private Sub GetFullSeriesFormula()
    
    With This
        .OldFullSeriesFormula = .FullSeriesFormula
        .FullSeriesFormula = vbNullString
        
        With .Chart
            'in case some data are currently hidden and thus the formula can't be
            'accessed, temporarily change '.PlotVisibleOnly' to 'False'
            Dim bPlotVisible As Boolean
            bPlotVisible = .PlotVisibleOnly
            .PlotVisibleOnly = False
        End With
        
        GetFullSeriesFormulaUsualAttempt
        GetFullSeriesFormulaSpecialAttempt
        
        .Chart.PlotVisibleOnly = bPlotVisible
    End With
    
End Sub


Private Sub GetFullSeriesFormulaUsualAttempt()
    With This
        'because there is a possibility that an error is risen, take care of that
        'situation
        '(inspired by <https://peltiertech.com/change-series-formula-improved-routines/>)
        On Error Resume Next
        .FullSeriesFormula = .Series.Formula
        On Error GoTo 0
    End With
End Sub


'temporarily change to column chart if series is inaccessible
Private Sub GetFullSeriesFormulaSpecialAttempt()
    
    'check, if special attempt not needed
    If Len(This.FullSeriesFormula) > 0 Then Exit Sub
    
    With This.Series
        'store current 'ChartType' and 'AxisGroup' of series ...
        Dim iChartType As XlChartType
        iChartType = .ChartType
        
        Dim iAxisGroup As XlAxisGroup
        iAxisGroup = .AxisGroup
        
        '... and change it temporarily to a column chart
        .ChartType = xlColumnClustered
        
        GetFullSeriesFormulaUsualAttempt
        
        'restore original 'ChartType' and 'AxisGroup'
        .ChartType = iChartType
        .AxisGroup = iAxisGroup
    End With
    
End Sub


Private Sub GetIsSeriesAccessible()
    With This
        .IsSeriesAccessible = (Len(.FullSeriesFormula) > 0)
    End With
End Sub


Private Function HasFormulaChanged() As Boolean
    With This
        HasFormulaChanged = (.OldFullSeriesFormula <> .FullSeriesFormula)
    End With
End Function


Private Sub GetSeriesFormulaParts()
    
    With This
        Dim MySeries As ISeriesFormulaParts
'NOTE: refactor so the second argument is not needed any more
        Set MySeries = SeriesFormulaParts.Create( _
                .FullSeriesFormula, _
                .IsSeriesInBubbleChart _
        )
    End With
    
    With MySeries
        Dim Element As eElement
        For Element = eElement.[_First] To eElement.[_Last]
            Set This.SeriesParts(Element) = _
                    SeriesPart.Create( _
                            This.ChartWorkbookName, _
                            .PartSeriesFormula(Element) _
                    )
        Next
    End With
    
    Set MySeries = Nothing
    
End Sub


'------------------------------------------------------------------------------
'returns the number of (data) points in the chart's SERIES
Private Function SERIESFormulaDataPoints() As Long
    With This
        SERIESFormulaDataPoints = .Series.Points.Count
    End With
End Function


'returns the range of the given point
Private Function PointRange( _
    ByVal Element As eElement, _
    ByVal PointNumber As Long _
        ) As Variant
    
    Dim sSeries As String
    Dim rng As Range
    Dim sWorksheetName As String
    Dim arRange() As String
    
    
    'exit if PointNumber is greater than the number of points existent
    'in 'SeriesNumber'
    If PointNumber > SERIESFormulaDataPoints() Then
        PointRange = "ERROR - BAD POINTNUMBER"
        Exit Function
    End If
    
    'exit if x or y values are not given as a range
    RefreshCachedProperties
    
    With This.SeriesParts(Element)
        If .EntryType <> eRange Then
'---
'TODO: raise an error (5?)
'      (otherwise is BubbleChart missing)
            If Element = eElement.eXValues Then
                PointRange = "ERROR - XVALUES ARE NOT A RANGE"
            ElseIf Element = eElement.eYValues Then
                PointRange = "ERROR - VALUES ARE NOT A RANGE"
            End If
'---
            Exit Function
        End If
        
        'store the 'Worksheet.Name'
        sWorksheetName = .RangeSheet
        
        'get the range string to parse for the address and store only the
        'range without the 'Worksheet.Name'
        Set rng = .Range
    End With
    sSeries = rng.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    'get an array with all cells of the SERIES
    arRange = SplitRangeToSingleCells(sSeries)
    
    PointRange = sWorksheetName & "!" & arRange(PointNumber - 1)
    
End Function


'NOTE: simplify using <https://stackoverflow.com/a/50062705/5776000>???
'
'function that returns an array where all cells of a range are separated
Private Function SplitRangeToSingleCells( _
    ByVal sRange As String _
        ) As Variant
    
    Dim arRange() As String
    Dim arRangeNew() As String
    Dim arCRange() As String
    Dim i As Long
    Dim j As Long
    Dim k As Long     'counter for number of elements
    Dim iStartRC As Long
    Dim iEndRC As Long
    Dim iRow As Long
    Dim sColumn As String
    
    
    'first split the 'Areas'
    arRange = Split(sRange, pcsSep)
    
    'initialize 'k'
    k = 0
    
    'now test, if in the areas are (continuous) ranges
    'if so, further split them
    For i = LBound(arRange) To UBound(arRange)
'NOTE: account for intersection ranges (space as separator)?
        If InStr(arRange(i), ":") Then
            'extract the numbers from the (continuous) range
            arCRange = Split(arRange(i), ":")
            iStartRC = OnlyDigits(arCRange(LBound(arCRange)))
            iEndRC = OnlyDigits(arCRange(UBound(arCRange)))
            'test if given row by row
            If iStartRC <> iEndRC Then
                'find column
                sColumn = Left$(arCRange(LBound(arCRange)), Len(arCRange(LBound(arCRange))) - Len(CStr(iStartRC)))
                
                'now create all the cells in the range
                If i = LBound(arRange) Then
                    ReDim arRangeNew(iEndRC - iStartRC)
                Else
                    ReDim Preserve arRangeNew(UBound(arRangeNew) + iEndRC - iStartRC + 1)
                End If
                For j = iStartRC To iEndRC
                    arRangeNew(k) = sColumn & CStr(j)
                    k = k + 1
                Next
            'else given column by column
            Else
                'store row number
                iRow = iStartRC
                'extract column (numbers)
                sColumn = Left$(arCRange(LBound(arCRange)), Len(arCRange(LBound(arCRange))) - Len(CStr(iRow)))
                iStartRC = ColumnLetterToNumber(sColumn)
                sColumn = Left$(arCRange(UBound(arCRange)), Len(arCRange(UBound(arCRange))) - Len(CStr(iRow)))
                iEndRC = ColumnLetterToNumber(sColumn)
                
                'now create all the cells in the range
                If i = LBound(arRange) Then
                    ReDim arRangeNew(iEndRC - iStartRC)
                Else
                    ReDim Preserve arRangeNew(UBound(arRangeNew) + iEndRC - iStartRC + 1)
                End If
                For j = iStartRC To iEndRC
                    arRangeNew(k) = ColumnNumberToLetter(j) & CStr(iRow)
                    k = k + 1
                Next
            End If
        Else
            If i = LBound(arRange) Then
                ReDim arRangeNew(k)
            Else
                ReDim Preserve arRangeNew(k)
            End If
            arRangeNew(k) = arRange(i)
            k = k + 1
        End If
    Next
    
    SplitRangeToSingleCells = arRangeNew
    
End Function


'==============================================================================
'inspired by <https://stackoverflow.com/a/7239408/5776000>
Private Function OnlyDigits( _
    sText As String _
        ) As Long
    
    Dim Digits As String
    Digits = vbNullString
    
    Dim i As Long
    For i = 1 To Len(sText)
'NOTE: test if these two tests are faster than 'If IsNumeric(sText, i, 1)'
        If Mid(sText, i, 1) >= "0" Then
            If Mid(sText, i, 1) <= "9" Then
                Digits = Digits + Mid(sText, i, 1)
            End If
        End If
    Next
    
    OnlyDigits = CLng(Digits)
    
End Function


'=============================================================================
Private Sub RaiseErrorIsNothing()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eChartSeriesError.ErrIsNothing, _
            Description:= _
                    "The 'Series' is 'Nothing'." & _
                    vbCrLf & _
                    "Please check beforehand for that."
End Sub


Private Sub RaiseErrorNotAccessible()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eChartSeriesError.ErrNotAccessible, _
            Description:= _
                    "The 'Series.Formula' is not accessible." & _
                    vbCrLf & _
                    "Hence, you can't access any property." & _
                    vbCrLf & _
                    "Please first check with '.IsSeriesAccessible'."
End Sub


Private Sub RaiseErrorNotBubbleChart()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eChartSeriesError.ErrNotInBubbleChart, _
            Description:= _
                    "The Series is not in a Bubble Chart." & _
                    vbCrLf & _
                    "Please first check with '.IsSeriesInBubbleChart'."
End Sub

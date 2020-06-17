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

'Sub ExampleUsage()
'    '(with 'srs' being a 'Series' object)
'    Dim MySeries As IChartSeries
'    Set MySeries = ChartSeries.Create(srs)
'    With MySeries
'        If Not .IsSeriesAccessible Then Exit Sub
'
'        If .XValues.EntryType = eEntryType.eRange Then
'            MsgBox .XValues.RangeString
'        Else
'            MsgBox .XValues.FormulaPart
'        End If
'    End With
'End Sub

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
    If Value Is Nothing Then RaiseErrorIsNothing
    
    Set This.Series = Value
End Property


'@Description("Return the full series formula ('ChartSeries')")
Private Property Get IChartSeries_FullFormula() As String
    With This
        If Not .IsSeriesAccessible Then RaiseErrorNotAccessible
        
        IChartSeries_FullFormula = .FullSeriesFormula
    End With
End Property


Private Property Get IChartSeries_SeriesName() As ISeriesPart
    With This
        If Not .IsSeriesAccessible Then RaiseErrorNotAccessible
        
        Set IChartSeries_SeriesName = .SeriesParts(eElement.eName)
    End With
End Property


Private Property Get IChartSeries_XValues() As ISeriesPart
    With This
        If Not .IsSeriesAccessible Then RaiseErrorNotAccessible
        
        Set IChartSeries_XValues = .SeriesParts(eElement.eXValues)
    End With
End Property


Private Property Get IChartSeries_Values() As ISeriesPart
    With This
        If Not .IsSeriesAccessible Then RaiseErrorNotAccessible
        
        Set IChartSeries_Values = .SeriesParts(eElement.eYValues)
    End With
End Property


Private Property Get IChartSeries_PlotOrder() As ISeriesPlotOrder
    With This
        If Not .IsSeriesAccessible Then RaiseErrorNotAccessible
        
        Set IChartSeries_PlotOrder = _
                SeriesPlotOrder.Create( _
                        .SeriesParts(eElement.ePlotOrder).FormulaPart _
                )
    End With
End Property


Private Property Get IChartSeries_BubbleSizes() As ISeriesPart
    With This
        If Not .IsSeriesAccessible Then RaiseErrorNotAccessible
        If Not .IsSeriesInBubbleChart Then RaiseErrorNotBubbleChart
        
        Set IChartSeries_BubbleSizes = .SeriesParts(eElement.eBubbleSizes)
    End With
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
    
    If TypeOf ChartParent Is Workbook Then
        GetWorkbookNameOfChartSheet ChartParent
    Else
        GetWorkbookNameOfChartObject ChartParent
    End If
End Sub


Private Sub GetWorkbookNameOfChartSheet( _
    ByVal ChartParent As Workbook _
)
    This.ChartWorkbookName = ChartParent.Name
End Sub


Private Sub GetWorkbookNameOfChartObject( _
    ByVal ChartParent As ChartObject _
)
    Dim ChartWorksheet As Worksheet
    Set ChartWorksheet = ChartParent.Parent
    
    Dim ChartWorkbook As Workbook
    Set ChartWorkbook = ChartWorksheet.Parent
    
    This.ChartWorkbookName = ChartWorkbook.Name
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

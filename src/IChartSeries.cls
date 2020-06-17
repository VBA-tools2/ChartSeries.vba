VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IChartSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an extension of the 'Series' object."

'@Folder("ChartSeries")
'@ModuleDescription("Describes an extension of the 'Series' object.")
'@Interface
'@Exposed

Option Explicit

Public Enum eChartSeriesError
    [_First] = vbObjectError + 1
    ErrIsNothing = [_First]
    ErrNotAccessible
    ErrNotInBubbleChart
    [_Last] = ErrNotInBubbleChart
End Enum

'@Description("Returns the full series formula")
Public Property Get FullFormula() As String
Attribute FullFormula.VB_Description = "Returns the full series formula"
End Property

'@Description("Returns the series 'Name' extended object")
Public Property Get SeriesName() As ISeriesPart
Attribute SeriesName.VB_Description = "Returns the series 'Name' extended object"
End Property

'@Description("Returns the series 'XValues' extended object")
Public Property Get XValues() As ISeriesPart
Attribute XValues.VB_Description = "Returns the series 'XValues' extended object"
End Property

'@Description("Returns the series 'Values' extended object")
Public Property Get Values() As ISeriesPart
Attribute Values.VB_Description = "Returns the series 'Values' extended object"
End Property

'@Description("Returns the series 'PlotOrder' object")
Public Property Get PlotOrder() As ISeriesPlotOrder
Attribute PlotOrder.VB_Description = "Returns the series 'PlotOrder' object"
End Property

'@Description("Returns the series 'BubbleSize' extended object")
Public Property Get BubbleSizes() As ISeriesPart
Attribute BubbleSizes.VB_Description = "Returns the series 'BubbleSize' extended object"
End Property

'@Description("Returns 'True' if the series is accessible")
Public Function IsSeriesAccessible() As Boolean
Attribute IsSeriesAccessible.VB_Description = "Returns 'True' if the series is accessible"
End Function

'@Description("Returns 'True' if the series is in a Bubble Chart")
Public Function IsSeriesInBubbleChart() As Boolean
Attribute IsSeriesInBubbleChart.VB_Description = "Returns 'True' if the series is in a Bubble Chart"
End Function

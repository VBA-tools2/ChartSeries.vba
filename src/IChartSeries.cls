VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IChartSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'@Folder("ChartSeries")
'@ModuleDescription("Describes an extension of the 'Series' object.")
'@Interface
'@Exposed

Option Explicit

'@Description("Returns the full series formula")
Public Property Get FullFormula() As String
End Property

'@Description("Returns the series 'Name' extended object")
Public Property Get SeriesName() As ISeriesPart
End Property

'@Description("Returns the series 'XValues' extended object")
Public Property Get XValues() As ISeriesPart
End Property

'@Description("Returns the series 'Values' extended object")
Public Property Get Values() As ISeriesPart
End Property

'@Description("Returns the series 'PlotOrder' object")
Public Property Get PlotOrder() As ISeriesPlotOrder
End Property

'@Description("Returns the series 'BubbleSize' extended object")
Public Property Get BubbleSizes() As ISeriesPart
End Property

Public Property Get NoOfPoints() As Long
End Property

Public Property Get PointSourceRange( _
    ByVal Element As Long, _
    ByVal PointNumber As Long _
        ) As Variant
End Property

'@Description("Returns 'True' if the series is accessible")
Public Function IsSeriesAccessible() As Boolean
End Function

'@Description("Returns 'True' if the series is in a Bubble Chart")
Public Function IsSeriesInBubbleChart() As Boolean
End Function

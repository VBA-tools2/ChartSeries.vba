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
'@Interface
'@Exposed

Option Explicit

'@Description("Return the full series formula ('IChartSeries')")
Public Property Get FullFormula() As String
End Property

Public Property Get SeriesName() As ISeriesPart
End Property

Public Property Get XValues() As ISeriesPart
End Property

Public Property Get Values() As ISeriesPart
End Property

Public Property Get PlotOrder() As ISeriesPlotOrder
End Property

Public Property Get BubbleSizes() As ISeriesPart
End Property

Public Property Get NoOfPoints() As Long
End Property

Public Property Get PointSourceRange( _
    ByVal Element As Long, _
    ByVal PointNumber As Long _
        ) As Variant
End Property

Public Function IsSeriesAccessible() As Boolean
End Function

Public Function IsSeriesInBubbleChart() As Boolean
End Function

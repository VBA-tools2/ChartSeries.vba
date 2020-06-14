VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISeriesPlotOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'@Folder("ChartSeries")
'@ModuleDescription("Describes an object of the series formula part 'PlotOrder'.")
'@Interface
'@Exposed

Option Explicit

Public Enum eSeriesPlotOrderError
    ErrNotNumericFormulaPart = vbObjectError + 1
End Enum

'@Description("Returns the type of the series formula part (e.g. Empty, Range, String)")
Public Property Get EntryType() As eEntryType
End Property

'@Description("Returns the PlotOrder value")
Public Property Get Value() As Byte
End Property

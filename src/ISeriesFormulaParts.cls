VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISeriesFormulaParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'@Folder("ChartSeries")
'@Interface
'@Exposed

Option Explicit

Public Enum eElement
    [_First] = 1
    eName = eElement.[_First]
    eXValues
    eYValues
    ePlotOrder
    eBubbleSizes
    [_Last] = eElement.eBubbleSizes
End Enum

'@Description("fill me")
Public Property Get PartSeriesFormula(ByVal Element As eElement) As String
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISeriesFormulaParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object to separate the string parts of the 'Series.Formula'."

'@Folder("ChartSeries")
'@ModuleDescription("Describes an object to separate the string parts of the 'Series.Formula'.")
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

Public Enum eSeriesFormulaPartsError
    ErrFullSeriesFormulaEmpty = vbObjectError + 1
End Enum

'@Description("Returns the full series formula")
Public Property Get PartSeriesFormula(ByVal Element As eElement) As String
Attribute PartSeriesFormula.VB_Description = "Returns the full series formula"
End Property

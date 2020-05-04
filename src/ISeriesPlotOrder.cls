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
'@Interface
'@Exposed

Option Explicit

Public Property Let FormulaPart(ByVal Value As String)
End Property

Public Property Get EntryType() As eEntryType
End Property

Public Property Get Value() As Byte
End Property

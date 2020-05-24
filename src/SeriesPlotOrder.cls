VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeriesPlotOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("ChartSeries")
'@ModuleDescription("An object that only handles the 'PlotOrder' part of a Series.")
'@PredeclaredId

Option Explicit

Implements ISeriesPlotOrder

Private Type TSeriesPart
    FormulaPart As String
    EntryType As eEntryType
    Value As Byte
End Type
Private This As TSeriesPart


'@Description("Returns the series formula part")
Private Property Get FormulaPart() As String
    FormulaPart = This.FormulaPart
End Property


Private Property Let ISeriesPlotOrder_FormulaPart(ByVal Value As String)
    This.FormulaPart = Value
    Run
End Property


Private Property Get ISeriesPlotOrder_EntryType() As eEntryType
    ISeriesPlotOrder_EntryType = This.EntryType
End Property


Private Property Get ISeriesPlotOrder_Value() As Byte
    ISeriesPlotOrder_Value = This.Value
End Property


'==============================================================================
Private Sub Class_Initialize()
    This.EntryType = eEntryType.eNotSet
End Sub


'==============================================================================
Private Sub Run()
    With This
        If Not IsNumeric(.FormulaPart) Then
'NOTE: refactor Description
            Err.Raise _
                Number:=vbObjectError + 10, _
                Description:="It should never happen that we land here"
            Exit Sub
        End If
        
        .EntryType = eInteger
        .Value = CByte(.FormulaPart)
    End With
End Sub

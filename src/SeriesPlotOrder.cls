VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeriesPlotOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@PredeclaredId
'@Folder("ChartSeries")
'@ModuleDescription("A class to handle each part of the SERIES string.")

Option Explicit

Implements ISeriesPlotOrder

Private Type TSeriesPart
    FormulaPart As String
    EntryType As eEntryType
    Value As Byte
End Type
Private This As TSeriesPart


Private Property Get FormulaPart() As String
    FormulaPart = This.FormulaPart
End Property


'@Description("...")
Private Property Let ISeriesPlotOrder_FormulaPart(ByVal Value As String)
    This.FormulaPart = Value
    Run
End Property


'@Description("Returns the Enum of the 'PartType'")
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


'Private Function IsSeriesValid( _
'    ByVal Value As Byte _
'        ) As Boolean
'    IsSeriesValid = (Value >= 0)
'End Function

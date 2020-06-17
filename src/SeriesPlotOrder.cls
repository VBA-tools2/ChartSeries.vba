VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeriesPlotOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that only handles the 'PlotOrder' part of a Series."

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


'@Description("Factory method to create this object")
Public Function Create( _
    ByVal FormulaPart As String _
        ) As ISeriesPlotOrder
Attribute Create.VB_Description = "Factory method to create this object"
    
    With New SeriesPlotOrder
        .FormulaPart = FormulaPart
        .RefreshCachedProperties
        Set Create = .Self
    End With
    
End Function


'@Description("Returns the object (itself)")
Friend Property Get Self() As ISeriesPlotOrder
Attribute Self.VB_Description = "Returns the object (itself)"
    Set Self = Me
End Property


'@Description("Sets the series formula part")
Friend Property Let FormulaPart(ByVal Value As String)
Attribute FormulaPart.VB_Description = "Sets the series formula part"
    This.FormulaPart = Value
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
Friend Sub RefreshCachedProperties()
    With This
        If Not IsNumeric(.FormulaPart) Then RaiseErrorNotNumericFormulaPart
        
        .EntryType = eInteger
        .Value = CByte(.FormulaPart)
    End With
End Sub


'==============================================================================
Private Sub RaiseErrorNotNumericFormulaPart()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eSeriesPlotOrderError.ErrNotNumericFormulaPart, _
            Description:= _
                    "This error should never be raised." & vbCrLf & _
                    "It seems you have found either an unhandled case " & _
                    "or a bug in the 'SeriesFormulaParts' class."
End Sub

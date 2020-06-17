VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISeriesPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object of the series formula part except the 'PlotOrder'."

'@Folder("ChartSeries")
'@ModuleDescription("Describes an object of the series formula part except the 'PlotOrder'.")
'@Interface
'@Exposed

Option Explicit

Public Enum eEntryType
    eNotSet = -1
    [_First] = 0
    eEmpty = eEntryType.[_First]
    eString
    eArray
    eInteger
    eRange
    eDefinedName
    [_Last] = eEntryType.eRange
End Enum

Public Enum eSeriesPartError
    ErrChartWorkbookNameEmpty = vbObjectError + 1
    ErrNotARange
    ErrNotADefinedName
End Enum

'@Description("Returns the series formula part")
Public Property Get FormulaPart() As String
Attribute FormulaPart.VB_Description = "Returns the series formula part"
End Property

'@Description("Returns the series formula part without surrounding characters")
Public Property Get CleanFormulaPart() As String
Attribute CleanFormulaPart.VB_Description = "Returns the series formula part without surrounding characters"
End Property

'@Description("Returns the type of the series formula part (e.g. Empty, Range, String)")
'"Type" is a keyword in VBA and therefore is not a valid variable name
Public Property Get EntryType() As eEntryType
Attribute EntryType.VB_Description = "Returns the type of the series formula part (e.g. Empty, Range, String)"
End Property

'@Description("Returns the range of the series formula part")
Public Property Get Range() As Range
Attribute Range.VB_Description = "Returns the range of the series formula part"
End Property

'@Description("Returns the cell part of the range")
Public Property Get RangeString() As String
Attribute RangeString.VB_Description = "Returns the cell part of the range"
End Property

'@Description("Returns the worksheet part of the range")
Public Property Get RangeSheet() As String
Attribute RangeSheet.VB_Description = "Returns the worksheet part of the range"
End Property

'@Description("Returns the workbook part of the range")
Public Property Get RangeBook() As String
Attribute RangeBook.VB_Description = "Returns the workbook part of the range"
End Property

'@Description("Returns the path part of the range")
Public Property Get RangePath() As String
Attribute RangePath.VB_Description = "Returns the path part of the range"
End Property

'@Description("Returns 'True' if the series formula part is a range")
Public Function IsRange() As Boolean
Attribute IsRange.VB_Description = "Returns 'True' if the series formula part is a range"
End Function

'@Description("Returns 'True' if the series formula part is a defined name range")
Public Function IsDefinedNameRange() As Boolean
Attribute IsDefinedNameRange.VB_Description = "Returns 'True' if the series formula part is a defined name range"
End Function

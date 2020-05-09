VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISeriesPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'@Folder("ChartSeries")
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

Public Property Get FormulaPart() As String
End Property

Public Property Get CleanFormulaPart() As String
End Property

'"Type" is a keyword in VBA and therefore is not a valid variable name
Public Property Get EntryType() As eEntryType
End Property

Public Property Get Range() As Range
End Property

Public Property Get RangeString() As String
End Property

Public Property Get RangeSheet() As String
End Property

Public Property Get RangeBook() As String
End Property

Public Property Get RangePath() As String
End Property

Public Function IsDefinedNameRange() As Boolean
End Function

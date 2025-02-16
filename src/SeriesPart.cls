VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeriesPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object to handle each part of the SERIES string."

'@Folder("ChartSeries")
'@ModuleDescription("An object to handle each part of the SERIES string.")
'@PredeclaredId

Option Explicit

Implements ISeriesPart


Private Type TSeriesPart
    'Properties, Functions, Subs
    FormulaPart As String
    CleanFormulaPart As String
    EntryType As eEntryType
    IsRange As Boolean
    Range As Range
    
    RangeString As String
    RangeSheet As String
    RangeBook As String
    RangePath As String
    
    'helper stuff
    ChartWorkbookName As String
    HasSurroundingChars As Boolean
    IsSingleAreaRange As Boolean
    WkbWksPart As String
    CleanWkbWksPart As String
    RawRangeString As String
    
    IsDefinedNameRange As Boolean
End Type
Private This As TSeriesPart


'@Description("Factory method to create this object")
Public Function Create( _
    ByVal ChartWorkbookName As String, _
    ByVal FormulaPart As String _
        ) As ISeriesPart
Attribute Create.VB_Description = "Factory method to create this object"
    
    With New SeriesPart
        .ChartWorkbookName = ChartWorkbookName
        .FormulaPart = FormulaPart
        .RefreshCachedProperties
        Set Create = .Self
    End With
    
End Function


'@Description("Returns the object (itself)")
Friend Property Get Self() As ISeriesPart
Attribute Self.VB_Description = "Returns the object (itself)"
    Set Self = Me
End Property


'@Description("Stores the 'Workbook.Name' where the chart is placed in")
Friend Property Let ChartWorkbookName(ByVal Value As String)
Attribute ChartWorkbookName.VB_Description = "Stores the 'Workbook.Name' where the chart is placed in"
    If Len(Value) = 0 Then RaiseErrorChartWorkbookNameEmpty
    
    This.ChartWorkbookName = Value
End Property


'@Description("Sets the series formula part")
Friend Property Let FormulaPart(ByVal Value As String)
Attribute FormulaPart.VB_Description = "Sets the series formula part"
    This.FormulaPart = Value
End Property


Private Property Get ISeriesPart_FormulaPart() As String
    ISeriesPart_FormulaPart = This.FormulaPart
End Property


Private Property Get ISeriesPart_CleanFormulaPart() As String
    ISeriesPart_CleanFormulaPart = This.CleanFormulaPart
End Property


Private Property Get ISeriesPart_EntryType() As eEntryType
    ISeriesPart_EntryType = This.EntryType
End Property


Private Property Get ISeriesPart_Range() As Range
    Set ISeriesPart_Range = This.Range
End Property


Private Property Get ISeriesPart_RangeString() As String
    With This
        If Not .IsRange Then RaiseErrorNotARange
        
        ISeriesPart_RangeString = .RangeString
    End With
End Property


Private Property Get ISeriesPart_RangeSheet() As String
    With This
        If Not .IsRange Then RaiseErrorNotARange
        
        ISeriesPart_RangeSheet = .RangeSheet
    End With
End Property


Private Property Get ISeriesPart_RangeBook() As String
    With This
        If Not .IsRange Then RaiseErrorNotARange
        
        ISeriesPart_RangeBook = .RangeBook
    End With
End Property


Private Property Get ISeriesPart_RangePath() As String
    With This
        If Not .IsRange Then RaiseErrorNotARange
        
        ISeriesPart_RangePath = .RangePath
    End With
End Property


Private Function ISeriesPart_IsRange() As Boolean
    ISeriesPart_IsRange = This.IsRange
End Function


Private Function ISeriesPart_IsDefinedNameRange() As Boolean
    With This
        If Not .EntryType = eDefinedName Then RaiseErrorNotADefinedName
        
        ISeriesPart_IsDefinedNameRange = .IsDefinedNameRange
    End With
End Function


'==============================================================================
Private Sub Class_Initialize()
    This.EntryType = eEntryType.eNotSet
End Sub


'==============================================================================
Friend Sub RefreshCachedProperties()
    
    SetEntryTypes
    
    GetIsSingleAreaRange
    GetHasSurroundingChars
    GetCleanFormulaPart
    
    With This
        If .EntryType = eRange Then
            HandleRangeCase
        ElseIf .EntryType = eDefinedName Then
            HandleDefinedNameCase
        Else
            Set .Range = Nothing
        End If
    End With
    
    GetIsRange
    
End Sub


Private Sub SetEntryTypes()
    
    With This
        Dim FormulaPartString As String
        FormulaPartString = .FormulaPart
        
        If Len(FormulaPartString) = 0 Then               'Empty
            .EntryType = eEmpty
        ElseIf Left$(FormulaPartString, 1) = """" Then   'String
            .EntryType = eString
        ElseIf Left$(FormulaPartString, 1) = "{" Then    'Literal Array
            .EntryType = eArray
        ElseIf IsNumeric(FormulaPartString) Then         'Integer (for PlotOrder)
            .EntryType = eInteger
        ElseIf Left$(FormulaPartString, 1) = "(" Then    'Multi Area Range
            .EntryType = eRange
        ElseIf IsDefinedName(FormulaPartString) Then     'Defined Name
            .EntryType = eDefinedName
        Else                                             'Single Area Range
            .EntryType = eRange
        End If
    End With
    
End Sub


Private Function IsDefinedName( _
    ByVal FormulaPartString As String _
        ) As Boolean
    
    Dim LastExclMarkPos As Long
    LastExclMarkPos = InStrRev(FormulaPartString, "!")
    
    Dim RelevantString As String
    RelevantString = Right$(FormulaPartString, Len(FormulaPartString) - LastExclMarkPos)
    
    IsDefinedName = (InStr(RelevantString, "$") = 0)
    
End Function


Private Sub GetIsSingleAreaRange()
    With This
        If .EntryType = eRange Then
            .IsSingleAreaRange = (Left$(.FormulaPart, 1) <> "(")
        End If
    End With
End Sub


Private Sub GetHasSurroundingChars()
    With This
        Select Case .EntryType
            Case eString, eArray
                .HasSurroundingChars = True
            Case eEmpty, eInteger, eDefinedName
                .HasSurroundingChars = False
            Case Else
                .HasSurroundingChars = Not .IsSingleAreaRange
        End Select
    End With
End Sub


Private Sub GetCleanFormulaPart()
    If This.HasSurroundingChars Then
        GetCleanFormulaPartWithSurroundingChars
    Else
        GetCleanFormulaPartWithoutSurroundingChars
    End If
End Sub


Private Sub GetCleanFormulaPartWithoutSurroundingChars()
    With This
        .CleanFormulaPart = .FormulaPart
    End With
End Sub


Private Sub GetCleanFormulaPartWithSurroundingChars()
    Const vbDoubleQuote As String = """"
    With This
        .CleanFormulaPart = .FormulaPart
        If Left$(.FormulaPart, 1) = """" Then
            .CleanFormulaPart = Replace(.CleanFormulaPart, vbDoubleQuote & vbDoubleQuote, vbDoubleQuote)
        End If
        .CleanFormulaPart = Mid$(.CleanFormulaPart, 2, Len(.CleanFormulaPart) - 2)
    End With
End Sub


Private Sub HandleRangeCase()
    GetWkbWksPart
    GetCleanWkbWksPart
    GetRawRangeString
    GetRangeString
    GetRangeWorksheet
    GetRangeWorkbook
    GetRangePath
    
    SetRange
End Sub


Private Sub GetWkbWksPart()
    If Left$(This.CleanFormulaPart, 1) = "'" Then
        HandleWkbWksPartWithSingleQuotes
    Else
        HandleWkbWksPartWithoutSingleQuotes
    End If
End Sub


Private Sub HandleWkbWksPartWithSingleQuotes()
    
    Const CriticalString As String = "''!"
    'replace with a char that is not allowed in 'Worksheet.Name's ...
    Const ReplacementString As String = "?"
    
    '... in case there is a "'!" in a Worksheet name
    Dim TempString As String
    TempString = Replace(This.CleanFormulaPart, CriticalString, ReplacementString)
    
    Dim SeparatorPosition As Long
    SeparatorPosition = InStr(TempString, "'!")
    
    Dim sWkbWksPart As String
    sWkbWksPart = Left$(TempString, SeparatorPosition)
    
    This.WkbWksPart = Replace(sWkbWksPart, ReplacementString, CriticalString)
    
End Sub


Private Sub HandleWkbWksPartWithoutSingleQuotes()
    With This
        Dim SeparatorPosition As Long
        SeparatorPosition = InStr(.CleanFormulaPart, "!")
        
        .WkbWksPart = Left$(.CleanFormulaPart, SeparatorPosition - 1)
    End With
End Sub


Private Sub GetCleanWkbWksPart()
    With This
        If Left$(.WkbWksPart, 1) = "'" Then
            .CleanWkbWksPart = Mid$(.WkbWksPart, 2, Len(.WkbWksPart) - 2)
        Else
            .CleanWkbWksPart = .WkbWksPart
        End If
    End With
End Sub


Private Sub GetRawRangeString()
    If This.IsSingleAreaRange Then
        GetSingleAreaRawRangeString
    Else
        GetMultiAreaRawRangeString
    End If
End Sub


Private Sub GetSingleAreaRawRangeString()
    
    Dim SingleAreaString As String
    SingleAreaString = This.CleanFormulaPart
    
    Dim StringLength As Long
    StringLength = Len(SingleAreaString)
    
    Dim LastExclMarkPosition As Long
    LastExclMarkPosition = InStrRev(SingleAreaString, "!")
    
    This.RawRangeString = _
            Right$(SingleAreaString, StringLength - LastExclMarkPosition)
    
End Sub


Private Sub GetMultiAreaRawRangeString()
    With This
        .RawRangeString = Replace( _
                .CleanFormulaPart, _
                .WkbWksPart & "!", _
                vbNullString _
        )
    End With
End Sub


Private Sub GetRangeString()
    With This
        .RangeString = Replace(.RawRangeString, "$", vbNullString)
    End With
End Sub


Private Sub GetRangeWorksheet()
    
    Dim sWkbWks As String
    sWkbWks = This.CleanWkbWksPart
    
    Dim StringLength As Long
    StringLength = Len(sWkbWks)
    
    Dim BracketPosition As Long
    BracketPosition = InStr(sWkbWks, "]")
    
    Dim TempString As String
    TempString = Right$(sWkbWks, StringLength - BracketPosition)
    
    'SingleQuotes are shown as double SingleQuotes in the SeriesFormula
    '(when they are part of a Worksheet.Name or string)
    This.RangeSheet = Replace(TempString, "''", "'")
    
End Sub


Private Sub GetRangeWorkbook()
    
    With This
        Dim ClosingBracketPosition As Long
        ClosingBracketPosition = InStr(.CleanWkbWksPart, "]")
        
        If ClosingBracketPosition = 0 Then
            .RangeBook = .ChartWorkbookName
        Else
            Dim OpeningBracketPosition As Long
            OpeningBracketPosition = InStr(.CleanWkbWksPart, "[") + 1
            
            Dim UncleanedBookName As String
            UncleanedBookName = _
                    Mid$( _
                            .CleanWkbWksPart, _
                            OpeningBracketPosition, _
                            ClosingBracketPosition - OpeningBracketPosition _
                    )
            
            .RangeBook = Replace$(UncleanedBookName, "''", "'")
        End If
    End With
    
End Sub


Private Sub GetRangePath()
    
    With This
        Dim LastBackslashPosition As Long
        LastBackslashPosition = InStrRev(.CleanWkbWksPart, "\")
        
        If LastBackslashPosition > 0 Then
            .RangePath = Left$(.CleanWkbWksPart, LastBackslashPosition - 1)
        End If
    End With
    
End Sub


Private Sub SetRange()
    
    With This
        If IsExternalRangeClosed Then
            HandleClosedExternalRange
        ElseIf Len(.FormulaPart) <= 255 Then
            HandleShortRangeCase
        Else
            HandleLongRangeCase
        End If
    End With
    
End Sub


'a backslash (\) is neither allowed in range names nor in worksheet names
Private Function IsExternalRangeClosed() As Boolean
    IsExternalRangeClosed = (InStr(This.FormulaPart, "\") > 0)
End Function


'external ranges can crash Excel (2010) (at least when the file is not open)
'thus only set the range when the source file is open
Private Sub HandleClosedExternalRange()
End Sub


Private Sub HandleShortRangeCase()
    
    With This
        Dim wkb As Workbook
        Set wkb = Workbooks(.RangeBook)
        
        Dim wks As Worksheet
        Set wks = wkb.Worksheets(.RangeSheet)
        
        Set .Range = wks.Range(.RangeString)
    End With
    
End Sub


Private Sub HandleLongRangeCase()
    
    With This
        Dim wkb As Workbook
        Set wkb = Workbooks(.RangeBook)
        
        Dim wks As Worksheet
        Set wks = wkb.Worksheets(.RangeSheet)
        
        If Len(.RangeString) <= 255 Then
            Set .Range = wks.Range(.RangeString)
        Else
            Dim arrRanges() As String
            arrRanges = Split(.RangeString, ",")
            
            Dim i As Long
            For i = LBound(arrRanges) To UBound(arrRanges)
                Dim rngUnion As Range
                Set rngUnion = Union2(rngUnion, wks.Range(arrRanges(i)))
            Next
            
            Set .Range = rngUnion
        End If
    End With
    
End Sub


'==============================================================================
Private Sub HandleDefinedNameCase()
    
    'multiple areas with defined names seem not to be allowed
    GetSingleAreaRawRangeString
    GetRangeString
    
    GetWkbWksPart
    GetCleanWkbWksPart
    
    GetRangeWorksheet
    GetRangeWorkbook
    GetRangePath
    
    'closed external global scope defined names lead to
    'an inaccessible 'Series.Formula' (in Excel 2010)
    If Not IsLocalScopeDefinedName Then
        With This
            .RangeBook = .RangeSheet
            .RangeSheet = vbNullString
        End With
    End If
    
    TrySettingRangeFromDefinedName
    GetIsDefinedNameRange
    
End Sub


Private Function IsLocalScopeDefinedName() As Boolean
    If IsExternalLocalScopeDefinedName Then
        IsLocalScopeDefinedName = True
    ElseIf IsInternalLocalScopeDefinedName Then
        IsLocalScopeDefinedName = True
    Else
        IsLocalScopeDefinedName = False
    End If
End Function


Private Function IsExternalLocalScopeDefinedName() As Boolean
    IsExternalLocalScopeDefinedName = (InStr(This.CleanWkbWksPart, "]") > 0)
End Function


Private Function IsInternalLocalScopeDefinedName() As Boolean
    With This
        IsInternalLocalScopeDefinedName = _
                DefinedNameExists( _
                        .CleanFormulaPart, _
                        Workbooks(.ChartWorkbookName) _
                )
    End With
End Function


Private Sub TrySettingRangeFromDefinedName()
    With This
        'can't be checked for closed workbooks
        If Len(.RangePath) > 0 Then
            Set .Range = Nothing
            Exit Sub
        End If
        
        On Error Resume Next
        If Len(.RangeSheet) > 0 Then
            Set .Range = Workbooks(.RangeBook).Worksheets(.RangeSheet).Range(.RangeString)
        Else
            Set .Range = Application.Range(.FormulaPart)
        End If
        On Error GoTo 0
    End With
End Sub


Private Sub GetIsDefinedNameRange()
    With This
        .IsDefinedNameRange = (Not .Range Is Nothing)
    End With
End Sub


Private Sub GetIsRange()
    With This
        If .EntryType = eRange Then
            .IsRange = True
        ElseIf .IsDefinedNameRange Then
            .IsRange = True
        Else
            .IsRange = False
        End If
    End With
End Sub


'==============================================================================
Private Sub RaiseErrorChartWorkbookNameEmpty()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eSeriesPartError.ErrChartWorkbookNameEmpty, _
            Description:= _
                    "This error should never be raised." & vbCrLf & _
                    "The 'ChartWorkbookName' is empty, which is impossible."
End Sub


Private Sub RaiseErrorNotARange()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eSeriesPartError.ErrNotARange, _
            Description:= _
                    "This element is neither a range nor a defined name." & vbCrLf & _
                    "Check with property 'EntryType' first."
End Sub


Private Sub RaiseErrorNotADefinedName()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eSeriesPartError.ErrNotADefinedName, _
            Description:= _
                    "This element is not a defined name." & vbCrLf & _
                    "Check with property 'EntryType' first."
End Sub

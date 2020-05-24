VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeriesPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("ChartSeries")
'@ModuleDescription("A class to handle each part of the SERIES string.")
'@PredeclaredId

Option Explicit

Implements ISeriesPart

'==============================================================================
Private Const SingleQuote As String = "'"
Private Const DoubleQuote As String = """"

Private Const pcsErrNotARange As String = _
        "This element is neither a range nor a defined name." & vbCrLf & _
        "Check with property 'EntryType'."
Private Const pcsErrNotADefinedName As String = _
        "This element is not a defined name." & vbCrLf & _
        "Check with property 'EntryType'."
'==============================================================================

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


Public Function Create( _
    ByVal ChartWorkbookName As String, _
    ByVal FormulaPart As String _
        ) As ISeriesPart
    
    With New SeriesPart
        .ChartWorkbookName = ChartWorkbookName
        .FormulaPart = FormulaPart
        .Run
        Set Create = .Self
    End With
    
End Function


Public Property Get Self() As ISeriesPart
    Set Self = Me
End Property


'@Description("...")
Public Property Let ChartWorkbookName(ByVal Value As String)
    This.ChartWorkbookName = Value
End Property


Public Property Let FormulaPart(ByVal Value As String)
    This.FormulaPart = Value
'    Run
End Property


Private Property Get ISeriesPart_FormulaPart() As String
    ISeriesPart_FormulaPart = This.FormulaPart
End Property


Private Property Get ISeriesPart_CleanFormulaPart() As String
    ISeriesPart_CleanFormulaPart = This.CleanFormulaPart
End Property


'@Description("Returns the Enum of the 'PartType'")
Private Property Get ISeriesPart_EntryType() As eEntryType
    ISeriesPart_EntryType = This.EntryType
End Property


Private Property Get ISeriesPart_Range() As Range
    Set ISeriesPart_Range = This.Range
End Property


'@Description("Return the range (string) where data are in")
Private Property Get ISeriesPart_RangeString() As String
    With This
        If .IsRange Then
            ISeriesPart_RangeString = .RangeString
        Else
            RaiseErrorNotARange
        End If
    End With
End Property


'@Description("Return the name (string) of the 'Worksheet', where the (range) data are on")
Private Property Get ISeriesPart_RangeSheet() As String
    With This
        If .IsRange Then
            ISeriesPart_RangeSheet = .RangeSheet
        Else
            RaiseErrorNotARange
        End If
    End With
End Property


'@Description("Return the name (string) of the 'Workbook', where the (range) data are on")
Private Property Get ISeriesPart_RangeBook() As String
    With This
        If .IsRange Then
            ISeriesPart_RangeBook = .RangeBook
        Else
            RaiseErrorNotARange
        End If
    End With
End Property


'@Description("Return the path (string) of the 'Workbook', where the (range) data are on")
Private Property Get ISeriesPart_RangePath() As String
    With This
        If .IsRange Then
            ISeriesPart_RangePath = .RangePath
        Else
            RaiseErrorNotARange
        End If
    End With
End Property


Private Function ISeriesPart_IsRange() As Boolean
    ISeriesPart_IsRange = This.IsRange
End Function


Private Function ISeriesPart_IsDefinedNameRange() As Boolean
    With This
        If .EntryType = eDefinedName Then
            ISeriesPart_IsDefinedNameRange = .IsDefinedNameRange
        Else
            RaiseErrorNotADefinedName
        End If
    End With
End Function


'==============================================================================
Private Sub Class_Initialize()
    This.EntryType = eEntryType.eNotSet
End Sub


'==============================================================================
Public Sub Run()
    
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
        Dim sFormulaPart As String
        sFormulaPart = .FormulaPart
        
        If Len(sFormulaPart) = 0 Then               'Empty
            .EntryType = eEmpty
        ElseIf Left$(sFormulaPart, 1) = DoubleQuote Then 'String
            .EntryType = eString
        ElseIf Left$(sFormulaPart, 1) = "{" Then    'Literal Array
            .EntryType = eArray
        ElseIf IsNumeric(sFormulaPart) Then         'Integer (for PlotOrder)
            .EntryType = eInteger
        ElseIf Left$(sFormulaPart, 1) = "(" Then    'Multi Area Range
            .EntryType = eRange
        ElseIf IsDefinedName(sFormulaPart) Then     'Defined Name
            .EntryType = eDefinedName
        Else                                        'Single Area Range
            .EntryType = eRange
        End If
    End With
    
End Sub


Private Function IsDefinedName( _
    ByVal sFormulaPart As String _
        ) As Boolean
    
    Dim ExclMarkPos As Long
    ExclMarkPos = InStrRev(sFormulaPart, "!")
    
    Dim RelevantString As String
    RelevantString = Right$(sFormulaPart, Len(sFormulaPart) - ExclMarkPos)
    
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
                If .IsSingleAreaRange Then
                    .HasSurroundingChars = False
                Else
                    .HasSurroundingChars = True
                End If
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
    With This
        .CleanFormulaPart = Mid$(.FormulaPart, 2, Len(.FormulaPart) - 2)
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
    If Left$(This.CleanFormulaPart, 1) = SingleQuote Then
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
    
    Dim iSeparator As Long
    iSeparator = InStr(TempString, "'!")
    
    Dim sWkbWksPart As String
    sWkbWksPart = Left$(TempString, iSeparator)
    
    This.WkbWksPart = Replace(sWkbWksPart, ReplacementString, CriticalString)
    
End Sub


Private Sub HandleWkbWksPartWithoutSingleQuotes()
    With This
        Dim iSeparator As Long
        iSeparator = InStr(.CleanFormulaPart, "!")
        
        .WkbWksPart = Left$(.CleanFormulaPart, iSeparator - 1)
    End With
End Sub


Private Sub GetCleanWkbWksPart()
    With This
        If Left$(.WkbWksPart, 1) = SingleQuote Then
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
    
    Dim iLen As Long
    iLen = Len(SingleAreaString)
    
    Dim iPos As Long
    iPos = InStrRev(SingleAreaString, "!")
    
    This.RawRangeString = Right$(SingleAreaString, iLen - iPos)
    
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
    
    Dim iLen As Long
    iLen = Len(sWkbWks)
    
    Dim iStart As Long
    iStart = InStr(sWkbWks, "]")
    
    Dim TempString As String
    TempString = Right$(sWkbWks, iLen - iStart)
    
    'SingleQuotes are shown as double SingleQuotes in the SeriesFormula
    '(when they are part of a Worksheet.Name or string)
    This.RangeSheet = Replace(TempString, SingleQuote & SingleQuote, SingleQuote)
    
End Sub


Private Sub GetRangeWorkbook()
    
    With This
        Dim iEnd As Long
        iEnd = InStr(.CleanWkbWksPart, "]")
        
        If iEnd = 0 Then
            .RangeBook = .ChartWorkbookName
        Else
            Dim iStart As Long
            iStart = InStr(.CleanWkbWksPart, "[") + 1
            
            .RangeBook = Mid$(.CleanWkbWksPart, iStart, iEnd - iStart)
        End If
    End With
    
End Sub


Private Sub GetRangePath()
    
    With This
        Dim iEnd As Long
        iEnd = InStrRev(.CleanWkbWksPart, "\")
        
        If iEnd > 0 Then
            .RangePath = Left$(.CleanWkbWksPart, iEnd - 1)
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


'TODO: not sure if something is left here to do
'external ranges can crash Excel (2010) (at least when the file is not open)
'thus only set the range when the source file is open
Private Sub HandleClosedExternalRange()
    
    With This
'        .Type = "Closed External Range"
    End With
    
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
    
    Dim sWks As String
    sWks = This.CleanWkbWksPart
    
    Dim wks As Worksheet
    Set wks = SetWorksheet(sWks)
    
    Dim RangeString As String
    RangeString = This.FormulaPart
    
    Dim sRng As String
    sRng = ExtractRange( _
            RangeString, _
            sWks _
    )
    
    If Len(sRng) <= 255 Then
        Set This.Range = wks.Range(sRng)
    Else
        Dim arrRanges() As String
        arrRanges = Split(sRng, ",")
        
        Dim i As Long
        For i = LBound(arrRanges) To UBound(arrRanges)
            Dim rngUnion As Range
            Set rngUnion = Union2(rngUnion, wks.Range(arrRanges(i)))
        Next
        
        Set This.Range = rngUnion
    End If
    
End Sub


Private Function SetWorksheet( _
    ByVal sWks As String _
        ) As Worksheet
    
    If Left$(sWks, 1) = SingleQuote Then
        Dim iLenWksName As Long
        iLenWksName = Len(sWks)
        
        Dim sWkbWks As String
        sWkbWks = Mid$(sWks, 2, iLenWksName - 2)
    Else
        sWkbWks = sWks
    End If
    
    Dim wkb As Workbook
    Set wkb = SetWorkbook(sWkbWks)
    
    Dim iClosingBracketPos As Long
    iClosingBracketPos = InStr(2, sWkbWks, "]")
    
    Dim sWorksheet As String
    sWorksheet = Right$(sWkbWks, Len(sWkbWks) - iClosingBracketPos)
    
    Dim wks As Worksheet
    Set wks = wkb.Worksheets(sWorksheet)
    
    Set SetWorksheet = wks
    
End Function


Private Function SetWorkbook( _
    ByVal sWkbWks As String _
        ) As Workbook
    
    If Left$(sWkbWks, 1) <> "[" Then
        Set SetWorkbook = Workbooks(This.ChartWorkbookName)
    Else
        Dim iClosingBracketPos As Long
        iClosingBracketPos = InStr(2, sWkbWks, "]")
        
        Dim sWorkbook As String
        sWorkbook = Mid$(sWkbWks, 2, iClosingBracketPos - 2)
        
        Set SetWorkbook = Workbooks(sWorkbook)
    End If
    
End Function


Private Function ExtractRange( _
    ByVal RangeString As String, _
    ByVal sWkbWks As String _
        ) As String
    
    Dim PossibleSingleQuote As String
    If Left$(RangeString, 1) = SingleQuote Then
        PossibleSingleQuote = SingleQuote
    End If
    
    Dim sRangePrefix As String
    sRangePrefix = PossibleSingleQuote & sWkbWks & PossibleSingleQuote & "!"
    
    Dim sRngWithoutWkbWksName As String
    sRngWithoutWkbWksName = Replace(RangeString, sRangePrefix, vbNullString)
    
    Dim sRng As String
    sRng = Replace(sRngWithoutWkbWksName, "$", vbNullString)
    
    ExtractRange = sRng
    
End Function


'==============================================================================
Private Sub HandleDefinedNameCase()
    
    'multiple areas with defined names seem not to be allowed
    GetSingleAreaRawRangeString
    GetRangeString
    
    GetIsDefinedNameRange
    SetDefinedNameRange
    
    If This.IsDefinedNameRange Then
        GetWkbWksPart
        GetCleanWkbWksPart
        
        GetRangeWorksheet
        GetRangeWorkbook
        
        'check if external local range
        If Not IsLocalScopeDefinedName Then
            With This
                .RangeBook = .RangeSheet
                .RangeSheet = vbNullString
            End With
        End If
    End If
    
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


Private Sub GetIsDefinedNameRange()
    
    With This
        On Error Resume Next
        Dim rng As Range
        Set rng = Application.Range(.FormulaPart)
        On Error GoTo 0
        
        .IsDefinedNameRange = (Not rng Is Nothing)
    End With
    
End Sub


Private Sub SetDefinedNameRange()
    With This
        If .IsDefinedNameRange Then
            Set .Range = Application.Range(.FormulaPart)
        Else
            Set .Range = Nothing
        End If
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
Private Sub RaiseErrorNotARange()
    Err.Raise _
            Number:=vbObjectError + 102, _
            Description:=pcsErrNotARange
End Sub


Private Sub RaiseErrorNotADefinedName()
    Err.Raise _
            Number:=vbObjectError + 103, _
            Description:=pcsErrNotADefinedName
End Sub

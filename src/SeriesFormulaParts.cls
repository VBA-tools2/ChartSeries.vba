VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeriesFormulaParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("ChartSeries")
'@ModuleDescription("An object separating the string parts of the 'Series.Formula'.")
'@PredeclaredId

Option Explicit

Implements ISeriesFormulaParts

'==============================================================================
Private Const pcsSep As String = ","
Private Const SingleQuote As String = "'"
Private Const DoubleQuote As String = """"
'==============================================================================

Private Type TSeriesFormulaParts
    FullSeriesFormula As String
    IsSeriesInBubbleChart As Boolean
    WasRunCalled As Boolean
    
    SeriesFormula As String
    
    RemainingFormulaPart(eElement.[_First] To eElement.[_Last] + 1) As String
    PartSeriesFormula(eElement.[_First] To eElement.[_Last]) As String
End Type
Private This As TSeriesFormulaParts


'@Description("Factory method to create this object")
Public Function Create( _
    ByVal FullSeriesFormula As String, _
    Optional ByVal IsSeriesInBubbleChart As Boolean = False _
        ) As ISeriesFormulaParts
    
    With New SeriesFormulaParts
        .FullSeriesFormula = FullSeriesFormula
        .IsSeriesInBubbleChart = IsSeriesInBubbleChart
        .GetSeriesFormulaParts
        Set Create = .Self
    End With
    
End Function


'@Description("Returns the object (itself)")
Public Property Get Self() As ISeriesFormulaParts
    Set Self = Me
End Property


'@Description("Sets the full series formula")
Public Property Let FullSeriesFormula(ByVal Value As String)
    This.FullSeriesFormula = Value
End Property


'@Description("Sets if the 'Series(.Formula)' is in a bubble chart")
Public Property Let IsSeriesInBubbleChart(ByVal Value As Boolean)
    This.IsSeriesInBubbleChart = Value
End Property


Private Property Get ISeriesFormulaParts_PartSeriesFormula(ByVal Element As eElement) As String
    ISeriesFormulaParts_PartSeriesFormula = This.PartSeriesFormula(Element)
End Property


'==============================================================================
Public Sub GetSeriesFormulaParts()
    
    If This.WasRunCalled Then Exit Sub
    
    GetSeriesFormula
    
    GetBubbleSizes
    GetPlotOrder
    GetValues eElement.eYValues
    GetValues eElement.eXValues
    GetSeriesName
    
    This.WasRunCalled = True
    
End Sub


Private Sub GetSeriesFormula()
    
    '==========================================================================
    Const csPrefix As String = "=SERIES("
    Const ciSuffix As Long = 1
    '==========================================================================
    
    
    Dim iLenPrefix As Long
    iLenPrefix = Len(csPrefix)
    
    With This
        Dim iLenSeriesFormula As Long
        iLenSeriesFormula = Len(.FullSeriesFormula)
        
        .SeriesFormula = _
                Mid$( _
                        .FullSeriesFormula, _
                        iLenPrefix + 1, _
                        iLenSeriesFormula - iLenPrefix - ciSuffix _
                )
    End With
    
End Sub


Private Sub GetBubbleSizes()
    Dim Element As eElement
    Element = eElement.eBubbleSizes
    
    With This
        If .IsSeriesInBubbleChart Then
            .RemainingFormulaPart(Element + 1) = .SeriesFormula
            GetValues Element
        Else
            .RemainingFormulaPart(Element) = .SeriesFormula
            .PartSeriesFormula(Element) = vbNullString
        End If
    End With
End Sub


Private Sub GetPlotOrder()
    Dim Element As eElement
    Element = eElement.ePlotOrder
    
    With This
        Dim LastComma As Long
        LastComma = InStrRev(.RemainingFormulaPart(Element + 1), pcsSep)
        
        .RemainingFormulaPart(Element) = _
                Left$(.RemainingFormulaPart(Element + 1), LastComma - 1)
        .PartSeriesFormula(Element) = _
                Mid$(.RemainingFormulaPart(Element + 1), LastComma + 1)
    End With
    
End Sub


Private Sub GetValues( _
    ByVal Element As eElement _
)
    
    Dim iMinCommata As Long
    iMinCommata = Element - 1
    
    With This
        Dim sCurrString As String
        sCurrString = .RemainingFormulaPart(Element + 1)
        
        Dim LenTitleYXValues As Long
        LenTitleYXValues = Len(sCurrString)
        
        Select Case Right$(sCurrString, 1)
            Case "}"                               'Literal Array
                Dim sOneLess As String
                Dim YXValues As String
                GetLiteralArrayYXValues _
                        sCurrString, _
                        LenTitleYXValues, _
                        sOneLess, _
                        YXValues
            Case ")"                               'Multi Area Range
                GetMultiAreaRangeYXValues _
                        LenTitleYXValues, _
                        sCurrString, _
                        sOneLess, _
                        YXValues
            Case pcsSep                            'Empty
                sOneLess = Left$(sCurrString, LenTitleYXValues - 1)
                YXValues = vbNullString
            Case Else                              'Single Area Range
                GetSingleRangeYXValues _
                        LenTitleYXValues, _
                        iMinCommata, _
                        sCurrString, _
                        sOneLess, _
                        YXValues
        End Select
        .PartSeriesFormula(Element) = YXValues
        .RemainingFormulaPart(Element) = sOneLess
    End With
    
End Sub


Private Sub GetLiteralArrayYXValues( _
    ByVal sCurrString As String, _
    ByVal LenTitleYXValues As String, _
    ByRef sOneLess As String, _
    ByRef YXValues As String _
)
    
    Dim iOpeningBracket As Long
    iOpeningBracket = InStrRev(sCurrString, "{")
    
    sOneLess = Left$(sCurrString, iOpeningBracket - 2)
    
    YXValues = Mid$( _
            sCurrString, _
            iOpeningBracket, _
            LenTitleYXValues - iOpeningBracket + 1 _
    )
    
End Sub


'Invalid characters in 'Worksheet.Name's: */:?[\]
'It is also invalid that 'Worksheet.Name's start or end with a (single) quote
'These characters cause (single) quotes:
'    ' ,!,",#,$,%,&,',(,),*,+,,,-,;,<,=,>,?,@,^,`,{,|,},~,,?,‚,„,‹,?,?,?,‘,’,”,•,›,?, ,¢,£,¥,¦,©,«,¬,®,»'
'Corresponding character numbers:
'    '32,33,34,35,36,37,38,39,40,41,42,43,44,45,59,60,61,62,63,64,94,96,123,124,125,126,127,129,130,132,139,141,143,144,145,146,148,149,155,157,160,162,163,165,166,169,171,172,174,187'
Private Sub GetMultiAreaRangeYXValues( _
    ByVal LenTitleYXValues As Long, _
    ByVal sCurrString As String, _
    ByRef sOneLess As String, _
    ByRef YXValues As String _
)
    
    Dim iExclMarkPos As Long
    iExclMarkPos = InStrRev(sCurrString, "!")
    
    If IsCharBeforeExclMarkSingleQuote( _
            sCurrString, _
            iExclMarkPos _
    ) Then
        HandleWithSingleQuoteCase _
                sCurrString, _
                sOneLess, _
                YXValues
    Else
        HandleWithoutSingleQuoteCase _
                LenTitleYXValues, _
                sCurrString, _
                sOneLess, _
                YXValues
    End If
End Sub


Private Function IsCharBeforeExclMarkSingleQuote( _
    ByVal sCurrString As String, _
    ByVal iExclMarkPos As Long _
        ) As Boolean
    
    IsCharBeforeExclMarkSingleQuote = _
            (Mid$(sCurrString, iExclMarkPos - 1, 1) = "'")
End Function


'Luckily 'Worksheet.Name's are not allowed to begin with a (single) quote
'and when they are used in 'Worksheet.Name's they are doubled.
'So in case ",('" is used in a 'Worksheet.Name' they will be shown as
'",(''".
Private Sub HandleWithSingleQuoteCase( _
    ByVal sCurrString As String, _
    ByRef sOneLess As String, _
    ByRef YXValues As String _
)
    
    Const CriticalString As String = ",(''"
    'replace with a char that is not allowed in 'Worksheet.Name's
    Const ReplacementString As String = "?"
    Dim TempString As String
    TempString = Replace(sCurrString, CriticalString, ReplacementString)
    
    Dim iOpeningBracket As Long
    iOpeningBracket = InStrRev(TempString, ",('")
    
    Dim sTempOneLess As String
    sTempOneLess = Left$(TempString, iOpeningBracket - 1)
    sOneLess = Replace(sTempOneLess, ReplacementString, CriticalString)
    
    Dim sTempYXValues As String
    sTempYXValues = Mid$( _
            TempString, _
            iOpeningBracket + 1, _
            Len(TempString) - iOpeningBracket _
    )
    YXValues = Replace(sTempYXValues, ReplacementString, CriticalString)
    
End Sub


Private Sub HandleWithoutSingleQuoteCase( _
    ByVal LenTitleYXValues As Long, _
    ByVal sCurrString As String, _
    ByRef sOneLess As String, _
    ByRef YXValues As String _
)
    
    Dim iOpeningBracket As Long
    iOpeningBracket = InStrRev(sCurrString, ",(")
    
    sOneLess = Left$(sCurrString, iOpeningBracket - 1)
    YXValues = Mid$( _
            sCurrString, _
            iOpeningBracket + 1, _
            LenTitleYXValues - iOpeningBracket _
    )
    
End Sub


'NOTE: refactor
Private Sub GetSingleRangeYXValues( _
    ByVal LenTitleYXValues As Long, _
    ByVal iMinCommata As Long, _
    ByVal sCurrString As String, _
    ByRef sOneLess As String, _
    ByRef YXValues As String _
)
    
    Dim PosOfLastExclMark As Long
    PosOfLastExclMark = InStrRev(sCurrString, "!")
    
    If Mid$(sCurrString, PosOfLastExclMark - 1, 1) = SingleQuote Then
        Dim iExclMark As Long
        iExclMark = InStrRev(sCurrString, "'!")
        
        Dim iTotalNoOfSeparators As Long
        iTotalNoOfSeparators = NoOfOccurrencesInString(sCurrString, pcsSep)
        
        Dim iNoOfSeparatorsInPossibleStringTitle As Long
        iNoOfSeparatorsInPossibleStringTitle = _
                GetNoOfSeparatorsInPossibleStringTitle(sCurrString)
        
        Dim iNoOfSeparatorsWOTheOnesInPossibleStringTitle As Long
        iNoOfSeparatorsWOTheOnesInPossibleStringTitle = _
                iTotalNoOfSeparators - iNoOfSeparatorsInPossibleStringTitle
        
        Dim iComma As Long
        If iNoOfSeparatorsWOTheOnesInPossibleStringTitle = iMinCommata Then
            iComma = InStrRev(sCurrString, pcsSep)
        Else
            iComma = InStrRev(sCurrString, pcsSep & SingleQuote, iExclMark)
        End If
        
        sOneLess = Left$(sCurrString, iComma - 1)
        YXValues = Right$( _
                sCurrString, _
                LenTitleYXValues - iComma _
        )
    Else
        Dim iNextComma As Long
        iNextComma = InStrRev(sCurrString, pcsSep)
        
        sOneLess = Left$(sCurrString, iNextComma - 1)
        YXValues = Right$(sCurrString, LenTitleYXValues - iNextComma)
    End If
    
End Sub


Private Function GetNoOfSeparatorsInPossibleStringTitle( _
    ByVal sCurrString As String _
        ) As Long
    
    'no string title
    If Left$(sCurrString, 1) <> DoubleQuote Then Exit Function
    
    Dim iNoOfQuotes As Long
    iNoOfQuotes = NoOfOccurrencesInString(sCurrString, DoubleQuote)
    
    If iNoOfQuotes = 2 Then
        Dim sTitleString As String
        sTitleString = GetStringTitle(sCurrString)
        
        GetNoOfSeparatorsInPossibleStringTitle = _
                NoOfOccurrencesInString(sTitleString, pcsSep)
    Else
        'then it gets complicated, because either the string title itself
        'contains one or more quotes or the 'Worksheet.Name' contains one or
        'more quotes
        '--> return a negative value and hope that it will work
        GetNoOfSeparatorsInPossibleStringTitle = -1
        Stop
    End If
    
End Function


Private Function GetStringTitle( _
    ByVal sCurrString As String _
        ) As String
    
    Dim iLastQuotePos As Long
    iLastQuotePos = InStrRev(sCurrString, DoubleQuote)
    
    Dim sTitleString As String
    sTitleString = Mid$(sCurrString, 2, iLastQuotePos - 2)
    
    GetStringTitle = sTitleString
    
End Function


Private Sub GetSeriesName()
    Dim Element As eElement
    Element = eElement.eName
    
    This.PartSeriesFormula(Element) = This.RemainingFormulaPart(Element + 1)
End Sub


'=============================================================================
'inspired by <https://stackoverflow.com/a/23357807/5776000>
Private Function NoOfOccurrencesInString( _
    ByVal Expression As String, _
    ByVal ToSearch As String _
        ) As Long
    
    Dim iNoOfOccurrencesInString As Long
    iNoOfOccurrencesInString = _
            Len(Expression) - Len(Replace(Expression, ToSearch, vbNullString))
    
    NoOfOccurrencesInString = iNoOfOccurrencesInString
    
End Function

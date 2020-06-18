VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeriesFormulaParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object separating the string parts of the 'Series.Formula'."

'@Folder("ChartSeries")
'@ModuleDescription("An object separating the string parts of the 'Series.Formula'.")
'@PredeclaredId

Option Explicit

Implements ISeriesFormulaParts


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
Attribute Create.VB_Description = "Factory method to create this object"
    
    With New SeriesFormulaParts
        .FullSeriesFormula = FullSeriesFormula
        .IsSeriesInBubbleChart = IsSeriesInBubbleChart
        .GetSeriesFormulaParts
        Set Create = .Self
    End With
    
End Function


'@Description("Returns the object (itself)")
Friend Property Get Self() As ISeriesFormulaParts
Attribute Self.VB_Description = "Returns the object (itself)"
    Set Self = Me
End Property


'@Description("Sets the full series formula")
Friend Property Let FullSeriesFormula(ByVal Value As String)
Attribute FullSeriesFormula.VB_Description = "Sets the full series formula"
    If Len(Value) = 0 Then RaiseErrorFullSeriesFormulaEmpty
    
    This.FullSeriesFormula = Value
End Property


'@Description("Sets if the 'Series(.Formula)' is in a bubble chart")
Friend Property Let IsSeriesInBubbleChart(ByVal Value As Boolean)
Attribute IsSeriesInBubbleChart.VB_Description = "Sets if the 'Series(.Formula)' is in a bubble chart"
    This.IsSeriesInBubbleChart = Value
End Property


Private Property Get ISeriesFormulaParts_PartSeriesFormula(ByVal Element As eElement) As String
    ISeriesFormulaParts_PartSeriesFormula = This.PartSeriesFormula(Element)
End Property


'==============================================================================
Friend Sub GetSeriesFormulaParts()
    
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
    
    
    Dim PrefixLength As Long
    PrefixLength = Len(csPrefix)
    
    With This
        Dim SeriesFormulaLength As Long
        SeriesFormulaLength = Len(.FullSeriesFormula)
        
        .SeriesFormula = _
                Mid$( _
                        .FullSeriesFormula, _
                        PrefixLength + 1, _
                        SeriesFormulaLength - PrefixLength - ciSuffix _
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
        LastComma = InStrRev(.RemainingFormulaPart(Element + 1), ",")
        
        .RemainingFormulaPart(Element) = _
                Left$(.RemainingFormulaPart(Element + 1), LastComma - 1)
        .PartSeriesFormula(Element) = _
                Mid$(.RemainingFormulaPart(Element + 1), LastComma + 1)
    End With
    
End Sub


Private Sub GetValues( _
    ByVal Element As eElement _
)
    
    With This
        Dim CurrentString As String
        CurrentString = .RemainingFormulaPart(Element + 1)
        
        Dim CurrentStringLength As Long
        CurrentStringLength = Len(CurrentString)
        
        Select Case Right$(CurrentString, 1)
            Case "}"                               'Literal Array
                Dim RemainingString As String
                Dim CurrentPartString As String
                GetLiteralArray _
                        CurrentString, _
                        CurrentStringLength, _
                        RemainingString, _
                        CurrentPartString
            Case ")"                               'Multi Area Range
                GetMultiAreaRange _
                        CurrentString, _
                        CurrentStringLength, _
                        RemainingString, _
                        CurrentPartString
            Case ","                               'Empty
                RemainingString = Left$(CurrentString, CurrentStringLength - 1)
                CurrentPartString = vbNullString
            Case Else                              'Single Area Range
                Dim MinCommataCount As Long
                MinCommataCount = Element - 1
                
                GetSingleAreaRange _
                        MinCommataCount, _
                        CurrentString, _
                        CurrentStringLength, _
                        RemainingString, _
                        CurrentPartString
        End Select
        .PartSeriesFormula(Element) = CurrentPartString
        .RemainingFormulaPart(Element) = RemainingString
    End With
    
End Sub


Private Sub GetLiteralArray( _
    ByVal CurrentString As String, _
    ByVal CurrentStringLength As String, _
    ByRef RemainingString As String, _
    ByRef CurrentPartString As String _
)
    
    Dim OpeningBracketPosition As Long
    OpeningBracketPosition = InStrRev(CurrentString, "{")
    
    RemainingString = Left$(CurrentString, OpeningBracketPosition - 2)
    
    CurrentPartString = Mid$( _
            CurrentString, _
            OpeningBracketPosition, _
            CurrentStringLength - OpeningBracketPosition + 1 _
    )
    
End Sub


'Invalid characters in 'Worksheet.Name's: */:?[\]
'It is also invalid that 'Worksheet.Name's start or end with a (single) quote
'These characters cause (single) quotes:
'    ' ,!,",#,$,%,&,',(,),*,+,,,-,;,<,=,>,?,@,^,`,{,|,},~,,?,‚,„,‹,?,?,?,‘,’,”,•,›,?, ,¢,£,¥,¦,©,«,¬,®,»'
'Corresponding character numbers:
'    '32,33,34,35,36,37,38,39,40,41,42,43,44,45,59,60,61,62,63,64,94,96,123,124,125,126,127,129,130,132,139,141,143,144,145,146,148,149,155,157,160,162,163,165,166,169,171,172,174,187'
Private Sub GetMultiAreaRange( _
    ByVal CurrentString As String, _
    ByVal CurrentStringLength As Long, _
    ByRef RemainingString As String, _
    ByRef CurrentPartString As String _
)
    
    Dim ExclMarkPosition As Long
    ExclMarkPosition = InStrRev(CurrentString, "!")
    
    If IsCharBeforeExclMarkSingleQuote( _
            CurrentString, _
            ExclMarkPosition _
    ) Then
        HandleWithSingleQuoteCase _
                CurrentString, _
                RemainingString, _
                CurrentPartString
    Else
        HandleWithoutSingleQuoteCase _
                CurrentString, _
                CurrentStringLength, _
                RemainingString, _
                CurrentPartString
    End If
End Sub


Private Function IsCharBeforeExclMarkSingleQuote( _
    ByVal CurrentString As String, _
    ByVal ExclMarkPosition As Long _
        ) As Boolean
    
    IsCharBeforeExclMarkSingleQuote = _
            (Mid$(CurrentString, ExclMarkPosition - 1, 1) = "'")
End Function


'Luckily 'Worksheet.Name's are not allowed to begin with a (single) quote
'and when they are used in 'Worksheet.Name's they are doubled.
'So in case ",('" is used in a 'Worksheet.Name' they will be shown as
'",(''".
Private Sub HandleWithSingleQuoteCase( _
    ByVal CurrentString As String, _
    ByRef RemainingString As String, _
    ByRef CurrentPartString As String _
)
    
    '==========================================================================
    Const CriticalString As String = ",(''"
    'replace with a char that is not allowed in 'Worksheet.Name's
    Const ReplacementString As String = "?"
    '==========================================================================
    
    Dim TempString As String
    TempString = Replace(CurrentString, CriticalString, ReplacementString)
    
    Dim OpeningBracketPosition As Long
    OpeningBracketPosition = InStrRev(TempString, ",('")
    
    Dim TempRemainingString As String
    TempRemainingString = Left$(TempString, OpeningBracketPosition - 1)
    RemainingString = Replace(TempRemainingString, ReplacementString, CriticalString)
    
    Dim TempCurrentPartString As String
    TempCurrentPartString = Mid$( _
            TempString, _
            OpeningBracketPosition + 1, _
            Len(TempString) - OpeningBracketPosition _
    )
    CurrentPartString = Replace( _
            TempCurrentPartString, _
            ReplacementString, _
            CriticalString _
    )
    
End Sub


Private Sub HandleWithoutSingleQuoteCase( _
    ByVal CurrentString As String, _
    ByVal CurrentStringLength As Long, _
    ByRef RemainingString As String, _
    ByRef CurrentPartString As String _
)
    
    Dim OpeningBracketPosition As Long
    OpeningBracketPosition = InStrRev(CurrentString, ",(")
    
    RemainingString = Left$(CurrentString, OpeningBracketPosition - 1)
    CurrentPartString = Mid$( _
            CurrentString, _
            OpeningBracketPosition + 1, _
            CurrentStringLength - OpeningBracketPosition _
    )
    
End Sub


Private Sub GetSingleAreaRange( _
    ByVal MinCommataCount As Long, _
    ByVal CurrentString As String, _
    ByVal CurrentStringLength As Long, _
    ByRef RemainingString As String, _
    ByRef CurrentPartString As String _
)
    
    Dim LastExclMarkPosition As Long
    LastExclMarkPosition = InStrRev(CurrentString, "!")
    
    If Mid$(CurrentString, LastExclMarkPosition - 1, 1) = "'" Then
        Dim ExclMarkPosition As Long
        ExclMarkPosition = InStrRev(CurrentString, "'!")
        
        Dim TotalNoOfSeparators As Long
        TotalNoOfSeparators = NoOfOccurrencesInString(CurrentString, ",")
        
        Dim CommaPosition As Long
        If TotalNoOfSeparators = MinCommataCount Then
            CommaPosition = InStrRev(CurrentString, ",")
        Else
            CommaPosition = InStrRev(CurrentString, ",'", ExclMarkPosition)
        End If
        
        RemainingString = Left$(CurrentString, CommaPosition - 1)
        CurrentPartString = Right$( _
                CurrentString, _
                CurrentStringLength - CommaPosition _
        )
    Else
        Dim NextCommaPosition As Long
        NextCommaPosition = InStrRev(CurrentString, ",")
        
        RemainingString = Left$(CurrentString, NextCommaPosition - 1)
        CurrentPartString = Right$(CurrentString, CurrentStringLength - NextCommaPosition)
    End If
    
End Sub


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
    
    NoOfOccurrencesInString = _
            Len(Expression) - Len(Replace(Expression, ToSearch, vbNullString))
    
End Function


'=============================================================================
Private Sub RaiseErrorFullSeriesFormulaEmpty()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eSeriesFormulaPartsError.ErrFullSeriesFormulaEmpty, _
            Description:= _
                    "This error should never be raised." & vbCrLf & _
                    "The 'FullSeriesFormula' is empty, " & _
                    "which should be caught by the calling routine."
End Sub

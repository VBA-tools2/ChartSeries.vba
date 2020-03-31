VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'@Exposed

'This Class module provides an easy way to access the items in a
'chart's SERIES formula. It can be exported and then imported into
'any project

'Developed by John Walkenbach, JWALK AND ASSOCIATES
'Copyright 1999. All rights reserved.
'May be used and distributed freely, but may not be sold.
'http://www.j-walk.com/ss/
'new website:
'<http://spreadsheetpage.com/index.php/file/a_class_module_to_manipulate_a_chart_series/>

'===================================
' PROPERTIES FOR CHARTSERIES OBJECT
'===================================

'Chart (read/write)
'ChartSeries (read/write)

'SeriesNameType (read-only)
'XValuesType (read-only)
'ValuesType (read-only)
'PlotOrderType (read-only)
'BubbleSizesType (read-only)

'SeriesName (read/write)
'XValues (read/write)
'Values (read/write)
'PlotOrder (read/write)
'BubbleSizes (read/write)

'DataSheet (read-only)
'NoOfPoints (read-only)
'PointSourceRange (read-only)

'Sub ExampleUsage()
'    Dim MySeries As clsChartSeries
'    Set MySeries = New clsChartSeries
'    With MySeries
'        .Chart = Sheets(1).ChartObjects(1).Chart
'        .ChartSeries = 1
'        If .XValuesType = "Range" Then
'            MsgBox .XValues.Address & " with " & .NoOfPoints & " points"
'        Else
'            MsgBox .XValues
'        End If
'    End With
'End Sub

Option Explicit

Private Type tElement
    Type  As String             'type of element ("Empty", "String", "Integer",
                                '"Array", "Named Range", "External Range",
                                '"Open External Range", "Closed External Range",
                                '"Range", "Long Range")
    Value As Variant            'value of element
    Range As Range              'if type is "Range" then it stores the range
    LongRangeString As String   'if type is "Long Range" return the range as string
End Type

Private Type TChartSeries
    CurrChart As Chart
    CurrSeries As Long
    OldFullSeriesFormula As String
    FullSeriesFormula As String
    SeriesFormula As String
    NameXValuesValuesPlotOrder As String
    NameXValuesValues As String
    NameXValues As String
    Name As String
    CurrData() As tElement
    CurrSheet() As String
    CurrBook() As String
End Type
Private This As TChartSeries

Private Enum eElement
    [_First] = 1
    eName = eElement.[_First]
    eXValues
    eYValues
    ePlotOrder
    eBubbleSizes
    [_Last] = eElement.eBubbleSizes
End Enum

'==============================================================================
'string to return when a series formula is inaccessible
Private Const pcsInaccessible As String = "inaccessible"
'string to return when a series number is invalid
Private Const pcsBadSeriesNumber As String = "ERROR - BAD SERIES NUMBER"
'string to return if 'BubbleSizes' or 'BubbleSizesType' is asked for a
'non-bubble chart
Private Const pcsBadBubbleChart As String = "Error - No Bubble Chart"

Private Const pcsSep As String = ","
Private Const pcsApostrophe As String = "'"
'==============================================================================


Public Property Get FullFormula() As String
    Call SERIESFormulaElement
    FullFormula = This.FullSeriesFormula
End Property

Public Property Let FullFormula(ByVal sFormula As String)
    With This
        .CurrChart.SeriesCollection(.CurrSeries).Formula = sFormula
    End With
End Property


Public Property Get Chart() As Chart
    Set Chart = This.CurrChart
End Property

Public Property Let Chart(ByVal cht As Chart)
    Set This.CurrChart = cht
End Property


Public Property Get ChartSeries() As Long
    ChartSeries = This.CurrSeries
End Property

Public Property Let ChartSeries(ByVal SeriesNum As Long)
    This.CurrSeries = SeriesNum
End Property


Public Property Get SeriesNameType() As String
    SeriesNameType = This.CurrData(eElement.eName).Type
End Property


Public Property Get SeriesName() As Variant
    Call SERIESFormulaElement
    With This.CurrData(eElement.eName)
        Select Case .Type
            Case "Range", "Named Range", "Open External Range"
                Set SeriesName = .Range
            Case "Long Range"
                SeriesName = .LongRangeString
            Case Else
                SeriesName = .Value
        End Select
    End With
End Property

'2do:
'With this only a string can be set.
'How can a range be set?
Public Property Let SeriesName(ByVal SName As Variant)
    With This
        .CurrChart.SeriesCollection(.CurrSeries).Name = SName
    End With
End Property


Public Property Get XValuesType() As String
    Call SERIESFormulaElement
    XValuesType = This.CurrData(eElement.eXValues).Type
End Property


Public Property Get XValues() As Variant
    Call SERIESFormulaElement
    With This.CurrData(eElement.eXValues)
        Select Case .Type
            Case "Range", "Named Range", "Open External Range"
                Set XValues = .Range
            Case "Long Range"
                XValues = .LongRangeString
            Case Else
                XValues = .Value
        End Select
    End With
End Property

Public Property Let XValues(ByVal XVals As Variant)
    With This
        If TypeName(XVals) = "Range" Then
            Dim sRange As String
            sRange = GetRangeAddress(XVals)
            .CurrChart.SeriesCollection(.CurrSeries).XValues = sRange
        Else
            .CurrChart.SeriesCollection(.CurrSeries).XValues = XVals
        End If
    End With
End Property


Public Property Get ValuesType() As String
    Call SERIESFormulaElement
    ValuesType = This.CurrData(eElement.eYValues).Type
End Property


Public Property Get Values() As Variant
    Call SERIESFormulaElement
    With This.CurrData(eElement.eYValues)
        Select Case .Type
            Case "Range", "Named Range", "Open External Range"
                Set Values = .Range
            Case "Long Range"
                Values = .LongRangeString
            Case Else
                Values = .Value
        End Select
    End With
End Property

Public Property Let Values(ByVal Vals As Variant)
    With This
        If TypeName(Vals) = "Range" Then
            Dim sRange As String
            sRange = GetRangeAddress(Vals)
            .CurrChart.SeriesCollection(.CurrSeries).Values = sRange
        Else
            .CurrChart.SeriesCollection(.CurrSeries).Values = Vals
        End If
    End With
End Property


Public Property Get PlotOrderType() As String
    Call SERIESFormulaElement
    PlotOrderType = This.CurrData(eElement.ePlotOrder).Type
End Property


Public Property Get PlotOrder() As Long
    Call SERIESFormulaElement
    PlotOrder = This.CurrData(eElement.ePlotOrder).Value
End Property

Public Property Let PlotOrder(ByVal PltOrder As Long)
    With This
        .CurrChart.SeriesCollection(.CurrSeries).PlotOrder = PltOrder
        .CurrSeries = PltOrder
    End With
End Property


Public Property Get BubbleSizesType() As String
    Call SERIESFormulaElement
    BubbleSizesType = This.CurrData(eElement.eBubbleSizes).Type
End Property


Public Property Get BubbleSizes() As Variant
    Call SERIESFormulaElement
    With This.CurrData(eElement.eBubbleSizes)
        If .Type = "Range" Then
            Set BubbleSizes = .Range
        Else
            BubbleSizes = .Value
        End If
    End With
End Property

Public Property Let BubbleSizes(ByVal Vals As Variant)
    With This
        .CurrChart.SeriesCollection(.CurrSeries).Values = Vals
    End With
End Property


'determine Worksheet, where the data of SERIES come from
Public Property Get DataSheet(ByVal Element As Long) As String
    Call SERIESFormulaElement
    DataSheet = RemovePossibleApostrophes(This.CurrSheet(Element))
End Property


'determine Workbook, where the data of SERIES come from
Public Property Get DataBook(ByVal Element As Long) As String
    Call SERIESFormulaElement
    DataBook = RemovePossibleApostrophes(This.CurrBook(Element))
End Property


'determine number of data points in SERIES
Public Property Get NoOfPoints() As Long
    NoOfPoints = SERIESFormulaDataPoints()
End Property


'determine the source range of Point 'PointNumber' of SERIES
'the order of the points is as stated in the SERIES, not as shown in the graph
Public Property Get PointSourceRange( _
    ByVal Element As Long, _
    ByVal PointNumber As Long _
        ) As Variant
    PointSourceRange = PointRange(Element, PointNumber)
End Property


'==============================================================================
Private Sub Class_Initialize()
    With This
        ReDim .CurrData(eElement.[_First] To eElement.[_Last])
        ReDim .CurrSheet(eElement.[_First] To eElement.[_Last])
        ReDim .CurrBook(eElement.[_First] To eElement.[_Last])
    End With
End Sub


'==============================================================================

'This function essentially parses and analyzes the given 'Element' of a
'SERIES formula
'
'Element 1: Series Name
'- Type:  Returns "Range", "External Range", "Empty", or "String"
'- Value: Can be a range, a literal value, or nothing
'Element 2: XValues.
'- Type:  Returns "Range", "External Range", "Empty", or "Array"
'- Value: Can be a range reference (including a non-contiguous range), a literal array, or nothing
'Element 3: Values.
'- Type:  Returns "Range", "External Range", "Empty", or "Array"
'- Value: Can be a range reference (including a non-contiguous range), or a literal array
'Element 4: PlotOrder.
'- Type:  Always returns "Integer"
'- Value: Must be an integer
'Element 5: BubbleSizes (for Bubble plots (only)).
'- Type:  Returns "Range", "External Range", or "Array"
'- Value: Can be a range reference (including a non-contiguous range), or a literal array
Private Sub SERIESFormulaElement()

    If Not IsCurrSeriesValid Then Exit Sub
    If Not GetFullSeriesFormula Then Exit Sub
    
    With This
        If .OldFullSeriesFormula <> .FullSeriesFormula Then
            Call GetSeriesFormula
            
            Call GetBubbleSizes
            Call GetPlotOrder
            Call GetValues(eElement.eYValues)
            Call GetValues(eElement.eXValues)
            Call GetSeriesName
            
            Call GetWorkbooks
            Call GetWorksheets
        End If
    End With
    
End Sub


Private Function IsCurrSeriesValid() As Boolean

    IsCurrSeriesValid = False
    
    With This
        If .CurrSeries < 1 Or .CurrSeries > .CurrChart.SeriesCollection.Count Then
            Dim i As eElement
            For i = eElement.eName To eElement.eYValues
                With .CurrData(i)
                    .Type = pcsBadSeriesNumber
                    .Value = pcsBadSeriesNumber
                End With
            Next
        Else
            IsCurrSeriesValid = True
        End If
    End With

End Function


Private Function GetFullSeriesFormula() As Boolean

    With This
        GetFullSeriesFormula = False
        .OldFullSeriesFormula = .FullSeriesFormula
        .FullSeriesFormula = vbNullString
        
        With .CurrChart
            'in case some data are currently hidden and thus the formula can't be
            'accessed, temporarily change '.PlotVisibleOnly' to 'False'
            Dim bPlotVisible As Boolean
            bPlotVisible = .PlotVisibleOnly
            .PlotVisibleOnly = False
            
            'because there is a possibility that an error is risen, take care of that
            'situation
            '(inspired by <https://peltiertech.com/change-series-formula-improved-routines/>)
            On Error Resume Next
            This.FullSeriesFormula = .SeriesCollection(This.CurrSeries).Formula
            'change to column chart if series is inaccessible
            If Len(This.FullSeriesFormula) = 0 Then
                'reset error number
                Err.Number = 0
                'store current 'ChartType' of series ...
                Dim iChartType As XlChartType
                iChartType = .SeriesCollection(This.CurrSeries).ChartType
                '... and change it temporarily to a column chart
                .SeriesCollection(This.CurrSeries).ChartType = xlColumnClustered
                
                'now try to read the formula again
                This.FullSeriesFormula = .SeriesCollection(This.CurrSeries).Formula
                
                'restore original 'ChartType'
                .SeriesCollection(This.CurrSeries).ChartType = iChartType
                
                'if still an error is returned, the formula is inaccessible
                '(this e.g. can be the case if there are #REF errors inside the given ranges)
                If Err.Number <> 0 Then
                    Dim i As eElement
                    For i = eElement.eName To eElement.eYValues
                        With This.CurrData(i)
                            .Type = pcsInaccessible
                            .Value = pcsInaccessible
                        End With
                    Next
                    This.FullSeriesFormula = pcsInaccessible
                    GoTo Exit_Function
                End If
            End If
            On Error GoTo 0
        End With
    End With
    
    GetFullSeriesFormula = True
    
Exit_Function:
    This.CurrChart.PlotVisibleOnly = bPlotVisible

End Function


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
    
    With This
        Select Case .CurrChart.ChartType
            Case xlBubble, xlBubble3DEffect
                Call GetValues(eElement.eBubbleSizes)
            Case Else
                .NameXValuesValuesPlotOrder = .SeriesFormula
                
                With This.CurrData(eElement.eBubbleSizes)
                    Set .Range = Nothing
                    .Type = pcsBadBubbleChart
                    .Value = pcsBadBubbleChart
                End With
        End Select
    End With
    
End Sub


Private Sub GetPlotOrder()

    With This
        Dim LastComma As Long
        LastComma = InStrRev(.NameXValuesValuesPlotOrder, pcsSep)
        
        .NameXValuesValues = Left$(.NameXValuesValuesPlotOrder, LastComma - 1)
        With .CurrData(eElement.ePlotOrder)
            .Type = "Integer"
            .Value = Mid$(This.NameXValuesValuesPlotOrder, LastComma + 1)
            Set .Range = Nothing
        End With
    End With

End Sub


Private Sub GetValues( _
    ByVal Element As eElement _
)
    
    With This
        If Element = eElement.eBubbleSizes Then
            Dim sCurrString As String
            sCurrString = .SeriesFormula
            
            Dim iMinCommata As Long
            iMinCommata = 4
        ElseIf Element = eElement.eYValues Then
            sCurrString = .NameXValuesValues
            iMinCommata = 2
        ElseIf Element = eElement.eXValues Then
            sCurrString = .NameXValues
            iMinCommata = 1
        End If
    End With
    
    Dim LenTitleYXValues As Long
    LenTitleYXValues = Len(sCurrString)
    
    Select Case Right$(sCurrString, 1)
        Case "}"    'Literal Array
            Dim iOpeningBracket As Long
            iOpeningBracket = InStrRev(sCurrString, "{")
            
            Dim sOneLess As String
            sOneLess = Left$(sCurrString, iOpeningBracket - 2)
            
            Dim YXValues As String
            YXValues = Mid$( _
                    sCurrString, _
                    iOpeningBracket + 1, _
                    LenTitleYXValues - iOpeningBracket - 1 _
            )
        Case ")"    'Multiple ranges
            Call GetMultiRangeYXValues( _
                    LenTitleYXValues, _
                    sCurrString, _
                    sOneLess, _
                    YXValues _
            )
        Case pcsSep    'Empty
            sOneLess = Left$(sCurrString, LenTitleYXValues - 1)
            YXValues = vbNullString
        Case Else   'A single range
            Call GetSingleRangeYXValues( _
                    LenTitleYXValues, _
                    iMinCommata, _
                    sCurrString, _
                    sOneLess, _
                    YXValues _
            )
    End Select
    
    Dim DummyElement As tElement
    With DummyElement
        If Len(YXValues) = 0 Then
            .Type = "Empty"
            Set .Range = Nothing
        ElseIf Right$(sCurrString, 1) = "}" Then
            .Type = "Array"
            Set .Range = Nothing
        Else
            If IsRangeExternal(YXValues) Then
                HandleExternalRange _
                        YXValues, _
                        DummyElement
            ElseIf IsNamedRange(YXValues) Then
                .Type = "Named Range"
                Set .Range = Range(YXValues)
            ElseIf Len(YXValues) <= 255 Then
                .Type = "Range"
                Set .Range = Range(YXValues)
            Else
                HandleLongRangeCase _
                        YXValues, _
                        DummyElement
            End If
        End If
        .Value = YXValues
    End With
    
    With This
        If Element = eElement.eBubbleSizes Then
            .CurrData(eElement.eBubbleSizes) = DummyElement
            .NameXValuesValuesPlotOrder = sOneLess
        ElseIf Element = eElement.eYValues Then
            .CurrData(eElement.eYValues) = DummyElement
            .NameXValues = sOneLess
        ElseIf Element = eElement.eXValues Then
            .CurrData(eElement.eXValues) = DummyElement
            .Name = sOneLess
        End If
    End With

End Sub


Private Sub GetMultiRangeYXValues( _
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
            iOpeningBracket + 2, _
            LenTitleYXValues - iOpeningBracket - 2 _
    )
End Sub


Private Sub GetSingleRangeYXValues( _
    ByVal LenTitleYXValues As Long, _
    ByVal iMinCommata As Long, _
    ByVal sCurrString As String, _
    ByRef sOneLess As String, _
    ByRef YXValues As String _
)

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
        iComma = InStrRev(sCurrString, ",'", iExclMark)
    End If
    
    sOneLess = Left$(sCurrString, iComma - 1)
    YXValues = Right$( _
            sCurrString, _
            LenTitleYXValues - iComma _
    )

End Sub


Private Function GetNoOfSeparatorsInPossibleStringTitle( _
    ByVal sCurrString As String _
        ) As Long

    'no string title
    If Left$(sCurrString, 1) <> Chr$(34) Then Exit Function
    
    Dim iNoOfQuotes As Long
    iNoOfQuotes = NoOfOccurrencesInString(sCurrString, Chr$(34))
    
    If iNoOfQuotes = 2 Then
        Dim sTitleString As String
        sTitleString = GetStringTitle(sCurrString)
        
        GetNoOfSeparatorsInPossibleStringTitle = _
                NoOfOccurrencesInString(sTitleString, pcsSep)
    Else
        'then it gets complicated, because either the string title itself
        'contains one or more quotes or the Worksheet.Name contains one or
        'more quotes
        '--> return a negative value and hope that it will work
        GetNoOfSeparatorsInPossibleStringTitle = -1
    End If

End Function


Private Function GetStringTitle( _
    ByVal sCurrString As String _
        ) As String

    Dim iLastQuotePos As Long
    iLastQuotePos = InStrRev(sCurrString, Chr$(34))
    
    Dim sTitleString As String
    sTitleString = Mid$(sCurrString, 2, iLastQuotePos - 2)
    
    GetStringTitle = sTitleString

End Function


Private Function IsRangeExternal( _
    ByVal RangeString As String _
        ) As Boolean
    
    IsRangeExternal = False
    
    If InStr(RangeString, "\") > 0 Then
        Dim bIsExternal As Boolean
        bIsExternal = True
    ElseIf InStr(RangeString, "[") > 0 And InStr(RangeString, "]") > 0 Then
        bIsExternal = True
    End If
    
    IsRangeExternal = bIsExternal

End Function


Private Function IsNamedRange( _
    ByVal RangeString As String _
        ) As Boolean

    If InStr(RangeString, "$") > 0 Then Exit Function
    
    Dim rng As Range
    Set rng = Range(RangeString)
    
    Dim wks As Worksheet
    On Error Resume Next
    Set wks = rng.Parent
    
    If wks Is Nothing Then
        Dim wkb As Workbook
        Set wkb = rng.Parent
    Else
        Set wkb = wks.Parent
    End If
    
    IsNamedRange = RangeExists(wkb, RangeString)

End Function


'external ranges can crash Excel (at least when the file is not open)
'thus only set the range when the source file is open
Private Sub HandleExternalRange( _
    ByVal RangeString As String, _
    ByRef DummyElement As tElement _
)

    Dim sWkbWks As String
    sWkbWks = ExtractWkbWksPartOfFullRange(RangeString)
    
    If InStr(sWkbWks, "\") = 0 Then
        HandleOpenExternalRange _
                sWkbWks, _
                RangeString, _
                DummyElement
    Else
        HandleClosedExternalRange _
                DummyElement
    End If

End Sub


Private Sub HandleOpenExternalRange( _
    ByVal sWkbWks As String, _
    ByVal RangeString As String, _
    ByRef DummyElement As tElement _
)

    Dim wks As Worksheet
    Set wks = SetWorksheet(sWkbWks)
    
    Dim sRng As String
    sRng = ExtractRange( _
            RangeString, _
            sWkbWks _
    )
    
    With DummyElement
        .Type = "Open External Range"
        Set .Range = wks.Range(sRng)
    End With

End Sub


Private Sub HandleClosedExternalRange( _
    ByRef DummyElement As tElement _
)

    With DummyElement
        .Type = "Closed External Range"
        Set .Range = Nothing
    End With

End Sub


Private Sub HandleLongRangeCase( _
    ByVal RangeString As String, _
    ByRef DummyElement As tElement _
)

    Dim sWks As String
    sWks = ExtractWkbWksPartOfFullRange(RangeString)
    
    Dim wks As Worksheet
    Set wks = SetWorksheet(sWks)
    
    Dim sRng As String
    sRng = ExtractRange( _
            RangeString, _
            sWks _
    )
    
    If Len(sRng) <= 255 Then
        With DummyElement
            .Type = "Range"
            Set .Range = wks.Range(sRng)
        End With
    Else
        Dim arrRanges() As String
        arrRanges = Split(sRng, ",")
        
        Dim i As Long
        For i = LBound(arrRanges) To UBound(arrRanges)
            Dim rngUnion As Range
            Set rngUnion = Union2(rngUnion, wks.Range(arrRanges(i)))
        Next
        
        With DummyElement
            .Type = "Long Range"
            Set .Range = rngUnion
            .LongRangeString = sRng
        End With
    End If
    
End Sub


Private Function ExtractWkbWksPartOfFullRange( _
    ByVal RangeString As String _
        ) As String
    
    Dim sApostrophe As String
    If Left$(RangeString, 1) = pcsApostrophe Then
        sApostrophe = pcsApostrophe
    End If
    
    Dim sSeparator As String
    sSeparator = sApostrophe & "!$"
    
    Dim iSeparator As Long
    iSeparator = InStr(RangeString, sSeparator)
    
    If iSeparator = 0 Then Exit Function
    
    If Len(sApostrophe) = 0 Then
        Dim sWkbWks As String
        sWkbWks = Left$(RangeString, iSeparator - 1)
    Else
        sWkbWks = Mid$(RangeString, 2, iSeparator - 2)
    End If
    
    ExtractWkbWksPartOfFullRange = sWkbWks
    
End Function


Private Function SetWorksheet( _
    ByVal sWks As String _
        ) As Worksheet
    
    If Left$(sWks, 1) = pcsApostrophe Then
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
        Set SetWorkbook = ActiveWorkbook
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
    
    Dim sApostrophe As String
    If Left$(RangeString, 1) = pcsApostrophe Then
        sApostrophe = pcsApostrophe
    End If
    
    Dim sSeparator As String
    sSeparator = sApostrophe & "!$"
    
    Dim sRangePrefix As String
    sRangePrefix = sApostrophe & sWkbWks & sApostrophe & "!"
    
    Dim sRngWithoutWkbWksName As String
    sRngWithoutWkbWksName = Replace(RangeString, sRangePrefix, vbNullString)
    
    Dim sRng As String
    sRng = Replace(sRngWithoutWkbWksName, "$", vbNullString)
    
    ExtractRange = sRng
    
End Function


Private Sub GetSeriesName()

    Dim LenName As Long
    Dim SeriesName As String
    
    
    LenName = Len(This.Name)
    
    With This.CurrData(eElement.eName)
        If LenName = 0 Then
            SeriesName = vbNullString
            .Type = "Empty"
            Set .Range = Nothing
        ElseIf Left$(This.Name, 1) = Chr$(34) Then
            SeriesName = Mid$(This.Name, 2, LenName - 2)
            .Type = "String"
            Set .Range = Nothing
        Else
            SeriesName = This.Name
            If IsRangeExternal(SeriesName) Then
                HandleExternalRange _
                        SeriesName, _
                        This.CurrData(eElement.eName)
            Else
                .Type = "Range"
                Set .Range = Range(SeriesName)
            End If
        End If
        .Value = SeriesName
    End With

End Sub


Private Sub GetWorkbooks()
    
    Dim sChartWkb As String
    sChartWkb = GetChartWkb
    
    With This
        Dim Element As eElement
        For Element = LBound(.CurrBook) To UBound(.CurrBook)
            'exit if SeriesName, x or y values are not given as a range
            Select Case .CurrData(Element).Type
                Case "Range", "Open External Range"
                    .CurrBook(Element) = GetCurrWorksheetOrBook(Element, sChartWkb, True)
                Case Else
                    Select Case Element
                        Case eElement.eName
                            .CurrBook(Element) = "ERROR - TITLE IS NOT A RANGE"
                        Case eElement.eXValues
                            .CurrBook(Element) = "ERROR - XVALUES ARE NOT A RANGE"
                        Case eElement.eYValues
                            .CurrBook(Element) = "ERROR - VALUES ARE NOT A RANGE"
                        Case eElement.eBubbleSizes
                            .CurrBook(Element) = pcsBadBubbleChart
                    End Select
            End Select
        Next
    End With
    
End Sub


Private Sub GetWorksheets()
    
    Dim sChartWkb As String
    sChartWkb = GetChartWkb
    
    With This
        Dim Element As eElement
        For Element = LBound(.CurrSheet) To UBound(.CurrSheet)
            'exit if SeriesName, x or y values are not given as a range
            Select Case .CurrData(Element).Type
                Case "Range", "Open External Range"
                    .CurrSheet(Element) = GetCurrWorksheetOrBook(Element, sChartWkb, False)
                Case Else
                    Select Case Element
                        Case eElement.eName
                            .CurrSheet(Element) = "ERROR - TITLE IS NOT A RANGE"
                        Case eElement.eXValues
                            .CurrSheet(Element) = "ERROR - XVALUES ARE NOT A RANGE"
                        Case eElement.eYValues
                            .CurrSheet(Element) = "ERROR - VALUES ARE NOT A RANGE"
                        Case eElement.eBubbleSizes
                            .CurrSheet(Element) = pcsBadBubbleChart
                    End Select
            End Select
        Next
    End With
    
End Sub


Private Function GetChartWkb() As String

    Dim crt As ChartObject
    Dim wks As Worksheet
    Dim wkb As Workbook
    
    With This.CurrChart
        On Error Resume Next
        Set wkb = .Parent
        
        If Err.Number <> 0 Then
            Set crt = .Parent
            Set wks = crt.Parent
            Set wkb = wks.Parent
        End If
        On Error GoTo 0
    End With
    
    GetChartWkb = wkb.Name

End Function


Private Function GetCurrWorksheetOrBook( _
    ByVal Element As eElement, _
    ByVal sChartWkb As String, _
    Optional ByVal SheetOrBook As Boolean = False _
        ) As String

    With This.CurrData(Element)
        If .Range Is Nothing Then
            GetCurrWorksheetOrBook = "External Range"
            Exit Function
        Else
            Dim rng As Range
            Set rng = .Range.Areas(1)
        End If
    End With
    
    Dim wks As Worksheet
    Set wks = rng.Parent
    
    If Not SheetOrBook Then
        GetCurrWorksheetOrBook = wks.Name
    Else
        Dim wkb As Workbook
        Set wkb = wks.Parent
        GetCurrWorksheetOrBook = wkb.Name
    End If

End Function


Private Function GetApostrophe( _
    ByVal Element As eElement _
        ) As String

    With This
        If Left$(This.CurrData(Element).Value, 2) = "(" & pcsApostrophe Or _
                Left$(This.CurrData(Element).Value, 1) = pcsApostrophe Then
            GetApostrophe = pcsApostrophe
        Else
            GetApostrophe = vbNullString
        End If
    End With

End Function


'------------------------------------------------------------------------------
'returns the number of (data) points in the chart's SERIES
Private Function SERIESFormulaDataPoints() As Long
    With This
        SERIESFormulaDataPoints = .CurrChart.SeriesCollection(.CurrSeries).Points.Count
    End With
End Function


'returns the range of the given point
Private Function PointRange( _
    ByVal Element As eElement, _
    ByVal PointNumber As Long _
        ) As Variant
    
    Dim sSeries As String
    Dim sSheetName As String
    Dim rng As Range
    Dim sWorksheetName As String
    Dim arRange() As String
    
    
    'exit if PointNumber is greater than the number of points existent
    'in 'CurrSeries'
    If PointNumber > SERIESFormulaDataPoints() Then
        PointRange = "ERROR - BAD POINTNUMBER"
        Exit Function
    End If
    
    'exit if x or y values are not given as a range
    Call SERIESFormulaElement
    
    With This
        If .CurrData(Element).Type <> "Range" Then
            If Element = eElement.eXValues Then
                PointRange = "ERROR - XVALUES ARE NOT A RANGE"
            ElseIf Element = eElement.eYValues Then
                PointRange = "ERROR - VALUES ARE NOT A RANGE"
            End If
            Exit Function
        End If
        
        'store the 'Worksheet.Name'
        sWorksheetName = .CurrSheet(Element)
        
        'get the range string to parse for the address and store only the
        'range without the 'Worksheet.Name'
        Set rng = .CurrData(Element).Range
    End With
    sSeries = rng.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    'get an array with all cells of the SERIES
    arRange = SplitRangeToSingleCells(sSeries)
    
    PointRange = sWorksheetName & "!" & arRange(PointNumber - 1)

End Function


'==============================================================================
Private Function RemovePossibleApostrophes( _
    ByVal sWorksheet As String _
        ) As String

    Dim iLen As Long
    Dim sClean As String
    
    
    If Left$(sWorksheet, 1) = pcsApostrophe Then
        iLen = Len(sWorksheet)
        sClean = Mid$(sWorksheet, 2, iLen - 2)
    Else
        sClean = sWorksheet
    End If
    
    RemovePossibleApostrophes = sClean

End Function


'function that returns an array where all cells of a range are separated
Private Function SplitRangeToSingleCells( _
    ByVal sRange As String _
        ) As Variant
    
    Dim arRange() As String
    Dim arRangeNew() As String
    Dim arCRange() As String
    Dim i As Long
    Dim j As Long
    Dim k As Long     'counter for number of elements
    Dim iStartRC As Long
    Dim iEndRC As Long
    Dim iRow As Long
    Dim sColumn As String
    
    
    'first split the 'Areas'
    arRange = Split(sRange, pcsSep)
    
    'initialize 'k'
    k = 0
    
    'now test, if in the areas are (continuous) ranges
    'if so, further split them
    For i = LBound(arRange) To UBound(arRange)
        If InStr(arRange(i), ":") Then
            'extract the numbers from the (continuous) range
            arCRange = Split(arRange(i), ":")
            iStartRC = ExtractNumber(arCRange(LBound(arCRange)))
            iEndRC = ExtractNumber(arCRange(UBound(arCRange)))
            'test if given row by row
            If iStartRC <> iEndRC Then
                'find column
                sColumn = Left$(arCRange(LBound(arCRange)), Len(arCRange(LBound(arCRange))) - Len(CStr(iStartRC)))
                
                'now create all the cells in the range
                If i = LBound(arRange) Then
                    ReDim arRangeNew(iEndRC - iStartRC)
                Else
                    ReDim Preserve arRangeNew(UBound(arRangeNew) + iEndRC - iStartRC + 1)
                End If
                For j = iStartRC To iEndRC
                    arRangeNew(k) = sColumn & CStr(j)
                    k = k + 1
                Next
            'else given column by column
            Else
                'store row number
                iRow = iStartRC
                'extract column (numbers)
                sColumn = Left$(arCRange(LBound(arCRange)), Len(arCRange(LBound(arCRange))) - Len(CStr(iRow)))
                iStartRC = ColumnLetterToNumber(sColumn)
                sColumn = Left$(arCRange(UBound(arCRange)), Len(arCRange(UBound(arCRange))) - Len(CStr(iRow)))
                iEndRC = ColumnLetterToNumber(sColumn)
                
                'now create all the cells in the range
                If i = LBound(arRange) Then
                    ReDim arRangeNew(iEndRC - iStartRC)
                Else
                    ReDim Preserve arRangeNew(UBound(arRangeNew) + iEndRC - iStartRC + 1)
                End If
                For j = iStartRC To iEndRC
                    arRangeNew(k) = ColumnNumberToLetter(j) & CStr(iRow)
                    k = k + 1
                Next
            End If
        Else
            If i = LBound(arRange) Then
                ReDim arRangeNew(k)
            Else
                ReDim Preserve arRangeNew(k)
            End If
            arRangeNew(k) = arRange(i)
            k = k + 1
        End If
    Next
    
    SplitRangeToSingleCells = arRangeNew
    
End Function


Private Function GetRangeAddress( _
    ByVal Vals As Variant _
        ) As String
    
    Dim iNoOfRangeAreas As Long
    iNoOfRangeAreas = Vals.Areas.Count
    
    If iNoOfRangeAreas = 1 Then
        Dim sRange As String
        sRange = "=" & CombineWorksheetAndRange(Vals)
    Else
        sRange = "=(" & CombineWorksheetAndRange(Vals.Areas(1))
        
        Dim i As Long
        For i = 2 To iNoOfRangeAreas
            sRange = sRange & pcsSep & CombineWorksheetAndRange(Vals.Areas(i))
        Next
        
        sRange = sRange & ")"
    End If
    
    GetRangeAddress = sRange
    
End Function


Private Function CombineWorksheetAndRange( _
    ByVal Vals As Variant _
        ) As String

    Dim wks As Worksheet
    Set wks = Vals.Parent
    
    Dim sWks As String
    sWks = wks.Name
    
    Dim sRange As String
    sRange = Vals.Address(False, False)
    
    CombineWorksheetAndRange = _
            pcsApostrophe & sWks & pcsApostrophe & "!" & sRange

End Function


'==============================================================================
'' slightly modified version from
'' <http://www.ozgrid.com/VBA/ExtractNum.htm>
'' the string is not given as a range but directly as string
'Extracts a number from a string containing text and numbers.
Private Function ExtractNumber( _
    ByVal sText As String, _
    Optional ByVal Take_decimal As Boolean = False, _
    Optional ByVal Take_negative As Boolean = False _
        ) As Double
    
    Dim iCount As Long
    Dim i As Long
    Dim iLoop As Long
    Dim strNeg As String
    Dim strDec As String
    Dim lNum As String
    Dim vVal As Variant
    
    
    If Take_decimal = True And Take_negative = True Then
        strNeg = "-" 'Negative Sign MUST be before 1st number.
        strDec = "."
    ElseIf Take_decimal = True And Take_negative = False Then
        strNeg = vbNullString
        strDec = "."
    ElseIf Take_decimal = False And Take_negative = True Then
        strNeg = "-"
        strDec = vbNullString
    End If
    iLoop = Len(sText)
    
    For iCount = iLoop To 1 Step -1
        vVal = Mid$(sText, iCount, 1)
        If IsNumeric(vVal) Or vVal = strNeg Or vVal = strDec Then
            i = i + 1
            lNum = Mid$(sText, iCount, 1) & lNum
            If IsNumeric(lNum) Then
                If CDbl(lNum) < 0 Then Exit For
            Else
                lNum = Replace(lNum, Left$(lNum, 1), vbNullString, , 1)
            End If
        End If
        If i = 1 And lNum <> vbNullString Then lNum = CDbl(Mid$(lNum, 1, 1))
    Next iCount
    
    ExtractNumber = CDbl(lNum)
    
End Function


'==============================================================================
'<http://www.excelforum.com/excel-programming/638861-vb-convert-column-letter-to-number.html>
Private Function ColumnLetterToNumber(ByVal strLetter As String) As Long
    ColumnLetterToNumber = ThisWorkbook.Worksheets(1).Columns(strLetter).Column
End Function

Private Function ColumnNumberToLetter(ByVal lngNumber As Long) As String
    ColumnNumberToLetter = Split(ThisWorkbook.Worksheets(1).Columns(lngNumber).Address(ColumnAbsolute:=False), ":")(0)
End Function


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

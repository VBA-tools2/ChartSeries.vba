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

'SeriesName (read/write)
'XValues (read/write)
'Values (read/write)
'PlotOrder (read/write)

'SeriesNameType (read-only)
'XValuesType (read-only)
'ValuesType (read-only)
'PlotOrderType (read-only)
'
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

Private CurrChart As Chart
Private CurrSeries As Long
Private pOldFullSeriesFormula As String
Private pFullSeriesFormula As String
Private pSeriesFormula As String
Private pNameXValuesValues As String
Private pNameXValues As String
Private pName As String
Private CurrData() As tElement
Private CurrSheet() As String

Private Enum eElement
    [_First] = 1
    eName = eElement.[_First]
    eXValues
    eYValues
    ePlotOrder
    [_Last] = eElement.ePlotOrder
End Enum

Private Type tElement
    Type  As String      'type of element ("Empty", "String", "Integer",
                         '"Array", "Range", External Range")
    Value As Variant     'value of element
    Range As Range       'if type is "Range" then it stores the range
End Type

'==============================================================================
'string to return when a series formula is inaccessible
Private Const pcsInaccessible As String = "inaccessible"
'string to return when a series number is invalid
Private Const pcsBadSeriesNumber As String = "ERROR - BAD SERIES NUMBER"
Private Const pcsSep As String = ","
Private Const pcsApostrophe As String = "'"
'==============================================================================


Public Property Let Formula(ByVal sFormula As String)
    CurrChart.SeriesCollection(CurrSeries).Formula = sFormula
End Property


Public Property Get Chart() As Chart
    Set Chart = CurrChart
End Property

Public Property Let Chart(ByVal cht As Chart)
    Set CurrChart = cht
End Property


Public Property Get ChartSeries() As Long
    ChartSeries = CurrSeries
End Property

Public Property Let ChartSeries(ByVal SeriesNum As Long)
    CurrSeries = SeriesNum
End Property


Public Property Get SeriesName() As Variant
    Call SERIESFormulaElement
    With CurrData(eElement.eName)
        If .Type = "Range" Then
            Set SeriesName = .Range
        Else
            SeriesName = .Value
        End If
    End With
End Property

'2do:
'With this only a string can be set.
'How can a range be set?
Public Property Let SeriesName(ByVal SName)
    CurrChart.SeriesCollection(CurrSeries).Name = SName
End Property


Public Property Get SeriesNameType() As String
    Call SERIESFormulaElement
    SeriesNameType = CurrData(eElement.eName).Type
End Property


Public Property Get XValues() As Variant
    Call SERIESFormulaElement
    With CurrData(eElement.eXValues)
        If .Type = "Range" Then
            Set XValues = .Range
        Else
            XValues = .Value
        End If
    End With
End Property

Public Property Let XValues(XVals)
    CurrChart.SeriesCollection(CurrSeries).XValues = XVals
End Property


Public Property Get XValuesType() As String
    Call SERIESFormulaElement
    XValuesType = CurrData(eElement.eXValues).Type
End Property


Public Property Get Values() As Variant
    Call SERIESFormulaElement
    With CurrData(eElement.eYValues)
        If .Type = "Range" Then
            Set Values = .Range
        Else
            Values = .Value
        End If
    End With
End Property

Public Property Let Values(Vals)
    CurrChart.SeriesCollection(CurrSeries).Values = Vals
End Property


Public Property Get ValuesType() As String
    Call SERIESFormulaElement
    ValuesType = CurrData(eElement.eYValues).Type
End Property


Public Property Get PlotOrder() As Long
    Call SERIESFormulaElement
    PlotOrder = CurrData(eElement.ePlotOrder).Value
End Property

Public Property Let PlotOrder(ByVal PltOrder As Long)
    CurrChart.SeriesCollection(CurrSeries).PlotOrder = PltOrder
End Property


Public Property Get PlotOrderType() As String
    Call SERIESFormulaElement
    PlotOrderType = CurrData(eElement.ePlotOrder).Type
End Property


'determine Worksheet, where the data of SERIES come from
Public Property Get DataSheet(ByVal Element As Long) As String
    Call SERIESFormulaElement
    DataSheet = RemovePossibleApostrophes(CurrSheet(Element))
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
    ReDim CurrData(eElement.[_First] To eElement.[_Last])
    ReDim CurrSheet(eElement.eName To eElement.eYValues)
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
Private Sub SERIESFormulaElement()

    If Not IsCurrSeriesValid Then Exit Sub
    If Not GetFullSeriesFormula Then Exit Sub
    
    If pOldFullSeriesFormula <> pFullSeriesFormula Then
        Call GetSeriesFormula
        Call GetPlotOrder
        Call GetYXValues(False)
        Call GetYXValues(True)
        Call GetSeriesName
        
        Call GetWorksheets
    End If
    
End Sub


Private Function IsCurrSeriesValid() As Boolean

    Dim i As eElement
    
    
    IsCurrSeriesValid = False
    
    If CurrSeries < 1 Or CurrSeries > CurrChart.SeriesCollection.Count Then
        For i = eElement.eName To eElement.eYValues
            With CurrData(i)
                .Type = pcsBadSeriesNumber
                .Value = pcsBadSeriesNumber
            End With
        Next
    Else
        IsCurrSeriesValid = True
    End If

End Function


Private Function GetFullSeriesFormula() As Boolean

    Dim iChartType As XlChartType
    Dim i As eElement
    
    
    GetFullSeriesFormula = False
    pOldFullSeriesFormula = pFullSeriesFormula
    pFullSeriesFormula = vbNullString
    
    With CurrChart
        'in case some data are currently hidden and thus the formula can't be
        'accessed, temporarily change '.PlotVisibleOnly' to 'False'
        Dim bPlotVisible As Boolean
        bPlotVisible = .PlotVisibleOnly
        .PlotVisibleOnly = False
        
        'because there is a possibility that an error is risen, take care of that
        'situation
        '(inspired by <https://peltiertech.com/change-series-formula-improved-routines/>)
        On Error Resume Next
        pFullSeriesFormula = .SeriesCollection(CurrSeries).Formula
        'change to column chart if series is inaccessible
        If Len(pFullSeriesFormula) = 0 Then
            'reset error number
            Err.Number = 0
            'store current 'ChartType' of series ...
            iChartType = .SeriesCollection(CurrSeries).ChartType
            '... and change it temporarily to a column chart
            .SeriesCollection(CurrSeries).ChartType = xlColumnClustered
            
            'now try to read the formula again
            pFullSeriesFormula = .SeriesCollection(CurrSeries).Formula
            
            'restore original 'ChartType'
            .SeriesCollection(CurrSeries).ChartType = iChartType
            
            'if still an error is returned, the formula is inaccessible
            '(this e.g. can be the case if there are #REF errors inside the given ranges)
            If Err.Number <> 0 Then
                For i = eElement.eName To eElement.eYValues
                    With CurrData(i)
                        .Type = pcsInaccessible
                        .Value = pcsInaccessible
                    End With
                Next
                pFullSeriesFormula = pcsInaccessible
                GoTo Exit_Function
            End If
        End If
        On Error GoTo 0
    End With
    
    GetFullSeriesFormula = True
    
Exit_Function:
    CurrChart.PlotVisibleOnly = bPlotVisible

End Function


Private Sub GetSeriesFormula()

    Dim iLenPrefix As Long
    Dim iLenSeriesFormula As Long
    
    '==========================================================================
    Const csPrefix As String = "=SERIES("
    Const ciSuffix As Long = 1
    '==========================================================================
    
    
    iLenPrefix = Len(csPrefix)
    iLenSeriesFormula = Len(pFullSeriesFormula)
    
    pSeriesFormula = _
            Mid( _
                    pFullSeriesFormula, _
                    iLenPrefix + 1, _
                    iLenSeriesFormula - iLenPrefix - ciSuffix _
            )

End Sub


Private Sub GetPlotOrder()

    Dim LastComma As Long
    
    
    LastComma = InStrRev(pSeriesFormula, pcsSep)
    pNameXValuesValues = Left(pSeriesFormula, LastComma - 1)
    With CurrData(eElement.ePlotOrder)
        .Type = "Integer"
        .Value = Mid(pSeriesFormula, LastComma + 1)
        Set .Range = Nothing
    End With

End Sub


Private Sub GetYXValues( _
    ByVal YValues_or_XValues As Boolean _
)

    Dim LenTitleYXValues As Long
    Dim iMinCommata As Long
    Dim iOpeningBracket As Long
    Dim YXValues As String
    Dim sCurrString As String
    Dim sOneLess As String
    Dim DummyElement As tElement
    
    
    If YValues_or_XValues = False Then
        sCurrString = pNameXValuesValues
        iMinCommata = 2
    Else
        sCurrString = pNameXValues
        iMinCommata = 1
    End If
    
    LenTitleYXValues = Len(sCurrString)
    
    Select Case Right(sCurrString, 1)
        Case "}"    'Literal Array
            iOpeningBracket = InStrRev(sCurrString, "{")
            sOneLess = Left(sCurrString, iOpeningBracket - 2)
            YXValues = Mid( _
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
            sOneLess = Left(sCurrString, LenTitleYXValues - 1)
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
    
    With DummyElement
        If Len(YXValues) = 0 Then
            .Type = "Empty"
            Set .Range = Nothing
        ElseIf Right(sCurrString, 1) = "}" Then
            .Type = "Array"
            Set .Range = Nothing
        Else
            'external ranges can crash Excel (at least when the file is not open)
            'thus only set the range when it is an "internal" range
            If IsRangeExternal(YXValues) Then
                .Type = "External Range"
                Set .Range = Nothing
            Else
                .Type = "Range"
                Set .Range = Range(YXValues)
            End If
        End If
        .Value = YXValues
    End With
    
    If YValues_or_XValues = False Then
        CurrData(eElement.eYValues) = DummyElement
        pNameXValues = sOneLess
    Else
        CurrData(eElement.eXValues) = DummyElement
        pName = sOneLess
    End If

End Sub


Private Sub GetMultiRangeYXValues( _
    ByVal LenTitleYXValues, _
    ByVal sCurrString, _
    ByRef sOneLess, _
    ByRef YXValues _
)

    Dim iOpeningBracket As Long
    iOpeningBracket = InStrRev(sCurrString, ",(")
    
    sOneLess = Left(sCurrString, iOpeningBracket - 1)
    YXValues = Mid( _
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
    
    sOneLess = Left(sCurrString, iComma - 1)
    YXValues = Right( _
            sCurrString, _
            LenTitleYXValues - iComma _
    )

End Sub


Private Function GetNoOfSeparatorsInPossibleStringTitle( _
    ByVal sCurrString As String _
        ) As Long

    'no string title
    If Left(sCurrString, 1) <> Chr(34) Then Exit Function
    
    Dim iNoOfQuotes As Long
    iNoOfQuotes = NoOfOccurrencesInString(sCurrString, Chr(34))
    
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
    iLastQuotePos = InStrRev(sCurrString, Chr(34))
    
    Dim sTitleString As String
    sTitleString = Mid(sCurrString, 2, iLastQuotePos - 2)
    
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


Private Sub GetSeriesName()

    Dim LenName As Long
    Dim SeriesName As String
    
    
    LenName = Len(pName)
    
    With CurrData(eElement.eName)
        If LenName = 0 Then
            SeriesName = vbNullString
            .Type = "Empty"
            Set .Range = Nothing
        ElseIf Left(pName, 1) = Chr(34) Then
            SeriesName = Mid(pName, 2, LenName - 2)
            .Type = "String"
            Set .Range = Nothing
        Else
            SeriesName = pName
            'external ranges can crash Excel (at least when the file is not open)
            'thus only set the range when it is an "internal" range
            If IsRangeExternal(SeriesName) Then
                .Type = "External Range"
                Set .Range = Nothing
            Else
                .Type = "Range"
                Set .Range = Range(SeriesName)
            End If
        End If
        .Value = SeriesName
    End With

End Sub


Private Sub GetWorksheets()
    
    Dim Element As eElement
    Dim sChartWkb As String
    
    
    sChartWkb = GetChartWkb
    
    For Element = LBound(CurrSheet) To UBound(CurrSheet)
        'exit if x or y values are not given as a range
        If CurrData(Element).Type <> "Range" Then
            Select Case Element
                Case eElement.eName
                    CurrSheet(Element) = "ERROR - TITLE IS NOT A RANGE"
                Case eElement.eXValues
                    CurrSheet(Element) = "ERROR - XVALUES ARE NOT A RANGE"
                Case eElement.eYValues
                    CurrSheet(Element) = "ERROR - VALUES ARE NOT A RANGE"
            End Select
        Else
            CurrSheet(Element) = GetCurrWorksheet(Element, sChartWkb)
        End If
    Next
    
End Sub


Private Function GetChartWkb() As String

    Dim crt As ChartObject
    Dim wks As Worksheet
    Dim wkb As Workbook
    
    With CurrChart
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


Private Function GetCurrWorksheet( _
    ByVal Element As eElement, _
    ByVal sChartWkb As String _
        ) As String

    With CurrData(Element)
        If .Range Is Nothing Then
            GetCurrWorksheet = "External Range"
            Exit Function
        Else
            Dim rng As Range
            Set rng = .Range.Areas(1)
        End If
    End With
    
    Dim wks As Worksheet
    Set wks = rng.Parent
    
    Dim wkb As Workbook
    Set wkb = wks.Parent
    
    Dim sWks As String
    sWks = wks.Name
    
    Dim sWkb As String
    sWkb = wkb.Name
    
    Dim sApostrophe As String
    sApostrophe = GetApostrophe(Element)
    
    Dim sWorksheet As String
    sWorksheet = sApostrophe
    
    If sWkb <> sChartWkb Then
        sWorksheet = sWorksheet & "[" & sWkb & "]"
    End If
    
    sWorksheet = sWorksheet & sWks & sApostrophe
    GetCurrWorksheet = sWorksheet

End Function


Private Function GetApostrophe( _
    ByVal Element As eElement _
        ) As String

    If Left(CurrData(Element).Value, 2) = "(" & pcsApostrophe Or _
            Left(CurrData(Element).Value, 1) = pcsApostrophe Then
        GetApostrophe = pcsApostrophe
    Else
        GetApostrophe = vbNullString
    End If

End Function


'------------------------------------------------------------------------------
'returns the number of (data) points in the chart's SERIES
Private Function SERIESFormulaDataPoints() As Long
    SERIESFormulaDataPoints = CurrChart.SeriesCollection(CurrSeries).Points.Count
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
    If CurrData(Element).Type <> "Range" Then
        If Element = eElement.eXValues Then
            PointRange = "ERROR - XVALUES ARE NOT A RANGE"
        ElseIf Element = eElement.eYValues Then
            PointRange = "ERROR - VALUES ARE NOT A RANGE"
        End If
        Exit Function
    End If
    
    'store the 'Worksheet.Name'
    sWorksheetName = CurrSheet(Element)
    
    'get the range string to parse for the address and store only the
    'range without the 'Worksheet.Name'
    Set rng = CurrData(Element).Range
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
    
    
    If Left(sWorksheet, 1) = pcsApostrophe Then
        iLen = Len(sWorksheet)
        sClean = Mid(sWorksheet, 2, iLen - 2)
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
                sColumn = Left(arCRange(LBound(arCRange)), Len(arCRange(LBound(arCRange))) - Len(CStr(iStartRC)))
                
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
                sColumn = Left(arCRange(LBound(arCRange)), Len(arCRange(LBound(arCRange))) - Len(CStr(iRow)))
                iStartRC = ColumnLetterToNumber(sColumn)
                sColumn = Left(arCRange(UBound(arCRange)), Len(arCRange(UBound(arCRange))) - Len(CStr(iRow)))
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
        vVal = Mid(sText, iCount, 1)
        If IsNumeric(vVal) Or vVal = strNeg Or vVal = strDec Then
            i = i + 1
            lNum = Mid(sText, iCount, 1) & lNum
            If IsNumeric(lNum) Then
                If CDbl(lNum) < 0 Then Exit For
            Else
                lNum = Replace(lNum, Left(lNum, 1), vbNullString, , 1)
            End If
        End If
        If i = 1 And lNum <> vbNullString Then lNum = CDbl(Mid(lNum, 1, 1))
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

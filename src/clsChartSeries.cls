VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'=======================
'last update: 2014-12-04
'=======================

'This Class module provides an easy way to access the items in a
'chart's SERIES formula. It can be exported and then imported into
'any project

'Developed by John Walkenbach, JWALK AND ASSOCIATES
'Copyright 1999. All rights reserved.
'May be used and distributed freely, but may not be sold.
'http://www.j-walk.com/ss/

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
'    Dim MySeries As New clsChartSeries
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
Dim CurrChart As Chart     'accessible to all procedures
Dim CurrSeries As Integer  'accessible to all procedures


Property Get Chart() As Chart
    Set Chart = CurrChart
End Property

Property Let Chart(Cht)
    Set CurrChart = Cht
End Property


Property Get ChartSeries()
    ChartSeries = CurrSeries
End Property

Property Let ChartSeries(SeriesNum)
    CurrSeries = SeriesNum
End Property


Property Get SeriesName() As Variant
    If SeriesNameType = "Range" Then
        Set SeriesName = Range(SERIESFormulaElement(CurrChart, CurrSeries, 1))
    Else
        SeriesName = SERIESFormulaElement(CurrChart, CurrSeries, 1)
    End If
End Property

Property Let SeriesName(SName)
    CurrChart.SeriesCollection(CurrSeries).Name = SName
End Property


Property Get SeriesNameType() As String
    SeriesNameType = SERIESFormulaElementType(CurrChart, CurrSeries, 1)
End Property


Property Get XValues() As Variant
    If XValuesType = "Range" Then
        Set XValues = Range(SERIESFormulaElement(CurrChart, CurrSeries, 2))
    Else
        XValues = SERIESFormulaElement(CurrChart, CurrSeries, 2)
    End If
End Property

Property Let XValues(XVals)
    CurrChart.SeriesCollection(CurrSeries).XValues = XVals
End Property


Property Get XValuesType() As String
    XValuesType = SERIESFormulaElementType(CurrChart, CurrSeries, 2)
End Property


Property Get Values() As Variant
    If ValuesType = "Range" Then
        Set Values = Range(SERIESFormulaElement(CurrChart, CurrSeries, 3))
    Else
        Values = SERIESFormulaElement(CurrChart, CurrSeries, 3)
    End If
End Property

Property Let Values(Vals)
    CurrChart.SeriesCollection(CurrSeries).Values = Vals
End Property


Property Get ValuesType() As String
    ValuesType = SERIESFormulaElementType(CurrChart, CurrSeries, 3)
End Property


Property Get PlotOrder()
    PlotOrder = SERIESFormulaElement(CurrChart, CurrSeries, 4)
End Property

Property Let PlotOrder(PltOrder)
    CurrChart.SeriesCollection(CurrSeries).PlotOrder = PltOrder
End Property


Property Get PlotOrderType() As String
    PlotOrderType = SERIESFormulaElementType(CurrChart, CurrSeries, 4)
End Property


'------------------------------------------------------------------------------

'Returns a string that describes the element of a chart's SERIES formula
'This function essentially parses and analyzes a SERIES formula
'
'Element 1: Series Name. Returns "Range" , "Empty", or "String"
'Element 2: XValues. Returns "Range", "Empty", or "Array"
'Element 3: Values. Returns "Range" or "Array"
'Element 4: PlotOrder. Always returns "Integer"
Private Function SERIESFormulaElementType(ChartObj, SeriesNum, Element) As String

    Dim SeriesFormula As String
    Dim FirstComma As Integer, SecondComma As Integer, LastComma As Integer
    Dim FirstParen As Integer, SecondParen As Integer
    Dim FirstBracket As Integer, SecondBracket As Integer
    
    'as above but reversed
    Dim LenSeriesFormula As Integer    'to store the length of 'SeriesFormula'
    Dim rSeriesFormula As String
    Dim rFirstComma As Integer, rSecondComma As Integer, rLastComma As Integer
    Dim rFirstParen As Integer, rSecondParen As Integer
    Dim rFirstBracket As Integer, rSecondBracket As Integer
    
    Dim SeriesName, XValues, Values, PlotOrder As Integer
    
    
    'Exit if Surface chart (surface chrarts do not have SERIES formulas)
    If ChartObj.ChartType >= 83 And ChartObj.ChartType <= 86 Then
        SERIESFormulaElementType = "ERROR - SURFACE CHART"
        Exit Function
    End If
    
    'Exit if nonexistent series is specified
    If SeriesNum > ChartObj.SeriesCollection.Count Or SeriesNum < 1 Then
        SERIESFormulaElementType = "ERROR - BAD SERIES NUMBER"
        Exit Function
    End If
    
    'Exit if element is > 4
    If Element > 4 Or Element < 1 Then
        SERIESFormulaElementType = "ERROR - BAD ELEMENT NUMBER"
        Exit Function
    End If
    
    'Get the SERIES formula
    SeriesFormula = ChartObj.SeriesCollection(SeriesNum).Formula
    rSeriesFormula = StrReverse(SeriesFormula)
    LenSeriesFormula = Len(SeriesFormula)
    
'-----
'adapt because string _can_ now be longer than 255
'(also in 'SERIESFormulaElement')
    'Exit if SERIES formula is longer than 255 characters
    '(test if length of SeriesFormula = 255 and the last character isn't a ")")
    If LenSeriesFormula = 255 And Left(rSeriesFormula, 1) <> ")" Then
        SERIESFormulaElementType = "ERROR - FORMULA TOO LONG"
        Exit Function
    End If
'-----
    
    '-----
    'Get the Fourth Element (Plot Order)
    '-----
    rLastComma = InStr(1, rSeriesFormula, ",")
    LastComma = LenSeriesFormula - rLastComma + 1
    
    If Element = 4 Then
        PlotOrder = Mid(SeriesFormula, LastComma + 1, Len(SeriesFormula) - LastComma - 1)
        SERIESFormulaElementType = "Integer"
        Exit Function
    End If
    
    '-----
    'Get the Third Element (Y Range)
    '-----
    If Mid(rSeriesFormula, rLastComma + 1, 1) = ")" Then
        'Multiple ranges
        rSecondParen = InStr(rLastComma, rSeriesFormula, ")")
        rFirstParen = InStr(rSecondParen, rSeriesFormula, "(,")
        
        FirstParen = LenSeriesFormula - rFirstParen + 1
        SecondParen = LenSeriesFormula - rSecondParen + 1
        Values = Mid(SeriesFormula, FirstParen + 1, SecondParen - FirstParen - 1)
        SecondComma = FirstParen - 1
    Else
        If Mid(rSeriesFormula, rLastComma + 1, 1) = "}" Then
            'Literal Array
            rSecondBracket = InStr(rLastComma, rSeriesFormula, "}")
            rFirstBracket = InStr(rSecondBracket, rSeriesFormula, "{,")
            
            FirstBracket = LenSeriesFormula - rFirstBracket + 1
            SecondBracket = LenSeriesFormula - rSecondBracket + 1
            Values = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            SecondComma = FirstBracket - 1
        Else
            'A single range
            rSecondComma = InStr(rLastComma + 1, rSeriesFormula, ",")
            
            SecondComma = LenSeriesFormula - rSecondComma + 1
            Values = Mid(SeriesFormula, SecondComma + 1, LastComma - SecondComma - 1)
        End If
    End If
    rSecondComma = LenSeriesFormula - SecondComma + 1
    If Element = 3 Then
        If IsRange(Values) Then
            SERIESFormulaElementType = "Range"
        Else
            SERIESFormulaElementType = "Array"
        End If
        Exit Function
    End If
    
    '-----
    'Get the Second Element (X Range)
    '-----
    If Mid(rSeriesFormula, rSecondComma + 1, 1) = ")" Then
        'Multiple ranges
        rSecondParen = InStr(rSecondComma, rSeriesFormula, ")")
        rFirstParen = InStr(rSecondParen, rSeriesFormula, "(,")
        
        FirstParen = LenSeriesFormula - rFirstParen + 1
        SecondParen = LenSeriesFormula - rSecondParen + 1
        XValues = Mid(SeriesFormula, FirstParen, SecondParen - FirstParen + 1)
        FirstComma = FirstParen - 1
    Else
        If Mid(rSeriesFormula, rSecondComma + 1, 1) = "}" Then
            'Literal Array
            rSecondBracket = InStr(rSecondComma, rSeriesFormula, "}")
            rFirstBracket = InStr(rSecondBracket, rSeriesFormula, "{,")
            
            FirstBracket = LenSeriesFormula - rFirstBracket + 1
            SecondBracket = LenSeriesFormula - rSecondBracket + 1
            XValues = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            FirstComma = FirstBracket - 1
        Else
            'A single range
            rFirstComma = InStr(rSecondComma + 1, rSeriesFormula, ",")
            
            FirstComma = LenSeriesFormula - rFirstComma + 1
            XValues = Mid(SeriesFormula, FirstComma + 1, SecondComma - FirstComma - 1)
        End If
    End If
    If Element = 2 Then
        If IsRange(XValues) Then
            SERIESFormulaElementType = "Range"
        Else
            If XValues = "" Then
                SERIESFormulaElementType = "Empty"
            Else
                SERIESFormulaElementType = "Array"
            End If
        End If
        Exit Function
    End If
    
    '-----
    'Get the First Element (Series Name)
    '-----
    FirstParen = InStr(1, SeriesFormula, "(")
    SeriesName = Mid(SeriesFormula, FirstParen + 1, FirstComma - FirstParen - 1)
    If Element = 1 Then
        If IsRange(SeriesName) Then
            SERIESFormulaElementType = "Range"
        Else
            If SeriesName = "" Then
                SERIESFormulaElementType = "Empty"
            Else
                If TypeName(SeriesName) = "String" Then
                    SERIESFormulaElementType = "String"
                End If
            End If
        End If
        Exit Function
    End If
    
End Function


'Returns one of four elements in a chart's SERIES formula (as a string)
'This function essentially parses and analyzes a SERIES formula
'
'Element 1: Series Name. Can be a range reference, a literal value, or nothing
'Element 2: XValues. Can be a range reference (including a non-contiguous range), a literal array, or nothing
'Element 3: Values. Can be a range reference (including a non-contiguous range), or a literal array
'Element 4: PlotOrder. Must be an integer
Private Function SERIESFormulaElement(ChartObj, SeriesNum, Element) As String
    
    Dim SeriesFormula As String
    Dim FirstComma As Integer, SecondComma As Integer, LastComma As Integer
    Dim FirstParen As Integer, SecondParen As Integer
    Dim FirstBracket As Integer, SecondBracket As Integer
    Dim StartY As Integer
    Dim SeriesName, XValues, Values, PlotOrder As Integer
    
    'as above but reversed
    Dim LenSeriesFormula As Integer    'to store the length of 'SeriesFormula'
    Dim rSeriesFormula As String
    Dim rFirstComma As Integer, rSecondComma As Integer, rLastComma As Integer
    Dim rFirstParen As Integer, rSecondParen As Integer
    Dim rFirstBracket As Integer, rSecondBracket As Integer
    
    
    'Exit if Surface chart (surface chrarts do not have SERIES formulas)
    If ChartObj.ChartType >= 83 And ChartObj.ChartType <= 86 Then
        SERIESFormulaElement = "ERROR - SURFACE CHART"
        Exit Function
    End If
    
    'Exit if nonexistent series is specified
    If SeriesNum > ChartObj.SeriesCollection.Count Or SeriesNum < 1 Then
        SERIESFormulaElement = "ERROR - BAD SERIES NUMBER"
        Exit Function
    End If
    
    'Exit if element is > 4
    If Element > 4 Then
        SERIESFormulaElement = "ERROR - BAD ELEMENT NUMBER"
        Exit Function
    End If
    
    'Get the SERIES formula
    SeriesFormula = ChartObj.SeriesCollection(SeriesNum).Formula
    rSeriesFormula = StrReverse(SeriesFormula)
    LenSeriesFormula = Len(SeriesFormula)
    
    'Exit if SERIES formula is longer than 255 characters
    '(test if length of SeriesFormula = 255 and the last character isn't a ")")
    If LenSeriesFormula = 255 And Left(rSeriesFormula, 1) <> ")" Then
        SERIESFormulaElement = "ERROR - FORMULA TOO LONG"
        Exit Function
    End If
    
    
    '-----
    'Get the Fourth Element (Plot Order)
    '-----
    rLastComma = InStr(1, rSeriesFormula, ",")
    LastComma = LenSeriesFormula - rLastComma + 1
    
    If Element = 4 Then
        PlotOrder = Mid(SeriesFormula, LastComma + 1, Len(SeriesFormula) - LastComma - 1)
        SERIESFormulaElement = PlotOrder
        Exit Function
    End If
    
    
    '-----
    'Get the Third Element (Y Range)
    '-----
    If Mid(rSeriesFormula, rLastComma + 1, 1) = ")" Then
        'Multiple ranges
        rSecondParen = InStr(rLastComma, rSeriesFormula, ")")
        rFirstParen = InStr(rSecondParen, rSeriesFormula, "(,")
        
        FirstParen = LenSeriesFormula - rFirstParen + 1
        SecondParen = LenSeriesFormula - rSecondParen + 1
        Values = Mid(SeriesFormula, FirstParen + 1, SecondParen - FirstParen - 1)
        SecondComma = FirstParen - 1
    Else
        If Mid(rSeriesFormula, rLastComma + 1, 1) = "}" Then
            'Literal Array
            rSecondBracket = InStr(rLastComma, rSeriesFormula, "}")
            rFirstBracket = InStr(rSecondBracket, rSeriesFormula, "{,")
            
            FirstBracket = LenSeriesFormula - rFirstBracket + 1
            SecondBracket = LenSeriesFormula - rSecondBracket + 1
            Values = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            SecondComma = FirstBracket - 1
        Else
            'A single range
            rSecondComma = InStr(rLastComma + 1, rSeriesFormula, ",")
            
            SecondComma = LenSeriesFormula - rSecondComma + 1
            Values = Mid(SeriesFormula, SecondComma + 1, LastComma - SecondComma - 1)
        End If
    End If
    rSecondComma = LenSeriesFormula - SecondComma + 1
    If Element = 3 Then
        SERIESFormulaElement = Values
        Exit Function
    End If
    
    '-----
    'Get the Second Element (X Range)
    '-----
    If Mid(rSeriesFormula, rSecondComma + 1, 1) = ")" Then
        'Multiple ranges
        rSecondParen = InStr(rSecondComma, rSeriesFormula, ")")
        rFirstParen = InStr(rSecondParen, rSeriesFormula, "(,")
        
        FirstParen = LenSeriesFormula - rFirstParen + 1
        SecondParen = LenSeriesFormula - rSecondParen + 1
        XValues = Mid(SeriesFormula, FirstParen, SecondParen - FirstParen + 1)
        FirstComma = FirstParen - 1
    Else
        If Mid(rSeriesFormula, rSecondComma + 1, 1) = "}" Then
            'Literal Array
            rSecondBracket = InStr(rSecondComma, rSeriesFormula, "}")
            rFirstBracket = InStr(rSecondBracket, rSeriesFormula, "{,")
            
            FirstBracket = LenSeriesFormula - rFirstBracket + 1
            SecondBracket = LenSeriesFormula - rSecondBracket + 1
            XValues = Mid(SeriesFormula, FirstBracket, SecondBracket - FirstBracket + 1)
            FirstComma = FirstBracket - 1
        Else
            'A single range
            rFirstComma = InStr(rSecondComma + 1, rSeriesFormula, ",")
            
            FirstComma = LenSeriesFormula - rFirstComma + 1
            XValues = Mid(SeriesFormula, FirstComma + 1, SecondComma - FirstComma - 1)
        End If
    End If
    If Element = 2 Then
        SERIESFormulaElement = XValues
        Exit Function
    End If
    
    '-----
    'Get the First Element (Series Name)
    '-----
    FirstParen = InStr(1, SeriesFormula, "(")
    SeriesName = Mid(SeriesFormula, FirstParen + 1, FirstComma - FirstParen - 1)
    If Element = 1 Then
        SERIESFormulaElement = SeriesName
        Exit Function
    End If
    
End Function


'Returns True if 'ref' is a Range
Private Function IsRange(ref) As Boolean
    Dim x As Range
    Dim sRef As String
    
    sRef = ref
    
    'if the string lenght of 'ref' exceeds 255 characters the below test will
    'fail. Therefore remove 'Worksheet.Name' which should shrink the length
    'to below 255 characters and then test
    If Len(sRef) > 255 Then
        Dim wksName As String
        wksName = ExtractWorksheetName(sRef, True)
        sRef = Replace(sRef, wksName, vbNullString)
    End If
    
    On Error Resume Next
    Set x = Range(sRef)
    If Err = 0 Then IsRange = True Else IsRange = False
End Function


'==============================================================================
'additions by Stefan Pinnow (2014-12-03)
'==============================================================================

'determine number of data points in SERIES
Property Get NoOfPoints() As Long
    NoOfPoints = SERIESFormulaDataPoints(CurrChart, CurrSeries)
End Property

'determine Worksheet, where the data of SERIES come from
Property Get DataSheet(XorY As Long) As String
    DataSheet = SERIESFormulaDataSheet(CurrChart, CurrSeries, XorY)
End Property

'determine the source range of Point 'PointNumber' of SERIES
'the order of the points is as stated in the SERIES, not as shown in the graph
Property Get PointSourceRange(PointNumber As Integer, XorY As Long) As Variant
    PointSourceRange = PointRange(CurrChart, CurrSeries, PointNumber, XorY)
End Property


'------------------------------------------------------------------------------

'returns the number of (data) points in the chart's SERIES
Private Function SERIESFormulaDataPoints(ChartObj, SeriesNum) As Long
    SERIESFormulaDataPoints = ChartObj.SeriesCollection(SeriesNum).Points.Count
End Function


'returns the 'Worksheet.Name' where the data from the SERIES are coming from
Private Function SERIESFormulaDataSheet( _
    ChartObj, SeriesNum, XorY, _
    Optional bWithExclMark As Boolean = False) _
        As Variant
    
    Dim rng As Range
    Dim sRngAddress As String
    
    'exit if x or y values are not given as a range
    If SERIESFormulaElementType(CurrChart, CurrSeries, XorY) <> "Range" Then
        If XorY = 2 Then
            SERIESFormulaDataSheet = "ERROR - XVALUES ARE NOT A RANGE"
        Else
            SERIESFormulaDataSheet = "ERROR - VALUES ARE NOT A RANGE"
        End If
        Exit Function
    End If
    
    'store range of (X)Values
    Set rng = Range(SERIESFormulaElement(CurrChart, CurrSeries, XorY))
    sRngAddress = rng.Address(External:=True)
    
    'call function that extracts the 'Worksheet.Name'
    SERIESFormulaDataSheet = ExtractWorksheetName(sRngAddress, bWithExclMark)
    
End Function


'returns the range of the given point
Private Function PointRange(ChartObj, SeriesNum, PointNumber, XorY) As Variant
    
    Dim sSeries As String
    Dim sSheetName As String
    Dim rng As Range
    Dim sWorksheetName As String
    Dim arRange() As String
    
    
    'exit if PointNumber is greater than the number of points existent
    'in 'SeriesNum'
    If PointNumber > SERIESFormulaDataPoints(CurrChart, CurrSeries) Then
        PointRange = "ERROR - BAD POINTNUMBER"
        Exit Function
    End If
    
    'exit if x or y values are not given as a range
    If SERIESFormulaElementType(CurrChart, CurrSeries, XorY) <> "Range" Then
        If XorY = 2 Then
            PointRange = "ERROR - XVALUES ARE NOT A RANGE"
        Else
            PointRange = "ERROR - VALUES ARE NOT A RANGE"
        End If
        Exit Function
    End If
    
    'store the 'Worksheet.Name'
    sWorksheetName = SERIESFormulaDataSheet(CurrChart, CurrSeries, XorY, True)
    
    'get the range string to parse for the address and store only the
    'range without the 'Worksheet.Name'
    Set rng = Range(SERIESFormulaElement(CurrChart, CurrSeries, XorY))
    sSeries = rng.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    'get an array with all cells of the SERIES
    arRange = SplitRangeToSingleCells(sSeries)
    
    PointRange = sWorksheetName & arRange(PointNumber - 1)

End Function


'==============================================================================

'function that returns the
Private Function ExtractWorksheetName( _
    sRngAddress As String, _
    Optional bWithExclMark As Boolean = False) _
        As String
    
    Dim sRng As String
    Dim iRightEnd As Integer
    Dim sWorkbookName As String
    Dim sApostroph As String
    
    'set 'sRng' to 'sRngAddress'
    sRng = sRngAddress
    
    'test, if string contains more than one cell
    'if so, truncate it to only one cell
    If InStr(sRng, ",") > 0 Then
        sRng = Left(sRng, InStr(sRng, ",") - 1)
    End If
    
    'remove cell reference part so only the 'Workbook.Name' and
    ''Worksheet.Name' remains
    iRightEnd = InStr(sRng, "!")
    sRng = Left(sRng, iRightEnd - 1)
    
    'check, if remaining string starts and ends with "'"
    'if yes, remove them
    'and store in 'sApostroph' apostroph
    If Left(sRng, 1) = "'" Then
        sRng = Mid(sRng, 2, Len(sRng) - 2)
        sApostroph = "'"
    End If
    
    'check, if a 'Workbook.Name' is in the string and if so, extract it
    If Left(sRng, 1) = "[" Then
        iRightEnd = InStrRev(sRng, "]")
        sWorkbookName = Mid(sRng, 2, iRightEnd - 2)
        
        'if 'sWorkbookName' is the same as the 'ActiveWorkbook.Name' then
        'only return the 'Worksheet.Name' otherwise the total remaining string
        If sWorkbookName = ActiveWorkbook.Name Then
            sRng = Right(sRng, Len(sRng) - iRightEnd)
        End If
    End If
    
    If bWithExclMark = True Then
        ExtractWorksheetName = sApostroph & sRng & sApostroph & "!"
    Else
        ExtractWorksheetName = sRng
    End If

End Function


'function that returns an array where all cells of a range are separated
Private Function SplitRangeToSingleCells(sRange As String) As Variant
    
    Dim arRange() As String, arRangeNew() As String
    Dim arCRange() As String
    Dim i As Integer, j As Integer
    Dim k As Integer     'counter for number of elements
    Dim iStartRC As Integer, iEndRC As Integer
    Dim iRow As Integer, sColumn As String
    
    'first split the 'Areas'
    arRange = Split(sRange, ",")
    
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
    sText As String, _
    Optional Take_decimal As Boolean = False, _
    Optional Take_negative As Boolean = False _
        ) As Double
    
    Dim iCount As Integer, i As Integer, iLoop As Integer
    Dim strNeg As String, strDec As String
    Dim lNum As String
    Dim vVal As Variant, vVal2 As Variant
    
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
                lNum = Replace(lNum, Left(lNum, 1), "", , 1)
            End If
        End If
        If i = 1 And lNum <> vbNullString Then lNum = CDbl(Mid(lNum, 1, 1))
    Next iCount
    
    ExtractNumber = CDbl(lNum)
    
End Function


' =============================================================================
' http://www.excelforum.com/excel-programming/638861-vb-convert-column-letter-to-number.html
' -----------------------------------------------------------------------------
Private Function ColumnLetterToNumber(ByVal strLetter As String) As Long
    ColumnLetterToNumber = ThisWorkbook.Worksheets(1).Columns(strLetter).Column
End Function

Private Function ColumnNumberToLetter(ByVal lngNumber As Long) As String
    ColumnNumberToLetter = Split(ThisWorkbook.Worksheets(1).Columns(lngNumber).Address(ColumnAbsolute:=False), ":")(0)
End Function

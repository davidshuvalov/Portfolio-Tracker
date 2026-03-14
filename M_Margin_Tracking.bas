Attribute VB_Name = "M_Margin_Tracking"

Sub CreateContractMarginTracking()
    ' Error handling and performance improvements for margin tracking
    
    On Error GoTo ErrorHandler
    
    ' Declare variables
    Dim wsPortInMarketShort As Worksheet, wsPortInMarketLong As Worksheet
    Dim wsPortfolio As Worksheet, wsPortfolioGraphs As Worksheet
    Dim wsContractMarginTracking As Worksheet
    Dim wsTotalPortfolioM2M As Worksheet
    Dim uniqueSymbols As Object, uniqueSectors As Object
    Dim startdate As Date, currentdate As Date
    Dim dataStartRow As Long, dataEndRow As Long
    
    ' Performance optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' License validation
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        GoTo CleanExit
    End If
    
    ' Validate required worksheets exist
    If Not ValidateRequiredWorksheets(wsPortfolio, wsPortfolioGraphs) Then
        GoTo CleanExit
    End If
    
    ' Initialize other worksheets
    Set wsPortInMarketShort = ThisWorkbook.Sheets("PortInMarketShort")
    Set wsPortInMarketLong = ThisWorkbook.Sheets("PortInMarketLong")
    Set wsTotalPortfolioM2M = ThisWorkbook.Sheets("TotalPortfolioM2M")
    
    ' Initialize column constants
    Call InitializeColumnConstantsManually
    
    ' Create new tracking sheet
    If Not CreateNewTrackingSheet(wsContractMarginTracking) Then
        GoTo CleanExit
    End If
    
    ' Initialize dictionaries for unique symbols and sectors
    Set uniqueSymbols = CreateObject("Scripting.Dictionary")
    Set uniqueSectors = CreateObject("Scripting.Dictionary")
    
    ' Calculate date ranges
    Dim yearsToConsider As Double
    yearsToConsider = GetNamedRangeValue("PortfolioPeriod")
    currentdate = wsPortfolio.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    startdate = DateAdd("yyyy", -Int(yearsToConsider), currentdate)
    startdate = DateAdd("m", -(yearsToConsider - Int(yearsToConsider)) * 12, startdate)
    
    ' Get last row and column
    Dim lastRow As Long, lastCol As Long
    lastRow = wsPortInMarketShort.Cells(wsPortInMarketShort.rows.count, 1).End(xlUp).row
    lastCol = wsPortInMarketShort.Cells(1, wsPortInMarketShort.Columns.count).End(xlToLeft).column
    
    ' Set data start row
    dataStartRow = 7
    dataEndRow = dataStartRow
    
    ' Populate unique symbols and sectors (optimized)
    If Not PopulateUniqueValues(wsPortfolio, wsPortInMarketShort, uniqueSymbols, uniqueSectors, lastCol) Then
        GoTo CleanExit
    End If
    
    
     Application.StatusBar = "Deep Thought is Thinking..."
    
    ' Initialize arrays for data storage
    Dim resultArray() As Variant
    Dim resultMarginArray() As Variant
    Dim sectorNetPositionsArray() As Variant
    Dim sectorAbsPositionsArray() As Variant
    
    ' Size arrays based on data dimensions
    ReDim resultArray(1 To lastRow, 1 To (uniqueSymbols.count + 2))
    ReDim resultMarginArray(1 To lastRow, 1 To uniqueSymbols.count)
    ReDim sectorNetPositionsArray(1 To lastRow, 1 To uniqueSectors.count)
    ReDim sectorAbsPositionsArray(1 To lastRow, 1 To uniqueSectors.count)
    
    ' Set up headers
    Call SetupArrayHeaders(resultArray, resultMarginArray, sectorNetPositionsArray, sectorAbsPositionsArray, uniqueSymbols, uniqueSectors)
    
    ' Process data
    If Not ProcessMarginData(wsPortfolio, wsPortInMarketShort, wsPortInMarketLong, _
                           resultArray, resultMarginArray, sectorNetPositionsArray, sectorAbsPositionsArray, _
                           uniqueSymbols, uniqueSectors, startdate, currentdate, lastRow, lastCol, _
                           dataStartRow, dataEndRow) Then
        GoTo CleanExit
    End If
    
    ' Create summary statistics
    Call CreateSummaryStatistics(wsContractMarginTracking, wsTotalPortfolioM2M, resultArray, dataStartRow, dataEndRow)
    
    ' Create visualizations
    Call CreateVisualizations(wsContractMarginTracking, resultArray, resultMarginArray, _
                            sectorNetPositionsArray, sectorAbsPositionsArray, _
                            uniqueSymbols, uniqueSectors, dataStartRow, dataEndRow)
    
    ' Format worksheet
    Call FormatWorksheet(wsContractMarginTracking, uniqueSymbols.count, uniqueSectors.count)
    
    ' Add navigation buttons
    Call AddNavigationButtons(wsContractMarginTracking)
    
    ' Create pie charts
    
    Call CreatePieCharts(wsContractMarginTracking, uniqueSymbols.count, uniqueSectors.count, dataStartRow, dataEndRow)
    
    ' Final cleanup and organization
    Call OrderVisibleTabsBasedOnList
    wsContractMarginTracking.Activate
    
    wsContractMarginTracking.Columns("B:B").ColumnWidth = 15
    
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
    
    MsgBox "Contract Margin Tracking with sector summaries, graphs, and statistics created successfully!", vbInformation
    
CleanExit:
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "An error occurred:" & vbNewLine & _
            "Error " & Err.Number & ": " & Err.Description & vbNewLine & _
            "Line: " & Erl
    MsgBox errMsg, vbCritical
    Resume CleanExit
End Sub

Private Function ValidateRequiredWorksheets(ByRef wsPortfolio As Worksheet, ByRef wsPortfolioGraphs As Worksheet) As Boolean
    On Error Resume Next
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set wsPortfolioGraphs = ThisWorkbook.Sheets("PortfolioGraphs")
    On Error GoTo 0
    
     Call InitializeColumnConstantsManually
    
    If wsPortfolio Is Nothing Then
        MsgBox "Error: 'Portfolio' sheet does not exist.", vbExclamation
        ValidateRequiredWorksheets = False
        Exit Function
    End If
    
    If wsPortfolio.Cells(2, COL_PORT_STRATEGY_NAME).value = "" Then
        MsgBox "Error: 'Portfolio' sheet exists but contains no data in row 2.", vbExclamation
        ValidateRequiredWorksheets = False
        Exit Function
    End If
    
    If wsPortfolioGraphs Is Nothing Then
        MsgBox "Error: 'PortfolioGraphs' sheet does not exist.", vbExclamation
        ValidateRequiredWorksheets = False
        Exit Function
    End If
    
    ValidateRequiredWorksheets = True
End Function

Private Function CreateNewTrackingSheet(ByRef wsContractMarginTracking As Worksheet) As Boolean
    
    Call Deletetab("ContractMarginTracking")
    
    Set wsContractMarginTracking = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsContractMarginTracking.name = "ContractMarginTracking"
    wsContractMarginTracking.Tab.Color = RGB(117, 219, 255)
    
          ' Set white background color for the entire worksheet
    'wsContractMarginTracking.Cells.Interior.Color = RGB(255, 255, 255)
    
    CreateNewTrackingSheet = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function PopulateUniqueValues(ByRef wsPortfolio As Worksheet, ByRef wsPortInMarketShort As Worksheet, _
                                    ByRef uniqueSymbols As Object, ByRef uniqueSectors As Object, _
                                    ByVal lastCol As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long, portfolioRow As Long
    Dim strategyName As String, symbol As Variant, sector As String
    Dim lastRowPort As Long
    
    lastRowPort = wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row
    
    ' Create array for faster lookup
    Dim portfolioData As Variant
    portfolioData = wsPortfolio.Range("A2:Z" & lastRowPort).value
    
    For i = 2 To lastCol
        strategyName = wsPortInMarketShort.Cells(1, i).value
        If strategyName <> "" Then
            portfolioRow = FindPortfolioRow(portfolioData, strategyName)
            
            If portfolioRow = -99 Then
                MsgBox "Error: Cannot find " & strategyName & " in 'Portfolio' tab.", vbExclamation
                PopulateUniqueValues = False
                Exit Function
            End If
            
            symbol = portfolioData(portfolioRow, COL_PORT_SYMBOL)
            sector = portfolioData(portfolioRow, COL_PORT_SECTOR)
            
            If Not uniqueSymbols.Exists(symbol) Then uniqueSymbols.Add symbol, symbol
            If Not uniqueSectors.Exists(sector) Then uniqueSectors.Add sector, sector
        End If
    Next i
    
    PopulateUniqueValues = True
    Exit Function
    
ErrorHandler:
    PopulateUniqueValues = False
End Function

Private Sub SetupArrayHeaders(ByRef resultArray() As Variant, ByRef resultMarginArray() As Variant, _
                            ByRef sectorNetPositionsArray() As Variant, ByRef sectorAbsPositionsArray() As Variant, _
                            ByRef uniqueSymbols As Object, ByRef uniqueSectors As Object)
    
    ' Set headers in result array
    resultArray(1, 1) = "Date"
    resultArray(1, 2) = "Total Margin"
    
    Dim symbolCol As Long, symbol As Variant
    symbolCol = 3
    For Each symbol In uniqueSymbols.keys
        resultArray(1, symbolCol) = symbol
        symbolCol = symbolCol + 1
    Next symbol
    
    symbolCol = 1
    For Each symbol In uniqueSymbols.keys
        resultMarginArray(1, symbolCol) = symbol
        symbolCol = symbolCol + 1
    Next symbol
    
    Dim sectorCol As Long, sector As Variant
    sectorCol = 1
    For Each sector In uniqueSectors.keys
        sectorNetPositionsArray(1, sectorCol) = sector
        sectorCol = sectorCol + 1
    Next sector
    
    sectorCol = 1
    For Each sector In uniqueSectors.keys
        sectorAbsPositionsArray(1, sectorCol) = sector
        sectorCol = sectorCol + 1
    Next sector
End Sub

Private Function FindPortfolioRow(ByRef portfolioData As Variant, ByVal strategyName As String) As Long
    On Error GoTo ErrorHandler
    
    ' Initialize to error value
    FindPortfolioRow = -99
    
    ' Input validation
    If IsEmpty(portfolioData) Or Len(Trim(strategyName)) = 0 Then
        Exit Function
    End If
    
    Dim i As Long
    Dim currentStrategy As Variant
    
    ' Loop through portfolio data array
    For i = LBound(portfolioData, 1) To UBound(portfolioData, 1)
        currentStrategy = portfolioData(i, COL_PORT_STRATEGY_NAME)
        
        ' Check for matching strategy name
        If Not IsError(currentStrategy) Then
            If Trim(CStr(currentStrategy)) = Trim(strategyName) Then
                FindPortfolioRow = i
                Exit Function
            End If
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    FindPortfolioRow = -99
    Debug.Print "Error in FindPortfolioRow: " & Err.Description & " for strategy: " & strategyName
End Function


Private Function ProcessMarginData(ByRef wsPortfolio As Worksheet, ByRef wsPortInMarketShort As Worksheet, _
                                 ByRef wsPortInMarketLong As Worksheet, ByRef resultArray() As Variant, _
                                 ByRef resultMarginArray() As Variant, ByRef sectorNetPositionsArray() As Variant, _
                                 ByRef sectorAbsPositionsArray() As Variant, ByRef uniqueSymbols As Object, _
                                 ByRef uniqueSectors As Object, ByVal startdate As Date, ByVal currentdate As Date, _
                                 ByVal lastRow As Long, ByVal lastCol As Long, _
                                 ByVal dataStartRow As Long, ByRef dataEndRow As Long) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Create symbol to sector mapping dictionary for faster lookup
    Dim symbolSectorMap As Object
    Set symbolSectorMap = CreateObject("Scripting.Dictionary")
    
    ' Cache portfolio data
    Dim portfolioData As Variant
    portfolioData = wsPortfolio.Range("A2:Z" & wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row).value
    
    ' Build symbol to sector mapping
    Dim i As Long
    For i = LBound(portfolioData, 1) To UBound(portfolioData, 1)
        If Not IsEmpty(portfolioData(i, COL_PORT_SYMBOL)) Then
            symbolSectorMap(CStr(portfolioData(i, COL_PORT_SYMBOL))) = CStr(portfolioData(i, COL_PORT_SECTOR))
        End If
    Next i
    
    ' Cache market data
    Dim shortData As Variant, longData As Variant
    shortData = wsPortInMarketShort.Range("A1:ZZ" & lastRow).value
    longData = wsPortInMarketLong.Range("A1:ZZ" & lastRow).value
    
    Dim currentRow As Long
    currentRow = 2
    
    ' Initialize sector arrays for each date
    Dim sectorNetPositions As Object, sectorAbsPositions As Object
    
    For dateRow = 2 To lastRow
        Dim currentDateInData As Date
        currentDateInData = shortData(dateRow, 1)
        
        If Not IsNonTradingDay(currentDateInData) Then
            If currentDateInData >= startdate And currentDateInData <= currentdate Then
                ' Reset sector tracking for this date
                Set sectorNetPositions = CreateObject("Scripting.Dictionary")
                Set sectorAbsPositions = CreateObject("Scripting.Dictionary")
                
                ' Initialize sector values
                Dim sector As Variant
                For Each sector In uniqueSectors.keys
                    sectorNetPositions(sector) = 0
                    sectorAbsPositions(sector) = 0
                Next sector
                
                resultArray(currentRow, 1) = currentDateInData
                Dim totalMargin As Double
                totalMargin = 0
                
                Dim symbolCol As Long
                symbolCol = 3
                
                For Each symbol In uniqueSymbols.keys
                    Dim netPosition As Double
                    netPosition = CalculateNetPosition(symbol, dateRow, shortData, longData, portfolioData, lastCol)
                    
                    Dim margin As Double
                    margin = GetMarginValue(symbol, portfolioData)
                    
                    resultArray(currentRow, symbolCol) = netPosition
                    resultMarginArray(currentRow, symbolCol - 2) = Abs(netPosition) * margin
                    
                     ' Update sector totals
                    If symbolSectorMap.Exists(CStr(symbol)) Then
                        Dim symbolSector As String
                        symbolSector = symbolSectorMap(CStr(symbol))
                        
                        ' Update net positions (keeping true net value for now)
                        sectorNetPositions(symbolSector) = sectorNetPositions(symbolSector) + (netPosition * margin)
                        
                        ' Update absolute positions (using absolute value)
                        sectorAbsPositions(symbolSector) = sectorAbsPositions(symbolSector) + (Abs(netPosition) * margin)
                    End If
                    
                    If Not IsError(margin) And margin <> 0 Then
                        totalMargin = totalMargin + Abs(netPosition) * margin
                    End If
                    
                    symbolCol = symbolCol + 1
                Next symbol
                
                resultArray(currentRow, 2) = totalMargin
                
                ' Transfer sector data to arrays, making net positions absolute at this point
                Dim sectorCol As Long
                sectorCol = 1
                For Each sector In uniqueSectors.keys
                    ' Convert net positions to absolute values here
                    sectorNetPositionsArray(currentRow, sectorCol) = Abs(sectorNetPositions(sector))
                    sectorAbsPositionsArray(currentRow, sectorCol) = sectorAbsPositions(sector)
                    sectorCol = sectorCol + 1
                Next sector
                
                currentRow = currentRow + 1
                dataEndRow = dataEndRow + 1
            End If
        End If
        
        Application.StatusBar = "Contract Tracking Running: " & Format((currentDateInData - startdate) / (currentdate - startdate), "0%") & " completed"
        
       
    Next dateRow
    
    ProcessMarginData = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in ProcessMarginData: " & Err.Description
    ProcessMarginData = False
End Function



Private Sub UpdateSectorSummaries(ByVal symbol As Variant, ByVal netPosition As Double, _
                                ByVal margin As Double, ByVal currentRow As Long, _
                                ByRef sectorNetPositionsArray() As Variant, _
                                ByRef sectorAbsPositionsArray() As Variant, _
                                ByRef uniqueSectors As Object)
    On Error GoTo ErrorHandler
    
    ' Get sector for the current symbol
    Dim sector As String
    sector = FindSectorValueMargin(symbol)
    
    ' Find sector index in uniqueSectors collection
    Dim sectorIndex As Long
    sectorIndex = GetSectorIndex(sector, uniqueSectors)
    
    ' If sector is found, update arrays
    If sectorIndex > 0 Then
        ' Update net positions (considers direction of position)
        If Not IsEmpty(sectorNetPositionsArray(currentRow, sectorIndex)) Then
            sectorNetPositionsArray(currentRow, sectorIndex) = _
                sectorNetPositionsArray(currentRow, sectorIndex) + (netPosition * margin)
        Else
            sectorNetPositionsArray(currentRow, sectorIndex) = netPosition * margin
        End If
        
        ' Update absolute positions (ignores direction)
        If Not IsEmpty(sectorAbsPositionsArray(currentRow, sectorIndex)) Then
            sectorAbsPositionsArray(currentRow, sectorIndex) = _
                sectorAbsPositionsArray(currentRow, sectorIndex) + (Abs(netPosition) * margin)
        Else
            sectorAbsPositionsArray(currentRow, sectorIndex) = Abs(netPosition) * margin
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in UpdateSectorSummaries: " & Err.Description & _
                " for symbol: " & symbol & ", sector: " & sector
End Sub

Private Function GetSectorIndex(ByVal sector As String, ByRef uniqueSectors As Object) As Long
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim sectorKey As Variant
    i = 1
    
    ' Loop through unique sectors to find index
    For Each sectorKey In uniqueSectors.keys
        If sectorKey = sector Then
            GetSectorIndex = i
            Exit Function
        End If
        i = i + 1
    Next sectorKey
    
    GetSectorIndex = -1
    Exit Function
    
ErrorHandler:
    GetSectorIndex = -1
    Debug.Print "Error in GetSectorIndex: " & Err.Description & " for sector: " & sector
End Function

Private Function FindSectorValueMargin(ByVal symbol As String) As String
    On Error GoTo ErrorHandler
    
    ' Initialize portfolio sheet
    Dim wsPortfolio As Worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    
    ' Find the symbol in portfolio
    Dim lastRow As Long
    lastRow = wsPortfolio.Cells(wsPortfolio.rows.count, COL_PORT_SYMBOL).End(xlUp).row
    
    Dim findRow As Range
    Set findRow = wsPortfolio.Range("A2:A" & lastRow).Find(What:=symbol, _
                                                          LookIn:=xlValues, _
                                                          LookAt:=xlWhole, _
                                                          MatchCase:=True)
    
    ' If symbol found, return sector
    If Not findRow Is Nothing Then
        FindSectorValueMargin = wsPortfolio.Cells(findRow.row, COL_PORT_SECTOR).value
    Else
        FindSectorValueMargin = "Unknown"
    End If
    
    Exit Function
    
ErrorHandler:
    FindSectorValueMargin = "Unknown"
    Debug.Print "Error in FindSectorValueMargin: " & Err.Description & " for symbol: " & symbol
End Function


Private Function CalculateNetPosition(ByVal symbol As Variant, ByVal dateRow As Long, _
                                    ByRef shortData As Variant, ByRef longData As Variant, _
                                    ByRef portfolioData As Variant, ByVal lastCol As Long) As Double
    On Error GoTo ErrorHandler
    
    Dim netPosition As Double
    Dim i As Long
    Dim shortValue As Double, longValue As Double
    Dim strategyName As String
    netPosition = 0
    
    ' First pass: Handle cases where position is either long or short but not both
    For i = 2 To lastCol
        strategyName = shortData(1, i)
        If strategyName <> "" Then
            If SymbolMatchesStrategy(symbol, strategyName, portfolioData) Then
                ' Get short and long values from cached data
                shortValue = CDbl(shortData(dateRow, i))
                longValue = CDbl(longData(dateRow, i))
                
                ' Calculate net position for single direction positions
                If longValue > 0 And shortValue = 0 Then
                    netPosition = netPosition + longValue
                ElseIf shortValue > 0 And longValue = 0 Then
                    netPosition = netPosition - shortValue
                End If
            End If
        End If
    Next i
    
    ' Second pass: Handle cases where there are both long and short positions
    For i = 2 To lastCol
        strategyName = shortData(1, i)
        If strategyName <> "" Then
            If SymbolMatchesStrategy(symbol, strategyName, portfolioData) Then
                shortValue = CDbl(shortData(dateRow, i))
                longValue = CDbl(longData(dateRow, i))
                
                ' Handle mixed positions
                If longValue > 0 And shortValue > 0 Then
                    If netPosition < 0 Then
                        netPosition = netPosition - shortValue
                    Else
                        netPosition = netPosition + longValue
                    End If
                End If
            End If
        End If
    Next i
    
    CalculateNetPosition = netPosition
    Exit Function
    
ErrorHandler:
    CalculateNetPosition = 0
    Debug.Print "Error in CalculateNetPosition: " & Err.Description & " for symbol: " & symbol
End Function

Private Function GetMarginValue(ByVal symbol As Variant, ByRef portfolioData As Variant) As Double
    On Error GoTo ErrorHandler
    
    ' Initialize to error value
    GetMarginValue = 0
    
    ' Input validation
    If IsEmpty(portfolioData) Or IsNull(symbol) Then
        Exit Function
    End If
    
    Dim i As Long
    Dim currentSymbol As Variant
    Dim contracts As Double
    Dim margin As Double
    
    ' Loop through portfolio data array
    For i = LBound(portfolioData, 1) To UBound(portfolioData, 1)
        currentSymbol = portfolioData(i, COL_PORT_SYMBOL)
        
        ' Check for matching symbol
        If Not IsError(currentSymbol) Then
            If CStr(currentSymbol) = CStr(symbol) Then
                ' Get contracts and margin values
                contracts = CDbl(portfolioData(i, COL_PORT_CONTRACTS))
                margin = CDbl(portfolioData(i, COL_PORT_MARGIN))
                
                ' Calculate margin per contract
                If contracts > 0 Then
                    GetMarginValue = margin / contracts
                    Exit Function
                End If
            End If
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    GetMarginValue = 0
    Debug.Print "Error in GetMarginValue: " & Err.Description & " for symbol: " & symbol
End Function


Private Function SymbolMatchesStrategy(ByVal symbol As Variant, ByVal strategyName As String, _
                                     ByRef portfolioData As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    ' Search through portfolio data for matching strategy
    For i = LBound(portfolioData, 1) To UBound(portfolioData, 1)
        If CStr(portfolioData(i, COL_PORT_STRATEGY_NAME)) = strategyName Then
            ' Check if symbol matches
            SymbolMatchesStrategy = (CStr(portfolioData(i, COL_PORT_SYMBOL)) = CStr(symbol))
            Exit Function
        End If
    Next i
    
    SymbolMatchesStrategy = False
    Exit Function
    
ErrorHandler:
    SymbolMatchesStrategy = False
End Function



Private Sub CreateSummaryStatistics(ByRef wsSheet As Worksheet, ByRef wsTotalPortfolioM2M As Worksheet, _
                                  ByRef resultArray() As Variant, ByVal dataStartRow As Long, _
                                  ByVal dataEndRow As Long)
    On Error GoTo ErrorHandler
    
    ' First, ensure we have the margin data in the worksheet
    Dim marginDataColumn As Long
    marginDataColumn = 14  ' Column N
    
    ' Write the margin data from resultArray to the worksheet
    Dim i As Long
    For i = dataStartRow To dataEndRow
        wsSheet.Cells(i, marginDataColumn).value = resultArray(i - dataStartRow + 1, 2) ' Column 2 of resultArray contains Total Margin
    Next i
    
    ' Create Range object for margin calculations
    Dim marginRange As Range
    Set marginRange = wsSheet.Range(wsSheet.Cells(dataStartRow, marginDataColumn), _
                                  wsSheet.Cells(dataEndRow, marginDataColumn))
    
    ' Write summary statistics headers
    With wsSheet
        .Cells(1, 1).value = "Margin Summary Statistics"
        .Cells(2, 1).value = "Average Margin"
        .Cells(3, 1).value = "Median Margin"
        .Cells(4, 1).value = "Maximum Margin"
        .Cells(5, 1).value = "Average Margin + Max Drawdown"
        .Cells(6, 1).value = "Median Margin + Max DrawDown"
        .Cells(7, 1).value = "Maximum Margin + Max DrawDown"
        
        ' Calculate and write statistics
        .Cells(2, 2).value = WorksheetFunction.Average(marginRange)
        .Cells(3, 2).value = WorksheetFunction.Median(marginRange)
        .Cells(4, 2).value = WorksheetFunction.Max(marginRange)
        .Cells(5, 2).value = .Cells(2, 2).value + WorksheetFunction.Max(wsTotalPortfolioM2M.Range("D:D").value)
        .Cells(6, 2).value = .Cells(3, 2).value + WorksheetFunction.Max(wsTotalPortfolioM2M.Range("D:D").value)
        .Cells(7, 2).value = .Cells(4, 2).value + WorksheetFunction.Max(wsTotalPortfolioM2M.Range("D:D").value)
        
        ' Format as currency
        .Range("B2:B7").NumberFormat = "$#,##0"
        
    End With
    
    ' Debug print to verify data
    Debug.Print "Margin Range Count: " & marginRange.Cells.count
    Debug.Print "First Margin Value: " & marginRange.Cells(1).value
    Debug.Print "Last Margin Value: " & marginRange.Cells(marginRange.Cells.count).value
    
    ' Create histogram data
    CreateHistogram wsSheet, marginRange, dataStartRow
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in CreateSummaryStatistics: " & Err.Description
End Sub


Private Sub CreateHistogram(ByRef wsSheet As Worksheet, ByRef marginRange As Range, ByVal dataStartRow As Long)
    Dim minMargin As Double, maxMargin As Double, binWidth As Double
    Dim binMin As Double, binMax As Double, numBins As Integer
    
    minMargin = Application.WorksheetFunction.Min(marginRange)
    maxMargin = Application.WorksheetFunction.Max(marginRange)
    
    ' Calculate optimal bin width
    If Application.WorksheetFunction.Round(maxMargin / 10, -6) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxMargin / 10, -6)
    ElseIf Application.WorksheetFunction.Round(maxMargin / 10, -5) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxMargin / 10, -5)
    ElseIf Application.WorksheetFunction.Round(maxMargin / 10, -4) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxMargin / 10, -4)
    ElseIf Application.WorksheetFunction.Round(maxMargin / 10, -3) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxMargin / 10, -3)
    ElseIf Application.WorksheetFunction.Round(maxMargin / 10, -2) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxMargin / 10, -2)
    ElseIf Application.WorksheetFunction.Round(maxMargin / 10, -1) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxMargin / 10, -1)
    Else
        binWidth = 1
    End If
    
    binMin = Application.WorksheetFunction.Floor(minMargin, binWidth)
    binMax = Application.WorksheetFunction.Ceiling(maxMargin, binWidth)
    numBins = (binMax - binMin) / binWidth + 1
    
    ' Create histogram data
    wsSheet.Cells(10, 1).value = "Margin Histogram"
    wsSheet.Cells(11, 1).value = "Bin"
    wsSheet.Cells(11, 2).value = "Frequency"
    
    Dim binRow As Long, binLabel As Double
    Dim totalCount As Long
    
    ' Calculate total count first
    totalCount = marginRange.rows.count
    
    binRow = 12
    For binLabel = binMin To binMax Step binWidth
        wsSheet.Cells(binRow, 1).value = Format(binLabel, "$#,##0") & " - " & Format(binLabel + binWidth, "$#,##0")
        wsSheet.Cells(binRow, 2).value = Application.WorksheetFunction.CountIfs( _
        marginRange, ">=" & binLabel, marginRange, "<" & binLabel + binWidth) / totalCount
        binRow = binRow + 1
    Next binLabel
    
    ' Format histogram bins and percentages
    wsSheet.Range("A10:A" & binRow).NumberFormat = "$#,##0 - $#,##0"
    wsSheet.Range("B10:B" & binRow).NumberFormat = "0.0%"
End Sub

Private Sub CreateVisualizations(ByRef wsSheet As Worksheet, ByRef resultArray() As Variant, _
                               ByRef resultMarginArray() As Variant, _
                               ByRef sectorNetPositionsArray() As Variant, _
                               ByRef sectorAbsPositionsArray() As Variant, _
                               ByRef uniqueSymbols As Object, ByRef uniqueSectors As Object, _
                               ByVal dataStartRow As Long, ByVal dataEndRow As Long)
    On Error GoTo ErrorHandler
    
    ' Create Total Margin Line Chart
    CreateMarginLineChart wsSheet, dataStartRow, dataEndRow
    
    ' Create Margin Histogram Chart
    CreateHistogramChart wsSheet
    
    ' Write arrays to worksheet
    WriteArraysToWorksheet wsSheet, resultArray, resultMarginArray, _
                          sectorNetPositionsArray, sectorAbsPositionsArray, _
                          uniqueSymbols.count, uniqueSectors.count, _
                          dataStartRow, dataEndRow
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in CreateVisualizations: " & Err.Description
End Sub

Private Sub CreateMarginLineChart(ByRef wsSheet As Worksheet, ByVal dataStartRow As Long, ByVal dataEndRow As Long)
    Dim graphChart As ChartObject
    
    Set graphChart = wsSheet.ChartObjects.Add(left:=wsSheet.Cells(6, 1).left, _
                                            Width:=500, top:=wsSheet.Cells(10, 1).top, Height:=220)
    
    With graphChart.chart
        .SetSourceData source:=wsSheet.Range("M" & dataStartRow & ":N" & dataEndRow)
        .ChartType = xlLine
        .HasTitle = True
        .chartTitle.text = "Total Margin Over Time"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .HasLegend = False
        .Axes(xlCategory, xlPrimary).AxisTitle.text = "Date"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.text = "Total Margin"
        .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
        .SeriesCollection(1).Border.Weight = xlThin
    End With
End Sub

Private Sub CreateHistogramChart(ByRef wsSheet As Worksheet)
    Dim graphChart As ChartObject
    Dim lastRow As Long
    
    lastRow = wsSheet.Cells(wsSheet.rows.count, 1).End(xlUp).row
    
    Set graphChart = wsSheet.ChartObjects.Add(left:=wsSheet.Cells(20, 1).left, _
                                            Width:=500, top:=wsSheet.Cells(26, 1).top, Height:=220)
    
    With graphChart.chart
        .SetSourceData source:=wsSheet.Range("A10:B" & lastRow)
        .ChartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.text = "Total Margin Histogram"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .HasLegend = True
        .Axes(xlCategory, xlPrimary).AxisTitle.text = "Margin Range"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.text = "Frequency"
        .Axes(xlCategory).TickLabels.NumberFormat = "$#,##0"
    End With
End Sub

Private Sub WriteArraysToWorksheet(ByRef wsSheet As Worksheet, ByRef resultArray() As Variant, _
                                 ByRef resultMarginArray() As Variant, _
                                 ByRef sectorNetPositionsArray() As Variant, _
                                 ByRef sectorAbsPositionsArray() As Variant, _
                                 ByVal symbolCount As Long, ByVal sectorCount As Long, _
                                 ByVal dataStartRow As Long, ByVal dataEndRow As Long)
    
    ' Write date and total margin columns
    wsSheet.Range(wsSheet.Cells(dataStartRow, 13), _
                 wsSheet.Cells(dataStartRow + UBound(resultArray) - 1, 14)).value = _
                 GetColumnsFromArray(resultArray, 1, 2)  ' Date and Total Margin
    
    ' Write contract positions
    wsSheet.Range(wsSheet.Cells(dataStartRow, 15), _
                 wsSheet.Cells(dataStartRow + UBound(resultArray) - 1, 14 + symbolCount)).value = _
                 GetColumnsFromArray(resultArray, 3, symbolCount)
    
    ' Write gross margin by contract
    wsSheet.Range(wsSheet.Cells(dataStartRow, symbolCount + 16), _
                 wsSheet.Cells(dataStartRow + UBound(resultMarginArray) - 1, symbolCount * 2 + 15)).value = resultMarginArray
    
    ' Write net margin by sector with headers
    Dim netMarginRange As Range
    Set netMarginRange = wsSheet.Range(wsSheet.Cells(dataStartRow, symbolCount * 2 + 17), _
                 wsSheet.Cells(dataStartRow + UBound(sectorNetPositionsArray) - 1, symbolCount * 2 + sectorCount + 16))
    netMarginRange.value = sectorNetPositionsArray
    
    ' Write gross margin by sector with headers
    Dim grossMarginRange As Range
    Set grossMarginRange = wsSheet.Range(wsSheet.Cells(dataStartRow, symbolCount * 2 + sectorCount + 18), _
                 wsSheet.Cells(dataStartRow + UBound(sectorAbsPositionsArray) - 1, symbolCount * 2 + sectorCount * 2 + 17))
    grossMarginRange.value = sectorAbsPositionsArray
    
    ' Write headers
    With wsSheet
        .Cells(dataStartRow - 1, 13).value = "Date"
        .Cells(dataStartRow - 1, 14).value = "Total Margin"
        .Cells(dataStartRow - 1, 15).value = "Contracts"
        .Cells(dataStartRow - 1, symbolCount + 16).value = "Gross Margin for each Contract"
        .Cells(dataStartRow - 1, symbolCount * 2 + 17).value = "Net Margin Amounts by Sector"
        .Cells(dataStartRow - 1, symbolCount * 2 + sectorCount + 18).value = "Gross Margin Amounts by Sector"
        
        ' Write individual sector headers for net margins
        Dim i As Long
        For i = 1 To UBound(sectorNetPositionsArray, 2)
            .Cells(dataStartRow - 1, symbolCount * 2 + 16 + i).value = sectorNetPositionsArray(1, i)
        Next i
        
        ' Write individual sector headers for gross margins
        For i = 1 To UBound(sectorAbsPositionsArray, 2)
            .Cells(dataStartRow - 1, symbolCount * 2 + sectorCount + 17 + i).value = sectorAbsPositionsArray(1, i)
        Next i
        
        ' Format headers
        .Range(.Cells(dataStartRow - 1, 13), .Cells(dataStartRow - 1, symbolCount * 2 + sectorCount * 2 + 17)).Font.Bold = True
    End With
End Sub


Private Sub FormatWorksheet(ByRef wsSheet As Worksheet, ByVal symbolCount As Long, ByVal sectorCount As Long)
    On Error GoTo ErrorHandler
    
    ' Auto-fit the columns with statistics
    wsSheet.Columns("A:B").AutoFit
    
    ' Format margin columns
    Dim col As Long
    For col = 13 To 14 ' Columns containing profits
        wsSheet.Columns(col).ColumnWidth = 15
    Next col
    
    wsSheet.Columns(14).NumberFormat = "$#,##0"
    
    ' Format contract columns
    For col = 15 To (symbolCount + 15)
        wsSheet.Columns(col).ColumnWidth = 7
        wsSheet.Columns(col).NumberFormat = "0.0"
    Next col
    
    ' Format margin and sector columns
    For col = (symbolCount + 16) To (symbolCount * 2 + sectorCount * 2 + 18)
        wsSheet.Columns(col).ColumnWidth = 10
        wsSheet.Columns(col).NumberFormat = "$#,##0"
    Next col
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in FormatWorksheet: " & Err.Description
End Sub

Private Function GetColumnsFromArray(arr As Variant, startCol As Long, colCount As Long) As Variant
    Dim result() As Variant
    Dim i As Long, j As Long
    
    ReDim result(1 To UBound(arr), 1 To colCount)
    
    For i = 1 To UBound(arr)
        For j = 1 To colCount
            result(i, j) = arr(i, startCol + j - 1)
        Next j
    Next i
    
    GetColumnsFromArray = result
End Function

Private Sub CreateSectorAnalysisCharts(ByRef wsSheet As Worksheet, ByVal symbolCount As Long, _
                                     ByVal sectorCount As Long, ByVal dataStartRow As Long, _
                                     ByVal dataEndRow As Long)
    ' Create Net Margin by Sector Chart
    CreateSectorChart wsSheet, "Net Margin by Sector Over Time", _
                     symbolCount * 2 + 17, symbolCount * 2 + sectorCount + 16, _
                     dataStartRow, dataEndRow, 40, 1
    
    ' Create Gross Margin by Sector Chart
    CreateSectorChart wsSheet, "Gross Margin by Sector Over Time", _
                     symbolCount * 2 + sectorCount + 18, symbolCount * 2 + sectorCount * 2 + 17, _
                     dataStartRow, dataEndRow, 40, 8
End Sub

Private Sub CreateSectorChart(ByRef wsSheet As Worksheet, ByVal chartTitle As String, _
                            ByVal startCol As Long, ByVal endCol As Long, _
                            ByVal dataStartRow As Long, ByVal dataEndRow As Long, _
                            ByVal chartRow As Long, ByVal chartCol As Long)
    Dim chartObj As ChartObject
    Set chartObj = wsSheet.ChartObjects.Add(left:=wsSheet.Cells(chartRow, chartCol).left, _
                                          Width:=600, Height:=300)
    
    With chartObj.chart
        ' Set source data including dates
        .SetSourceData source:=wsSheet.Range( _
            wsSheet.Cells(dataStartRow - 1, 13), _
            wsSheet.Cells(dataEndRow, endCol))
        
        ' Set chart type and format
        .ChartType = xlLine
        .HasTitle = True
        .chartTitle.text = chartTitle
        
        ' Format axes
        With .Axes(xlCategory, xlPrimary)
            .CategoryType = xlTimeScale
            .TickLabels.NumberFormat = "mm/dd/yyyy"
            .HasTitle = True
            .AxisTitle.text = "Date"
        End With
        
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.text = "Margin Amount"
            .NumberFormat = "$#,##0"
        End With
        
        ' Add and position legend
        .HasLegend = True
        .Legend.position = xlLegendPositionRight
        
        ' Format series
        Dim series As series
        For Each series In .SeriesCollection
            With series
                .MarkerStyle = xlMarkerStyleNone
                .Format.line.Weight = 2
            End With
        Next series
    End With
End Sub


Private Sub AddSectorSummaryTables(ByRef wsSheet As Worksheet, ByVal symbolCount As Long, _
                                 ByVal sectorCount As Long, ByVal dataStartRow As Long, _
                                 ByVal dataEndRow As Long)
    ' Calculate latest date summary
    Dim latestRow As Long
    latestRow = dataEndRow
    
    ' Write Net Margin Summary
    WriteSectorSummary wsSheet, "Net Margin by Sector Summary", _
                      symbolCount * 2 + 17, symbolCount * 2 + sectorCount + 16, _
                      dataStartRow, dataEndRow, 20, 1
    
    ' Write Gross Margin Summary
    WriteSectorSummary wsSheet, "Gross Margin by Sector Summary", _
                      symbolCount * 2 + sectorCount + 18, symbolCount * 2 + sectorCount * 2 + 17, _
                      dataStartRow, dataEndRow, 20, 8
End Sub

Private Sub WriteSectorSummary(ByRef wsSheet As Worksheet, ByVal title As String, _
                             ByVal startCol As Long, ByVal endCol As Long, _
                             ByVal dataStartRow As Long, ByVal dataEndRow As Long, _
                             ByVal summaryRow As Long, ByVal summaryCol As Long)
    With wsSheet
        ' Write title
        .Cells(summaryRow, summaryCol).value = title
        .Cells(summaryRow, summaryCol).Font.Bold = True
        
        ' Write headers
        .Cells(summaryRow + 1, summaryCol).value = "Sector"
        .Cells(summaryRow + 1, summaryCol + 1).value = "Latest Value"
        .Cells(summaryRow + 1, summaryCol + 2).value = "Average"
        .Cells(summaryRow + 1, summaryCol + 3).value = "Maximum"
        .Range(.Cells(summaryRow + 1, summaryCol), .Cells(summaryRow + 1, summaryCol + 3)).Font.Bold = True
        
        ' Write sector data
        Dim col As Long, row As Long
        row = summaryRow + 2
        
        For col = startCol To endCol
            .Cells(row, summaryCol).value = .Cells(dataStartRow - 1, col).value
            .Cells(row, summaryCol + 1).Formula = "=" & .Cells(dataEndRow, col).Address
            .Cells(row, summaryCol + 2).Formula = "=AVERAGE(" & .Range(.Cells(dataStartRow, col), .Cells(dataEndRow, col)).Address & ")"
            .Cells(row, summaryCol + 3).Formula = "=MAX(" & .Range(.Cells(dataStartRow, col), .Cells(dataEndRow, col)).Address & ")"
            row = row + 1
        Next col
        
        ' Format numbers
        .Range(.Cells(summaryRow + 2, summaryCol + 1), .Cells(row - 1, summaryCol + 3)).NumberFormat = "$#,##0"
        
        ' AutoFit columns
        .Range(.Cells(summaryRow, summaryCol), .Cells(row - 1, summaryCol + 3)).Columns.AutoFit
    End With
End Sub

Private Sub AddNavigationButtons(ByRef wsSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    
    ' Delete Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 4).left, _
                                 top:=wsSheet.Cells(2, 4).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteContractMarginTracking"
    End With
    
    ' Summary Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(4, 4).left, _
                                 top:=wsSheet.Cells(4, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With
    
    ' Portfolio Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(4, 7).left, _
                                 top:=wsSheet.Cells(4, 4).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With
    
    ' Control Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 7).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With
    
    ' Strategies Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(4, 10).left, _
                                 top:=wsSheet.Cells(4, 4).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies"
    End With
    
    ' Inputs Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 10).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs"
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddNavigationButtons: " & Err.Description
End Sub

Private Sub CreatePieCharts(ByRef wsSheet As Worksheet, ByVal symbolCount As Long, _
                          ByVal sectorCount As Long, ByVal dataStartRow As Long, _
                          ByVal dataEndRow As Long)
    On Error GoTo ErrorHandler
    
    ' Calculate dates for filtering
    Dim currentdate As Date
    currentdate = wsSheet.Cells(dataEndRow, 13).value
    Dim lastYearStartDate As Date
    lastYearStartDate = DateAdd("yyyy", -1, currentdate)
    
    ' Define row for totals
    Dim totalRow As Long
    totalRow = 2
    
    ' Calculate and write totals
    CalculateAndWriteTotals wsSheet, symbolCount, sectorCount, _
                           dataStartRow, dataEndRow, totalRow, lastYearStartDate
    
    ' Create pie charts for different data sets
    CreatePieChartSet wsSheet, symbolCount, "Gross Margin for each Contract", _
                     totalRow, symbolCount + 16, symbolCount * 2 + 15, 42
                     
    CreatePieChartSet wsSheet, sectorCount, "Net Margin Amounts by Sector", _
                     totalRow, symbolCount * 2 + 17, symbolCount * 2 + sectorCount + 16, 60
                     
    CreatePieChartSet wsSheet, sectorCount, "Gross Margin Amounts by Sector", _
                     totalRow, symbolCount * 2 + sectorCount + 18, symbolCount * 2 + sectorCount * 2 + 17, 78
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in CreatePieCharts: " & Err.Description
End Sub

Private Sub CalculateAndWriteTotals(ByRef wsSheet As Worksheet, ByVal symbolCount As Long, _
                                  ByVal sectorCount As Long, ByVal dataStartRow As Long, _
                                  ByVal dataEndRow As Long, ByVal totalRow As Long, _
                                  ByVal lastYearStartDate As Date)
    Dim startCol As Long, endCol As Long
    Dim j As Long, i As Long
    Dim totalSums As Double, lastYearSums As Double
    
    startCol = symbolCount + 16
    endCol = symbolCount * 2 + sectorCount * 2 + 17
    
    For j = startCol To endCol
        lastYearSums = 0
        If wsSheet.Cells(dataStartRow, j).value <> "" Then
            ' Write headers
            wsSheet.Cells(totalRow - 1, j).value = wsSheet.Cells(dataStartRow, j).value
            wsSheet.Cells(totalRow + 1, j).value = wsSheet.Cells(dataStartRow, j).value
            
            ' Calculate total sums
            totalSums = Application.sum(wsSheet.Range(wsSheet.Cells(dataStartRow + 1, j), _
                                                    wsSheet.Cells(dataEndRow, j)))
            
            ' Calculate last year sums
            For i = dataStartRow + 1 To dataEndRow
                If wsSheet.Cells(i, 13).value >= lastYearStartDate Then
                    lastYearSums = lastYearSums + wsSheet.Cells(i, j).value
                End If
            Next i
            
            ' Write totals
            wsSheet.Cells(totalRow, j).value = totalSums
            wsSheet.Cells(totalRow + 2, j).value = lastYearSums
        End If
    Next j
End Sub

Private Sub CreatePieChartSet(ByRef wsSheet As Worksheet, ByVal count As Long, _
                            ByVal titleBase As String, ByVal totalRow As Long, _
                            ByVal startCol As Long, ByVal endCol As Long, _
                            ByVal rowPosition As Long)
    ' Create Total Period Chart
    Dim RangeTotal As Range
    Set RangeTotal = wsSheet.Range(wsSheet.Cells(totalRow - 1, startCol), _
                                  wsSheet.Cells(totalRow, endCol))
    
    CreateSinglePieChart wsSheet, RangeTotal, titleBase & " - Total", _
                        rowPosition, 1
    
    ' Create Last Year Chart
    Dim RangeLastYear As Range
    Set RangeLastYear = wsSheet.Range(wsSheet.Cells(totalRow + 1, startCol), _
                                    wsSheet.Cells(totalRow + 2, endCol))
    
    CreateSinglePieChart wsSheet, RangeLastYear, titleBase & " - Last Year", _
                        rowPosition, 6
End Sub

Private Sub CreateSinglePieChart(ByRef wsSheet As Worksheet, ByRef sourceRange As Range, _
                               ByVal title As String, ByVal row As Long, ByVal col As Long)
    Dim chartObj As ChartObject
    Set chartObj = wsSheet.ChartObjects.Add(left:=wsSheet.Cells(row, col).left, _
                                          top:=wsSheet.Cells(row, col).top, _
                                          Width:=300, Height:=260)
    
    With chartObj.chart
        .ChartType = xlPie
        .SetSourceData source:=sourceRange
        .HasTitle = True
        .chartTitle.text = title
        .HasLegend = True
        .Legend.position = xlLegendPositionRight
        
        With .SeriesCollection(1)
            .ApplyDataLabels
            .DataLabels.ShowPercentage = True
            .DataLabels.ShowCategoryName = True
            .DataLabels.ShowValue = False
        End With
    End With
End Sub





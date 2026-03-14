Attribute VB_Name = "S_LeaveOneOut"
Sub RunLeaveOneOutAnalysis()
    Dim wsPortfolio    As Worksheet
    Dim wsLeaveOneOut  As Worksheet
    Dim currentdate    As Date, startdate As Date, endDate As Date
    Dim yearsToConsider As Double
    Dim MCTradeType    As String
    Dim ceaseType      As String
    Dim requiredMargin As Double, margin As Double
    Dim startingEquity As Double
    Dim numScenarios   As Long
    Dim tradeAdjustment As Double
    Dim pnlResults     As Variant
    Dim averageTradesPerYear As Long
    Dim AverageTrade   As Double
    Dim numStrategies  As Long
    Dim i As Long, j As Long, K As Long
    Dim TradeCount As Long
    
    ' NEW: User selection variables
    Dim analysisType As String
    Dim sortingMetric As String
    
    '— 0) License & setup —
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license.", vbCritical
        Exit Sub
    End If
    Call InitializeColumnConstantsManually
    
    ' NEW: Ask user for analysis type
    analysisType = InputBox("Choose analysis type:" & vbCrLf & _
                           "1 = Monte Carlo (randomized scenarios)" & vbCrLf & _
                           "2 = Chronological Backtest (actual sequence)", _
                           "Leave-One-Out Analysis Type", "1")
    
    ' Check if user pressed Cancel
    If analysisType = "" Then Exit Sub
    
    If analysisType <> "1" And analysisType <> "2" Then
        MsgBox "Invalid selection. Using Monte Carlo method.", vbInformation
        analysisType = "1"
    End If
    
    ' NEW: Ask user for sorting metric
    sortingMetric = InputBox("Choose sorting metric:" & vbCrLf & _
                            "1 = Return/Max DD Impact (default)" & vbCrLf & _
                            "2 = Return/Avg DD Impact" & vbCrLf & _
                            "3 = Profit/Stdev Benefit" & vbCrLf & _
                            "4 = Strategy Number (original order)" & vbCrLf & _
                            "5 = Return Impact", _
                            "Sorting Metric", "1")
    
    ' Check if user pressed Cancel
    If sortingMetric = "" Then Exit Sub
    
    If sortingMetric <> "1" And sortingMetric <> "2" And sortingMetric <> "3" And sortingMetric <> "4" And sortingMetric <> "5" Then
        MsgBox "Invalid selection. Using Return/Max DD Impact.", vbInformation
        sortingMetric = "1"
    End If
    
    '— 1) Get Portfolio sheet & inputs —
    On Error Resume Next
      Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    On Error GoTo 0
    If wsPortfolio Is Nothing Then
        MsgBox "'Portfolio' sheet is missing!", vbExclamation: Exit Sub
    End If
    
    yearsToConsider = GetNamedRangeValue("PortfolioPeriod")
    currentdate = wsPortfolio.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    startdate = DateAdd("yyyy", -Int(yearsToConsider), currentdate)
    startdate = DateAdd("m", -(yearsToConsider - Int(yearsToConsider)) * 12, startdate)
    endDate = currentdate
    
    MCTradeType = GetNamedRangeValue("PortMCTradeType")
    ceaseType = GetNamedRangeValue("PortfolioCeaseTradingType")
    requiredMargin = GetNamedRangeValue("PortfolioCeaseTrading")
    startingEquity = GetNamedRangeValue("PortfolioStartingEquity")
    
    ' Only get Monte Carlo specific parameters if needed
    If analysisType = "1" Then
        numScenarios = GetNamedRangeValue("PortfolioSimulations")
        tradeAdjustment = GetNamedRangeValue("PortfolioMCTradeAdjustment")
    End If
    
    ' Compute actual margin threshold
    If LCase(ceaseType) = "percentage" Then
        margin = (1 - requiredMargin) * startingEquity
    Else
        margin = requiredMargin
    End If
    
    '— 2) Build PnL array (daily or weekly) —
    If LCase(MCTradeType) = "daily" Then
        pnlResults = CleanPortfolioDailyPnL(startdate, endDate)
        averageTradesPerYear = 252
    ElseIf LCase(MCTradeType) = "weekly" Then
        pnlResults = ConvertDailyToWeeklyPnL(startdate, endDate)
        averageTradesPerYear = 52
    Else
        MsgBox "MCTradeType must be 'Daily' or 'Weekly'.", vbExclamation
        Exit Sub
    End If
    
    If IsEmpty(pnlResults) Then
        MsgBox "No PnL data returned.", vbExclamation
        Exit Sub
    End If
    
    '— 3) Compute AverageTrade —
    AverageTrade = 0
    numStrategies = UBound(pnlResults, 2)
    TradeCount = 0
    
    For i = 1 To UBound(pnlResults, 1)
        Dim daySum As Double: daySum = 0
        For j = 1 To numStrategies
            daySum = daySum + pnlResults(i, j)
        Next j
        AverageTrade = AverageTrade + daySum
        TradeCount = TradeCount + 1
    Next i
    If TradeCount > 0 Then AverageTrade = AverageTrade / TradeCount
    
  
    
    On Error Resume Next
      Application.DisplayAlerts = False
      ThisWorkbook.Sheets("LeaveOneOut").Delete
      Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsLeaveOneOut = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsLeaveOneOut.name = "LeaveOneOut"
    wsLeaveOneOut.Tab.Color = RGB(117, 219, 255)
    wsLeaveOneOut.Cells.Interior.Color = RGB(255, 255, 255)
    
    '— 5) Run the selected analysis —
    If analysisType = "1" Then
        ' Monte Carlo Analysis
        Call RankStrategiesLeaveOneOut( _
            wsLeaveOneOut, _
            pnlResults, _
            margin, _
            averageTradesPerYear, _
            startingEquity, _
            numScenarios, _
            tradeAdjustment, _
            AverageTrade, _
            MCTradeType, _
            sortingMetric _
        )
    Else
        ' Chronological Backtest Analysis
        Call RankStrategiesLeaveOneOutBacktest( _
            wsLeaveOneOut, _
            pnlResults, _
            margin, _
            startingEquity, _
            AverageTrade, _
            MCTradeType, _
            sortingMetric _
        )
    End If
    
    '— 6) Add navigation buttons —
    AddNavigationButtonsLeaveOneOut wsLeaveOneOut
    
    ' Set window properties
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
    
    '— 7) Reorder tabs & finish up —
    Call OrderVisibleTabsBasedOnList
    Application.ScreenUpdating = True
    wsLeaveOneOut.Activate
    
    Dim analysisTypeName As String
    If analysisType = "1" Then analysisTypeName = "Monte Carlo" Else analysisTypeName = "Chronological Backtest"
    MsgBox analysisTypeName & " Leave-One-Out analysis complete!", vbInformation
End Sub


Function CalculateMetricsFromResults(results As Variant, startingEquity As Double) As Variant
    ' Calculates key metrics from Monte Carlo simulation results
    ' Returns array with: median return, median return/drawdown, median return/avg drawdown, risk of ruin
    
    Dim metrics(1 To 4) As Double
    Dim i As Long, ruinCount As Long
    
    ' Extract arrays for calculations
    Dim returnArray() As Double
    Dim returnToMaxDDArray() As Double
    Dim returnToAvgDDArray() As Double
    Dim ruinArray() As Double
    
    ReDim returnArray(1 To UBound(results, 1))
    ReDim returnToMaxDDArray(1 To UBound(results, 1))
    ReDim ruinArray(1 To UBound(results, 1))
    
    ' For average drawdown, we'll use a placeholder calculation if it's not available
    Dim hasAvgDrawdown As Boolean
    hasAvgDrawdown = (UBound(results, 2) >= 7)
    
    If hasAvgDrawdown Then
        ReDim returnToAvgDDArray(1 To UBound(results, 1))
    End If
    
    ' Calculate array values and count ruined scenarios
    ruinCount = 0
    For i = 1 To UBound(results, 1)
        returnArray(i) = results(i, 3)  ' Return percentage
        returnToMaxDDArray(i) = results(i, 4)  ' Return/Max Drawdown
        
        If hasAvgDrawdown Then
            returnToAvgDDArray(i) = results(i, 2) / (results(i, 7) + 0.000001)  ' Return/Avg Drawdown
        End If
        
        If results(i, 6) = 1 Then
            ruinCount = ruinCount + 1
        End If
    Next i
    
    ' Calculate median values
    metrics(1) = Application.WorksheetFunction.Median(returnArray)  ' Median return
    metrics(2) = Application.WorksheetFunction.Median(returnToMaxDDArray)  ' Median return/max drawdown
    
    If hasAvgDrawdown Then
        metrics(3) = Application.WorksheetFunction.Median(returnToAvgDDArray)  ' Median return/avg drawdown
    Else
        ' If avg drawdown not available, we'll use a factor of max drawdown as approximation
        metrics(3) = metrics(2) * 1.5  ' Assumption: avg DD typically lower than max DD
    End If
    
    ' Calculate risk of ruin
    metrics(4) = ruinCount / UBound(results, 1)
    
    CalculateMetricsFromResults = metrics
End Function



Function RankStrategiesByMetric(metrics As Variant, metricColumn As Long) As Long()
    ' Returns an array of rankings (1 = best) based on the specified metric column
    ' Higher metric values are considered better
    
    Dim numStrategies As Long, i As Long, j As Long
    Dim temp As Long
    Dim sortedIndices() As Long
    Dim ranks() As Long
    
    numStrategies = UBound(metrics, 1)
    
    ' Initialize arrays
    ReDim sortedIndices(1 To numStrategies)
    ReDim ranks(1 To numStrategies)
    
    ' Initialize indices
    For i = 1 To numStrategies
        sortedIndices(i) = i
    Next i
    
    ' Simple bubble sort by metric value (descending order - higher is better)
    For i = 1 To numStrategies - 1
        For j = i + 1 To numStrategies
            If metrics(sortedIndices(i), metricColumn) < metrics(sortedIndices(j), metricColumn) Then
                temp = sortedIndices(i)
                sortedIndices(i) = sortedIndices(j)
                sortedIndices(j) = temp
            End If
        Next j
    Next i
    
    ' Assign ranks based on sorted order
    For i = 1 To numStrategies
        ranks(sortedIndices(i)) = i
    Next i
    
    RankStrategiesByMetric = ranks
End Function

Sub RankStrategiesLeaveOneOut(wsPortfolioMC As Worksheet, pnlResults As Variant, requiredMargin As Double, _
                              averageTradesPerYear As Long, startingEquity As Double, _
                              numScenarios As Long, tradeAdjustment As Double, _
                              AverageTrade As Double, MCTradeType As String, _
                              sortingMetric As String)
    ' Performs "leave-one-out" sensitivity analysis on strategies
    ' Creates a ranking table and radar chart of strategy contributions
    ' Supports both Daily and Weekly multi-strategy data
    
    Dim baselineResults As Variant
    Dim dailyProfitTracking() As Double
    Dim dailyDrawdownTracking() As Double
    Dim dailyMaxDrawdownTracking() As Double
    Dim baselineMetrics As Variant
    Dim numStrategies As Long, i As Long, j As Long, K As Long
    Dim strategyNames() As String
    Dim Symbols() As String
    Dim strategyNumbers() As Long
    Dim metrics() As Variant
    Dim tableStartRow As Long, tableStartCol As Long
    Dim chartStartRow As Long, chartStartCol As Long
    Dim gRandIdx As Variant
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing baseline portfolio performance..."
    
    ' Define start positions
    tableStartRow = 6
    tableStartCol = 12  ' Column L
    chartStartRow = 6
    chartStartCol = 2   ' Column B
    
    ' Determine if we're using multi-strategy data (either daily or weekly)
    If MCTradeType <> "Daily" And MCTradeType <> "Weekly" Then
        MsgBox "Strategy contribution analysis requires daily or weekly multi-strategy data.", vbExclamation
        Exit Sub
    End If
    
    ' Determine number of strategies
    numStrategies = UBound(pnlResults, 2)
    If numStrategies < 2 Then
        MsgBox "At least two strategies are required for comparative analysis.", vbExclamation
        Exit Sub
    End If
    
    ' Get strategy names and numbers from the Portfolio sheet
    ReDim strategyNames(1 To numStrategies)
    ReDim strategyNumbers(1 To numStrategies)
    ReDim Symbols(1 To numStrategies)
    On Error Resume Next
    For i = 1 To numStrategies
        strategyNames(i) = ThisWorkbook.Sheets("Portfolio").Cells(i + 1, COL_PORT_STRATEGY_NAME).value
        strategyNumbers(i) = ThisWorkbook.Sheets("Portfolio").Cells(i + 1, COL_PORT_STRATEGY_NUMBER).value
        Symbols(i) = ThisWorkbook.Sheets("Portfolio").Cells(i + 1, COL_PORT_SYMBOL).value
        If strategyNames(i) = "" Then strategyNames(i) = "Strategy " & i
        If strategyNumbers(i) = 0 Then strategyNumbers(i) = i
        If Symbols(i) = "" Then Symbols(i) = "Error"
    Next i
    On Error GoTo 0
    
    
    gRandIdx = GenerateRandomIndexMatrix(pnlResults, numScenarios, averageTradesPerYear)
    
    ' Run baseline simulation with all strategies
    baselineResults = RunMonteCarloWithTracking(pnlResults, requiredMargin, _
                                              averageTradesPerYear, startingEquity, _
                                              numScenarios, tradeAdjustment, AverageTrade, _
                                              MCTradeType, dailyProfitTracking, _
                                              dailyDrawdownTracking, dailyMaxDrawdownTracking, gRandIdx)
    
    ' Calculate baseline metrics - store each component separately using DOLLAR amounts
    Dim baselineReturn As Double
    Dim baselineMaxDD As Double
    Dim baselineAvgDD As Double
    Dim baselineReturnToMaxDD As Double
    Dim baselineReturnToAvgDD As Double
    
    ' Calculate direct metrics from raw results - ALL IN DOLLARS
    baselineReturn = CalculateMedianReturn(baselineResults)  ' Now returns dollars
    baselineMaxDD = CalculateMedianMaxDrawdown(baselineResults)  ' Dollars
    baselineAvgDD = CalculateMedianAvgDrawdown(baselineResults)  ' Dollars
    
    ' Calculate the ratios - NOW BOTH ARE IN DOLLARS
    If baselineMaxDD > 0.01 Then ' Small threshold for division
        baselineReturnToMaxDD = baselineReturn / baselineMaxDD
    Else
        baselineReturnToMaxDD = baselineReturn * 1000  ' Large number if no meaningful drawdown
    End If
    
    If baselineAvgDD > 0.01 Then ' Small threshold for division
        baselineReturnToAvgDD = baselineReturn / baselineAvgDD
    Else
        baselineReturnToAvgDD = baselineReturn * 1000  ' Large number if no meaningful drawdown
    End If
    
    ' Create results array for each strategy's impact - EXPANDED to include return impact
    ReDim metrics(1 To numStrategies, 1 To 8)  ' Added one more column for return impact
    
    ' Run simulations excluding one strategy at a time
    For i = 1 To numStrategies
        Application.StatusBar = "Analyzing without strategy " & i & " of " & numStrategies & ": " & strategyNames(i)
        
        ' Create a modified version of pnlResults without the current strategy
        Dim modifiedPnlResults As Variant
        modifiedPnlResults = ExcludeStrategyFromPnL(pnlResults, i)
        
        ' Run Monte Carlo with this strategy excluded
        Dim excludedResults As Variant
        excludedResults = RunMonteCarloWithTracking(modifiedPnlResults, requiredMargin, _
                                                  averageTradesPerYear, startingEquity, _
                                                  numScenarios, tradeAdjustment, AverageTrade, _
                                                  MCTradeType, dailyProfitTracking, _
                                                  dailyDrawdownTracking, dailyMaxDrawdownTracking, gRandIdx)
        
        ' Calculate excluded metrics directly - ALL IN DOLLARS
        Dim excludedReturn As Double
        Dim excludedMaxDD As Double
        Dim excludedAvgDD As Double
        Dim excludedReturnToMaxDD As Double
        Dim excludedReturnToAvgDD As Double
        
        excludedReturn = CalculateMedianReturn(excludedResults)  ' Now returns dollars
        excludedMaxDD = CalculateMedianMaxDrawdown(excludedResults)  ' Dollars
        excludedAvgDD = CalculateMedianAvgDrawdown(excludedResults)  ' Dollars
        
        ' Calculate the ratios - NOW BOTH ARE IN DOLLARS
        If excludedMaxDD > 0.01 Then
            excludedReturnToMaxDD = excludedReturn / excludedMaxDD
        Else
            excludedReturnToMaxDD = excludedReturn * 1000
        End If
        
        If excludedAvgDD > 0.01 Then
            excludedReturnToAvgDD = excludedReturn / excludedAvgDD
        Else
            excludedReturnToAvgDD = excludedReturn * 1000
        End If
        
        ' Store the difference between baseline and excluded metrics
        metrics(i, 1) = i  ' Strategy index
        metrics(i, 2) = strategyNumbers(i)  ' Strategy number
        metrics(i, 3) = strategyNames(i)  ' Strategy name
        metrics(i, 8) = Symbols(i)  ' Symbol (moved to column 8)
        
        ' NEW: Calculate Return Impact (column 4) - Using percentage change logic as requested
        If Abs(baselineReturn) > 0.00001 Then
            metrics(i, 4) = (baselineReturn - excludedReturn) / baselineReturn
        Else
            metrics(i, 4) = 0
        End If
        
        ' Calculate Profit/Stdev Benefit (moved to column 5)
        Dim baselineStdev As Double, excludedStdev As Double
        Dim baselineProfitStdev As Double, excludedProfitStdev As Double
        
        ' Calculate standard deviation directly from PnL data for consistency
        baselineStdev = CalculatePortfolioStdevFromPnL(pnlResults)
        excludedStdev = CalculatePortfolioStdevFromPnL(modifiedPnlResults)
        
        ' Calculate Profit/Stdev ratios
        If baselineStdev > 0 Then baselineProfitStdev = baselineReturn / baselineStdev Else baselineProfitStdev = baselineReturn
        If excludedStdev > 0 Then excludedProfitStdev = excludedReturn / excludedStdev Else excludedProfitStdev = excludedReturn
        
        ' Store Profit/Stdev Benefit (moved to column 5)
        If Abs(excludedProfitStdev) > 0.00001 Then
            metrics(i, 5) = (baselineProfitStdev - excludedProfitStdev) / Abs(excludedProfitStdev)
        Else
            metrics(i, 5) = 0
        End If
        
        ' Impact on Return/Max DD as percentage change (moved to column 6)
        If Abs(baselineReturnToMaxDD) > 0.00001 Then
            metrics(i, 6) = (baselineReturnToMaxDD - excludedReturnToMaxDD) / baselineReturnToMaxDD
        Else
            metrics(i, 6) = 0
        End If
        
        ' Impact on Return/Avg DD as percentage change (moved to column 7)
        If Abs(baselineReturnToAvgDD) > 0.00001 Then
            metrics(i, 7) = (baselineReturnToAvgDD - excludedReturnToAvgDD) / baselineReturnToAvgDD
        Else
            metrics(i, 7) = 0
        End If
    Next i
    
    ' Sort metrics by selected criterion (updated for new column structure)
    Call SortLeaveOneOutMetricsWithReturnImpact(metrics, sortingMetric, numStrategies)
    
    ' Calculate ranks for each metric (updated column references)
    Dim rankReturnImpact() As Long
    Dim rankReturn() As Long
    Dim rankReturnMaxDD() As Long
    Dim rankReturnAvgDD() As Long
    
    rankReturnImpact = RankStrategiesByMetric(metrics, 4)  ' Rank by Return Impact
    rankReturn = RankStrategiesByMetric(metrics, 5)  ' Rank by Profit/Stdev benefit
    rankReturnMaxDD = RankStrategiesByMetric(metrics, 6)  ' Rank by return/max drawdown impact
    rankReturnAvgDD = RankStrategiesByMetric(metrics, 7)  ' Rank by return/avg drawdown impact
    
    ' Get sorting description
    Dim sortDescription As String
    Select Case sortingMetric
        Case "1": sortDescription = "Return/Max DD Impact"
        Case "2": sortDescription = "Return/Avg DD Impact"
        Case "3": sortDescription = "Profit/Stdev Benefit"
        Case "4": sortDescription = "Strategy Number"
        Case "5": sortDescription = "Return Impact"  ' New option
        Case Else: sortDescription = "Return/Max DD Impact"
    End Select
    
    ' Output results to worksheet
    With wsPortfolioMC
        ' Title for strategy analysis section
        .Cells(tableStartRow - 1, tableStartCol).value = "Strategy Contribution Analysis (" & MCTradeType & ") - Sorted by " & sortDescription
        .Cells(tableStartRow - 1, tableStartCol).Font.Bold = True
        .Cells(tableStartRow - 1, tableStartCol).Font.Size = 14
        .Range(.Cells(tableStartRow - 1, tableStartCol), .Cells(tableStartRow - 1, tableStartCol + 10)).Merge  ' Expanded merge range
        .Cells(tableStartRow - 1, tableStartCol).HorizontalAlignment = xlCenter
        .Cells(tableStartRow - 1, tableStartCol).Interior.Color = RGB(0, 102, 204)
        .Cells(tableStartRow - 1, tableStartCol).Font.Color = RGB(255, 255, 255)
        
        ' Add border around the title
        With .Range(.Cells(tableStartRow - 1, tableStartCol), .Cells(tableStartRow - 1, tableStartCol + 10)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        
        ' Column headers - UPDATED with new Return Impact column
        .Cells(tableStartRow, tableStartCol).value = "Strategy #"
        .Cells(tableStartRow, tableStartCol + 1).value = "Symbol"
        .Cells(tableStartRow, tableStartCol + 2).value = "Strategy Name"
        .Cells(tableStartRow, tableStartCol + 3).value = "Return Impact"           ' NEW COLUMN
        .Cells(tableStartRow, tableStartCol + 4).value = "Return/Max DD Impact"
        .Cells(tableStartRow, tableStartCol + 5).value = "Return/Avg DD Impact"
        .Cells(tableStartRow, tableStartCol + 6).value = "Profit/Stdev Benefit"
        .Cells(tableStartRow, tableStartCol + 7).value = "Return Rank"             ' NEW RANKING
        .Cells(tableStartRow, tableStartCol + 8).value = "Return/Max DD Rank"
        .Cells(tableStartRow, tableStartCol + 9).value = "Return/Avg DD Rank"
        .Cells(tableStartRow, tableStartCol + 10).value = "P/Stdev Rank"
        
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).Interior.Color = RGB(224, 224, 224)
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).Font.Bold = True
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).WrapText = True
        
        .Cells(tableStartRow, tableStartCol).ColumnWidth = 12.5
        .Cells(tableStartRow, tableStartCol + 1).ColumnWidth = 10
        .Cells(tableStartRow, tableStartCol + 2).ColumnWidth = 75
        '.Cells(tableStartRow, tableStartCol + 3).ColumnWidth = 15  ' Width for Return Impact
        
        ' Populate strategy data
        For i = 1 To numStrategies
            .Cells(tableStartRow + i, tableStartCol).value = metrics(i, 2)  ' Strategy number
            .Cells(tableStartRow + i, tableStartCol + 1).value = metrics(i, 8)  ' Symbol
            .Cells(tableStartRow + i, tableStartCol + 2).value = metrics(i, 3)  ' Strategy name
            
            ' NEW: Return Impact column (percentage-based)
            .Cells(tableStartRow + i, tableStartCol + 3).value = metrics(i, 4)  ' Return Impact
            .Cells(tableStartRow + i, tableStartCol + 3).NumberFormat = "+0.0%;-0.0%;0.0%"  ' Percentage format
            
            ' Updated column assignments
            .Cells(tableStartRow + i, tableStartCol + 4).value = metrics(i, 6)  ' Return/Max DD Impact
            .Cells(tableStartRow + i, tableStartCol + 5).value = metrics(i, 7)  ' Return/Avg DD Impact
            .Cells(tableStartRow + i, tableStartCol + 6).value = metrics(i, 5)  ' Profit/Stdev Benefit
            
            ' Format the percentage cells
            .Cells(tableStartRow + i, tableStartCol + 4).NumberFormat = "+0.0%;-0.0%;0.0%"
            .Cells(tableStartRow + i, tableStartCol + 5).NumberFormat = "+0.0%;-0.0%;0.0%"
            .Cells(tableStartRow + i, tableStartCol + 6).NumberFormat = "+0.0%;-0.0%;0.0%"
            
            ' Updated ranking assignments
            .Cells(tableStartRow + i, tableStartCol + 7).value = rankReturnImpact(i)    ' Return Impact Rank
            .Cells(tableStartRow + i, tableStartCol + 8).value = rankReturnMaxDD(i)      ' Return/Max DD Rank
            .Cells(tableStartRow + i, tableStartCol + 9).value = rankReturnAvgDD(i)      ' Return/Avg DD Rank
            .Cells(tableStartRow + i, tableStartCol + 10).value = rankReturn(i)          ' P/Stdev Rank
            
            ' Color coding for impacts (columns 3-6)
            For j = 3 To 6
                If .Cells(tableStartRow + i, tableStartCol + j).value > 0 Then
                    .Cells(tableStartRow + i, tableStartCol + j).Interior.Color = RGB(198, 239, 206)  ' Light green
                ElseIf .Cells(tableStartRow + i, tableStartCol + j).value < 0 Then
                    .Cells(tableStartRow + i, tableStartCol + j).Interior.Color = RGB(255, 199, 206)  ' Light red
                End If
            Next j
        Next i
        
        ' Add borders to the table
        With .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow + numStrategies, tableStartCol + 10)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        
        ' Create the radar chart that references the existing table data
        Dim strategyRange As Range
        Dim returnRange As Range
        Dim returnMaxDDRange As Range
        Dim returnAvgDDRange As Range
        
        Set strategyRange = .Range(.Cells(tableStartRow + 1, tableStartCol), .Cells(tableStartRow + numStrategies, tableStartCol))
        Set returnRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 6), .Cells(tableStartRow + numStrategies, tableStartCol + 6))        ' Profit/Stdev
        Set returnMaxDDRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 4), .Cells(tableStartRow + numStrategies, tableStartCol + 4))    ' Return/Max DD
        Set returnAvgDDRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 5), .Cells(tableStartRow + numStrategies, tableStartCol + 5))    ' Return/Avg DD
        
        ' Create the chart
        .Shapes.AddChart2(201, xlRadar, .Cells(chartStartRow, chartStartCol).left, _
                        .Cells(chartStartRow, chartStartCol).top, 400, 350).Select
        
        ' Configure the chart
        With ActiveChart
            .HasTitle = True
            .chartTitle.text = "Strategy Contribution Profile"
            
            .HasLegend = True
            .Legend.position = xlLegendPositionBottom
            
            ' Clear any existing series
            Do While .SeriesCollection.count > 0
                .SeriesCollection(1).Delete
            Loop
            
            ' Add series manually using the table ranges
            .SeriesCollection.NewSeries
            .SeriesCollection(1).name = "Profit/Stdev Benefit"
            .SeriesCollection(1).values = returnRange
            .SeriesCollection(1).XValues = strategyRange
            
            .SeriesCollection.NewSeries
            .SeriesCollection(2).name = "Return/Max DD"
            .SeriesCollection(2).values = returnMaxDDRange
            .SeriesCollection(2).XValues = strategyRange
            
            .SeriesCollection.NewSeries
            .SeriesCollection(3).name = "Return/Avg DD"
            .SeriesCollection(3).values = returnAvgDDRange
            .SeriesCollection(3).XValues = strategyRange
            
            ' Format the series with different colors - NO MARKERS
            .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(65, 140, 240)   ' Blue for Return
            .SeriesCollection(1).Format.line.Weight = 2.5
            .SeriesCollection(1).MarkerStyle = xlMarkerStyleNone  ' No markers
            
            .SeriesCollection(2).Format.line.ForeColor.RGB = RGB(252, 180, 65)   ' Orange for Return/Max DD
            .SeriesCollection(2).Format.line.Weight = 2.5
            .SeriesCollection(2).MarkerStyle = xlMarkerStyleNone  ' No markers
            
            .SeriesCollection(3).Format.line.ForeColor.RGB = RGB(127, 96, 170)   ' Purple for Return/Avg DD
            .SeriesCollection(3).Format.line.Weight = 2.5
            .SeriesCollection(3).MarkerStyle = xlMarkerStyleNone  ' No markers
            
            ' Format chart area and plot area
            .ChartArea.Format.fill.ForeColor.RGB = RGB(255, 255, 255)
            .PlotArea.Format.fill.ForeColor.RGB = RGB(242, 242, 242)
            
            ' Format gridlines
            On Error Resume Next  ' In case axes don't have gridlines
            .Axes(xlCategory).MajorGridlines.Format.line.ForeColor.RGB = RGB(191, 191, 191)
            .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(191, 191, 191)
            On Error GoTo 0
        End With
        
        ' Add explanation below the table
        .Cells(tableStartRow + numStrategies + 2, tableStartCol).value = "Interpretation Guide:"
        .Cells(tableStartRow + numStrategies + 2, tableStartCol).Font.Bold = True
        
        .Cells(tableStartRow + numStrategies + 3, tableStartCol).value = "• Return Impact shows the percentage impact on returns when strategy is removed"
        .Cells(tableStartRow + numStrategies + 4, tableStartCol).value = "• Positive values indicate the strategy improves that metric"
        .Cells(tableStartRow + numStrategies + 5, tableStartCol).value = "• Strategies are sorted by " & sortDescription
        .Cells(tableStartRow + numStrategies + 6, tableStartCol).value = "• Lower rank numbers (1 = best) indicate more important strategies"
        
        .Range(.Cells(tableStartRow + numStrategies + 3, tableStartCol), _
              .Cells(tableStartRow + numStrategies + 6, tableStartCol + 10)).Interior.Color = RGB(242, 242, 242)
    End With
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Sub RankStrategiesLeaveOneOutBacktest(wsLeaveOneOut As Worksheet, pnlResults As Variant, _
                                      requiredMargin As Double, startingEquity As Double, _
                                      AverageTrade As Double, MCTradeType As String, _
                                      sortingMetric As String)
    ' Performs "leave-one-out" sensitivity analysis using chronological backtests
    ' Exactly matches the original Monte Carlo function but with chronological order
    
    Dim baselineResults As Variant
    Dim numStrategies As Long, i As Long, j As Long, K As Long
    Dim strategyNames() As String
    Dim Symbols() As String
    Dim strategyNumbers() As Long
    Dim metrics() As Variant
    Dim tableStartRow As Long, tableStartCol As Long
    Dim chartStartRow As Long, chartStartCol As Long
    
    ' Baseline metrics - EXACT SAME VARIABLES AS ORIGINAL
    Dim baselineReturn As Double
    Dim baselineMaxDD As Double
    Dim baselineAvgDD As Double
    Dim baselineReturnToMaxDD As Double
    Dim baselineReturnToAvgDD As Double
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing baseline portfolio performance..."
    
    ' Define start positions
    tableStartRow = 6
    tableStartCol = 12  ' Column L
    chartStartRow = 6
    chartStartCol = 2   ' Column B
    
    ' Determine number of strategies
    numStrategies = UBound(pnlResults, 2)
    If numStrategies < 2 Then
        MsgBox "At least two strategies are required for comparative analysis.", vbExclamation
        Exit Sub
    End If
    
    ' Get strategy names and numbers from the Portfolio sheet - SAME AS ORIGINAL
    ReDim strategyNames(1 To numStrategies)
    ReDim strategyNumbers(1 To numStrategies)
    ReDim Symbols(1 To numStrategies)
    On Error Resume Next
    For i = 1 To numStrategies
        strategyNames(i) = ThisWorkbook.Sheets("Portfolio").Cells(i + 1, COL_PORT_STRATEGY_NAME).value
        strategyNumbers(i) = ThisWorkbook.Sheets("Portfolio").Cells(i + 1, COL_PORT_STRATEGY_NUMBER).value
        Symbols(i) = ThisWorkbook.Sheets("Portfolio").Cells(i + 1, COL_PORT_SYMBOL).value
        If strategyNames(i) = "" Then strategyNames(i) = "Strategy " & i
        If strategyNumbers(i) = 0 Then strategyNumbers(i) = i
        If Symbols(i) = "" Then Symbols(i) = "Error"
    Next i
    On Error GoTo 0
    
    ' Run baseline backtest with ALL strategies - REPLACE MONTE CARLO WITH CHRONOLOGICAL
    baselineResults = RunChronologicalBacktestWithTracking(pnlResults, requiredMargin, startingEquity, AverageTrade)
    
    ' Calculate baseline metrics - EXACT SAME LOGIC AS ORIGINAL
    baselineReturn = CalculateMedianReturn(baselineResults)
    baselineMaxDD = CalculateMedianMaxDrawdown(baselineResults)
    baselineAvgDD = CalculateMedianAvgDrawdown(baselineResults)
    
    ' Calculate the ratios - EXACT SAME AS ORIGINAL
    If baselineMaxDD > 0 Then baselineReturnToMaxDD = baselineReturn / baselineMaxDD Else baselineReturnToMaxDD = 0
    If baselineAvgDD > 0 Then baselineReturnToAvgDD = baselineReturn / baselineAvgDD Else baselineReturnToAvgDD = 0
    
    ' Create results array for each strategy's impact - EXPANDED to include dollar profits
    ReDim metrics(1 To numStrategies, 1 To 8)  ' Added one more column for dollar profits
    
    ' Run backtests excluding one strategy at a time - SAME LOOP AS ORIGINAL
    For i = 1 To numStrategies
        Application.StatusBar = "Analyzing without strategy " & i & " of " & numStrategies & ": " & strategyNames(i)
        
        ' Create a modified version of pnlResults without the current strategy - SAME AS ORIGINAL
        Dim modifiedPnlResults As Variant
        modifiedPnlResults = ExcludeStrategyFromPnL(pnlResults, i)
        
        ' Run chronological backtest with this strategy excluded - REPLACE MONTE CARLO
        Dim excludedResults As Variant
        excludedResults = RunChronologicalBacktestWithTracking(modifiedPnlResults, requiredMargin, startingEquity, AverageTrade)
        
        ' Calculate excluded metrics directly - SAME AS ORIGINAL
        Dim excludedReturn As Double
        Dim excludedMaxDD As Double
        Dim excludedAvgDD As Double
        Dim excludedReturnToMaxDD As Double
        Dim excludedReturnToAvgDD As Double
        
        excludedReturn = CalculateMedianReturn(excludedResults)
        excludedMaxDD = CalculateMedianMaxDrawdown(excludedResults)
        excludedAvgDD = CalculateMedianAvgDrawdown(excludedResults)
        
        ' Calculate the ratios - SAME AS ORIGINAL
        If excludedMaxDD > 0 Then excludedReturnToMaxDD = excludedReturn / excludedMaxDD Else excludedReturnToMaxDD = 0
        If excludedAvgDD > 0 Then excludedReturnToAvgDD = excludedReturn / excludedAvgDD Else excludedReturnToAvgDD = 0
        
        ' Store the difference between baseline and excluded metrics
        metrics(i, 1) = i  ' Strategy index
        metrics(i, 2) = strategyNumbers(i)  ' Strategy number
        metrics(i, 3) = strategyNames(i)  ' Strategy name
        metrics(i, 8) = Symbols(i)  ' Symbol (moved to column 8)
        
        ' NEW: Calculate Dollar Profits Impact (column 4) - Using percentage change logic
        If Abs(baselineReturn) > 0.00001 Then
            metrics(i, 4) = (baselineReturn - excludedReturn) / baselineReturn
        Else
            metrics(i, 4) = 0
        End If
        
        ' Calculate Profit/Stdev Benefit using proper standard deviation from PnL data (moved to column 5)
        Dim baselineStdev As Double, excludedStdev As Double
        Dim baselineProfitStdev As Double, excludedProfitStdev As Double
        
        ' Calculate standard deviation from the actual PnL data
        baselineStdev = CalculateBacktestStdev(pnlResults)
        excludedStdev = CalculateBacktestStdev(modifiedPnlResults)
        
        ' Calculate Profit/Stdev ratios
        If baselineStdev > 0 Then baselineProfitStdev = baselineReturn / baselineStdev Else baselineProfitStdev = baselineReturn
        If excludedStdev > 0 Then excludedProfitStdev = excludedReturn / excludedStdev Else excludedProfitStdev = excludedReturn
        
        ' Store Profit/Stdev Benefit (moved to column 5)
        If Abs(excludedProfitStdev) > 0.00001 Then
            metrics(i, 5) = (baselineProfitStdev - excludedProfitStdev) / Abs(excludedProfitStdev)
        Else
            metrics(i, 5) = 0
        End If
        
        ' Impact on Return/Max DD as percentage change (moved to column 6)
        If Abs(baselineReturnToMaxDD) > 0.00001 Then
            Dim returnMaxDDImpact As Double
            returnMaxDDImpact = (baselineReturnToMaxDD - excludedReturnToMaxDD) / baselineReturnToMaxDD
            
            ' Filter out very small values that would display as "-0"
            If Abs(returnMaxDDImpact) < 0.001 Then
                metrics(i, 6) = 0
            Else
                metrics(i, 6) = returnMaxDDImpact
            End If
        Else
            metrics(i, 6) = 0
        End If
        
        ' Impact on Return/Avg DD as percentage change (moved to column 7)
        If Abs(baselineReturnToAvgDD) > 0.00001 Then
            metrics(i, 7) = (baselineReturnToAvgDD - excludedReturnToAvgDD) / baselineReturnToAvgDD
        Else
            metrics(i, 7) = 0
        End If
    Next i
    
    ' Sort metrics by selected criterion
    Call SortLeaveOneOutMetricsWithReturnImpact(metrics, sortingMetric, numStrategies)
    
    ' Calculate ranks for each metric - UPDATED for new column structure
    Dim rankDollarProfits() As Long
    Dim rankReturn() As Long
    Dim rankReturnMaxDD() As Long
    Dim rankReturnAvgDD() As Long
    
    rankDollarProfits = RankStrategiesByMetric(metrics, 4)  ' Rank by Dollar Profits (now percentage-based)
    rankReturn = RankStrategiesByMetric(metrics, 5)  ' Rank by Profit/Stdev benefit
    rankReturnMaxDD = RankStrategiesByMetric(metrics, 6)  ' Rank by return/max drawdown impact
    rankReturnAvgDD = RankStrategiesByMetric(metrics, 7)  ' Rank by return/avg drawdown impact
    
    ' Get sorting description
    Dim sortDescription As String
    Select Case sortingMetric
        Case "1": sortDescription = "Return/Max DD Impact"
        Case "2": sortDescription = "Return/Avg DD Impact"
        Case "3": sortDescription = "Profit/Stdev Benefit"
        Case "4": sortDescription = "Strategy Number"
        Case "5": sortDescription = "Return Impact"  ' Updated name to reflect percentage calculation
        Case Else: sortDescription = "Return/Max DD Impact"
    End Select
    
    ' Output results to worksheet - UPDATED with new Dollar Profits column
    With wsLeaveOneOut
        ' Title for strategy analysis section
        .Cells(tableStartRow - 1, tableStartCol).value = "Strategy Contribution Analysis (" & MCTradeType & ") - Sorted by " & sortDescription
        .Cells(tableStartRow - 1, tableStartCol).Font.Bold = True
        .Cells(tableStartRow - 1, tableStartCol).Font.Size = 14
        .Range(.Cells(tableStartRow - 1, tableStartCol), .Cells(tableStartRow - 1, tableStartCol + 10)).Merge  ' Expanded merge range
        .Cells(tableStartRow - 1, tableStartCol).HorizontalAlignment = xlCenter
        .Cells(tableStartRow - 1, tableStartCol).Interior.Color = RGB(0, 102, 204)
        .Cells(tableStartRow - 1, tableStartCol).Font.Color = RGB(255, 255, 255)
        
        ' Add border around the title
        With .Range(.Cells(tableStartRow - 1, tableStartCol), .Cells(tableStartRow - 1, tableStartCol + 10)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        
        ' Column headers - UPDATED with new Return Impact column
        .Cells(tableStartRow, tableStartCol).value = "Strategy #"
        .Cells(tableStartRow, tableStartCol + 1).value = "Symbol"
        .Cells(tableStartRow, tableStartCol + 2).value = "Strategy Name"
        .Cells(tableStartRow, tableStartCol + 3).value = "Return Impact"           ' NEW COLUMN (percentage-based)
        .Cells(tableStartRow, tableStartCol + 4).value = "Return/Max DD Impact"
        .Cells(tableStartRow, tableStartCol + 5).value = "Return/Avg DD Impact"
        .Cells(tableStartRow, tableStartCol + 6).value = "Profit/Stdev Benefit"
        .Cells(tableStartRow, tableStartCol + 7).value = "Return Rank"             ' NEW RANKING
        .Cells(tableStartRow, tableStartCol + 8).value = "Return/Max DD Rank"
        .Cells(tableStartRow, tableStartCol + 9).value = "Return/Avg DD Rank"
        .Cells(tableStartRow, tableStartCol + 10).value = "P/Stdev Rank"
        
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).Interior.Color = RGB(224, 224, 224)
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).Font.Bold = True
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).WrapText = True
        
        .Cells(tableStartRow, tableStartCol).ColumnWidth = 12.5
        .Cells(tableStartRow, tableStartCol + 1).ColumnWidth = 10
        .Cells(tableStartRow, tableStartCol + 2).ColumnWidth = 75
      '  .Cells(tableStartRow, tableStartCol + 3).ColumnWidth = 15  ' Width for Return Impact
        
        ' Populate strategy data
        For i = 1 To numStrategies
            .Cells(tableStartRow + i, tableStartCol).value = metrics(i, 2)  ' Strategy number
            .Cells(tableStartRow + i, tableStartCol + 1).value = metrics(i, 8)  ' Symbol
            .Cells(tableStartRow + i, tableStartCol + 2).value = metrics(i, 3)  ' Strategy name
            
            ' NEW: Return Impact column (percentage-based)
            .Cells(tableStartRow + i, tableStartCol + 3).value = metrics(i, 4)  ' Return Impact
            .Cells(tableStartRow + i, tableStartCol + 3).NumberFormat = "+0.0%;-0.0%;0.0%"  ' Percentage format
            
            ' Updated column assignments
            .Cells(tableStartRow + i, tableStartCol + 4).value = metrics(i, 6)  ' Return/Max DD Impact
            .Cells(tableStartRow + i, tableStartCol + 5).value = metrics(i, 7)  ' Return/Avg DD Impact
            .Cells(tableStartRow + i, tableStartCol + 6).value = metrics(i, 5)  ' Profit/Stdev Benefit
            
            ' Format the percentage cells
            .Cells(tableStartRow + i, tableStartCol + 4).NumberFormat = "+0.0%;-0.0%;0.0%"
            .Cells(tableStartRow + i, tableStartCol + 5).NumberFormat = "+0.0%;-0.0%;0.0%"
            .Cells(tableStartRow + i, tableStartCol + 6).NumberFormat = "+0.0%;-0.0%;0.0%"
            
            ' Updated ranking assignments
            .Cells(tableStartRow + i, tableStartCol + 7).value = rankDollarProfits(i)    ' Return Impact Rank
            .Cells(tableStartRow + i, tableStartCol + 8).value = rankReturnMaxDD(i)      ' Return/Max DD Rank
            .Cells(tableStartRow + i, tableStartCol + 9).value = rankReturnAvgDD(i)      ' Return/Avg DD Rank
            .Cells(tableStartRow + i, tableStartCol + 10).value = rankReturn(i)          ' P/Stdev Rank
            
            ' Color coding for impacts (columns 3-6)
            For j = 3 To 6
                If .Cells(tableStartRow + i, tableStartCol + j).value > 0 Then
                    .Cells(tableStartRow + i, tableStartCol + j).Interior.Color = RGB(198, 239, 206)  ' Light green
                ElseIf .Cells(tableStartRow + i, tableStartCol + j).value < 0 Then
                    .Cells(tableStartRow + i, tableStartCol + j).Interior.Color = RGB(255, 199, 206)  ' Light red
                End If
            Next j
        Next i
        
        ' Add borders to the table
        With .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow + numStrategies, tableStartCol + 10)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        
        ' Create the radar chart that references the existing table data
        Dim strategyRange As Range
        Dim returnRange As Range
        Dim returnMaxDDRange As Range
        Dim returnAvgDDRange As Range
        
        Set strategyRange = .Range(.Cells(tableStartRow + 1, tableStartCol), .Cells(tableStartRow + numStrategies, tableStartCol))
        Set returnRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 6), .Cells(tableStartRow + numStrategies, tableStartCol + 6))        ' Profit/Stdev
        Set returnMaxDDRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 4), .Cells(tableStartRow + numStrategies, tableStartCol + 4))    ' Return/Max DD
        Set returnAvgDDRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 5), .Cells(tableStartRow + numStrategies, tableStartCol + 5))    ' Return/Avg DD
        
        ' Create the chart
        .Shapes.AddChart2(201, xlRadar, .Cells(chartStartRow, chartStartCol).left, _
                        .Cells(chartStartRow, chartStartCol).top, 400, 350).Select
        
        ' Configure the chart
        With ActiveChart
            .HasTitle = True
            .chartTitle.text = "Strategy Contribution Profile"
            
            .HasLegend = True
            .Legend.position = xlLegendPositionBottom
            
            ' Clear any existing series
            Do While .SeriesCollection.count > 0
                .SeriesCollection(1).Delete
            Loop
            
            ' Add series manually using the table ranges
            .SeriesCollection.NewSeries
            .SeriesCollection(1).name = "Profit/Stdev Benefit"
            .SeriesCollection(1).values = returnRange
            .SeriesCollection(1).XValues = strategyRange
            
            .SeriesCollection.NewSeries
            .SeriesCollection(2).name = "Return/Max DD"
            .SeriesCollection(2).values = returnMaxDDRange
            .SeriesCollection(2).XValues = strategyRange
            
            .SeriesCollection.NewSeries
            .SeriesCollection(3).name = "Return/Avg DD"
            .SeriesCollection(3).values = returnAvgDDRange
            .SeriesCollection(3).XValues = strategyRange
            
            ' Format the series with different colors - NO MARKERS
            .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(65, 140, 240)   ' Blue for Return
            .SeriesCollection(1).Format.line.Weight = 2.5
            .SeriesCollection(1).MarkerStyle = xlMarkerStyleNone  ' No markers
            
            .SeriesCollection(2).Format.line.ForeColor.RGB = RGB(252, 180, 65)   ' Orange for Return/Max DD
            .SeriesCollection(2).Format.line.Weight = 2.5
            .SeriesCollection(2).MarkerStyle = xlMarkerStyleNone  ' No markers
            
            .SeriesCollection(3).Format.line.ForeColor.RGB = RGB(127, 96, 170)   ' Purple for Return/Avg DD
            .SeriesCollection(3).Format.line.Weight = 2.5
            .SeriesCollection(3).MarkerStyle = xlMarkerStyleNone  ' No markers
            
            ' Format chart area and plot area
            .ChartArea.Format.fill.ForeColor.RGB = RGB(255, 255, 255)
            .PlotArea.Format.fill.ForeColor.RGB = RGB(242, 242, 242)
            
            ' Format gridlines
            On Error Resume Next  ' In case axes don't have gridlines
            .Axes(xlCategory).MajorGridlines.Format.line.ForeColor.RGB = RGB(191, 191, 191)
            .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(191, 191, 191)
            On Error GoTo 0
        End With
        
        ' Add explanation below the table
        .Cells(tableStartRow + numStrategies + 2, tableStartCol).value = "Interpretation Guide:"
        .Cells(tableStartRow + numStrategies + 2, tableStartCol).Font.Bold = True
        
        .Cells(tableStartRow + numStrategies + 3, tableStartCol).value = "• Return Impact shows the percentage impact on returns when strategy is removed"
        .Cells(tableStartRow + numStrategies + 4, tableStartCol).value = "• Positive values indicate the strategy improves that metric"
        .Cells(tableStartRow + numStrategies + 5, tableStartCol).value = "• Strategies are sorted by " & sortDescription
        .Cells(tableStartRow + numStrategies + 6, tableStartCol).value = "• Lower rank numbers (1 = best) indicate more important strategies"
        
        .Range(.Cells(tableStartRow + numStrategies + 3, tableStartCol), _
              .Cells(tableStartRow + numStrategies + 6, tableStartCol + 10)).Interior.Color = RGB(242, 242, 242)
    End With
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Sub SortLeaveOneOutMetricsWithReturnImpact(ByRef metrics As Variant, sortingMetric As String, numStrategies As Long)
    Dim i As Long, j As Long, K As Long
    Dim tempRow As Variant
    Dim sortColumn As Long
    Dim sortDescending As Boolean
    
    ReDim tempRow(1 To UBound(metrics, 2))
    
    ' Updated sort column mappings for new structure
    Select Case sortingMetric
        Case "1"
            sortColumn = 6  ' Return/Max DD Impact (now in column 6)
            sortDescending = True
        Case "2"
            sortColumn = 7  ' Return/Avg DD Impact (now in column 7)
            sortDescending = True
        Case "3"
            sortColumn = 5  ' Profit/Stdev Benefit (now in column 5)
            sortDescending = True
        Case "4"
            sortColumn = 2  ' Strategy Number
            sortDescending = False
        Case "5"
            sortColumn = 4  ' Return Impact (new column 4)
            sortDescending = True
        Case Else
            sortColumn = 6  ' Default to Return/Max DD Impact
            sortDescending = True
    End Select
    
    ' Only sort if not keeping original order
    If sortingMetric <> "4" Then
        For i = 1 To numStrategies - 1
            For j = i + 1 To numStrategies
                Dim shouldSwap As Boolean
                
                If sortDescending Then
                    shouldSwap = (metrics(i, sortColumn) < metrics(j, sortColumn))
                Else
                    shouldSwap = (metrics(i, sortColumn) > metrics(j, sortColumn))
                End If
                
                If shouldSwap Then
                    For K = 1 To UBound(metrics, 2)
                        tempRow(K) = metrics(i, K)
                        metrics(i, K) = metrics(j, K)
                        metrics(j, K) = tempRow(K)
                    Next K
                End If
            Next j
        Next i
    End If
End Sub

Function RunChronologicalBacktestWithTracking(pnlResults As Variant, requiredMargin As Double, _
                                              startingEquity As Double, AverageTrade As Double) As Variant
    ' Runs a single chronological backtest and returns results in the same format as Monte Carlo
    ' Uses the ORIGINAL PnL data without any adjustments
    ' NO RUIN CHECKING
    
    Dim numPeriods As Long, numStrategies As Long
    Dim i As Long, j As Long
    Dim equity As Double, peakEquity As Double, maxDrawdown As Double
    Dim periodPnL As Double, totalReturn As Double
    Dim results(1 To 1, 1 To 8) As Double  ' Single row to match Monte Carlo format
    Dim drawdownSum As Double, drawdownCount As Long, avgDrawdown As Double
    
    numPeriods = UBound(pnlResults, 1)
    numStrategies = UBound(pnlResults, 2)
    
    ' Initialize
    equity = startingEquity
    peakEquity = startingEquity
    maxDrawdown = 0
    drawdownSum = 0
    drawdownCount = 0
    
    ' Run chronological backtest using ORIGINAL PnL data (no adjustments)
    For i = 1 To numPeriods
        ' Calculate period P&L (sum across all strategies) - USE ORIGINAL DATA
        periodPnL = 0
        For j = 1 To numStrategies
            periodPnL = periodPnL + pnlResults(i, j)  ' No adjustments applied
        Next j
        
        ' Update equity
        equity = equity + periodPnL
        
        ' Update peak and drawdown tracking - TRACK DOLLAR AMOUNTS
        If equity > peakEquity Then
            peakEquity = equity
        End If
        
        ' Calculate current dollar drawdown
        Dim currentDollarDrawdown As Double
        currentDollarDrawdown = peakEquity - equity
        
        ' Track maximum dollar drawdown
        If currentDollarDrawdown > maxDrawdown Then
            maxDrawdown = currentDollarDrawdown  ' Store dollar amount, not percentage
        End If
        
        ' Track for average drawdown calculation (in dollars)
        If currentDollarDrawdown > 0 Then
            drawdownSum = drawdownSum + currentDollarDrawdown
            drawdownCount = drawdownCount + 1
        End If
    Next i
    
    ' Calculate average drawdown in dollars
    If drawdownCount > 0 Then
        avgDrawdown = drawdownSum / drawdownCount
    Else
        avgDrawdown = 0
    End If
    
    ' Calculate final metrics
    Dim totalDollarReturn As Double
    totalDollarReturn = equity - startingEquity
    
    ' Populate results array to match Monte Carlo format
    results(1, 1) = 1        ' Scenario number
    results(1, 2) = totalDollarReturn           ' Total profit in DOLLARS
    results(1, 3) = totalDollarReturn           ' Dollar return
    
    ' Calculate Return/Max DD as dollars/dollars
    If maxDrawdown > 0.01 Then  ' Small threshold to avoid division by very small numbers
        results(1, 4) = totalDollarReturn / maxDrawdown  ' Dollar Return / Dollar Max DD
    Else
        results(1, 4) = totalDollarReturn * 1000  ' No meaningful drawdown
    End If
    
    results(1, 5) = maxDrawdown           ' Max drawdown in DOLLARS
    results(1, 6) = 0        ' No ruin flag (always 0)
    results(1, 7) = numPeriods   ' Total periods
    results(1, 8) = avgDrawdown  ' Average drawdown in DOLLARS
    
    RunChronologicalBacktestWithTracking = results
End Function

' Helper function to calculate the median return from results
Function CalculateMedianReturn(results As Variant) As Double
    ' Extract DOLLAR returns from column 2 instead of column 3 (percentage)
    Dim returnArray() As Double
    Dim i As Long
    
    ReDim returnArray(1 To UBound(results, 1))
    
    For i = 1 To UBound(results, 1)
        returnArray(i) = results(i, 2)  ' CHANGED: Use column 2 (dollar return) instead of column 3 (percentage)
    Next i
    
    CalculateMedianReturn = Application.WorksheetFunction.Median(returnArray)
End Function

Function CalculateMedianMaxDrawdown(results As Variant) As Double
    ' Extract DOLLAR max drawdowns from column 5
    Dim maxDDArray() As Double
    Dim i As Long
    
    ReDim maxDDArray(1 To UBound(results, 1))
    
    For i = 1 To UBound(results, 1)
        maxDDArray(i) = results(i, 5)  ' Dollar max drawdown
    Next i
    
    CalculateMedianMaxDrawdown = Application.WorksheetFunction.Median(maxDDArray)
End Function

Function CalculateMedianAvgDrawdown(results As Variant) As Double
    ' Extract DOLLAR average drawdowns from column 7
    Dim avgDDArray() As Double
    Dim i As Long
    
    ReDim avgDDArray(1 To UBound(results, 1))
    
    For i = 1 To UBound(results, 1)
        avgDDArray(i) = results(i, 7)  ' Dollar average drawdown
    Next i
    
    CalculateMedianAvgDrawdown = Application.WorksheetFunction.Median(avgDDArray)
End Function


Function ExcludeStrategyFromPnL(pnlResults As Variant, strategyIndex As Long, Optional MCTradeType As String = "Daily") As Variant
    ' Creates a copy of pnlResults with one strategy excluded
    ' Works with both daily and weekly multi-strategy data using the same approach
    ' Parameters:
    '   pnlResults - 2D array of PnL data (periods x strategies)
    '   strategyIndex - Index of the strategy to exclude (zero out)
    '   MCTradeType - "Daily" or "Weekly" (for documentation only, doesn't affect behavior)
    
    Dim numPeriods As Long, numStrategies As Long
    Dim i As Long, j As Long
    Dim modifiedResults As Variant
    
    ' Get dimensions
    numPeriods = UBound(pnlResults, 1)
    numStrategies = UBound(pnlResults, 2)
    
    ' Create a copy of the original array
    ReDim modifiedResults(1 To numPeriods, 1 To numStrategies)
    
    ' Copy all data, zeroing out the target strategy
    For i = 1 To numPeriods
        For j = 1 To numStrategies
            If j = strategyIndex Then
                modifiedResults(i, j) = 0  ' Zero out this strategy
            Else
                modifiedResults(i, j) = pnlResults(i, j)  ' Copy other strategies as-is
            End If
        Next j
    Next i
    
    ExcludeStrategyFromPnL = modifiedResults
End Function


'— return the date of the nth weekday in a month (using vbMonday: Mon=1…Sun=7)—
Function NthWeekdayOfMonth( _
    ByVal Y As Long, _
    ByVal m As Long, _
    ByVal weekdayIndex As Long, _
    ByVal n As Long _
) As Date
    Dim d As Date, cnt As Long
    For d = DateSerial(Y, m, 1) To DateSerial(Y, m + 1, 0)
        If Weekday(d, vbMonday) = weekdayIndex Then
            cnt = cnt + 1
            If cnt = n Then
                NthWeekdayOfMonth = d
                Exit Function
            End If
        End If
    Next d
End Function

'— return the last Monday of a given month/year —
Function LastMondayOfMonth( _
    ByVal Y As Long, _
    ByVal m As Long _
) As Date
    Dim d As Date
    d = DateSerial(Y, m + 1, 0) ' last day of month
    Do While Weekday(d, vbMonday) <> 1 ' Monday=1
        d = d - 1
    Loop
    LastMondayOfMonth = d
End Function





Sub AddNavigationButtonsLeaveOneOut(ws As Worksheet)
    ' Add navigation buttons to the worksheet
    Dim btn As Object
    
    ' Create delete button
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 1).left + 30, top:=ws.Cells(35, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteLeaveOneOut" ' Make sure to create this sub to handle deletion
    End With

    ' Create a button to return to the Summary page
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 4).left + 30, top:=ws.Cells(35, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary" ' Assign the macro to run when the button is clicked
    End With
 
    ' Create a button to return to the Portfolio page
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 1).left + 30, top:=ws.Cells(38, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio" ' Assign the macro to run when the button is clicked
    End With
    
    ' Create a button to return to the Control page
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 4).left + 30, top:=ws.Cells(38, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl" ' Assign the macro to run when the button is clicked
    End With
    
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 1).left + 30, top:=ws.Cells(41, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies" ' Assign the macro to run when the button is clicked
    End With
    
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 4).left + 30, top:=ws.Cells(41, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs" ' Assign the macro to run when the button is clicked
    End With
End Sub



Function CalculateBacktestStdev(pnlResults As Variant) As Double
    ' Calculate standard deviation from PnL data - matches diversificator function
    Dim i As Long, j As Long, numPeriods As Long
    Dim periodReturns() As Double, avgReturn As Double, variance As Double
    
    numPeriods = UBound(pnlResults, 1)
    ReDim periodReturns(1 To numPeriods)
    
    ' Sum all strategies' returns for each period
    For i = 1 To numPeriods
        periodReturns(i) = 0
        For j = 1 To UBound(pnlResults, 2)
            periodReturns(i) = periodReturns(i) + pnlResults(i, j)
        Next j
    Next i
    
    ' Calculate average return
    avgReturn = 0
    For i = 1 To numPeriods
        avgReturn = avgReturn + periodReturns(i)
    Next i
    avgReturn = avgReturn / numPeriods
    
    ' Calculate variance
    variance = 0
    For i = 1 To numPeriods
        variance = variance + (periodReturns(i) - avgReturn) ^ 2
    Next i
    variance = variance / (numPeriods - 1)
    
    CalculateBacktestStdev = Sqr(variance)
End Function

Function CalculatePortfolioStdevFromPnL(pnlResults As Variant) As Double
    ' This is identical to CalculateBacktestStdev - use the same logic
    CalculatePortfolioStdevFromPnL = CalculateBacktestStdev(pnlResults)
End Function


Function GenerateRandomIndexMatrix(pnlResults As Variant, numScenarios As Long, numPeriods As Long) As Variant
    Dim i As Long, p As Long
    ReDim gRandIdx(1 To numScenarios, 1 To numPeriods)
    Randomize Int(Rnd * 100000) + 1  ' or a fixed seed of your choice
    For i = 1 To numScenarios
        For p = 1 To numPeriods
            ' pick a day at random (1 to numPeriods)
            gRandIdx(i, p) = Int(Rnd * UBound(pnlResults, 1)) + 1
        Next p
    Next i
    
    GenerateRandomIndexMatrix = gRandIdx
End Function

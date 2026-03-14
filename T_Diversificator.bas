Attribute VB_Name = "T_Diversificator"
Sub RunDiversificationAnalysis()
    Dim wsPortfolio    As Worksheet
    Dim wsSizingGraphs As Worksheet
    Dim wsDiversification  As Worksheet
    Dim currentdate    As Date, startdate As Date, endDate As Date
    Dim yearsToConsider As Double
    Dim MCTradeType    As String
    Dim ceaseType      As String
    Dim requiredMargin As Double, margin As Double
    Dim startingEquity As Double
    Dim pnlResults     As Variant
    Dim AverageTrade   As Double
    Dim numStrategies  As Long
    Dim i As Long, j As Long, K As Long
    Dim TradeCount As Long
    Dim analysisMethod As String
    Dim sortingMetric As String
    
    '— 0) License & setup —
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license.", vbCritical
        Exit Sub
    End If
    Call InitializeColumnConstantsManually
    
    ' Ask user which method to use
    analysisMethod = InputBox("Choose analysis method:" & vbCrLf & _
                             "1 = Randomized (shows average diversification benefit)" & vbCrLf & _
                             "2 = Greedy (shows optimal sequential selection)", _
                             "Diversification Analysis Method", "1")
    
    ' Check if user pressed Cancel
    If analysisMethod = "" Then Exit Sub
    
    If analysisMethod <> "1" And analysisMethod <> "2" Then
        MsgBox "Invalid selection. Using Randomized method.", vbInformation
        analysisMethod = "1"
    End If
    
    ' Ask user which metric to sort by - UPDATED TO INCLUDE RETURN IMPACT
    sortingMetric = InputBox("Choose sorting metric:" & vbCrLf & _
                            "1 = Profit/Max DD Benefit (default)" & vbCrLf & _
                            "2 = Profit/Avg DD Benefit" & vbCrLf & _
                            "3 = Profit/Stdev Benefit" & vbCrLf & _
                            "4 = Strategy Number (original order)" & vbCrLf & _
                            "5 = Return Impact", _
                            "Sorting Metric", "1")
    
    ' Check if user pressed Cancel
    If sortingMetric = "" Then Exit Sub
    
    If sortingMetric <> "1" And sortingMetric <> "2" And sortingMetric <> "3" And sortingMetric <> "4" And sortingMetric <> "5" Then
        MsgBox "Invalid selection. Using Profit/Max DD Benefit.", vbInformation
        sortingMetric = "1"
    End If
    
    ' Rest of the function remains the same...
    ' [All the existing code for getting inputs, creating sheets, etc.]
    
    ' If randomized method, ask for number of iterations
    Dim numIterations As Long
    If analysisMethod = "1" Then
        Dim iterationsInput As String
        iterationsInput = InputBox("Enter number of random iterations:" & vbCrLf & _
                                  "Recommended: 100-300" & vbCrLf & _
                                  "More iterations = more accurate but slower", _
                                  "Number of Iterations", "100")
        
        ' Check if user pressed Cancel
        If iterationsInput = "" Then Exit Sub
        
        On Error Resume Next
        numIterations = CLng(iterationsInput)
        On Error GoTo 0
        
        If numIterations < 10 Or numIterations > 1000 Or numIterations = 0 Then
            MsgBox "Invalid number of iterations. Using 100.", vbInformation
            numIterations = 100
        End If
    Else
        numIterations = 1  ' Not used for greedy method
    End If
    
    '— 1) Get Portfolio sheet & inputs —
    On Error Resume Next
      Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    On Error GoTo 0
    If wsPortfolio Is Nothing Then
        MsgBox "'Portfolio' sheet is missing!", vbExclamation: Exit Sub
    End If
    
    On Error Resume Next
      Set wsSizingGraphs = ThisWorkbook.Sheets("SizingGraphs")
    On Error GoTo 0
    If wsPortfolio Is Nothing Then
        MsgBox "'SizingGraphs' sheet is missing!", vbExclamation: Exit Sub
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
    
    ' Compute actual margin threshold
    If LCase(ceaseType) = "percentage" Then
        margin = (1 - requiredMargin) * startingEquity
    Else
        margin = requiredMargin
    End If
    
    '— 2) Build PnL array —
    If LCase(MCTradeType) = "daily" Then
        pnlResults = CleanPortfolioDailyPnL(startdate, endDate)
    ElseIf LCase(MCTradeType) = "weekly" Then
        pnlResults = ConvertDailyToWeeklyPnL(startdate, endDate)
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
    
    '— 4) Create / clear the Diversification sheet —
    On Error Resume Next
      Application.DisplayAlerts = False
      ThisWorkbook.Sheets("Diversificator").Delete
      Application.DisplayAlerts = True
    On Error GoTo 0
    
    If SheetExists("wsSizingGraphs") Then
      Set wsDiversification = ThisWorkbook.Sheets.Add(After:=wsSizingGraphs)
    Else
      Set wsDiversification = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    End If
      
    wsDiversification.name = "Diversificator"
    wsDiversification.Tab.Color = RGB(117, 219, 255)
    wsDiversification.Cells.Interior.Color = RGB(255, 255, 255)
    
    '— 5) Run the selected diversification analysis —
    If analysisMethod = "1" Then
        Call RankStrategiesByDiversificationBenefit( _
            wsDiversification, _
            pnlResults, _
            margin, _
            startingEquity, _
            AverageTrade, _
            MCTradeType, _
            numIterations, _
            sortingMetric _
        )
    Else
        Call RankStrategiesByGreedySelection( _
            wsDiversification, _
            pnlResults, _
            margin, _
            startingEquity, _
            AverageTrade, _
            MCTradeType, _
            sortingMetric _
        )
    End If
    '— 6) Add navigation buttons —
    AddNavigationButtonsDiversificator wsDiversification
    
    ' Set window properties
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
    
    '— 7) Reorder tabs & finish up —
  '  Call OrderVisibleTabsBasedOnList
    Application.ScreenUpdating = True
    wsDiversification.Activate
    
    MsgBox "Diversification analysis complete!", vbInformation
End Sub


Sub RankStrategiesByDiversificationBenefit(wsDiversification As Worksheet, pnlResults As Variant, _
                                          requiredMargin As Double, startingEquity As Double, _
                                          AverageTrade As Double, MCTradeType As String, _
                                          numIterations As Long, sortingMetric As String)
    ' ULTRA-OPTIMIZED VERSION - Fixed with proper error handling and status updates
    
    Dim numStrategies As Long, numDays As Long
    Dim i As Long, j As Long, K As Long, iteration As Long
    Dim strategyNames() As String, Symbols() As String, strategyNumbers() As Long
    Dim contributionMatrix() As Double
    Dim medianContributions() As Double
    Dim metrics() As Variant
    Dim tableStartRow As Long, tableStartCol As Long
    Dim chartStartRow As Long, chartStartCol As Long
    
    On Error GoTo ErrorHandler
    
    ' CRITICAL: Excel performance optimizations - but keep status bar working
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ' Note: We'll temporarily enable ScreenUpdating for status bar updates
    
    numStrategies = UBound(pnlResults, 2)
    numDays = UBound(pnlResults, 1)
    tableStartRow = 6: tableStartCol = 12
    chartStartRow = 6: chartStartCol = 2
    
    If numStrategies < 2 Then
        MsgBox "At least two strategies are required.", vbExclamation
        GoTo Cleanup
    End If
    
    ' Status update with screen refresh
    Application.ScreenUpdating = True
    Application.StatusBar = "Getting strategy information..."
    Application.ScreenUpdating = False
    
    ' Get strategy information
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
        If Symbols(i) = "" Then Symbols(i) = "S" & i
    Next i
    On Error GoTo ErrorHandler
    
    ' OPTIMIZATION 1: Pre-calculate ALL individual strategy metrics
    Dim individualMetrics() As Double
    ReDim individualMetrics(1 To numStrategies, 1 To 7)
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Pre-calculating individual strategy metrics..."
    Application.ScreenUpdating = False
    
    For i = 1 To numStrategies
        Call CalculateAllStrategyMetricsFast(pnlResults, i, startingEquity, individualMetrics, i)
    Next i
    
    ' OPTIMIZATION 2: Pre-generate ALL random orders
    Dim allRandomOrders() As Long
    ReDim allRandomOrders(1 To numIterations, 1 To numStrategies)
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Pre-generating " & numIterations & " random orders..."
    Application.ScreenUpdating = False
    
    For iteration = 1 To numIterations
        ' Generate random order using existing function or inline method
        Call GenerateRandomOrderInline(numStrategies, allRandomOrders, iteration)
    Next iteration
    
    ' Initialize contribution matrix
    ReDim contributionMatrix(1 To numStrategies, 1 To numIterations, 1 To 7)
    
    ' OPTIMIZATION 3: Ultra-fast batch processing
    Dim progressUpdate As Long
    progressUpdate = IIf(numIterations >= 100, 10, 5)
    
    For iteration = 1 To numIterations
        If iteration Mod progressUpdate = 0 Or iteration = 1 Then
            Application.ScreenUpdating = True
            Application.StatusBar = "Processing iteration " & iteration & " of " & numIterations & " (" & Format(iteration / numIterations, "0%") & ")"
            Application.ScreenUpdating = False
        End If
        
        If iteration Mod 25 = 0 Then
            DoEvents  ' Allow other processes
        End If
        
        ' Process entire iteration efficiently
        Call ProcessIterationFast(pnlResults, allRandomOrders, iteration, numStrategies, numDays, _
                                 startingEquity, individualMetrics, contributionMatrix)
    Next iteration
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Calculating median results..."
    Application.ScreenUpdating = False
    
    ' Calculate median contributions
    ReDim medianContributions(1 To numStrategies, 1 To 7)
    For i = 1 To numStrategies
        For j = 1 To 7
            medianContributions(i, j) = CalculateMedianFast(contributionMatrix, i, j, numIterations)
        Next j
    Next i
    
    ' Prepare results array
    ReDim metrics(1 To numStrategies, 1 To 12)
    For i = 1 To numStrategies
        metrics(i, 1) = i
        metrics(i, 2) = strategyNumbers(i)
        metrics(i, 3) = strategyNames(i)
        metrics(i, 4) = medianContributions(i, 1)  ' Return Impact
        metrics(i, 5) = medianContributions(i, 2)  ' P/MaxDD Contribution
        metrics(i, 6) = medianContributions(i, 3)  ' P/AvgDD Contribution
        metrics(i, 7) = medianContributions(i, 4)  ' P/Stdev Contribution
        metrics(i, 8) = Symbols(i)
        metrics(i, 9) = medianContributions(i, 5)  ' Actual P/MaxDD
        metrics(i, 10) = medianContributions(i, 6) ' Actual P/AvgDD
        metrics(i, 11) = medianContributions(i, 7) ' Actual P/Stdev
    Next i
    
    ' Sort and output
    Call SortDiversificationMetricsWithReturn(metrics, sortingMetric, numStrategies)
    
    Dim rankReturnImpact() As Long, rankProfitMaxDD() As Long, rankProfitAvgDD() As Long, rankProfitStdev() As Long
    rankReturnImpact = RankStrategiesByMetric(metrics, 4)
    rankProfitMaxDD = RankStrategiesByMetric(metrics, 5)
    rankProfitAvgDD = RankStrategiesByMetric(metrics, 6)
    rankProfitStdev = RankStrategiesByMetric(metrics, 7)
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Creating output..."
    
    Call OutputDiversificationResultsWithReturn(wsDiversification, metrics, rankReturnImpact, rankProfitMaxDD, rankProfitAvgDD, rankProfitStdev, _
                                    tableStartRow, tableStartCol, chartStartRow, chartStartCol, MCTradeType, numStrategies, numIterations, sortingMetric)
    
    GoTo Cleanup

ErrorHandler:
    MsgBox "Error in diversification analysis: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
End Sub

Sub CalculateAllStrategyMetricsFast(pnlResults As Variant, strategyIndex As Long, startingEquity As Double, _
                                   ByRef individualMetrics() As Double, outputRow As Long)
    ' Calculate ALL metrics for a single strategy in ONE PASS
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim equity As Double, peakEquity As Double, maxDrawdown As Double
    Dim periodPnL As Double, totalReturn As Double
    Dim drawdownSum As Double, drawdownCount As Long
    Dim sum As Double, sumSquares As Double, mean As Double
    Dim numPeriods As Long
    
    numPeriods = UBound(pnlResults, 1)
    
    ' Initialize
    equity = startingEquity
    peakEquity = startingEquity
    maxDrawdown = 0
    drawdownSum = 0
    drawdownCount = 0
    sum = 0
    sumSquares = 0
    
    ' Single pass through data
    For i = 1 To numPeriods
        periodPnL = pnlResults(i, strategyIndex)
        
        ' Running totals
        sum = sum + periodPnL
        sumSquares = sumSquares + periodPnL * periodPnL
        
        ' Equity tracking
        equity = equity + periodPnL
        
        If equity > peakEquity Then
            peakEquity = equity
        End If
        
        Dim currentDrawdown As Double
        currentDrawdown = peakEquity - equity
        
        If currentDrawdown > maxDrawdown Then
            maxDrawdown = currentDrawdown
        End If
        
        If currentDrawdown > 0 Then
            drawdownSum = drawdownSum + currentDrawdown
            drawdownCount = drawdownCount + 1
        End If
    Next i
    
    ' Calculate final metrics
    totalReturn = sum
    mean = sum / numPeriods
    
    Dim variance As Double, stdev As Double, avgDrawdown As Double
    If numPeriods > 1 Then
        variance = (sumSquares - sum * mean) / (numPeriods - 1)
        stdev = Sqr(Abs(variance))  ' Abs to handle potential rounding errors
    Else
        stdev = 0
    End If
    
    If drawdownCount > 0 Then
        avgDrawdown = drawdownSum / drawdownCount
    Else
        avgDrawdown = 0
    End If
    
    ' Store results
    individualMetrics(outputRow, 1) = totalReturn
    individualMetrics(outputRow, 2) = maxDrawdown
    individualMetrics(outputRow, 3) = avgDrawdown
    individualMetrics(outputRow, 4) = stdev
    
    ' Calculate ratios with error checking
    If maxDrawdown > 0.01 Then
        individualMetrics(outputRow, 5) = totalReturn / maxDrawdown
    Else
        individualMetrics(outputRow, 5) = IIf(totalReturn > 0, totalReturn * 1000, 0)
    End If
    
    If avgDrawdown > 0.01 Then
        individualMetrics(outputRow, 6) = totalReturn / avgDrawdown
    Else
        individualMetrics(outputRow, 6) = IIf(totalReturn > 0, totalReturn * 1000, 0)
    End If
    
    If stdev > 0.01 Then
        individualMetrics(outputRow, 7) = totalReturn / stdev
    Else
        individualMetrics(outputRow, 7) = totalReturn
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Fill with safe defaults
    individualMetrics(outputRow, 1) = 0
    individualMetrics(outputRow, 2) = 0
    individualMetrics(outputRow, 3) = 0
    individualMetrics(outputRow, 4) = 0
    individualMetrics(outputRow, 5) = 0
    individualMetrics(outputRow, 6) = 0
    individualMetrics(outputRow, 7) = 0
End Sub

Sub GenerateRandomOrderInline(numStrategies As Long, ByRef allRandomOrders() As Long, iteration As Long)
    ' Generate random order for specific iteration
    Dim i As Long, j As Long, temp As Long
    
    ' Initialize with sequential order
    For i = 1 To numStrategies
        allRandomOrders(iteration, i) = i
    Next i
    
    ' Fisher-Yates shuffle
    For i = numStrategies To 2 Step -1
        j = Int(Rnd() * i) + 1
        temp = allRandomOrders(iteration, i)
        allRandomOrders(iteration, i) = allRandomOrders(iteration, j)
        allRandomOrders(iteration, j) = temp
    Next i
End Sub

Sub ProcessIterationFast(pnlResults As Variant, allRandomOrders() As Long, iteration As Long, _
                        numStrategies As Long, numDays As Long, startingEquity As Double, _
                        individualMetrics() As Double, ByRef contributionMatrix() As Double)
    ' Process entire iteration efficiently
    On Error GoTo ErrorHandler
    
    Dim i As Long, currentStrategy As Long
    Dim portfolioMetrics(1 To 2, 1 To 4) As Double
    
    For i = 1 To numStrategies
        currentStrategy = allRandomOrders(iteration, i)
        
        If i = 1 Then
            ' First strategy: use pre-calculated metrics
            contributionMatrix(currentStrategy, iteration, 1) = 0  ' Return Impact
            contributionMatrix(currentStrategy, iteration, 2) = 0  ' P/MaxDD Contribution
            contributionMatrix(currentStrategy, iteration, 3) = 0  ' P/AvgDD Contribution
            contributionMatrix(currentStrategy, iteration, 4) = 0  ' P/Stdev Contribution
            contributionMatrix(currentStrategy, iteration, 5) = individualMetrics(currentStrategy, 5)
            contributionMatrix(currentStrategy, iteration, 6) = individualMetrics(currentStrategy, 6)
            contributionMatrix(currentStrategy, iteration, 7) = individualMetrics(currentStrategy, 7)
        Else
            ' Calculate both portfolios in single pass
            Call CalculateBothPortfoliosFast(pnlResults, allRandomOrders, iteration, i - 1, i, _
                                           numDays, startingEquity, portfolioMetrics)
            
            ' Extract metrics
            Dim oldReturn As Double, oldMaxDD As Double, oldAvgDD As Double, oldStdev As Double
            Dim newReturn As Double, newMaxDD As Double, newAvgDD As Double, newStdev As Double
            
            oldReturn = portfolioMetrics(1, 1)
            oldMaxDD = portfolioMetrics(1, 2)
            oldAvgDD = portfolioMetrics(1, 3)
            oldStdev = portfolioMetrics(1, 4)
            
            newReturn = portfolioMetrics(2, 1)
            newMaxDD = portfolioMetrics(2, 2)
            newAvgDD = portfolioMetrics(2, 3)
            newStdev = portfolioMetrics(2, 4)
            
            ' Calculate contributions safely
            ' Return Impact
            If Abs(oldReturn) > 0.00001 Then
                contributionMatrix(currentStrategy, iteration, 1) = (newReturn - oldReturn) / oldReturn
            Else
                contributionMatrix(currentStrategy, iteration, 1) = 0
            End If
            
            ' P/MaxDD Contribution
            Dim oldPMaxDD As Double, newPMaxDD As Double, maxDDContrib As Double
            oldPMaxDD = IIf(oldMaxDD > 0.01, oldReturn / oldMaxDD, IIf(oldReturn > 0, oldReturn * 1000, 0))
            newPMaxDD = IIf(newMaxDD > 0.01, newReturn / newMaxDD, IIf(newReturn > 0, newReturn * 1000, 0))
            
            If Abs(oldPMaxDD) > 0.001 Then
                maxDDContrib = (newPMaxDD - oldPMaxDD) / Abs(oldPMaxDD)
                contributionMatrix(currentStrategy, iteration, 2) = IIf(Abs(maxDDContrib) < 0.005, 0, maxDDContrib)
            Else
                contributionMatrix(currentStrategy, iteration, 2) = 0
            End If
            
            ' P/AvgDD Contribution
            Dim oldPAvgDD As Double, newPAvgDD As Double, avgDDContrib As Double
            oldPAvgDD = IIf(oldAvgDD > 0.01, oldReturn / oldAvgDD, IIf(oldReturn > 0, oldReturn * 1000, 0))
            newPAvgDD = IIf(newAvgDD > 0.01, newReturn / newAvgDD, IIf(newReturn > 0, newReturn * 1000, 0))
            
            If Abs(oldPAvgDD) > 0.001 Then
                avgDDContrib = (newPAvgDD - oldPAvgDD) / Abs(oldPAvgDD)
                contributionMatrix(currentStrategy, iteration, 3) = IIf(Abs(avgDDContrib) < 0.005, 0, avgDDContrib)
            Else
                contributionMatrix(currentStrategy, iteration, 3) = 0
            End If
            
            ' P/Stdev Contribution
            Dim oldPStdev As Double, newPStdev As Double, stdevContrib As Double
            oldPStdev = IIf(oldStdev > 0.01, oldReturn / oldStdev, oldReturn)
            newPStdev = IIf(newStdev > 0.01, newReturn / newStdev, newReturn)
            
            If Abs(oldPStdev) > 0.001 Then
                stdevContrib = (newPStdev - oldPStdev) / Abs(oldPStdev)
                contributionMatrix(currentStrategy, iteration, 4) = IIf(Abs(stdevContrib) < 0.005, 0, stdevContrib)
            Else
                contributionMatrix(currentStrategy, iteration, 4) = 0
            End If
            
            ' Store actual ratios
            contributionMatrix(currentStrategy, iteration, 5) = newPMaxDD
            contributionMatrix(currentStrategy, iteration, 6) = newPAvgDD
            contributionMatrix(currentStrategy, iteration, 7) = newPStdev
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    ' Continue with next strategy if error occurs
    Resume Next
End Sub

Sub CalculateBothPortfoliosFast(pnlResults As Variant, allRandomOrders() As Long, iteration As Long, _
                               oldSize As Long, newSize As Long, numDays As Long, _
                               startingEquity As Double, ByRef portfolioMetrics() As Double)
    ' Calculate metrics for both portfolios in single pass
    On Error GoTo ErrorHandler
    
    Dim day As Long, strategy As Long
    Dim oldDayPnL As Double, newDayPnL As Double
    
    ' Portfolio tracking variables
    Dim oldEquity As Double, oldPeakEquity As Double, oldMaxDD As Double
    Dim oldSum As Double, oldSumSquares As Double, oldDrawdownSum As Double, oldDrawdownCount As Long
    Dim newEquity As Double, newPeakEquity As Double, newMaxDD As Double
    Dim newSum As Double, newSumSquares As Double, newDrawdownSum As Double, newDrawdownCount As Long
    
    ' Initialize
    oldEquity = startingEquity: oldPeakEquity = startingEquity: oldMaxDD = 0
    oldSum = 0: oldSumSquares = 0: oldDrawdownSum = 0: oldDrawdownCount = 0
    newEquity = startingEquity: newPeakEquity = startingEquity: newMaxDD = 0
    newSum = 0: newSumSquares = 0: newDrawdownSum = 0: newDrawdownCount = 0
    
    ' Single pass through all days
    For day = 1 To numDays
        oldDayPnL = 0
        newDayPnL = 0
        
        ' Calculate old portfolio P&L
        For strategy = 1 To oldSize
            oldDayPnL = oldDayPnL + pnlResults(day, allRandomOrders(iteration, strategy))
        Next strategy
        
        ' Calculate new portfolio P&L (includes one additional strategy)
        For strategy = 1 To newSize
            newDayPnL = newDayPnL + pnlResults(day, allRandomOrders(iteration, strategy))
        Next strategy
        
        ' Update old portfolio metrics
        oldSum = oldSum + oldDayPnL
        oldSumSquares = oldSumSquares + oldDayPnL * oldDayPnL
        oldEquity = oldEquity + oldDayPnL
        
        If oldEquity > oldPeakEquity Then oldPeakEquity = oldEquity
        Dim oldCurrentDD As Double
        oldCurrentDD = oldPeakEquity - oldEquity
        If oldCurrentDD > oldMaxDD Then oldMaxDD = oldCurrentDD
        If oldCurrentDD > 0 Then
            oldDrawdownSum = oldDrawdownSum + oldCurrentDD
            oldDrawdownCount = oldDrawdownCount + 1
        End If
        
        ' Update new portfolio metrics
        newSum = newSum + newDayPnL
        newSumSquares = newSumSquares + newDayPnL * newDayPnL
        newEquity = newEquity + newDayPnL
        
        If newEquity > newPeakEquity Then newPeakEquity = newEquity
        Dim newCurrentDD As Double
        newCurrentDD = newPeakEquity - newEquity
        If newCurrentDD > newMaxDD Then newMaxDD = newCurrentDD
        If newCurrentDD > 0 Then
            newDrawdownSum = newDrawdownSum + newCurrentDD
            newDrawdownCount = newDrawdownCount + 1
        End If
    Next day
    
    ' Calculate final metrics
    portfolioMetrics(1, 1) = oldSum
    portfolioMetrics(1, 2) = oldMaxDD
    portfolioMetrics(1, 3) = IIf(oldDrawdownCount > 0, oldDrawdownSum / oldDrawdownCount, 0)
    
    If numDays > 1 Then
        Dim oldMean As Double, oldVariance As Double
        oldMean = oldSum / numDays
        oldVariance = (oldSumSquares - oldSum * oldMean) / (numDays - 1)
        portfolioMetrics(1, 4) = Sqr(Abs(oldVariance))
    Else
        portfolioMetrics(1, 4) = 0
    End If
    
    portfolioMetrics(2, 1) = newSum
    portfolioMetrics(2, 2) = newMaxDD
    portfolioMetrics(2, 3) = IIf(newDrawdownCount > 0, newDrawdownSum / newDrawdownCount, 0)
    
    If numDays > 1 Then
        Dim newMean As Double, newVariance As Double
        newMean = newSum / numDays
        newVariance = (newSumSquares - newSum * newMean) / (numDays - 1)
        portfolioMetrics(2, 4) = Sqr(Abs(newVariance))
    Else
        portfolioMetrics(2, 4) = 0
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Return safe defaults
    portfolioMetrics(1, 1) = 0: portfolioMetrics(1, 2) = 0: portfolioMetrics(1, 3) = 0: portfolioMetrics(1, 4) = 0
    portfolioMetrics(2, 1) = 0: portfolioMetrics(2, 2) = 0: portfolioMetrics(2, 3) = 0: portfolioMetrics(2, 4) = 0
End Sub

Function CalculateMedianFast(contributionMatrix() As Double, strategyIndex As Long, metricIndex As Long, numIterations As Long) As Double
    ' Fast median calculation with error handling
    On Error GoTo ErrorHandler
    
    Dim values() As Double
    Dim i As Long
    
    ReDim values(1 To numIterations)
    For i = 1 To numIterations
        values(i) = contributionMatrix(strategyIndex, i, metricIndex)
    Next i
    
    CalculateMedianFast = Application.WorksheetFunction.Median(values)
    Exit Function
    
ErrorHandler:
    CalculateMedianFast = 0
End Function





Sub CalculatePortfolioMetricsPair(pnlResults As Variant, allRandomOrders() As Long, iteration As Long, _
                                 oldPortfolioSize As Long, newPortfolioSize As Long, numDays As Long, _
                                 startingEquity As Double, ByRef portfolioMetrics() As Double)
    ' Calculate metrics for BOTH old and new portfolios in SINGLE PASS - ultimate optimization
    
    Dim day As Long, strategy As Long
    Dim oldDayPnL As Double, newDayPnL As Double
    
    ' Old portfolio tracking
    Dim oldEquity As Double, oldPeakEquity As Double, oldMaxDD As Double
    Dim oldSum As Double, oldSumSquares As Double, oldDrawdownSum As Double, oldDrawdownCount As Long
    
    ' New portfolio tracking
    Dim newEquity As Double, newPeakEquity As Double, newMaxDD As Double
    Dim newSum As Double, newSumSquares As Double, newDrawdownSum As Double, newDrawdownCount As Long
    
    ' Initialize
    oldEquity = startingEquity: oldPeakEquity = startingEquity: oldMaxDD = 0
    oldSum = 0: oldSumSquares = 0: oldDrawdownSum = 0: oldDrawdownCount = 0
    
    newEquity = startingEquity: newPeakEquity = startingEquity: newMaxDD = 0
    newSum = 0: newSumSquares = 0: newDrawdownSum = 0: newDrawdownCount = 0
    
    ' SINGLE PASS: Calculate ALL metrics for BOTH portfolios simultaneously
    For day = 1 To numDays
        oldDayPnL = 0
        newDayPnL = 0
        
        ' Calculate day P&L for old portfolio (first oldPortfolioSize strategies)
        For strategy = 1 To oldPortfolioSize
            oldDayPnL = oldDayPnL + pnlResults(day, allRandomOrders(iteration, strategy))
        Next strategy
        
        ' Calculate day P&L for new portfolio (first newPortfolioSize strategies)
        For strategy = 1 To newPortfolioSize
            newDayPnL = newDayPnL + pnlResults(day, allRandomOrders(iteration, strategy))
        Next strategy
        
        ' Update old portfolio metrics
        oldSum = oldSum + oldDayPnL
        oldSumSquares = oldSumSquares + oldDayPnL * oldDayPnL
        oldEquity = oldEquity + oldDayPnL
        
        If oldEquity > oldPeakEquity Then oldPeakEquity = oldEquity
        Dim oldCurrentDD As Double
        oldCurrentDD = oldPeakEquity - oldEquity
        If oldCurrentDD > oldMaxDD Then oldMaxDD = oldCurrentDD
        If oldCurrentDD > 0 Then
            oldDrawdownSum = oldDrawdownSum + oldCurrentDD
            oldDrawdownCount = oldDrawdownCount + 1
        End If
        
        ' Update new portfolio metrics
        newSum = newSum + newDayPnL
        newSumSquares = newSumSquares + newDayPnL * newDayPnL
        newEquity = newEquity + newDayPnL
        
        If newEquity > newPeakEquity Then newPeakEquity = newEquity
        Dim newCurrentDD As Double
        newCurrentDD = newPeakEquity - newEquity
        If newCurrentDD > newMaxDD Then newMaxDD = newCurrentDD
        If newCurrentDD > 0 Then
            newDrawdownSum = newDrawdownSum + newCurrentDD
            newDrawdownCount = newDrawdownCount + 1
        End If
    Next day
    
    ' Calculate final metrics for old portfolio
    portfolioMetrics(1, 1) = oldSum ' Return
    portfolioMetrics(1, 2) = oldMaxDD ' Max DD
    portfolioMetrics(1, 3) = IIf(oldDrawdownCount > 0, oldDrawdownSum / oldDrawdownCount, 0) ' Avg DD
    
    Dim oldMean As Double, oldVariance As Double
    oldMean = oldSum / numDays
    oldVariance = (oldSumSquares - oldSum * oldMean) / (numDays - 1)
    portfolioMetrics(1, 4) = Sqr(oldVariance) ' Stdev
    
    ' Calculate final metrics for new portfolio
    portfolioMetrics(2, 1) = newSum ' Return
    portfolioMetrics(2, 2) = newMaxDD ' Max DD
    portfolioMetrics(2, 3) = IIf(newDrawdownCount > 0, newDrawdownSum / newDrawdownCount, 0) ' Avg DD
    
    Dim newMean As Double, newVariance As Double
    newMean = newSum / numDays
    newVariance = (newSumSquares - newSum * newMean) / (numDays - 1)
    portfolioMetrics(2, 4) = Sqr(newVariance) ' Stdev
End Sub

' ===== OPTIMIZED HELPER FUNCTIONS THAT REPLICATE EXACT FUNCTIONALITY =====

Function RunChronologicalBacktestSingleStrategy(pnlResults As Variant, strategyIndex As Long, _
                                               requiredMargin As Double, startingEquity As Double, _
                                               AverageTrade As Double) As Variant
    ' EXACT same logic as RunChronologicalBacktestWithTracking but for single strategy
    Dim numPeriods As Long
    Dim i As Long
    Dim equity As Double, peakEquity As Double, maxDrawdown As Double
    Dim periodPnL As Double, totalReturn As Double
    Dim results(1 To 1, 1 To 8) As Double
    Dim drawdownSum As Double, drawdownCount As Long, avgDrawdown As Double
    
    numPeriods = UBound(pnlResults, 1)
    
    ' Initialize
    equity = startingEquity
    peakEquity = startingEquity
    maxDrawdown = 0
    drawdownSum = 0
    drawdownCount = 0
    
    ' Run chronological backtest using ORIGINAL PnL data (EXACT same logic)
    For i = 1 To numPeriods
        ' Get period P&L for this strategy only
        periodPnL = pnlResults(i, strategyIndex)
        
        ' Update equity
        equity = equity + periodPnL
        
        ' Update peak and drawdown tracking - EXACT same logic
        If equity > peakEquity Then
            peakEquity = equity
        End If
        
        ' Calculate current dollar drawdown
        Dim currentDollarDrawdown As Double
        currentDollarDrawdown = peakEquity - equity
        
        ' Track maximum dollar drawdown
        If currentDollarDrawdown > maxDrawdown Then
            maxDrawdown = currentDollarDrawdown
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
    totalReturn = equity - startingEquity
    
    ' Populate results array to match EXACT format
    results(1, 1) = 1
    results(1, 2) = totalReturn
    results(1, 3) = totalReturn
    
    If maxDrawdown > 0.01 Then
        results(1, 4) = totalReturn / maxDrawdown
    Else
        results(1, 4) = totalReturn * 1000
    End If
    
    results(1, 5) = maxDrawdown
    results(1, 6) = 0
    results(1, 7) = numPeriods
    results(1, 8) = avgDrawdown
    
    RunChronologicalBacktestSingleStrategy = results
End Function

Function CalculatePortfolioReturn(pnlResults As Variant, randomOrder() As Long, numStrategies As Long) As Double
    ' Calculate total return for portfolio using EXACT same logic
    Dim i As Long, j As Long, totalReturn As Double
    
    For i = 1 To UBound(pnlResults, 1)
        Dim periodPnL As Double
        periodPnL = 0
        
        ' Sum across selected strategies
        For j = 1 To numStrategies
            periodPnL = periodPnL + pnlResults(i, randomOrder(j))
        Next j
        
        totalReturn = totalReturn + periodPnL
    Next i
    
    CalculatePortfolioReturn = totalReturn
End Function

Function CalculatePortfolioMaxDD(pnlResults As Variant, randomOrder() As Long, numStrategies As Long, startingEquity As Double) As Double
    ' Calculate max drawdown using EXACT same logic as RunChronologicalBacktestWithTracking
    Dim i As Long, j As Long
    Dim equity As Double, peakEquity As Double, maxDrawdown As Double
    
    equity = startingEquity
    peakEquity = startingEquity
    maxDrawdown = 0
    
    For i = 1 To UBound(pnlResults, 1)
        Dim periodPnL As Double
        periodPnL = 0
        
        ' Sum across selected strategies
        For j = 1 To numStrategies
            periodPnL = periodPnL + pnlResults(i, randomOrder(j))
        Next j
        
        equity = equity + periodPnL
        
        If equity > peakEquity Then
            peakEquity = equity
        End If
        
        Dim currentDollarDrawdown As Double
        currentDollarDrawdown = peakEquity - equity
        
        If currentDollarDrawdown > maxDrawdown Then
            maxDrawdown = currentDollarDrawdown
        End If
    Next i
    
    CalculatePortfolioMaxDD = maxDrawdown
End Function

Function CalculatePortfolioAvgDD(pnlResults As Variant, randomOrder() As Long, numStrategies As Long, startingEquity As Double) As Double
    ' Calculate average drawdown using EXACT same logic
    Dim i As Long, j As Long
    Dim equity As Double, peakEquity As Double
    Dim drawdownSum As Double, drawdownCount As Long
    
    equity = startingEquity
    peakEquity = startingEquity
    drawdownSum = 0
    drawdownCount = 0
    
    For i = 1 To UBound(pnlResults, 1)
        Dim periodPnL As Double
        periodPnL = 0
        
        ' Sum across selected strategies
        For j = 1 To numStrategies
            periodPnL = periodPnL + pnlResults(i, randomOrder(j))
        Next j
        
        equity = equity + periodPnL
        
        If equity > peakEquity Then
            peakEquity = equity
        End If
        
        Dim currentDollarDrawdown As Double
        currentDollarDrawdown = peakEquity - equity
        
        If currentDollarDrawdown > 0 Then
            drawdownSum = drawdownSum + currentDollarDrawdown
            drawdownCount = drawdownCount + 1
        End If
    Next i
    
    If drawdownCount > 0 Then
        CalculatePortfolioAvgDD = drawdownSum / drawdownCount
    Else
        CalculatePortfolioAvgDD = 0
    End If
End Function


Function CalculateStrategyStdev(pnlResults As Variant, strategyIndex As Long) As Double
    ' Calculate single strategy standard deviation
    Dim i As Long, sum As Double, sumSquares As Double, n As Long, mean As Double
    
    n = UBound(pnlResults, 1)
    
    For i = 1 To n
        sum = sum + pnlResults(i, strategyIndex)
    Next i
    mean = sum / n
    
    For i = 1 To n
        sumSquares = sumSquares + (pnlResults(i, strategyIndex) - mean) ^ 2
    Next i
    
    CalculateStrategyStdev = Sqr(sumSquares / (n - 1))
End Function


' OPTIMIZED VERSION OF GREEDY SELECTION - Much faster
' REQUIRED: Add this helper function to your module first
Sub UpdateStatusWithRefresh(statusText As String)
    Application.ScreenUpdating = True
    Application.StatusBar = statusText
    Application.ScreenUpdating = False
End Sub

Sub RankStrategiesByGreedySelection(wsDiversification As Worksheet, pnlResults As Variant, _
                                           requiredMargin As Double, startingEquity As Double, _
                                           AverageTrade As Double, MCTradeType As String, _
                                           Optional sortingMetric As String = "1")
    ' ENHANCED greedy selection with comprehensive status updates
    
    Dim numStrategies As Long, numDays As Long
    Dim i As Long, j As Long, K As Long
    Dim strategyNames() As String, Symbols() As String, strategyNumbers() As Long
    Dim metrics() As Variant
    Dim selectionOrder() As Long
    Dim strategySelected() As Boolean
    Dim tableStartRow As Long, tableStartCol As Long
    Dim chartStartRow As Long, chartStartCol As Long
    
    On Error GoTo ErrorHandler
    
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    numStrategies = UBound(pnlResults, 2)
    numDays = UBound(pnlResults, 1)
    tableStartRow = 6: tableStartCol = 12
    chartStartRow = 6: chartStartCol = 2
    
    If numStrategies < 2 Then
        MsgBox "At least two strategies are required.", vbExclamation
        GoTo Cleanup
    End If
    
    ' ENHANCED: Detailed initialization status
    Call UpdateStatusWithRefresh("Initializing greedy selection for " & numStrategies & " strategies...")
    
    ' Initialize arrays
    ReDim strategyNames(1 To numStrategies)
    ReDim strategyNumbers(1 To numStrategies)
    ReDim Symbols(1 To numStrategies)
    ReDim selectionOrder(1 To numStrategies)
    ReDim strategySelected(1 To numStrategies)
    
    Call UpdateStatusWithRefresh("Loading strategy information from Portfolio sheet...")
    
    ' Get strategy information
    For i = 1 To numStrategies
        On Error Resume Next
        strategyNames(i) = ThisWorkbook.Sheets("Portfolio").Cells(i + 1, COL_PORT_STRATEGY_NAME).value
        strategyNumbers(i) = ThisWorkbook.Sheets("Portfolio").Cells(i + 1, COL_PORT_STRATEGY_NUMBER).value
        Symbols(i) = ThisWorkbook.Sheets("Portfolio").Cells(i + 1, COL_PORT_SYMBOL).value
        On Error GoTo ErrorHandler
        
        If strategyNames(i) = "" Then strategyNames(i) = "Strategy " & i
        If strategyNumbers(i) = 0 Then strategyNumbers(i) = i
        If Symbols(i) = "" Then Symbols(i) = "S" & i
        strategySelected(i) = False
    Next i
    
    ' OPTIMIZATION: Pre-calculate individual strategy metrics
    Dim individualMetrics() As Double ' Return, MaxDD, AvgDD, Stdev, P/MaxDD
    ReDim individualMetrics(1 To numStrategies, 1 To 5)
    
    Call UpdateStatusWithRefresh("Pre-calculating individual strategy metrics for " & numStrategies & " strategies...")
    
    For i = 1 To numStrategies
        ' ENHANCED: Progress updates during individual calculations
        If i Mod 3 = 0 Or i = 1 Or i = numStrategies Then
            Call UpdateStatusWithRefresh("Pre-calculating metrics: strategy " & i & " of " & numStrategies & " (" & Format(i / numStrategies, "0%") & ")")
        End If
        
        Call CalculateIndividualStrategyMetrics(pnlResults, i, startingEquity, individualMetrics, i)
    Next i
    
    ' Step 1: Find best individual strategy
    Dim bestStrategy As Long, bestRatio As Double
    bestStrategy = 1: bestRatio = individualMetrics(1, 5)
    
    Call UpdateStatusWithRefresh("Step 1: Finding best individual strategy...")
    
    For i = 2 To numStrategies
        If individualMetrics(i, 5) > bestRatio Then
            bestRatio = individualMetrics(i, 5)
            bestStrategy = i
        End If
    Next i
    
    strategySelected(bestStrategy) = True
    selectionOrder(1) = bestStrategy
    
    Call UpdateStatusWithRefresh("Step 1 complete: Selected strategy " & bestStrategy & " (" & strategyNames(bestStrategy) & ") as foundation")
    
    ' Step 2: Greedy selection using optimized calculations
    For selectionStep = 2 To numStrategies
        ' ENHANCED: More detailed status with percentage and step info
        Call UpdateStatusWithRefresh("Step 2: Greedy selection step " & selectionStep & " of " & numStrategies & " (" & Format((selectionStep - 1) / (numStrategies - 1), "0%") & " complete)")
        
        If selectionStep Mod 3 = 0 Then DoEvents
        
        Dim bestNextStrategy As Long, bestImprovement As Double
        bestNextStrategy = -1: bestImprovement = -999999
        
        ' Calculate current portfolio metrics efficiently
        Dim currentMetrics(1 To 4) As Double ' Return, MaxDD, AvgDD, Stdev
        Call CalculateCurrentPortfolioMetrics(pnlResults, strategySelected, startingEquity, currentMetrics)
        
        Dim currentRatio As Double
        currentRatio = IIf(currentMetrics(2) > 0.01, currentMetrics(1) / currentMetrics(2), currentMetrics(1) * 1000)
        
        ' Count remaining candidates for better progress tracking
        Dim candidatesRemaining As Long, candidatesTested As Long
        candidatesRemaining = 0
        For i = 1 To numStrategies
            If Not strategySelected(i) Then candidatesRemaining = candidatesRemaining + 1
        Next i
        
        candidatesTested = 0
        
        ' Test each unselected strategy
        For i = 1 To numStrategies
            If Not strategySelected(i) Then
                candidatesTested = candidatesTested + 1
                
                ' ENHANCED: Progress updates during candidate testing
                If candidatesTested Mod 2 = 0 Or candidatesTested = 1 Or candidatesTested = candidatesRemaining Then
                    Call UpdateStatusWithRefresh("Step " & selectionStep & ": Testing candidate " & candidatesTested & " of " & candidatesRemaining & " (strategy " & i & ")")
                End If
                
                ' Calculate new portfolio metrics with strategy i added
                Dim newMetrics(1 To 4) As Double
                Call CalculatePortfolioMetricsWithAddedStrategy(pnlResults, strategySelected, i, startingEquity, newMetrics)
                
                Dim newRatio As Double
                newRatio = IIf(newMetrics(2) > 0.01, newMetrics(1) / newMetrics(2), newMetrics(1) * 1000)
                
                ' Calculate improvement
                Dim improvement As Double
                If Abs(currentRatio) > 0.001 Then
                    improvement = (newRatio - currentRatio) / Abs(currentRatio)
                Else
                    improvement = newRatio - currentRatio
                End If
                
                If improvement > bestImprovement Then
                    bestImprovement = improvement
                    bestNextStrategy = i
                End If
            End If
        Next i
        
        ' Add best strategy found
        If bestNextStrategy > 0 Then
            strategySelected(bestNextStrategy) = True
            selectionOrder(selectionStep) = bestNextStrategy
            Call UpdateStatusWithRefresh("Step " & selectionStep & ": Selected strategy " & bestNextStrategy & " (" & strategyNames(bestNextStrategy) & ") - improvement: " & Format(bestImprovement, "0.0%"))
        Else
            ' Fallback: add first unselected strategy
            For i = 1 To numStrategies
                If Not strategySelected(i) Then
                    strategySelected(i) = True
                    selectionOrder(selectionStep) = i
                    Call UpdateStatusWithRefresh("Step " & selectionStep & ": Added strategy " & i & " (" & strategyNames(i) & ") - no improvement found")
                    Exit For
                End If
            Next i
        End If
    Next selectionStep
    
    ' Step 3: Calculate final metrics and contributions
    ReDim metrics(1 To numStrategies, 1 To 12)
    
    Call UpdateStatusWithRefresh("Step 3: Calculating final contributions and portfolio metrics...")
    
    Dim Step As Long
    
    For Step = 1 To numStrategies
        ' ENHANCED: Progress updates for final calculations
        If Step Mod 2 = 0 Or Step = 1 Or Step = numStrategies Then
            Call UpdateStatusWithRefresh("Step 3: Calculating metrics for portfolio step " & Step & " of " & numStrategies & " (" & Format(Step / numStrategies, "0%") & ")")
        End If
        
        Dim strategyIndex As Long
        strategyIndex = selectionOrder(Step)
        
        ' Calculate portfolio metrics at this step
        Dim stepMetrics(1 To 4) As Double
        Call CalculateStepPortfolioMetrics(pnlResults, selectionOrder, Step, startingEquity, stepMetrics)
        
        ' Calculate contributions
        Dim contributions(1 To 4) As Double
        If Step = 1 Then
            ' First strategy has zero contributions
            contributions(1) = 0: contributions(2) = 0: contributions(3) = 0: contributions(4) = 0
        Else
            ' Calculate previous step metrics
            Dim prevMetrics(1 To 4) As Double
            Call CalculateStepPortfolioMetrics(pnlResults, selectionOrder, Step - 1, startingEquity, prevMetrics)
            
            ' Calculate contributions
            Call CalculateContributions(prevMetrics, stepMetrics, contributions)
        End If
        
        ' Store results
        metrics(Step, 1) = Step
        metrics(Step, 2) = strategyNumbers(strategyIndex)
        metrics(Step, 3) = strategyNames(strategyIndex)
        metrics(Step, 4) = contributions(1)  ' Return Impact
        metrics(Step, 5) = contributions(2)  ' P/MaxDD contribution
        metrics(Step, 6) = contributions(3)  ' P/AvgDD contribution
        metrics(Step, 7) = contributions(4)  ' P/Stdev contribution
        metrics(Step, 8) = Symbols(strategyIndex)
        
        ' Calculate and store actual ratios
        Dim pMaxDD As Double, pAvgDD As Double, pStdev As Double
        pMaxDD = IIf(stepMetrics(2) > 0.01, stepMetrics(1) / stepMetrics(2), stepMetrics(1))
        pAvgDD = IIf(stepMetrics(3) > 0.01, stepMetrics(1) / stepMetrics(3), stepMetrics(1))
        pStdev = IIf(stepMetrics(4) > 0.01, stepMetrics(1) / stepMetrics(4), stepMetrics(1))
        
        metrics(Step, 9) = pMaxDD
        metrics(Step, 10) = pAvgDD
        metrics(Step, 11) = pStdev
    Next Step
    
    Call UpdateStatusWithRefresh("Sorting results and calculating rankings...")
    
    ' Sort and calculate ranks
    If sortingMetric <> "1" And sortingMetric <> "" Then
        Call SortDiversificationMetricsWithReturn(metrics, sortingMetric, numStrategies)
    End If
    
    Dim rankReturnImpact() As Long, rankProfitMaxDD() As Long, rankProfitAvgDD() As Long, rankProfitStdev() As Long
    
    If sortingMetric = "1" Or sortingMetric = "" Then
        ' Keep greedy order
        ReDim rankReturnImpact(1 To numStrategies)
        ReDim rankProfitMaxDD(1 To numStrategies)
        ReDim rankProfitAvgDD(1 To numStrategies)
        ReDim rankProfitStdev(1 To numStrategies)
        For i = 1 To numStrategies
            rankReturnImpact(i) = i: rankProfitMaxDD(i) = i: rankProfitAvgDD(i) = i: rankProfitStdev(i) = i
        Next i
    Else
        rankReturnImpact = RankStrategiesByMetric(metrics, 4)
        rankProfitMaxDD = RankStrategiesByMetric(metrics, 5)
        rankProfitAvgDD = RankStrategiesByMetric(metrics, 6)
        rankProfitStdev = RankStrategiesByMetric(metrics, 7)
    End If
    
    ' Output results
    Call UpdateStatusWithRefresh("Creating final output tables and charts...")
    
    Call OutputGreedyResultsWithReturn(wsDiversification, metrics, rankReturnImpact, rankProfitMaxDD, rankProfitAvgDD, rankProfitStdev, _
                            tableStartRow, tableStartCol, chartStartRow, chartStartCol, MCTradeType, numStrategies, sortingMetric)
    
    Call UpdateStatusWithRefresh("Greedy selection analysis complete!")
    
    GoTo Cleanup

ErrorHandler:
    MsgBox "Error in enhanced greedy selection: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
End Sub

' HELPER FUNCTIONS - Add these if you don't already have them

Sub CalculateIndividualStrategyMetrics(pnlResults As Variant, strategyIndex As Long, startingEquity As Double, _
                                      ByRef individualMetrics() As Double, outputRow As Long)
    ' Calculate metrics for single strategy efficiently
    Dim i As Long, numDays As Long
    Dim equity As Double, peakEquity As Double, maxDD As Double
    Dim sum As Double, sumSquares As Double, drawdownSum As Double, drawdownCount As Long
    
    numDays = UBound(pnlResults, 1)
    equity = startingEquity: peakEquity = startingEquity: maxDD = 0
    sum = 0: sumSquares = 0: drawdownSum = 0: drawdownCount = 0
    
    For i = 1 To numDays
        Dim dayPnL As Double
        dayPnL = pnlResults(i, strategyIndex)
        
        sum = sum + dayPnL
        sumSquares = sumSquares + dayPnL * dayPnL
        equity = equity + dayPnL
        
        If equity > peakEquity Then peakEquity = equity
        Dim currentDD As Double
        currentDD = peakEquity - equity
        If currentDD > maxDD Then maxDD = currentDD
        If currentDD > 0 Then
            drawdownSum = drawdownSum + currentDD
            drawdownCount = drawdownCount + 1
        End If
    Next i
    
    Dim totalReturn As Double, avgDD As Double, stdev As Double
    totalReturn = sum
    avgDD = IIf(drawdownCount > 0, drawdownSum / drawdownCount, 0)
    
    If numDays > 1 Then
        Dim mean As Double, variance As Double
        mean = sum / numDays
        variance = (sumSquares - sum * mean) / (numDays - 1)
        stdev = Sqr(Abs(variance))
    Else
        stdev = 0
    End If
    
    individualMetrics(outputRow, 1) = totalReturn
    individualMetrics(outputRow, 2) = maxDD
    individualMetrics(outputRow, 3) = avgDD
    individualMetrics(outputRow, 4) = stdev
    individualMetrics(outputRow, 5) = IIf(maxDD > 0.01, totalReturn / maxDD, IIf(totalReturn > 0, totalReturn * 1000, 0))
End Sub

Sub CalculateCurrentPortfolioMetrics(pnlResults As Variant, strategySelected() As Boolean, startingEquity As Double, _
                                    ByRef currentMetrics() As Double)
    ' Calculate portfolio metrics for currently selected strategies
    Dim i As Long, j As Long, numDays As Long, numStrategies As Long
    Dim equity As Double, peakEquity As Double, maxDD As Double
    Dim sum As Double, sumSquares As Double, drawdownSum As Double, drawdownCount As Long
    
    numDays = UBound(pnlResults, 1)
    numStrategies = UBound(pnlResults, 2)
    equity = startingEquity: peakEquity = startingEquity: maxDD = 0
    sum = 0: sumSquares = 0: drawdownSum = 0: drawdownCount = 0
    
    For i = 1 To numDays
        Dim dayPnL As Double: dayPnL = 0
        For j = 1 To numStrategies
            If strategySelected(j) Then dayPnL = dayPnL + pnlResults(i, j)
        Next j
        
        sum = sum + dayPnL
        sumSquares = sumSquares + dayPnL * dayPnL
        equity = equity + dayPnL
        
        If equity > peakEquity Then peakEquity = equity
        Dim currentDD As Double
        currentDD = peakEquity - equity
        If currentDD > maxDD Then maxDD = currentDD
        If currentDD > 0 Then
            drawdownSum = drawdownSum + currentDD
            drawdownCount = drawdownCount + 1
        End If
    Next i
    
    currentMetrics(1) = sum ' Return
    currentMetrics(2) = maxDD ' Max DD
    currentMetrics(3) = IIf(drawdownCount > 0, drawdownSum / drawdownCount, 0) ' Avg DD
    
    If numDays > 1 Then
        Dim mean As Double, variance As Double
        mean = sum / numDays
        variance = (sumSquares - sum * mean) / (numDays - 1)
        currentMetrics(4) = Sqr(Abs(variance)) ' Stdev
    Else
        currentMetrics(4) = 0
    End If
End Sub

Sub CalculatePortfolioMetricsWithAddedStrategy(pnlResults As Variant, strategySelected() As Boolean, addedStrategy As Long, _
                                              startingEquity As Double, ByRef newMetrics() As Double)
    ' Calculate portfolio metrics with one additional strategy
    Dim i As Long, j As Long, numDays As Long, numStrategies As Long
    Dim equity As Double, peakEquity As Double, maxDD As Double
    Dim sum As Double, sumSquares As Double, drawdownSum As Double, drawdownCount As Long
    
    numDays = UBound(pnlResults, 1)
    numStrategies = UBound(pnlResults, 2)
    equity = startingEquity: peakEquity = startingEquity: maxDD = 0
    sum = 0: sumSquares = 0: drawdownSum = 0: drawdownCount = 0
    
    For i = 1 To numDays
        Dim dayPnL As Double: dayPnL = 0
        For j = 1 To numStrategies
            If strategySelected(j) Or j = addedStrategy Then dayPnL = dayPnL + pnlResults(i, j)
        Next j
        
        sum = sum + dayPnL
        sumSquares = sumSquares + dayPnL * dayPnL
        equity = equity + dayPnL
        
        If equity > peakEquity Then peakEquity = equity
        Dim currentDD As Double
        currentDD = peakEquity - equity
        If currentDD > maxDD Then maxDD = currentDD
        If currentDD > 0 Then
            drawdownSum = drawdownSum + currentDD
            drawdownCount = drawdownCount + 1
        End If
    Next i
    
    newMetrics(1) = sum
    newMetrics(2) = maxDD
    newMetrics(3) = IIf(drawdownCount > 0, drawdownSum / drawdownCount, 0)
    
    If numDays > 1 Then
        Dim mean As Double, variance As Double
        mean = sum / numDays
        variance = (sumSquares - sum * mean) / (numDays - 1)
        newMetrics(4) = Sqr(Abs(variance))
    Else
        newMetrics(4) = 0
    End If
End Sub

Sub CalculateStepPortfolioMetrics(pnlResults As Variant, selectionOrder() As Long, numSteps As Long, _
                                 startingEquity As Double, ByRef stepMetrics() As Double)
    ' Calculate portfolio metrics for first numSteps strategies in selection order
    Dim i As Long, j As Long, K As Long, numDays As Long
    Dim equity As Double, peakEquity As Double, maxDD As Double
    Dim sum As Double, sumSquares As Double, drawdownSum As Double, drawdownCount As Long
    
    numDays = UBound(pnlResults, 1)
    equity = startingEquity: peakEquity = startingEquity: maxDD = 0
    sum = 0: sumSquares = 0: drawdownSum = 0: drawdownCount = 0
    
    For i = 1 To numDays
        Dim dayPnL As Double: dayPnL = 0
        For j = 1 To numSteps
            dayPnL = dayPnL + pnlResults(i, selectionOrder(j))
        Next j
        
        sum = sum + dayPnL
        sumSquares = sumSquares + dayPnL * dayPnL
        equity = equity + dayPnL
        
        If equity > peakEquity Then peakEquity = equity
        Dim currentDD As Double
        currentDD = peakEquity - equity
        If currentDD > maxDD Then maxDD = currentDD
        If currentDD > 0 Then
            drawdownSum = drawdownSum + currentDD
            drawdownCount = drawdownCount + 1
        End If
    Next i
    
    stepMetrics(1) = sum
    stepMetrics(2) = maxDD
    stepMetrics(3) = IIf(drawdownCount > 0, drawdownSum / drawdownCount, 0)
    
    If numDays > 1 Then
        Dim mean As Double, variance As Double
        mean = sum / numDays
        variance = (sumSquares - sum * mean) / (numDays - 1)
        stepMetrics(4) = Sqr(Abs(variance))
    Else
        stepMetrics(4) = 0
    End If
End Sub

Sub CalculateContributions(prevMetrics() As Double, stepMetrics() As Double, ByRef contributions() As Double)
    ' Calculate diversification contributions
    Dim prevReturn As Double, stepReturn As Double
    Dim prevMaxDD As Double, stepMaxDD As Double
    Dim prevAvgDD As Double, stepAvgDD As Double
    Dim prevStdev As Double, stepStdev As Double
    
    prevReturn = prevMetrics(1): stepReturn = stepMetrics(1)
    prevMaxDD = prevMetrics(2): stepMaxDD = stepMetrics(2)
    prevAvgDD = prevMetrics(3): stepAvgDD = stepMetrics(3)
    prevStdev = prevMetrics(4): stepStdev = stepMetrics(4)
    
    ' Return Impact
    contributions(1) = IIf(Abs(prevReturn) > 0.00001, (stepReturn - prevReturn) / prevReturn, 0)
    If Abs(contributions(1)) < 0.005 Then contributions(1) = 0
    
    ' P/MaxDD Contribution
    Dim prevPMaxDD As Double, stepPMaxDD As Double
    prevPMaxDD = IIf(prevMaxDD > 0.01, prevReturn / prevMaxDD, IIf(prevReturn > 0, prevReturn * 1000, 0))
    stepPMaxDD = IIf(stepMaxDD > 0.01, stepReturn / stepMaxDD, IIf(stepReturn > 0, stepReturn * 1000, 0))
    
    If Abs(prevPMaxDD) > 0.001 Then
        contributions(2) = (stepPMaxDD - prevPMaxDD) / Abs(prevPMaxDD)
        If Abs(contributions(2)) < 0.005 Then contributions(2) = 0
    Else
        contributions(2) = 0
    End If
    
    ' P/AvgDD Contribution
    Dim prevPAvgDD As Double, stepPAvgDD As Double
    prevPAvgDD = IIf(prevAvgDD > 0.01, prevReturn / prevAvgDD, IIf(prevReturn > 0, prevReturn * 1000, 0))
    stepPAvgDD = IIf(stepAvgDD > 0.01, stepReturn / stepAvgDD, IIf(stepReturn > 0, stepReturn * 1000, 0))
    
    If Abs(prevPAvgDD) > 0.001 Then
        contributions(3) = (stepPAvgDD - prevPAvgDD) / Abs(prevPAvgDD)
        If Abs(contributions(3)) < 0.005 Then contributions(3) = 0
    Else
        contributions(3) = 0
    End If
    
    ' P/Stdev Contribution
    Dim prevPStdev As Double, stepPStdev As Double
    prevPStdev = IIf(prevStdev > 0.01, prevReturn / prevStdev, prevReturn)
    stepPStdev = IIf(stepStdev > 0.01, stepReturn / stepStdev, stepReturn)
    
    If Abs(prevPStdev) > 0.001 Then
        contributions(4) = (stepPStdev - prevPStdev) / Abs(prevPStdev)
        If Abs(contributions(4)) < 0.005 Then contributions(4) = 0
    Else
        contributions(4) = 0
    End If
End Sub


Function CalculatePortfolioStdev(portfolioData As Variant) As Double
    ' Calculate standard deviation for portfolio data
    ' This version works with portfolio arrays where unused strategies have 0 values
    
    Dim i As Long, j As Long
    Dim dayReturns() As Double
    Dim sum As Double, sumSquares As Double, mean As Double, variance As Double
    Dim numDays As Long, numStrategies As Long
    Dim dayPnL As Double
    
    On Error GoTo ErrorHandler
    
    numDays = UBound(portfolioData, 1)
    numStrategies = UBound(portfolioData, 2)
    
    ReDim dayReturns(1 To numDays)
    
    ' Calculate daily P&L for the portfolio
    For i = 1 To numDays
        dayPnL = 0
        For j = 1 To numStrategies
            dayPnL = dayPnL + portfolioData(i, j)
        Next j
        dayReturns(i) = dayPnL
        sum = sum + dayPnL
    Next i
    
    ' Calculate mean
    mean = sum / numDays
    
    ' Calculate variance
    For i = 1 To numDays
        sumSquares = sumSquares + (dayReturns(i) - mean) ^ 2
    Next i
    
    If numDays > 1 Then
        variance = sumSquares / (numDays - 1)
        CalculatePortfolioStdev = Sqr(Abs(variance))  ' Abs to handle rounding errors
    Else
        CalculatePortfolioStdev = 0
    End If
    
    Exit Function
    
ErrorHandler:
    CalculatePortfolioStdev = 0
End Function

Function CalculateMedianReturn(results As Variant) As Double
    ' Calculate median return from backtest results
    ' Assumes results array has return data in column 2
    
    On Error GoTo ErrorHandler
    
    If UBound(results, 1) = 1 Then
        ' Single result
        CalculateMedianReturn = results(1, 2)
    Else
        ' Multiple results - use median
        Dim returns() As Double
        Dim i As Long
        
        ReDim returns(1 To UBound(results, 1))
        For i = 1 To UBound(results, 1)
            returns(i) = results(i, 2)
        Next i
        
        CalculateMedianReturn = Application.WorksheetFunction.Median(returns)
    End If
    
    Exit Function
    
ErrorHandler:
    CalculateMedianReturn = 0
End Function

Function CalculateMedianMaxDrawdown(results As Variant) As Double
    ' Calculate median max drawdown from backtest results
    ' Assumes results array has max drawdown data in column 5
    
    On Error GoTo ErrorHandler
    
    If UBound(results, 1) = 1 Then
        ' Single result
        CalculateMedianMaxDrawdown = results(1, 5)
    Else
        ' Multiple results - use median
        Dim drawdowns() As Double
        Dim i As Long
        
        ReDim drawdowns(1 To UBound(results, 1))
        For i = 1 To UBound(results, 1)
            drawdowns(i) = results(i, 5)
        Next i
        
        CalculateMedianMaxDrawdown = Application.WorksheetFunction.Median(drawdowns)
    End If
    
    Exit Function
    
ErrorHandler:
    CalculateMedianMaxDrawdown = 0
End Function

Function CalculateMedianAvgDrawdown(results As Variant) As Double
    ' Calculate median average drawdown from backtest results
    ' Assumes results array has avg drawdown data in column 8
    
    On Error GoTo ErrorHandler
    
    If UBound(results, 1) = 1 Then
        ' Single result
        CalculateMedianAvgDrawdown = results(1, 8)
    Else
        ' Multiple results - use median
        Dim avgDrawdowns() As Double
        Dim i As Long
        
        ReDim avgDrawdowns(1 To UBound(results, 1))
        For i = 1 To UBound(results, 1)
            avgDrawdowns(i) = results(i, 8)
        Next i
        
        CalculateMedianAvgDrawdown = Application.WorksheetFunction.Median(avgDrawdowns)
    End If
    
    Exit Function
    
ErrorHandler:
    CalculateMedianAvgDrawdown = 0
End Function

Function RankStrategiesByMetric(metrics As Variant, metricColumn As Long) As Long()
    ' Rank strategies by a specific metric column
    ' Returns array of ranks (1 = best)
    
    Dim numStrategies As Long
    Dim ranks() As Long
    Dim values() As Double
    Dim i As Long, j As Long, rank As Long
    
    On Error GoTo ErrorHandler
    
    numStrategies = UBound(metrics, 1)
    ReDim ranks(1 To numStrategies)
    ReDim values(1 To numStrategies)
    
    ' Extract values
    For i = 1 To numStrategies
        If IsNumeric(metrics(i, metricColumn)) Then
            values(i) = CDbl(metrics(i, metricColumn))
        Else
            values(i) = 0  ' Handle null/empty values
        End If
    Next i
    
    ' Calculate ranks (1 = highest value)
    For i = 1 To numStrategies
        rank = 1
        For j = 1 To numStrategies
            If j <> i And values(j) > values(i) Then
                rank = rank + 1
            End If
        Next j
        ranks(i) = rank
    Next i
    
    RankStrategiesByMetric = ranks
    Exit Function
    
ErrorHandler:
    ReDim ranks(1 To numStrategies)
    For i = 1 To numStrategies
        ranks(i) = i  ' Default sequential ranking
    Next i
    RankStrategiesByMetric = ranks
End Function








Sub OutputGreedyResultsWithReturn(ws As Worksheet, metrics As Variant, rankReturnImpact() As Long, _
                       rankProfitMaxDD() As Long, rankProfitAvgDD() As Long, rankProfitStdev() As Long, _
                       tableStartRow As Long, tableStartCol As Long, _
                       chartStartRow As Long, chartStartCol As Long, _
                       MCTradeType As String, numStrategies As Long, _
                       sortingMetric As String)
    
    Dim sortDescription As String
    Select Case sortingMetric
        Case "1"
            sortDescription = "Greedy Selection Order"
        Case "2"
            sortDescription = "Profit/Avg DD Benefit"
        Case "3"
            sortDescription = "Profit/Stdev Benefit"
        Case "4"
            sortDescription = "Strategy Number"
        Case "5"
            sortDescription = "Return Impact"
        Case Else
            sortDescription = "Greedy Selection Order"
    End Select
    
    With ws
        .Cells(tableStartRow - 1, tableStartCol).value = "Greedy Strategy Selection Analysis (" & MCTradeType & ") - Sorted by " & sortDescription
        .Cells(tableStartRow - 1, tableStartCol).Font.Bold = True
        .Cells(tableStartRow - 1, tableStartCol).Font.Size = 14
        .Range(.Cells(tableStartRow - 1, tableStartCol), .Cells(tableStartRow - 1, tableStartCol + 10)).Merge  ' Expanded for new column
        .Cells(tableStartRow - 1, tableStartCol).HorizontalAlignment = xlCenter
        .Cells(tableStartRow - 1, tableStartCol).Interior.Color = RGB(0, 102, 204)
        .Cells(tableStartRow - 1, tableStartCol).Font.Color = RGB(255, 255, 255)
        
        With .Range(.Cells(tableStartRow - 1, tableStartCol), .Cells(tableStartRow - 1, tableStartCol + 10)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        
        ' Column headers - UPDATED TO INCLUDE RETURN IMPACT
        .Cells(tableStartRow, tableStartCol).value = "Strategy #"
        .Cells(tableStartRow, tableStartCol + 1).value = "Symbol"
        .Cells(tableStartRow, tableStartCol + 2).value = "Strategy Name"
        .Cells(tableStartRow, tableStartCol + 3).value = "Return Impact"           ' NEW COLUMN
        .Cells(tableStartRow, tableStartCol + 4).value = "Profit/Max DD Benefit"
        .Cells(tableStartRow, tableStartCol + 5).value = "Profit/Avg DD Benefit"
        .Cells(tableStartRow, tableStartCol + 6).value = "Profit/Stdev Benefit"
        .Cells(tableStartRow, tableStartCol + 7).value = "Return Rank"             ' NEW RANKING
        .Cells(tableStartRow, tableStartCol + 8).value = "P/MaxDD Rank"
        .Cells(tableStartRow, tableStartCol + 9).value = "P/AvgDD Rank"
        .Cells(tableStartRow, tableStartCol + 10).value = "P/Stdev Rank"
        
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).Interior.Color = RGB(224, 224, 224)
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).Font.Bold = True
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).WrapText = True
        
        .Cells(tableStartRow, tableStartCol).ColumnWidth = 15
        .Cells(tableStartRow, tableStartCol + 1).ColumnWidth = 10
        .Cells(tableStartRow, tableStartCol + 2).ColumnWidth = 60
       '.Cells(tableStartRow, tableStartCol + 3).ColumnWidth = 15  ' Width for Return Impact
        
        ' Calculate cumulative values - UPDATED FOR RETURN IMPACT
        Dim cumulativeBenefit1 As Double, cumulativeBenefit2 As Double, cumulativeBenefit3 As Double, cumulativeBenefit4 As Double
        cumulativeBenefit1 = 0: cumulativeBenefit2 = 0: cumulativeBenefit3 = 0: cumulativeBenefit4 = 0
        
        For i = 1 To numStrategies
            .Cells(tableStartRow + i, tableStartCol).value = metrics(i, 2)      ' Strategy number
            .Cells(tableStartRow + i, tableStartCol + 1).value = metrics(i, 8)  ' Symbol
            .Cells(tableStartRow + i, tableStartCol + 2).value = metrics(i, 3)  ' Strategy name
            
            ' Individual contribution benefits (null for first strategy)
            If i = 1 Then
                .Cells(tableStartRow + i, tableStartCol + 3).value = ""  ' Return Impact
                .Cells(tableStartRow + i, tableStartCol + 4).value = ""  ' Profit/Max DD
                .Cells(tableStartRow + i, tableStartCol + 5).value = ""  ' Profit/Avg DD
                .Cells(tableStartRow + i, tableStartCol + 6).value = ""  ' Profit/Stdev
                
                ' Gray out first strategy cells
                .Cells(tableStartRow + i, tableStartCol + 3).Interior.Color = RGB(240, 240, 240)
                .Cells(tableStartRow + i, tableStartCol + 4).Interior.Color = RGB(240, 240, 240)
                .Cells(tableStartRow + i, tableStartCol + 5).Interior.Color = RGB(240, 240, 240)
                .Cells(tableStartRow + i, tableStartCol + 6).Interior.Color = RGB(240, 240, 240)
            Else
                ' NEW: Return Impact
                .Cells(tableStartRow + i, tableStartCol + 3).value = metrics(i, 4)
                .Cells(tableStartRow + i, tableStartCol + 3).NumberFormat = "+0.0%;-0.0%;0.0%"
                
                ' Updated column assignments
                .Cells(tableStartRow + i, tableStartCol + 4).value = metrics(i, 5)  ' Profit/Max DD
                .Cells(tableStartRow + i, tableStartCol + 4).NumberFormat = "+0.0%;-0.0%;0.0%"
                
                .Cells(tableStartRow + i, tableStartCol + 5).value = metrics(i, 6)  ' Profit/Avg DD
                .Cells(tableStartRow + i, tableStartCol + 5).NumberFormat = "+0.0%;-0.0%;0.0%"
                
                .Cells(tableStartRow + i, tableStartCol + 6).value = metrics(i, 7)  ' Profit/Stdev
                .Cells(tableStartRow + i, tableStartCol + 6).NumberFormat = "+0.0%;-0.0%;0.0%"
                
                ' Color coding for individual benefits
                For j = 3 To 6
                    If .Cells(tableStartRow + i, tableStartCol + j).value > 0 Then
                        .Cells(tableStartRow + i, tableStartCol + j).Interior.Color = RGB(198, 239, 206)  ' Light green
                    ElseIf .Cells(tableStartRow + i, tableStartCol + j).value < 0 Then
                        .Cells(tableStartRow + i, tableStartCol + j).Interior.Color = RGB(255, 199, 206)  ' Light red
                    End If
                Next j
            End If
            
            ' Calculate cumulative benefits (sum of all improvements up to this point)
            If i = 1 Then
                ' First strategy: cumulative is zero (no previous improvements)
                cumulativeBenefit1 = 0  ' Return Impact
                cumulativeBenefit2 = 0  ' Profit/Max DD
                cumulativeBenefit3 = 0  ' Profit/Avg DD
                cumulativeBenefit4 = 0  ' Profit/Stdev
            Else
                ' Add this strategy's contribution to running totals
                cumulativeBenefit1 = cumulativeBenefit1 + metrics(i, 4)  ' Return Impact
                cumulativeBenefit2 = cumulativeBenefit2 + metrics(i, 5)  ' Profit/Max DD
                cumulativeBenefit3 = cumulativeBenefit3 + metrics(i, 6)  ' Profit/Avg DD
                cumulativeBenefit4 = cumulativeBenefit4 + metrics(i, 7)  ' Profit/Stdev
            End If
            
            ' Updated ranking assignments
            .Cells(tableStartRow + i, tableStartCol + 7).value = rankReturnImpact(i)    ' Return Impact Rank
            .Cells(tableStartRow + i, tableStartCol + 8).value = rankProfitMaxDD(i)     ' Profit/Max DD Rank
            .Cells(tableStartRow + i, tableStartCol + 9).value = rankProfitAvgDD(i)     ' Profit/Avg DD Rank
            .Cells(tableStartRow + i, tableStartCol + 10).value = rankProfitStdev(i)    ' Profit/Stdev Rank
        Next i
        
        With .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow + numStrategies, tableStartCol + 10)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        
        ' Create radar chart for the diversification benefits
        Dim strategyRange As Range, returnRange As Range, maxDDRange As Range, avgDDRange As Range, stdevRange As Range
        
        Set strategyRange = .Range(.Cells(tableStartRow + 1, tableStartCol), .Cells(tableStartRow + numStrategies, tableStartCol))
        Set returnRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 3), .Cells(tableStartRow + numStrategies, tableStartCol + 3))  ' Return Impact
        Set maxDDRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 4), .Cells(tableStartRow + numStrategies, tableStartCol + 4))   ' Profit/Max DD
        Set avgDDRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 5), .Cells(tableStartRow + numStrategies, tableStartCol + 5))   ' Profit/Avg DD
        Set stdevRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 6), .Cells(tableStartRow + numStrategies, tableStartCol + 6))   ' Profit/Stdev
        
        .Shapes.AddChart2(201, xlRadar, .Cells(chartStartRow, chartStartCol).left, _
                        .Cells(chartStartRow, chartStartCol).top, 400, 350).Select
        
        With ActiveChart
            .HasTitle = True
            .chartTitle.text = "Greedy Selection Benefit Profile"
            
            .HasLegend = True
            .Legend.position = xlLegendPositionBottom
            
            Do While .SeriesCollection.count > 0
                .SeriesCollection(1).Delete
            Loop
            
            .SeriesCollection.NewSeries
            .SeriesCollection(1).name = "Return Impact"
            .SeriesCollection(1).values = returnRange
            .SeriesCollection(1).XValues = strategyRange
            
            .SeriesCollection.NewSeries
            .SeriesCollection(2).name = "P/Max DD Benefit"
            .SeriesCollection(2).values = maxDDRange
            .SeriesCollection(2).XValues = strategyRange
            
            .SeriesCollection.NewSeries
            .SeriesCollection(3).name = "P/Avg DD Benefit"
            .SeriesCollection(3).values = avgDDRange
            .SeriesCollection(3).XValues = strategyRange
            
            .SeriesCollection.NewSeries
            .SeriesCollection(4).name = "P/Stdev Benefit"
            .SeriesCollection(4).values = stdevRange
            .SeriesCollection(4).XValues = strategyRange
            
            ' Format the series with different colors - NO MARKERS
            .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(220, 20, 60)    ' Crimson for Return Impact
            .SeriesCollection(1).Format.line.Weight = 2.5
            .SeriesCollection(1).MarkerStyle = xlMarkerStyleNone
            
            .SeriesCollection(2).Format.line.ForeColor.RGB = RGB(65, 140, 240)   ' Blue for P/Max DD
            .SeriesCollection(2).Format.line.Weight = 2.5
            .SeriesCollection(2).MarkerStyle = xlMarkerStyleNone
            
            .SeriesCollection(3).Format.line.ForeColor.RGB = RGB(252, 180, 65)   ' Orange for P/Avg DD
            .SeriesCollection(3).Format.line.Weight = 2.5
            .SeriesCollection(3).MarkerStyle = xlMarkerStyleNone
            
            .SeriesCollection(4).Format.line.ForeColor.RGB = RGB(127, 96, 170)   ' Purple for P/Stdev
            .SeriesCollection(4).Format.line.Weight = 2.5
            .SeriesCollection(4).MarkerStyle = xlMarkerStyleNone
            
            .ChartArea.Format.fill.ForeColor.RGB = RGB(255, 255, 255)
            .PlotArea.Format.fill.ForeColor.RGB = RGB(242, 242, 242)
            
            On Error Resume Next
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.text = "Strategy Number"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.text = "Benefit"
            .Axes(xlCategory).MajorGridlines.Format.line.ForeColor.RGB = RGB(191, 191, 191)
            .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(191, 191, 191)
            On Error GoTo 0
        End With
        
        .Cells(tableStartRow + numStrategies + 2, tableStartCol).value = "Greedy Selection Interpretation:"
        .Cells(tableStartRow + numStrategies + 2, tableStartCol).Font.Bold = True
        
        .Cells(tableStartRow + numStrategies + 3, tableStartCol).value = "• Return Impact shows percentage impact on returns when strategy was added"
        .Cells(tableStartRow + numStrategies + 4, tableStartCol).value = "• Individual benefit shows contribution when added at that step (empty for first)"
        .Cells(tableStartRow + numStrategies + 5, tableStartCol).value = "• Strategies sorted by " & sortDescription
        .Cells(tableStartRow + numStrategies + 6, tableStartCol).value = "• Chart shows diversification benefit profile for each strategy"
        
        .Range(.Cells(tableStartRow + numStrategies + 3, tableStartCol), _
              .Cells(tableStartRow + numStrategies + 6, tableStartCol + 10)).Interior.Color = RGB(242, 242, 242)
    End With
End Sub



Sub OutputDiversificationResultsWithReturn(ws As Worksheet, metrics As Variant, rankReturnImpact() As Long, _
                                rankProfitMaxDD() As Long, rankProfitAvgDD() As Long, rankProfitStdev() As Long, _
                                tableStartRow As Long, tableStartCol As Long, _
                                chartStartRow As Long, chartStartCol As Long, _
                                MCTradeType As String, numStrategies As Long, numIterations As Long, _
                                sortingMetric As String)
    
    Dim sortDescription As String
    Select Case sortingMetric
        Case "1"
            sortDescription = "Profit/Max DD Benefit"
        Case "2"
            sortDescription = "Profit/Avg DD Benefit"
        Case "3"
            sortDescription = "Profit/Stdev Benefit"
        Case "4"
            sortDescription = "Strategy Number"
        Case "5"
            sortDescription = "Return Impact"
        Case Else
            sortDescription = "Profit/Max DD Benefit"
    End Select
    
    With ws
        .Cells(tableStartRow - 1, tableStartCol).value = "Strategy Diversification Analysis (" & MCTradeType & ") - Sorted by " & sortDescription
        .Cells(tableStartRow - 1, tableStartCol).Font.Bold = True
        .Cells(tableStartRow - 1, tableStartCol).Font.Size = 14
        .Range(.Cells(tableStartRow - 1, tableStartCol), .Cells(tableStartRow - 1, tableStartCol + 10)).Merge  ' Expanded for new column
        .Cells(tableStartRow - 1, tableStartCol).HorizontalAlignment = xlCenter
        .Cells(tableStartRow - 1, tableStartCol).Interior.Color = RGB(0, 102, 204)
        .Cells(tableStartRow - 1, tableStartCol).Font.Color = RGB(255, 255, 255)
        
        With .Range(.Cells(tableStartRow - 1, tableStartCol), .Cells(tableStartRow - 1, tableStartCol + 10)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        
        ' Column headers - UPDATED TO INCLUDE RETURN IMPACT
        .Cells(tableStartRow, tableStartCol).value = "Strategy #"
        .Cells(tableStartRow, tableStartCol + 1).value = "Symbol"
        .Cells(tableStartRow, tableStartCol + 2).value = "Strategy Name"
        .Cells(tableStartRow, tableStartCol + 3).value = "Return Impact"          ' NEW COLUMN
        .Cells(tableStartRow, tableStartCol + 4).value = "Profit/Max DD Benefit"
        .Cells(tableStartRow, tableStartCol + 5).value = "Profit/Avg DD Benefit"
        .Cells(tableStartRow, tableStartCol + 6).value = "Profit/Stdev Benefit"
        .Cells(tableStartRow, tableStartCol + 7).value = "Return Rank"            ' NEW RANKING
        .Cells(tableStartRow, tableStartCol + 8).value = "P/MaxDD Rank"
        .Cells(tableStartRow, tableStartCol + 9).value = "P/AvgDD Rank"
        .Cells(tableStartRow, tableStartCol + 10).value = "P/Stdev Rank"
        
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).Interior.Color = RGB(224, 224, 224)
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).Font.Bold = True
        .Range(.Cells(tableStartRow, tableStartCol), .Cells(tableStartRow, tableStartCol + 10)).WrapText = True
        
        ' Set column widths
        .Cells(tableStartRow, tableStartCol).ColumnWidth = 12.5
        .Cells(tableStartRow, tableStartCol + 1).ColumnWidth = 10
        .Cells(tableStartRow, tableStartCol + 2).ColumnWidth = 75
        '.Cells(tableStartRow, tableStartCol + 3).ColumnWidth = 15  ' Width for Return Impact
        
        ' Calculate cumulative benefits based on sorted order
        Dim cumulativeBenefit1 As Double, cumulativeBenefit2 As Double, cumulativeBenefit3 As Double, cumulativeBenefit4 As Double
        cumulativeBenefit1 = 0: cumulativeBenefit2 = 0: cumulativeBenefit3 = 0: cumulativeBenefit4 = 0
        
        ' Populate the data rows
        For i = 1 To numStrategies
            .Cells(tableStartRow + i, tableStartCol).value = metrics(i, 2)      ' Strategy number
            .Cells(tableStartRow + i, tableStartCol + 1).value = metrics(i, 8)  ' Symbol
            .Cells(tableStartRow + i, tableStartCol + 2).value = metrics(i, 3)  ' Strategy name
            
            ' NEW: Return Impact column
            .Cells(tableStartRow + i, tableStartCol + 3).value = metrics(i, 4)  ' Return Impact
            .Cells(tableStartRow + i, tableStartCol + 3).NumberFormat = "+0.0%;-0.0%;0.0%"
            
            ' Updated column assignments
            .Cells(tableStartRow + i, tableStartCol + 4).value = metrics(i, 5)  ' Profit/Max DD Benefit
            .Cells(tableStartRow + i, tableStartCol + 4).NumberFormat = "+0.0%;-0.0%;0.0%"
            
            .Cells(tableStartRow + i, tableStartCol + 5).value = metrics(i, 6)  ' Profit/Avg DD Benefit
            .Cells(tableStartRow + i, tableStartCol + 5).NumberFormat = "+0.0%;-0.0%;0.0%"
            
            .Cells(tableStartRow + i, tableStartCol + 6).value = metrics(i, 7)  ' Profit/Stdev Benefit
            .Cells(tableStartRow + i, tableStartCol + 6).NumberFormat = "+0.0%;-0.0%;0.0%"
            
            ' Updated ranking assignments
            .Cells(tableStartRow + i, tableStartCol + 7).value = rankReturnImpact(i)    ' Return Impact Rank
            .Cells(tableStartRow + i, tableStartCol + 8).value = rankProfitMaxDD(i)     ' Profit/Max DD Rank
            .Cells(tableStartRow + i, tableStartCol + 9).value = rankProfitAvgDD(i)     ' Profit/Avg DD Rank
            .Cells(tableStartRow + i, tableStartCol + 10).value = rankProfitStdev(i)    ' Profit/Stdev Rank
            
            ' Color coding for benefits (columns 3-6)
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
        
        ' Create radar chart for the diversification benefits
        Dim strategyRange As Range, returnRange As Range, maxDDRange As Range, avgDDRange As Range, stdevRange As Range
        
        Set strategyRange = .Range(.Cells(tableStartRow + 1, tableStartCol), .Cells(tableStartRow + numStrategies, tableStartCol))
        Set returnRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 3), .Cells(tableStartRow + numStrategies, tableStartCol + 3))    ' Return Impact
        Set maxDDRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 4), .Cells(tableStartRow + numStrategies, tableStartCol + 4))     ' Profit/Max DD
        Set avgDDRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 5), .Cells(tableStartRow + numStrategies, tableStartCol + 5))     ' Profit/Avg DD
        Set stdevRange = .Range(.Cells(tableStartRow + 1, tableStartCol + 6), .Cells(tableStartRow + numStrategies, tableStartCol + 6))     ' Profit/Stdev
        
        .Shapes.AddChart2(201, xlRadar, .Cells(chartStartRow, chartStartCol).left, _
                        .Cells(chartStartRow, chartStartCol).top, 400, 350).Select
        
        With ActiveChart
            .HasTitle = True
            .chartTitle.text = "Strategy Diversification Profile"
            
            .HasLegend = True
            .Legend.position = xlLegendPositionBottom
            
            Do While .SeriesCollection.count > 0
                .SeriesCollection(1).Delete
            Loop
            
            .SeriesCollection.NewSeries
            .SeriesCollection(1).name = "Return Impact"
            .SeriesCollection(1).values = returnRange
            .SeriesCollection(1).XValues = strategyRange
            
            .SeriesCollection.NewSeries
            .SeriesCollection(2).name = "P/Max DD Benefit"
            .SeriesCollection(2).values = maxDDRange
            .SeriesCollection(2).XValues = strategyRange
            
            .SeriesCollection.NewSeries
            .SeriesCollection(3).name = "P/Avg DD Benefit"
            .SeriesCollection(3).values = avgDDRange
            .SeriesCollection(3).XValues = strategyRange
            
            .SeriesCollection.NewSeries
            .SeriesCollection(4).name = "P/Stdev Benefit"
            .SeriesCollection(4).values = stdevRange
            .SeriesCollection(4).XValues = strategyRange
            
            ' Format the series with different colors - NO MARKERS
            .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(220, 20, 60)    ' Crimson for Return Impact
            .SeriesCollection(1).Format.line.Weight = 2.5
            .SeriesCollection(1).MarkerStyle = xlMarkerStyleNone
            
            .SeriesCollection(2).Format.line.ForeColor.RGB = RGB(65, 140, 240)   ' Blue for P/Max DD
            .SeriesCollection(2).Format.line.Weight = 2.5
            .SeriesCollection(2).MarkerStyle = xlMarkerStyleNone
            
            .SeriesCollection(3).Format.line.ForeColor.RGB = RGB(252, 180, 65)   ' Orange for P/Avg DD
            .SeriesCollection(3).Format.line.Weight = 2.5
            .SeriesCollection(3).MarkerStyle = xlMarkerStyleNone
            
            .SeriesCollection(4).Format.line.ForeColor.RGB = RGB(127, 96, 170)   ' Purple for P/Stdev
            .SeriesCollection(4).Format.line.Weight = 2.5
            .SeriesCollection(4).MarkerStyle = xlMarkerStyleNone
            
            .ChartArea.Format.fill.ForeColor.RGB = RGB(255, 255, 255)
            .PlotArea.Format.fill.ForeColor.RGB = RGB(242, 242, 242)
            
            On Error Resume Next
            .Axes(xlCategory).MajorGridlines.Format.line.ForeColor.RGB = RGB(191, 191, 191)
            .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(191, 191, 191)
            On Error GoTo 0
        End With
        
        .Cells(tableStartRow + numStrategies + 2, tableStartCol).value = "Interpretation Guide:"
        .Cells(tableStartRow + numStrategies + 2, tableStartCol).Font.Bold = True
        
        .Cells(tableStartRow + numStrategies + 3, tableStartCol).value = "• Return Impact shows percentage impact on returns when strategy is added"
        .Cells(tableStartRow + numStrategies + 4, tableStartCol).value = "• Other benefits show average diversification improvement for each strategy"
        .Cells(tableStartRow + numStrategies + 5, tableStartCol).value = "• Based on " & numIterations & " random order simulations"
        .Cells(tableStartRow + numStrategies + 6, tableStartCol).value = "• Lower rank numbers (1 = best) indicate more valuable strategies"
        
        .Range(.Cells(tableStartRow + numStrategies + 3, tableStartCol), _
              .Cells(tableStartRow + numStrategies + 6, tableStartCol + 10)).Interior.Color = RGB(242, 242, 242)
    End With
End Sub




' Alternative approach using a UserForm for more professional dropdown selection
Sub ShowSortingOptionsForm()
    ' This would create a more professional interface
    ' You would need to create a UserForm with ComboBox controls
    ' Here's the basic structure:
    
    Dim sortingChoice As String
    
    ' Create a simple form or use Application.InputBox with validation
    sortingChoice = InputBox("Select sorting criteria:" & vbCrLf & vbCrLf & _
                           "1 - Profit/Max DD Benefit (Best diversifiers first)" & vbCrLf & _
                           "2 - Profit/Avg DD Benefit (Consistent performers first)" & vbCrLf & _
                           "3 - Profit/Stdev Benefit (Risk reducers first)" & vbCrLf & _
                           "4 - Strategy Number (Original order)" & vbCrLf & vbCrLf & _
                           "Enter choice (1-4):", _
                           "Sort Output By", "1")
    
    ' Validate input
    If sortingChoice = "" Then Exit Sub
    If Not (sortingChoice = "1" Or sortingChoice = "2" Or sortingChoice = "3" Or sortingChoice = "4") Then
        MsgBox "Invalid selection. Please choose 1, 2, 3, or 4.", vbExclamation
        Exit Sub
    End If
    
    ' Continue with analysis using sortingChoice
    ' (This would integrate with your main analysis routine)
End Sub


Function GenerateRandomStrategyOrder(numStrategies As Long) As Long()
    Dim randomOrder() As Long
    Dim i As Long, j As Long, temp As Long
    
    ReDim randomOrder(1 To numStrategies)
    For i = 1 To numStrategies
        randomOrder(i) = i
    Next i
    
    For i = numStrategies To 2 Step -1
        j = Int(Rnd() * i) + 1
        temp = randomOrder(i)
        randomOrder(i) = randomOrder(j)
        randomOrder(j) = temp
    Next i
    
    GenerateRandomStrategyOrder = randomOrder
End Function

Function CreateEmptyPortfolio(pnlResults As Variant) As Variant
    Dim emptyPortfolio As Variant
    Dim i As Long, j As Long
    
    ReDim emptyPortfolio(1 To UBound(pnlResults, 1), 1 To UBound(pnlResults, 2))
    
    For i = 1 To UBound(pnlResults, 1)
        For j = 1 To UBound(pnlResults, 2)
            emptyPortfolio(i, j) = 0
        Next j
    Next i
    
    CreateEmptyPortfolio = emptyPortfolio
End Function

Function BuildIncrementalPortfolio(pnlResults As Variant, strategyOrder() As Long, numStrategiesToInclude As Long) As Variant
    Dim portfolio As Variant
    Dim i As Long, j As Long, K As Long
    Dim includeStrategy As Boolean
    
    ReDim portfolio(1 To UBound(pnlResults, 1), 1 To UBound(pnlResults, 2))
    
    For i = 1 To UBound(pnlResults, 1)
        For j = 1 To UBound(pnlResults, 2)
            includeStrategy = False
            For K = 1 To numStrategiesToInclude
                If strategyOrder(K) = j Then
                    includeStrategy = True
                    Exit For
                End If
            Next K
            
            If includeStrategy Then
                portfolio(i, j) = pnlResults(i, j)
            Else
                portfolio(i, j) = 0
            End If
        Next j
    Next i
    
    BuildIncrementalPortfolio = portfolio
End Function

Function CreateSingleStrategyPortfolio(pnlResults As Variant, strategyIndex As Long) As Variant
    Dim portfolio As Variant
    Dim i As Long, j As Long
    
    ReDim portfolio(1 To UBound(pnlResults, 1), 1 To UBound(pnlResults, 2))
    
    For i = 1 To UBound(pnlResults, 1)
        For j = 1 To UBound(pnlResults, 2)
            If j = strategyIndex Then
                portfolio(i, j) = pnlResults(i, j)
            Else
                portfolio(i, j) = 0
            End If
        Next j
    Next i
    
    CreateSingleStrategyPortfolio = portfolio
End Function

Function BuildSelectedStrategiesPortfolio(pnlResults As Variant, strategySelected() As Boolean) As Variant
    Dim portfolio As Variant
    Dim i As Long, j As Long
    
    ReDim portfolio(1 To UBound(pnlResults, 1), 1 To UBound(pnlResults, 2))
    
    For i = 1 To UBound(pnlResults, 1)
        For j = 1 To UBound(pnlResults, 2)
            If strategySelected(j) Then
                portfolio(i, j) = pnlResults(i, j)
            Else
                portfolio(i, j) = 0
            End If
        Next j
    Next i
    
    BuildSelectedStrategiesPortfolio = portfolio
End Function

Function BuildSequentialPortfolio(pnlResults As Variant, selectionOrder() As Long, numToInclude As Long) As Variant
    Dim portfolio As Variant
    Dim i As Long, j As Long, K As Long
    Dim includeStrategy As Boolean
    
    ReDim portfolio(1 To UBound(pnlResults, 1), 1 To UBound(pnlResults, 2))
    
    For i = 1 To UBound(pnlResults, 1)
        For j = 1 To UBound(pnlResults, 2)
            includeStrategy = False
            For K = 1 To numToInclude
                If selectionOrder(K) = j Then
                    includeStrategy = True
                    Exit For
                End If
            Next K
            
            If includeStrategy Then
                portfolio(i, j) = pnlResults(i, j)
            Else
                portfolio(i, j) = 0
            End If
        Next j
    Next i
    
    BuildSequentialPortfolio = portfolio
End Function

Function CalculateDiversificationContributionWithReturn(oldPortfolio As Variant, newPortfolio As Variant, _
                                            requiredMargin As Double, startingEquity As Double, _
                                            AverageTrade As Double) As Variant
    Dim oldResults As Variant, newResults As Variant
    Dim oldReturn As Double, newReturn As Double
    Dim oldMaxDD As Double, newMaxDD As Double, oldAvgDD As Double, newAvgDD As Double
    Dim oldProfitMaxDD As Double, newProfitMaxDD As Double
    Dim oldProfitAvgDD As Double, newProfitAvgDD As Double
    Dim oldProfitStdev As Double, newProfitStdev As Double
    Dim contribution(1 To 4) As Double  ' Expanded to include return impact
    
    oldResults = RunChronologicalBacktestWithTracking(oldPortfolio, requiredMargin, startingEquity, AverageTrade)
    newResults = RunChronologicalBacktestWithTracking(newPortfolio, requiredMargin, startingEquity, AverageTrade)
    
    ' Calculate returns and drawdowns (all in dollars)
    oldReturn = CalculateMedianReturn(oldResults)
    newReturn = CalculateMedianReturn(newResults)
    oldMaxDD = CalculateMedianMaxDrawdown(oldResults)
    newMaxDD = CalculateMedianMaxDrawdown(newResults)
    oldAvgDD = CalculateMedianAvgDrawdown(oldResults)
    newAvgDD = CalculateMedianAvgDrawdown(newResults)
    
    ' Calculate Return Impact (percentage change in dollar returns)
    If Abs(oldReturn) > 0.00001 Then
        contribution(1) = (newReturn - oldReturn) / oldReturn
    Else
        contribution(1) = 0
    End If
    
    ' Calculate ratios - dollars/dollars
    If oldMaxDD > 0.01 Then oldProfitMaxDD = oldReturn / oldMaxDD Else oldProfitMaxDD = oldReturn * 1000
    If newMaxDD > 0.01 Then newProfitMaxDD = newReturn / newMaxDD Else newProfitMaxDD = newReturn * 1000
    If oldAvgDD > 0.01 Then oldProfitAvgDD = oldReturn / oldAvgDD Else oldProfitAvgDD = oldReturn * 1000
    If newAvgDD > 0.01 Then newProfitAvgDD = newReturn / newAvgDD Else newProfitAvgDD = newReturn * 1000
    
    ' Calculate Profit/Stdev ratios
    oldProfitStdev = CalculatePortfolioStdev(oldPortfolio)
    newProfitStdev = CalculatePortfolioStdev(newPortfolio)
    
    Dim oldReturnStdev As Double, newReturnStdev As Double
    If oldProfitStdev > 0.01 Then oldReturnStdev = oldReturn / oldProfitStdev Else oldReturnStdev = oldReturn
    If newProfitStdev > 0.01 Then newReturnStdev = newReturn / newProfitStdev Else newReturnStdev = newReturn
    
    ' Calculate percentage contributions WITH THRESHOLDS
    If Abs(oldProfitMaxDD) > 0.001 Then
        Dim maxDDContrib As Double
        maxDDContrib = (newProfitMaxDD - oldProfitMaxDD) / Abs(oldProfitMaxDD)
        If Abs(maxDDContrib) < 0.005 Then  ' 0.5% threshold
            contribution(2) = 0
        Else
            contribution(2) = maxDDContrib
        End If
    Else
        contribution(2) = 0
    End If
    
    If Abs(oldProfitAvgDD) > 0.001 Then
        Dim avgDDContrib As Double
        avgDDContrib = (newProfitAvgDD - oldProfitAvgDD) / Abs(oldProfitAvgDD)
        If Abs(avgDDContrib) < 0.005 Then  ' 0.5% threshold
            contribution(3) = 0
        Else
            contribution(3) = avgDDContrib
        End If
    Else
        contribution(3) = 0
    End If
    
    If Abs(oldReturnStdev) > 0.001 Then
        Dim stdevContrib As Double
        stdevContrib = (newReturnStdev - oldReturnStdev) / Abs(oldReturnStdev)
        If Abs(stdevContrib) < 0.005 Then  ' 0.5% threshold
            contribution(4) = 0
        Else
            contribution(4) = stdevContrib
        End If
    Else
        contribution(4) = 0
    End If
    
    CalculateDiversificationContributionWithReturn = contribution
End Function



Function CalculateMedianFromArray(contributionMatrix() As Double, strategyIndex As Long, metricIndex As Long, numIterations As Long) As Double
    Dim values() As Double
    Dim i As Long
    
    ReDim values(1 To numIterations)
    For i = 1 To numIterations
        values(i) = contributionMatrix(strategyIndex, i, metricIndex)
    Next i
    
    CalculateMedianFromArray = Application.WorksheetFunction.Median(values)
End Function

Sub SortDiversificationMetricsWithReturn(ByRef metrics As Variant, sortingMetric As String, numStrategies As Long)
    ' Sort the metrics array based on the selected criteria - UPDATED FOR RETURN IMPACT
    Dim i As Long, j As Long, K As Long
    Dim tempRow As Variant
    Dim sortColumn As Long
    Dim sortAscending As Boolean
    
    ReDim tempRow(1 To UBound(metrics, 2))
    
    ' Determine sort column and direction
    Select Case sortingMetric
        Case "1"
            sortColumn = 5  ' Profit/Max DD Benefit (moved due to new Return Impact column)
            sortAscending = False  ' Descending (best first)
        Case "2"
            sortColumn = 6  ' Profit/Avg DD Benefit (moved due to new Return Impact column)
            sortAscending = False  ' Descending (best first)
        Case "3"
            sortColumn = 7  ' Profit/Stdev Benefit (moved due to new Return Impact column)
            sortAscending = False  ' Descending (best first)
        Case "4"
            sortColumn = 2  ' Strategy Number
            sortAscending = True   ' Ascending (original order)
        Case "5"
            sortColumn = 4  ' Return Impact (new column)
            sortAscending = False  ' Descending (best first)
        Case Else
            sortColumn = 5  ' Default to Profit/Max DD
            sortAscending = False
    End Select
    
    ' Bubble sort implementation
    For i = 1 To numStrategies - 1
        For j = i + 1 To numStrategies
            Dim shouldSwap As Boolean
            
            ' Handle potential null/empty values for first strategy
            Dim value1 As Double, value2 As Double
            
            ' Get values, treating empty/null as 0 for sorting purposes
            If IsEmpty(metrics(i, sortColumn)) Or metrics(i, sortColumn) = "" Then
                value1 = 0
            Else
                value1 = CDbl(metrics(i, sortColumn))
            End If
            
            If IsEmpty(metrics(j, sortColumn)) Or metrics(j, sortColumn) = "" Then
                value2 = 0
            Else
                value2 = CDbl(metrics(j, sortColumn))
            End If
            
            ' Determine if we should swap
            If sortAscending Then
                shouldSwap = (value1 > value2)
            Else
                shouldSwap = (value1 < value2)
            End If
            
            If shouldSwap Then
                ' Swap rows
                For K = 1 To UBound(metrics, 2)
                    tempRow(K) = metrics(i, K)
                    metrics(i, K) = metrics(j, K)
                    metrics(j, K) = tempRow(K)
                Next K
            End If
        Next j
    Next i
End Sub




Sub AddNavigationButtonsDiversificator(ws As Worksheet)
    ' Add navigation buttons to the worksheet
    Dim btn As Object
    
    ' Create delete button
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 1).left + 30, top:=ws.Cells(35, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteDiversificator" ' Make sure to create this sub to handle deletion
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
        .OnAction = "GoToPortfolio" ' Assign the macro to run when the button is coldMaxDD licked
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

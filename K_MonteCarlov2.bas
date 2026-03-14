Attribute VB_Name = "K_MonteCarlov2"

Sub UpdatedRunPortfolioMonteCarloSimulation()
    ' Updated main function for running Monte Carlo simulation with charts integrated into MC tab
    Dim wsPortfolio As Worksheet
    Dim wsPortfolioMC As Worksheet
    Dim wsTotalPortfolioM2M As Worksheet
    Dim startingMargin As Double
    Dim startingEquity As Double
    Dim pnlResults As Variant
    Dim requiredMargin As Double
    Dim averageTradesPerYear As Long
    Dim numScenarios As Long
    Dim results As Variant
    Dim targetRiskOfRuin As Double
    Dim tolerance As Double
    Dim currentRiskOfRuin As Double
    Dim yearsToConsider As Double
    Dim ruinedCount As Long
    Dim count As Long
    Dim startdate As Date
    Dim endDate As Date
    Dim tradeAdjustment As Double
    Dim summaryRow As Long
    Dim solveRisk As String, ceaseTradingType As String
    Dim maxProfit As Double, margin As Double
    Dim binWidth As Double
    Dim MCTradeType As String
    Dim dailyProfitTracking() As Double
    Dim dailyDrawdownTracking() As Double
    Dim dailyMaxDrawdownTracking() As Double
    Dim dailyStats As Variant
    Dim Outputsamples As Long
    
    ' Check license validity
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        Exit Sub
    End If

    Call InitializeColumnConstantsManually
    
    On Error Resume Next
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsPortfolio Is Nothing Then
        MsgBox "Error: 'Portfolio' sheet does not exist.", vbExclamation
        Exit Sub
    End If

    ' Exit and show error if the sheet exists but has no data in row 2
    If wsPortfolio.Cells(2, COL_PORT_STRATEGY_NAME).value = "" Then
        MsgBox "Error: 'Portfolio' sheet exists but contains no data in row 2.", vbExclamation
        Exit Sub
    End If
    
    ' Check if the PortfolioMC tab already exists, delete if it does
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("PortfolioMC").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Set up the Portfolio Summary tab
    Set wsPortfolioMC = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsPortfolioMC.name = "PortfolioMC"
    wsPortfolioMC.Tab.Color = RGB(117, 219, 255) ' blue
    
    ' Set white background color for the entire worksheet
    wsPortfolioMC.Cells.Interior.Color = RGB(255, 255, 255)
    
    Set wsTotalPortfolioM2M = ThisWorkbook.Sheets("TotalPortfolioM2M")
    
    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    Application.ScreenUpdating = False
    
    ' Set up parameters
    ceaseTradingType = GetNamedRangeValue("PortfolioCeaseTradingType")
    requiredMargin = GetNamedRangeValue("PortfolioCeaseTrading")
    startingEquity = GetNamedRangeValue("PortfolioStartingEquity")
    numScenarios = GetNamedRangeValue("PortfolioSimulations")
    tradeAdjustment = GetNamedRangeValue("PortfolioMCTradeAdjustment")
    targetRiskOfRuin = GetNamedRangeValue("PortfolioRiskRuinTarget")
    tolerance = GetNamedRangeValue("PortfolioRiskRuinTolerance")
    solveRisk = GetNamedRangeValue("Solve_Risk_Ruin")
    MCTradeType = GetNamedRangeValue("PortMCTradeType")
    Outputsamples = GetNamedRangeValue("Port_MC_OutputSamples")
    
    ' Dates and trading information
    yearsToConsider = GetNamedRangeValue("PortfolioPeriod")
    currentdate = wsPortfolio.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    startdate = DateAdd("yyyy", -Int(yearsToConsider), currentdate)
    startdate = DateAdd("m", -(yearsToConsider - Int(yearsToConsider)) * 12, startdate)
    endDate = currentdate
    
  
    
    ' Get the PnL Data - UPDATED SECTION FOR WEEKLY CONVERSION
    If MCTradeType = "Daily" Then
        ' Use the standard daily cleaning function
        pnlResults = CleanPortfolioDailyPnL(startdate, endDate)
    ElseIf MCTradeType = "Weekly" Then
        ' Use the new weekly conversion function instead of reading from TotalPortfolioM2M
        pnlResults = ConvertDailyToWeeklyPnL(startdate, endDate)
    Else
        ' Fallback to original code for any other type
        Dim lastRow As Long
        If Application.WorksheetFunction.CountA(wsTotalPortfolioM2M.Columns(14)) > 0 Then
            lastRow = wsTotalPortfolioM2M.Cells(wsTotalPortfolioM2M.rows.count, 14).End(xlUp).row
        Else
            lastRow = 1 ' Default to the first row if no data is found
        End If

        Dim i As Long
        ReDim pnlResults(1 To lastRow - 1)
        For i = 2 To lastRow
            pnlResults(i - 1) = wsTotalPortfolioM2M.Cells(i, 14).value
        Next i
    End If
    
    
      ' Set the appropriate number of trading periods per year
    If MCTradeType = "Daily" Then
            averageTradesPerYear = 252
            x = UBound(pnlResults)
            
    Else
            averageTradesPerYear = 52  ' weekly trades
    End If
    
    ' Check if we have valid PnL data
    If IsEmpty(pnlResults) Then
        MsgBox "No valid PnL data found.", vbExclamation
        Exit Sub
    End If
    
    ' Debug.Print "Total PnL Before Removal: $" & Format(sumBefore, "#,##0")
    
    ' Remove top performers
    Dim removeTopPercent As Double
    removeTopPercent = Range("Remove_Top_Percent").value
    Debug.Print "Removing Top " & removeTopPercent & "% of positive days"
    pnlResults = RemoveTopPerformers(pnlResults, removeTopPercent, MCTradeType)

    ' Calculate average trade
    Dim AverageTrade As Double, j As Long, K As Long
    
    AverageTrade = 0
    
    If MCTradeType = "Daily" Or MCTradeType = "Weekly" Then
        ' For daily/weekly multi-strategy data, calculate average across all strategies and periods
        Dim numStrategies As Long, TradeCount As Long
        numStrategies = UBound(pnlResults, 2)
        TradeCount = 0
        
        For j = 1 To UBound(pnlResults, 1)
            For K = 1 To numStrategies
                If pnlResults(j, K) <> 0 Then
                    AverageTrade = AverageTrade + pnlResults(j, K)
                End If
            Next K
            TradeCount = TradeCount + 1
        Next j
        
        If TradeCount > 0 Then
            AverageTrade = AverageTrade / TradeCount
        End If
    Else
        ' For other data types, calculate simple average
        For j = 1 To UBound(pnlResults)
            AverageTrade = AverageTrade + pnlResults(j)
        Next j
        
        If UBound(pnlResults) > 0 Then
            AverageTrade = AverageTrade / UBound(pnlResults)
        End If
    End If
    
    count = 0 ' Initialize the iteration count

    ' Do While loop to adjust starting equity based on risk of ruin
    Do
        ' Calculate margin based on cease trading type
        If ceaseTradingType = "Percentage" Then
            margin = (1 - requiredMargin) * startingEquity
        Else
            margin = requiredMargin
        End If
        
        ' Run Monte Carlo simulation with equity tracking
        results = RunMonteCarloWithTracking(pnlResults, margin, averageTradesPerYear, startingEquity, _
                                           numScenarios, tradeAdjustment, AverageTrade, MCTradeType, _
                                           dailyProfitTracking, dailyDrawdownTracking, dailyMaxDrawdownTracking)
        
        ' Make sure we have valid results
        Dim rowCount As Long
        rowCount = UBound(results, 1)
        
        If rowCount < 1 Then
            MsgBox "No valid scenarios to evaluate.", vbExclamation
            Exit Sub
        End If
        
        ' Calculate risk of ruin
        ruinedCount = 0
        For K = 1 To rowCount
            If results(K, 6) = 1 Then
                ruinedCount = ruinedCount + 1
            End If
        Next K
        
        ' Calculate the current risk of ruin as a percentage
        currentRiskOfRuin = ruinedCount / numScenarios
        
        ' Check if current risk of ruin is within tolerance
        If solveRisk = "Yes" Then
            If currentRiskOfRuin > targetRiskOfRuin + tolerance Then
                startingEquity = startingEquity * 1.05
            ElseIf currentRiskOfRuin < targetRiskOfRuin - tolerance Then
                startingEquity = startingEquity * 0.991
            End If
        End If
        
              
        count = count + 1
        
        Application.StatusBar = "Monte Carlo Running: " & count & " runs completed"
        
    Loop While Abs(currentRiskOfRuin - targetRiskOfRuin) > tolerance And count < 500 And solveRisk = "Yes"

    ' Calculate summary statistics after exiting the loop
    Dim medianReturn As Double, medianDrawdown As Double, medianProfit As Double, medianReturnToDrawdown As Double
    
    ' Calculate metrics
    medianReturn = WorksheetFunction.Median(Application.index(results, 0, 3))
    medianDrawdown = WorksheetFunction.Median(Application.index(results, 0, 5)) / startingEquity
    medianProfit = WorksheetFunction.Median(Application.index(results, 0, 2))
    medianReturnToDrawdown = WorksheetFunction.Median(Application.index(results, 0, 4))
    
    ' Calculate daily equity statistics for chart data
    dailyStats = CalculateDailyEquityStatistics(dailyProfitTracking)
    
    ' Create equity progression chart directly in the PortfolioMC sheet
    Call CreateProfitProgressionChart(wsPortfolioMC, dailyStats, "Profit Over Time", 28, 16)
    Call CreateProfitSamplePathsChart(wsPortfolioMC, dailyProfitTracking, Outputsamples & " Sample Profit Paths", Outputsamples, 2, 16)
    
    ' After creating equity charts, add drawdown charts
    Dim drawdownStats As Variant, maxdrawdownStats As Variant
    
    maxdrawdownStats = CalculateDailyEquityStatistics(dailyMaxDrawdownTracking)
        
    ' Create drawdown progression chart
    Call CreateDrawdownProgressionChart(wsPortfolioMC, maxdrawdownStats, "Max Drawdown Over Time", 28, 29)
    Call CreateDrawdownSamplePathsChart(wsPortfolioMC, dailyMaxDrawdownTracking, Outputsamples & " Sample Max Drawdown Paths", Outputsamples, 2, 29)
       
    drawdownStats = CalculateDailyEquityStatistics(dailyDrawdownTracking)  ' Reuse same function for stats
        
    ' Create drawdown progression chart
    Call CreateDrawdownProgressionChart(wsPortfolioMC, drawdownStats, "Avg Drawdown Over Time", 28, 42)
    Call CreateDrawdownSamplePathsChart(wsPortfolioMC, dailyDrawdownTracking, Outputsamples & " Sample Avg Drawdown Paths", Outputsamples, 2, 42)
    
    ' Calculate margin based on cease trading type
    If ceaseTradingType = "Percentage" Then
        margin = (1 - requiredMargin) * startingEquity
    Else
        margin = requiredMargin
    End If
    
    ' Output summary statistics to the worksheet
    Call OutputMonteCarloSummaryMetrics(wsPortfolioMC, results, startingEquity, margin, _
                                      yearsToConsider, medianReturnToDrawdown, currentRiskOfRuin)

    ' Generate histograms
    maxProfit = WorksheetFunction.Max(Application.index(results, 0, 2))
 
    ' Calculate bin width as one-tenth of maxprofit, rounded appropriately
    If Application.WorksheetFunction.Round(maxProfit / 10, -6) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxProfit / 10, -6)
    ElseIf Application.WorksheetFunction.Round(maxProfit / 10, -5) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxProfit / 10, -5)
    ElseIf Application.WorksheetFunction.Round(maxProfit / 10, -4) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxProfit / 10, -4)
    ElseIf Application.WorksheetFunction.Round(maxProfit / 10, -3) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxProfit / 10, -3)
    ElseIf Application.WorksheetFunction.Round(maxProfit / 10, -2) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxProfit / 10, -2)
    ElseIf Application.WorksheetFunction.Round(maxProfit / 10, -1) <> 0 Then
        binWidth = Application.WorksheetFunction.Round(maxProfit / 10, -1)
    Else
        binWidth = 1
    End If
    
    Dim returnArray() As Variant
    Dim profitArray() As Variant
    Dim drawdownArray() As Variant
   
    ' Convert to 1D array
    ReDim returnArray(LBound(results, 1) To UBound(results, 1))
    ReDim profitArray(LBound(results, 1) To UBound(results, 1))
    ReDim drawdownArray(LBound(results, 1) To UBound(results, 1))
    For i = LBound(results, 1) To UBound(results, 1)
        returnArray(i) = results(i, 3)
        drawdownArray(i) = results(i, 5) / startingEquity
        profitArray(i) = results(i, 2)
    Next i

    summaryRow = 3
    
    ' Generate histograms
    CreateHistogram wsPortfolioMC, returnArray, "Return Histogram", 7, 60, summaryRow, 2, 0.05
    CreateHistogram wsPortfolioMC, drawdownArray, "Max Drawdown Histogram", 7, 63, summaryRow, 18, 0.05
    CreateHistogram wsPortfolioMC, profitArray, "Profit Histogram", 7, 66, summaryRow, 34, binWidth
    
    ' Set window properties
    With ThisWorkbook.Windows(1)
        .Zoom = 70 ' Set zoom level to 70%
    End With

    ' Autofit columns for readability
    wsPortfolioMC.Columns("A:D").ColumnWidth = 16
    
    wsPortfolioMC.Range("A1").Select
    
    ' Add navigation buttons
    AddNavigationButtons wsPortfolioMC
    
    ' Order tabs
    Call OrderVisibleTabsBasedOnList
    
    ' Reset status bar and enable screen updating
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    ' Activate the Monte Carlo sheet
    wsPortfolioMC.Activate
End Sub


Function CalculateDailyEquityStatistics(dailyProfitTracking() As Double) As Variant
    ' Calculate daily statistics from the equity tracking array and return as 2D array
    ' Returns array with stats for each day:
    '   Column 1: Day number
    '   Column 2: Average
    '   Column 3: Median
    '   Column 4: 10th percentile
    '   Column 5: 25th percentile
    '   Column 6: 75th percentile
    '   Column 7: 90th percentile
    '   Column 8: 99th percentile
    '   Column 9: 1st percentile
    '   Column 10: Minimum value
    '   Column 11: Maximum value
    
    
    
    Dim numScenarios As Long, numDays As Long
    Dim i As Long, j As Long
    Dim dailyStats As Variant
    
    ' Get dimensions
    numScenarios = UBound(dailyProfitTracking, 1)
    numDays = UBound(dailyProfitTracking, 2)
    
    ' Initialize output array
    ReDim dailyStats(0 To numDays, 1 To 13)
    
    ' Create a temporary array for sorting values for each day
    Dim tempValues() As Double
    ReDim tempValues(1 To numScenarios)
    
    ' Calculate statistics for each day
    For j = 0 To numDays
        ' Set day number
        dailyStats(j, 1) = j
        
        ' Extract values for this day across all scenarios
        For i = 1 To numScenarios
            tempValues(i) = dailyProfitTracking(i, j)
        Next i
        
        ' Calculate basic statistics
        dailyStats(j, 2) = Application.WorksheetFunction.Average(tempValues)  ' Average
        dailyStats(j, 3) = Application.WorksheetFunction.Median(tempValues)   ' Median
        dailyStats(j, 10) = Application.WorksheetFunction.Min(tempValues)     ' Minimum
        dailyStats(j, 11) = Application.WorksheetFunction.Max(tempValues)     ' Maximum
        
        ' Calculate percentiles directly using Excel's built-in PERCENTILE function
        dailyStats(j, 4) = Application.WorksheetFunction.percentile(tempValues, 0.1)    ' 10th percentile
        dailyStats(j, 5) = Application.WorksheetFunction.percentile(tempValues, 0.25)  ' 25th percentile
        dailyStats(j, 6) = Application.WorksheetFunction.percentile(tempValues, 0.75)   ' 75th percentile
        dailyStats(j, 7) = Application.WorksheetFunction.percentile(tempValues, 0.9)    ' 90th percentile
        dailyStats(j, 8) = Application.WorksheetFunction.percentile(tempValues, 0.99)  ' 99th percentile
        dailyStats(j, 9) = Application.WorksheetFunction.percentile(tempValues, 0.01)  ' 1st percentile
        dailyStats(j, 12) = Application.WorksheetFunction.percentile(tempValues, 0.95)  ' 95th percentile
        dailyStats(j, 13) = Application.WorksheetFunction.percentile(tempValues, 0.05)  '5th percentile
    Next j
    
    CalculateDailyEquityStatistics = dailyStats
End Function




Sub AddNavigationButtons(ws As Worksheet)
    ' Add navigation buttons to the worksheet
    Dim btn As Object
    
    ' Create delete button
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 1).left + 30, top:=ws.Cells(35, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeletePortfolioMC" ' Make sure to create this sub to handle deletion
    End With

    ' Create a button to return to the Summary page
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 3).left + 30, top:=ws.Cells(35, 1).top, Width:=100, Height:=25)
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
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 3).left + 30, top:=ws.Cells(38, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl" ' Assign the macro to run when the button is clicked
    End With
    
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 1).left + 30, top:=ws.Cells(41, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies" ' Assign the macro to run when the button is clicked
    End With
    
    Set btn = ws.Buttons.Add(left:=ws.Cells(1, 3).left + 30, top:=ws.Cells(41, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs" ' Assign the macro to run when the button is clicked
    End With
End Sub


Sub OutputMonteCarloSummaryMetrics(wsPortfolioMC As Worksheet, results As Variant, _
                                   startingEquity As Double, margin As Double, _
                                   yearsToConsider As Double, medianReturnToDrawdown As Double, _
                                   currentRiskOfRuin As Double)
    ' Function to output Monte Carlo simulation summary metrics to the PortfolioMC sheet
    ' Parameters:
    '   wsPortfolioMC - Worksheet to output to
    '   results - 2D array of simulation results
    '   startingEquity - Initial portfolio value
    '   margin - Minimum portfolio value
    '   yearsToConsider - Number of years in the input data
    '   medianReturnToDrawdown - Median return to drawdown ratio
    '   currentRiskOfRuin - Calculated risk of ruin
    
    Dim summaryRow As Long
    Dim percentiles As Variant
    Dim percentileLabels As Variant
    Dim i As Long, j As Long
    
    summaryRow = 3
    
    With wsPortfolioMC
        ' Set column widths for A to D
        .Columns("A:D").ColumnWidth = 16
        
        ' Title Formatting
        .Cells(1, 1).value = "Monte Carlo Simulation Summary"
        .Range(.Cells(1, 1), .Cells(1, 4)).Merge
        .Cells(1, 1).HorizontalAlignment = xlCenter
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Interior.Color = RGB(0, 102, 204) ' Light blue background for title
        .Cells(1, 1).Font.Color = RGB(255, 255, 255) ' White font for title
        
        ' Add border around the main title
        With .Range(.Cells(1, 1), .Cells(1, 4)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0) ' Black border
        End With
    
        ' Header row background for basic metrics
        .Range(.Cells(summaryRow, 1), .Cells(summaryRow + 4, 2)).Interior.Color = RGB(224, 224, 224) ' Light grey background for headers
        .Range(.Cells(summaryRow, 1), .Cells(summaryRow + 4, 2)).Font.Bold = True ' Bold header labels
        .Range(.Cells(summaryRow, 1), .Cells(summaryRow + 4, 2)).WrapText = True ' Wrap text
    
        ' Populate the basic metrics table with merged cells
        For i = 0 To 4
            ' Merge cells for each label
            .Range(.Cells(summaryRow + i, 1), .Cells(summaryRow + i, 2)).Merge
            .Cells(summaryRow + i, 1).HorizontalAlignment = xlLeft
        Next i
        
        .Cells(summaryRow, 1).value = "Starting Capital"
        .Cells(summaryRow, 3).value = startingEquity
        .Cells(summaryRow, 3).NumberFormat = "$#,##0"
        
        .Cells(summaryRow + 1, 1).value = "Minimum Portfolio Value"
        .Cells(summaryRow + 1, 3).value = margin
        .Cells(summaryRow + 1, 3).NumberFormat = "$#,##0"
        
        .Cells(summaryRow + 2, 1).value = "Input Data Period (last years)"
        .Cells(summaryRow + 2, 3).value = yearsToConsider & " years"
        .Cells(summaryRow + 2, 3).HorizontalAlignment = xlLeft
        
        .Cells(summaryRow + 3, 1).value = "Median Return to Drawdown"
        .Cells(summaryRow + 3, 3).value = medianReturnToDrawdown
        .Cells(summaryRow + 3, 3).NumberFormat = "0.0"
        
        .Cells(summaryRow + 4, 1).value = "Risk of Ruin"
        .Cells(summaryRow + 4, 3).value = currentRiskOfRuin
        .Cells(summaryRow + 4, 3).NumberFormat = "0.0%"
        
        ' Formatting for basic metrics value cells
        .Range(.Cells(summaryRow, 3), .Cells(summaryRow + 4, 3)).Interior.Color = RGB(242, 242, 242) ' Light grey for values
    
        ' Apply borders around the basic metrics table
        With .Range(.Cells(summaryRow, 1), .Cells(summaryRow + 4, 3)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0) ' Black border
        End With
        
        ' Add percentile tables title
        .Cells(summaryRow + 6, 1).value = "Percentile Distribution (Absolute Values)"
        .Cells(summaryRow + 6, 1).Font.Bold = True
        .Cells(summaryRow + 6, 1).Font.Size = 12
        .Range(.Cells(summaryRow + 6, 1), .Cells(summaryRow + 6, 4)).Merge
        .Cells(summaryRow + 6, 1).HorizontalAlignment = xlCenter
        .Cells(summaryRow + 6, 1).Interior.Color = RGB(0, 102, 204)
        .Cells(summaryRow + 6, 1).Font.Color = RGB(255, 255, 255)
        
        ' Add border around the percentile table title
        With .Range(.Cells(summaryRow + 6, 1), .Cells(summaryRow + 6, 4)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0) ' Black border
        End With
        
        ' Column headers for absolute values table
        .Cells(summaryRow + 7, 1).value = "Percentile"
        .Cells(summaryRow + 7, 2).value = "Profit ($)"
        .Cells(summaryRow + 7, 3).value = "Max Drawdown ($)"
        .Cells(summaryRow + 7, 4).value = "Average Drawdown ($)"
        .Range(.Cells(summaryRow + 7, 1), .Cells(summaryRow + 7, 4)).Interior.Color = RGB(224, 224, 224)
        .Range(.Cells(summaryRow + 7, 1), .Cells(summaryRow + 7, 4)).Font.Bold = True
        .Range(.Cells(summaryRow + 7, 1), .Cells(summaryRow + 7, 4)).WrapText = True
        
        ' Percentile rows for absolute values
        percentiles = Array(0.01, 0.05, 0.1, 0.25, 0.5, 0.75, 0.9, 0.95, 0.99)
        percentileLabels = Array("1%", "5%", "10%", "25%", "50%", "75%", "90%", "95%", "99%")
        
        ' Extract arrays for calculations
        Dim profitArray() As Double
        Dim maxDrawdownArray() As Double
        Dim avgDrawdownArray() As Double
        
        ReDim profitArray(1 To UBound(results, 1))
        ReDim maxDrawdownArray(1 To UBound(results, 1))
        ReDim avgDrawdownArray(1 To UBound(results, 1))
        
        For i = 1 To UBound(results, 1)
            profitArray(i) = results(i, 2)  ' Profit column
            maxDrawdownArray(i) = results(i, 5)  ' Max Drawdown column
            ' Assuming avg drawdown is in column 7, adjust if needed
            If UBound(results, 2) >= 7 Then
                avgDrawdownArray(i) = results(i, 7)
            Else
                ' If average drawdown isn't available, use max drawdown * 0.6 as an approximation
                avgDrawdownArray(i) = results(i, 5) * 0.6
            End If
        Next i
        
        ' Populate absolute values table
        For i = 0 To UBound(percentiles)
            j = i + 1
            
            ' If this is the 50% percentile (median), rename it and highlight it
            If percentiles(i) = 0.5 Then
                .Cells(summaryRow + 7 + j, 1).value = "Median"
                .Cells(summaryRow + 7 + j, 1).HorizontalAlignment = xlRight
                .Range(.Cells(summaryRow + 7 + j, 1), .Cells(summaryRow + 7 + j, 4)).Interior.Color = RGB(255, 255, 0) ' Yellow highlight
                .Range(.Cells(summaryRow + 7 + j, 1), .Cells(summaryRow + 7 + j, 4)).Font.Bold = True
                
            Else
                .Cells(summaryRow + 7 + j, 1).value = percentileLabels(i)
                .Range(.Cells(summaryRow + 7 + j, 1), .Cells(summaryRow + 7 + j, 4)).Interior.Color = RGB(242, 242, 242)
            End If
            
            ' Calculate percentile values
            .Cells(summaryRow + 7 + j, 2).value = Application.WorksheetFunction.percentile(profitArray, percentiles(i))
            .Cells(summaryRow + 7 + j, 2).NumberFormat = "$#,##0"
            
            .Cells(summaryRow + 7 + j, 3).value = Application.WorksheetFunction.percentile(maxDrawdownArray, percentiles(i))
            .Cells(summaryRow + 7 + j, 3).NumberFormat = "$#,##0"
            
            .Cells(summaryRow + 7 + j, 4).value = Application.WorksheetFunction.percentile(avgDrawdownArray, percentiles(i))
            .Cells(summaryRow + 7 + j, 4).NumberFormat = "$#,##0"
        Next i
        
        ' Add borders to absolute values table
        With .Range(.Cells(summaryRow + 7, 1), .Cells(summaryRow + 7 + UBound(percentiles) + 1, 4)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        
        ' Percentage of starting capital table title
        .Cells(summaryRow + 19, 1).value = "Percentile Distribution (Percentage of Starting Capital)"
        .Cells(summaryRow + 19, 1).Font.Bold = True
        .Cells(summaryRow + 19, 1).Font.Size = 12
        .Range(.Cells(summaryRow + 19, 1), .Cells(summaryRow + 19, 4)).Merge
        .Cells(summaryRow + 19, 1).HorizontalAlignment = xlCenter
        .Cells(summaryRow + 19, 1).Interior.Color = RGB(0, 102, 204)
        .Cells(summaryRow + 19, 1).Font.Color = RGB(255, 255, 255)
        
        ' Add border around the percentage table title
        With .Range(.Cells(summaryRow + 19, 1), .Cells(summaryRow + 19, 4)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0) ' Black border
        End With
        
        ' Column headers for percentage table
        .Cells(summaryRow + 20, 1).value = "Percentile"
        .Cells(summaryRow + 20, 2).value = "Profit (%)"
        .Cells(summaryRow + 20, 3).value = "Max Drawdown (%)"
        .Cells(summaryRow + 20, 4).value = "Average Drawdown (%)"
        .Range(.Cells(summaryRow + 20, 1), .Cells(summaryRow + 20, 4)).Interior.Color = RGB(224, 224, 224)
        .Range(.Cells(summaryRow + 20, 1), .Cells(summaryRow + 20, 4)).Font.Bold = True
        .Range(.Cells(summaryRow + 20, 1), .Cells(summaryRow + 20, 4)).WrapText = True
        
        ' Populate percentage table
        For i = 0 To UBound(percentiles)
            j = i + 1
            
            ' If this is the 50% percentile (median), rename it and highlight it
            If percentiles(i) = 0.5 Then
                .Cells(summaryRow + 20 + j, 1).value = "Median"
                .Cells(summaryRow + 20 + j, 1).HorizontalAlignment = xlRight
                .Range(.Cells(summaryRow + 20 + j, 1), .Cells(summaryRow + 20 + j, 4)).Interior.Color = RGB(255, 255, 0) ' Yellow highlight
                .Range(.Cells(summaryRow + 20 + j, 1), .Cells(summaryRow + 20 + j, 4)).Font.Bold = True
            Else
                .Cells(summaryRow + 20 + j, 1).value = percentileLabels(i)
                .Range(.Cells(summaryRow + 20 + j, 1), .Cells(summaryRow + 20 + j, 4)).Interior.Color = RGB(242, 242, 242)
            End If
            
            ' Calculate percentile values as percentage of starting capital
            .Cells(summaryRow + 20 + j, 2).value = Application.WorksheetFunction.percentile(profitArray, percentiles(i)) / startingEquity
            .Cells(summaryRow + 20 + j, 2).NumberFormat = "0.0%"
            
            .Cells(summaryRow + 20 + j, 3).value = Application.WorksheetFunction.percentile(maxDrawdownArray, percentiles(i)) / startingEquity
            .Cells(summaryRow + 20 + j, 3).NumberFormat = "0.0%"
            
            .Cells(summaryRow + 20 + j, 4).value = Application.WorksheetFunction.percentile(avgDrawdownArray, percentiles(i)) / startingEquity
            .Cells(summaryRow + 20 + j, 4).NumberFormat = "0.0%"
        Next i
        
        ' Add borders to percentage table
        With .Range(.Cells(summaryRow + 20, 1), .Cells(summaryRow + 20 + UBound(percentiles) + 1, 4)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End With
End Sub

Function RunMonteCarloWithTracking(pnlResults As Variant, requiredMargin As Double, _
                                averageTradesPerYear As Long, startingEquity As Double, _
                                numScenarios As Long, tradeAdjustment As Double, _
                                AverageTrade As Double, MCTradeType As String, _
                                ByRef dailyProfitTracking() As Double, _
                                ByRef dailyDrawdownTracking() As Double, _
                                ByRef dailyMaxDrawdownTracking() As Double, _
                                Optional randIdx As Variant) As Variant
                                
    ' Updated Monte Carlo simulation that handles both daily and weekly data
    ' Also tracks equity values over time for visualization
    ' Parameters:
    '   pnlResults - Historical PnL data (can be 1D array for weekly or 2D array for daily)
    '   requiredMargin - Minimum equity required (either fixed value or % of starting)
    '   averageTradesPerYear - Number of trades to simulate per year
    '   startingEquity - Initial portfolio equity
    '   numScenarios - Number of Monte Carlo scenarios to run
    '   tradeAdjustment - Adjustment factor for average trade
    '   averageTrade - Average trade value
    '   MCTradeType - "Daily" or "Weekly" sampling mode
    '   dailyProfitTracking - Output array for tracking equity over time (scenarios x days)
    
    Dim results() As Variant
    Dim i As Long, j As Long, K As Long
    Dim marginThreshold As Double
    Dim trade As Double
    Dim equity As Double
    Dim peakEquity As Double
    Dim maxDrawdown As Double
    Dim drawdown As Double
    Dim adjustedAvgTrade As Double
    Dim numStrategies As Long
    Dim isDaily As Boolean
    Dim strategyPnL As Double
    Dim avgDrawdown As Double
    Dim drawdowntotal As Double
    Dim randomperiod As Long
    
    ' Determine if we're using daily multi-strategy data or weekly single-strategy data
    isDaily = (MCTradeType = "Daily")
    
    ' For daily data, pnlResults is a 2D array (days x strategies)
    ' For weekly data, pnlResults is a 1D array of values
   
        ' Check if pnlResults is a 2D array
        On Error Resume Next
        numStrategies = UBound(pnlResults, 2)
        On Error GoTo 0
        
        If numStrategies = 0 Then
            ' Invalid format, exit function
            MsgBox "Invalid data format for daily PnL. Expected 2D array.", vbExclamation
            Exit Function
        End If
   
    Dim useRandIdx As Boolean
    useRandIdx = (Not IsEmpty(randIdx) And IsArray(randIdx))

    ' Initialize result arrays
    ReDim results(1 To numScenarios, 1 To 8)
   ' Initialize tracking arrays
    ReDim dailyProfitTracking(1 To numScenarios, 0 To averageTradesPerYear)
    ReDim dailyDrawdownTracking(1 To numScenarios, 0 To averageTradesPerYear)
    ReDim dailyMaxDrawdownTracking(1 To numScenarios, 0 To averageTradesPerYear)
    
    
    ' Set starting values for all scenarios (day 0)
    For i = 1 To numScenarios
        dailyProfitTracking(i, 0) = 0
        dailyDrawdownTracking(i, 0) = 0  ' No drawdown at start
        dailyMaxDrawdownTracking(i, 0) = 0  ' No drawdown at start
    Next i
    ' Calculate adjusted average trade based on trade adjustment
    adjustedAvgTrade = AverageTrade * (1 - tradeAdjustment)
    
    ' Determine margin threshold
    marginThreshold = requiredMargin
    
    ' Perform Monte Carlo simulation for each scenario
    For i = 1 To numScenarios
        equity = startingEquity
        peakEquity = startingEquity
        maxDrawdown = 0
        avgDrawdown = 0
        drawdowntotal = 0
        ' Simulate trades for the specified time period
        For j = 1 To averageTradesPerYear
            ' Reset trade PnL for this day
            trade = 0
                
                If useRandIdx Then
                    randomperiod = randIdx(i, j)
                Else
                    randomperiod = Int(Rnd * UBound(pnlResults, 1)) + 1
                End If
               
                
                ' Sum PnL across all strategies for the selected day
            For K = 1 To numStrategies
                strategyPnL = pnlResults(randomperiod, K)
                    
                    ' Apply trade adjustment
                trade = trade + strategyPnL
            Next K
                
                ' Apply overall trade adjustment
            trade = trade - adjustedAvgTrade
  
            
            ' Apply the trade PnL to equity
            equity = equity + trade
            
            ' Record the equity value for this day
            dailyProfitTracking(i, j) = equity - startingEquity
            

            
            
            ' Check and update peak equity and drawdown
            If equity > peakEquity Then
                peakEquity = equity
                drawdown = 0  ' If we're at a new peak, current drawdown is 0
            Else
                drawdown = (peakEquity - equity) / peakEquity
                If drawdown > maxDrawdown Then
                    maxDrawdown = drawdown
                End If
            End If
            
            drawdowntotal = drawdowntotal + drawdown
            
            avgDrawdown = drawdowntotal / j
            
            dailyDrawdownTracking(i, j) = avgDrawdown * peakEquity  ' Track current drawdown
            dailyMaxDrawdownTracking(i, j) = maxDrawdown * peakEquity  ' Track current drawdown
            
            ' Check for ruin
            If equity <= marginThreshold Then
                ' Mark as ruined
                results(i, 6) = 1
                
                ' Fill remaining days with the ruin value
                For K = j + 1 To averageTradesPerYear
                    dailyProfitTracking(i, K) = equity - startingEquity
                    dailyDrawdownTracking(i, K) = avgDrawdown * peakEquity   ' Track current drawdown
                    dailyMaxDrawdownTracking(i, K) = maxDrawdown * peakEquity   ' Track current drawdown
                Next K
                
                Exit For
            End If
        Next j
        
        ' Store results for this scenario
        results(i, 1) = equity                                 ' Final equity
        results(i, 2) = equity - startingEquity                ' Net profit
        results(i, 3) = (equity - startingEquity) / startingEquity  ' Return %
        results(i, 5) = maxDrawdown * peakEquity
        results(i, 4) = IIf(maxDrawdown > 0, results(i, 2) / results(i, 5), 0)   ' Return/Drawdown
        ' Store average drawdown and return/average drawdown ratio
        results(i, 7) = avgDrawdown * peakEquity               ' Average drawdown
        results(i, 8) = IIf(avgDrawdown > 0, results(i, 2) / results(i, 7), 0)   ' Return/Average Drawdown

        If results(i, 6) <> 1 Then results(i, 6) = 0           ' Ensure ruin flag is set
    Next i
    
    RunMonteCarloWithTracking = results
End Function




Sub CreateProfitProgressionChart(ws As Worksheet, dailyStats As Variant, chartTitle As String, row As Long, col As Long)
    ' Create a chart from the equity statistics data in the specified worksheet
    ' Shows median and percentiles: 1%, 10%, 25%, 75%, 90%, 99%
    '
    ' Parameters:
    '   ws - Worksheet where the chart will be placed
    '   dailyStats - Array containing daily statistics
    '   chartTitle - Title for the chart
    
    Dim chartObj As ChartObject
    Dim cht As chart
    Dim numDays As Long
    
    ' Determine number of data rows
    numDays = UBound(dailyStats, 1)
    
    ' Delete any existing equity progression charts
    On Error Resume Next
    For Each chartObj In ws.ChartObjects
        If InStr(chartObj.chart.chartTitle.text, chartTitle) > 0 Then
            chartObj.Delete
        End If
    Next chartObj
    On Error GoTo 0
    
    ' Position chart in a good location - to the right of the summary table, starting at row 3
    Set chartObj = ws.ChartObjects.Add(left:=ws.Cells(row, col).left, top:=ws.Cells(row, col).top, Width:=600, Height:=350)
    Set cht = chartObj.chart
    
    ' Set chart type
    cht.ChartType = xlLine
    
    ' Add data series
    With cht
        ' Add median series (make this prominent)
        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = "Median"
        .SeriesCollection(1).values = Application.index(dailyStats, 0, 3) ' Median (column 3)
        .SeriesCollection(1).XValues = Application.index(dailyStats, 0, 1) ' Day (column 1)
        .SeriesCollection(1).Format.line.Weight = 2
        .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(0, 176, 80) ' Green
        
        ' Add 1% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(2).name = "1% Percentile"
        .SeriesCollection(2).values = Application.index(dailyStats, 0, 9) ' 1% percentile (column 9)
        .SeriesCollection(2).XValues = Application.index(dailyStats, 0, 1) ' Day
        .SeriesCollection(2).Format.line.Weight = 2
        .SeriesCollection(2).Format.line.DashStyle = msoLineDash
        .SeriesCollection(2).Format.line.ForeColor.RGB = RGB(192, 0, 0) ' Dark Red
        
        ' Add 99% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(3).name = "5% Percentile"
        .SeriesCollection(3).values = Application.index(dailyStats, 0, 13) ' 5th percentile (column 8)
        .SeriesCollection(3).XValues = Application.index(dailyStats, 0, 1) ' Day (column 1)
        .SeriesCollection(3).Format.line.Weight = 2
        .SeriesCollection(3).Format.line.DashStyle = msoLineDash
        .SeriesCollection(3).Format.line.ForeColor.RGB = RGB(50, 100, 160) ' Purple
        
        
        ' Add 10% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(4).name = "10% Percentile"
        .SeriesCollection(4).values = Application.index(dailyStats, 0, 4) ' 10th percentile (column 4)
        .SeriesCollection(4).XValues = Application.index(dailyStats, 0, 1) ' Day (column 1)
        .SeriesCollection(4).Format.line.Weight = 2
        .SeriesCollection(4).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(4).Format.line.ForeColor.RGB = RGB(255, 102, 0) ' Orange
        
        ' Add 25% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(5).name = "25% Percentile"
        .SeriesCollection(5).values = Application.index(dailyStats, 0, 5) ' 25th percentile (column 5)
        .SeriesCollection(5).XValues = Application.index(dailyStats, 0, 1) ' Day (column 1)
        .SeriesCollection(5).Format.line.Weight = 1.5
        .SeriesCollection(5).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(5).Format.line.ForeColor.RGB = RGB(255, 192, 0) ' Gold
        
        ' Add 75% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(6).name = "75% Percentile"
        .SeriesCollection(6).values = Application.index(dailyStats, 0, 6) ' 75th percentile (column 6)
        .SeriesCollection(6).XValues = Application.index(dailyStats, 0, 1) ' Day (column 1)
        .SeriesCollection(6).Format.line.Weight = 1.5
        .SeriesCollection(6).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(6).Format.line.ForeColor.RGB = RGB(112, 173, 71) ' Light Green
        
        ' Add 90% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(7).name = "90% Percentile"
        .SeriesCollection(7).values = Application.index(dailyStats, 0, 7) ' 90th percentile (column 7)
        .SeriesCollection(7).XValues = Application.index(dailyStats, 0, 1) ' Day (column 1)
        .SeriesCollection(7).Format.line.Weight = 2
        .SeriesCollection(7).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(7).Format.line.ForeColor.RGB = RGB(0, 112, 192) ' Blue
    
        ' Add 90% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(8).name = "95% Percentile"
        .SeriesCollection(8).values = Application.index(dailyStats, 0, 12) ' 95th percentile (column 7)
        .SeriesCollection(8).XValues = Application.index(dailyStats, 0, 1) ' Day (column 1)
        .SeriesCollection(8).Format.line.Weight = 2
        .SeriesCollection(8).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(8).Format.line.ForeColor.RGB = RGB(0, 112, 100) ' Blue
    
        ' Add 99% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(9).name = "99% Percentile"
        .SeriesCollection(9).values = Application.index(dailyStats, 0, 8) ' 99th percentile (column 8)
        .SeriesCollection(9).XValues = Application.index(dailyStats, 0, 1) ' Day (column 1)
        .SeriesCollection(9).Format.line.Weight = 2
        .SeriesCollection(9).Format.line.DashStyle = msoLineDash
        .SeriesCollection(9).Format.line.ForeColor.RGB = RGB(112, 48, 160) ' Purple
        
        
        
        

        ' Format chart
        .HasTitle = True
        .chartTitle.text = chartTitle
        .chartTitle.Font.Size = 14
        .chartTitle.Font.Bold = True
        
        ' Add axis titles
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.text = "Trading Day"
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.text = "Portfolio Profit ($)"
        
        ' Format axes
        .Axes(xlCategory).MajorGridlines.Format.line.Visible = msoFalse
        .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(200, 200, 200)
        
        ' Auto-scale Y-axis to fit data (adjust to include the data range with a small margin)
        Dim minValue As Double, maxValue As Double
        Dim startingEquity As Double
        
        ' Get starting equity value
       ' startingEquity = dailyStats(0, 2)  ' Starting equity from average column
        
        ' Find min and max values
        minValue = Application.WorksheetFunction.Min(Application.index(dailyStats, 0, 10))  ' Min values (column 10)
        maxValue = Application.WorksheetFunction.Max(Application.index(dailyStats, 0, 11)) ' Max values (column 11)
        
    
        
        ' Add a margin to make the chart look better
        Dim yRange As Double
        yRange = maxValue - minValue
        minValue = minValue - (yRange * 0.05)  ' 5% margin on bottom
        maxValue = maxValue + (yRange * 0.05)  ' 5% margin on top
        
        ' Set axis scales
        '.Axes(xlValue).MinimumScale = minValue
        '.Axes(xlValue).MaximumScale = maxValue
        
        ' Format plot area
        .PlotArea.Format.fill.ForeColor.RGB = RGB(242, 242, 242)
        
        ' Format to show currency on Y axis
        .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
        
        ' Add legend
        .HasLegend = True
        .Legend.position = xlLegendPositionBottom
        
        ' Format grid lines
        .Axes(xlValue).MajorGridlines.Format.line.DashStyle = msoLineDash
        .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(200, 200, 200)
    End With
End Sub

Sub CreateDistributionChart(ws As Worksheet, dailyProfitTracking() As Double, tradingDay As Long, chartTitle As String, left As Double, top As Double)
    ' Create a histogram showing the distribution of equity values on a specific day
    '
    ' Parameters:
    '   ws - Worksheet where the chart will be placed
    '   dailyProfitTracking - Array with equity tracking data (scenarios x days)
    '   tradingDay - The trading day to analyze
    '   chartTitle - Title for the chart
    '   left - Left position for the chart
    '   top - Top position for the chart
    
    Dim chartObj As ChartObject
    Dim cht As chart
    Dim numScenarios As Long
    Dim i As Long, bin As Long
    Dim minValue As Double, maxValue As Double
    Dim numBins As Long, binWidth As Double
    Dim tempValues() As Double
    Dim binCounts() As Long
    Dim binEdges() As Double
    
    ' Get dimensions
    numScenarios = UBound(dailyProfitTracking, 1)
    
    ' Delete any existing distribution charts for this day
    On Error Resume Next
    For Each chartObj In ws.ChartObjects
        If InStr(chartObj.chart.chartTitle.text, "Distribution Day " & tradingDay) > 0 Then
            chartObj.Delete
        End If
    Next chartObj
    On Error GoTo 0
    
    ' Extract values for this day
    ReDim tempValues(1 To numScenarios)
    For i = 1 To numScenarios
        tempValues(i) = dailyProfitTracking(i, tradingDay)
    Next i
    
    ' Find min and max values
    minValue = Application.WorksheetFunction.Min(tempValues)
    maxValue = Application.WorksheetFunction.Max(tempValues)
    
    ' Create bins for histogram
    numBins = 15 ' Number of bins for distribution
    binWidth = (maxValue - minValue) / numBins
    
    If binWidth <= 0 Then
        ' Handle case where all values are the same
        binWidth = 1
        numBins = 1
    End If
    
    ReDim binCounts(1 To numBins)
    ReDim binEdges(0 To numBins)
    
    ' Set bin edges
    For i = 0 To numBins
        binEdges(i) = minValue + (i * binWidth)
    Next i
    
    ' Count values in each bin
    For i = 1 To numScenarios
        ' Find which bin this value belongs to
        For bin = 1 To numBins
            If tempValues(i) < binEdges(bin) Or bin = numBins Then
                binCounts(bin) = binCounts(bin) + 1
                Exit For
            End If
        Next bin
    Next i
    
    ' Create temp range for histogram data
    Dim tempRange As Range
    Dim DataRange As Range
    
    ' Use cells at far right of worksheet for temporary data storage (column 100+)
    Set tempRange = ws.Range(ws.Cells(1, 100), ws.Cells(numBins + 1, 101))
    tempRange.Clear
    
    ' Add bin edges and counts
    For i = 1 To numBins
        ws.Cells(i, 100).value = binEdges(i - 1) ' Bin start
        ws.Cells(i, 101).value = binCounts(i)    ' Count
    Next i
    
    Set DataRange = ws.Range(ws.Cells(1, 100), ws.Cells(numBins, 101))
    
    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left:=left, top:=top, Width:=350, Height:=250)
    Set cht = chartObj.chart
    
    ' Set chart type and data
    With cht
        .ChartType = xlColumnClustered
        
        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = "Scenarios"
        .SeriesCollection(1).values = ws.Range(ws.Cells(1, 101), ws.Cells(numBins, 101))
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(1, 100), ws.Cells(numBins, 100))
        
        ' Format series
        .SeriesCollection(1).Format.fill.ForeColor.RGB = RGB(91, 155, 213) ' Blue columns
        
        ' Format chart
        .HasTitle = True
        .chartTitle.text = chartTitle & " (Day " & tradingDay & ")"
        .chartTitle.Font.Size = 12
        .chartTitle.Font.Bold = True
        
        ' Format axes
        .Axes(xlCategory).TickLabels.NumberFormat = "$#,##0"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.text = "Equity Value"
        
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.text = "Number of Scenarios"
        
        ' Format plot area
        .PlotArea.Format.fill.ForeColor.RGB = RGB(242, 242, 242)
        
        ' No legend needed
        .HasLegend = False
    End With
    
    ' Clear temporary data
    tempRange.Clear
End Sub


Sub CreateProfitSamplePathsChart(ws As Worksheet, dailyProfitTracking() As Double, chartTitle As String, numPaths As Long, row As Long, col As Long)
    ' Create a chart showing randomly selected sample paths from the Monte Carlo simulation
    '
    ' Parameters:
    '   ws - Worksheet where the chart will be placed
    '   dailyProfitTracking - Array with equity tracking data (scenarios x days)
    '   chartTitle - Title for the chart
    '   numPaths - Number of sample paths to display (recommended: 10-20)
    
    Dim chartObj As ChartObject
    Dim cht As chart
    Dim numScenarios As Long, numDays As Long
    Dim i As Long, j As Long, pathIndex As Long
    Dim randomPaths() As Long
    Dim minValue As Double, maxValue As Double
    Dim dayLabels() As Long
    
    ' Get dimensions
    numScenarios = UBound(dailyProfitTracking, 1)
    numDays = UBound(dailyProfitTracking, 2)
    
    ' Create array of day labels (0 to numDays)
    ReDim dayLabels(0 To numDays)
    For i = 0 To numDays
        dayLabels(i) = i
    Next i
    
    ' Select random paths (without replacement)
    ReDim randomPaths(1 To numPaths)
    Randomize
    
    For i = 1 To numPaths
        ' Generate unique random path indices
        Do
            pathIndex = Int(Rnd() * numScenarios) + 1
            ' Check if already selected
            Dim isDuplicate As Boolean
            isDuplicate = False
            For j = 1 To i - 1
                If randomPaths(j) = pathIndex Then
                    isDuplicate = True
                    Exit For
                End If
            Next j
        Loop While isDuplicate
        
        randomPaths(i) = pathIndex
    Next i
    
    ' Delete any existing sample path charts
    On Error Resume Next
    For Each chartObj In ws.ChartObjects
        If InStr(chartObj.chart.chartTitle.text, chartTitle) > 0 Then
            chartObj.Delete
        End If
    Next chartObj
    On Error GoTo 0
    
    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left:=ws.Cells(row, col).left, top:=ws.Cells(row, col).top, Width:=600, Height:=350)
    Set cht = chartObj.chart
    
    ' Set chart type
    cht.ChartType = xlLine
    
    ' Add data series for each random path
    With cht
        ' Add each random path as a series
        For i = 1 To numPaths
            .SeriesCollection.NewSeries
            .SeriesCollection(i).name = "Path " & i
            
            ' Extract equity values for this path
            Dim pathValues() As Double
            ReDim pathValues(0 To numDays)
            For j = 0 To numDays
                pathValues(j) = dailyProfitTracking(randomPaths(i), j)
            Next j
            
            .SeriesCollection(i).values = pathValues
            .SeriesCollection(i).XValues = dayLabels
            
            ' Format line with semi-transparent colors
            .SeriesCollection(i).Format.line.Weight = 1.5
            
            ' Create different colors for different paths
            Dim r As Byte, g As Byte, b As Byte
            r = 50 + Int(Rnd() * 205)  ' Random RGB values
            g = 50 + Int(Rnd() * 205)
            b = 50 + Int(Rnd() * 205)
            
            .SeriesCollection(i).Format.line.ForeColor.RGB = RGB(r, g, b)
            .SeriesCollection(i).Format.line.Transparency = 0.5 ' 50% transparency
        Next i
        
             
        ' Format chart
        .HasTitle = True
        .chartTitle.text = chartTitle
        .chartTitle.Font.Size = 14
        .chartTitle.Font.Bold = True
        
        ' Add axis titles
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.text = "Trading Day"
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.text = "Portfolio Profit ($)"
        
        ' Format axes
        .Axes(xlCategory).MajorGridlines.Format.line.Visible = msoFalse
        .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(200, 200, 200)
        
        ' Find min and max values across all selected paths
        minValue = dailyProfitTracking(randomPaths(1), 0)
        maxValue = minValue
        Dim startingEquity As Double
        startingEquity = dailyProfitTracking(1, 0)  ' Get starting equity
        
        For i = 1 To numPaths
            For j = 0 To numDays
                If dailyProfitTracking(randomPaths(i), j) < minValue Then
                    minValue = dailyProfitTracking(randomPaths(i), j)
                ElseIf dailyProfitTracking(randomPaths(i), j) > maxValue Then
                    maxValue = dailyProfitTracking(randomPaths(i), j)
                End If
            Next j
        Next i
        
        ' Ensure starting equity is included in the range
      '  If startingEquity < minValue Then
      '      minValue = startingEquity
      '  End If
      '  If startingEquity > maxValue Then
      '      maxValue = startingEquity
      '  End If
        
        ' Add a margin to make the chart look better
        Dim yRange As Double
        yRange = maxValue - minValue
        minValue = minValue - (yRange * 0.05)  ' 5% margin on bottom
        maxValue = maxValue + (yRange * 0.05)  ' 5% margin on top
        
        ' Set axis scales
       ' .Axes(xlValue).MinimumScale = minValue
       ' .Axes(xlValue).MaximumScale = maxValue
        
        ' Format plot area
        .PlotArea.Format.fill.ForeColor.RGB = RGB(242, 242, 242)
        
        ' Format to show currency on Y axis
        .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
        
        ' Add legend
        .HasLegend = False  ' Hide legend since there are many paths
    End With
End Sub

Sub CreateHistogram(ws As Worksheet, values As Variant, title As String, numBins As Long, startCol As Long, summaryRow As Long, chartHeight As Long, roundToNearest As Double)
    Dim i As Long
    Dim minValue As Double, maxValue As Double, binSize As Double
    Dim binCount As Object
    Dim binKey As Variant
    Dim startRow As Long
    Dim chart As ChartObject
    Dim displayFormat As String
    Dim totalCount As Long
    
    Set binCount = CreateObject("Scripting.Dictionary")
    startRow = summaryRow
    totalCount = UBound(values) - LBound(values) + 1

    ' Determine currency display format and unit based on the rounding parameter
    If roundToNearest >= 1 Then
        If roundToNearest >= 1000000 Then
            displayFormat = "$#,##0,,\M"  ' Millions
            unitLabel = "M"
        ElseIf roundToNearest >= 1000 Then
            displayFormat = "$#,##0,K"  ' Thousands
            unitLabel = "K"
        Else
            displayFormat = "$#,##0"  ' Basic currency
            unitLabel = ""
        End If
    Else
        displayFormat = "0%"  ' Percentage format
        unitLabel = "%"
    End If

    ' Calculate minimum and maximum values to define bin range
    minValue = Application.WorksheetFunction.Min(values)
    maxValue = Application.WorksheetFunction.Max(values)
    
    ' Calculate bin size for 10 bins and round to the nearest specified value
    binSize = (maxValue - minValue) / numBins
    binSize = Application.WorksheetFunction.Round(binSize / roundToNearest, 0) * roundToNearest
    binSize = Application.WorksheetFunction.Max(binSize, roundToNearest)
    
    ' Initialize bins with rounded bin size
    For i = 0 To numBins - 1
        Dim binLower As Double
        binLower = minValue + (i * binSize)
        binLower = Application.WorksheetFunction.Round(binLower / roundToNearest, 0) * roundToNearest
        binCount(binLower) = 0
    Next i

    ' Populate bin counts based on values
    For i = LBound(values) To UBound(values)
        If Not IsNumeric(values(i)) Then
            MsgBox "Non-numeric value found: " & values(i), vbExclamation
            Exit Sub
        End If
        
        Dim bin As Double
        bin = Application.WorksheetFunction.RoundDown((values(i) - minValue) / binSize, 0) * binSize + minValue
        bin = Application.WorksheetFunction.Round(bin / roundToNearest, 0) * roundToNearest
        If bin >= minValue And bin <= maxValue Then
            If Not binCount.Exists(bin) Then binCount(bin) = 0
            binCount(bin) = binCount(bin) + 1
        End If
    Next i

    ' Output histogram title and headers
    ws.Cells(2, startCol).value = title
    ws.Cells(startRow, startCol).value = "Bin Range"
    ws.Cells(startRow, startCol + 1).value = "Percentage"
    startRow = startRow + 1

    ' Display bins with percentages
    For Each binKey In binCount.keys
        ' Format the bin range display
        If roundToNearest >= 1 Then
            ws.Cells(startRow, startCol).value = Format(binKey, displayFormat) & " to " & Format(binKey + binSize, displayFormat)
        Else
            ws.Cells(startRow, startCol).value = Format(binKey * 100, "0") & "% to " & Format((binKey + binSize) * 100, "0") & "%"
        End If
        
        ' Percentage value
        ws.Cells(startRow, startCol + 1).value = binCount(binKey) / totalCount
        ws.Cells(startRow, startCol + 1).NumberFormat = "0.0%"
        
        startRow = startRow + 1
    Next binKey

    ' Create the histogram chart
    Set chart = ws.ChartObjects.Add(left:=ws.Cells(chartHeight, 6).left, Width:=440, top:=ws.Cells(chartHeight, 6).top, Height:=220)
    With chart.chart
        .SetSourceData source:=ws.Range(ws.Cells(summaryRow + 1, startCol), ws.Cells(startRow - 1, startCol + 1))
        .ChartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.text = title
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.text = "Bin Range"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.text = "Percentage"
        .Axes(xlValue).TickLabels.NumberFormat = "0.0%"
        
        ' Format data labels
        With .SeriesCollection(1)
            .HasDataLabels = True
            .DataLabels.position = xlLabelPositionOutsideEnd
            .DataLabels.NumberFormat = "0.0%"
        End With
        
        .HasLegend = False
    End With
End Sub




Sub CreateDrawdownProgressionChart(ws As Worksheet, drawdownStats As Variant, chartTitle As String, row As Long, col As Long)
    ' Create a chart from the drawdown statistics data in the specified worksheet
    ' Shows median and percentiles: 1%, 10%, 25%, 75%, 90%, 99%
    '
    ' Parameters:
    '   ws - Worksheet where the chart will be placed
    '   drawdownStats - Array containing daily statistics
    '   chartTitle - Title for the chart
    
    Dim chartObj As ChartObject
    Dim cht As chart
    Dim numDays As Long
    
    ' Determine number of data rows
    numDays = UBound(drawdownStats, 1)
    
    ' Delete any existing drawdown progression charts
    On Error Resume Next
    For Each chartObj In ws.ChartObjects
        If InStr(chartObj.chart.chartTitle.text, chartTitle) > 0 Then
            chartObj.Delete
        End If
    Next chartObj
    On Error GoTo 0
    
    ' Position chart below equity progression chart
    Set chartObj = ws.ChartObjects.Add(left:=ws.Cells(row, col).left, top:=ws.Cells(row, col).top, Width:=600, Height:=350)
    Set cht = chartObj.chart
    
    ' Set chart type
    cht.ChartType = xlLine
    
    ' Add data series
    With cht
         ' Add median series (make this prominent)
        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = "Median"
        .SeriesCollection(1).values = Application.index(drawdownStats, 0, 3) ' Median (column 3)
        .SeriesCollection(1).XValues = Application.index(drawdownStats, 0, 1) ' Day (column 1)
        .SeriesCollection(1).Format.line.Weight = 2
        .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(0, 176, 80) ' Green
        
        ' Add 1% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(2).name = "1% Percentile"
        .SeriesCollection(2).values = Application.index(drawdownStats, 0, 9) ' 1% percentile (column 9)
        .SeriesCollection(2).XValues = Application.index(drawdownStats, 0, 1) ' Day
        .SeriesCollection(2).Format.line.Weight = 2
        .SeriesCollection(2).Format.line.DashStyle = msoLineDash
        .SeriesCollection(2).Format.line.ForeColor.RGB = RGB(192, 0, 0) ' Dark Red
        
        ' Add 99% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(3).name = "5% Percentile"
        .SeriesCollection(3).values = Application.index(drawdownStats, 0, 13) ' 5th percentile (column 8)
        .SeriesCollection(3).XValues = Application.index(drawdownStats, 0, 1) ' Day (column 1)
        .SeriesCollection(3).Format.line.Weight = 2
        .SeriesCollection(3).Format.line.DashStyle = msoLineDash
        .SeriesCollection(3).Format.line.ForeColor.RGB = RGB(50, 100, 160) ' Purple
        
        
        ' Add 10% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(4).name = "10% Percentile"
        .SeriesCollection(4).values = Application.index(drawdownStats, 0, 4) ' 10th percentile (column 4)
        .SeriesCollection(4).XValues = Application.index(drawdownStats, 0, 1) ' Day (column 1)
        .SeriesCollection(4).Format.line.Weight = 2
        .SeriesCollection(4).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(4).Format.line.ForeColor.RGB = RGB(255, 102, 0) ' Orange
        
        ' Add 25% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(5).name = "25% Percentile"
        .SeriesCollection(5).values = Application.index(drawdownStats, 0, 5) ' 25th percentile (column 5)
        .SeriesCollection(5).XValues = Application.index(drawdownStats, 0, 1) ' Day (column 1)
        .SeriesCollection(5).Format.line.Weight = 1.5
        .SeriesCollection(5).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(5).Format.line.ForeColor.RGB = RGB(255, 192, 0) ' Gold
        
        ' Add 75% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(6).name = "75% Percentile"
        .SeriesCollection(6).values = Application.index(drawdownStats, 0, 6) ' 75th percentile (column 6)
        .SeriesCollection(6).XValues = Application.index(drawdownStats, 0, 1) ' Day (column 1)
        .SeriesCollection(6).Format.line.Weight = 1.5
        .SeriesCollection(6).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(6).Format.line.ForeColor.RGB = RGB(112, 173, 71) ' Light Green
        
        ' Add 90% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(7).name = "90% Percentile"
        .SeriesCollection(7).values = Application.index(drawdownStats, 0, 7) ' 90th percentile (column 7)
        .SeriesCollection(7).XValues = Application.index(drawdownStats, 0, 1) ' Day (column 1)
        .SeriesCollection(7).Format.line.Weight = 2
        .SeriesCollection(7).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(7).Format.line.ForeColor.RGB = RGB(0, 112, 192) ' Blue
    
        ' Add 90% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(8).name = "95% Percentile"
        .SeriesCollection(8).values = Application.index(drawdownStats, 0, 12) ' 95th percentile (column 7)
        .SeriesCollection(8).XValues = Application.index(drawdownStats, 0, 1) ' Day (column 1)
        .SeriesCollection(8).Format.line.Weight = 2
        .SeriesCollection(8).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(8).Format.line.ForeColor.RGB = RGB(0, 112, 100) ' Blue
    
        ' Add 99% percentile series
        .SeriesCollection.NewSeries
        .SeriesCollection(9).name = "99% Percentile"
        .SeriesCollection(9).values = Application.index(drawdownStats, 0, 8) ' 99th percentile (column 8)
        .SeriesCollection(9).XValues = Application.index(drawdownStats, 0, 1) ' Day (column 1)
        .SeriesCollection(9).Format.line.Weight = 2
        .SeriesCollection(9).Format.line.DashStyle = msoLineDash
        .SeriesCollection(9).Format.line.ForeColor.RGB = RGB(112, 48, 160) ' Purple

        
        ' Format chart
        .HasTitle = True
        .chartTitle.text = chartTitle
        .chartTitle.Font.Size = 14
        .chartTitle.Font.Bold = True
        
        ' Add axis titles
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.text = "Trading Day"
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.text = "Drawdown ($)"
        
        ' Format axes
        .Axes(xlCategory).MajorGridlines.Format.line.Visible = msoFalse
        .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(200, 200, 200)
        
        ' Auto-scale Y-axis
        Dim minValue As Double, maxValue As Double
        
        ' Find min and max values
        minValue = Application.WorksheetFunction.Min(Application.index(drawdownStats, 0, 10))  ' Min values
        maxValue = Application.WorksheetFunction.Max(Application.index(drawdownStats, 0, 11)) ' Max values
        
        ' Add a margin to make the chart look better
        Dim yRange As Double
        yRange = maxValue - minValue
        minValue = minValue - (yRange * 0.05)  ' 5% margin on bottom
        maxValue = maxValue + (yRange * 0.05)  ' 5% margin on top
        
        ' Set axis scales and invert for drawdowns
       ' .Axes(xlValue).MinimumScale = minValue
       ' .Axes(xlValue).MaximumScale = maxValue
        '.Axes(xlValue).Reverse = True  ' Correct syntax for inverting axis
    
        ' Format plot area
        .PlotArea.Format.fill.ForeColor.RGB = RGB(242, 242, 242)
        
        ' Format to show percentage on Y axis
        .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
        
        ' Add legend
        .HasLegend = True
        .Legend.position = xlLegendPositionBottom
        
        ' Format grid lines
        .Axes(xlValue).MajorGridlines.Format.line.DashStyle = msoLineDash
        .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(200, 200, 200)
    End With
End Sub


Sub CreateDrawdownSamplePathsChart(ws As Worksheet, dailyDrawdownTracking() As Double, chartTitle As String, numPaths As Long, row As Long, col As Long)
    ' Create a chart showing randomly selected sample drawdown paths
    '
    ' Parameters:
    '   ws - Worksheet where the chart will be placed
    '   dailyDrawdownTracking - Array with drawdown tracking data (scenarios x days)
    '   chartTitle - Title for the chart
    '   numPaths - Number of sample paths to display
    
    Dim chartObj As ChartObject
    Dim cht As chart
    Dim numScenarios As Long, numDays As Long
    Dim i As Long, j As Long, pathIndex As Long
    Dim randomPaths() As Long
    Dim minValue As Double, maxValue As Double
    Dim dayLabels() As Long
    
    ' Get dimensions
    numScenarios = UBound(dailyDrawdownTracking, 1)
    numDays = UBound(dailyDrawdownTracking, 2)
    
    ' Create array of day labels (0 to numDays)
    ReDim dayLabels(0 To numDays)
    For i = 0 To numDays
        dayLabels(i) = i
    Next i
    
    ' Select random paths (without replacement)
    ReDim randomPaths(1 To numPaths)
    Randomize
    
    For i = 1 To numPaths
        ' Generate unique random path indices
        Do
            pathIndex = Int(Rnd() * numScenarios) + 1
            ' Check if already selected
            Dim isDuplicate As Boolean
            isDuplicate = False
            For j = 1 To i - 1
                If randomPaths(j) = pathIndex Then
                    isDuplicate = True
                    Exit For
                End If
            Next j
        Loop While isDuplicate
        
        randomPaths(i) = pathIndex
    Next i
    
    ' Delete any existing sample path charts
    On Error Resume Next
    For Each chartObj In ws.ChartObjects
        If InStr(chartObj.chart.chartTitle.text, "Sample Drawdown Paths") > 0 Then
            chartObj.Delete
        End If
    Next chartObj
    On Error GoTo 0
    
    ' Create chart
    Set chartObj = ws.ChartObjects.Add(left:=ws.Cells(row, col).left, top:=ws.Cells(row, col).top, Width:=600, Height:=350)
    Set cht = chartObj.chart
    
    ' Set chart type
    cht.ChartType = xlLine
    
    ' Add data series for each random path
    With cht
        ' Add each random path as a series
        For i = 1 To numPaths
            .SeriesCollection.NewSeries
            .SeriesCollection(i).name = "Path " & i
            
            ' Extract drawdown values for this path
            Dim pathValues() As Double
            ReDim pathValues(0 To numDays)
            For j = 0 To numDays
                pathValues(j) = dailyDrawdownTracking(randomPaths(i), j)
            Next j
            
            .SeriesCollection(i).values = pathValues
            .SeriesCollection(i).XValues = dayLabels
            
            ' Format line with semi-transparent red colors (since drawdowns are negative)
            .SeriesCollection(i).Format.line.Weight = 1.5
            
            ' Create different shades of red for different paths
            Dim r As Byte, g As Byte, b As Byte
            r = 205 + Int(Rnd() * 50)  ' Bright red base
            g = Int(Rnd() * 100)       ' Limited green
            b = Int(Rnd() * 100)       ' Limited blue
            
            .SeriesCollection(i).Format.line.ForeColor.RGB = RGB(r, g, b)
            .SeriesCollection(i).Format.line.Transparency = 0.5 ' 50% transparency
        Next i
        
        ' Add reference line for zero drawdown
        .SeriesCollection.NewSeries
        .SeriesCollection(numPaths + 1).name = "Zero Drawdown"
        .SeriesCollection(numPaths + 1).values = Array(0, 0) ' Zero drawdown line
        .SeriesCollection(numPaths + 1).XValues = Array(0, numDays)
        .SeriesCollection(numPaths + 1).Format.line.Weight = 2
        .SeriesCollection(numPaths + 1).Format.line.DashStyle = msoLineDashDot
        .SeriesCollection(numPaths + 1).Format.line.ForeColor.RGB = RGB(0, 176, 80) ' Green
        
        ' Format chart
        .HasTitle = True
        .chartTitle.text = chartTitle
        .chartTitle.Font.Size = 14
        .chartTitle.Font.Bold = True
        
        ' Add axis titles
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.text = "Trading Day"
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.text = "Drawdown ($)"
        
        ' Format axes
        .Axes(xlCategory).MajorGridlines.Format.line.Visible = msoFalse
        .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(200, 200, 200)
        
        ' Find min and max values across all selected paths
        minValue = 0  ' Drawdowns start at 0
        maxValue = 0
        
        For i = 1 To numPaths
            For j = 0 To numDays
                If dailyDrawdownTracking(randomPaths(i), j) > maxValue Then
                    maxValue = dailyDrawdownTracking(randomPaths(i), j)
                End If
            Next j
        Next i
        
        ' Add a margin to make the chart look better
        maxValue = maxValue * 1.05  ' 5% margin
        
        ' Set axis scales and invert the axis (since drawdowns are typically shown as negative)
       ' .Axes(xlValue).MinimumScale = minValue
      '  .Axes(xlValue).MaximumScale = maxValue
       ' .Axes(xlValue).ReversePlot = True  ' Invert the axis
        
        ' Format plot area
        .PlotArea.Format.fill.ForeColor.RGB = RGB(242, 242, 242)
        
        ' Format to show percentage on Y axis
        .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
        
        ' No legend needed since there are many paths
        .HasLegend = False
        
        ' Format grid lines
        .Axes(xlValue).MajorGridlines.Format.line.DashStyle = msoLineDash
        .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(200, 200, 200)
    End With
End Sub

Function RemoveTopPerformers(pnlResults As Variant, removeTopPercent As Double, MCTradeType As String) As Variant
    ' Removes specified percentage of highest positive PnL periods (days or weeks)
    ' Parameters:
    '   pnlResults - Original PnL data (1D or 2D array)
    '   removeTopPercent - Percentage of top performers to remove (0-100)
    '   MCTradeType - "Daily" or "Weekly"
    ' Returns:
    '   Modified PnL data with top performers completely removed (array size reduced)
    
    Dim i As Long, j As Long, K As Long
    
    ' If no removal requested, return original data
    If removeTopPercent <= 0 Then
        RemoveTopPerformers = pnlResults
        Exit Function
    End If
    
    ' For 2D array (periods × strategies)
    Dim numPeriods As Long, numStrategies As Long
    numPeriods = UBound(pnlResults, 1)
    numStrategies = UBound(pnlResults, 2)
    
    ' Calculate period totals
    Dim periodTotals() As Double
    Dim periodIndices() As Long
    ReDim periodTotals(1 To numPeriods)
    ReDim periodIndices(1 To numPeriods)
    
    ' Sum each period's PnL and store original indices
    For i = 1 To numPeriods
        periodTotals(i) = 0
        periodIndices(i) = i
        For j = 1 To numStrategies
            periodTotals(i) = periodTotals(i) + pnlResults(i, j)
        Next j
    Next i
    
    ' Calculate number of positive periods to remove
    Dim positivePeriods As Long, periodsToRemove As Long
    For i = 1 To numPeriods
        If periodTotals(i) > 0 Then positivePeriods = positivePeriods + 1
    Next i
    
    periodsToRemove = Int(positivePeriods * (removeTopPercent))
    
    If periodsToRemove > 0 Then
        ' Sort period totals in descending order while keeping track of original indices
        Dim tempTotal As Double, tempIndex As Long
        For i = 1 To numPeriods - 1
            For j = i + 1 To numPeriods
                If periodTotals(i) < periodTotals(j) Then
                    ' Swap totals
                    tempTotal = periodTotals(i)
                    periodTotals(i) = periodTotals(j)
                    periodTotals(j) = tempTotal
                    
                    ' Swap indices
                    tempIndex = periodIndices(i)
                    periodIndices(i) = periodIndices(j)
                    periodIndices(j) = tempIndex
                End If
            Next j
        Next i
        
        ' Debug output of top values being removed
        Debug.Print "Removing top " & periodsToRemove & " " & MCTradeType & " periods out of " & positivePeriods & " positive periods"
        Debug.Print "Top 5 " & MCTradeType & " PnLs being removed:"
        For i = 1 To Application.WorksheetFunction.Min(5, periodsToRemove)
            Debug.Print MCTradeType & " " & periodIndices(i) & ": $" & Format(periodTotals(i), "#,##0")
        Next i
        
        ' Create a new smaller array excluding top performers
        Dim newNumPeriods As Long
        newNumPeriods = numPeriods - periodsToRemove
        Dim reducedPnL As Variant
        ReDim reducedPnL(1 To newNumPeriods, 1 To numStrategies)
        
        ' Create a lookup array to quickly determine if a period should be removed
        Dim periodsToRemoveMap() As Boolean
        ReDim periodsToRemoveMap(1 To numPeriods)
        For i = 1 To periodsToRemove
            periodsToRemoveMap(periodIndices(i)) = True
        Next i
        
        ' Copy only non-removed periods to the new array
        K = 1
        For i = 1 To numPeriods
            If Not periodsToRemoveMap(i) Then
                For j = 1 To numStrategies
                    reducedPnL(K, j) = pnlResults(i, j)
                Next j
                K = K + 1
            End If
        Next i
        
        RemoveTopPerformers = reducedPnL
    Else
        RemoveTopPerformers = pnlResults
    End If
    
End Function



Function CleanPortfolioDailyPnL( _
    startdate As Date, _
    endDate As Date _
) As Variant
    Dim ws          As Worksheet
    Dim lastRow     As Long, lastCol As Long
    Dim rawDate     As Variant
    Dim d           As Date, targetDate As Date
    Dim dict        As Object
    Dim rowPnL()    As Double
    Dim i As Long, j As Long
    
    ' for filtering & sorting
    Dim keysArr     As Variant
    Dim idx         As Long, tmpD As Date
    Dim SumPnL      As Double
    
    Dim stratCount  As Long
    
    ' 1) load sheet & size
    Set ws = ThisWorkbook.Sheets("PortfolioDailyM2M")
    With ws
        lastRow = .Cells(.rows.count, 1).End(xlUp).row
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).column
    End With
    stratCount = lastCol - 1
    If stratCount < 1 Then
        MsgBox "No strategy columns found!", vbExclamation
        CleanPortfolioDailyPnL = Empty: Exit Function
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 2) read, shift weekends & aggregate
    For i = 2 To lastRow
        rawDate = ws.Cells(i, 1).value
        If IsDate(rawDate) Then
            d = dateValue(rawDate)
            Select Case Weekday(d, vbMonday)
              Case 6: targetDate = d - 1   ' Sat?Fri
              Case 7: targetDate = d + 1   ' Sun?Mon
              Case Else: targetDate = d
            End Select
            If targetDate >= startdate And targetDate <= endDate Then
                ReDim rowPnL(1 To stratCount)
                For j = 1 To stratCount
                    If IsNumeric(ws.Cells(i, j + 1).value) Then
                        rowPnL(j) = CDbl(ws.Cells(i, j + 1).value)
                    Else
                        rowPnL(j) = 0
                    End If
                Next j
                
                If dict.Exists(targetDate) Then
                    Dim prev() As Double
                    prev = dict(targetDate)
                    For j = 1 To stratCount
                        prev(j) = prev(j) + rowPnL(j)
                    Next j
                    dict(targetDate) = prev
                Else
                    dict.Add targetDate, rowPnL
                End If
            End If
        End If
    Next i
    
    ' 3) drop every date where total PnL = 0
    keysArr = dict.keys
    For idx = LBound(keysArr) To UBound(keysArr)
        d = CDate(keysArr(idx))
        SumPnL = 0
        Dim arr() As Double
        arr = dict(d)
        For j = LBound(arr) To UBound(arr)
            SumPnL = SumPnL + arr(j)
        Next j
        If SumPnL = 0 Then dict.Remove d
    Next idx
    
    ' 4) sort remaining dates
    keysArr = dict.keys
    For i = LBound(keysArr) To UBound(keysArr) - 1
        For idx = i + 1 To UBound(keysArr)
            If CDate(keysArr(i)) > CDate(keysArr(idx)) Then
                tmpD = keysArr(i)
                keysArr(i) = keysArr(idx)
                keysArr(idx) = tmpD
            End If
        Next idx
    Next i
    
    ' 5) build result array: rows = non-zero trading days
    Dim outRows As Long
    outRows = UBound(keysArr) - LBound(keysArr) + 1
    Dim result() As Variant
    ReDim result(1 To outRows, 0 To stratCount)
    
    For idx = LBound(keysArr) To UBound(keysArr)
        i = idx - LBound(keysArr) + 1
        result(i, 0) = keysArr(idx)
        arr = dict(keysArr(idx))
        For j = 1 To stratCount
            result(i, j) = arr(j)
        Next j
    Next idx
    
    CleanPortfolioDailyPnL = result
End Function



Function ConvertDailyToWeeklyPnL(startdate As Date, endDate As Date) As Variant
    Dim dailyArr    As Variant
    Dim weekDict    As Object
    Dim totalDays   As Long, stratCount As Long
    Dim i As Long, j As Long
    Dim d As Date, wk As Date
    Dim tmp()       As Double
    Dim keys        As Variant
    Dim nWeeks      As Long
    Dim resultArr   As Variant

    '—— 1) Grab the cleaned daily PnL (dates in col 0, PnLs in 1…N) ——
    dailyArr = CleanPortfolioDailyPnL(startdate, endDate)
    If IsEmpty(dailyArr) Then
        ConvertDailyToWeeklyPnL = Empty
        Exit Function
    End If

    totalDays = UBound(dailyArr, 1)
    stratCount = UBound(dailyArr, 2)

    Set weekDict = CreateObject("Scripting.Dictionary")

    '—— 2) Bucket each calendar date into its Monday ——
    For i = 1 To totalDays
        d = dailyArr(i, 0)
        ' back up to Monday of that week
        wk = d - (Weekday(d, vbMonday) - 1)

        If Not weekDict.Exists(wk) Then
            ReDim tmp(1 To stratCount)
            weekDict.Add wk, tmp
        End If

        tmp = weekDict(wk)
        For j = 1 To stratCount
            tmp(j) = tmp(j) + dailyArr(i, j)
        Next j
        weekDict(wk) = tmp
    Next i

    '—— 3) Sort the week-start dates and build the output array ——
    keys = weekDict.keys
    Call QuickSortStrings(keys, LBound(keys), UBound(keys))  ' assumes you have your string-sort routine

    nWeeks = UBound(keys) - LBound(keys) + 1
    ReDim resultArr(1 To nWeeks, 0 To stratCount)

    For i = 1 To nWeeks
        resultArr(i, 0) = CDate(keys(i - 1 + LBound(keys)))
        tmp = weekDict(resultArr(i, 0))
        For j = 1 To stratCount
            resultArr(i, j) = tmp(j)
        Next j
    Next i

    ConvertDailyToWeeklyPnL = resultArr
End Function


' Helper function to sort dates
Private Sub SortDates(ByRef dateArray() As Date)
    Dim i As Long, j As Long
    Dim tempDate As Date
    
    For i = LBound(dateArray) To UBound(dateArray) - 1
        For j = i + 1 To UBound(dateArray)
            If dateArray(i) > dateArray(j) Then
                tempDate = dateArray(i)
                dateArray(i) = dateArray(j)
                dateArray(j) = tempDate
            End If
        Next j
    Next i
End Sub



'— In-place QuickSort for string arrays (lexical = chronological on "yyyy-ww") —
Private Sub QuickSortStrings(arr As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim pivot As String, tmp As String
    Dim i As Long, j As Long
    pivot = arr((lo + hi) \ 2)
    i = lo: j = hi
    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortStrings arr, lo, j
    If i < hi Then QuickSortStrings arr, i, hi
End Sub

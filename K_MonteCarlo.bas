Attribute VB_Name = "K_MonteCarlo"

Sub RunMonteCarloSimulation(i As Long)
    Dim wsSummary As Worksheet
    Dim strategyName As String
    Dim startingMargin As Double
    Dim startingEquity As Double
    Dim pnlDays As Variant
    Dim requiredMargin As Double
    Dim averageTradesPerYear As Long
    Dim numScenarios As Long
    Dim results As Variant
    Dim targetRiskOfRuin As Double
    Dim tolerance As Double
    Dim currentRiskOfRuin As Double
    Dim medianReturn As Double
    Dim ruinedCount As Long
    Dim count As Long
    Dim tradeOption As String
    Dim startdate As Date
    Dim endDate As Date
    Dim tradeAdjustment As Double
    
    Dim backtestPeriod As String
    Dim AverageTrade As Double


' Set the Summary worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    ' Get required values from the summary sheet
    strategyName = wsSummary.Cells(i, COL_STRATEGY_NAME).value
    requiredMargin = wsSummary.Cells(i, COL_MARGIN).value
    startingEquity = wsSummary.Cells(i, COL_MARGIN).value * GetNamedRangeValue("MC_StartingEquity")
    backtestPeriod = GetNamedRangeValue("MC_Period")
    numScenarios = GetNamedRangeValue("MC_Simulations")
    tradeOption = GetNamedRangeValue("MC_Trade_Option")
    tradeAdjustment = GetNamedRangeValue("MC_TradeAdjustment")
    targetRiskOfRuin = GetNamedRangeValue("MC_RiskRuinTarget") ' 10%
    tolerance = GetNamedRangeValue("MC_RiskRuinThreshold") ' 2% tolerance
    
    startdate = wsSummary.Cells(i, COL_START_DATE).value
    
    If backtestPeriod = "IS" Then
        endDate = wsSummary.Cells(i, COL_OOS_BEGIN_DATE).value
    Else
        endDate = wsSummary.Cells(i, COL_LAST_DATE_ON_FILE).value
    End If
        
    
    If tradeOption = "Closed Trade" Then
        ' Capture non-zero PNL days
        averageTradesPerYear = wsSummary.Cells(i, COL_TRADES_PER_YEAR).value
        pnlDays = GetNonZeroPNLDays(strategyName, startdate, endDate)
        If UBound(pnlDays) = -1 Then
            MsgBox "No non-zero Closed Trade PNL days found for strategy: " & strategyName, vbExclamation
            Exit Sub
        End If
    ElseIf tradeOption = "Marked 2 Market" Then
        averageTradesPerYear = 300
        pnlDays = GetNonZeroDailyPNLDays(strategyName, startdate, endDate)
        If UBound(pnlDays) = -1 Then
            MsgBox "No non-zero Daily PNL days found for strategy: " & strategyName, vbExclamation
            Exit Sub
        End If
    Else
        
        MsgBox "Error with Monte Carlo Trade Option Input", vbExclamation
        Exit Sub
     
    End If
    
   
    
    AverageTrade = 0
    For j = 1 To UBound(pnlDays)
            AverageTrade = AverageTrade + pnlDays(j)
    Next j
    
    If UBound(pnlDays) > 0 Then AverageTrade = AverageTrade / UBound(pnlDays)
    
    ' Set the desired risk of ruin and tolerance
   
    
    Do
        ' Run Monte Carlo simulation
    
        results = RunMonteCarlo(pnlDays, requiredMargin, averageTradesPerYear, startingEquity, numScenarios, tradeAdjustment, AverageTrade)
        
        Dim rowCount As Long
        rowCount = UBound(results, 1)
        
        If rowCount < 1 Then
            MsgBox "No valid scenarios to evaluate.", vbExclamation
            Exit Sub
        End If
        
        
        
        ruinedCount = 0
        prob0count = 0
        For K = 1 To rowCount
            If results(K, 6) = 1 Then
                ruinedCount = ruinedCount + 1
            End If
        Next K
                
                       
        ' Calculate the current risk of ruin as a percentage
        currentRiskOfRuin = ruinedCount / numScenarios
        
        ' Calculate the median of ReturnToDrawdown values
        medianReturn = WorksheetFunction.Median(Application.index(results, 0, 4))
        
        
        ' Check if current risk of ruin is within tolerance
        If currentRiskOfRuin > targetRiskOfRuin + tolerance Then
            '
            startingEquity = startingEquity * 1.05
        ElseIf currentRiskOfRuin < targetRiskOfRuin - tolerance Then
            startingEquity = startingEquity * 0.991
        End If
        count = count + 1

    Loop While Abs(currentRiskOfRuin - targetRiskOfRuin) > tolerance Or count < 100
    
    ' Save the results to the summary tab
    wsSummary.Cells(i, COL_BACKTEST_MC).value = medianReturn
    wsSummary.Cells(i, COL_NOTIONAL_CAPITAL).value = startingEquity
    'wsSummary.Cells(i, COL_RISK_RUIN).value = currentRiskOfRuin
    wsSummary.Cells(i, COL_EXPECTED_ANNUAL_RETURN).value = wsSummary.Cells(i, COL_EXPECTED_ANNUAL_PROFIT).value / (startingEquity + 0.001)
    wsSummary.Cells(i, COL_ACTUAL_ANNUAL_RETURN).value = wsSummary.Cells(i, COL_ACTUAL_ANNUAL_PROFIT).value / (startingEquity + 0.001)
    
    
End Sub

Function GetNonZeroPNLDays(strategyName As String, startdate As Date, endDate As Date) As Variant
    Dim wsClosedTradePNL As Worksheet
    Dim lastRow As Long
    Dim pnlDays As Collection
    Dim i As Long
    Dim strategyColumn As Long
    Dim dateColumn As Long
    
    ' Set the ClosedTradePNL worksheet
    Set wsClosedTradePNL = ThisWorkbook.Sheets("ClosedTradePNL")
    lastRow = EndRowByCutoffSimple(wsClosedTradePNL, 1)


    Set pnlDays = New Collection


    ' Find the  name column
    For i = 1 To wsClosedTradePNL.Cells(1, wsClosedTradePNL.Columns.count).End(xlToLeft).column
        If wsClosedTradePNL.Cells(1, i).value = "Date" Then
            dateColumn = i
        ElseIf wsClosedTradePNL.Cells(1, i).value = strategyName Then
            strategyColumn = i
        End If
        Next i

    ' Check if the strategy columstrategyn was found
    If strategyColumn = 0 Then
        MsgBox strategyName & " not found in the ClosedTradePNL sheet.", vbExclamation
       
        GetNonZeroPNLDays = Array()
        Exit Function
    End If

    ' Loop through each row and add non-zero PNL values to pnlDays
    For i = 2 To lastRow
        If wsClosedTradePNL.Cells(i, dateColumn).value >= startdate And wsClosedTradePNL.Cells(i, dateColumn).value <= endDate Then
            If wsClosedTradePNL.Cells(i, strategyColumn).value <> 0 Then
                pnlDays.Add wsClosedTradePNL.Cells(i, strategyColumn).value
            End If
        End If
    Next i
    ' Convert collection to array
    Dim pnlDaysArray() As Variant
    If pnlDays.count > 0 Then
        ReDim pnlDaysArray(1 To pnlDays.count)
        For i = 1 To pnlDays.count
            pnlDaysArray(i) = pnlDays(i)
        Next i
        GetNonZeroPNLDays = pnlDaysArray
    Else

        GetNonZeroPNLDays = Array()
    End If
End Function

Sub RunPortfolioMonteCarloSimulation()
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
    Set wsPortfolioMC = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
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
    
    ' Dates and trading information
    yearsToConsider = GetNamedRangeValue("PortfolioPeriod")
    currentdate = wsPortfolio.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    startdate = DateAdd("yyyy", -Int(yearsToConsider), currentdate)
    startdate = DateAdd("m", -(yearsToConsider - Int(yearsToConsider)) * 12, startdate)
    endDate = currentdate
    
    
    If MCTradeType = daily Then
        averageTradesPerYear = 282 'based on average
    Else
        averageTradesPerYear = 52
    End If
    
    ' Get the PnL Days
    
    If MCTradeType = daily Then
        pnlResults = GetNonZeroDailyPortfolioPNLDays(startdate, endDate)
    Else
        Dim lastRow As Long
    
    ' Check if the column has data, and safely find the last row
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
    
    
    If UBound(pnlResults) = -1 Then
        MsgBox "No results found", vbExclamation
        Exit Sub
    End If


    Dim AverageTrade As Double
    AverageTrade = 0
    For j = 1 To UBound(pnlResults)
            AverageTrade = AverageTrade + pnlResults(j)
    Next j
    
    If UBound(pnlResults) > 0 Then AverageTrade = AverageTrade / UBound(pnlResults)

    
    count = 0 ' Initialize the iteration count

    ' Do While loop to adjust starting equity based on risk of ruin
    Do
        ' Run Monte Carlo simulation and gather results
        If ceaseTradingType = "Percentage" Then
            margin = (1 - requiredMargin) * startingEquity
        Else
            margin = requiredMargin
        End If
        
        results = RunMonteCarlo(pnlResults, margin, averageTradesPerYear, startingEquity, numScenarios, tradeAdjustment, AverageTrade)
        
        
           Dim rowCount As Long
        rowCount = UBound(results, 1)
        
        If rowCount < 1 Then
            MsgBox "No valid scenarios to evaluate.", vbExclamation
            Exit Sub
        End If
        
        
        ruinedCount = 0
        prob0count = 0
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
                '
                startingEquity = startingEquity * 1.05
            ElseIf currentRiskOfRuin < targetRiskOfRuin - tolerance Then
                startingEquity = startingEquity * 0.991
            End If
        End If
        count = count + 1
        
             
  
        Application.StatusBar = "Monte Carlo Running: " & count & " runs completed"
        
    Loop While Abs(currentRiskOfRuin - targetRiskOfRuin) > tolerance And count < 500 And solveRisk = "Yes"

    ' Calculate summary statistics after exiting the loop
    Dim avgProfit As Double, avgReturn As Double, avgMaxDrawdown As Double, avgReturnToDrawdown As Double
    Dim medianReturn As Double
    
    ' Calculate the median of ReturnToDrawdown values
    medianReturn = WorksheetFunction.Median(Application.index(results, 0, 3))
    medianDrawdown = WorksheetFunction.Median(Application.index(results, 0, 5))
    medianProfit = WorksheetFunction.Median(Application.index(results, 0, 2))
    medianReturnToDrawdown = WorksheetFunction.Median(Application.index(results, 0, 4))
    
    avgReturn = WorksheetFunction.Average(Application.index(results, 0, 3))
    avgMaxDrawdown = WorksheetFunction.Average(Application.index(results, 0, 5))
    avgProfit = WorksheetFunction.Average(Application.index(results, 0, 2))
    avgReturnToDrawdown = WorksheetFunction.Average(Application.index(results, 0, 4))
    
    ' Output summary metrics
    summaryRow = 3
    With wsPortfolioMC
        ' Title Formatting
        .Cells(1, 1).value = "Monte Carlo Simulation Summary"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Interior.Color = RGB(0, 102, 204) ' Light blue background for title
        .Cells(1, 1).Font.Color = RGB(255, 255, 255) ' White font for title
    
        ' Header row background
        .Range(.Cells(summaryRow, 1), .Cells(summaryRow + 13, 1)).Interior.Color = RGB(224, 224, 224) ' Light grey background for headers
        .Range(.Cells(summaryRow, 1), .Cells(summaryRow + 13, 1)).Font.Bold = True ' Bold header labels
    
    
    
       requiredMargin = GetNamedRangeValue("PortfolioCeaseTrading")
   
    numScenarios = GetNamedRangeValue("PortfolioSimulations")
    tradeAdjustment = GetNamedRangeValue("PortfolioMCTradeAdjustment")
    targetRiskOfRuin = GetNamedRangeValue("PortfolioRiskRuinTarget")
    tolerance = GetNamedRangeValue("PortfolioRiskRuinTolerance")
    
    ' Dates and trading information
    yearsToConsider = GetNamedRangeValue("PortfolioPeriod")
    
    
        ' Populate the table with values
        .Cells(summaryRow, 1).value = "Starting Capital"
        .Cells(summaryRow, 2).value = startingEquity
        .Cells(summaryRow, 2).NumberFormat = "$#,##0"
        .Cells(summaryRow + 1, 1).value = "Minimum Portfolio Value"
        .Cells(summaryRow + 1, 2).value = margin
        .Cells(summaryRow + 1, 2).NumberFormat = "$#,##0"
        .Cells(summaryRow + 2, 1).value = "Backtest Period"
        .Cells(summaryRow + 2, 2).value = yearsToConsider & " years"
        .Cells(summaryRow + 3, 1).value = "Average Profit ($)"
        .Cells(summaryRow + 3, 2).value = avgProfit
        .Cells(summaryRow + 3, 2).NumberFormat = "$#,##0"
        .Cells(summaryRow + 4, 1).value = "Median Profit ($)"
        .Cells(summaryRow + 4, 2).value = medianProfit
        .Cells(summaryRow + 4, 2).NumberFormat = "$#,##0"
        
        .Cells(summaryRow + 5, 1).value = "Average Return (%)"
        .Cells(summaryRow + 5, 2).value = avgReturn
        .Cells(summaryRow + 5, 2).NumberFormat = "#%"
        .Cells(summaryRow + 6, 1).value = "Median Return (%)"
        .Cells(summaryRow + 6, 2).value = medianReturn
        .Cells(summaryRow + 6, 2).NumberFormat = "#%"
        
        .Cells(summaryRow + 7, 1).value = "Average Max Drawdown ($)"
        .Cells(summaryRow + 7, 2).value = avgMaxDrawdown * startingEquity
        .Cells(summaryRow + 7, 2).NumberFormat = "$#,##0"
        .Cells(summaryRow + 8, 1).value = "Median Max Drawdown($)"
        .Cells(summaryRow + 8, 2).value = medianDrawdown * startingEquity
        .Cells(summaryRow + 8, 2).NumberFormat = "$#,##0"
        
        
        .Cells(summaryRow + 9, 1).value = "Average Max Drawdown"
        .Cells(summaryRow + 9, 2).value = avgMaxDrawdown
        .Cells(summaryRow + 9, 2).NumberFormat = "#%"
        .Cells(summaryRow + 10, 1).value = "Median Max Drawdown"
        .Cells(summaryRow + 10, 2).value = medianDrawdown
        .Cells(summaryRow + 10, 2).NumberFormat = "#%"
        
        .Cells(summaryRow + 11, 1).value = "Average Return to Drawdown"
        .Cells(summaryRow + 11, 2).value = avgReturnToDrawdown
        .Cells(summaryRow + 11, 2).NumberFormat = "0.0"
        .Cells(summaryRow + 12, 1).value = "Median Return to Drawdown"
        .Cells(summaryRow + 12, 2).value = medianReturnToDrawdown
        .Cells(summaryRow + 12, 2).NumberFormat = "0.0"
        
        .Cells(summaryRow + 13, 1).value = "Risk of Ruin"
        .Cells(summaryRow + 13, 2).value = currentRiskOfRuin
        .Cells(summaryRow + 13, 2).NumberFormat = "0.0%"
        
        ' Formatting for value cells
         .Range(.Cells(summaryRow, 2), .Cells(summaryRow + 13, 2)).Interior.Color = RGB(242, 242, 242) ' Light grey for values
    
        ' Apply borders around the entire table
        With .Range(.Cells(summaryRow, 1), .Cells(summaryRow + 13, 2)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0) ' Black border
        End With
    
        ' Autofit columns to the content
        .Columns("A:Z").AutoFit
    End With

  maxProfit = WorksheetFunction.Max(Application.index(results, 0, 2))
 
 ' Calculate bin width as one-tenth of maxprofit, rounded to the nearest 1,000
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
        drawdownArray(i) = results(i, 5)
        profitArray(i) = results(i, 2)
    Next i

    

   ' Generate histograms for Profit, Return, and Max Drawdown with custom rounding in separate columns starting from column 15 (O)
    CreateHistogram wsPortfolioMC, returnArray, "Return Histogram", 10, 14, summaryRow, 2, 0.05
    CreateHistogram wsPortfolioMC, drawdownArray, "Max Drawdown Histogram", 10, 17, summaryRow, 16, 0.05
    CreateHistogram wsPortfolioMC, profitArray, "Profit Histogram", 10, 20, summaryRow, 30, binWidth

    
        ' Create delete button
    Dim btn As Object
    Set btn = wsPortfolioMC.Buttons.Add(left:=wsPortfolioMC.Cells(1, 1).left + 30, top:=wsPortfolioMC.Cells(18, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeletePortfolioMC" ' Make sure to create this sub to handle deletion
    End With

 ' Create a button to return to the Summary page
    Set btn = wsPortfolioMC.Buttons.Add(left:=wsPortfolioMC.Cells(1, 1).left + 30, _
                                    top:=wsPortfolioMC.Cells(21, 1).top, _
                                    Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary" ' Assign the macro to run when the button is clicked
    End With
 
 ' Create a button to return to the Portfolio page
    Set btn = wsPortfolioMC.Buttons.Add(left:=wsPortfolioMC.Cells(1, 1).left + 30, _
                                    top:=wsPortfolioMC.Cells(24, 1).top, _
                                    Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio" ' Assign the macro to run when the button is clicked
    End With
 ' Create a button to return to the Control page
    Set btn = wsPortfolioMC.Buttons.Add(left:=wsPortfolioMC.Cells(1, 1).left + 30, _
                                    top:=wsPortfolioMC.Cells(27, 1).top, _
                                    Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl" ' Assign the macro to run when the button is clicked
    End With
    
    Set btn = wsPortfolioMC.Buttons.Add(left:=wsPortfolioMC.Cells(1, 1).left + 30, _
                                    top:=wsPortfolioMC.Cells(30, 1).top, _
                                    Width:=100, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies" ' Assign the macro to run when the button is clicked
    End With
    
    Set btn = wsPortfolioMC.Buttons.Add(left:=wsPortfolioMC.Cells(1, 1).left + 30, _
                                    top:=wsPortfolioMC.Cells(33, 1).top, _
                                    Width:=100, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs" ' Assign the macro to run when the button is clicked
    End With
    
    
    With ThisWorkbook.Windows(1)
        .Zoom = 70 ' Set zoom level to 70%
    End With


 ' Autofit columns in the PortfolioGraphs sheet for readability
    wsPortfolioMC.Columns("A:Z").AutoFit
    
    Call OrderVisibleTabsBasedOnList
    
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    
    wsPortfolioMC.Activate
        
End Sub





Function GetNonZeroDailyPNLDays(strategyName As String, startdate As Date, endDate As Date) As Variant
    Dim wsDailyM2MEquity As Worksheet
    Dim lastRow As Long
    Dim pnlDays As Collection
    Dim i As Long
    Dim strategyColumn As Long
    Dim dateColumn As Long
    Dim nonZeroCount As Long
    Dim expectedTradingDays As Long
    Dim daysWithZeros As Long
    
    ' Set the DailyM2MEquity worksheet
    Set wsDailyM2MEquity = ThisWorkbook.Sheets("DailyM2MEquity")
    lastRow = EndRowByCutoffSimple(wsDailyM2MEquity, 1)

    Set pnlDays = New Collection

    ' Find the strategy name column and the date column
    For i = 1 To wsDailyM2MEquity.Cells(1, wsDailyM2MEquity.Columns.count).End(xlToLeft).column
        If wsDailyM2MEquity.Cells(1, i).value = "Date" Then
            dateColumn = i
        ElseIf wsDailyM2MEquity.Cells(1, i).value = strategyName Then
            strategyColumn = i
        End If
    Next i

    ' Check if both columns were found
    If strategyColumn = 0 Or dateColumn = 0 Then
        MsgBox strategyName & " or Date column not found in the DailyM2MEquity sheet.", vbExclamation
        'ReDim pnlDaysArray(1 To 0) ' Return an empty array
        GetNonZeroDailyPNLDays = Array()
        Exit Function
    End If

    ' Loop through each row within the date range and add non-zero PNL values to pnlDays
    For i = 2 To lastRow
        If wsDailyM2MEquity.Cells(i, dateColumn).value >= startdate And wsDailyM2MEquity.Cells(i, dateColumn).value <= endDate Then
            If wsDailyM2MEquity.Cells(i, strategyColumn).value <> 0 Then
                pnlDays.Add wsDailyM2MEquity.Cells(i, strategyColumn).value
            End If
        End If
    Next i

    ' Calculate the expected trading days using the 252 trading days per year assumption
    yearsDifference = (endDate - startdate) / 365.25
    expectedTradingDays = Application.WorksheetFunction.RoundDown(yearsDifference * 252, 0)
    
    ' Calculate how many zero values to add to match expected trading days
    nonZeroCount = pnlDays.count
    daysWithZeros = expectedTradingDays - nonZeroCount

    ' Add zeros to the pnlDays collection to match the expected trading days
    For i = 1 To daysWithZeros
        pnlDays.Add 0
    Next i

    ' Convert collection to array
    Dim pnlDaysArray() As Variant
    If pnlDays.count > 0 Then
        ReDim pnlDaysArray(1 To pnlDays.count)
        For i = 1 To pnlDays.count
            pnlDaysArray(i) = pnlDays(i)
        Next i
        GetNonZeroDailyPNLDays = pnlDaysArray
    Else
        GetNonZeroDailyPNLDays = Array()
    End If
End Function


Function GetNonZeroDailyPortfolioPNLDays(startdate As Date, endDate As Date) As Variant
    Dim wsTotalPortfolioM2M As Worksheet
    Dim lastRow As Long
    Dim pnlDays As Collection
    Dim i As Long
    Dim dateColumn As Long
    Dim nonZeroCount As Long
    Dim expectedTradingDays As Long
    Dim daysWithZeros As Long
    Dim strategyColumn As Long
    
    ' Set the DailyM2MEquity worksheet
    Set wsTotalPortfolioM2M = ThisWorkbook.Sheets("TotalPortfolioM2M")
    lastRow = wsTotalPortfolioM2M.Cells(wsTotalPortfolioM2M.rows.count, 1).End(xlUp).row

    Set pnlDays = New Collection

    ' Find the strategy name column and the date column
    For i = 1 To wsTotalPortfolioM2M.Cells(1, wsTotalPortfolioM2M.Columns.count).End(xlToLeft).column
        If wsTotalPortfolioM2M.Cells(1, i).value = "Date" Then
            dateColumn = i
        ElseIf wsTotalPortfolioM2M.Cells(1, i).value = "Total Daily Profit" Then
            strategyColumn = i
        End If
    Next i

    ' Check if both columns were found
    If strategyColumn = 0 Or dateColumn = 0 Then
        MsgBox "Daily profits or Date column not found in the TotalPortfolioM2M sheet.", vbExclamation
        GetNonZeroDailyPortfolioPNLDays = Nothing
        Exit Function
    End If

    ' Loop through each row within the date range and add non-zero PNL values to pnlDays
    For i = 2 To lastRow
        If wsTotalPortfolioM2M.Cells(i, dateColumn).value >= startdate And wsTotalPortfolioM2M.Cells(i, dateColumn).value <= endDate Then
            If wsTotalPortfolioM2M.Cells(i, strategyColumn).value <> 0 Then
                pnlDays.Add wsTotalPortfolioM2M.Cells(i, strategyColumn).value
            End If
        End If
    Next i

    ' Calculate the expected trading days using the 252 trading days per year assumption
    yearsDifference = (endDate - startdate) / 365.25
    expectedTradingDays = Application.WorksheetFunction.RoundDown(yearsDifference * 252, 0)
    
    ' Calculate how many zero values to add to match expected trading days
    nonZeroCount = pnlDays.count
    daysWithZeros = expectedTradingDays - nonZeroCount

    ' Add zeros to the pnlDays collection to match the expected trading days
    For i = 1 To daysWithZeros
        pnlDays.Add 0
    Next i

    ' Convert collection to array
    Dim pnlDaysArray() As Variant
    If pnlDays.count > 0 Then
        ReDim pnlDaysArray(1 To pnlDays.count)
        For i = 1 To pnlDays.count
            pnlDaysArray(i) = pnlDays(i)
        Next i
        GetNonZeroDailyPortfolioPNLDays = pnlDaysArray
    Else
        GetNonZeroDailyPortfolioPNLDays = Nothing
    End If
End Function



Function RunMonteCarlo(pnlResults As Variant, marginThreshold As Double, averageTradesPerYear As Long, startingMargin As Double, numScenarios As Long, tradeAdjustment As Double, AverageTrade As Double) As Variant
    Dim totalTrades As Long, randomIndices() As Long
    Dim equity() As Double, profit() As Double, returnToDrawdown() As Double
    Dim maxDrawdown() As Double, ruinedScenarios() As Long
    Dim i As Long, j As Long, tradeIndex As Long
    Dim peakEquity As Double, drawdown As Double, adjustedTradeFactor As Double

    ' Precompute reusable values
    totalTrades = UBound(pnlResults)
    adjustedTradeFactor = AverageTrade * (1 - tradeAdjustment)

    ' Pre-allocate arrays
    ReDim equity(1 To numScenarios)
    ReDim profit(1 To numScenarios)
    ReDim returnToDrawdown(1 To numScenarios)
    ReDim maxDrawdown(1 To numScenarios)
    ReDim ruinedScenarios(1 To numScenarios)
    
    
    
    ReDim randomIndices(1 To (numScenarios * averageTradesPerYear))

    ' Generate random indices in bulk
    For i = 1 To UBound(randomIndices)
        randomIndices(i) = Int(totalTrades * Rnd + 1)
    Next i

    Dim randomIndexPointer As Long
    randomIndexPointer = 1

    ' Run scenarios
    For i = 1 To numScenarios
        Dim currentEquity As Double
        currentEquity = startingMargin
        peakEquity = startingMargin
        maxDrawdown(i) = 0
        ruinedScenarios(i) = 0

        ' Simulate trades
        For j = 1 To averageTradesPerYear
            tradeIndex = randomIndices(randomIndexPointer)
            randomIndexPointer = randomIndexPointer + 1

            Dim tradeReturn As Double
            tradeReturn = pnlResults(tradeIndex) - adjustedTradeFactor

            ' Update equity
            currentEquity = currentEquity + tradeReturn

            ' Update peak equity and drawdown
            If currentEquity > peakEquity Then
                peakEquity = currentEquity
            Else
                drawdown = 1 - IIf(peakEquity = 0, 0, (currentEquity / peakEquity + 0.000001))
                If drawdown > maxDrawdown(i) Then maxDrawdown(i) = drawdown
            End If

            ' Check for margin threshold breach
            If currentEquity < marginThreshold Then
                ruinedScenarios(i) = 1
                Exit For
            End If
        Next j

        ' Store results for this scenario
        equity(i) = currentEquity
        profit(i) = currentEquity - startingMargin
        returnToDrawdown(i) = IIf(maxDrawdown(i) = 0, 4, ((currentEquity / startingMargin) - 1) / (maxDrawdown(i) + 0.00001))
    Next i

    ' Return results as a 2D array
    Dim results() As Variant
    ReDim results(1 To numScenarios, 1 To 6)
    For i = 1 To numScenarios
        results(i, 1) = equity(i)
        results(i, 2) = profit(i)
        results(i, 3) = (equity(i) / startingMargin - 1)
        results(i, 4) = returnToDrawdown(i)
        results(i, 5) = maxDrawdown(i)
        results(i, 6) = ruinedScenarios(i)
    Next i

    RunMonteCarlo = results
    
End Function



Sub RunAllMonteCarloSimulations()
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the Summary worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    ' Get the last row in the Summary sheet
    lastRow = wsSummary.Cells(wsSummary.rows.count, COL_STRATEGY_NAME).End(xlUp).row
    
    ' Loop through each strategy in the summary sheet
    For i = 2 To lastRow ' Assuming the first row is headers
        ' Call the Monte Carlo simulation for each strategy
        RunMonteCarloSimulation i
    Next i

    MsgBox "Monte Carlo simulations completed for all strategies.", vbInformation
End Sub



Sub RunAllMC()
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim simulationsRun As Long

    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        Exit Sub
    End If

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    
    ' Set the summary worksheet
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsSummary Is Nothing Then
        MsgBox "Error: 'Summary' sheet does not exist.", vbExclamation
        Exit Sub
    End If

    ' Exit and show error if the sheet exists but has no data in row 2
    If wsSummary.Cells(2, COL_STRATEGY_NAME).value = "" Then
        MsgBox "Error: 'Summary' sheet exists but contains no data in row 2.", vbExclamation
        Exit Sub
    End If

    ' Find the last row in the Summary sheet
    lastRow = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row

    ' Initialize counter for simulations run
    simulationsRun = 0

    ' Loop through rows to run Monte Carlo simulations based on status
    For i = 2 To lastRow
        
            Call RunMonteCarloSimulation(i)
            simulationsRun = simulationsRun + 1
        
        Application.StatusBar = "Running Monte Carlo on all strategies: " & Format(i / lastRow, "0%") & " completed"
        
    Next i
    
    
    Application.StatusBar = False
       
    ' Display a message based on the number of simulations run
    If simulationsRun > 0 Then
        MsgBox simulationsRun & " Monte Carlo simulations completed successfully.", vbInformation
    Else
        MsgBox "No Monte Carlo simulations were run. No entries found with status '" & status & "'.", vbExclamation
    End If
End Sub




Sub RunMultipleMC(status As String, prompt As Boolean)
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim simulationsRun As Long


    
    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    
    ' Set the summary worksheet
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsSummary Is Nothing Then
        MsgBox "Error: 'Summary' sheet does not exist.", vbExclamation
        Exit Sub
    End If

    ' Exit and show error if the sheet exists but has no data in row 2
    If wsSummary.Cells(2, COL_STRATEGY_NAME).value = "" Then
        MsgBox "Error: 'Summary' sheet exists but contains no data in row 2.", vbExclamation
        Exit Sub
    End If

    ' Find the last row in the Summary sheet
    lastRow = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row

    ' Initialize counter for simulations run
    simulationsRun = 0

    ' Loop through rows to run Monte Carlo simulations based on status
    For i = 2 To lastRow
        If status = wsSummary.Cells(i, COL_STATUS).value Then
            Call RunMonteCarloSimulation(i)
            simulationsRun = simulationsRun + 1
        End If
        
        Application.StatusBar = "Running Monte Carlo on all strategies: " & Format(i / lastRow, "0%") & " completed"

    Next i
    
    Application.StatusBar = "False"
    ' Display a message based on the number of simulations run
    
    If prompt Then
        If simulationsRun > 0 Then
            MsgBox simulationsRun & " Monte Carlo simulations completed successfully.", vbInformation
        Else
            MsgBox "No Monte Carlo simulations were run. No entries found with status '" & status & "'.", vbExclamation
        End If
    End If
End Sub



Sub RunMultipleMCLive()
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        Exit Sub
    End If
    Call RunMultipleMC(GetNamedRangeValue("Port_Status"), True)
End Sub

Sub RunMultipleMCPassedandLive()
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        Exit Sub
    End If
    Call RunMultipleMC(GetNamedRangeValue("Port_Status"), False)
    Call RunMultipleMC(GetNamedRangeValue("Pass_Status"), True)
End Sub







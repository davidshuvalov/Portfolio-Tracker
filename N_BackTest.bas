Attribute VB_Name = "N_BackTest"
Option Explicit

Public strategyList As String
Public symbolList As String
Public sectorList As String



Sub GenerateBackTest()
    Dim wsBacktest As Worksheet, wsDailyM2M As Worksheet, wsSummary As Worksheet, wsClosedTrade As Worksheet, wsBackTestGraphs As Worksheet, wsTotalBackTest As Worksheet
    Dim wsBackTestM2M As Worksheet
    Dim strategyName As String, lastRowBacktest As Long, lastRowSummary As Long
    Dim fromDate As Date, toDate As Date
    Dim numContracts As Double, row As Long, col As Long
    Dim summaryStartDate As Date, summaryEndDate As Date
    Dim dailyEquity As Double, totalEquity As Double
    Dim errorsFound As Boolean
    Dim dictStrategies As Object
    Dim earliestDate As Date
    Dim latestDate As Date, currentdate As Date
    Dim totalPnL As Double, winningPnL As Double, losingPnL As Double
    Dim peakProfit As Double, totalCumulativeProfit As Double, totalDailyProfit As Double, currentDrawdown As Double, startingEquity As Double
    Dim drawdownpercent As Double
    Dim totalTrades As Long, winningTrades As Long, losingTrades As Long
    Dim portfolioWinRate As Double, avgProfit As Double, avgLoss As Double
    Dim riskToReward As Double, edge As Double
    Dim summaryRow As Long
    Dim overRideStart As Date, overRideEnd As Date
    
    
    
    
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        Exit Sub
    End If
    
    
    ' Initialize column constants manually
    Call InitializeColumnConstantsManually


    On Error Resume Next
    Set wsDailyM2M = ThisWorkbook.Sheets("DailyM2MEquity")
    Set wsClosedTrade = ThisWorkbook.Sheets("ClosedTradePNL")

    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsDailyM2M Is Nothing Then
        MsgBox "Error: 'DailyM2MEquity' sheet does not exist.", vbExclamation
        Exit Sub
    End If
    
    If wsClosedTrade Is Nothing Then
        MsgBox "Error: 'ClosedTradePNL' sheet does not exist.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsSummary Is Nothing Then
        MsgBox "Error: 'Summary' sheet does not exist.", vbExclamation
        Exit Sub
    End If

    ' Initialize worksheets
   
    On Error Resume Next
     Set wsBacktest = ThisWorkbook.Sheets("Backtest")
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsBacktest Is Nothing Then
        MsgBox "Error: 'Backtest' sheet does not exist.", vbExclamation
        Exit Sub
    End If



    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Get last rows
    lastRowBacktest = wsBacktest.Cells(wsBacktest.rows.count, 1).End(xlUp).row
    lastRowSummary = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row

    If lastRowBacktest = 1 Then
            MsgBox "Error: No strategies found in the Backtest tab.", vbExclamation
            Exit Sub
    End If

    overRideStart = GetNamedRangeValue("BackTest_Start")
    overRideEnd = GetNamedRangeValue("BackTest_End")


    ' Initialize dictionary for overlap check
    Set dictStrategies = CreateObject("Scripting.Dictionary")

    ' Error checking: Validate date ranges in Backtest tab
    
    ' Error checking: Validate strategy names in Backtest exist in Summary
    Dim found As Integer
    found = 0
    
    For row = 2 To lastRowBacktest
        strategyName = wsBacktest.Cells(row, 1).value ' Column A: Strategy Name
        ' Check if the strategy exists in the Summary tab
        For summaryRow = 2 To lastRowSummary
            If CompareStrategyNames(wsSummary.Cells(summaryRow, COL_STRATEGY_NAME).value, strategyName) Then
                found = 1
                Exit For
            End If
        Next summaryRow
        If found = 0 Then
            MsgBox "Error: Strategy '" & strategyName & "' in row " & row & " of Backtest tab is not found in Summary tab.", vbExclamation
            Exit Sub
        End If
        found = 0
    Next row
    
    
    Dim response As Integer
    Dim ignoreWarnings As Boolean
    ignoreWarnings = False
        
    For row = 2 To lastRowBacktest
        ' Check for type mismatches before proceeding
        If Not IsDate(wsBacktest.Cells(row, 3).value) Then
            MsgBox "Error: Invalid date format in row " & row & ", column C (From Date)." & vbNewLine & _
                   "Please correct the date format before continuing.", vbCritical, "Type Mismatch Error"
            Exit Sub
        End If
        
        If Not IsDate(wsBacktest.Cells(row, 4).value) Then
            MsgBox "Error: Invalid date format in row " & row & ", column D (To Date)." & vbNewLine & _
                   "Please correct the date format before continuing.", vbCritical, "Type Mismatch Error"
            Exit Sub
        End If
        
        If Not IsNumeric(wsBacktest.Cells(row, 2).value) Then
            MsgBox "Error: Invalid numeric value in row " & row & ", column B (Contracts)." & vbNewLine & _
                   "Please enter a valid number before continuing.", vbCritical, "Type Mismatch Error"
            Exit Sub
        End If
        
        strategyName = wsBacktest.Cells(row, 1).value ' Assume column A has strategy names
        numContracts = wsBacktest.Cells(row, 2).value ' Assume column B has contracts
        fromDate = wsBacktest.Cells(row, 3).value     ' Assume column C has from dates
        toDate = wsBacktest.Cells(row, 4).value       ' Assume column D has to dates
        
        
        ' Validate date range is within Summary limits
        For summaryRow = 2 To lastRowSummary
            If CompareStrategyNames(wsSummary.Cells(summaryRow, COL_STRATEGY_NAME).value, strategyName) Then
                summaryStartDate = wsSummary.Cells(summaryRow, COL_START_DATE).value ' Start of Modelling
                summaryEndDate = wsSummary.Cells(summaryRow, COL_LAST_DATE_ON_FILE).value   ' Last Date On File
                
                If Not ignoreWarnings And (fromDate < summaryStartDate Or toDate > summaryEndDate + 31) Then
                    response = MsgBox("Warning: Strategy '" & strategyName & "' in row " & row & _
                                    " has date range outside Summary limits." & vbNewLine & vbNewLine & _
                                    "Click 'Abort' to cancel operation" & vbNewLine & _
                                    "Click 'Continue' to continue anyway" & vbNewLine & _
                                    "Click 'Ignore' to continue and suppress future warnings", _
                                    vbAbortRetryIgnore + vbExclamation, _
                                    "Date Range Warning")
                    
                    Select Case response
                        Case vbAbort  ' Cancel
                            Exit Sub
                        Case vbOK  ' OK/Continue
                            ' Do nothing and continue
                        Case vbIgnore ' Ignore all warnings
                            ignoreWarnings = True
                    End Select
                End If
                Exit For
            End If
        Next summaryRow

        ' Check for overlapping date ranges
        If Not dictStrategies.Exists(strategyName) Then
            dictStrategies.Add strategyName, Array(fromDate, toDate)
        Else
            Dim datesArray As Variant
            datesArray = dictStrategies(strategyName)
            If fromDate <= datesArray(1) And toDate >= datesArray(0) Then
                MsgBox "Error: Overlapping date ranges for strategy '" & strategyName & "' in row " & row & ".", vbExclamation
                Exit Sub
            Else
                dictStrategies(strategyName) = Array(Application.Min(fromDate, datesArray(0)), Application.Max(toDate, datesArray(1)))
            End If
        End If
    Next row

    
    
    
    ' Create or clear "TotalBackTest" sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not ThisWorkbook.Sheets("TotalBackTest") Is Nothing Then
        ThisWorkbook.Sheets("TotalBackTest").Delete
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsTotalBackTest = ThisWorkbook.Sheets.Add(After:=wsSummary)
    wsTotalBackTest.name = "TotalBackTest"
    wsTotalBackTest.Tab.Color = RGB(71, 211, 89)
    
     ' Set white background color for the entire worksheet
    wsTotalBackTest.Cells.Interior.Color = RGB(255, 255, 255)
    
    
    ' Create or clear "BackTestM2MEquity" sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not ThisWorkbook.Sheets("BackTestM2MEquity") Is Nothing Then
        ThisWorkbook.Sheets("BackTestM2MEquity").Delete
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsBackTestM2M = ThisWorkbook.Sheets.Add(After:=wsTotalBackTest)
    wsBackTestM2M.name = "BackTestM2MEquity"
    wsBackTestM2M.Tab.Color = RGB(71, 211, 89)
    
    
     ' Set white background color for the entire worksheet
    wsBackTestM2M.Cells.Interior.Color = RGB(255, 255, 255)
    
      
    ' === 1. Create Dictionaries for Strategies and Strategy Details ===
    Dim dictUniqueStrategies As Object, dictStrategyDetails As Object
    Set dictUniqueStrategies = CreateObject("Scripting.Dictionary")
    Set dictStrategyDetails = CreateObject("Scripting.Dictionary")
    

    
    ' === 2. Collect Unique Strategy Names from Backtest Sheet ===

    For row = 2 To lastRowBacktest
        strategyName = wsBacktest.Cells(row, 1).value
        If Not IsEmpty(strategyName) And Not dictUniqueStrategies.Exists(strategyName) Then
            dictUniqueStrategies.Add strategyName, dictUniqueStrategies.count + 2 ' +2: col 1 for date
        End If
    Next row
    
    ' === 3. Collect Strategy Details (Symbol & Sector) from Summary Tab ===
    Dim symbol As String, sector As String
    
    For row = 2 To lastRowSummary
        strategyName = wsSummary.Cells(row, COL_STRATEGY_NAME).value
        If dictUniqueStrategies.Exists(strategyName) Then
            symbol = wsSummary.Cells(row, COL_SYMBOL).value
            sector = wsSummary.Cells(row, COL_SECTOR).value
            dictStrategyDetails(strategyName) = Array(symbol, sector)
        End If
    Next row
    
    ' === 4. Set Up Headers in BackTestM2MEquity ===
    With wsBackTestM2M
        .Cells(3, 1).value = "Date"
        
        Dim strategy As Variant
        For Each strategy In dictUniqueStrategies
            Dim colNum As Long
            colNum = dictUniqueStrategies(strategy)
            .Cells(1, colNum).value = strategy
            If dictStrategyDetails.Exists(strategy) Then
                .Cells(2, colNum).value = dictStrategyDetails(strategy)(0) ' Symbol
                .Cells(3, colNum).value = dictStrategyDetails(strategy)(1) ' Sector
            End If
        Next strategy
    End With
    
    
    
    ' === Collect Unique Symbols and Sectors from Backtest Strategies ===
    Dim dictUniqueSymbols As Object, dictUniqueSectors As Object
  
    
    Set dictUniqueSymbols = CreateObject("Scripting.Dictionary")
    Set dictUniqueSectors = CreateObject("Scripting.Dictionary")

    
    Dim dictSymbolProfits As Object, dictSectorProfits As Object
    Set dictSymbolProfits = CreateObject("Scripting.Dictionary")
    Set dictSectorProfits = CreateObject("Scripting.Dictionary")
    
    Dim backtestSymbol As String, backtestSector As String
    
    For Each strategy In dictUniqueStrategies.keys
        If dictStrategyDetails.Exists(strategy) Then
            backtestSymbol = dictStrategyDetails(strategy)(0)
            backtestSector = dictStrategyDetails(strategy)(1)
            
            If Not IsEmpty(backtestSymbol) And Not dictUniqueSymbols.Exists(backtestSymbol) Then
                dictUniqueSymbols.Add backtestSymbol, 0
            End If
            
            If Not IsEmpty(backtestSector) And Not dictUniqueSectors.Exists(backtestSector) Then
                dictUniqueSectors.Add backtestSector, 0
            End If
        End If
    Next strategy
    
    
    Dim lastRowDailyM2M As Double
        ' Get last rows
    lastRowBacktest = wsBacktest.Cells(wsBacktest.rows.count, 1).End(xlUp).row
    lastRowDailyM2M = wsDailyM2M.Cells(wsDailyM2M.rows.count, 1).End(xlUp).row
    
    ' === 5. Find Earliest and Latest Dates from Backtest ===
    earliestDate = wsBacktest.Cells(2, 3).value ' From Date
    latestDate = wsBacktest.Cells(2, 4).value   ' To Date
    
    For row = 2 To lastRowBacktest
        If wsBacktest.Cells(row, 3).value < earliestDate Then
            earliestDate = wsBacktest.Cells(row, 3).value
        End If
        If wsBacktest.Cells(row, 4).value > latestDate Then
            latestDate = wsBacktest.Cells(row, 4).value
        End If
    Next row
    
    ' === 5a. Override Dates if necessary ===
    
    If overRideStart > earliestDate Then earliestDate = overRideStart
    If overRideEnd < latestDate Then latestDate = overRideEnd
    
    
    ' === 6. Set Up Headers in TotalBackTest Sheet ===
    With wsTotalBackTest
        .Cells(1, 1).value = "Date"
        .Cells(1, 2).value = "Total Daily Profits"
        .Cells(1, 3).value = "Total Cumulative Profits"
        .Cells(1, 4).value = "Total Drawdown"
        .Cells(1, 5).value = "Year"
        .Cells(1, 6).value = "Month"
        .Cells(1, 7).value = "Week"
        .Cells(1, 8).value = "Drawdown Percentage"
    End With
    
    ' === 7. Initialize Key Variables ===
    Dim outputRow As Long, m2mOutputRow As Long

    
    outputRow = 1
    m2mOutputRow = 1
    peakProfit = 0
    totalCumulativeProfit = 0
    
    totalPnL = 0
    winningPnL = 0
    losingPnL = 0
    totalTrades = 0
    winningTrades = 0
    losingTrades = 0
    
    ' === 8. Load Data into Arrays for Fast Lookup ===
    Dim backtestData As Variant, dailyM2MData As Variant, closedTradeData As Variant
    backtestData = wsBacktest.Range("A2:D" & lastRowBacktest).value
    
    Dim lastRowDaily As Long, lastColDaily As Long
    lastRowDaily = wsDailyM2M.Cells(wsDailyM2M.rows.count, 1).End(xlUp).row
    lastColDaily = wsDailyM2M.Cells(1, wsDailyM2M.Columns.count).End(xlToLeft).column

    ' Capture full range including headers and all data
    dailyM2MData = wsDailyM2M.Range(wsDailyM2M.Cells(1, 1), _
                                wsDailyM2M.Cells(lastRowDaily, lastColDaily)).value
    
    Dim lastRowClosed As Long, lastColClosed As Long
    lastRowClosed = wsClosedTrade.Cells(wsClosedTrade.rows.count, 1).End(xlUp).row
    lastColClosed = wsClosedTrade.Cells(1, wsClosedTrade.Columns.count).End(xlToLeft).column
    
    ' Capture full range including headers and all data
    closedTradeData = wsClosedTrade.Range(wsClosedTrade.Cells(1, 1), _
                                          wsClosedTrade.Cells(lastRowClosed, lastColClosed)).value
    

    
    ' === 10. Create Column Lookup from DailyM2M Headers ===
    Dim strategyColumnLookup As Object
    Set strategyColumnLookup = CreateObject("Scripting.Dictionary")
    
    For col = 2 To UBound(dailyM2MData, 2)
        strategyColumnLookup(dailyM2MData(1, col)) = col
    Next col
    
    ' === 11. Prepare Bulk Output Arrays ===
    Dim totalBackTestOutput() As Variant, m2mOutput() As Variant
    ReDim totalBackTestOutput(1 To lastRowDailyM2M, 1 To 8)
    ReDim m2mOutput(1 To lastRowDailyM2M, 1 To (dictUniqueStrategies.count + 2))
    
    Dim dailyRow As Long, strategyColumn As Long
    
    ' === Track Profit by Symbol and Sector (Only from Backtest) ===
    Dim strategySymbol As String, strategySector As String
    
    ' === READY FOR MAIN LOOP ===
    Application.StatusBar = "Initialization complete. Starting backtest..."
    
    For dailyRow = 2 To lastRowDailyM2M
        currentdate = wsDailyM2M.Cells(dailyRow, 1).value
        totalDailyProfit = 0
        
        
        
        If currentdate >= earliestDate And currentdate <= latestDate Then
            ' === Strategy Profits for the Day ===
            Dim strategyProfits As Object
            Set strategyProfits = CreateObject("Scripting.Dictionary")
            
            ' === First Loop: Calculate Profits for Active Strategies ===
            Dim i As Long
            For i = 1 To UBound(backtestData, 1)
                strategyName = backtestData(i, 1)
                numContracts = backtestData(i, 2)
                fromDate = backtestData(i, 3)
                toDate = backtestData(i, 4)
                
                If currentdate >= fromDate And currentdate <= toDate Then
                    If strategyColumnLookup.Exists(strategyName) Then
                        col = strategyColumnLookup(strategyName)
                        Dim dailyProfit As Double
                        dailyProfit = dailyM2MData(dailyRow, col) * numContracts
                        totalDailyProfit = totalDailyProfit + dailyProfit
                        
                         ' Get Symbol and Sector for the Strategy
                        If dictStrategyDetails.Exists(strategyName) Then
                            strategySymbol = dictStrategyDetails(strategyName)(0)
                            strategySector = dictStrategyDetails(strategyName)(1)
                            
                            ' === Track Symbol Profits ===
                            If Not dictSymbolProfits.Exists(strategySymbol) Then
                                dictSymbolProfits(strategySymbol) = 0
                            End If
                            dictSymbolProfits(strategySymbol) = dictSymbolProfits(strategySymbol) + dailyProfit
                            
                            ' === Track Sector Profits ===
                            If Not dictSectorProfits.Exists(strategySector) Then
                                dictSectorProfits(strategySector) = 0
                            End If
                            dictSectorProfits(strategySector) = dictSectorProfits(strategySector) + dailyProfit
                        End If
                        
                        
                        ' Add to strategy profits
                        If Not strategyProfits.Exists(strategyName) Then
                            strategyProfits.Add strategyName, dailyProfit
                        Else
                            strategyProfits(strategyName) = strategyProfits(strategyName) + dailyProfit
                        End If
                        
                        ' Process PnL data
                        Dim pnlValue As Double
                        pnlValue = closedTradeData(dailyRow, col) * numContracts
                        If pnlValue <> 0 Then
                            totalTrades = totalTrades + 1
                            totalPnL = totalPnL + pnlValue
                            
                            If pnlValue > 0 Then
                                winningTrades = winningTrades + 1
                                winningPnL = winningPnL + pnlValue
                            Else
                                losingTrades = losingTrades + 1
                                losingPnL = losingPnL + pnlValue
                            End If
                        End If
                    End If
                End If
            Next i
            
            ' === Second Loop: Write Results for All Strategies ===
            
            
            m2mOutput(m2mOutputRow, 1) = currentdate
            
            For Each strategy In dictUniqueStrategies
                strategyColumn = dictUniqueStrategies(strategy)
                Dim IsActive As Boolean: IsActive = False
            
                
                
                For i = 1 To UBound(backtestData, 1)
                    If backtestData(i, 1) = strategy Then
                        If currentdate >= backtestData(i, 3) And _
                           currentdate <= backtestData(i, 4) Then
                            IsActive = True
                            Exit For
                        End If
                    End If
                Next i
                
                
                ' Store output in M2M array
                If IsActive Then
                    If strategyProfits.Exists(strategy) Then
                        m2mOutput(m2mOutputRow, strategyColumn) = strategyProfits(strategy)
                    Else
                        m2mOutput(m2mOutputRow, strategyColumn) = 0
                    End If
                Else
                    m2mOutput(m2mOutputRow, strategyColumn) = CVErr(xlErrNA)
                End If
            
            Next strategy
            
            ' === Update Cumulative Profit and Drawdown ===
            totalCumulativeProfit = totalCumulativeProfit + totalDailyProfit
            If totalCumulativeProfit > peakProfit Then peakProfit = totalCumulativeProfit
            currentDrawdown = peakProfit - totalCumulativeProfit
            drawdownpercent = currentDrawdown / (startingEquity + peakProfit + 0.000001)
            
            ' === Store Data for TotalBackTest Output ===
            totalBackTestOutput(outputRow, 1) = currentdate
            totalBackTestOutput(outputRow, 2) = totalDailyProfit
            totalBackTestOutput(outputRow, 3) = totalCumulativeProfit
            totalBackTestOutput(outputRow, 4) = currentDrawdown
            totalBackTestOutput(outputRow, 5) = Year(currentdate)
            totalBackTestOutput(outputRow, 6) = Month(currentdate)
            totalBackTestOutput(outputRow, 7) = Application.WorksheetFunction.WeekNum(currentdate, vbSunday)
            totalBackTestOutput(outputRow, 8) = drawdownpercent
            
            outputRow = outputRow + 1
            m2mOutputRow = m2mOutputRow + 1
            
            Application.StatusBar = "Backtest Running: " & Format((currentdate - earliestDate) / (latestDate - earliestDate), "0%") & " completed"
        End If
    Next dailyRow
    
    
    
    
    ' === BULK WRITE OUTPUTS TO SHEETS ===
    ' Write TotalBackTest results
    wsTotalBackTest.Range("A2").Resize(outputRow, 8).value = totalBackTestOutput
    
    
    ' === Output Symbol and Sector Profits to TotalBackTest (Starting Column S, Row 1) ===
    Dim symbolStartCol As Long, sectorStartCol As Long
    Dim symbolStartRow As Long, sectorStartRow As Long
    
    ' Set starting column (S = 19) and row
    symbolStartCol = 19 ' Column S
    sectorStartCol = symbolStartCol + 3 ' Start sector output 3 columns after symbols
    symbolStartRow = 1
    sectorStartRow = 1
    
    ' === Output Symbol Profits to TotalBackTest ===
    wsTotalBackTest.Cells(symbolStartRow, symbolStartCol).value = "Symbol Profits (Backtest Only)"
    wsTotalBackTest.Cells(symbolStartRow, symbolStartCol).Font.Bold = True
    symbolStartRow = symbolStartRow + 1
    
    Dim symbolunique As Variant
    For Each symbolunique In dictSymbolProfits.keys
        wsTotalBackTest.Cells(symbolStartRow, symbolStartCol).value = symbolunique
        wsTotalBackTest.Cells(symbolStartRow, symbolStartCol + 1).value = dictSymbolProfits(symbolunique)
        wsTotalBackTest.Cells(symbolStartRow, symbolStartCol + 1).NumberFormat = "$#,##0.00"
        symbolStartRow = symbolStartRow + 1
    Next symbolunique
    
    ' === Output Sector Profits to TotalBackTest ===
    wsTotalBackTest.Cells(sectorStartRow, sectorStartCol).value = "Sector Profits (Backtest Only)"
    wsTotalBackTest.Cells(sectorStartRow, sectorStartCol).Font.Bold = True
    sectorStartRow = sectorStartRow + 1
    
    Dim sectorunique As Variant
    For Each sectorunique In dictSectorProfits.keys
        wsTotalBackTest.Cells(sectorStartRow, sectorStartCol).value = sectorunique
        wsTotalBackTest.Cells(sectorStartRow, sectorStartCol + 1).value = dictSectorProfits(sectorunique)
        wsTotalBackTest.Cells(sectorStartRow, sectorStartCol + 1).NumberFormat = "$#,##0.00"
        sectorStartRow = sectorStartRow + 1
    Next sectorunique
    
    ' === Auto-fit Columns for Display ===
    With wsTotalBackTest
        .Columns(symbolStartCol).AutoFit
        .Columns(symbolStartCol + 1).AutoFit
        .Columns(sectorStartCol).AutoFit
        .Columns(sectorStartCol + 1).AutoFit
    End With
    
    
    ' Write BackTestM2M results
    wsBackTestM2M.Range("A4").Resize(m2mOutputRow, dictUniqueStrategies.count + 2).value = m2mOutput
 
      ' Add benchmark data if enabled
    Application.StatusBar = "Adding benchmark data..."
    Call CalculateBenchmarkData(wsTotalBackTest, wsBackTestM2M, earliestDate, latestDate)
    


    ' Autofit columns for readability
    wsTotalBackTest.Columns.AutoFit

    ThisWorkbook.Sheets("TotalBackTest").Visible = xlSheetHidden
    ThisWorkbook.Sheets("BackTestM2MEquity").Visible = xlSheetHidden
    
    
    Dim wsTotalGraphs As Worksheet
    Set wsTotalGraphs = ThisWorkbook.Sheets("TotalBackTest")
    

    Call CreatePortfolioGraphs(wsTotalGraphs, "BackTestGraphs", wsSummary)
    
    
     
    
    Set wsBackTestGraphs = ThisWorkbook.Sheets("BackTestGraphs")
     ' Calculate portfolio-level statistics
    If totalTrades > 0 Then
        portfolioWinRate = winningTrades / totalTrades
    Else
        portfolioWinRate = 0
    End If
    
    If winningTrades > 0 Then
        avgProfit = winningPnL / winningTrades
    Else
        avgProfit = 0
    End If
    
    If losingTrades > 0 Then
        avgLoss = losingPnL / losingTrades
    Else
        avgLoss = 0
    End If
    
    If avgLoss <> 0 Then
        riskToReward = Abs(avgProfit / avgLoss)
    Else
        riskToReward = 0
    End If
    
    If totalTrades > 0 Then
        edge = totalPnL / totalTrades
    Else
        edge = 0
    End If
    
     
     
    Dim FindMaxDrawdownRow As Long, lastRow As Long
    
    lastRow = wsBackTestGraphs.Cells(wsBackTestGraphs.rows.count, "B").End(xlUp).row
    
    ' Loop through each cell in column B
    For i = 1 To lastRow
        If wsBackTestGraphs.Cells(i, 2).value = "Maximum Drawdown (%)" Then
            FindMaxDrawdownRow = i
            Exit For
        End If
    Next i
     
     
     Application.StatusBar = "Updating Graphs"
    
    outputRow = FindMaxDrawdownRow + 2
    
    ' Output the results
    wsBackTestGraphs.Cells(outputRow, 2).value = "Portfolio Stats"
    wsBackTestGraphs.Cells(outputRow + 1, 2).value = "Total Trades"
    wsBackTestGraphs.Cells(outputRow + 2, 2).value = "Win Rate (%)"
    wsBackTestGraphs.Cells(outputRow + 3, 2).value = "Avg Profit"
    wsBackTestGraphs.Cells(outputRow + 4, 2).value = "Avg Loss"
    wsBackTestGraphs.Cells(outputRow + 5, 2).value = "Risk to Reward"
    wsBackTestGraphs.Cells(outputRow + 6, 2).value = "Edge"
    
    wsBackTestGraphs.Cells(outputRow + 1, 3).value = totalTrades
    wsBackTestGraphs.Cells(outputRow + 1, 3).NumberFormat = "#,##0"
    wsBackTestGraphs.Cells(outputRow + 2, 3).value = portfolioWinRate
    wsBackTestGraphs.Cells(outputRow + 2, 3).NumberFormat = "0%"
    wsBackTestGraphs.Cells(outputRow + 3, 3).value = avgProfit
    wsBackTestGraphs.Cells(outputRow + 3, 3).NumberFormat = "$#,##0"
    wsBackTestGraphs.Cells(outputRow + 4, 3).value = avgLoss
    wsBackTestGraphs.Cells(outputRow + 4, 3).NumberFormat = "$#,##0"
    wsBackTestGraphs.Cells(outputRow + 5, 3).value = riskToReward
    wsBackTestGraphs.Cells(outputRow + 5, 3).NumberFormat = "0.0"
    wsBackTestGraphs.Cells(outputRow + 6, 3).value = edge
    wsBackTestGraphs.Cells(outputRow + 6, 3).NumberFormat = "$#,##0"
    
    With wsBackTestGraphs
        ' Add Portfolio Stats header
        .Cells(outputRow + 0, 2).value = "Portfolio Stats"
        .Cells(outputRow + 0, 2).Font.Bold = True
        .Cells(outputRow + 0, 2).Font.Size = 14
        .Cells(outputRow + 0, 2).HorizontalAlignment = xlCenter
    
        ' Format labels and data
        Dim rowStart As Integer
        rowStart = outputRow + 1 ' Starting row for stats
    
        Dim statLabels As Variant
        statLabels = Array("Total Trades", "Win Rate (%)", "Avg Profit", "Avg Loss", "Risk to Reward", "Edge")
        
        Dim statValues As Variant
        statValues = Array(totalTrades, portfolioWinRate, avgProfit, avgLoss, riskToReward, edge)
        
        For i = LBound(statLabels) To UBound(statLabels)
            .Cells(rowStart + i, 2).value = statLabels(i)
            .Cells(rowStart + i, 3).value = statValues(i)
    
            ' Bold labels
            .Cells(rowStart + i, 2).Font.Bold = True
            .Cells(rowStart + i, 2).Font.Size = 12
            .Cells(rowStart + i, 2).HorizontalAlignment = xlLeft
            
            ' Align data to the right
            .Cells(rowStart + i, 3).Font.Size = 12
            .Cells(rowStart + i, 3).HorizontalAlignment = xlRight
            
            ' Add borders to cells
            .Cells(rowStart + i, 2).Borders.LineStyle = xlContinuous
            .Cells(rowStart + i, 3).Borders.LineStyle = xlContinuous
        Next i
    
        ' Apply column width adjustments
        .Columns(1).AutoFit
        .Columns(2).AutoFit
    End With
    
    With wsBackTestGraphs
        .Columns(1).ColumnWidth = 16
        .Columns(2).ColumnWidth = 16
        .Columns(3).ColumnWidth = 16
    End With
    
    
    ' Format BackTestM2MEquity sheet
    With wsBackTestM2M
        ' Format headers
        .Range(.Cells(1, 1), .Cells(1, dictUniqueStrategies.count + 1)).Font.Bold = True
        
        ' Add number formatting for profit columns
        .Range(.Cells(2, 2), .Cells(m2mOutputRow - 1, dictUniqueStrategies.count + 1)).NumberFormat = "#,##0.00"
        
        ' Format date column
        .Range(.Cells(2, 1), .Cells(m2mOutputRow - 1, 1)).NumberFormat = "dd/mm/yyyy"
    End With
    
    
    wsBackTestGraphs.Tab.Color = RGB(255, 255, 0)
    
    Call CreateStrategyAnalysis
        
    
 '   Call AddAllButtonsGraphs(wsBackTestGraphs)
    
    Call OrderVisibleTabsBasedOnList
    
    wsBackTestGraphs.Activate
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
    MsgBox "BackTest Graphs created successfully!", vbInformation



End Sub


Sub AddAllButtonsGraphs(ws As Worksheet)
    ' Remove old buttons
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    
     ' Navigation Buttons
    Dim btnDelete As Button
    Set btnDelete = ws.Buttons.Add(left:=ws.Cells(10, 1).left + 30, top:=ws.Cells(28, 1).top, Width:=100, Height:=25)
    With btnDelete
        .Caption = "Delete Tab"
        .OnAction = "DeleteBackTestGraphs"
    End With

    Dim btnSummary As Button
    Set btnSummary = ws.Buttons.Add(left:=ws.Cells(10, 3).left, top:=ws.Cells(28, 1).top, Width:=100, Height:=25)
    With btnSummary
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With

    Dim btnPortfolio As Button
    Set btnPortfolio = ws.Buttons.Add(left:=ws.Cells(10, 1).left + 30, top:=ws.Cells(31, 1).top, Width:=100, Height:=25)
    With btnPortfolio
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With

    Dim btnControl As Button
    Set btnControl = ws.Buttons.Add(left:=ws.Cells(10, 3).left, top:=ws.Cells(31, 1).top, Width:=100, Height:=25)
    With btnControl
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With

    Dim btnStrategies As Button
    Set btnStrategies = ws.Buttons.Add(left:=ws.Cells(10, 1).left + 30, top:=ws.Cells(34, 1).top, Width:=100, Height:=25)
    With btnStrategies
        .Caption = "Back to Strategies"
        .OnAction = "GoToStrategies"
    End With

    Dim btnInputs As Button
    Set btnInputs = ws.Buttons.Add(left:=ws.Cells(10, 3).left, top:=ws.Cells(34, 1).top, Width:=100, Height:=25)
    With btnInputs
        .Caption = "Back to Inputs"
        .OnAction = "GoToInputs"
    End With
    
End Sub


Sub ConsolidateTradingPeriods()
    Dim wsBacktest As Worksheet
    Dim lastRow As Long
    Dim DataRange As Range
    Dim outputRange As Range
    Dim result As Collection
    Dim key As Variant
    Dim dict As Object
    Dim strategyDict As Object
    Dim i As Long
    Dim strategyName As String
    Dim numContracts As Double
    Dim fromDate As Date, toDate As Date
    Dim entry As Variant
    Dim consolidated As Boolean
    Dim j As Long, K As Long
    Dim hasOverlaps As Boolean
    
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        Exit Sub
    End If

    ' Initialize worksheets
    On Error Resume Next
    Set wsBacktest = ThisWorkbook.Sheets("Backtest")
    On Error GoTo 0

    If wsBacktest Is Nothing Then
        MsgBox "Error: 'Backtest' sheet does not exist.", vbExclamation
        Exit Sub
    End If

    ' Clear previous results
    wsBacktest.Range("E1:Z" & wsBacktest.rows.count).ClearContents

    With wsBacktest.Range("E1")
        .value = "Consolidated ->"
        .Font.Bold = True
        .Font.Color = RGB(255, 105, 180)
        .HorizontalAlignment = xlCenter
    End With

    lastRow = wsBacktest.Cells(wsBacktest.rows.count, 1).End(xlUp).row
    Set DataRange = wsBacktest.Range("A2:D" & lastRow)
    Set dict = CreateObject("Scripting.Dictionary")
    Set strategyDict = CreateObject("Scripting.Dictionary")

    ' First pass: Collect all date ranges and check for overlaps within same strategy
    For i = 1 To DataRange.rows.count
        strategyName = DataRange.Cells(i, 1).value
        numContracts = DataRange.Cells(i, 2).value
        fromDate = DataRange.Cells(i, 3).value
        toDate = DataRange.Cells(i, 4).value

        ' Validate dates
        If fromDate > toDate Then
            MsgBox "Warning: Invalid date range found in row " & (i + 1) & ". From Date is later than To Date.", vbExclamation
            Exit Sub
        End If

        ' Check for overlaps within same strategy (different contracts)
        If Not strategyDict.Exists(strategyName) Then
            Set strategyDict(strategyName) = CreateObject("Scripting.Dictionary")
        End If
        
        ' Check overlap with existing periods for this strategy
        Dim overlappingContracts As String
        overlappingContracts = CheckOverlap(strategyDict(strategyName), fromDate, toDate, numContracts)
        
        If overlappingContracts <> "" Then
            hasOverlaps = True
            Dim msg As String
            msg = "Strategy '" & strategyName & "' has overlapping periods:" & vbCrLf & _
                  "Row " & (i + 1) & ": " & Format(fromDate, "mm/dd/yyyy") & " to " & Format(toDate, "mm/dd/yyyy") & _
                  " with " & numContracts & " contracts" & vbCrLf & _
                  "Overlaps with period using " & overlappingContracts & " contracts" & vbCrLf & vbCrLf & _
                  "Please review and correct the contract numbers before proceeding."
            
            If MsgBox(msg & vbCrLf & vbCrLf & "Do you want to continue processing other strategies?", _
                     vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If

        ' Add period to strategy dictionary
        If Not strategyDict(strategyName).Exists(numContracts) Then
            Set strategyDict(strategyName).Item(numContracts) = New Collection
        End If
        strategyDict(strategyName).Item(numContracts).Add Array(fromDate, toDate)

        ' Add to main consolidation dictionary
        key = strategyName & "|" & numContracts
        If Not dict.Exists(key) Then
            Set result = New Collection
            result.Add Array(fromDate, toDate)
            Set dict.Item(key) = result
        Else
            dict.Item(key).Add Array(fromDate, toDate)
        End If
    Next i

    ' Second pass: Sort and consolidate date ranges for each key
    For Each key In dict.keys
        Set result = dict.Item(key)
        
        ' Convert collection to array for sorting
        Dim dateArray() As Variant
        ReDim dateArray(1 To result.count, 1 To 2)
        
        For i = 1 To result.count
            dateArray(i, 1) = result(i)(0)  ' From date
            dateArray(i, 2) = result(i)(1)  ' To date
        Next i
        
        ' Sort by start date
        Call QuickSort(dateArray, 1, UBound(dateArray, 1))
        
        ' Clear original collection and add sorted, consolidated ranges
        Set result = New Collection
        
        ' Add first range
        result.Add Array(dateArray(1, 1), dateArray(1, 2))
        
        ' Consolidate overlapping ranges
        For i = 2 To UBound(dateArray, 1)
            entry = result(result.count)  ' Get last consolidated range
            
            ' Check for overlap or adjacent dates
            If dateArray(i, 1) <= entry(1) + 1 Then
                ' Merge ranges if overlapping or adjacent
                If dateArray(i, 2) > entry(1) Then
                    result.Remove result.count
                    result.Add Array(entry(0), dateArray(i, 2))
                End If
            Else
                ' Add new non-overlapping range
                result.Add Array(dateArray(i, 1), dateArray(i, 2))
            End If
        Next i
        
        ' Update dictionary with consolidated ranges
        Set dict.Item(key) = result
    Next key

    ' Output headers
    Set outputRange = wsBacktest.Range("F1")
    With wsBacktest.Range("F1:I1")
        outputRange.value = "Strategy Name"
        outputRange.Offset(0, 1).value = "Number of Contracts"
        outputRange.Offset(0, 2).value = "From Date"
        outputRange.Offset(0, 3).value = "To Date"
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .WrapText = True
    End With

    ' Output results
    Dim outputRow As Long
    outputRow = 2
    For Each key In dict.keys
        Set result = dict.Item(key)
        For Each entry In result
            wsBacktest.Cells(outputRow, 6).value = Split(key, "|")(0)
            wsBacktest.Cells(outputRow, 7).value = Split(key, "|")(1)
            wsBacktest.Cells(outputRow, 8).value = entry(0)
            wsBacktest.Cells(outputRow, 9).value = entry(1)
            outputRow = outputRow + 1
        Next
    Next

    ' Format output
    wsBacktest.Columns("A:Z").AutoFit
    wsBacktest.rows(1).WrapText = True
    wsBacktest.Activate

    If hasOverlaps Then
        MsgBox "Consolidation completed with warnings. Please review the overlapping periods mentioned earlier.", vbExclamation
    Else
        MsgBox "Consolidation complete!"
    End If
End Sub

Private Function CheckOverlap(strategyPeriods As Object, newFromDate As Date, newToDate As Date, newContracts As Double) As String
    Dim contracts As Variant
    Dim periodCollection As Collection
    Dim period As Variant
    
    For Each contracts In strategyPeriods.keys
        If contracts <> newContracts Then  ' Only check different contract numbers
            Set periodCollection = strategyPeriods(contracts)
            For Each period In periodCollection
                ' Check if periods overlap
                If (newFromDate <= period(1)) And (newToDate >= period(0)) Then
                    CheckOverlap = contracts & " contracts"
                    Exit Function
                End If
            Next period
        End If
    Next contracts
    
    CheckOverlap = ""  ' No overlap found
End Function

' QuickSort implementation for date ranges
Private Sub QuickSort(arr() As Variant, low As Long, high As Long)
    Dim pivot As Date
    Dim temp As Variant
    Dim i As Long
    Dim j As Long
    
    If low < high Then
        pivot = arr((low + high) \ 2, 1)
        i = low
        j = high
        
        Do
            Do While arr(i, 1) < pivot
                i = i + 1
            Loop
            
            Do While arr(j, 1) > pivot
                j = j - 1
            Loop
            
            If i <= j Then
                ' Swap dates
                temp = arr(i, 1)
                arr(i, 1) = arr(j, 1)
                arr(j, 1) = temp
                
                ' Swap corresponding end dates
                temp = arr(i, 2)
                arr(i, 2) = arr(j, 2)
                arr(j, 2) = temp
                
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        
        If low < j Then Call QuickSort(arr, low, j)
        If i < high Then Call QuickSort(arr, i, high)
    End If
End Sub



Private Function CompareStrategyNames(name1 As String, name2 As String) As Boolean
    ' Safely compare two strategy names, ignoring case and trimming whitespace
    ' Returns True if the names match, False otherwise
    
    ' Handle Null or Empty values
    If IsNull(name1) Or IsNull(name2) Then
        CompareStrategyNames = False
        Exit Function
    End If
    
    ' Convert to string and trim
    Dim str1 As String, str2 As String
    str1 = Trim(CStr(name1))
    str2 = Trim(CStr(name2))
    
    ' Compare using case-insensitive comparison
    CompareStrategyNames = (StrComp(str1, str2, vbTextCompare) = 0)
End Function


Public Sub CreateStrategyAnalysis()
    ' Create or get the analysis sheet
    Dim wsBacktestDetails As Worksheet
    
    Call Deletetab("BacktestDetails")
    
    On Error Resume Next
    Set wsBacktestDetails = ThisWorkbook.Sheets("BacktestDetails")
    On Error GoTo 0
    
    If wsBacktestDetails Is Nothing Then
        Set wsBacktestDetails = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("BackTestGraphs"))
        wsBacktestDetails.name = "BacktestDetails"
        wsBacktestDetails.Tab.Color = RGB(255, 255, 0) ' Yellow tab color
        wsBacktestDetails.Cells.Interior.Color = RGB(255, 255, 255) ' White background
    End If
    
    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    
    ' ============================
    ' Set Up Interface
    ' ============================
    With wsBacktestDetails
        ' Title
        .Range("B1:L1").Merge
        .Range("B1:B1").value = "BackTest - Details Tab"
        With .Range("B1:B1")
            .Font.Bold = True
            .Font.Size = 14
            .Font.name = "Calibri"
            .Interior.Color = RGB(220, 230, 241)
        End With
        
        ' Category Dropdown
        .Range("A2:A2").value = "Select Category:"
        With .Range("A2:A2")
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlRight
        End With
        .Range("B2:L2").Merge
        With .Range("B2:L2")
            .Interior.Color = RGB(225, 225, 225)
            .Borders.LineStyle = xlContinuous
            .Font.Size = 12
            .HorizontalAlignment = xlLeft
        End With
        With .Range("B2").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="Strategy,Symbol,Sector"
        End With
        
        ' Dependent Dropdown
        .Range("A3:A3").value = "Select Option:"
        With .Range("A3:A3")
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlRight
        End With
        .Range("B3:L3").Merge
        With .Range("B3:L3")
            .Interior.Color = RGB(225, 225, 225)
            .Borders.LineStyle = xlContinuous
            .Font.Size = 12
            .HorizontalAlignment = xlLeft
        End With
        
        
         ' Format columns
       .Columns(1).EntireColumn.ColumnWidth = 25
       .Columns(2).EntireColumn.ColumnWidth = 18
       .Range(.Cells(1, COL_BACKSTRAT_DATE), .Cells(1, COL_BACKSTRAT_WEEK)).EntireColumn.ColumnWidth = 15
       .Columns(COL_BACKSTRAT_DATE).NumberFormat = "dd/mm/yyyy"
       .Range(.Columns(COL_BACKSTRAT_DAILY_PL), .Columns(COL_BACKSTRAT_DRAWDOWN)).NumberFormat = "#,##0.00"
       .Range(.Columns(COL_BACKSTRAT_YEAR), .Columns(COL_BACKSTRAT_WEEK)).NumberFormat = "0"
        
        ' ============================
        ' Collect Unique Lists
        ' ============================
        Call CollectUniqueLists
        
        ' ============================
        ' Add Buttons
        ' ============================
        Call AddAllButtons(wsBacktestDetails)
        
        ' ============================
        ' Automatically Update Dropdown
        ' ============================
        Range("B2").value = "Strategy" ' default value
        Call UpdateSecondDropdown
        
        ' Set zoom level
        
        
        
 ' Set up data columns using predefined constants
        With .Range(.Cells(1, COL_BACKSTRAT_DATE), .Cells(1, COL_BACKSTRAT_WEEK))
            .value = Array("Date", "Daily P&L", "Cumulative P&L", "Drawdown", "Year", "Month", "Week")
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 241)
            .Borders.LineStyle = xlContinuous
        End With
        
        
      
        
        ThisWorkbook.Windows(1).Zoom = 70
    End With
End Sub



Sub CollectUniqueLists()
    Dim wsM2M As Worksheet, wsStrategy As Worksheet
    Set wsM2M = ThisWorkbook.Sheets("BackTestM2MEquity")
    Set wsStrategy = ThisWorkbook.Sheets("BacktestDetails")

    Dim strategyDict As Object, symbolDict As Object, sectorDict As Object
    Set strategyDict = CreateObject("Scripting.Dictionary")
    Set symbolDict = CreateObject("Scripting.Dictionary")
    Set sectorDict = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = wsM2M.Cells(1, wsM2M.Columns.count).End(xlToLeft).column

    Dim col As Long
    For col = 2 To lastCol
        Dim strategyVal As String, symbolVal As String, sectorVal As String
        strategyVal = wsM2M.Cells(1, col).value
        symbolVal = wsM2M.Cells(2, col).value
        sectorVal = wsM2M.Cells(3, col).value

        If strategyVal <> "" And Not strategyDict.Exists(strategyVal) Then strategyDict(strategyVal) = ""
        If symbolVal <> "" And Not symbolDict.Exists(symbolVal) Then symbolDict(symbolVal) = ""
        If sectorVal <> "" And Not sectorDict.Exists(sectorVal) Then sectorDict(sectorVal) = ""
    Next col

    ' Write lists to hidden columns
    wsStrategy.Columns("BX:BZ").ClearContents
    Dim stratRng As Range, symRng As Range, secRng As Range

    ' Strategy List
    Set stratRng = wsStrategy.Range("BX1")
    stratRng.Resize(strategyDict.count, 1).value = Application.Transpose(strategyDict.keys)
    CreateNamedRange "StrategyList", wsStrategy, stratRng, strategyDict.count

    ' Symbol List
    Set symRng = wsStrategy.Range("BY1")
    symRng.Resize(symbolDict.count, 1).value = Application.Transpose(symbolDict.keys)
    CreateNamedRange "SymbolList", wsStrategy, symRng, symbolDict.count

    ' Sector List
    Set secRng = wsStrategy.Range("BZ1")
    secRng.Resize(sectorDict.count, 1).value = Application.Transpose(sectorDict.keys)
    CreateNamedRange "SectorList", wsStrategy, secRng, sectorDict.count

    ' Hide the columns
    wsStrategy.Columns("BX:BZ").Hidden = True
End Sub

Sub CreateNamedRange(name As String, ws As Worksheet, startCell As Range, count As Long)
    On Error Resume Next
    ThisWorkbook.Names(name).Delete
    On Error GoTo 0

    If count > 0 Then
        Dim rng As Range
        Set rng = ws.Range(startCell, startCell.Offset(count - 1, 0))
        ThisWorkbook.Names.Add name:=name, RefersTo:=rng
    End If
End Sub

Sub UpdateSecondDropdown()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BacktestDetails")

    Dim category As String
    category = ws.Range("B2").value

    Dim namedRange As String
    Select Case category
        Case "Strategy": namedRange = "StrategyList"
        Case "Symbol": namedRange = "SymbolList"
        Case "Sector": namedRange = "SectorList"
        Case Else: namedRange = ""
    End Select

    ' === Handle Validation Removal with Error Handling ===
    On Error Resume Next
    ws.Range("B3").Validation.Delete
    On Error GoTo 0

    ' === Add New Validation and Default to First Value ===
    If namedRange <> "" Then
        With ws.Range("B3").Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=" & namedRange
            .IgnoreBlank = True
            .InCellDropdown = True
        End With

        ' === Default B3 to the First Value in Named Range ===
        On Error Resume Next
        ws.Range("B3").value = ThisWorkbook.Names(namedRange).RefersToRange.Cells(1, 1).value
        On Error GoTo 0
    Else
        ws.Range("B3").value = ""
    End If
End Sub


Sub AddAllButtons(ws As Worksheet)
    ' Remove old buttons
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0

    ' ============================
    ' Main Update Options Button
    ' ============================
    Dim btnUpdate As Button
    Set btnUpdate = ws.Buttons.Add(left:=ws.Cells(5, 1).left + 10, top:=ws.Cells(5, 1).top, Width:=200, Height:=25)
    With btnUpdate
        .OnAction = "UpdateSecondDropdown"
        .Caption = "Update Options"
        .Font.Size = 12
        .Font.Bold = True
        .name = "btnUpdateOptions"
    End With

    ' Strategy Button
    Dim btnStrategy As Button
    Set btnStrategy = ws.Buttons.Add(left:=ws.Cells(5, 4).left + 10, top:=ws.Cells(5, 1).top, Width:=200, Height:=25)
    With btnStrategy
        .OnAction = "RunAnalyze"
        .Caption = "Open Detailed Graphs"
        .Font.Size = 12
        .Font.Bold = True
        .name = "btnStrategy"
    End With

    ' Navigation Buttons
    Dim btnDelete As Button
    Set btnDelete = ws.Buttons.Add(left:=ws.Cells(10, 1).left + 10, top:=ws.Cells(10, 1).top, Width:=100, Height:=25)
    With btnDelete
        .Caption = "Delete Tab"
        .OnAction = "DeleteBacktestDetails"
    End With

    Dim btnSummary As Button
    Set btnSummary = ws.Buttons.Add(left:=ws.Cells(10, 2).left, top:=ws.Cells(10, 1).top, Width:=100, Height:=25)
    With btnSummary
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With

    Dim btnPortfolio As Button
    Set btnPortfolio = ws.Buttons.Add(left:=ws.Cells(10, 1).left + 10, top:=ws.Cells(13, 1).top, Width:=100, Height:=25)
    With btnPortfolio
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With

    Dim btnControl As Button
    Set btnControl = ws.Buttons.Add(left:=ws.Cells(10, 2).left, top:=ws.Cells(13, 1).top, Width:=100, Height:=25)
    With btnControl
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With

    Dim btnStrategies As Button
    Set btnStrategies = ws.Buttons.Add(left:=ws.Cells(10, 1).left + 10, top:=ws.Cells(16, 1).top, Width:=100, Height:=25)
    With btnStrategies
        .Caption = "Back to Strategies"
        .OnAction = "GoToStrategies"
    End With

    Dim btnInputs As Button
    Set btnInputs = ws.Buttons.Add(left:=ws.Cells(10, 2).left, top:=ws.Cells(16, 1).top, Width:=100, Height:=25)
    With btnInputs
        .Caption = "Back to Inputs"
        .OnAction = "GoToInputs"
    End With
    
        
    
End Sub






Public Sub RunAnalyze()
    ' Get the worksheets
    Dim wsBacktestDetails As Worksheet
    Dim wsBackTestM2M As Worksheet
    Set wsBacktestDetails = ThisWorkbook.Sheets("BacktestDetails")  ' Updated sheet name
    Set wsBackTestM2M = ThisWorkbook.Sheets("BackTestM2MEquity")
    Dim AnnualprofitsCol As Long
    Dim selectionType As Long
    Dim row As Double
    Dim i As Long
    
      ' Get selected value based on selection type
    Dim selectedOption As String, selectedValue As String
    
    selectedOption = Range("B2").value
    
     
    
    
    Select Case selectedOption
        Case "Strategy": selectionType = 1
        Case "Symbol": selectionType = 2
        Case "Sector": selectionType = 3
        Case Else
            MsgBox "Invalid selection type!", vbExclamation
            Exit Sub
    End Select
    
    If selectedOption = "" Then
        MsgBox "Please select a Category first!", vbExclamation
        Exit Sub
    End If
    
    
    selectedValue = Range("B3").value
    
    If selectedValue = "" Then
        MsgBox "Please select an Option first!", vbExclamation
        Exit Sub
    End If
    
    
     
    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    
    Application.ScreenUpdating = False
    
  
    ' Clear existing data and charts
    wsBacktestDetails.Range(wsBacktestDetails.Cells(2, COL_BACKSTRAT_DATE), _
        wsBacktestDetails.Cells(1000, COL_BACKSTRAT_WEEK)).Clear
    On Error Resume Next
    wsBacktestDetails.ChartObjects.Delete
    On Error GoTo 0
    
    ' Get last row of data
    Dim lastRow As Long
    lastRow = wsBackTestM2M.Cells(wsBackTestM2M.rows.count, 1).End(xlUp).row
    

  ' Find strategy column in BackTestM2MEquity
    Dim lastCol As Long
    lastCol = wsBackTestM2M.Cells(1, wsBackTestM2M.Columns.count).End(xlToLeft).column
    
      ' Collect matching columns based on selection
    Dim matchingCols As Collection
    Set matchingCols = New Collection

    
    Dim col As Long
    If selectionType = 1 Then
        ' Strategy: Find single column match
        For col = 2 To lastCol
            If wsBackTestM2M.Cells(1, col).value = selectedValue Then
                matchingCols.Add col
                Exit For
            End If
        Next col
    ElseIf selectionType = 2 Then
        ' Symbol: Find all columns for this symbol
        For col = 2 To lastCol
            If wsBackTestM2M.Cells(2, col).value = selectedValue Then
                matchingCols.Add col
            End If
        Next col
    ElseIf selectionType = 3 Then
        ' Sector: Find all columns for this sector
        For col = 2 To lastCol
            If wsBackTestM2M.Cells(3, col).value = selectedValue Then
                matchingCols.Add col
            End If
        Next col
    End If
    
    If matchingCols.count = 0 Then
        MsgBox "No matching columns found for " & selectedValue, vbExclamation
        Exit Sub
    End If


    ' Variables for calculations
    Dim outputRow As Long
    outputRow = 2
    Dim peakProfit As Double
    peakProfit = 0
    Dim cumulativeProfit As Double
    cumulativeProfit = 0
    
    ' Create dictionaries for data analysis
    Dim annualProfitsDict As Object
    Dim annualMaxDrawdownDict As Object
    Dim monthlyProfitsDict As Object
    Set annualProfitsDict = CreateObject("Scripting.Dictionary")
    Set annualMaxDrawdownDict = CreateObject("Scripting.Dictionary")
    Set monthlyProfitsDict = CreateObject("Scripting.Dictionary")
    
    Dim distinctMonthsDict As Object
    Set distinctMonthsDict = CreateObject("Scripting.Dictionary")
    
    ' Create dictionaries for monthly, weekly, and average monthly profits
    Dim weeklyProfitsDict As Object
    Dim monthOnlyProfitsDict As Object
    

    Set weeklyProfitsDict = CreateObject("Scripting.Dictionary")
    Set monthOnlyProfitsDict = CreateObject("Scripting.Dictionary")
    
    ' Initialize month names
    Dim monthNames(1 To 12) As String
    monthNames(1) = "January": monthNames(2) = "February": monthNames(3) = "March"
    monthNames(4) = "April": monthNames(5) = "May": monthNames(6) = "June"
    monthNames(7) = "July": monthNames(8) = "August": monthNames(9) = "September"
    monthNames(10) = "October": monthNames(11) = "November": monthNames(12) = "December"

    ' Initialize month count array
    Dim monthCountArray(1 To 12) As Long
    Dim monthProfitArray(1 To 12) As Double
    Dim monthKeyString As String
    
           
             
    
    
    
    ' Process data
    For row = 4 To lastRow
        Dim currentdate As Date
        currentdate = wsBackTestM2M.Cells(row, 1).value
        
        ' Sum daily profit from all matching columns and track errors
        Dim dailyProfit As Double: dailyProfit = 0
        Dim errorCount As Long: errorCount = 0
        Dim currentCol As Variant
        
        For Each currentCol In matchingCols
            If IsError(wsBackTestM2M.Cells(row, currentCol).value) Then
                errorCount = errorCount + 1
            Else
                dailyProfit = dailyProfit + wsBackTestM2M.Cells(row, currentCol).value
            End If
        Next currentCol
        
        ' Check if all columns were errors
        If errorCount = matchingCols.count Then
            ' Output error if all columns are errors
            With wsBacktestDetails
                .Cells(outputRow, COL_BACKSTRAT_DATE).value = currentdate
                .Cells(outputRow, COL_BACKSTRAT_DAILY_PL).Formula = "=NA()"
                .Cells(outputRow, COL_BACKSTRAT_CUM_PL).Formula = "=NA()"
                .Cells(outputRow, COL_BACKSTRAT_DRAWDOWN).Formula = "=NA()"
                .Cells(outputRow, COL_BACKSTRAT_YEAR).value = Year(currentdate)
                .Cells(outputRow, COL_BACKSTRAT_MONTH).value = Month(currentdate)
                .Cells(outputRow, COL_BACKSTRAT_WEEK).value = Application.WorksheetFunction.WeekNum(currentdate, vbSunday)
            End With
            GoTo NextRow
        End If
        
          ' Calculate cumulative profit and drawdown
        cumulativeProfit = cumulativeProfit + dailyProfit
        If cumulativeProfit > peakProfit Then peakProfit = cumulativeProfit
        Dim currentDrawdown As Double
        currentDrawdown = peakProfit - cumulativeProfit
        
        
         ' Write Daily Results
        With wsBacktestDetails
            .Cells(outputRow, COL_BACKSTRAT_DATE).value = currentdate
            .Cells(outputRow, COL_BACKSTRAT_DAILY_PL).value = dailyProfit
            .Cells(outputRow, COL_BACKSTRAT_CUM_PL).value = cumulativeProfit
            .Cells(outputRow, COL_BACKSTRAT_DRAWDOWN).value = currentDrawdown
            .Cells(outputRow, COL_BACKSTRAT_YEAR).value = Year(currentdate)
            .Cells(outputRow, COL_BACKSTRAT_MONTH).value = Month(currentdate)
            .Cells(outputRow, COL_BACKSTRAT_WEEK).value = Application.WorksheetFunction.WeekNum(currentdate, vbSunday)
        End With
        
        ' Annual Data
        Dim currentYear As String
        currentYear = Year(currentdate)
        If Not annualProfitsDict.Exists(currentYear) Then
            annualProfitsDict.Add currentYear, 0
            annualMaxDrawdownDict.Add currentYear, 0
        End If
        annualProfitsDict(currentYear) = annualProfitsDict(currentYear) + dailyProfit
        If currentDrawdown > annualMaxDrawdownDict(currentYear) Then
            annualMaxDrawdownDict(currentYear) = currentDrawdown
        End If
        
        ' Monthly Data
        Dim currentMonth As String
        currentMonth = currentYear & "-" & Format(Month(currentdate), "00")
        If Not monthlyProfitsDict.Exists(currentMonth) Then
            monthlyProfitsDict.Add currentMonth, 0
        End If
        monthlyProfitsDict(currentMonth) = monthlyProfitsDict(currentMonth) + dailyProfit
        
        ' Weekly Data
        Dim currentWeek As String
        currentWeek = currentMonth & ": W" & Application.WorksheetFunction.WeekNum(currentdate, vbSunday)
        If Not weeklyProfitsDict.Exists(currentWeek) Then
            weeklyProfitsDict.Add currentWeek, 0
        End If
        weeklyProfitsDict(currentWeek) = weeklyProfitsDict(currentWeek) + dailyProfit
        
        ' Month-Only Data (for Average Monthly Profits)
        'Dim monthKeyString As String
        monthKeyString = Year(currentdate) & "-" & Format(Month(currentdate), "00")
        If Not monthOnlyProfitsDict.Exists(monthKeyString) Then
            monthOnlyProfitsDict.Add monthKeyString, 0
        End If
        
        
        ' Count Distinct Months
        Dim monthIndex As Integer
        monthIndex = Month(currentdate)
        
        monthProfitArray(monthIndex) = monthProfitArray(monthIndex) + dailyProfit
        
        If Not distinctMonthsDict.Exists(monthKeyString) Then
            distinctMonthsDict.Add monthKeyString, True
            monthCountArray(monthIndex) = monthCountArray(monthIndex) + 1
        End If
        'outputRow = outputRow + 1
        
NextRow:
   outputRow = outputRow + 1
        
    Next row
    
    
    AnnualprofitsCol = COL_BACKSTRAT_WEEK + 2
    
    'Create annual summary table with formatting
    With wsBacktestDetails
       .Cells(1, AnnualprofitsCol).value = "Annual Summary"
       .Cells(1, AnnualprofitsCol).Font.Bold = True
       .Cells(1, AnnualprofitsCol).Font.Size = 12
       .Cells(1, AnnualprofitsCol).Interior.Color = RGB(220, 230, 241)
       
       With .Range(.Cells(2, AnnualprofitsCol), .Cells(2, AnnualprofitsCol + 2))
           .value = Array("Year", "Annual Profit", "Max Drawdown")
           .Font.Bold = True
           .Interior.Color = RGB(220, 230, 241)
           .Borders.LineStyle = xlContinuous
       End With
    End With
    
    ' Fill annual summary data
    Dim summaryRow As Long
    summaryRow = 3
    Dim yearKey As Variant
    For Each yearKey In annualProfitsDict.keys
       With wsBacktestDetails
           .Cells(summaryRow, AnnualprofitsCol).value = yearKey
           .Cells(summaryRow, AnnualprofitsCol + 1).value = annualProfitsDict(yearKey)
           .Cells(summaryRow, AnnualprofitsCol + 2).value = annualMaxDrawdownDict(yearKey)
       End With
       summaryRow = summaryRow + 1
    Next yearKey
    
    ' ===============================
    ' MONTHLY PROFITS (LAST 5 YEARS)
    ' ===============================
    Dim monthlyStartCol As Long
    monthlyStartCol = AnnualprofitsCol + 4
    
    With wsBacktestDetails
        .Cells(1, monthlyStartCol).value = "Monthly Profits (Last 5 Years)"
        .Cells(2, monthlyStartCol).value = "Month"
        .Cells(2, monthlyStartCol + 1).value = "Profit"
        .Range(.Cells(2, monthlyStartCol), .Cells(2, monthlyStartCol + 1)).Font.Bold = True
    End With
    
    Dim monthlyRow As Long: monthlyRow = 3
    Dim monthKeyCheck As Variant
    For Each monthKeyCheck In monthlyProfitsDict.keys
        With wsBacktestDetails
            .Cells(monthlyRow, monthlyStartCol).value = monthKeyCheck
            .Cells(monthlyRow, monthlyStartCol + 1).value = monthlyProfitsDict(monthKeyCheck)
        End With
        monthlyRow = monthlyRow + 1
    Next monthKeyCheck
    
    ' ===============================
    ' WEEKLY PROFITS (LAST 52 WEEKS)
    ' ===============================
    Dim weeklyStartCol As Long
    weeklyStartCol = monthlyStartCol + 3
    
    With wsBacktestDetails
        .Cells(1, weeklyStartCol).value = "Weekly Profits (Last 52 Weeks)"
        .Cells(2, weeklyStartCol).value = "Week"
        .Cells(2, weeklyStartCol + 1).value = "Profit"
        .Range(.Cells(2, weeklyStartCol), .Cells(2, weeklyStartCol + 1)).Font.Bold = True
    End With
    
    Dim weeklyRow As Long: weeklyRow = 3
    Dim weekKey As Variant
    For Each weekKey In weeklyProfitsDict.keys
        With wsBacktestDetails
            .Cells(weeklyRow, weeklyStartCol).value = weekKey
            .Cells(weeklyRow, weeklyStartCol + 1).value = weeklyProfitsDict(weekKey)
        End With
        weeklyRow = weeklyRow + 1
    Next weekKey
    
    ' ==============================
    ' Average Monthly Profits Table
    ' ==============================
    
    Dim avgMonthlyStartCol As Long
    avgMonthlyStartCol = 3 + weeklyStartCol
    
    With wsBacktestDetails
        .Cells(1, avgMonthlyStartCol).value = "Average Monthly Profits (Distinct Months)"
        .Cells(2, avgMonthlyStartCol).value = "Month"
        .Cells(2, avgMonthlyStartCol + 1).value = "Average Profit"
        .Range(.Cells(2, avgMonthlyStartCol), .Cells(2, avgMonthlyStartCol + 1)).Font.Bold = True
    End With
    
    Dim avgMonthlyRow As Long: avgMonthlyRow = 3
    Dim totalMonthProfit As Double, totalMonthCount As Long
    
    For i = 1 To 12
        With wsBacktestDetails
            .Cells(avgMonthlyRow, avgMonthlyStartCol).value = monthNames(i)
            
            ' Calculate average only if distinct months exist
            If monthCountArray(i) > 0 Then
                .Cells(avgMonthlyRow, avgMonthlyStartCol + 1).value = monthProfitArray(i) / monthCountArray(i)
            Else
                .Cells(avgMonthlyRow, avgMonthlyStartCol + 1).value = 0
            End If
        End With
        avgMonthlyRow = avgMonthlyRow + 1
    Next i
    
    
    
    ' Format data ranges
    With wsBacktestDetails
       ' Format data grid
       With .Range(.Cells(1, COL_BACKSTRAT_DATE), .Cells(outputRow - 1, COL_BACKSTRAT_WEEK))
           .Borders.LineStyle = xlContinuous
           .Font.Size = 11
       End With
       
       ' Format summary table
       With .Range(.Cells(1, AnnualprofitsCol), .Cells(summaryRow - 1, AnnualprofitsCol + 2))
           .Borders.LineStyle = xlContinuous
           .Font.Size = 11
       End With
       
       ' Number formatting
       .Range(.Cells(3, AnnualprofitsCol + 1), .Cells(summaryRow - 1, AnnualprofitsCol + 2)).NumberFormat = "#,##0.00"
    End With
    
    
    ' Apply border formatting
    With wsBacktestDetails
        Dim summaryRange As Range
        Set summaryRange = .Range(.Cells(1, monthlyStartCol), .Cells(monthlyRow - 1, monthlyStartCol + 1))
        summaryRange.Borders.LineStyle = xlContinuous
    
        Set summaryRange = .Range(.Cells(1, weeklyStartCol), .Cells(weeklyRow - 1, weeklyStartCol + 1))
        summaryRange.Borders.LineStyle = xlContinuous
    
        Set summaryRange = .Range(.Cells(1, avgMonthlyStartCol), .Cells(avgMonthlyRow - 1, avgMonthlyStartCol + 1))
        summaryRange.Borders.LineStyle = xlContinuous
    End With
 
    ' Cumulative Profits Chart
    
    Dim cht As ChartObject
    
    Set cht = wsBacktestDetails.ChartObjects.Add(left:=300, Width:=600, top:=100, Height:=300)
    With cht.chart
        .ChartType = xlLine
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .values = wsBacktestDetails.Range(wsBacktestDetails.Cells(2, COL_BACKSTRAT_CUM_PL), _
                     wsBacktestDetails.Cells(outputRow - 1, COL_BACKSTRAT_CUM_PL))
            .XValues = wsBacktestDetails.Range(wsBacktestDetails.Cells(2, COL_BACKSTRAT_DATE), _
                      wsBacktestDetails.Cells(outputRow - 1, COL_BACKSTRAT_DATE))
            .name = "Cumulative P&L"
        End With
        
        ' Format Title
        .HasTitle = True
        .chartTitle.text = selectedValue & " - Cumulative P&L"
        .HasLegend = False
        
        ' Format Y Axis
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.text = "Profit/Loss ($)"
            .TickLabels.NumberFormat = "$#,##0"
            .MajorGridlines.Format.line.Visible = True
            .MajorGridlines.Format.line.ForeColor.RGB = RGB(240, 240, 240)
        End With
        
        ' Format X Axis
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.text = "Date"
            .TickLabels.NumberFormat = "mmm-yy"
            .TickLabelPosition = xlTickLabelPositionLow
        End With
    End With
    
    ' Drawdown Chart
    Set cht = wsBacktestDetails.ChartObjects.Add(left:=950, Width:=600, top:=100, Height:=300)
    With cht.chart
        .ChartType = xlLine
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .values = wsBacktestDetails.Range(wsBacktestDetails.Cells(2, COL_BACKSTRAT_DRAWDOWN), _
                     wsBacktestDetails.Cells(outputRow - 1, COL_BACKSTRAT_DRAWDOWN))
            .XValues = wsBacktestDetails.Range(wsBacktestDetails.Cells(2, COL_BACKSTRAT_DATE), _
                      wsBacktestDetails.Cells(outputRow - 1, COL_BACKSTRAT_DATE))
            .name = "Drawdown"
            .Format.line.ForeColor.RGB = RGB(255, 0, 0) ' Red line for drawdown
        End With
        
        ' Format Title
        .HasTitle = True
        .chartTitle.text = selectedValue & " - Drawdown"
        .HasLegend = False
        
        ' Format Y Axis
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.text = "Drawdown ($)"
            .TickLabels.NumberFormat = "$#,##0"
            .MajorGridlines.Format.line.Visible = True
            .MajorGridlines.Format.line.ForeColor.RGB = RGB(240, 240, 240)
        End With
        
        ' Format X Axis
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.text = "Date"
            .TickLabels.NumberFormat = "mmm-yy"
            .TickLabelPosition = xlTickLabelPositionLow
        End With
    End With
    
    ' Annual Profit Bar Chart
    Set cht = wsBacktestDetails.ChartObjects.Add(left:=300, Width:=600, top:=450, Height:=300)
    With cht.chart
        .ChartType = xlColumnClustered
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .values = wsBacktestDetails.Range(wsBacktestDetails.Cells(3, AnnualprofitsCol + 1), _
                     wsBacktestDetails.Cells(summaryRow - 1, AnnualprofitsCol + 1))
            .XValues = wsBacktestDetails.Range(wsBacktestDetails.Cells(3, AnnualprofitsCol), _
                      wsBacktestDetails.Cells(summaryRow - 1, AnnualprofitsCol))
            .name = "Annual Profit"
        End With
        
        ' Format Title
        .HasTitle = True
        .chartTitle.text = selectedValue & " - Annual Profits"
        .HasLegend = False
        
        ' Format Y Axis
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.text = "Profit/Loss ($)"
            .TickLabels.NumberFormat = "$#,##0"
            .MajorGridlines.Format.line.Visible = True
            .MajorGridlines.Format.line.ForeColor.RGB = RGB(240, 240, 240)
        End With
        
        ' Format X Axis
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.text = "Year"
            .MajorGridlines.Format.line.Visible = False
        End With
    End With
    
    
    ' ===========================
    ' Monthly Profits Chart
    ' ===========================
    Dim monthlyChart As ChartObject
    Set monthlyChart = wsBacktestDetails.ChartObjects.Add(left:=950, Width:=600, top:=450, Height:=300)
    
    With monthlyChart.chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.text = selectedValue & "Monthly Profits (Last 5 Years)"
        
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .XValues = wsBacktestDetails.Range(wsBacktestDetails.Cells(3, monthlyStartCol), _
                                               wsBacktestDetails.Cells(monthlyRow - 1, monthlyStartCol))
            .values = wsBacktestDetails.Range(wsBacktestDetails.Cells(3, monthlyStartCol + 1), _
                                              wsBacktestDetails.Cells(monthlyRow - 1, monthlyStartCol + 1))
            .name = "Monthly Profits"
        End With
        
        ' Format X and Y Axes
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.text = "Month"
            .TickLabels.Orientation = xlUpward
        End With
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.text = "Profit"
            .TickLabels.NumberFormat = "$#,##0"
        End With
        
        .HasLegend = False
    End With
    
    ' ===========================
    ' Weekly Profits Chart
    ' ===========================
    Dim weeklyChart As ChartObject
    Set weeklyChart = wsBacktestDetails.ChartObjects.Add(left:=300, Width:=600, top:=800, Height:=300)
    
    With weeklyChart.chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.text = selectedValue & "Weekly Profits (Last 52 Weeks)"
        
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .XValues = wsBacktestDetails.Range(wsBacktestDetails.Cells(3, weeklyStartCol), _
                                               wsBacktestDetails.Cells(weeklyRow - 1, weeklyStartCol))
            .values = wsBacktestDetails.Range(wsBacktestDetails.Cells(3, weeklyStartCol + 1), _
                                              wsBacktestDetails.Cells(weeklyRow - 1, weeklyStartCol + 1))
            .name = "Weekly Profits"
        End With
        
        ' Format X and Y Axes
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.text = "Week"
            .TickLabels.Orientation = xlUpward
        End With
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.text = "Profit"
            .TickLabels.NumberFormat = "$#,##0"
        End With
        
        .HasLegend = False
    End With
      
      
    ' ================================
    ' Average Monthly Profits Chart
    ' ================================
    Dim avgMonthlyChart As ChartObject
    Set avgMonthlyChart = wsBacktestDetails.ChartObjects.Add(left:=950, Width:=600, top:=800, Height:=300)
    
    With avgMonthlyChart.chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.text = selectedValue & "Average Monthly Profits"
        
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .XValues = wsBacktestDetails.Range(wsBacktestDetails.Cells(3, avgMonthlyStartCol), _
                                               wsBacktestDetails.Cells(avgMonthlyRow - 1, avgMonthlyStartCol))
            .values = wsBacktestDetails.Range(wsBacktestDetails.Cells(3, avgMonthlyStartCol + 1), _
                                              wsBacktestDetails.Cells(avgMonthlyRow - 1, avgMonthlyStartCol + 1))
            .name = "Average Monthly Profits"
        End With
        
        ' Format X and Y Axes
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.text = "Month"
            .TickLabels.Orientation = xlUpward
        End With
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.text = "Average Profit"
            .TickLabels.NumberFormat = "$#,##0"
        End With
        
        .HasLegend = False
    End With
    
    Application.ScreenUpdating = True
End Sub


Sub ToggleDisplayModeBacktest()
    ' Toggle the DisplayAsPercentage named range
    Dim currentValue As Boolean
    
    ' Check if the named range exists
    On Error Resume Next
    currentValue = Range("DisplayAsPercentage").value
    
    ' If it doesn't exist, create it and set to True
    If Err.Number <> 0 Then
        On Error GoTo 0
        ' Create the named range
        ThisWorkbook.Names.Add name:="DisplayAsPercentage", RefersTo:=True
        MsgBox "Display mode set to Percentage. Regenerating graphs...", vbInformation
    Else
        On Error GoTo 0
        ' Toggle the existing value
        Range("DisplayAsPercentage").value = Not currentValue
        
        If Not currentValue Then
            MsgBox "Display mode set to Percentage. Regenerating graphs...", vbInformation
        Else
            MsgBox "Display mode set to Dollar. Regenerating graphs...", vbInformation
        End If
    End If
    
    ' Regenerate the portfolio graphs with the new display mode
    Application.ScreenUpdating = False
    
    ' Get the active sheet so we can return to it
    Dim activeSheet As Worksheet
    Set activeSheet = activeSheet
    
    ' Get references to the required sheets
    Dim wsPortfolio As Worksheet
    Dim wsTotalGraphs As Worksheet
    
    ' You may need to adjust these sheet names based on your workbook structure
    Set wsPortfolio = ThisWorkbook.Sheets("Summary")
    Set wsTotalGraphs = ThisWorkbook.Sheets("TotalBackTest")
    
    ' Delete the current PortfolioGraphs sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("BackTestGraphs").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Call your graph generation routine
    Call CreatePortfolioGraphs(wsTotalGraphs, "BackTestGraphs", wsPortfolio)
    
   
    ' Activate the new PortfolioGraphs sheet
    ThisWorkbook.Sheets("BackTestGraphs").Activate
    
    Application.ScreenUpdating = True
End Sub

   

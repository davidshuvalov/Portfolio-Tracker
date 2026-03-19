Attribute VB_Name = "J_Portfolio_Setup"


Sub CreatePortfolioSummary()
    Dim wsSummary As Worksheet
    Dim wsAverageTrueRange As Worksheet
    Dim wsDailyM2M As Worksheet
    Dim wsClosedTrade As Worksheet
    Dim wsPortfolio As Worksheet
    Dim wsInMarketShort As Worksheet
    Dim wsInMarketLong As Worksheet
    Dim wsPortfolioDailyM2M As Worksheet
    Dim wsPortInMarketShort As Worksheet
    Dim wsPortInMarketLong As Worksheet
    Dim wsPortClosedTrade As Worksheet
    Dim wsTotalPortfolioM2M As Worksheet
    Dim wsPortfolioGraphs As Worksheet
    Dim wsStrategies As Worksheet
    Dim contractsTable As ListObject
    Dim lastRow As Long
    Dim contractCount As Double
    Dim i As Long, x As Long
    Dim portfolioRow As Long
    Dim strategyColumn As Long
    Dim margin As Double
    Dim profitRangeData As Variant, dateRangeData As Variant, shortRangeData As Variant, longRangeData As Variant
    Dim lastRowStrategiesNew As Long
    Dim lastATRRow As Long
    Dim ATR_Flag As Long
    Dim counter As Long
    Dim LastColEquity As Long
    Dim startdate As Date
    Dim currentdate As Date
    
    
    ' Initialize column constants manually
    
    ATR_Flag = 1
    
    Call InitializeColumnConstantsManually

    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        Exit Sub
    End If

    
    ' Check if "Summary" sheet exists and has data in row 2
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
    
    
    Call RemoveFilter(wsSummary)
    
    
    Call Deletetab("PortfolioGraphs")
    Call Deletetab("PortfolioMC")
    Call Deletetab("SizingGraphs")
    Call Deletetab("SectorTypeGraphs")
    Call Deletetab("ContractMarginTracking")
    Call Deletetab("DrawdownCorrelations")
    Call Deletetab("NegativeCorrelations")
    Call Deletetab("Correlations")
    Call Deletetab("Portfolio")
    Call Deletetab("status Changes")
    Call Deletetab("PortInMarketShort")
    Call Deletetab("PortInMarketLong")
    Call Deletetab("PortfolioDailyM2M")
    Call Deletetab("PortClosedTrade")
    Call Deletetab("TotalPortfolioM2M")
    Call Deletetab("StrategiesOld")
    
    
    ' Check if "Strategies" sheet exists and has data in row 2
    On Error Resume Next
    Set wsStrategies = ThisWorkbook.Sheets("Strategies") ' Adjust this if needed
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsStrategies Is Nothing Then
        MsgBox "Error: 'Strategies' sheet does not exist.", vbExclamation
        Exit Sub
    End If

   On Error Resume Next
    Set wsAverageTrueRange = ThisWorkbook.Sheets("AverageTrueRange") ' Adjust this if needed
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsAverageTrueRange Is Nothing Then
        ATR_Flag = 0
    End If
    
   Dim LiveCheck As Long
    lastRow = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row
    LiveCheck = 0
   For i = 2 To lastRow ' Assuming first row is headers
        If wsSummary.Cells(i, COL_STATUS).value = GetNamedRangeValue("Port_Status") Then
            LiveCheck = 1
        End If
    Next i
    
   If LiveCheck = 0 Then
        MsgBox "No Live strategies selected in the Summary tab"
        
        Exit Sub
    End If
         
    
    Set wsDailyM2M = ThisWorkbook.Sheets("DailyM2MEquity")
    Set wsClosedTrade = ThisWorkbook.Sheets("ClosedTradePNL")
    Set wsInMarketShort = ThisWorkbook.Sheets("InMarketShort") ' Adjust this if needed
    Set wsInMarketLong = ThisWorkbook.Sheets("InMarketLong") ' Adjust this if needed
    
    ' Turn off screen updating
       Application.ScreenUpdating = False
    Application.EnableEvents = False

     ' Set up the Portfolio Summary tab
    Set wsPortfolio = ThisWorkbook.Sheets.Add(After:=wsSummary)
    wsPortfolio.name = "Portfolio"
    wsPortfolio.Tab.Color = RGB(0, 176, 240)
    
    ' Set white background color for the entire worksheet
    wsPortfolio.Cells.Interior.Color = RGB(255, 255, 255)

    ' Set up the PortfolioDailyM2M tab
    Set wsPortfolioDailyM2M = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsPortfolioDailyM2M.name = "PortfolioDailyM2M"
    wsPortfolioDailyM2M.Tab.Color = RGB(117, 219, 255)
    Set wsPortClosedTrade = ThisWorkbook.Sheets.Add(After:=wsPortfolioDailyM2M)
    wsPortClosedTrade.name = "PortClosedTrade"
    wsPortClosedTrade.Tab.Color = RGB(117, 219, 255)
    Set wsTotalPortfolioM2M = ThisWorkbook.Sheets.Add(After:=wsPortClosedTrade)
    wsTotalPortfolioM2M.name = "TotalPortfolioM2M"
    wsTotalPortfolioM2M.Tab.Color = RGB(117, 219, 255)
    Set wsPortInMarketShort = ThisWorkbook.Sheets.Add(After:=wsTotalPortfolioM2M)
    wsPortInMarketShort.name = "PortInMarketShort"
    wsPortInMarketShort.Tab.Color = RGB(117, 219, 255)
    Set wsPortInMarketLong = ThisWorkbook.Sheets.Add(After:=wsPortInMarketShort)
    wsPortInMarketLong.name = "PortInMarketLong"
    wsPortInMarketLong.Tab.Color = RGB(117, 219, 255)
    
    Application.StatusBar = "Macro Running..."
   
   ' Initialize arrays for DailyM2M, InMarketShort, and InMarketLong with dates and strategy names
    lastRow = EndRowByCutoffSimple(wsDailyM2M, 1)
    
    numCols = wsDailyM2M.Cells(1, wsDailyM2M.Columns.count).End(xlToLeft).column
    ReDim combinedProfitArray(1 To lastRow, 1 To numCols)
    ReDim combinedClosedArray(1 To lastRow, 1 To numCols)
    ReDim combinedShortArray(1 To lastRow, 1 To numCols)
    ReDim combinedLongArray(1 To lastRow, 1 To numCols)
    

    ' Fill in the date column for each array and strategy names in the first row
    For i = 2 To lastRow
        combinedProfitArray(i, 1) = wsDailyM2M.Cells(i, 1).value
        combinedClosedArray(i, 1) = wsClosedTrade.Cells(i, 1).value
        combinedShortArray(i, 1) = wsDailyM2M.Cells(i, 1).value
        combinedLongArray(i, 1) = wsDailyM2M.Cells(i, 1).value
    Next i
   
      
    ' Set headers for Portfolio tab
    With wsPortfolio
        .Cells(1, COL_PORT_STRATEGYCOUNT).value = "Strategy Count"
        .Cells(1, COL_PORT_STRATEGY_NUMBER).value = "Strategy Number"
        .Cells(1, COL_PORT_CREATE_DETAILED_TAB).value = "Detailed Summary"
        .Cells(1, COL_PORT_OPEN_CODE_TAB).value = "Code Tab"
        .Cells(1, COL_PORT_CODE_TEXT).value = "Text File"
        .Cells(1, COL_PORT_FOLDER).value = "Folder"
        .Cells(1, COL_PORT_STRATEGY_NAME).value = "Strategy Name"
        .Cells(1, COL_PORT_CONTRACTS).value = "Contracts"
        .Cells(1, COL_PORT_SYMBOL).value = "Symbol"
        .Cells(1, COL_PORT_TIMEFRAME).value = "Bar Size"
        .Cells(1, COL_PORT_SECTOR).value = "Sector"
        .Cells(1, COL_PORT_TYPE).value = "Strategy Type"
        .Cells(1, COL_PORT_HORIZON).value = "Horizon"
        .Cells(1, COL_PORT_LONGSHORT).value = "Long/Short"
        
        .Cells(1, COL_PORT_MARGIN).value = "Margin"
        
        .Cells(1, COL_PORT_CONTRACT_SIZE).value = "Full Contract Size"
        .Cells(1, COL_PORT_EST_CONTRACTS).value = "Est Vol Contract Size"

        .Cells(1, COL_PORT_STATUS).value = "Status"
        .Cells(1, COL_PORT_ELIGIBILITY).value = "Eligibility"
        .Cells(1, COL_PORT_NEXT_OPT_DATE).value = "Next Opt Date"
        .Cells(1, COL_PORT_LAST_OPT_DATE).value = "Last Opt Date"
        .Cells(1, COL_PORT_LAST_DATE_ON_FILE).value = "Last Date On File"
        .Cells(1, COL_PORT_CURRENT_POSITION).value = "Current Position"
  
        .Cells(1, COL_PORT_EXPECTED_ANNUAL_PROFIT).value = "Expected Annual Profit"
        .Cells(1, COL_PORT_ACTUAL_ANNUAL_PROFIT).value = "Actual Annual Return"
        .Cells(1, COL_PORT_NOTIONAL_CAPITAL).value = "Notional Capital"
        .Cells(1, COL_PORT_IS_ANNUAL_SD_IS).value = "Annual Standard Deviation (IS)"
        .Cells(1, COL_PORT_IS_ANNUAL_SD_ISOOS).value = "Annual Standard Deviation (IS + OOS)"
        
        
        .Cells(1, COL_PORT_TRADES_PER_YEAR).value = "Trades Per Year"
        
        .Cells(1, COL_PORT_PERCENT_TIME_IN_MARKET).value = "Percent Time in Market"
        .Cells(1, COL_PORT_AVG_TRADE_LENGTH).value = "Average Trade in Market (days)"
        
        .Cells(1, COL_PORT_AVG_IS_OOS_TRADE).value = "Avg Trade (IS+OOS)"
        
        .Cells(1, COL_PORT_AVG_PROFIT_IS_OOS_TRADE).value = "Avg Profitable Trade (IS+OOS)"
        .Cells(1, COL_PORT_AVG_LOSS_IS_OOS_TRADE).value = "Avg Unprofitable Trade (IS+OOS)"
        .Cells(1, COL_PORT_LARGEST_WIN_IS_OOS_TRADE).value = " Largest Profitable Trade (IS+OOS)"
        .Cells(1, COL_PORT_LARGEST_LOSS_IS_OOS_TRADE).value = "Largest Unprofitable Trade (IS+OOS)"
        
        .Cells(1, COL_PORT_MAX_IS_OOS_DRAWDOWN).value = "Max Drawdown (IS+OOS)"
        .Cells(1, COL_PORT_AVG_IS_OOS_DRAWDOWN).value = "Avg Drawdown (IS+OOS)"
        .Cells(1, COL_PORT_MAX_DRAWDOWN_LAST_12_MONTHS).value = "Max Drawdown (Last 12 Months)"

        
        .Cells(1, COL_PORT_PROFIT_LAST_1_MONTH).value = "Profit Last 1 Month"
        .Cells(1, COL_PORT_PROFIT_LAST_3_MONTHS).value = "Profit Last 3 Months"
        .Cells(1, COL_PORT_PROFIT_LAST_6_MONTHS).value = "Profit Last 6 Months"
        .Cells(1, COL_PORT_PROFIT_LAST_9_MONTHS).value = "Profit Last 9 Months"
        .Cells(1, COL_PORT_PROFIT_LAST_12_MONTHS).value = "Profit Last 12 Months"
        .Cells(1, COL_PORT_PROFIT_SINCE_OOS_START).value = "Profit Since OOS Start"
        .Cells(1, COL_PORT_COUNT_PROFIT_MONTHS).value = wsSummary.Cells(1, COL_COUNT_PROFIT_MONTHS).value

        .Cells(1, COL_PORT_ATR_LAST_1_MONTH).value = "ATR Last 1 Month"
        .Cells(1, COL_PORT_ATR_LAST_3_MONTHS).value = "ATR Last 3 Months"
        .Cells(1, COL_PORT_ATR_LAST_6_MONTHS).value = "ATR Last 6 Months"
        .Cells(1, COL_PORT_ATR_LAST_12_MONTHS).value = "ATR Last 12 Months"
        .Cells(1, COL_PORT_ATR_LAST_24_MONTHS).value = "ATR Last 24 Months"
        .Cells(1, COL_PORT_ATR_LAST_60_MONTHS).value = "ATR Last 60 Months"
        .Cells(1, COL_PORT_ATR_ALL_DATA).value = "ATR All Time"
        
        
    End With

    'inputs for estimated contract sizing
    Dim contractSize As Double, accountSize As Double, Ratio As Double, ATROption As String, PercContPort As Double, marginMutliple As Double
    Dim ATRColumn As Long
    
    
    accountSize = GetNamedRangeValue("PortfolioStartingEquity")
    marginMultiple = GetNamedRangeValue("marginMulti")
    Ratio = GetNamedRangeValue("Contract_Size_Ratio")
    ATROption = GetNamedRangeValue("ATRContractOption")
    PercContPort = GetNamedRangeValue("PercentContractPortfolio")
    
    'get the column of the chosen ATR
    Select Case ATROption
       Case "ATR Last 1 Month"
                                ATRColumn = COL_PORT_ATR_LAST_1_MONTH
       Case "ATR Last 3 Months"
                                ATRColumn = COL_PORT_ATR_LAST_3_MONTHS
       Case "ATR Last 6 Months"
                                ATRColumn = COL_PORT_ATR_LAST_6_MONTHS
       Case "ATR Last 12 Months"
                                ATRColumn = COL_PORT_ATR_LAST_12_MONTHS
       Case "ATR Last 24 Months"
                                ATRColumn = COL_PORT_ATR_LAST_24_MONTHS
       Case "ATR Last 60 Months"
                                ATRColumn = COL_PORT_ATR_LAST_60_MONTHS
       Case "ATR All Time"
                                ATRColumn = COL_PORT_ATR_ALL_DATA
       Case Else
                                ATRColumn = COL_PORT_ATR_LAST_3_MONTHS
    End Select
    

    ' Loop through the live strategies in the Summary sheet
    lastRow = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row
    
    If ATR_Flag = 1 Then
        lastATRRow = wsAverageTrueRange.Cells(wsAverageTrueRange.rows.count, 1).End(xlUp).row
    End If
    
    portfolioRow = 2
    
    Dim Live_Count As Long, z As Long
    Live_Count = 0
    
    For i = 2 To lastRow ' Assuming first row is headers
        If wsSummary.Cells(i, COL_STATUS).value = GetNamedRangeValue("Port_Status") Then Live_Count = Live_Count + 1
    Next i
    
    z = 0
    
    For i = 2 To lastRow ' Assuming first row is headers
        
        
        
        If wsSummary.Cells(i, COL_STATUS).value = GetNamedRangeValue("Port_Status") Then
            z = z + 1
            Application.StatusBar = "Portfolio Sorting: " & Format((z) / (Live_Count), "0%") & " completed"
            
            margin = wsSummary.Cells(i, COL_MARGIN).value
            strategyName = wsSummary.Cells(i, COL_STRATEGY_NAME).value
        
            ' Fill in the strategy name in the first row
            combinedProfitArray(1, portfolioRow) = strategyName
            combinedShortArray(1, portfolioRow) = strategyName
            combinedLongArray(1, portfolioRow) = strategyName
            combinedClosedArray(1, portfolioRow) = strategyName
            
            ' Fill in the portfolio details
            wsPortfolio.Cells(portfolioRow, COL_PORT_STRATEGYCOUNT).value = portfolioRow - 1
            wsPortfolio.Cells(portfolioRow, COL_PORT_STRATEGY_NUMBER).value = wsSummary.Cells(i, COL_STRATEGY_NUMBER).value
            wsPortfolio.Cells(portfolioRow, COL_PORT_STRATEGY_NAME).value = strategyName
            wsPortfolio.Cells(portfolioRow, COL_PORT_SYMBOL).value = wsSummary.Cells(i, COL_SYMBOL).value
            wsPortfolio.Cells(portfolioRow, COL_PORT_TIMEFRAME).value = wsSummary.Cells(i, COL_TIMEFRAME).value
            wsPortfolio.Cells(portfolioRow, COL_PORT_STATUS).value = wsSummary.Cells(i, COL_STATUS).value
            wsPortfolio.Cells(portfolioRow, COL_PORT_ELIGIBILITY).value = wsSummary.Cells(i, COL_ELIGIBILITY).value
                        
            wsPortfolio.Cells(portfolioRow, COL_PORT_SECTOR).value = wsSummary.Cells(i, COL_SECTOR).value
            
            wsPortfolio.Cells(portfolioRow, COL_PORT_NEXT_OPT_DATE).value = wsSummary.Cells(i, COL_NEXT_OPT_DATE).value
            wsPortfolio.Cells(portfolioRow, COL_PORT_LAST_OPT_DATE).value = wsSummary.Cells(i, COL_LAST_OPT_DATE).value
            wsPortfolio.Cells(portfolioRow, COL_PORT_LAST_DATE_ON_FILE).value = wsSummary.Cells(i, COL_LAST_DATE_ON_FILE).value
            wsPortfolio.Cells(portfolioRow, COL_PORT_LONGSHORT).value = wsSummary.Cells(i, COL_LONG_SHORT).value
            
            
            Dim lastRowContract As Long
            lastRowContract = wsStrategies.Cells(wsStrategies.rows.count, 1).End(xlUp).row
            contractCount = 0
            wsPortfolio.Cells(portfolioRow, COL_PORT_CONTRACTS).value = contractCount
            ' Loop through the contracts table to find a match in column B (Strategy Name)
           
           
            For contractRow = 2 To lastRowContract
                If wsStrategies.Cells(contractRow, COL_STRAT_STRATEGY_NAME).value = wsPortfolio.Cells(portfolioRow, COL_PORT_STRATEGY_NAME).value Then ' Check column B for strategy name
                    contractCount = wsStrategies.Cells(contractRow, COL_STRAT_CONTRACTS).value ' Get the number of contracts from column E
                    wsPortfolio.Cells(portfolioRow, COL_PORT_TYPE).value = wsStrategies.Cells(contractRow, COL_STRAT_TYPE).value ' Get the number of contracts from column E
                    wsPortfolio.Cells(portfolioRow, COL_PORT_HORIZON).value = wsStrategies.Cells(contractRow, COL_STRAT_HORIZON).value ' Get the number of contracts from column E
                           wsPortfolio.Cells(portfolioRow, COL_PORT_CONTRACTS).value = contractCount
                    
                    
                    Exit For
                End If
            Next contractRow



            ' Scale metrics by number of contracts
            wsPortfolio.Cells(portfolioRow, COL_PORT_MARGIN).value = margin * contractCount
     
            

            
            
            wsPortfolio.Cells(portfolioRow, COL_PORT_EXPECTED_ANNUAL_PROFIT).value = wsSummary.Cells(i, COL_EXPECTED_ANNUAL_PROFIT).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_ACTUAL_ANNUAL_PROFIT).value = wsSummary.Cells(i, COL_ACTUAL_ANNUAL_PROFIT).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_NOTIONAL_CAPITAL).value = wsSummary.Cells(i, COL_NOTIONAL_CAPITAL).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_IS_ANNUAL_SD_IS).value = wsSummary.Cells(i, COL_IS_ANNUAL_SD_IS).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_IS_ANNUAL_SD_ISOOS).value = wsSummary.Cells(i, COL_IS_ANNUAL_SD_ISOOS).value * contractCount
            

            
            wsPortfolio.Cells(portfolioRow, COL_PORT_TRADES_PER_YEAR).value = wsSummary.Cells(i, COL_TRADES_PER_YEAR).value
            wsPortfolio.Cells(portfolioRow, COL_PORT_PERCENT_TIME_IN_MARKET).value = wsSummary.Cells(i, COL_PERCENT_TIME_IN_MARKET).value
            wsPortfolio.Cells(portfolioRow, COL_PORT_AVG_TRADE_LENGTH).value = wsSummary.Cells(i, COL_AVG_TRADE_LENGTH).value
            
            
            wsPortfolio.Cells(portfolioRow, COL_PORT_AVG_IS_OOS_TRADE).value = wsSummary.Cells(i, COL_AVG_IS_OOS_TRADE).value * contractCount
            
            wsPortfolio.Cells(portfolioRow, COL_PORT_AVG_PROFIT_IS_OOS_TRADE).value = wsSummary.Cells(i, COL_AVG_PROFIT_IS_OOS_TRADE).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_AVG_LOSS_IS_OOS_TRADE).value = wsSummary.Cells(i, COL_AVG_LOSS_IS_OOS_TRADE).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_LARGEST_WIN_IS_OOS_TRADE).value = wsSummary.Cells(i, COL_LARGEST_WIN_IS_OOS_TRADE).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_LARGEST_LOSS_IS_OOS_TRADE).value = wsSummary.Cells(i, COL_LARGEST_LOSS_IS_OOS_TRADE).value * contractCount
            
            
            wsPortfolio.Cells(portfolioRow, COL_PORT_MAX_IS_OOS_DRAWDOWN).value = wsSummary.Cells(i, COL_WORST_IS_OOS_DRAWDOWN).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_AVG_IS_OOS_DRAWDOWN).value = wsSummary.Cells(i, COL_AVG_IS_OOS_DRAWDOWN).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_MAX_DRAWDOWN_LAST_12_MONTHS).value = wsSummary.Cells(i, COL_MAX_DRAWDOWN_LAST_12_MONTHS).value * contractCount
               
           ' wsPortfolio.Cells(portfolioRow, COL_PORT_MAX_DRAWDOWN_PERCENT).value = wsSummary.Cells(i, COL_MAX_DRAWDOWN_PERCENT).value
                
        
            wsPortfolio.Cells(portfolioRow, COL_PORT_PROFIT_LAST_1_MONTH).value = wsSummary.Cells(i, COL_PROFIT_LAST_1_MONTH).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_PROFIT_LAST_3_MONTHS).value = wsSummary.Cells(i, COL_PROFIT_LAST_3_MONTHS).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_PROFIT_LAST_6_MONTHS).value = wsSummary.Cells(i, COL_PROFIT_LAST_6_MONTHS).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_PROFIT_LAST_9_MONTHS).value = wsSummary.Cells(i, COL_PROFIT_LAST_9_MONTHS).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_PROFIT_LAST_12_MONTHS).value = wsSummary.Cells(i, COL_PROFIT_LAST_12_MONTHS).value * contractCount
            wsPortfolio.Cells(portfolioRow, COL_PORT_PROFIT_SINCE_OOS_START).value = wsSummary.Cells(i, COL_PROFIT_SINCE_OOS_START).value * contractCount
            
     
            wsPortfolio.Cells(portfolioRow, COL_PORT_COUNT_PROFIT_MONTHS).value = wsSummary.Cells(i, COL_COUNT_PROFIT_MONTHS).value
            
            wsPortfolio.Cells(portfolioRow, COL_PORT_CURRENT_POSITION).value = wsSummary.Cells(i, COL_CURRENT_POSITION).value * contractCount
     
            
            If ATR_Flag = 1 Then
                For x = 2 To lastATRRow
                    If wsPortfolio.Cells(portfolioRow, COL_PORT_SYMBOL).value = wsAverageTrueRange.Cells(x, 1).value Then
                    
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_1_MONTH).value = wsAverageTrueRange.Cells(x, 2).value * contractCount
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_3_MONTHS).value = wsAverageTrueRange.Cells(x, 3).value * contractCount
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_6_MONTHS).value = wsAverageTrueRange.Cells(x, 4).value * contractCount
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_12_MONTHS).value = wsAverageTrueRange.Cells(x, 5).value * contractCount
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_24_MONTHS).value = wsAverageTrueRange.Cells(x, 6).value * contractCount
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_60_MONTHS).value = wsAverageTrueRange.Cells(x, 7).value * contractCount
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_ALL_DATA).value = wsAverageTrueRange.Cells(x, 8).value * contractCount
                    End If
                Next x
            Else
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_1_MONTH).value = 0
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_3_MONTHS).value = 0
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_6_MONTHS).value = 0
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_12_MONTHS).value = 0
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_24_MONTHS).value = 0
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_LAST_60_MONTHS).value = 0
                      wsPortfolio.Cells(portfolioRow, COL_PORT_ATR_ALL_DATA).value = 0
            End If
  
            
            LastColEquity = wsDailyM2M.Cells(1, wsDailyM2M.Columns.count).End(xlToLeft).column
    
        
            If ATR_Flag = 1 And contractCount > 0 Then
                contractSize = margin * marginMultiple * Ratio + (wsPortfolio.Cells(portfolioRow, ATRColumn).value / contractCount) * (1 - Ratio)
            Else
                contractSize = margin * marginMultiple * Ratio
            End If
            
                   
            
            wsPortfolio.Cells(portfolioRow, COL_PORT_CONTRACT_SIZE).value = contractSize
            
            If contractSize > 0 And accountSize > 0 Then
                wsPortfolio.Cells(portfolioRow, COL_PORT_EST_CONTRACTS).value = Round(PercContPort / (contractSize / accountSize), 1)
            Else
                wsPortfolio.Cells(portfolioRow, COL_PORT_EST_CONTRACTS).value = ""
            End If
            
            
            
            strategyColumn = -99
            
            On Error Resume Next
            For counter = 2 To LastColEquity
                      If strategyName = wsDailyM2M.Cells(1, counter).value Then strategyColumn = counter
            Next counter
            On Error GoTo 0

            ' Exit and show error if the sheet doesn't exist
            If strategyColumn = -99 Then
        
                Application.ScreenUpdating = True
                Application.EnableEvents = Truera
                Application.StatusBar = False
                
                Call GoToControl
                MsgBox "Error: Cannot match " & strategyName & " in the tab 'DailyM2MEquity'", vbExclamation
                Exit Sub
            End If
            
            If Not IsError(strategyColumn) Then
                ' Define lastRow for current strategyColumn to capture full data
                lastRowDaily = EndRowByCutoffSimple(wsDailyM2M, 1)
            
                ' Load ranges into arrays
                profitRangeData = wsDailyM2M.Range(wsDailyM2M.Cells(2, strategyColumn), wsDailyM2M.Cells(lastRowDaily, strategyColumn)).value
                ClosedRangeData = wsClosedTrade.Range(wsClosedTrade.Cells(2, strategyColumn), wsClosedTrade.Cells(lastRowDaily, strategyColumn)).value
                shortRangeData = wsInMarketShort.Range(wsInMarketShort.Cells(2, strategyColumn), wsInMarketShort.Cells(lastRowDaily, strategyColumn)).value
                longRangeData = wsInMarketLong.Range(wsInMarketLong.Cells(2, strategyColumn), wsInMarketLong.Cells(lastRowDaily, strategyColumn)).value
                dateRangeData = wsDailyM2M.Range(wsDailyM2M.Cells(2, 1), wsDailyM2M.Cells(lastRowDaily, 1)).value
            
                Dim wsContractMultiples As Worksheet
                Dim contractMultipleRow As Long
                Dim contractMultipleColumn As Long
                Dim row As Long
                Dim col As Long
                Dim yearFound As Boolean
                Dim rowFound As Boolean
                Dim currentYear As Long
                Dim lastProcessedYear As Long
                Dim contractMultiple As Double
                Dim reweightATR As Boolean
                
                ' Check if reweight_PORT_ATR is "Yes"
                reweightATR = (GetNamedRangeValue("reweight_PORT_ATR") = "yes" Or GetNamedRangeValue("reweight_PORT_ATR") = "Yes")
                reweightindexATR = (GetNamedRangeValue("reweight_index_only") = "yes" Or GetNamedRangeValue("reweight_index_only") = "Yes")
                
                If reweightindexATR And wsPortfolio.Cells(portfolioRow, COL_PORT_SECTOR).value <> "Index" Then reweightATR = False
                                
                ' Check and set the ContractMultiples sheet if reweighting is required
                If reweightATR Then
                    On Error Resume Next
                    Set wsContractMultiples = ThisWorkbook.Sheets("ContractMultiples")
                    On Error GoTo 0
                
                    If wsContractMultiples Is Nothing Then
                    
                        Application.ScreenUpdating = True
                        Application.EnableEvents = True
                        Application.StatusBar = False
                        
                        Call GoToControl
                        
                        
                        MsgBox "Error: 'ContractMultiples' sheet does not exist. If the 'Buy and Hold' strategies are not set up, make sure the re-weight portfolio contracts on historical ATR is set to *No* in the inputs tab", vbExclamation
                        Exit Sub
                    End If
                
                    ' Find the strategy symbol row in the ContractMultiples sheet
                    rowFound = False
                    For row = 2 To wsContractMultiples.Cells(wsContractMultiples.rows.count, 1).End(xlUp).row
                        If wsContractMultiples.Cells(row, 1).value = wsPortfolio.Cells(portfolioRow, COL_PORT_SYMBOL).value Then
                            contractMultipleRow = row
                            rowFound = True
                            Exit For
                        End If
                    Next row
                
                    If Not rowFound Then
        
                        Application.ScreenUpdating = True
                        Application.EnableEvents = True
                        Application.StatusBar = False
                        
                        Call GoToControl
                        
                        MsgBox "Strategy symbol not found in 'ContractMultiples' sheet: " & wsPortfolio.Cells(portfolioRow, COL_PORT_SYMBOL).value, vbExclamation
                        Exit Sub
                    End If
                End If
                
                ' Loop through the data
                lastProcessedYear = -1 ' Initialize to an invalid year
                For j = 1 To UBound(profitRangeData, 1)
                    ' Validate date before processing
                    If Not IsEmpty(dateRangeData(j, 1)) And IsDate(dateRangeData(j, 1)) Then
                        currentYear = Year(dateRangeData(j, 1))
                
                        ' Only calculate the contract multiple if the year has changed
                        If reweightATR And currentYear <> lastProcessedYear Then
                            yearFound = False
                
                            ' Search for the year in the first row of the ContractMultiples sheet
                            For col = 2 To wsContractMultiples.Cells(1, wsContractMultiples.Columns.count).End(xlToLeft).column
                                If wsContractMultiples.Cells(1, col).value = currentYear Then
                                    contractMultipleColumn = col
                                    yearFound = True
                                    Exit For
                                End If
                            Next col
                
                            ' Retrieve contract multiple or default to 1 if the year is not found
                            If yearFound Then
                                contractMultiple = wsContractMultiples.Cells(contractMultipleRow, contractMultipleColumn).value
                                If IsEmpty(contractMultiple) Then
                                    Debug.Print "Empty contract multiple for Year: " & currentYear & " in column " & contractMultipleColumn
                                    contractMultiple = 1 ' Default to 1 if empty
                                End If
                            Else
                                
                                contractMultiple = 1 ' Default to 1 if the year is not found
                            End If
                
                            ' Update last processed year
                            lastProcessedYear = currentYear
                        End If
                    Else
                        ' Handle invalid or missing date
                        
                        contractMultiple = 1 ' Default value
                    End If
                    
                    If reweightATR = False Then contractMultiple = 1
                
                    ' Populate combined arrays with scaled values
                    combinedProfitArray(j + 1, portfolioRow) = profitRangeData(j, 1) * contractCount * contractMultiple
                    combinedClosedArray(j + 1, portfolioRow) = ClosedRangeData(j, 1) * contractCount * contractMultiple
                    combinedShortArray(j + 1, portfolioRow) = shortRangeData(j, 1) * contractCount
                    combinedLongArray(j + 1, portfolioRow) = longRangeData(j, 1) * contractCount
                Next j
            End If
            
            
            ' Increment portfolio row
            portfolioRow = portfolioRow + 1
        End If
        
        
    Next i
    
    
    
    Application.StatusBar = "Combining PortfolioDailyM2M in arrays..."
    ' Write combined arrays to sheets in one go
    wsPortfolioDailyM2M.Range("A1").Resize(UBound(combinedProfitArray, 1), UBound(combinedProfitArray, 2)).value = combinedProfitArray
    
    
    Application.StatusBar = "Combining PortInMarketShort in arrays..."
    wsPortInMarketShort.Range("A1").Resize(UBound(combinedShortArray, 1), UBound(combinedShortArray, 2)).value = combinedShortArray
    
     Application.StatusBar = "Combining PortInMarketLong in arrays..."
    wsPortInMarketLong.Range("A1").Resize(UBound(combinedLongArray, 1), UBound(combinedLongArray, 2)).value = combinedLongArray
    
       Application.StatusBar = "Combining PortClosedTrade in arrays..."
    wsPortClosedTrade.Range("A1").Resize(UBound(combinedClosedArray, 1), UBound(combinedClosedArray, 2)).value = combinedClosedArray
    
    yearsToConsider = Range("PortfolioPeriod").value
    currentdate = wsPortfolio.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    startdate = DateAdd("yyyy", -Int(yearsToConsider), currentdate)
    startdate = DateAdd("m", -(yearsToConsider - Int(yearsToConsider)) * 12, startdate)
    
    
    ' Optional: Format the PortfolioDailyM2MEquity tab
    With wsPortfolioDailyM2M
        .Cells(1, 1).value = "Date"
    End With
    
    With wsPortInMarketShort
        .Cells(1, 1).value = "Date"
    End With
    
    With wsPortInMarketLong
        .Cells(1, 1).value = "Date"
    End With
    
    With wsPortClosedTrade
        .Cells(1, 1).value = "Date"
    End With
    
    Application.StatusBar = "Writing to TotalPortfolioM2M..."
    
      ' Summing up daily M2M for each day across strategies and writing to TotalPortfolioM2M
    Dim lastRowDailyM2M As Long, lastColPortfolioM2M As Long
    lastRowDailyM2M = wsPortfolioDailyM2M.Cells(wsPortfolioDailyM2M.rows.count, 1).End(xlUp).row
    lastColPortfolioM2M = wsPortfolioDailyM2M.Cells(1, wsPortfolioDailyM2M.Columns.count).End(xlToLeft).column

    ' Set headers in TotalPortfolioM2M
    wsTotalPortfolioM2M.Cells(1, 1).value = "Date"
    wsTotalPortfolioM2M.Cells(1, 2).value = "Total Daily Profit"
    wsTotalPortfolioM2M.Cells(1, 3).value = "Total Cumulative P/L"
    wsTotalPortfolioM2M.Cells(1, 4).value = "Total Drawdown"
    wsTotalPortfolioM2M.Cells(1, 5).value = "Year"
    wsTotalPortfolioM2M.Cells(1, 6).value = "Month"
    wsTotalPortfolioM2M.Cells(1, 7).value = "Week"
    wsTotalPortfolioM2M.Cells(1, 8).value = "Drawdown Percent"
    
    Dim peakProfit As Double, currentEquity As Double, currentDrawdown As Double, drawdownpercent As Double, startingEquity As Double
    
    
    ' === Initialize Symbol and Sector Dictionaries ===
    Dim dictSymbolProfits As Object, dictSectorProfits As Object
    Set dictSymbolProfits = CreateObject("Scripting.Dictionary")
    Set dictSectorProfits = CreateObject("Scripting.Dictionary")
    
    Dim dictStrategyDetails As Object
    Set dictStrategyDetails = CreateObject("Scripting.Dictionary")
    
    ' Collect strategy details from Portfolio tab (Symbol & Sector)
    For i = 2 To wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row
        Dim portStrategy As String
        portStrategy = wsPortfolio.Cells(i, COL_PORT_STRATEGY_NAME).value
        
        If Not dictStrategyDetails.Exists(portStrategy) Then
            dictStrategyDetails.Add portStrategy, Array(wsPortfolio.Cells(i, COL_PORT_SYMBOL).value, wsPortfolio.Cells(i, COL_PORT_SECTOR).value)
        End If
    Next i
    
     ' Find last row and column in PortfolioDailyM2M
    lastRowDailyM2M = wsPortfolioDailyM2M.Cells(wsPortfolioDailyM2M.rows.count, 1).End(xlUp).row
    lastColPortfolioM2M = wsPortfolioDailyM2M.Cells(1, wsPortfolioDailyM2M.Columns.count).End(xlToLeft).column

    ' Initialize drawdown and equity tracking
    peakProfit = 0
    currentDrawdown = 0
    currentEquity = 0
    drawdownpercent = 0
    startingEquity = GetNamedRangeValue("PortfolioStartingEquity")
    ' Loop through each row and only process dates within the specified range
    outputRow = 2 ' Start output on the second row in TotalPortfolioM2M

    ' Loop through each row and only process dates within the specified range
    Dim dailyTotal As Double
    For row = 2 To lastRowDailyM2M
        ' Get the date of the current row
        Dim currentDateInData As Date
        
       
        
        currentDateInData = wsPortfolioDailyM2M.Cells(row, 1).value

        ' Check if the date is within the start and end date range
        If currentDateInData >= startdate And currentDateInData <= currentdate Then
        
            Application.StatusBar = "Portfolio Results Calculating: " & Format((currentDateInData - startdate) / (currentdate - startdate), "0%") & " completed"
        
        
            ' Copy the date to TotalPortfolioM2M
            wsTotalPortfolioM2M.Cells(outputRow, 1).value = currentDateInData

            ' Calculate daily total profit across all strategies
            dailyTotal = 0
            For col = 2 To lastColPortfolioM2M
                dailyTotal = dailyTotal + wsPortfolioDailyM2M.Cells(row, col).value
            Next col
            
            ' Write total daily profit and calculate cumulative P/L
            wsTotalPortfolioM2M.Cells(outputRow, 2).value = dailyTotal
            If outputRow > 2 Then
                wsTotalPortfolioM2M.Cells(outputRow, 3).value = dailyTotal + wsTotalPortfolioM2M.Cells(outputRow - 1, 3).value
            Else
                wsTotalPortfolioM2M.Cells(outputRow, 3).value = dailyTotal
            End If
            currentEquity = wsTotalPortfolioM2M.Cells(outputRow, 3).value

            ' Update peak profit if current equity is higher
            If currentEquity > peakProfit Then
                peakProfit = currentEquity
            End If
            

            ' Calculate drawdown as the difference between peak and current equity
            currentDrawdown = peakProfit - currentEquity
            drawdownpercent = currentDrawdown / (startingEquity + peakProfit + 0.000001)
            
            wsTotalPortfolioM2M.Cells(outputRow, 4).value = currentDrawdown

            ' Add the year to the Year column
            wsTotalPortfolioM2M.Cells(outputRow, 5).value = Year(currentDateInData)
            wsTotalPortfolioM2M.Cells(outputRow, 6).value = Month(currentDateInData)
            wsTotalPortfolioM2M.Cells(outputRow, 7).value = Application.WorksheetFunction.WeekNum(currentDateInData, vbSunday)
            wsTotalPortfolioM2M.Cells(outputRow, 8).value = drawdownpercent
            
            
            ' === Track Symbol and Sector Profits (Only from Portfolio) ===
            For col = 2 To lastColPortfolioM2M
                Dim colStrategy As String, dailyProfit As Double
                dailyProfit = wsPortfolioDailyM2M.Cells(row, col).value
                colStrategy = wsPortfolioDailyM2M.Cells(1, col).value ' Strategy name from header
                
                If dailyProfit <> 0 And dictStrategyDetails.Exists(colStrategy) Then
                    Dim sym As String, sec As String
                    sym = dictStrategyDetails(colStrategy)(0)
                    sec = dictStrategyDetails(colStrategy)(1)
                    
                    ' Accumulate Symbol Profits
                    If Not dictSymbolProfits.Exists(sym) Then dictSymbolProfits(sym) = 0
                    dictSymbolProfits(sym) = dictSymbolProfits(sym) + dailyProfit
                    
                    ' Accumulate Sector Profits
                    If Not dictSectorProfits.Exists(sec) Then dictSectorProfits(sec) = 0
                    dictSectorProfits(sec) = dictSectorProfits(sec) + dailyProfit
                End If
            Next col
            
            ' Increment the outputRow for the next eligible row
            outputRow = outputRow + 1
        End If
        
        
    
    Next row

    ' === Output Symbol and Sector Profits to TotalPortfolioM2M ===
    Dim symbolStartCol As Long, sectorStartCol As Long
    Dim outputSymbolRow As Long, outputSectorRow As Long
    
    symbolStartCol = 19 ' Column S
    sectorStartCol = symbolStartCol + 3 ' Start sector output 3 columns after symbols
    outputSymbolRow = 1
    outputSectorRow = 1
    
    ' === Output Symbol Profits ===
    wsTotalPortfolioM2M.Cells(outputSymbolRow, symbolStartCol).value = "Symbol Profits (Portfolio Only)"
    wsTotalPortfolioM2M.Cells(outputSymbolRow, symbolStartCol).Font.Bold = True
    outputSymbolRow = outputSymbolRow + 1
    
    Dim symbol As Variant
    For Each symbol In dictSymbolProfits.keys
        wsTotalPortfolioM2M.Cells(outputSymbolRow, symbolStartCol).value = symbol
        wsTotalPortfolioM2M.Cells(outputSymbolRow, symbolStartCol + 1).value = dictSymbolProfits(symbol)
        wsTotalPortfolioM2M.Cells(outputSymbolRow, symbolStartCol + 1).NumberFormat = "$#,##0.00"
        outputSymbolRow = outputSymbolRow + 1
    Next symbol
    
    ' === Output Sector Profits ===
    wsTotalPortfolioM2M.Cells(outputSectorRow, sectorStartCol).value = "Sector Profits (Portfolio Only)"
    wsTotalPortfolioM2M.Cells(outputSectorRow, sectorStartCol).Font.Bold = True
    outputSectorRow = outputSectorRow + 1
    
    Dim sector As Variant
    For Each sector In dictSectorProfits.keys
        wsTotalPortfolioM2M.Cells(outputSectorRow, sectorStartCol).value = sector
        wsTotalPortfolioM2M.Cells(outputSectorRow, sectorStartCol + 1).value = dictSectorProfits(sector)
        wsTotalPortfolioM2M.Cells(outputSectorRow, sectorStartCol + 1).NumberFormat = "$#,##0.00"
        outputSectorRow = outputSectorRow + 1
    Next sector
    
    ' Auto-fit for neat display
    With wsTotalPortfolioM2M
        .Columns(symbolStartCol).AutoFit
        .Columns(symbolStartCol + 1).AutoFit
        .Columns(sectorStartCol).AutoFit
        .Columns(sectorStartCol + 1).AutoFit
    End With


    ' Add benchmark data if enabled
    Application.StatusBar = "Adding benchmark data..."
    Call CalculateBenchmarkData(wsTotalPortfolioM2M, wsPortfolioDailyM2M, startdate, currentdate)
    
    ThisWorkbook.Sheets("PortfolioDailyM2M").Visible = xlSheetHidden
    ThisWorkbook.Sheets("TotalPortfolioM2M").Visible = xlSheetHidden
    ThisWorkbook.Sheets("PortClosedTrade").Visible = xlSheetHidden
    ThisWorkbook.Sheets("PortInMarketShort").Visible = xlSheetHidden
    ThisWorkbook.Sheets("PortInMarketLong").Visible = xlSheetHidden
    
    Application.StatusBar = "Danger: Portfolio Probabilities Approaching 1..."
    Call FormatPortfolioTable
    
    Application.StatusBar = "Buttons..."
    Call SetupPortfolioButtonsAndStrategyTabCreation
    
    Call SetupPortfolioButtonsforCodeTabCreation
    
    
    
    If GetNamedRangeValue("open_on_portfolio") = "Yes" Then
        Application.StatusBar = "Sizing Graphs..."
        Call CreateSectorTypeGraphs
        Call CreateSizingGraphs
        
    End If
    
    
    Application.StatusBar = "Portfolio Graphs..."
    Call CreatePortfolioGraphs(wsTotalPortfolioM2M, "PortfolioGraphs", wsPortfolio)
   
    Application.StatusBar = "Portfolio Metrics..."
    Set wsPortfolioGraphs = ThisWorkbook.Sheets("PortfolioGraphs")
    Dim FindMaxDrawdownRow As Long
    lastRow = wsPortfolioGraphs.Cells(wsPortfolioGraphs.rows.count, "B").End(xlUp).row
    
    ' Loop through each cell in column B
    For i = 1 To lastRow
        If wsPortfolioGraphs.Cells(i, 2).value = "Maximum Drawdown (%)" Then
            FindMaxDrawdownRow = i
            Exit For
        End If
    Next i
    
    
    Call PortMetrics(FindMaxDrawdownRow + 2)
    

    'Call OrderVisibleTabsBasedOnList
    
    wsPortfolio.Cells(1, COL_PORT_STRATEGY_NAME).VerticalAlignment = xlBottom
    
    Application.StatusBar = "More Buttons!!!..."
    
    Call CreateSummaryButtons(wsPortfolio, COL_PORT_STRATEGY_NAME, "Portfolio")
        

    ' Find the last row and last column in the summary sheet
    lastRow = wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row

    
    ' Apply AutoFilter to the table
    wsPortfolio.Range(wsPortfolio.Cells(1, 1), wsPortfolio.Cells(portfolioRow - 1, COL_PORT_ATR_ALL_DATA)).AutoFilter
     wsSummary.Range(wsSummary.Cells(1, 1), wsSummary.Cells(lastRow, COL_SHARPE_ISOOS)).AutoFilter
        
     ' Turn on screen updating
       Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "Portfolio Summary and Portfolio Graph tabs created successfully!"
End Sub




Sub FormatPortfolioTable()
    Dim wsPortfolio As Worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")

    Dim lastRow As Long, lastCol As Long
    Dim ExcelDateFormat As String
    
    ExcelDateFormat = GetNamedRangeValue("DateFormat")
    
    
    lastRow = wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row
    lastCol = wsPortfolio.Cells(1, wsPortfolio.Columns.count).End(xlToLeft).column

    ' Store the currently active sheet
    Dim currentSheet As Worksheet
    Set currentSheet = activeSheet
    
    ' Switch to wsSummary, freeze panes, and switch back
     Dim freezecol As String
    freezecol = GetNamedRangeValue("FreezePanesColumn2")
    
    wsPortfolio.Activate
    wsPortfolio.Range(freezecol & "2").Select
    ActiveWindow.FreezePanes = True
    
    ' Apply formatting to the header row
    With wsPortfolio.rows(1)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True ' Enable wrap text for the header row
        .Interior.Color = RGB(220, 230, 241) ' Light blue header background
    End With
    
    ' AutoFit columns for better readability
    wsPortfolio.Columns("A:BZ").AutoFit
    
    ' Apply borders to the table
    With wsPortfolio.Range(wsPortfolio.Cells(1, 1), wsPortfolio.Cells(lastRow, lastCol))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With

    ' Number formatting for certain columns (dollars and percentages)
    Dim currentdate As Date
    currentdate = Date
    Dim ATRThreshold As Double
    
    ATRThreshold = GetNamedRangeValue("ATRCheck")
    
    For i = 2 To lastRow
        Dim nextOptDate As Variant
        Dim lastOptDate As Variant
        Dim daysDiff As Long
        
        ' Check Next Option Date
        nextOptDate = wsPortfolio.Cells(i, COL_PORT_NEXT_OPT_DATE).value
        If IsDate(nextOptDate) Then
            daysDiff = Abs(DateDiff("d", nextOptDate, currentdate))
            If daysDiff <= 7 Then
                wsPortfolio.Cells(i, COL_PORT_NEXT_OPT_DATE).Interior.Color = RGB(255, 0, 0) ' Red
            ElseIf daysDiff <= 15 Then
                wsPortfolio.Cells(i, COL_PORT_NEXT_OPT_DATE).Interior.Color = RGB(255, 255, 0) ' Yellow
            End If
        End If
        
        ' Check Last Option Date
        lastOptDate = wsPortfolio.Cells(i, COL_PORT_LAST_OPT_DATE).value
        If IsDate(lastOptDate) Then
            daysDiff = Abs(DateDiff("d", lastOptDate, currentdate))
            If daysDiff <= 7 Then
                wsPortfolio.Cells(i, COL_PORT_LAST_OPT_DATE).Interior.Color = RGB(255, 0, 0) ' Red
            ElseIf daysDiff <= 15 Then
                wsPortfolio.Cells(i, COL_PORT_LAST_OPT_DATE).Interior.Color = RGB(255, 255, 0) ' Yellow
            End If
        End If
    
    
        If wsPortfolio.Cells(i, COL_PORT_EST_CONTRACTS).value < wsPortfolio.Cells(i, COL_PORT_CONTRACTS).value Then
                wsPortfolio.Cells(i, COL_PORT_EST_CONTRACTS).Interior.Color = RGB(255, 255, 0) ' Red
        End If
    
        If wsPortfolio.Cells(i, COL_PORT_ATR_LAST_1_MONTH).value > (1 + ATRThreshold) * wsPortfolio.Cells(i, COL_PORT_ATR_LAST_3_MONTHS) _
        Or wsPortfolio.Cells(i, COL_PORT_ATR_LAST_1_MONTH).value > (1 + ATRThreshold) * wsPortfolio.Cells(i, COL_PORT_ATR_LAST_6_MONTHS) Then
        
        wsPortfolio.Cells(i, COL_PORT_EST_CONTRACTS).Interior.Color = RGB(255, 0, 0) ' Yellow
        wsPortfolio.Cells(i, COL_PORT_ATR_LAST_1_MONTH).Interior.Color = RGB(255, 0, 0) ' Yellow
        End If
        
        
        
        
    Next i
    
    
    ' Conditional formatting for "Eligibility"
    Dim EligibilityRange As Range
    Set EligibilityRange = wsPortfolio.Range(wsPortfolio.Cells(2, COL_PORT_ELIGIBILITY), wsPortfolio.Cells(lastRow, COL_PORT_ELIGIBILITY))
    With EligibilityRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""No""")
        .Interior.Color = RGB(255, 199, 206) ' Red for Quit
        .Font.Bold = True
    End With
    With EligibilityRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Yes""")
        .Interior.Color = RGB(198, 239, 206) ' Green for Continue
        .Font.Bold = True
    End With

    
    wsPortfolio.Columns(COL_PORT_CONTRACTS).NumberFormat = "0.0"
    wsPortfolio.Columns(COL_PORT_MARGIN).NumberFormat = "$#,##0"
    
    
    wsPortfolio.Columns(COL_PORT_CONTRACT_SIZE).NumberFormat = "$#,##0"
    
     wsPortfolio.Columns(COL_PORT_EST_CONTRACTS).NumberFormat = "0.0"
    
    
    wsPortfolio.Columns(COL_PORT_EXPECTED_ANNUAL_PROFIT).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_ACTUAL_ANNUAL_PROFIT).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_NOTIONAL_CAPITAL).NumberFormat = "$#,##0"
    
    wsPortfolio.Columns(COL_PORT_IS_ANNUAL_SD_IS).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_IS_ANNUAL_SD_ISOOS).NumberFormat = "$#,##0"
    
   
    wsPortfolio.Columns(COL_PORT_TRADES_PER_YEAR).NumberFormat = "0;-0;0"
    
    wsPortfolio.Columns(COL_PORT_PERCENT_TIME_IN_MARKET).NumberFormat = "0%"
    wsPortfolio.Columns(COL_PORT_AVG_TRADE_LENGTH).NumberFormat = "0.00"
    


    wsPortfolio.Columns(COL_PORT_AVG_IS_OOS_TRADE).NumberFormat = "$#,##0"
    
    wsPortfolio.Columns(COL_PORT_AVG_PROFIT_IS_OOS_TRADE).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_AVG_LOSS_IS_OOS_TRADE).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_LARGEST_WIN_IS_OOS_TRADE).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_LARGEST_LOSS_IS_OOS_TRADE).NumberFormat = "$#,##0"
    
    wsPortfolio.Columns(COL_PORT_MAX_IS_OOS_DRAWDOWN).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_AVG_IS_OOS_DRAWDOWN).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_MAX_DRAWDOWN_LAST_12_MONTHS).NumberFormat = "$#,##0"
    
  '  wsPortfolio.Columns(COL_PORT_MAX_DRAWDOWN_PERCENT).NumberFormat = "0%"
    
   
    wsPortfolio.Columns(COL_PORT_PROFIT_LAST_1_MONTH).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_PROFIT_LAST_3_MONTHS).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_PROFIT_LAST_6_MONTHS).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_PROFIT_LAST_9_MONTHS).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_PROFIT_LAST_12_MONTHS).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_PROFIT_SINCE_OOS_START).NumberFormat = "$#,##0"
    
    
    wsPortfolio.Columns(COL_PORT_COUNT_PROFIT_MONTHS).NumberFormat = "0;-0;0"

    
    wsPortfolio.Columns(COL_PORT_ATR_LAST_1_MONTH).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_ATR_LAST_3_MONTHS).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_ATR_LAST_6_MONTHS).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_ATR_LAST_12_MONTHS).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_ATR_LAST_24_MONTHS).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_ATR_LAST_60_MONTHS).NumberFormat = "$#,##0"
    wsPortfolio.Columns(COL_PORT_ATR_ALL_DATA).NumberFormat = "$#,##0"
    ' Date formatting for date columns
    
    If ExcelDateFormat <> "US" Then
        wsPortfolio.Columns(COL_PORT_NEXT_OPT_DATE).NumberFormat = "dd/mm/yyyy" ' Next Opt Date
        wsPortfolio.Columns(COL_PORT_LAST_OPT_DATE).NumberFormat = "dd/mm/yyyy" ' Last Opt Date
        wsPortfolio.Columns(COL_PORT_LAST_DATE_ON_FILE).NumberFormat = "dd/mm/yyyy" ' OOS Begin Date
    Else
        wsPortfolio.Columns(COL_PORT_NEXT_OPT_DATE).NumberFormat = "mm/dd/yyyy" ' Next Opt Date
        wsPortfolio.Columns(COL_PORT_LAST_OPT_DATE).NumberFormat = "mm/dd/yyyy" ' Last Opt Date
        wsPortfolio.Columns(COL_PORT_LAST_DATE_ON_FILE).NumberFormat = "mm/dd/yyyy" ' OOS Begin Date
    End If

    ' Highlight negative profits in red
    Dim col As Long
    For col = COL_PORT_PROFIT_LAST_1_MONTH To COL_PORT_PROFIT_SINCE_OOS_START ' Columns containing profits
        wsPortfolio.Columns(col).FormatConditions.Delete ' Clear existing conditions
        wsPortfolio.Columns(col).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        wsPortfolio.Columns(col).FormatConditions(1).Font.Color = RGB(255, 0, 0) ' Red font for negative values
    Next col
    
    For col = COL_PORT_EXPECTED_ANNUAL_PROFIT To COL_PORT_ACTUAL_ANNUAL_PROFIT ' Columns containing profits
        wsPortfolio.Columns(col).FormatConditions.Delete ' Clear existing conditions
        wsPortfolio.Columns(col).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        wsPortfolio.Columns(col).FormatConditions(1).Font.Color = RGB(255, 0, 0) ' Red font for negative values
    Next col

    For col = COL_PORT_SYMBOL To COL_PORT_PROFIT_SINCE_OOS_START ' Columns containing profits
        wsPortfolio.Columns(col).ColumnWidth = 12
    Next col
        For col = COL_PORT_STRATEGYCOUNT To COL_PORT_OPEN_CODE_TAB ' Columns containing profits
        wsPortfolio.Columns(col).ColumnWidth = 10
    Next col
    For col = COL_PORT_SYMBOL To COL_PORT_STATUS ' Columns containing profits
        wsPortfolio.Columns(col).ColumnWidth = 10
    Next col
    
    wsPortfolio.Columns(COL_PORT_CONTRACTS).ColumnWidth = 12
    
    wsPortfolio.Columns(COL_PORT_CREATE_DETAILED_TAB).ColumnWidth = 11

    For col = COL_PORT_OPEN_CODE_TAB To COL_PORT_FOLDER ' Columns containing profits
        wsPortfolio.Columns(col).ColumnWidth = 7.5
    Next col
    
    wsPortfolio.Columns(COL_PORT_STRATEGY_NAME).ColumnWidth = 87.5
    
    If GetNamedRangeValue("startcolwidthoverride") <> "" Then
        For col = COL_PORT_STRATEGYCOUNT To COL_PORT_FOLDER ' Columns containing profits
            If GetNamedRangeValue("startcolwidthoverride") < 7.5 Then wsPortfolio.Cells(1, col).WrapText = False
            wsPortfolio.Columns(col).ColumnWidth = GetNamedRangeValue("startcolwidthoverride")
        Next col
    End If
    
    
    With ThisWorkbook.Windows(1)
        .Zoom = 70 ' Set zoom level to 70%
    End With
    
End Sub




Sub OpenSizingSectorGraphs()

    Dim wsPortfolio As Worksheet
    
    
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Call InitializeColumnConstantsManually
    
    ' Check if "Summary" sheet exists and has data in row 2
    On Error Resume Next
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsPortfolio Is Nothing Then
        MsgBox "Error: 'Portfolio' sheet does not exist.", vbExclamation
         Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Exit and show error if the sheet exists but has no data in row 2
    If wsPortfolio.Cells(2, 1).value = "" Then
        MsgBox "Error: 'Portfolio' sheet exists but contains no data in row 2.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Call Deletetab("SizingGraphs")
    Call Deletetab("SectorTypeGraphs")

    Call CreateSizingGraphs
    Call CreateSectorTypeGraphs
    
    Call OrderVisibleTabsBasedOnList
    
    Sheets("SizingGraphs").Activate
    
    Application.ScreenUpdating = True

End Sub


Sub PortMetrics(outputRow As Long)
    Dim wsPortClosedTrade As Worksheet
    Dim wsPortfolioGraphs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim totalPnL As Double, winningPnL As Double, losingPnL As Double
    Dim totalTrades As Long, winningTrades As Long, losingTrades As Long
    Dim portfolioWinRate As Double, avgProfit As Double, avgLoss As Double
    Dim riskToReward As Double, edge As Double
    Dim col As Long, row As Long
    Dim pnlValue As Double

    ' Set the worksheet containing the data
    Set wsPortClosedTrade = ThisWorkbook.Sheets("PortClosedTrade")
    Set wsPortfolioGraphs = ThisWorkbook.Sheets("PortfolioGraphs")
    
    ' Find the last row and last column
    lastRow = wsPortClosedTrade.Cells(wsPortClosedTrade.rows.count, 1).End(xlUp).row
    lastCol = wsPortClosedTrade.Cells(1, wsPortClosedTrade.Columns.count).End(xlToLeft).column
    
    ' Initialize variables
    totalPnL = 0
    winningPnL = 0
    losingPnL = 0
    totalTrades = 0
    winningTrades = 0
    losingTrades = 0
    

    ' Loop through each strategy column (starting from column 2 to skip 'Date')
    For col = 2 To lastCol
        For row = 2 To lastRow ' Start from row 2 to skip headers
            pnlValue = wsPortClosedTrade.Cells(row, col).value
            
            ' Only consider non-zero PnL entries
            If pnlValue <> 0 Then
                totalTrades = totalTrades + 1
                totalPnL = totalPnL + pnlValue
                
                If pnlValue > 0 Then
                    winningTrades = winningTrades + 1
                    winningPnL = winningPnL + pnlValue
                ElseIf pnlValue < 0 Then
                    losingTrades = losingTrades + 1
                    losingPnL = losingPnL + pnlValue
                End If
            End If
        Next row
    Next col
    
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
    
     
    ' Output the results
    wsPortfolioGraphs.Cells(outputRow, 2).value = "Portfolio Stats"
    wsPortfolioGraphs.Cells(outputRow + 1, 2).value = "Total Trades"
    wsPortfolioGraphs.Cells(outputRow + 2, 2).value = "Win Rate (%)"
    wsPortfolioGraphs.Cells(outputRow + 3, 2).value = "Avg Profit"
    wsPortfolioGraphs.Cells(outputRow + 4, 2).value = "Avg Loss"
    wsPortfolioGraphs.Cells(outputRow + 5, 2).value = "Risk to Reward"
    wsPortfolioGraphs.Cells(outputRow + 6, 2).value = "Edge"
    
    wsPortfolioGraphs.Cells(outputRow + 1, 3).value = totalTrades
    wsPortfolioGraphs.Cells(outputRow + 1, 3).NumberFormat = "#,##0"
    wsPortfolioGraphs.Cells(outputRow + 2, 3).value = portfolioWinRate
    wsPortfolioGraphs.Cells(outputRow + 2, 3).NumberFormat = "0%"
    wsPortfolioGraphs.Cells(outputRow + 3, 3).value = avgProfit
    wsPortfolioGraphs.Cells(outputRow + 3, 3).NumberFormat = "$#,##0"
    wsPortfolioGraphs.Cells(outputRow + 4, 3).value = avgLoss
    wsPortfolioGraphs.Cells(outputRow + 4, 3).NumberFormat = "$#,##0"
    wsPortfolioGraphs.Cells(outputRow + 5, 3).value = riskToReward
    wsPortfolioGraphs.Cells(outputRow + 5, 3).NumberFormat = "0.0"
    wsPortfolioGraphs.Cells(outputRow + 6, 3).value = edge
    wsPortfolioGraphs.Cells(outputRow + 6, 3).NumberFormat = "$#,##0"
    
    With wsPortfolioGraphs
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
End With
    
End Sub


Sub CreatePortfolioGraphs(wsTotalGraphs As Worksheet, PortfolioGraphsTab As String, wsPortfolio As Worksheet)
    'Dim wsTotalGraphs As Worksheet
    Dim wsPortfolioGraphs As Worksheet
    Dim lastRow As Long
    Dim annualProfitsDict As Object
    Dim annualMaxDrawdownDict As Object
    Dim yearKey As Variant
    Dim i As Long
    Dim currentYear As String
    Dim annualProfit As Double
    Dim drawdown As Double
    ' At the beginning of the function, add:
    Dim displayAsPercentage As Boolean
    
    
    ' Set the sheet references
    'Set wsTotalGraphs = ThisWorkbook.Sheets("TotalPortfolioM2M")
    
     On Error Resume Next
    displayAsPercentage = GetNamedRangeValue("DisplayAsPercentage")
    On Error GoTo 0
    
    ' If named range doesn't exist, default to dollar display
    If Err.Number <> 0 Then
        displayAsPercentage = False
    End If
    
    ' Create a new sheet for summary graphs and tables if it doesn't exist
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(PortfolioGraphsTab).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsPortfolioGraphs = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsPortfolioGraphs.name = PortfolioGraphsTab
    
    If PortfolioGraphsTab = "PortfolioGraphs" Then
        wsPortfolioGraphs.Tab.Color = RGB(0, 176, 240)
    Else
    wsPortfolioGraphs.Tab.Color = RGB(255, 255, 0) ' Yellow tab color
    End If
    
    ' Set white background color for the entire worksheet
    wsPortfolioGraphs.Cells.Interior.Color = RGB(255, 255, 255)

 
  ' Determine the last row in TotalPortfolioM2M
    lastRow = wsTotalGraphs.Cells(wsTotalGraphs.rows.count, 1).End(xlUp).row
    
    ' Initialize dictionaries for storing annual profits and max drawdowns
    Set annualProfitsDict = CreateObject("Scripting.Dictionary")
    Set annualMaxDrawdownDict = CreateObject("Scripting.Dictionary")
    Set monthlyProfitsDict = CreateObject("Scripting.Dictionary")
    Set weeklyProfitsDict = CreateObject("Scripting.Dictionary")
    Set monthOnlyProfitsDict = CreateObject("Scripting.Dictionary")
    
    ' Initialize an array to store the count of each month (1 to 12)
    Dim monthCountArray(1 To 12) As Long
    
        ' Define month names
        Dim monthNames(1 To 12) As String
        monthNames(1) = "January"
        monthNames(2) = "February"
        monthNames(3) = "March"
        monthNames(4) = "April"
        monthNames(5) = "May"
        monthNames(6) = "June"
        monthNames(7) = "July"
        monthNames(8) = "August"
        monthNames(9) = "September"
        monthNames(10) = "October"
        monthNames(11) = "November"
        monthNames(12) = "December"
    
    
    ' Loop through each row to calculate annual profits and max drawdown
    For i = 2 To lastRow
        currentYear = wsTotalGraphs.Cells(i, 5).value ' Column 5 contains the year
        annualProfit = wsTotalGraphs.Cells(i, 2).value ' Column 2 contains the Total Daily Profit
        drawdown = wsTotalGraphs.Cells(i, 4).value ' Column 4 contains the Total Drawdown
        currentMonth = currentYear & "-" & wsTotalGraphs.Cells(i, 6).value ' Year-Month in Columns 5 and 6
        currentWeek = currentMonth & ": W" & (day(wsTotalGraphs.Cells(i, 1).value) - 1) \ 7 + 1
        monthOnly = wsTotalGraphs.Cells(i, 6).value
        
        ' Accumulate annual profits
        If Not annualProfitsDict.Exists(currentYear) Then
            annualProfitsDict.Add currentYear, 0
        End If
        annualProfitsDict(currentYear) = annualProfitsDict(currentYear) + annualProfit
        
        ' Accumulate monthly profits
        If Not monthlyProfitsDict.Exists(currentMonth) Then
            monthlyProfitsDict.Add currentMonth, 0
            monthCountArray(wsTotalGraphs.Cells(i, 6).value) = monthCountArray(wsTotalGraphs.Cells(i, 6).value) + 1
        End If
        monthlyProfitsDict(currentMonth) = monthlyProfitsDict(currentMonth) + annualProfit
   
        ' Accumulate MonthOnly profits
        If Not monthOnlyProfitsDict.Exists(monthOnly) Then
            monthOnlyProfitsDict.Add monthOnly, 0
        End If
        monthOnlyProfitsDict(monthOnly) = monthOnlyProfitsDict(monthOnly) + annualProfit
  
  
  
         ' Accumulate weekly profits
        If Not weeklyProfitsDict.Exists(currentWeek) Then
            weeklyProfitsDict.Add currentWeek, 0
        End If
        weeklyProfitsDict(currentWeek) = weeklyProfitsDict(currentWeek) + annualProfit
        
        ' Track maximum drawdown for each year
        If Not annualMaxDrawdownDict.Exists(currentYear) Then
            annualMaxDrawdownDict.Add currentYear, drawdown
        Else
            If drawdown > annualMaxDrawdownDict(currentYear) Then
                annualMaxDrawdownDict(currentYear) = drawdown
            End If
        End If
    Next i
    
       ' Generate table for annual profits and maximum drawdown
    Dim row As Long
    row = 1
    wsPortfolioGraphs.Cells(row, 1).value = "Year"
    wsPortfolioGraphs.Cells(row, 2).value = "Annual Profit"
    wsPortfolioGraphs.Cells(row, 3).value = "Annual Max Drawdown"
    wsPortfolioGraphs.Cells(row, 3).HorizontalAlignment = xlLeft
    row = row + 1
    
    For Each yearKey In annualProfitsDict.keys
        wsPortfolioGraphs.Cells(row, 1).value = yearKey
        wsPortfolioGraphs.Cells(row, 2).value = annualProfitsDict(yearKey)
        wsPortfolioGraphs.Cells(row, 3).value = annualMaxDrawdownDict(yearKey)
        row = row + 1
    Next yearKey

    ' Apply formatting to the table
    With wsPortfolioGraphs.Range("A1:C1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 230, 241) ' Light blue background for headers
        .WrapText = True
    End With
    wsPortfolioGraphs.Range("A2:A" & row - 1).HorizontalAlignment = xlCenter
    
    wsPortfolioGraphs.Range("B2:C" & row - 1).NumberFormat = "$#,##0" ' Dollar format with no decimals

    ' Autofit columns for readability
    For col = 1 To 3 ' Columns containing profits
        wsPortfolioGraphs.Columns(col).ColumnWidth = 20
    Next col

    ' Calculate Average Annual Profit and Maximum Drawdown
    Dim totalProfit As Double, avgProfit As Double, maxDrawdown As Double
    totalProfit = Application.WorksheetFunction.sum(wsPortfolioGraphs.Range("B2:B" & row - 1))
    If annualProfitsDict.count <> 0 Then
      avgProfit = totalProfit / annualProfitsDict.count
    Else
       avgProfit = 0
    End If
    maxDrawdown = Application.WorksheetFunction.Max(wsPortfolioGraphs.Range("C2:C" & row - 1))
    
    Dim startingEquity As Double
    startingEquity = GetNamedRangeValue("PortfolioStartingEquity")
    
    ' Display Average Annual Profit and Maximum Drawdown
    wsPortfolioGraphs.Cells(row + 1, 2).value = "Average Annual Profit ($)"
    wsPortfolioGraphs.Cells(row + 1, 2).Font.Bold = True
    wsPortfolioGraphs.Cells(row + 1, 2).HorizontalAlignment = xlRight
    wsPortfolioGraphs.Cells(row + 1, 3).value = avgProfit
    wsPortfolioGraphs.Cells(row + 1, 3).NumberFormat = "$#,##0"
        
    wsPortfolioGraphs.Cells(row + 2, 2).value = "Average Annual Profit (%)"
    wsPortfolioGraphs.Cells(row + 2, 2).Font.Bold = True
    wsPortfolioGraphs.Cells(row + 2, 2).HorizontalAlignment = xlRight
    wsPortfolioGraphs.Cells(row + 2, 3).value = avgProfit / startingEquity
    wsPortfolioGraphs.Cells(row + 2, 3).NumberFormat = "0%"
    
    wsPortfolioGraphs.Cells(row + 3, 2).value = "Maximum Drawdown ($)"
    wsPortfolioGraphs.Cells(row + 3, 2).Font.Bold = True
    wsPortfolioGraphs.Cells(row + 3, 2).HorizontalAlignment = xlRight
    wsPortfolioGraphs.Cells(row + 3, 3).value = maxDrawdown
    wsPortfolioGraphs.Cells(row + 3, 3).NumberFormat = "$#,##0"
    
    wsPortfolioGraphs.Cells(row + 4, 2).value = "Maximum Drawdown (%)"
    wsPortfolioGraphs.Cells(row + 4, 2).Font.Bold = True
    wsPortfolioGraphs.Cells(row + 4, 2).HorizontalAlignment = xlRight
    wsPortfolioGraphs.Cells(row + 4, 3).value = maxDrawdown / startingEquity
    wsPortfolioGraphs.Cells(row + 4, 3).NumberFormat = "0%"
    
    
    
    

    ' Apply formatting to the table
    With wsPortfolioGraphs.Range("A1:B1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 230, 241) ' Light blue background for headers
    End With
    wsPortfolioGraphs.Range("A2:A" & row - 1).HorizontalAlignment = xlCenter
    wsPortfolioGraphs.Range("B2:B" & row - 1).NumberFormat = "$#,##0" ' Dollar format with no decimals
  

    ' Autofit columns for readability
    wsPortfolioGraphs.Columns("A:C").AutoFit

    
    ' Write Monthly Profits Table
    Dim monthlyStartRow As Long
    Dim monthlyRow As Long
    monthlyStartRow = 1
    wsTotalGraphs.Cells(monthlyStartRow, 10).value = "Month"
    wsTotalGraphs.Cells(monthlyStartRow, 11).value = "Monthly Profit"
    monthlyRow = monthlyStartRow + 1
    For Each monthKey In monthlyProfitsDict.keys
        wsTotalGraphs.Cells(monthlyRow, 10).value = monthKey
        wsTotalGraphs.Cells(monthlyRow, 11).value = monthlyProfitsDict(monthKey)
        monthlyRow = monthlyRow + 1
    Next monthKey
    
    
     ' Write Weekly Profits Table
    Dim weeklyStartRow As Long
    Dim weeklyRow As Long
    weeklyStartRow = 1
    wsTotalGraphs.Cells(weeklyStartRow, 13).value = "Week"
    wsTotalGraphs.Cells(weeklyStartRow, 14).value = "Weekly Profit"
    weeklyRow = weeklyStartRow + 1
    For Each weekKey In weeklyProfitsDict.keys
        wsTotalGraphs.Cells(weeklyRow, 13).value = weekKey
        wsTotalGraphs.Cells(weeklyRow, 14).value = weeklyProfitsDict(weekKey)
        weeklyRow = weeklyRow + 1
    Next weekKey
    
    
    ' Declare and initialize a Collection
    Dim sortedMonthKeys As Collection
    Set sortedMonthKeys = New Collection
    
    ' Add keys from MonthOnlyProfitsDict to the Collection
    'Dim monthKey As Variant
    For Each monthKey In monthOnlyProfitsDict.keys
        sortedMonthKeys.Add CInt(monthKey) ' Ensure keys are treated as integers
    Next monthKey
    
    ' Convert the Collection to an array for sorting
    Dim monthKeysArray() As Long
    
    ReDim monthKeysArray(1 To sortedMonthKeys.count)
    
    i = 1
    For Each monthKey In sortedMonthKeys
        monthKeysArray(i) = monthKey
        i = i + 1
    Next monthKey
    

    
    ' Sort the array in chronological order using a custom sorting function
    Call SortMonthsArray(monthKeysArray)
    
    ' Declare dictionaries to store aggregated profits and counts
    Dim MonthAggregatedProfits As Object
    Set MonthAggregatedProfits = CreateObject("Scripting.Dictionary")
    
    Dim monthCount As Object
    Set monthCount = CreateObject("Scripting.Dictionary")
    
    ' Initialize dictionaries for all months
    Dim mIndex As Integer
    For mIndex = 1 To 12
        MonthAggregatedProfits(monthNames(mIndex)) = 0
        monthCount(monthNames(mIndex)) = 0
    Next mIndex
    
    ' Aggregate profits and counts per month
    For i = LBound(monthKeysArray) To UBound(monthKeysArray)
        currentMonth = monthNames(monthKeysArray(i)) ' Get the month name from the numeric value
        
        ' Ensure MonthOnlyProfitsDict contains valid keys
        If monthOnlyProfitsDict.Exists(monthKeysArray(i)) Then
            MonthAggregatedProfits(currentMonth) = MonthAggregatedProfits(currentMonth) + monthOnlyProfitsDict(monthKeysArray(i))
            monthCount(currentMonth) = monthCount(currentMonth) + 1
        End If
    Next i
    
    ' Write Monthly Profits Table
    Dim MonthOnlyStartRow As Long
    Dim MonthOnlyRow As Long
    
    MonthOnlyStartRow = 1
    wsTotalGraphs.Cells(MonthOnlyStartRow, 16).value = "Month"
    wsTotalGraphs.Cells(MonthOnlyStartRow, 17).value = "Total Monthly Profit"
    MonthOnlyRow = MonthOnlyStartRow + 1
    
    For i = LBound(monthKeysArray) To UBound(monthKeysArray)
        wsTotalGraphs.Cells(MonthOnlyRow, 16).value = monthNames(monthKeysArray(i))
        wsTotalGraphs.Cells(MonthOnlyRow, 17).value = monthOnlyProfitsDict(monthKeysArray(i))
        MonthOnlyRow = MonthOnlyRow + 1
    Next i
    
    ' Write Average Monthly Profits Table

    
    MonthOnlyRow = MonthOnlyRow + 2 ' Leave some space below the monthly profits table
    wsTotalGraphs.Cells(MonthOnlyRow, 16).value = "Month"
    wsTotalGraphs.Cells(MonthOnlyRow, 17).value = "Average Monthly Profit"
    MonthOnlyRow = MonthOnlyRow + 1
    
    ' Calculate and write averages in chronological order
    Dim avgProfit2 As Double
    
    For mIndex = 1 To 12
        currentMonth = monthNames(mIndex)
        If monthCount(currentMonth) > 0 Then
            avgProfit2 = MonthAggregatedProfits(currentMonth) / monthCountArray(mIndex)
        Else
            avgProfit2 = 0
        End If
        
        wsTotalGraphs.Cells(MonthOnlyRow, 16).value = currentMonth
        wsTotalGraphs.Cells(MonthOnlyRow, 17).value = avgProfit2
        MonthOnlyRow = MonthOnlyRow + 1
    Next mIndex
    
    ' Clean up
    Set MonthAggregatedProfits = Nothing
    Set monthCount = Nothing
    ' Generate cumulative profit chart (using columns A and C for Date and Cumulative P/L)
    
    
    Dim chart As ChartObject
   
    startingEquity = GetNamedRangeValue("PortfolioStartingEquity")
    
    ' Add this at the beginning of your chart generation code
    Dim chartTitleSuffix As String
    If displayAsPercentage Then
        chartTitleSuffix = " (% of $" & Format(startingEquity, "#,##0") & ")"
    Else
        chartTitleSuffix = " ($)"
    End If

    
    
    Set chart = wsPortfolioGraphs.ChartObjects.Add(left:=300, Width:=600, top:=0, Height:=350)
    With chart.chart
        .SetSourceData source:=wsTotalGraphs.Range("A2:A" & lastRow & ",C2:C" & lastRow) ' Date and Cumulative Profit columns
        .ChartType = xlLine
        .HasTitle = True
        
        If displayAsPercentage Then
            ' Create a temporary column for percentage calculations
            Dim tempCol As Range
            Set tempCol = wsTotalGraphs.Range("Z2:Z" & lastRow)
            ' Calculate percentage values
            For i = 2 To lastRow
                wsTotalGraphs.Cells(i, 26).value = wsTotalGraphs.Cells(i, 3).value / startingEquity
            Next i
            
            ' Update chart data source to use percentage column
            .SeriesCollection(1).values = wsTotalGraphs.Range("Z2:Z" & lastRow)
            .chartTitle.text = "Cumulative Profits Over Time" & chartTitleSuffix
        Else
            .chartTitle.text = "Cumulative Profits Over Time" & chartTitleSuffix
        End If
        
        ' Set axis titles safely
        If .Axes.count > 0 Then
            ' Category axis (X-axis)
            If Not .Axes(xlCategory, xlPrimary).HasTitle Then
                .Axes(xlCategory, xlPrimary).HasTitle = True
            End If
            .Axes(xlCategory, xlPrimary).AxisTitle.text = "Date"
            
            ' Value axis (Y-axis)
            If Not .Axes(xlValue, xlPrimary).HasTitle Then
                .Axes(xlValue, xlPrimary).HasTitle = True
            End If
            
            If displayAsPercentage Then
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Cumulative Profit (%)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0.00%"
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Cumulative Profit ($)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "$#,##0"
            End If
        End If
        
        .SeriesCollection(1).Border.Weight = xlMedium ' Increase line thickness
        .HasLegend = False ' Remove legend
    End With
    
    ' Add benchmark to the chart if enabled
    Call ModifyPortfolioChartForBenchmark(chart, wsTotalGraphs, lastRow, startingEquity, displayAsPercentage)
    
    
    
    ' Generate drawdown chart
    Set chart = wsPortfolioGraphs.ChartObjects.Add(left:=950, Width:=600, top:=0, Height:=350)
    With chart.chart
        .SetSourceData source:=wsTotalGraphs.Range("A2:A" & lastRow & ",D2:D" & lastRow) ' Date and DrawDown
        .SeriesCollection.NewSeries
        .ChartType = xlLine
        .HasTitle = True
        
        If displayAsPercentage Then
            ' Create a temporary column for percentage calculations
            Dim tempColDrawdown As Range
            Set tempColDrawdown = wsTotalGraphs.Range("AA2:AA" & lastRow)
            ' Calculate percentage values
            For i = 2 To lastRow
                wsTotalGraphs.Cells(i, 27).value = wsTotalGraphs.Cells(i, 4).value / startingEquity
            Next i
            
            ' Update chart data source to use percentage column
            .SeriesCollection(1).values = wsTotalGraphs.Range("AA2:AA" & lastRow)
            .chartTitle.text = "Drawdown Over Time" & chartTitleSuffix
        Else
            .SeriesCollection(1).values = wsTotalGraphs.Range("D2:D" & lastRow) ' Use Drawdown data
            .chartTitle.text = "Drawdown Over Time" & chartTitleSuffix
        End If
        
        .SeriesCollection(1).XValues = wsTotalGraphs.Range("A2:A" & lastRow)
        .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(255, 0, 0) ' Red line for drawdown
        
        ' Set axis titles safely
        If .Axes.count > 0 Then
            ' Category axis (X-axis)
            If Not .Axes(xlCategory, xlPrimary).HasTitle Then
                .Axes(xlCategory, xlPrimary).HasTitle = True
            End If
            .Axes(xlCategory, xlPrimary).AxisTitle.text = "Date"
            
            ' Value axis (Y-axis)
            If Not .Axes(xlValue, xlPrimary).HasTitle Then
                .Axes(xlValue, xlPrimary).HasTitle = True
            End If
            
            If displayAsPercentage Then
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Drawdown (%)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0.00%"
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Drawdown ($)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "$#,##0"
            End If
        End If
        
        .SeriesCollection(1).Border.Weight = xlMedium ' Increase line thickness
        .HasLegend = False ' Remove legend
    End With
    
    ' Generate drawdown percent chart (always in percent)
  '  Set chart = wsPortfolioGraphs.ChartObjects.Add(left:=1600, Width:=600, top:=0, Height:=350)
  '  With chart.chart
  '      .SetSourceData source:=wsTotalGraphs.Range("A2:A" & lastRow & ",H2:H" & lastRow) ' Date and DrawDown Percent
  '      .SeriesCollection.NewSeries
  '      .SeriesCollection(1).values = wsTotalGraphs.Range("H2:H" & lastRow) ' Use Drawdown Percent data
  '      .SeriesCollection(1).xValues = wsTotalGraphs.Range("A2:A" & lastRow)
  '      .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(255, 0, 0) ' Red line for drawdown
  '      .ChartType = xlLine
  '      .HasTitle = True
  ''      .chartTitle.Text = "Drawdown Percent Over Time"
   '
        ' Set axis titles safely
  '      If .Axes.count > 0 Then
   '         ' Category axis (X-axis)
   '         If Not .Axes(xlCategory, xlPrimary).HasTitle Then
   '             .Axes(xlCategory, xlPrimary).HasTitle = True
   '         End If
   '         .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Date"
   '
   '         ' Value axis (Y-axis)
   '         If Not .Axes(xlValue, xlPrimary).HasTitle Then
   '             .Axes(xlValue, xlPrimary).HasTitle = True
    '        End If
    ''        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Drawdown Percent"
    ''        .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0%" ' Format Y-axis as percentage
    ''    End If
        
     '   .SeriesCollection(1).Border.Weight = xlMedium ' Increase line thickness
    '    .HasLegend = False ' Remove legend
   ' End With
    
    ' Generate annual profit bar chart
    Set chart = wsPortfolioGraphs.ChartObjects.Add(left:=300, Width:=600, top:=400, Height:=350)
    With chart.chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        
        ' Set the X and Y values explicitly
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = wsPortfolioGraphs.Range("A2:A" & row - 1) ' Set X-axis to Year
        
        If displayAsPercentage Then
            ' Create a temporary column for percentage calculations
            Dim tempColAnnual As Range
            Set tempColAnnual = wsTotalGraphs.Range("AB2:AB" & row)
            
            ' Calculate percentage values using the annual profits from the table
            For i = 2 To row - 1
                wsTotalGraphs.Cells(i, 28).value = wsPortfolioGraphs.Cells(i, 2).value / startingEquity
            Next i
            
            .SeriesCollection(1).values = wsTotalGraphs.Range("AB2:AB" & (row - 1))
            .chartTitle.text = "Annual Profits" & chartTitleSuffix
        Else
            .SeriesCollection(1).values = wsPortfolioGraphs.Range("B2:B" & row - 1)
            .chartTitle.text = "Annual Profits" & chartTitleSuffix
        End If
        
        ' Set axis titles safely
        If .Axes.count > 0 Then
            ' Category axis (X-axis)
            If Not .Axes(xlCategory, xlPrimary).HasTitle Then
                .Axes(xlCategory, xlPrimary).HasTitle = True
            End If
            .Axes(xlCategory, xlPrimary).AxisTitle.text = "Year"
            
            ' Value axis (Y-axis)
            If Not .Axes(xlValue, xlPrimary).HasTitle Then
                .Axes(xlValue, xlPrimary).HasTitle = True
            End If
            
            If displayAsPercentage Then
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Annual Profit (%)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0.00%"
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Annual Profit ($)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "$#,##0"
            End If
        End If
        
        ' Remove the legend
        .HasLegend = False
    End With
    
    ' Generate Monthly Profit Chart
    Set chart = wsPortfolioGraphs.ChartObjects.Add(left:=950, Width:=600, top:=400, Height:=350)
    With chart.chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        
        If displayAsPercentage Then
            ' Create a temporary column for percentage calculations
            Dim tempColMonthly As Range
            Set tempColMonthly = wsTotalGraphs.Range("AC2:AC" & monthlyRow)
            ' Calculate percentage values
            For i = 2 To monthlyRow - 1
                wsTotalGraphs.Cells(i, 29).value = wsTotalGraphs.Cells(i, 11).value / startingEquity
            Next i
            
            .SeriesCollection.NewSeries
            .SeriesCollection(1).values = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(monthlyRow - 61, 2), 29), wsTotalGraphs.Cells(monthlyRow - 1, 29))
            .SeriesCollection(1).XValues = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(monthlyRow - 61, 2), 10), wsTotalGraphs.Cells(monthlyRow - 1, 10))
            .chartTitle.text = "Monthly Profits (Last 5 Years)" & chartTitleSuffix
        Else
            .SeriesCollection.NewSeries
            .SeriesCollection(1).values = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(monthlyRow - 61, 2), 11), wsTotalGraphs.Cells(monthlyRow - 1, 11))
            .SeriesCollection(1).XValues = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(monthlyRow - 61, 2), 10), wsTotalGraphs.Cells(monthlyRow - 1, 10))
            .chartTitle.text = "Monthly Profits (Last 5 Years)" & chartTitleSuffix
        End If
        
        ' Set axis titles safely
        If .Axes.count > 0 Then
            ' Category axis (X-axis)
            If Not .Axes(xlCategory, xlPrimary).HasTitle Then
                .Axes(xlCategory, xlPrimary).HasTitle = True
            End If
            .Axes(xlCategory, xlPrimary).AxisTitle.text = "Month"
            
            ' Value axis (Y-axis)
            If Not .Axes(xlValue, xlPrimary).HasTitle Then
                .Axes(xlValue, xlPrimary).HasTitle = True
            End If
            
            If displayAsPercentage Then
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Profit (%)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0.00%"
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Profit ($)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "$#,##0"
            End If
        End If
        
        .HasLegend = False ' Remove legend
    End With
    
    ' Generate Weekly Profit Chart
    Set chart = wsPortfolioGraphs.ChartObjects.Add(left:=300, Width:=600, top:=800, Height:=350)
    With chart.chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        
        If displayAsPercentage Then
            ' Create a temporary column for percentage calculations
            Dim tempColWeekly As Range
            Set tempColWeekly = wsTotalGraphs.Range("AD2:AD" & weeklyRow)
            ' Calculate percentage values
            For i = 2 To weeklyRow - 1
                wsTotalGraphs.Cells(i, 30).value = wsTotalGraphs.Cells(i, 14).value / startingEquity
            Next i
            
            .SeriesCollection.NewSeries
            .SeriesCollection(1).values = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(weeklyRow - 53, 2), 30), wsTotalGraphs.Cells(weeklyRow - 1, 30))
            .SeriesCollection(1).XValues = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(weeklyRow - 53, 2), 13), wsTotalGraphs.Cells(weeklyRow - 1, 13))
            .chartTitle.text = "Weekly Profits (Last 52 Weeks)" & chartTitleSuffix
        Else
            .SeriesCollection.NewSeries
            .SeriesCollection(1).values = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(weeklyRow - 53, 2), 14), wsTotalGraphs.Cells(weeklyRow - 1, 14))
            .SeriesCollection(1).XValues = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(weeklyRow - 53, 2), 13), wsTotalGraphs.Cells(weeklyRow - 1, 13))
            .chartTitle.text = "Weekly Profits (Last 52 Weeks)" & chartTitleSuffix
        End If
        
        ' Set axis titles safely
        If .Axes.count > 0 Then
            ' Category axis (X-axis)
            If Not .Axes(xlCategory, xlPrimary).HasTitle Then
                .Axes(xlCategory, xlPrimary).HasTitle = True
            End If
            .Axes(xlCategory, xlPrimary).AxisTitle.text = "Week"
            
            ' Value axis (Y-axis)
            If Not .Axes(xlValue, xlPrimary).HasTitle Then
                .Axes(xlValue, xlPrimary).HasTitle = True
            End If
            
            If displayAsPercentage Then
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Profit (%)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0.00%"
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Profit ($)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "$#,##0"
            End If
        End If
        
        .HasLegend = False ' Remove legend
    End With
    
    ' Generate Average Monthly Profit Chart
    Set chart = wsPortfolioGraphs.ChartObjects.Add(left:=950, Width:=600, top:=800, Height:=350)
    With chart.chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        
        If displayAsPercentage Then
            ' Create a temporary column for percentage calculations
            Dim tempColAvgMonthly As Range
            Set tempColAvgMonthly = wsTotalGraphs.Range("AE2:AE" & MonthOnlyRow)
            ' Calculate percentage values
            For i = MonthOnlyRow - 12 To MonthOnlyRow - 1
                wsTotalGraphs.Cells(i, 31).value = wsTotalGraphs.Cells(i, 17).value / startingEquity
            Next i
            
            .SeriesCollection.NewSeries
            .SeriesCollection(1).values = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(MonthOnlyRow - 12, 2), 31), wsTotalGraphs.Cells(MonthOnlyRow - 1, 31))
            .SeriesCollection(1).XValues = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(MonthOnlyRow - 12, 2), 16), wsTotalGraphs.Cells(MonthOnlyRow - 1, 16))
            .chartTitle.text = "Average Monthly Profits" & chartTitleSuffix
        Else
            .SeriesCollection.NewSeries
            .SeriesCollection(1).values = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(MonthOnlyRow - 12, 2), 17), wsTotalGraphs.Cells(MonthOnlyRow - 1, 17))
            .SeriesCollection(1).XValues = wsTotalGraphs.Range(wsTotalGraphs.Cells(Application.Max(MonthOnlyRow - 12, 2), 16), wsTotalGraphs.Cells(MonthOnlyRow - 1, 16))
            .chartTitle.text = "Average Monthly Profits" & chartTitleSuffix
        End If
        
        ' Set axis titles safely
        If .Axes.count > 0 Then
            ' Category axis (X-axis)
            If Not .Axes(xlCategory, xlPrimary).HasTitle Then
                .Axes(xlCategory, xlPrimary).HasTitle = True
            End If
            .Axes(xlCategory, xlPrimary).AxisTitle.text = "Month"
            
            ' Value axis (Y-axis)
            If Not .Axes(xlValue, xlPrimary).HasTitle Then
                .Axes(xlValue, xlPrimary).HasTitle = True
            End If
            
            If displayAsPercentage Then
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Profit (%)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0.00%"
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Profit ($)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "$#,##0"
            End If
        End If
        
        .HasLegend = False ' Remove legend
    End With
    
    ' === Step 3: Create Symbol Profit Bar Chart (Using Columns S:T) ===
    Dim symbolChart As ChartObject
    Set symbolChart = wsPortfolioGraphs.ChartObjects.Add(left:=300, Width:=600, top:=1200, Height:=350)
    With symbolChart.chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        
        ' Use columns S (Symbol) and T (Total Profit)
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = wsTotalGraphs.Range("S2:S" & wsTotalGraphs.Cells(wsTotalGraphs.rows.count, 19).End(xlUp).row)
        
        If displayAsPercentage Then
            ' Create a temporary column for percentage calculations
            Dim tempColSymbol As Range
            Dim symbolLastRow As Long
            symbolLastRow = wsTotalGraphs.Cells(wsTotalGraphs.rows.count, 20).End(xlUp).row
            Set tempColSymbol = wsTotalGraphs.Range("AF2:AF" & symbolLastRow)
            
            ' Calculate percentage values
            For i = 2 To symbolLastRow
                wsTotalGraphs.Cells(i, 32).value = wsTotalGraphs.Cells(i, 20).value / startingEquity
            Next i
            
            .SeriesCollection(1).values = wsTotalGraphs.Range("AF2:AF" & symbolLastRow)
            .chartTitle.text = "Symbol Profits" & chartTitleSuffix
        Else
            .SeriesCollection(1).values = wsTotalGraphs.Range("T2:T" & wsTotalGraphs.Cells(wsTotalGraphs.rows.count, 20).End(xlUp).row)
            .chartTitle.text = "Symbol Profits" & chartTitleSuffix
        End If
        
        ' Set axis titles safely
        If .Axes.count > 0 Then
            ' Category axis (X-axis)
            If Not .Axes(xlCategory, xlPrimary).HasTitle Then
                .Axes(xlCategory, xlPrimary).HasTitle = True
            End If
            .Axes(xlCategory, xlPrimary).AxisTitle.text = "Symbol"
            
            ' Value axis (Y-axis)
            If Not .Axes(xlValue, xlPrimary).HasTitle Then
                .Axes(xlValue, xlPrimary).HasTitle = True
            End If
            
            If displayAsPercentage Then
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Total Profit (%)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0.00%"
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Total Profit ($)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "$#,##0"
            End If
        End If
        
        .SeriesCollection(1).Format.line.Weight = 2.25
        .HasLegend = False
    End With
    
    ' === Step 4: Create Sector Profit Bar Chart (Using Columns V:W) ===
    Dim sectorChart As ChartObject
    Set sectorChart = wsPortfolioGraphs.ChartObjects.Add(left:=950, Width:=600, top:=1200, Height:=350)
    With sectorChart.chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        
        ' Use columns V (Sector) and W (Total Profit)
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = wsTotalGraphs.Range("V2:V" & wsTotalGraphs.Cells(wsTotalGraphs.rows.count, 22).End(xlUp).row)
        
        If displayAsPercentage Then
            ' Create a temporary column for percentage calculations
            Dim tempColSector As Range
            Dim sectorLastRow As Long
            sectorLastRow = wsTotalGraphs.Cells(wsTotalGraphs.rows.count, 23).End(xlUp).row
            Set tempColSector = wsTotalGraphs.Range("AG2:AG" & sectorLastRow)
            
            ' Calculate percentage values
            For i = 2 To sectorLastRow
                wsTotalGraphs.Cells(i, 33).value = wsTotalGraphs.Cells(i, 23).value / startingEquity
            Next i
            
            .SeriesCollection(1).values = wsTotalGraphs.Range("AG2:AG" & sectorLastRow)
            .chartTitle.text = "Sector Profits" & chartTitleSuffix
        Else
            .SeriesCollection(1).values = wsTotalGraphs.Range("W2:W" & wsTotalGraphs.Cells(wsTotalGraphs.rows.count, 23).End(xlUp).row)
            .chartTitle.text = "Sector Profits" & chartTitleSuffix
        End If
        
        ' Set axis titles safely
        If .Axes.count > 0 Then
            ' Category axis (X-axis)
            If Not .Axes(xlCategory, xlPrimary).HasTitle Then
                .Axes(xlCategory, xlPrimary).HasTitle = True
            End If
            .Axes(xlCategory, xlPrimary).AxisTitle.text = "Sector"
            
            ' Value axis (Y-axis)
            If Not .Axes(xlValue, xlPrimary).HasTitle Then
                .Axes(xlValue, xlPrimary).HasTitle = True
            End If
            
            If displayAsPercentage Then
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Total Profit (%)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0.00%"
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.text = "Total Profit ($)"
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "$#,##0"
            End If
        End If
        
        .SeriesCollection(1).Format.line.Weight = 2.25
        .HasLegend = False
    End With
    
    
    ' Add toggle button for display mode
    Dim btnToggle As Object
    Set btnToggle = wsPortfolioGraphs.Buttons.Add(left:=wsPortfolioGraphs.Cells(2, 1).left + 30, top:=wsPortfolioGraphs.Cells(32, 1).top, Width:=180, Height:=25)
    With btnToggle
        If displayAsPercentage Then
            .Caption = "Switch to Dollar Display"
        Else
            .Caption = "Switch to Percentage Display"
        End If
        If PortfolioGraphsTab = "PortfolioGraphs" Then
            .OnAction = "ToggleDisplayMode"
        Else
            .OnAction = "ToggleDisplayModeBacktest"
        End If
    End With



    ' Autofit columns for readability
    wsPortfolioGraphs.Columns("A:C").ColumnWidth = 16
    ' Create delete button
    Dim btn As Object
    Set btn = wsPortfolioGraphs.Buttons.Add(left:=wsPortfolioGraphs.Cells(2, 1).left + 30, top:=wsPortfolioGraphs.Cells(35, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Delete Tab"
        .OnAction = "Delete" & PortfolioGraphsTab ' Make sure to create this sub to handle deletion
    End With

    ' Create navigation buttons for different sections
    Set btn = wsPortfolioGraphs.Buttons.Add(left:=wsPortfolioGraphs.Cells(2, 3).left, top:=wsPortfolioGraphs.Cells(35, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With
    Set btn = wsPortfolioGraphs.Buttons.Add(left:=wsPortfolioGraphs.Cells(2, 1).left + 30, top:=wsPortfolioGraphs.Cells(38, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With
    Set btn = wsPortfolioGraphs.Buttons.Add(left:=wsPortfolioGraphs.Cells(2, 3).left, top:=wsPortfolioGraphs.Cells(38, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With
    
    Set btn = wsPortfolioGraphs.Buttons.Add(left:=wsPortfolioGraphs.Cells(2, 1).left + 30, top:=wsPortfolioGraphs.Cells(41, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Back to Strategies"
        .OnAction = "GoToStrategies"
    End With
    Set btn = wsPortfolioGraphs.Buttons.Add(left:=wsPortfolioGraphs.Cells(2, 3).left, top:=wsPortfolioGraphs.Cells(41, 1).top, Width:=100, Height:=25)
    With btn
        .Caption = "Back to Inputs"
        .OnAction = "GoToInputs"
    End With

    ' Set zoom level
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
End Sub



Sub CreateSizingGraphs()
    Dim wsPortfolio As Worksheet
    Dim wsSizingGraphs As Worksheet
    Dim lastRow As Long
    Dim chart As ChartObject
    Dim chartLeft As Long, chartTop As Long
    Dim i As Long
    Dim labelarray() As String

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    
    ' Set the sheet references
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    
    ' Create or clear the PortfolioGraphs sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("SizingGraphs").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsSizingGraphs = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsSizingGraphs.name = "SizingGraphs"
    wsSizingGraphs.Tab.Color = RGB(0, 176, 240)
    
    ' Set white background color for the entire worksheet
    wsSizingGraphs.Cells.Interior.Color = RGB(255, 255, 255)
    
    ' Determine the last row in Portfolio
    lastRow = wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row

    ' Set positions for charts
    chartLeft = 150
    chartTop = 20

    ' Create an array for column numbers and titles
    Dim chartColumns As Variant, chartTitles As Variant
    chartColumns = Array(COL_PORT_MAX_IS_OOS_DRAWDOWN, COL_PORT_AVG_IS_OOS_DRAWDOWN, COL_PORT_MAX_DRAWDOWN_LAST_12_MONTHS, COL_PORT_NOTIONAL_CAPITAL, _
                         COL_PORT_AVG_IS_OOS_TRADE, COL_PORT_LARGEST_LOSS_IS_OOS_TRADE, COL_PORT_AVG_LOSS_IS_OOS_TRADE, COL_PORT_EXPECTED_ANNUAL_PROFIT, COL_PORT_MARGIN, COL_PORT_IS_ANNUAL_SD_ISOOS, COL_PORT_ATR_LAST_1_MONTH, COL_PORT_ATR_LAST_3_MONTHS, _
                         COL_PORT_ATR_LAST_6_MONTHS, COL_PORT_ATR_LAST_12_MONTHS, COL_PORT_ATR_LAST_24_MONTHS, COL_PORT_ATR_LAST_60_MONTHS, COL_PORT_ATR_ALL_DATA)
    chartTitles = Array("Max Drawdown", "Average Drawdown", "Max Drawdown (Last 12 Months)", "Notional Capital", "Average Trade Size", _
                        "Largest Unprofitable Trade (IS+OOS)", "Avg Unprofitable Trade (IS+OOS)", "Expected Annual Profit", "Margin", "Annual Standard Deviation", "ATR (Last Month)", "ATR (Last 3 Months)", "ATR (Last 6 Months)", "ATR (Last 12 Months)", "ATR (Last 24 Months)", "ATR (Last 60 Months)", "ATR (All Time)")
    

    
        ' Create label array for X-axis combining Strategy Count and Symbol
    ReDim labelarray(1 To lastRow - 1)
    For i = 2 To lastRow
        labelarray(i - 1) = wsPortfolio.Cells(i, COL_PORT_STRATEGYCOUNT).value & " - " & wsPortfolio.Cells(i, COL_PORT_SYMBOL).value
    Next i
    
    wsSizingGraphs.Range("AA1").value = "Strategy List"
    For i = 2 To lastRow
        wsSizingGraphs.Range("AA" & i).value = wsPortfolio.Cells(i, COL_PORT_STRATEGYCOUNT).value & " - " & wsPortfolio.Cells(i, COL_PORT_SYMBOL).value & " - " & wsPortfolio.Cells(i, COL_PORT_STRATEGY_NAME).value
    Next i
    
    
      ' Loop through each column in chartColumns array
    Dim startEquity As Double
    Dim sizingOption As Boolean
    
    
    ' Get the Portfolio Starting Equity
    startEquity = GetNamedRangeValue("PortfolioStartingEquity")
    
    ' Get the sizing option (Dollar or Percentage)
    sizingOption = GetNamedRangeValue("DisplayAsPercentage")
    
    ' Add label button for the metric dropdown
    Dim labelBtnSizing As Object
    Set labelBtnSizing = wsSizingGraphs.Buttons.Add(Left:=150, Top:=5, Width:=115, Height:=22)
    With labelBtnSizing
        .Caption = "Select Metric:"
        .OnAction = ""
    End With

    ' Add dropdown for strategy metric selection
    Dim ddSizing As DropDown
    Set ddSizing = wsSizingGraphs.DropDowns.Add(Left:=272, Top:=5, Width:=350, Height:=22)
    With ddSizing
        .name = "SizingMetricDropDown"
        For i = LBound(chartTitles) To UBound(chartTitles)
            .AddItem chartTitles(i)
        Next i
        .AddItem "Count of Strategies per Contract"
        .ListIndex = 1
        .OnAction = "RefreshSizingGraph"
    End With


    Dim contract As String
    Dim contractCounts As Variant
    Dim currentCount As Long
   
    ' Extract unique contracts and their counts
    Set uniqueContracts = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        contract = wsPortfolio.Cells(i, COL_PORT_SYMBOL).value
        If uniqueContracts.Exists(contract) Then
            uniqueContracts(contract) = uniqueContracts(contract) + 1
        Else
            uniqueContracts.Add contract, 1
        End If
    Next i
    
    Dim key As Variant
    Dim value As Variant
    
    ' Write unique contracts and counts to SizingGraphs sheet
    wsSizingGraphs.Cells(1, 40).value = "Unique Contract"
    wsSizingGraphs.Cells(1, 41).value = "Count"
    j = 2
    
    For Each key In uniqueContracts.keys
        ' Write the contract (key) and its count (value) to the worksheet
        wsSizingGraphs.Cells(j, 40).value = key ' Key is the contract name
        wsSizingGraphs.Cells(j, 41).value = uniqueContracts(key) ' Value is the count
        j = j + 1
    Next key



    ' Create initial chart for the default dropdown selection
    Call RefreshSizingGraph



    ' Autofit columns in the PortfolioGraphs sheet for readability
    wsSizingGraphs.Columns("A:Z").AutoFit



 ' Create delete button
    Dim btn As Object
    Set btn = wsSizingGraphs.Buttons.Add(left:=wsSizingGraphs.Cells(1, 1).left + 10, top:=wsSizingGraphs.Cells(1, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteSizingGraphs" ' Make sure to create this sub to handle deletion
    End With

    ' Create navigation buttons for different sections
    Set btn = wsSizingGraphs.Buttons.Add(left:=wsSizingGraphs.Cells(1, 1).left + 10, top:=wsSizingGraphs.Cells(4, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With
    Set btn = wsSizingGraphs.Buttons.Add(left:=wsSizingGraphs.Cells(1, 1).left + 10, top:=wsSizingGraphs.Cells(7, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With
    Set btn = wsSizingGraphs.Buttons.Add(left:=wsSizingGraphs.Cells(1, 1).left + 10, top:=wsSizingGraphs.Cells(10, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With
    Set btn = wsSizingGraphs.Buttons.Add(left:=wsSizingGraphs.Cells(1, 1).left + 10, top:=wsSizingGraphs.Cells(13, 1).top + 10, Width:=100, Height:=25)
      With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies"
    End With
    Set btn = wsSizingGraphs.Buttons.Add(left:=wsSizingGraphs.Cells(1, 1).left + 10, top:=wsSizingGraphs.Cells(16, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs"
    End With


    ' Set zoom level
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With

   
End Sub


Sub CreateSectorTypeGraphs()
    Dim wsPortfolio As Worksheet
    Dim wsSectorTypeGraphs As Worksheet
    Dim lastRow As Long
    Dim chart As ChartObject
    Dim chartLeft As Long, chartTop As Long
    Dim i As Long
    Dim sectorDict As Object
    Dim sector As String
    Dim sectorRow As Long
    Dim sectorLabels() As String
    Dim column As Variant
    Dim chartColumns As Variant, chartTitles As Variant

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually

    ' Set the sheet references
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")

    ' Create or clear the SectorTypeGraphs sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("SectorTypeGraphs").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsSectorTypeGraphs = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsSectorTypeGraphs.name = "SectorTypeGraphs"
    wsSectorTypeGraphs.Tab.Color = RGB(0, 176, 240)
    
      ' Set white background color for the entire worksheet
    wsSectorTypeGraphs.Cells.Interior.Color = RGB(255, 255, 255)
    
    ' Determine the last row in Portfolio
    lastRow = wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row

    ' Initialize column arrays
    chartColumns = Array(COL_PORT_MAX_IS_OOS_DRAWDOWN, COL_PORT_AVG_IS_OOS_DRAWDOWN, COL_PORT_MAX_DRAWDOWN_LAST_12_MONTHS, COL_PORT_NOTIONAL_CAPITAL, _
                         COL_PORT_AVG_IS_OOS_TRADE, COL_PORT_EXPECTED_ANNUAL_PROFIT, COL_PORT_MARGIN, COL_PORT_IS_ANNUAL_SD_ISOOS, COL_PORT_ATR_LAST_1_MONTH, COL_PORT_ATR_LAST_3_MONTHS, _
                         COL_PORT_ATR_LAST_6_MONTHS, COL_PORT_ATR_LAST_12_MONTHS, COL_PORT_ATR_LAST_24_MONTHS, COL_PORT_ATR_LAST_60_MONTHS, COL_PORT_ATR_ALL_DATA)
    chartTitles = Array("Max Drawdown", "Average Drawdown", "Max Drawdown (Last 12 Months)", "Notional Capital", "Average Trade Size", _
                        "Expected Annual Profit", "Margin", "Annual Standard Deviation", "ATR (Last Month)", "ATR (Last 3 Months)", "ATR (Last 6 Months)", "ATR (Last 12 Months)", "ATR (Last 24 Months)", "ATR (Last 60 Months)", "ATR (All Time)")
    

    
    ' Collect selected columns based on Yes/No inputs
    Set selectedColumns = New Collection
    If GetNamedRangeValue("GraphSymbol") = "Yes" Then selectedColumns.Add COL_PORT_SYMBOL
    If GetNamedRangeValue("GraphSector") = "Yes" Then selectedColumns.Add COL_PORT_SECTOR
    If GetNamedRangeValue("GraphType") = "Yes" Then selectedColumns.Add COL_PORT_TYPE
    If GetNamedRangeValue("GraphHorizon") = "Yes" Then selectedColumns.Add COL_PORT_HORIZON
    If GetNamedRangeValue("GraphTimeFrame") = "Yes" Then selectedColumns.Add COL_PORT_TIMEFRAME
    If GetNamedRangeValue("GraphLongShort") = "Yes" Then selectedColumns.Add COL_PORT_LONGSHORT

   
      ' Check if at least one column is selected
    If selectedColumns.count = 0 Then
        MsgBox "No columns selected for the X-axis. Please update the Config tab.", vbExclamation
        Exit Sub
    End If

    ' Initialize dictionary for aggregation
    Set aggregationDict = CreateObject("Scripting.Dictionary")
    Dim categoryCounts As Object
    Set categoryCounts = CreateObject("Scripting.Dictionary")
    
     ' Aggregate data by unique combinations
    For i = 2 To lastRow
        ' Create a unique key for the combination of selected columns
        Dim categoryKey As String
        categoryKey = ""
        For Each column In selectedColumns
            If Trim(wsPortfolio.Cells(i, column).value) = "" Then
                categoryKey = categoryKey & "N/A|"
            Else
                categoryKey = categoryKey & wsPortfolio.Cells(i, column).value & "|"
            End If
        Next column
        categoryKey = left(categoryKey, Len(categoryKey) - 1) ' Remove trailing "|"

        ' Initialize or aggregate data for this key
        If Not aggregationDict.Exists(categoryKey) Then
            aggregationDict.Add categoryKey, Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) ' One slot for each metric
            categoryCounts.Add categoryKey, 0 ' Initialize count
        End If

        ' Aggregate values for this key
        Dim currentValues As Variant
        currentValues = aggregationDict(categoryKey)
        For column = LBound(chartColumns) To UBound(chartColumns)
            If IsNumeric(wsPortfolio.Cells(i, chartColumns(column)).value) Then
                currentValues(column) = currentValues(column) + wsPortfolio.Cells(i, chartColumns(column)).value
            End If
        Next column
        aggregationDict(categoryKey) = currentValues

        ' Increment count for this key
        categoryCounts(categoryKey) = categoryCounts(categoryKey) + 1
    Next i

    ' Debug: Check if the dictionary contains valid data
    If aggregationDict.count = 0 Then
        MsgBox "No valid data was aggregated. Please check the input data.", vbExclamation
        Exit Sub
    End If

    ' Prepare data for charts
    Dim dictKeys As Variant
    dictKeys = aggregationDict.keys
    ReDim uniqueLabels(1 To aggregationDict.count)
    ReDim aggregatedValues(1 To aggregationDict.count, 1 To UBound(chartColumns) + 1)
    ReDim counts(1 To aggregationDict.count)

    For K = 0 To aggregationDict.count - 1
        uniqueLabels(K + 1) = dictKeys(K)
        Dim values As Variant
        values = aggregationDict(dictKeys(K))
        For column = LBound(chartColumns) To UBound(chartColumns)
            aggregatedValues(K + 1, column + 1) = values(column)
        Next column
        counts(K + 1) = categoryCounts(dictKeys(K)) ' Get count
    Next K

    ' Write data to worksheet
    Dim dataStartRow As Long, dataStartCol As Long
    dataStartRow = 1
    dataStartCol = 40

    ' Write headers
    wsSectorTypeGraphs.Cells(dataStartRow, dataStartCol).value = "Category Key"
    wsSectorTypeGraphs.Cells(dataStartRow, dataStartCol + 1).value = "Count"
    For column = LBound(chartTitles) To UBound(chartTitles)
        wsSectorTypeGraphs.Cells(dataStartRow, dataStartCol + column + 2).value = chartTitles(column)
    Next column

    ' Write data
    For row = 1 To UBound(uniqueLabels)
        wsSectorTypeGraphs.Cells(dataStartRow + row, dataStartCol).value = uniqueLabels(row) ' Write labels
        wsSectorTypeGraphs.Cells(dataStartRow + row, dataStartCol + 1).value = counts(row) ' Write counts
        For column = LBound(chartColumns) To UBound(chartColumns)
            wsSectorTypeGraphs.Cells(dataStartRow + row, dataStartCol + column + 2).value = aggregatedValues(row, column + 1) ' Write values
        Next column
    Next row

    ' Create the other metric graphs
    Dim startEquity As Double
    Dim sizingOption As Boolean
    
    
    ' Get the Portfolio Starting Equity
    startEquity = GetNamedRangeValue("PortfolioStartingEquity")
    
    ' Get the sizing option (Dollar or Percentage)
    sizingOption = GetNamedRangeValue("DisplayAsPercentage")
    
    ' Add label button for the metric dropdown
    Dim labelBtnSector As Object
    Set labelBtnSector = wsSectorTypeGraphs.Buttons.Add(Left:=150, Top:=5, Width:=115, Height:=22)
    With labelBtnSector
        .Caption = "Select Metric:"
        .OnAction = ""
    End With

    ' Add dropdown for sector metric selection
    Dim ddSector As DropDown
    Set ddSector = wsSectorTypeGraphs.DropDowns.Add(Left:=272, Top:=5, Width:=350, Height:=22)
    With ddSector
        .name = "SectorMetricDropDown"
        For column = LBound(chartTitles) To UBound(chartTitles)
            .AddItem chartTitles(column)
        Next column
        .AddItem "Count of Categories"
        .ListIndex = 1
        .OnAction = "RefreshSectorTypeGraph"
    End With

    ' Create initial chart for the default dropdown selection
    Call RefreshSectorTypeGraph

    ' Autofit columns for readability
    wsSectorTypeGraphs.Columns("A:Z").AutoFit

' Create delete button
    Dim btn As Object
    Set btn = wsSectorTypeGraphs.Buttons.Add(left:=wsSectorTypeGraphs.Cells(1, 1).left + 10, top:=wsSectorTypeGraphs.Cells(1, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteSectorTypeGraphs" ' Make sure to create this sub to handle deletion
    End With

    ' Create navigation buttons for different sections
    Set btn = wsSectorTypeGraphs.Buttons.Add(left:=wsSectorTypeGraphs.Cells(1, 1).left + 10, top:=wsSectorTypeGraphs.Cells(4, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With
    Set btn = wsSectorTypeGraphs.Buttons.Add(left:=wsSectorTypeGraphs.Cells(1, 1).left + 10, top:=wsSectorTypeGraphs.Cells(7, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With
    Set btn = wsSectorTypeGraphs.Buttons.Add(left:=wsSectorTypeGraphs.Cells(1, 1).left + 10, top:=wsSectorTypeGraphs.Cells(10, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With
    Set btn = wsSectorTypeGraphs.Buttons.Add(left:=wsSectorTypeGraphs.Cells(1, 1).left + 10, top:=wsSectorTypeGraphs.Cells(13, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies"
    End With
    Set btn = wsSectorTypeGraphs.Buttons.Add(left:=wsSectorTypeGraphs.Cells(1, 1).left + 10, top:=wsSectorTypeGraphs.Cells(16, 1).top + 10, Width:=100, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs"
    End With


    ' Set zoom level
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
End Sub


Sub RefreshSizingGraph()
    Dim ws As Worksheet
    Dim wsPortfolio As Worksheet
    Dim dd As DropDown
    Dim selectedIdx As Integer
    Dim chartColumns As Variant, chartTitles As Variant
    Dim co As ChartObject
    Dim startEquity As Double
    Dim sizingOption As Boolean
    Dim lastRow As Long
    Dim lastContractRow As Long
    Dim i As Long, j As Integer
    Dim labelarray() As String
    Dim valuesRange As Range
    Dim transformedValues() As Double
    Dim newChart As ChartObject

    Call InitializeColumnConstantsManually

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SizingGraphs")
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    On Error GoTo 0

    If ws Is Nothing Or wsPortfolio Is Nothing Then Exit Sub

    chartColumns = Array(COL_PORT_MAX_IS_OOS_DRAWDOWN, COL_PORT_AVG_IS_OOS_DRAWDOWN, COL_PORT_MAX_DRAWDOWN_LAST_12_MONTHS, COL_PORT_NOTIONAL_CAPITAL, _
                         COL_PORT_AVG_IS_OOS_TRADE, COL_PORT_LARGEST_LOSS_IS_OOS_TRADE, COL_PORT_AVG_LOSS_IS_OOS_TRADE, COL_PORT_EXPECTED_ANNUAL_PROFIT, COL_PORT_MARGIN, COL_PORT_IS_ANNUAL_SD_ISOOS, COL_PORT_ATR_LAST_1_MONTH, COL_PORT_ATR_LAST_3_MONTHS, _
                         COL_PORT_ATR_LAST_6_MONTHS, COL_PORT_ATR_LAST_12_MONTHS, COL_PORT_ATR_LAST_24_MONTHS, COL_PORT_ATR_LAST_60_MONTHS, COL_PORT_ATR_ALL_DATA)
    chartTitles = Array("Max Drawdown", "Average Drawdown", "Max Drawdown (Last 12 Months)", "Notional Capital", "Average Trade Size", _
                        "Largest Unprofitable Trade (IS+OOS)", "Avg Unprofitable Trade (IS+OOS)", "Expected Annual Profit", "Margin", "Annual Standard Deviation", "ATR (Last Month)", "ATR (Last 3 Months)", "ATR (Last 6 Months)", "ATR (Last 12 Months)", "ATR (Last 24 Months)", "ATR (Last 60 Months)", "ATR (All Time)")

    ' Get dropdown selection (ListIndex is 1-based; convert to 0-based)
    Set dd = ws.DropDowns("SizingMetricDropDown")
    selectedIdx = dd.ListIndex - 1

    startEquity = GetNamedRangeValue("PortfolioStartingEquity")
    sizingOption = GetNamedRangeValue("DisplayAsPercentage")
    lastRow = wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row

    ' Delete existing charts
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    If selectedIdx = UBound(chartTitles) + 1 Then
        ' Count of Strategies per Contract chart
        lastContractRow = ws.Cells(ws.rows.count, 40).End(xlUp).row
        Set newChart = ws.ChartObjects.Add(Left:=150, Top:=35, Width:=1050, Height:=550)
        With newChart.chart
            .ChartType = xlColumnClustered
            .HasTitle = True
            .chartTitle.text = "Count of Strategies per each contract"
            .HasLegend = False
            .SeriesCollection.NewSeries
            .SeriesCollection(1).XValues = ws.Range(ws.Cells(2, 40), ws.Cells(lastContractRow, 40))
            .SeriesCollection(1).values = ws.Range(ws.Cells(2, 41), ws.Cells(lastContractRow, 41))
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.text = "Symbol"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.text = "Count"
        End With
    Else
        ' Metric chart: read values directly from Portfolio sheet
        ReDim labelarray(1 To lastRow - 1)
        For i = 2 To lastRow
            labelarray(i - 1) = wsPortfolio.Cells(i, COL_PORT_STRATEGYCOUNT).value & " - " & wsPortfolio.Cells(i, COL_PORT_SYMBOL).value
        Next i

        Set valuesRange = wsPortfolio.Range(wsPortfolio.Cells(2, chartColumns(selectedIdx)), wsPortfolio.Cells(lastRow, chartColumns(selectedIdx)))
        ReDim transformedValues(1 To valuesRange.rows.count)
        For j = 1 To valuesRange.rows.count
            If sizingOption Then
                transformedValues(j) = valuesRange.Cells(j, 1).value / startEquity
            Else
                transformedValues(j) = valuesRange.Cells(j, 1).value
            End If
        Next j

        Set newChart = ws.ChartObjects.Add(Left:=150, Top:=35, Width:=1050, Height:=550)
        With newChart.chart
            .ChartType = xlColumnClustered
            .HasTitle = True
            If sizingOption Then
                .chartTitle.text = chartTitles(selectedIdx) & " (as % of " & Format(startEquity, "#,##0") & ")"
            Else
                .chartTitle.text = chartTitles(selectedIdx)
            End If
            .HasLegend = False
            .SeriesCollection.NewSeries
            .SeriesCollection(1).XValues = labelarray
            .SeriesCollection(1).values = transformedValues
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.text = "Strategy / Symbol"
            .Axes(xlValue).HasTitle = True
            If sizingOption Then
                .Axes(xlValue).AxisTitle.text = chartTitles(selectedIdx) & " (% of " & Format(startEquity, "#,##0") & ")"
                .Axes(xlValue).TickLabels.NumberFormat = "0.0%"
            Else
                .Axes(xlValue).AxisTitle.text = chartTitles(selectedIdx)
                .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
            End If
        End With
    End If
End Sub


Sub RefreshSectorTypeGraph()
    Dim ws As Worksheet
    Dim dd As DropDown
    Dim selectedIdx As Integer
    Dim chartTitles As Variant
    Dim co As ChartObject
    Dim startEquity As Double
    Dim sizingOption As Boolean
    Dim lastDataRow As Long
    Dim dataCol As Long
    Dim j As Integer
    Dim valuesRange As Range
    Dim transformedValues() As Double
    Dim newChart As ChartObject

    Call InitializeColumnConstantsManually

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SectorTypeGraphs")
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    chartTitles = Array("Max Drawdown", "Average Drawdown", "Max Drawdown (Last 12 Months)", "Notional Capital", "Average Trade Size", _
                        "Expected Annual Profit", "Margin", "Annual Standard Deviation", "ATR (Last Month)", "ATR (Last 3 Months)", "ATR (Last 6 Months)", "ATR (Last 12 Months)", "ATR (Last 24 Months)", "ATR (Last 60 Months)", "ATR (All Time)")

    ' Get dropdown selection (ListIndex is 1-based; convert to 0-based)
    Set dd = ws.DropDowns("SectorMetricDropDown")
    selectedIdx = dd.ListIndex - 1

    startEquity = GetNamedRangeValue("PortfolioStartingEquity")
    sizingOption = GetNamedRangeValue("DisplayAsPercentage")

    ' Find last data row in col 40 (Category Key)
    lastDataRow = ws.Cells(ws.rows.count, 40).End(xlUp).row

    ' Delete existing charts
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    Set newChart = ws.ChartObjects.Add(Left:=150, Top:=35, Width:=1050, Height:=550)

    If selectedIdx = UBound(chartTitles) + 1 Then
        ' Count of Categories chart (col 41 = counts)
        With newChart.chart
            .ChartType = xlColumnClustered
            .HasTitle = True
            .chartTitle.text = "Count of Categories"
            .HasLegend = False
            .SeriesCollection.NewSeries
            .SeriesCollection(1).XValues = ws.Range(ws.Cells(2, 40), ws.Cells(lastDataRow, 40))
            .SeriesCollection(1).values = ws.Range(ws.Cells(2, 41), ws.Cells(lastDataRow, 41))
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.text = "Category"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.text = "Count"
        End With
    Else
        ' Metric chart: col 42 = metric 0, col 43 = metric 1, etc.
        dataCol = 42 + selectedIdx

        Set valuesRange = ws.Range(ws.Cells(2, dataCol), ws.Cells(lastDataRow, dataCol))
        ReDim transformedValues(1 To valuesRange.rows.count)
        For j = 1 To valuesRange.rows.count
            If sizingOption Then
                transformedValues(j) = valuesRange.Cells(j, 1).value / startEquity
            Else
                transformedValues(j) = valuesRange.Cells(j, 1).value
            End If
        Next j

        With newChart.chart
            .ChartType = xlColumnClustered
            .HasTitle = True
            If sizingOption Then
                .chartTitle.text = chartTitles(selectedIdx) & " (as % of " & Format(startEquity, "#,##0") & ")"
            Else
                .chartTitle.text = chartTitles(selectedIdx)
            End If
            .HasLegend = False
            .SeriesCollection.NewSeries
            .SeriesCollection(1).XValues = ws.Range(ws.Cells(2, 40), ws.Cells(lastDataRow, 40))
            .SeriesCollection(1).values = transformedValues
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.text = "Sector Type"
            .Axes(xlValue).HasTitle = True
            If sizingOption Then
                .Axes(xlValue).AxisTitle.text = chartTitles(selectedIdx) & " (% of " & Format(startEquity, "#,##0") & ")"
                .Axes(xlValue).TickLabels.NumberFormat = "0.0%"
            Else
                .Axes(xlValue).AxisTitle.text = chartTitles(selectedIdx)
                .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
            End If
        End With
    End If
End Sub




Sub SetupPortfolioButtonsAndStrategyTabCreation()
    Dim wsPortfolio As Worksheet
    Dim btn As Object
    Dim lastRowSummary As Long
    Dim summaryRow As Long
    Dim buttonLeft As Double
    Dim buttonTop As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    
    ' Set the Summary worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    
    ' Find the last row with data in the Summary sheet
    lastRowSummary = wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row
    
    ' Loop through each strategy and add a button for creating detailed tabs in Column B
    For summaryRow = 2 To lastRowSummary
        ' Set the button's size and position to fit within the cell in Column B
        buttonLeft = wsPortfolio.Cells(summaryRow, COL_PORT_CREATE_DETAILED_TAB).left + 2
        buttonTop = wsPortfolio.Cells(summaryRow, COL_PORT_CREATE_DETAILED_TAB).top + 1
        buttonWidth = wsPortfolio.Cells(summaryRow, COL_PORT_CREATE_DETAILED_TAB).Width - 2
        buttonHeight = wsPortfolio.Cells(summaryRow, COL_PORT_CREATE_DETAILED_TAB).Height - 1
        
        ' Add button for creating strategy tab
        Set btn = wsPortfolio.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
        With btn
            .OnAction = "ButtonClickHandlerPort" ' Assign the macro to handle the button click
            .Caption = "<>" ' Clearer caption
            .name = "CreateTabBtn" & summaryRow ' Assign a unique name to each button
        End With
    Next summaryRow
End Sub



Sub SetupButtonsforCodeTabCreation()
    Dim wsSummary As Worksheet
    Dim btn As Object
    Dim lastRowSummary As Long
    Dim summaryRow As Long
    Dim buttonLeft As Double
    Dim buttonTop As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    
    ' Set the Summary worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    ' Find the last row with data in the Summary sheet
    lastRowSummary = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row
    
    ' Loop through each strategy and add a button for creating detailed tabs in Column B
    For summaryRow = 2 To lastRowSummary
        ' Set the button's size and position to fit within the cell in Column B
        buttonLeft = wsSummary.Cells(summaryRow, COL_OPEN_CODE_TAB).left + 2
        buttonTop = wsSummary.Cells(summaryRow, COL_OPEN_CODE_TAB).top + 1
        buttonWidth = wsSummary.Cells(summaryRow, COL_OPEN_CODE_TAB).Width - 2
        buttonHeight = wsSummary.Cells(summaryRow, COL_OPEN_CODE_TAB).Height - 1
        
        ' Add button for creating strategy tab
        Set btn = wsSummary.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
        With btn
            .OnAction = "ButtonClickHandlerCode" ' Assign the macro to handle the button click
            .Caption = "~~"
            .name = "CodeTabBtn" & summaryRow ' Assign a unique name to each button
        End With
    
    
        buttonLeft = wsSummary.Cells(summaryRow, COL_CODE_TEXT).left + 2
        buttonTop = wsSummary.Cells(summaryRow, COL_CODE_TEXT).top + 1
        buttonWidth = wsSummary.Cells(summaryRow, COL_CODE_TEXT).Width - 2
        buttonHeight = wsSummary.Cells(summaryRow, COL_CODE_TEXT).Height - 1
        
        ' Add button for creating strategy tab
        Set btn = wsSummary.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
        With btn
            .OnAction = "ButtonClickHandlerCodeText" ' Assign the macro to handle the button click
            .Caption = "**"
            .name = "CodeTextBtn" & summaryRow ' Assign a unique name to each button
        End With
        
        
        buttonLeft = wsSummary.Cells(summaryRow, COL_FOLDER).left + 2
        buttonTop = wsSummary.Cells(summaryRow, COL_FOLDER).top + 1
        buttonWidth = wsSummary.Cells(summaryRow, COL_FOLDER).Width - 2
        buttonHeight = wsSummary.Cells(summaryRow, COL_FOLDER).Height - 1
        
        ' Add button for creating strategy tab
        Set btn = wsSummary.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
        With btn
            .OnAction = "ButtonClickHandlerCodeFolder" ' Assign the macro to handle the button click
            .Caption = "%%"
            .name = "FolderOpenBtn" & summaryRow ' Assign a unique name to each button
        End With
    
    
    Next summaryRow
End Sub





Sub SetupPortfolioButtonsforCodeTabCreation()
    Dim wsPortfolio As Worksheet
    Dim btn As Object
    Dim lastRowSummary As Long
    Dim summaryRow As Long
    Dim buttonLeft As Double
    Dim buttonTop As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    
    ' Set the Summary worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    
    ' Find the last row with data in the Summary sheet
    lastRowSummary = wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row
    
    ' Loop through each strategy and add a button for creating detailed tabs in Column B
    For summaryRow = 2 To lastRowSummary
        ' Set the button's size and position to fit within the cell in Column B
        buttonLeft = wsPortfolio.Cells(summaryRow, COL_PORT_OPEN_CODE_TAB).left + 2
        buttonTop = wsPortfolio.Cells(summaryRow, COL_PORT_OPEN_CODE_TAB).top + 1
        buttonWidth = wsPortfolio.Cells(summaryRow, COL_PORT_OPEN_CODE_TAB).Width - 2
        buttonHeight = wsPortfolio.Cells(summaryRow, COL_PORT_OPEN_CODE_TAB).Height - 1
        
        ' Add button for creating strategy tab
        Set btn = wsPortfolio.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
        With btn
            .OnAction = "ButtonClickHandlerCodePort" ' Assign the macro to handle the button click
            .Caption = "~~"
            .name = "CodeTabBtn" & summaryRow ' Assign a unique name to each button
        End With
        
        
        
        buttonLeft = wsPortfolio.Cells(summaryRow, COL_PORT_CODE_TEXT).left + 2
        buttonTop = wsPortfolio.Cells(summaryRow, COL_PORT_CODE_TEXT).top + 1
        buttonWidth = wsPortfolio.Cells(summaryRow, COL_PORT_CODE_TEXT).Width - 2
        buttonHeight = wsPortfolio.Cells(summaryRow, COL_PORT_CODE_TEXT).Height - 1
        
        ' Add button for creating strategy tab
        Set btn = wsPortfolio.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
        With btn
            .OnAction = "ButtonClickHandlerCodePortText" ' Assign the macro to handle the button click
            .Caption = "**"
            .name = "CodeTextBtnText" & summaryRow ' Assign a unique name to each button
        End With
        
        
        buttonLeft = wsPortfolio.Cells(summaryRow, COL_PORT_FOLDER).left + 2
        buttonTop = wsPortfolio.Cells(summaryRow, COL_PORT_FOLDER).top + 1
        buttonWidth = wsPortfolio.Cells(summaryRow, COL_PORT_FOLDER).Width - 2
        buttonHeight = wsPortfolio.Cells(summaryRow, COL_PORT_FOLDER).Height - 1
        
        ' Add button for creating strategy tab
        Set btn = wsPortfolio.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
        With btn
            .OnAction = "ButtonClickHandlerCodePortFolder" ' Assign the macro to handle the button click
            .Caption = "%%"
            .name = "FolderOpenBtn" & summaryRow ' Assign a unique name to each button
        End With
        
        
    Next summaryRow
End Sub


Sub ButtonClickHandlerCodePort()
    Dim clickedButton As Object
    Dim wsPortfolio As Worksheet
    Dim buttonRow As Long
    Dim gStrategyName As String
    Dim gStrategyNumber As Long

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually

    ' Set the Summary worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")

    ' Get the button that was clicked
    On Error Resume Next
    Set clickedButton = wsPortfolio.Buttons(Application.Caller)
    On Error GoTo 0
    
    If clickedButton Is Nothing Then
        MsgBox "Button not recognized. Please ensure the button is linked correctly.", vbExclamation
        Exit Sub
    End If

    ' Find the row of the button (based on the button's position)
    buttonRow = clickedButton.TopLeftCell.row

    ' Set the global variable to the strategy number in that row
    gStrategyName = wsPortfolio.Cells(buttonRow, COL_PORT_STRATEGY_NAME).value
    gStrategyNumber = wsPortfolio.Cells(buttonRow, COL_PORT_STRATEGY_NUMBER).value
    ' Call the macro to create the strategy tab
    OpenStrategyCodeFile gStrategyName, gStrategyNumber, "tab"
    
End Sub



Sub ButtonClickHandlerCodePortText()
    Dim clickedButton As Object
    Dim wsPortfolio As Worksheet
    Dim buttonRow As Long
    Dim gStrategyName As String
    Dim gStrategyNumber As Long

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually

    ' Set the Summary worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")

    ' Get the button that was clicked
    On Error Resume Next
    Set clickedButton = wsPortfolio.Buttons(Application.Caller)
    On Error GoTo 0
    
    If clickedButton Is Nothing Then
        MsgBox "Button not recognized. Please ensure the button is linked correctly.", vbExclamation
        Exit Sub
    End If

    ' Find the row of the button (based on the button's position)
    buttonRow = clickedButton.TopLeftCell.row

    ' Set the global variable to the strategy number in that row
    gStrategyName = wsPortfolio.Cells(buttonRow, COL_PORT_STRATEGY_NAME).value
    gStrategyNumber = wsPortfolio.Cells(buttonRow, COL_PORT_STRATEGY_NUMBER).value
    ' Call the macro to create the strategy tab
    OpenStrategyCodeFile gStrategyName, gStrategyNumber, "file"
    
End Sub


Sub ButtonClickHandlerCodePortFolder()
    Dim clickedButton As Object
    Dim wsPortfolio As Worksheet
    Dim buttonRow As Long
    Dim gStrategyName As String
    Dim gStrategyNumber As Long

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually

    ' Set the Summary worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")

    ' Get the button that was clicked
    On Error Resume Next
    Set clickedButton = wsPortfolio.Buttons(Application.Caller)
    On Error GoTo 0
    
    If clickedButton Is Nothing Then
        MsgBox "Button not recognized. Please ensure the button is linked correctly.", vbExclamation
        Exit Sub
    End If

    ' Find the row of the button (based on the button's position)
    buttonRow = clickedButton.TopLeftCell.row

    ' Set the global variable to the strategy number in that row
    gStrategyName = wsPortfolio.Cells(buttonRow, COL_PORT_STRATEGY_NAME).value
    gStrategyNumber = wsPortfolio.Cells(buttonRow, COL_PORT_STRATEGY_NUMBER).value
    ' Call the macro to create the strategy tab
    OpenStrategyCodeFile gStrategyName, gStrategyNumber, "folder"
    
End Sub


Sub ToggleDisplayMode()
    ' Toggle the DisplayAsPercentage named range
    Dim currentValue As Boolean
    
    ' Check if the named range exists
    On Error Resume Next
    currentValue = GetNamedRangeValue("DisplayAsPercentage")
    
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
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set wsTotalGraphs = ThisWorkbook.Sheets("TotalPortfolioM2M")
    
    ' Delete the current PortfolioGraphs sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("PortfolioGraphs").Delete
    ThisWorkbook.Sheets("SizingGraphs").Delete
    ThisWorkbook.Sheets("SectorTypeGraphs").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Call your graph generation routine
    
    Call CreateSizingGraphs
    Call CreateSectorTypeGraphs
    
    Call CreatePortfolioGraphs(wsTotalGraphs, "PortfolioGraphs", wsPortfolio)
    
    ' Activate the new PortfolioGraphs sheet
    ThisWorkbook.Sheets("PortfolioGraphs").Activate
    
    Application.ScreenUpdating = True
End Sub

Sub CalculateBenchmarkData(wsTotalGraphs As Worksheet, wsPortfolioDailyM2M As Worksheet, startdate As Date, currentdate As Date)
    ' This sub calculates benchmark data and adds it to the TotalPortfolioM2M sheet
    
    Dim benchmarkEnabled As Boolean
    Dim benchmarkStrategy As String
    Dim benchmarkContracts As Double
    Dim strategyCol As Long
    Dim lastRow As Long
    Dim outputRow As Long
    Dim wsDailyM2M As Worksheet
    
    ' Step 1: Check if benchmarking is enabled
    On Error Resume Next
    benchmarkEnabled = GetNamedRangeValue("BenchMarkOption") = "Yes"
    On Error GoTo 0
    
    ' Get the DailyM2MEquity sheet which contains all strategy data
    On Error Resume Next
    Set wsDailyM2M = ThisWorkbook.Sheets("DailyM2MEquity")
    On Error GoTo 0
    
    If wsDailyM2M Is Nothing Then
        MsgBox "Error: Could not find DailyM2MEquity sheet.", vbExclamation
        Exit Sub
    End If
    
    If Not benchmarkEnabled Then
        Exit Sub ' Benchmark not enabled, exit the sub
    End If
    
    ' Step 2: Get benchmark strategy name and contracts
    On Error Resume Next
    benchmarkStrategy = GetNamedRangeValue("BenchMarkStrategy")
    benchmarkContracts = GetNamedRangeValue("BenchMarkContracts")
    On Error GoTo 0
    
    ' Validate benchmark data
    If benchmarkStrategy = "" Then
        Exit Sub ' No benchmark strategy specified
    End If
    
    If benchmarkContracts = 0 Then
        benchmarkContracts = 1 ' Default to 1 contract if not specified
    End If
    
    ' Step 3: Find the column with the benchmark strategy in the DailyM2MEquity sheet
    lastRow = wsDailyM2M.Cells(wsDailyM2M.rows.count, 1).End(xlUp).row
    
    ' Find the column with the benchmark strategy
    strategyCol = -1
    For col = 2 To wsDailyM2M.Cells(1, wsDailyM2M.Columns.count).End(xlToLeft).column
        If wsDailyM2M.Cells(1, col).value = benchmarkStrategy Then
            strategyCol = col
            Exit For
        End If
    Next col
    
    ' Exit if benchmark strategy not found
    If strategyCol = -1 Then
        MsgBox "Error: Benchmark strategy '" & benchmarkStrategy & "' not found in DailyM2MEquity sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Step 4: Add benchmark data to the TotalPortfolioM2M sheet
    ' Find where to add the benchmark data (column 40 as requested)
    Dim benchmarkStartCol As Long
    benchmarkStartCol = 40
    
    ' Add benchmark headers
    wsTotalGraphs.Cells(1, benchmarkStartCol).value = "Benchmark Date"
    wsTotalGraphs.Cells(1, benchmarkStartCol + 1).value = "Benchmark Daily Profit"
    wsTotalGraphs.Cells(1, benchmarkStartCol + 2).value = "Benchmark Cumulative P/L"
    
    ' Calculate benchmark P&L similar to portfolio P&L
    Dim benchmarkTotal As Double
    Dim dailyProfit As Double
    
    outputRow = 2 ' Start on second row
    benchmarkTotal = 0 ' Initialize cumulative P&L
    
    
    
    Dim currentDateInData As Date
    Dim lastRowDailyM2M As Integer

    
    
    lastRowDailyM2M = wsPortfolioDailyM2M.Cells(wsPortfolioDailyM2M.rows.count, 1).End(xlUp).row
    
    For row = 2 To lastRow
        currentDateInData = wsDailyM2M.Cells(row, 1).value
        ' Check if the date is within the start and end date range
        If currentDateInData >= startdate And currentDateInData <= currentdate Then
        
            Application.StatusBar = "Portfolio Results Calculating: " & Format((currentDateInData - startdate) / (currentdate - startdate), "0%") & " completed"
        
        
            ' Copy the date to TotalPortfolioM2M
            wsTotalGraphs.Cells(outputRow, benchmarkStartCol).value = currentDateInData

        
            ' Calculate daily profit scaled by contract count
            dailyProfit = wsDailyM2M.Cells(row, strategyCol).value * benchmarkContracts
            wsTotalGraphs.Cells(outputRow, benchmarkStartCol + 1).value = dailyProfit
            
            ' Calculate cumulative P&L
            benchmarkTotal = benchmarkTotal + dailyProfit
            wsTotalGraphs.Cells(outputRow, benchmarkStartCol + 2).value = benchmarkTotal
            
            outputRow = outputRow + 1
         End If
        
    Next row
End Sub

Sub ModifyPortfolioChartForBenchmark(chart As ChartObject, wsTotalGraphs As Worksheet, lastRow As Long, startingEquity As Double, displayAsPercentage As Boolean)
    ' This sub adds benchmark data to the specified chart
    
    Dim benchmarkEnabled As Boolean
    Dim benchmarkStrategy As String
    
    ' Check if benchmarking is enabled
    On Error Resume Next
    benchmarkEnabled = GetNamedRangeValue("BenchMarkOption") = "Yes"
    benchmarkStrategy = GetNamedRangeValue("BenchMarkStrategy")
    On Error GoTo 0
    
    If Not benchmarkEnabled Or benchmarkStrategy = "" Then
        Exit Sub ' Benchmark not enabled or not configured
    End If
    
    ' Find benchmark data columns (using column 42 for cumulative P/L as per your updated code)
    Dim benchmarkCol As Long
    benchmarkCol = 42 ' Column 42 for benchmark cumulative P/L
    
    ' First, check if the benchmark data exists in the TotalPortfolioM2M sheet
    If wsTotalGraphs.Cells(1, benchmarkCol).value <> "Benchmark Cumulative P/L" Then
        Exit Sub ' Benchmark data not found
    End If
    
    With chart.chart
        ' Add a new series for the benchmark
        .SeriesCollection.NewSeries
        
        ' Set X values (dates) - using column 40 for benchmark dates
        .SeriesCollection(2).XValues = wsTotalGraphs.Range(wsTotalGraphs.Cells(2, 40), wsTotalGraphs.Cells(lastRow, 40))
        
        If displayAsPercentage Then
            ' Create a temporary column for percentage calculations if not already done
            Dim tempColBenchmark As Range
            Set tempColBenchmark = wsTotalGraphs.Range("AQ2:AQ" & lastRow)
            
            ' Calculate percentage values
            For i = 2 To lastRow
                wsTotalGraphs.Cells(i, 43).value = wsTotalGraphs.Cells(i, benchmarkCol).value / startingEquity
            Next i
            
            ' Use percentage values for the benchmark
            .SeriesCollection(2).values = wsTotalGraphs.Range("AQ2:AQ" & lastRow)
        Else
            ' Use dollar values for the benchmark
            .SeriesCollection(2).values = wsTotalGraphs.Range(wsTotalGraphs.Cells(2, benchmarkCol), _
                                                          wsTotalGraphs.Cells(lastRow, benchmarkCol))
        End If
        
        ' Format the benchmark series
        .SeriesCollection(2).name = benchmarkStrategy & " (Benchmark)"
        .SeriesCollection(2).Border.colorIndex = 3 ' Red for benchmark
        .SeriesCollection(2).Border.Weight = xlMedium ' Medium line thickness
        
        ' Turn on legend since we now have multiple series
        .HasLegend = True
        .Legend.position = xlLegendPositionBottom
    End With
End Sub

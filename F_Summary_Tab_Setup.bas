Attribute VB_Name = "F_Summary_Tab_Setup"


Sub UpdateStrategySummaryWithArray(Optional ByVal resetStrategiestab As String = "No")
    

    
    ' Define the worksheet names
    Dim wsDetails As Worksheet
    Dim wsM2MEquity As Worksheet
    Dim wsSummary As Worksheet
    Dim wsFolderLocations As Worksheet ' Add this line
    Dim wsStrategies As Worksheet
    Dim wsLatestPosition As Worksheet

    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        Exit Sub
    End If

    Call RetrieveAllFolderData("Yes")

' Check if "MW Folder Locations" sheet exists and has data in row 2
    On Error Resume Next
    Set wsFolderLocations = ThisWorkbook.Sheets("MW Folder Locations") ' Add this line
    On Error GoTo 0
    
    
    On Error Resume Next
    Set wsStrategies = ThisWorkbook.Sheets("Strategies") ' Add this line
    On Error GoTo 0

    On Error Resume Next
    Set wsLatestPosition = ThisWorkbook.Sheets("LatestPositionData") ' Add this line
    On Error GoTo 0


    ' Exit and show error if the sheet doesn't exist
    If wsFolderLocations Is Nothing Then
        MsgBox "Error: 'MW Folder Locations' sheet does not exist.", vbExclamation
        Exit Sub
    End If
    
    
    
    
    If wsStrategies Is Nothing Then
        MsgBox "Error: 'Strategies Folder Locations' sheet does not exist.", vbExclamation
        Exit Sub
    End If
    
    
    If wsLatestPosition Is Nothing Then
        MsgBox "Error: 'Latest Position' sheet does not exist.", vbExclamation
        Exit Sub
    End If
    
    
    ' Exit and show error if the sheet exists but has no data in row 2
    If wsFolderLocations.Cells(2, 1).value = "" Then
        MsgBox "Error: 'MW Folder Locations' sheet exists but contains no data in row 2.", vbExclamation
        Exit Sub
    End If
    
    
   

    If wsFolderLocations.Cells(1, 1).value <> "Folder Count" Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Please correct issues highlighted in the 'MW Folder Locations' or 'Strategies' tab first before continuing...", vbExclamation
        Exit Sub
    End If
    
        If wsStrategies.Cells(2, 2).value = "Duplicate Strategy" Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Please remove duplicates from the 'Strategies' tab first before continuing...", vbExclamation
        Exit Sub
    End If
  

' Check if "DailyM2MEquity" sheet exists and has data in row 2
    On Error Resume Next
    Set wsM2MEquity = ThisWorkbook.Sheets("DailyM2MEquity") ' Add this line
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsM2MEquity Is Nothing Then
        MsgBox "Error: 'DailyM2MEquity' sheet does not exist.", vbExclamation
        Exit Sub
    End If

    ' Exit and show error if the sheet exists but has no data in row 2
    If wsM2MEquity.Cells(2, 2).value = "" Then
        MsgBox "Error: 'DailyM2MEquity' sheet exists but contains no data in row 2.", vbExclamation
        Exit Sub
    End If

' Check if "Walkforward Details" sheet exists and has data in row 2
    On Error Resume Next
    Set wsDetails = ThisWorkbook.Sheets("Walkforward Details") ' Add this line
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsDetails Is Nothing Then
        MsgBox "Error: 'Walkforward Details' sheet does not exist.", vbExclamation
        Exit Sub
    End If

    ' Exit and show error if the sheet exists but has no data in row 2
    If wsDetails.Cells(2, 2).value = "" Then
        MsgBox "Error: 'Walkforward Details' sheet exists but contains no data in row 2.", vbExclamation
        Exit Sub
    End If
 


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' Create a new sheet for summary table if it doesn't exist
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsSummary = ThisWorkbook.Sheets.Add(After:=wsFolderLocations)
    wsSummary.name = "Summary"
    wsSummary.Tab.Color = RGB(240, 0, 0)
    
     ' Set white background color for the entire worksheet
    wsSummary.Cells.Interior.Color = RGB(255, 255, 255)
    
    
    ' Create the header array from the "Walkforward Details" tab
    Dim headerArray() As Variant
    headerArray = CreateHeaderArray(wsDetails)
       
    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    
    Call DeleteAllStrategyTabsNoPrompt
    
     Application.StatusBar = "May the tracker be with you..."
    
    
    ' Define columns for output in Summary
    wsSummary.Cells(1, COL_STRATEGY_NUMBER).value = "Strategy Number"
    wsSummary.Cells(1, COL_CREATE_DETAILED_TAB).value = "Detailed Summary"
    wsSummary.Cells(1, COL_OPEN_CODE_TAB).value = "Code Tab"
    wsSummary.Cells(1, COL_CODE_TEXT).value = "Text File"
    wsSummary.Cells(1, COL_FOLDER).value = "Folder"
    wsSummary.Cells(1, COL_STRATEGY_NAME).value = "Strategy Name"
    wsSummary.Cells(1, COL_SYMBOL).value = "Symbol"
    wsSummary.Cells(1, COL_SECTOR).value = "Sector"
    wsSummary.Cells(1, COL_TIMEFRAME).value = "Bar Size"
    wsSummary.Cells(1, COL_MARGIN).value = "Margin"
    wsSummary.Cells(1, COL_STATUS).value = "Status (Input Column)" ' Added heading for Status
    wsSummary.Cells(1, COL_ELIGIBILITY).value = "Eligibile for Trading?" ' Added heading for Status
    wsSummary.Cells(1, COL_SYMBOL_RANK).value = "Sybl Rank"
    wsSummary.Cells(1, COL_SECTOR_RANK).value = "Sect Rank"
    wsSummary.Cells(1, COL_NEXT_OPT_DATE).value = "Next Opt Date"
    wsSummary.Cells(1, COL_LAST_OPT_DATE).value = "Last Opt Date"
    wsSummary.Cells(1, COL_OOS_BEGIN_DATE).value = "OOS Begin Date"
    wsSummary.Cells(1, COL_LAST_DATE_ON_FILE).value = "Last Date On File"
    wsSummary.Cells(1, COL_EXPECTED_ANNUAL_PROFIT).value = "Expected Annual Profit"
    wsSummary.Cells(1, COL_ACTUAL_ANNUAL_PROFIT).value = "Actual Annual Profit"
    wsSummary.Cells(1, COL_RETURN_EFFICIENCY).value = "Return Efficiency"
    wsSummary.Cells(1, COL_EXPECTED_ANNUAL_RETURN).value = "Expected Annual Return"
    wsSummary.Cells(1, COL_ACTUAL_ANNUAL_RETURN).value = "Actual Annual Return"
    wsSummary.Cells(1, COL_RUN_MC).value = "Run Monte Carlo" ' Added heading
    wsSummary.Cells(1, COL_BACKTEST_MC).value = "Monte Carlo" ' Added heading
    wsSummary.Cells(1, COL_CLOSEDTRADEMC).value = "Closed Trade MC" ' Added heading
    
    'wsSummary.Cells(1, COL_RISK_RUIN).value = "Risk of Ruin" ' Added heading
    wsSummary.Cells(1, COL_NOTIONAL_CAPITAL).value = "Notional Capital" ' Added heading
    wsSummary.Cells(1, COL_LAST_12_MONTHS).value = "Last 12 Months"
    wsSummary.Cells(1, COL_IS_MONTE_CARLO).value = "MW Monte Carlo (IS)"
    wsSummary.Cells(1, COL_IS_OOS_MONTE_CARLO).value = "MW Monte Carlo (IS + OOS)"
    
    
    
    wsSummary.Cells(1, COL_LONG_SHORT).value = "Long / Short"
    
    wsSummary.Cells(1, COL_PROFIT_FACTOR).value = "Profit Factor (IS + OOS)"
    wsSummary.Cells(1, COL_PROFIT_LONG_FACTOR).value = "Long Profit Factor (IS + OOS)"
    wsSummary.Cells(1, COL_PROFIT_SHORT_FACTOR).value = "Short Profit Factor (IS + OOS)"
    wsSummary.Cells(1, COL_GROSS_LONG_NETPROFIT).value = "Long Net Profit (IS + OOS)"
    wsSummary.Cells(1, COL_GROSS_LONG_PROFIT).value = "Long Gross Profit (IS + OOS)"
    wsSummary.Cells(1, COL_GROSS_LONG_LOSS).value = "Long Gross Loss (IS + OOS)"
    wsSummary.Cells(1, COL_GROSS_SHORT_NETPROFIT).value = "Short Net Profit (IS + OOS)"
    wsSummary.Cells(1, COL_GROSS_SHORT_PROFIT).value = "Short Gross Profit (IS + OOS)"
    wsSummary.Cells(1, COL_GROSS_SHORT_LOSS).value = "Short Gross Loss (IS + OOS)"
    
    
    wsSummary.Cells(1, COL_IS_ANNUAL_SD_IS).value = "Annual Standard Deviation (IS)"
    wsSummary.Cells(1, COL_IS_ANNUAL_SD_ISOOS).value = "Annual Standard Deviation (IS + OOS)"
    wsSummary.Cells(1, COL_R_DD_12MONTH).value = "Return to Drawdown Last 12 Months"
    wsSummary.Cells(1, COL_R_DD_OOS).value = "Return to Drawdown (OOS)"
    wsSummary.Cells(1, COL_BACKTEST_WINRATE).value = "Winrate (IS)"
    wsSummary.Cells(1, COL_OOS_WINRATE).value = "Winrate (OSS)"
    wsSummary.Cells(1, COL_OVERALL_WINRATE).value = "Winrate (IS + OSS)"
    wsSummary.Cells(1, COL_TRADES_PER_YEAR).value = "Trades Per Year"
    wsSummary.Cells(1, COL_PERCENT_TIME_IN_MARKET).value = "Percent Time in Market"
    wsSummary.Cells(1, COL_AVG_TRADE_LENGTH).value = "Average Trade in Market (days)"
    
    wsSummary.Cells(1, COL_AVG_IS_OOS_TRADE).value = "Average Trade (IS+OOS)"
    wsSummary.Cells(1, COL_AVG_PROFIT_IS_OOS_TRADE).value = "Avg Profitable Trade (IS+OOS)"
    wsSummary.Cells(1, COL_AVG_LOSS_IS_OOS_TRADE).value = "Avg Unprofitable Trade (IS+OOS)"
    wsSummary.Cells(1, COL_LARGEST_WIN_IS_OOS_TRADE).value = " Largest Profitable Trade (IS+OOS)"
    wsSummary.Cells(1, COL_LARGEST_LOSS_IS_OOS_TRADE).value = "Largest Unprofitable Trade (IS+OOS)"
    wsSummary.Cells(1, COL_TOTAL_IS_PROFIT).value = "Total Profit (IS)"
    wsSummary.Cells(1, COL_TOTAL_IS_OSS_PROFIT).value = "Total Profit (IS+OOS)"
    wsSummary.Cells(1, COL_ANNUALIZED_NET_PROFIT_IS_OOS).value = "IS+OOS Annualized Net Profit"
    
    
    
    wsSummary.Cells(1, COL_WORST_BACKTEST_DRAWDOWN).value = "Max Drawdown (IS)"
    wsSummary.Cells(1, COL_WORST_IS_OOS_DRAWDOWN).value = "Max Drawdown (IS + OOS)"
    wsSummary.Cells(1, COL_AVG_BACKTEST_DRAWDOWN).value = "Average Drawdown (IS)"
    wsSummary.Cells(1, COL_AVG_IS_OOS_DRAWDOWN).value = "Average Drawdown (IS + OOS)"
    wsSummary.Cells(1, COL_MAX_OOS_DRAWDOWN).value = "Max Drawdown (OOS)"
    wsSummary.Cells(1, COL_AVG_OOS_DRAWDOWN).value = "Avg Drawdown (OOS)"
    wsSummary.Cells(1, COL_MAX_DRAWDOWN_LAST_12_MONTHS).value = "Max Drawdown (Last 12 Months)"
   ' wsSummary.Cells(1, COL_MAX_DRAWDOWN_PERCENT).value = "Max Drawdown %"
    wsSummary.Cells(1, COL_PROFIT_LAST_1_MONTH).value = "Profit Last 1 Month"
    wsSummary.Cells(1, COL_PROFIT_LAST_3_MONTHS).value = "Profit Last 3 Months"
    wsSummary.Cells(1, COL_PROFIT_LAST_6_MONTHS).value = "Profit Last 6 Months"
    wsSummary.Cells(1, COL_PROFIT_LAST_9_MONTHS).value = "Profit Last 9 Months"
    wsSummary.Cells(1, COL_PROFIT_LAST_12_MONTHS).value = "Profit Last 12 Months"
    wsSummary.Cells(1, COL_PROFIT_SINCE_OOS_START).value = "Profit Since OOS Start"
    wsSummary.Cells(1, COL_COUNT_PROFIT_MONTHS).value = "# of " & GetNamedRangeValue("EligibilityGreaterThan") & " months in the last " & GetNamedRangeValue("EligibilityTotalMonths") & " months"
   
    wsSummary.Cells(1, COL_EFFICIENCY_LAST_1_MONTH).value = "Efficiency Last 1 Month"
    wsSummary.Cells(1, COL_EFFICIENCY_LAST_3_MONTHS).value = "Efficiency Last 3 Months"
    wsSummary.Cells(1, COL_EFFICIENCY_LAST_6_MONTHS).value = "Efficiency Last 6 Months"
    wsSummary.Cells(1, COL_EFFICIENCY_LAST_9_MONTHS).value = "Efficiency Last 9 Months"
    wsSummary.Cells(1, COL_EFFICIENCY_LAST_12_MONTHS).value = "Efficiency Last 12 Months"
    wsSummary.Cells(1, COL_EFFICIENCY_SINCE_OOS_START).value = "Efficiency Since OOS Start"
    wsSummary.Cells(1, COL_INCUBATION_STATUS).value = "Incubation Status"
    wsSummary.Cells(1, COL_INCUBATION_DATE).value = "Incubation Passed Date" ' Added heading for Incubation Date
    wsSummary.Cells(1, COL_QUITTING_STATUS).value = "Quitting Status"
    wsSummary.Cells(1, COL_QUITTING_DATE).value = "Quitting Date"
    wsSummary.Cells(1, COL_CURRENT_POSITION).value = "Current Position"
    
    
    wsSummary.Cells(1, COL_WF_IN_OUT).value = "Walkforward In/Out" ' Added heading
    wsSummary.Cells(1, COL_ANCHORED).value = "Anchored/Unachored" ' Added heading
    wsSummary.Cells(1, COL_FITNESS).value = "Fitness Function" ' Added heading
    wsSummary.Cells(1, COL_SESSION).value = "Session" ' Added heading
    wsSummary.Cells(1, COL_OOS_PERIOD).value = "Out of Sample Period"
    wsSummary.Cells(1, COL_DATA2_SYMBOL).value = "Symbol Data 2"
    wsSummary.Cells(1, COL_DATA2_TIMEFRAME).value = "Bar Sizw Data 2"
    wsSummary.Cells(1, COL_DATA3_SYMBOL).value = "Symbol Data 3"
    wsSummary.Cells(1, COL_DATA3_TIMEFRAME).value = "Bar Size Data 3"
    wsSummary.Cells(1, COL_START_DATE).value = "Start of Modelling"
    wsSummary.Cells(1, COL_TRADINGDAYS_IS).value = "Annual Trading Days (IS)"
    wsSummary.Cells(1, COL_TRADINGDAYS_ISOOS).value = "Annual Trading Days (IS + OOS)"
    wsSummary.Cells(1, COL_SHARPE_IS).value = "Daily Sharpe (IS)"
    wsSummary.Cells(1, COL_SHARPE_ISOOS).value = "Daily Sharpe (IS + OOS)"
    
     
    Dim oosPeriodYears As Double
    Dim startdate As Date
    Dim OOSTradeProfit As Double
    Dim OOSTrades As Double
    Dim tempmargin As Double, source As String, choice As String
    Dim allPeriodYears As Double
        ' Variables for processing
    Dim i As Long, lastRow As Long
    Dim isMaxDrawdown As Double, quitPercent As Double, quitDollar As Double
    Dim incubationFlag As Boolean, quitFlag As Boolean
    Dim temp As String, lastRowFolder As Long, counter As Long
    
    ' Assuming incubation parameters and quit parameters are defined in fixed cells
    Min_Incubation_Profit = GetNamedRangeValue("Min_Incubation_Profit")
    Incubation_Period = GetNamedRangeValue("Incubation_Period")
    quitPercent = GetNamedRangeValue("Quit_percent")
    quitDollar = GetNamedRangeValue("Quit_Dollar")
    
    
    ' Find the last row in the Walkforward Details sheet
    lastRow = wsDetails.Cells(wsDetails.rows.count, 1).End(xlUp).row
    lastRowFolder = wsStrategies.Cells(wsStrategies.rows.count, 1).End(xlUp).row
    LastColEquity = wsM2MEquity.Cells(1, wsM2MEquity.Columns.count).End(xlToLeft).column
    ' Loop through each row in the Walkforward Details sheet
    
    
    i = 2 ' Start from the second row
    Do While i <= lastRow
        On Error GoTo ErrorHandler ' Enable error handling for each row

        'symbol 1
        ' Strip the "@" character and remove ".D" if present
        strippedSymbol = wsDetails.Cells(i, FindColumnByHeader(headerArray, "Symbol Data1")).value
        strippedSymbol = Replace(strippedSymbol, "@", "")
        strippedSymbol = Replace(strippedSymbol, "$", "")
        strippedSymbol = Replace(strippedSymbol, ".D", "")
        wsSummary.Cells(i, COL_SYMBOL).value = UCase(strippedSymbol)
        wsSummary.Cells(i, COL_SECTOR).value = FindSectorValue(strippedSymbol)
        
        
        'symbol 2
        strippedSymbol = wsDetails.Cells(i, FindColumnByHeader(headerArray, "Symbol Data2")).value
        strippedSymbol = Replace(strippedSymbol, "@", "")
        strippedSymbol = Replace(strippedSymbol, "$", "")
        strippedSymbol = Replace(strippedSymbol, ".D", "")
        wsSummary.Cells(i, COL_DATA2_SYMBOL).value = UCase(strippedSymbol)
        
        'symbol 3
        strippedSymbol = wsDetails.Cells(i, FindColumnByHeader(headerArray, "Symbol Data3")).value
        strippedSymbol = Replace(strippedSymbol, "@", "")
        strippedSymbol = Replace(strippedSymbol, "$", "")
        strippedSymbol = Replace(strippedSymbol, ".D", "")
        wsSummary.Cells(i, COL_DATA3_SYMBOL).value = UCase(strippedSymbol)
        
        
       ' Get strategy name, symbol, and timeframe
        wsSummary.Cells(i, COL_STRATEGY_NUMBER).value = i - 1
        wsSummary.Cells(i, COL_STRATEGY_NAME).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "Strategy Name")).value
        
        wsSummary.Cells(i, COL_TIMEFRAME).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "Interval Data1")).value
        
        
        
        temp = "n/a"
        temp2 = ""
        On Error Resume Next
        For counter = 2 To lastRowFolder
            If wsSummary.Cells(i, COL_STRATEGY_NAME).value = wsStrategies.Cells(counter, COL_STRAT_STRATEGY_NAME).value Then
                temp = wsStrategies.Cells(counter, COL_STRAT_STATUS).value
                temp2 = wsStrategies.Cells(counter, COL_STRAT_CLOSEDTRADEMC).value
            End If
        Next counter
        On Error GoTo 0
        
        
        On Error Resume Next
        For counter = 2 To lastRowFolder + 10
            If wsSummary.Cells(i, COL_STRATEGY_NAME).value = wsLatestPosition.Cells(counter, 1).value Then
                 wsSummary.Cells(i, COL_CURRENT_POSITION).value = wsLatestPosition.Cells(counter, 2).value
            End If
        Next counter
        On Error GoTo 0
        
        
       
        
        
            ' Exit and show error if the sheet doesn't exist
       If temp = "n/a" Then
          Application.ScreenUpdating = True
          Application.EnableEvents = True
          Application.StatusBar = False
          Call OrderVisibleTabsBasedOnList
          Call GoToControl
          MsgBox "Error: Cannot find " & wsSummary.Cells(i, COL_STRATEGY_NAME).value & " in 'Folder Locations' tab.", vbExclamation
          Exit Sub
       End If
        
       ' If status is "Yes", delete the row and skip to the next iteration
        If GetNamedRangeValue("BuyandHoldinSummary") = "No" And temp = GetNamedRangeValue("BuyandHoldStatus") Then
            wsSummary.rows(i).Delete ' Delete the row if status is "Yes"
            lastRow = lastRow - 1 ' Adjust lastRow since a row was deleted
            GoTo SkipIteration ' Skip the rest of the processing for this row
        End If
            
        wsSummary.Cells(i, COL_STATUS).value = temp
        wsSummary.Cells(i, COL_CLOSEDTRADEMC).value = temp2
        
        'Not reopt required for buy and hold
        If wsSummary.Cells(i, COL_STATUS).value <> GetNamedRangeValue("BuyandHoldStatus") Then
            wsSummary.Cells(i, COL_NEXT_OPT_DATE).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "OOS Estimated Reopt Date")).value
        Else
            wsSummary.Cells(i, COL_NEXT_OPT_DATE).value = ""
        End If
        
        
        If CDate(wsSummary.Cells(i, COL_NEXT_OPT_DATE).value) > 0 And wsSummary.Cells(i, COL_NEXT_OPT_DATE).value <> "" Then
        
               ' Find previous reopt date
            If wsDetails.Cells(i, FindColumnByHeader(headerArray, "OUT Period Type")).value = "Month" Then
                 wsSummary.Cells(i, COL_LAST_OPT_DATE).value = CDate(wsSummary.Cells(i, COL_NEXT_OPT_DATE).value - Round(wsDetails.Cells(i, FindColumnByHeader(headerArray, "OUT Period Length")).value * 30.5, 0))
            ElseIf wsDetails.Cells(i, FindColumnByHeader(headerArray, "OUT Period Type")).value = "Year" Then
                 wsSummary.Cells(i, COL_LAST_OPT_DATE).value = CDate(wsSummary.Cells(i, COL_NEXT_OPT_DATE).value - Round(wsDetails.Cells(i, FindColumnByHeader(headerArray, "OUT Period Length")).value * 365.25, 0))
            ElseIf wsDetails.Cells(i, FindColumnByHeader(headerArray, "OUT Period Type")).value = "Trading Days" Then
                 wsSummary.Cells(i, COL_LAST_OPT_DATE).value = CDate(wsSummary.Cells(i, COL_NEXT_OPT_DATE).value - Round(wsDetails.Cells(i, FindColumnByHeader(headerArray, "OUT Period Length")).value * 365.25 / 252, 0))
            End If
            
        Else
            wsSummary.Cells(i, COL_LAST_OPT_DATE).value = ""
        End If
        
            
        
      
        
        
        wsSummary.Cells(i, COL_WF_IN_OUT).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "IN Period Length")).value & " " & wsDetails.Cells(i, FindColumnByHeader(headerArray, "IN Period Type")).value & "/ " & wsDetails.Cells(i, FindColumnByHeader(headerArray, "OUT Period Length")).value & " " & wsDetails.Cells(i, FindColumnByHeader(headerArray, "OUT Period Type")).value
        wsSummary.Cells(i, COL_ANCHORED).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "Anchored")).value
        wsSummary.Cells(i, COL_FITNESS).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "Fitness Function")).value
        wsSummary.Cells(i, COL_SESSION).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "Session Name Data1")).value
        
        
        wsSummary.Cells(i, COL_DATA2_TIMEFRAME).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "Interval Data2")).value
        
        wsSummary.Cells(i, COL_DATA3_TIMEFRAME).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "Interval Data3")).value
        
        
        
        Dim useCut As Boolean, cutVal As Variant
        useCut = (UCase$(Trim$(Range("PT_UseCutoff").value)) = "YES")
        cutVal = Range("PT_CutoffDate").value
        
        Dim rawBeg As Variant, rawEnd As Variant, newBeg As Variant, newEnd As Variant
        rawBeg = wsDetails.Cells(i, FindColumnByHeader(headerArray, "OOS Begin Date")).value
        rawEnd = wsDetails.Cells(i, FindColumnByHeader(headerArray, "OOS End Date")).value
        
        Call ResolveOOSDates(rawBeg, rawEnd, useCut, cutVal, newBeg, newEnd)
        
        wsSummary.Cells(i, COL_OOS_BEGIN_DATE).value = newBeg
        wsSummary.Cells(i, COL_LAST_DATE_ON_FILE).value = newEnd  ' already adjusted per rules
                
        ' OOS Period calculation
        If ((IsDate(newBeg) And _
           IsDate(newEnd)) And _
           newBeg <> newEnd) Then
            oosPeriodYears = DateDiff("d", newBeg, newEnd) / 365.25
        Else
            oosPeriodYears = 0 ' Set default if no valid dates found
        End If

        ' Find start date of strategy
        
        If Not (IsDate(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Begin Date")).value)) Then
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Application.StatusBar = False
            Call OrderVisibleTabsBasedOnList
            Call GoToControl
            MsgBox "Invalid date in 'IS Begin Date' for row " & i, vbExclamation
            Exit Sub
        End If
        
        If wsDetails.Cells(i, FindColumnByHeader(headerArray, "IN Period Type")).value = "Month" Then
            startdate = CDate(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Begin Date")).value + _
                       Round(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IN Period Length")).value * 30.5, 0))
        ElseIf wsDetails.Cells(i, FindColumnByHeader(headerArray, "IN Period Type")).value = "Year" Then
             startdate = CDate(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Begin Date")).value + _
                       Round(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IN Period Length")).value * 365.25, 0))
        ElseIf wsDetails.Cells(i, FindColumnByHeader(headerArray, "IN Period Type")).value = "Trading Days" Then
            startdate = CDate(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Begin Date")).value + _
                       Round(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IN Period Length")).value * 365.25 / 252, 0))
        End If
        
        wsSummary.Cells(i, COL_START_DATE).value = startdate
        
        ' Calculate all trading period
        If IsDate(startdate) And IsDate(newEnd) Then
            allPeriodYears = DateDiff("d", startdate, newEnd) / 365.25
        Else
            allPeriodYears = 1 ' Set default if dates are invalid
        End If

        ' Populate other data fields
        wsSummary.Cells(i, COL_EXPECTED_ANNUAL_PROFIT).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Annualized Net Profit")).value
        
        wsSummary.Cells(i, COL_ACTUAL_ANNUAL_PROFIT).value = IIf(oosPeriodYears = 0, 0, wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS/OOS Change in Net Profit")).value / (oosPeriodYears + 0.000001))

        'lawman edit
        If wsSummary.Cells(i, COL_EXPECTED_ANNUAL_PROFIT).value < 0.001 And wsSummary.Cells(i, COL_EXPECTED_ANNUAL_PROFIT).value > -0.001 Then
            wsSummary.Cells(i, COL_RETURN_EFFICIENCY).value = 0
        Else
            wsSummary.Cells(i, COL_RETURN_EFFICIENCY).value = wsSummary.Cells(i, COL_ACTUAL_ANNUAL_PROFIT).value / wsSummary.Cells(i, COL_EXPECTED_ANNUAL_PROFIT).value   ' Return efficiency
        End If
        
        wsSummary.Cells(i, COL_IS_MONTE_CARLO).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Monte Carlo")).value
        wsSummary.Cells(i, COL_IS_OOS_MONTE_CARLO).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Monte Carlo")).value

    
        wsSummary.Cells(i, COL_BACKTEST_WINRATE).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Percent Trades Profitable")).value
       
        
        ' OOS trades
        OOSTradeProfit = wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Trades Profitable")).value - wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Trades Profitable")).value
        OOSTrades = wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS/OOS Change in Total Trades")).value

        wsSummary.Cells(i, COL_OOS_WINRATE).value = IIf(OOSTrades = 0, 0, OOSTradeProfit / (OOSTrades + 0.000001)) ' Avoid division by zero

        wsSummary.Cells(i, COL_OVERALL_WINRATE).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Percent Trades Profitable")).value
        
        wsSummary.Cells(i, COL_TRADES_PER_YEAR).value = Round(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Total Trades")).value / (allPeriodYears + 0.000001), 0)
        wsSummary.Cells(i, COL_PERCENT_TIME_IN_MARKET).value = wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Percent Time In Market")).value
        wsSummary.Cells(i, COL_AVG_TRADE_LENGTH).value = IIf(wsSummary.Cells(i, COL_TRADES_PER_YEAR).value = 0, 0, (265 * wsSummary.Cells(i, COL_PERCENT_TIME_IN_MARKET).value) / (wsSummary.Cells(i, COL_TRADES_PER_YEAR).value + 0.000001))
        
        
        wsSummary.Cells(i, COL_AVG_IS_OOS_TRADE).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Avg Trade")).value)
        
        wsSummary.Cells(i, COL_AVG_PROFIT_IS_OOS_TRADE).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Avg Profitable Trade")).value)
        wsSummary.Cells(i, COL_AVG_LOSS_IS_OOS_TRADE).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Avg Unprofitable Trade")).value)
        wsSummary.Cells(i, COL_LARGEST_WIN_IS_OOS_TRADE).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Largest Profitable Trade")).value)
        wsSummary.Cells(i, COL_LARGEST_LOSS_IS_OOS_TRADE).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Largest Unprofitable Trade")).value)
       

        
        
        wsSummary.Cells(i, COL_TOTAL_IS_PROFIT).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Net Profit")).value)
        wsSummary.Cells(i, COL_TOTAL_IS_OSS_PROFIT).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Net Profit")).value)
        wsSummary.Cells(i, COL_ANNUALIZED_NET_PROFIT_IS_OOS).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Annualized Net Profit")).value)

        
        wsSummary.Cells(i, COL_WORST_BACKTEST_DRAWDOWN).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Max DD")).value)
        wsSummary.Cells(i, COL_WORST_IS_OOS_DRAWDOWN).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Max DD")).value)
        wsSummary.Cells(i, COL_AVG_BACKTEST_DRAWDOWN).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Avg DD")).value)
        wsSummary.Cells(i, COL_AVG_IS_OOS_DRAWDOWN).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Avg DD")).value)


        wsSummary.Cells(i, COL_TRADINGDAYS_IS).value = Round(Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Total Trading Days")).value) / Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Total Calendar Days")).value + 0.0000001) * 365.25, 0)
        wsSummary.Cells(i, COL_TRADINGDAYS_ISOOS).value = IIf(Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Total Calendar Days")).value) = 0, 0, Round(Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Total Trading Days")).value) / Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Total Calendar Days")).value + 0.0000001) * 365.25, 0))
        wsSummary.Cells(i, COL_SHARPE_IS).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS Sharpe Ratio")).value)
        wsSummary.Cells(i, COL_SHARPE_ISOOS).value = Abs(wsDetails.Cells(i, FindColumnByHeader(headerArray, "IS+OOS Sharpe Ratio")).value)

        wsSummary.Cells(i, COL_IS_ANNUAL_SD_IS).value = Abs(wsSummary.Cells(i, COL_EXPECTED_ANNUAL_PROFIT).value / ((Sqr(365.25) * wsSummary.Cells(i, COL_SHARPE_IS).value) + 0.00000001))
        wsSummary.Cells(i, COL_IS_ANNUAL_SD_ISOOS).value = Abs((wsSummary.Cells(i, COL_AVG_IS_OOS_TRADE).value * wsSummary.Cells(i, COL_TRADES_PER_YEAR).value) / ((Sqr(365.25) * wsSummary.Cells(i, COL_SHARPE_IS).value) + 0.00000001))
   
        ' Get the strategy name and find its corresponding column in wsM2MEquity
        Dim strategyName As String
        strategyName = wsSummary.Cells(i, COL_STRATEGY_NAME).value ' Assuming strategy name is in column 3

        Dim strategyColumn As Long
        
        strategyColumn = -99
        On Error Resume Next
        
            For counter = 2 To LastColEquity
                If strategyName = wsM2MEquity.Cells(1, counter).value Then strategyColumn = counter
            Next counter
            
        On Error GoTo 0
        
            ' Exit and show error if the sheet doesn't exist
        If strategyColumn = -99 Then
           Application.ScreenUpdating = True
           Application.EnableEvents = True
           Application.StatusBar = False
           Call OrderVisibleTabsBasedOnList
           Call GoToControl
           MsgBox "Error: Cannot find " & strategyName & " in 'DailyM2MEquity' tab.", vbExclamation
           Exit Sub
        End If
         
        
        ' Ensure the strategy column is valid
        If Not IsError(strategyColumn) Then
            ' Get OOS Begin and End Dates for the strategy
            Dim OOSBeginDate As Date
            Dim OOSEndDate As Date
            OOSBeginDate = newBeg ' OOS Begin Date
            OOSEndDate = newEnd   ' OOS End Date
            
            ' Get IS Annualized Profit
            Dim ISAnnualizedProfit As Double
            ISAnnualizedProfit = wsSummary.Cells(i, COL_EXPECTED_ANNUAL_PROFIT).value ' IS Annualized Profit (column 11)
            
            ' Call the function to calculate profit, drawdown, incubation, and quit flags
            Call CalculateProfitAndDrawdown(wsM2MEquity, strategyColumn, OOSBeginDate, OOSEndDate, wsSummary, i, ISAnnualizedProfit, wsSummary.Cells(i, COL_IS_ANNUAL_SD_IS).value)
        End If
        
        If IsDate(OOSBeginDate) And IsDate(OOSEndDate) Then
            Dim yearsDiff As Long
            Dim monthsDiff As Long
            Dim totalMonths As Long

            ' Calculate total months between the two dates
            totalMonths = DateDiff("m", OOSBeginDate, OOSEndDate)
            yearsDiff = totalMonths \ 12
            monthsDiff = totalMonths Mod 12
            
            ' Format the OOS period as "X years, Y months"
            wsSummary.Cells(i, COL_OOS_PERIOD).value = yearsDiff & " years, " & monthsDiff & " months"
        Else
            wsSummary.Cells(i, COL_OOS_PERIOD).value = "N/A" ' Handle invalid dates
        End If
        
        
        wsSummary.Cells(i, COL_R_DD_12MONTH).value = IIf(Abs(wsSummary.Cells(i, COL_MAX_DRAWDOWN_LAST_12_MONTHS).value) < 10, 10, wsSummary.Cells(i, COL_PROFIT_LAST_12_MONTHS).value / (wsSummary.Cells(i, COL_MAX_DRAWDOWN_LAST_12_MONTHS).value + 0.0001))
        wsSummary.Cells(i, COL_R_DD_OOS).value = IIf(Abs(wsSummary.Cells(i, COL_MAX_OOS_DRAWDOWN).value) < 10, 10, wsSummary.Cells(i, COL_PROFIT_SINCE_OOS_START).value / (wsSummary.Cells(i, COL_MAX_OOS_DRAWDOWN).value + 0.0001))
       
        ' Calculate profit factors and related metrics from Long_Trades and Short_Trades
        Call CalculateTradeProfitFactors(wsSummary, i, wsSummary.Cells(i, COL_STRATEGY_NAME).value)
        
        ' Determine strategy type
        wsSummary.Cells(i, COL_LONG_SHORT).value = DetermineStrategyType(wsSummary.Cells(i, COL_GROSS_LONG_PROFIT).value, wsSummary.Cells(i, COL_GROSS_LONG_LOSS).value, wsSummary.Cells(i, COL_GROSS_SHORT_PROFIT).value, wsSummary.Cells(i, COL_GROSS_SHORT_LOSS).value)
      
        
        
            
        
        Call LookupMarginRequirements(wsSummary, wsDetails, i, wsSummary.Cells(i, COL_SYMBOL).value, headerArray)
        
        If GetNamedRangeValue("Run_MC") = "Yes" Then
            Call RunMonteCarloSimulation(i)
        Else
            Dim notional_capital As Double
            notional_capital = GetNamedRangeValue("MC_StartingEquity") * wsSummary.Cells(i, COL_MARGIN).value
            wsSummary.Cells(i, COL_NOTIONAL_CAPITAL).value = notional_capital
            wsSummary.Cells(i, COL_EXPECTED_ANNUAL_RETURN).value = wsSummary.Cells(i, COL_EXPECTED_ANNUAL_PROFIT).value / (notional_capital + 0.001)
            wsSummary.Cells(i, COL_ACTUAL_ANNUAL_RETURN).value = wsSummary.Cells(i, COL_ACTUAL_ANNUAL_PROFIT).value / (notional_capital + 0.001)
        End If
        
        
        Call EligibilityTracking(wsSummary, i)
        

        On Error GoTo 0 ' Disable error handling after successful row processing
    
     Application.StatusBar = "Summary Tracking Calculating: " & Format(i / lastRow, "0%") & " completed"
    
SkipIteration:
        i = i + 1 ' Move to the next row
    Loop
        
    Application.StatusBar = "Ranking..."
    
    Call RankStrategiesInSummary
        
    Application.StatusBar = "Reordering..."
        
    Call ReorderSummaryTab

    Application.StatusBar = "Formatting..."
    Call FormatSummaryTable

    Call ApplyConditionalFormatting

    Application.StatusBar = "Sparklines..."
    Call AggregateWeeklyProfitsAndAddSparklines

    
    Dim validationRange As Range
    Set validationRange = wsSummary.Range(wsSummary.Cells(2, COL_STATUS), wsSummary.Cells(lastRow, COL_STATUS))
    Call CreateStatusDropdown(validationRange)
    


    Application.StatusBar = "Just more buttons... Don't Panic!"
    Call SetupButtonsAndStrategyTabCreation
    
    Call SetupButtonsforCodeTabCreation
    
    Call SetupButtonsforMC

    
    'If resetStrategiestab = "No" Then Call OrderVisibleTabsBasedOnList
    
    

    ' Hide columns from startCol to the last column
    wsSummary.Range(wsSummary.Columns(COL_WF_IN_OUT), wsSummary.Columns(COL_SHARPE_ISOOS)).EntireColumn.Hidden = True
   
    

    
    wsSummary.Cells(1, COL_STRATEGY_NAME).VerticalAlignment = xlBottom
    

    Call CreateSummaryButtons(wsSummary, COL_STRATEGY_NAME, "Summary")
     

    
    ' Apply AutoFilter to the Summary Table
    lastRow = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row
    ' Apply AutoFilter to the table
    wsSummary.Range(wsSummary.Cells(1, 1), wsSummary.Cells(lastRow, COL_SHARPE_ISOOS)).AutoFilter
    
    
    If resetStrategiestab = "No" Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Application.EnableEvents = True
        MsgBox "Strategy Summary has been updated successfully!"
        
        
    End If
    
    
    
    
    
    
   
    
    
    
    
    Exit Sub

ErrorHandler:
    'MsgBox "Error occurred in row " & i & ": " & Err.Description
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    Resume Next ' Continue to next row even after error

End Sub



' Function to create an array of headers and column numbers from the Walkforward Details sheet
Function CreateHeaderArray(ws As Worksheet) As Variant
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).column
    
    ' Create an array to hold the header names and their corresponding column numbers
    Dim headerArray() As Variant
    ReDim headerArray(1 To lastCol, 1 To 2)
    
    ' Loop through the first row to store header names and their column numbers in the array
    Dim colNum As Long
    For colNum = 1 To lastCol
        headerArray(colNum, 1) = ws.Cells(1, colNum).value  ' Store the header name
        headerArray(colNum, 2) = colNum  ' Store the column number
    Next colNum
    
    CreateHeaderArray = headerArray
End Function

' Function to search for a header in the array and return the corresponding column number
Function FindColumnByHeader(headerArray As Variant, headerName As String) As Long
    Dim i As Long
    For i = LBound(headerArray, 1) To UBound(headerArray, 1)
        If headerArray(i, 1) = headerName Then
            FindColumnByHeader = headerArray(i, 2)
            Exit Function
        End If
    Next i
    FindColumnByHeader = 0 ' Return 0 if the header is not found
End Function
    

Sub EligibilityTracking(wsSummary As Worksheet, i As Long)

    Dim ProfitMonth1 As String
    Dim ProfitMonth3 As String
    Dim ProfitMonth6 As String
    Dim LossMonth1 As String
    Dim LossMonth3 As String
    Dim LossMonth6 As String
    Dim ProfitMonth3or6 As String
    Dim ProfitMonth9 As String
    Dim ProfitMonth12 As String
    Dim ProfitMonthOOS As String
    
    Dim EfficiencyMonth1 As String
    Dim EfficiencyMonth3 As String
    Dim EfficiencyMonth6 As String
    Dim EfficiencyMonth1Loss As String
    Dim EfficiencyMonth3Loss As String
    Dim EfficiencyMonth6Loss As String
    Dim EfficiencyMonth9 As String
    Dim EfficiencyMonth12 As String
    Dim EfficiencyMonthOOS As String
    Dim efficiencyRatio As Double
    
    Dim IncubationStatus As String
    Dim QuittingStatus As String
    
    Dim AddUserFilter As String
    Dim AddUserFilterValue As String
    Dim AddUserFilterMin As Double
    
    
    
    Dim PositiveMonthsCheck As String, MinPositiveMonths As Long
        
    Dim EligibilityStatus As String
    
    ProfitMonth1 = GetNamedRangeValue("Eligibility1MonthProfit")
    ProfitMonth3 = GetNamedRangeValue("Eligibility3MonthProfit")
    ProfitMonth6 = GetNamedRangeValue("Eligibility6MonthProfit")
    LossMonth1 = GetNamedRangeValue("Eligibility1MonthLosses")
    LossMonth3 = GetNamedRangeValue("Eligibility3MonthLosses")
    LossMonth6 = GetNamedRangeValue("Eligibility6MonthLosses")
    ProfitMonth3or6 = GetNamedRangeValue("Eligibility3or6MonthProfit")
    
    ProfitMonth9 = GetNamedRangeValue("Eligibility9MonthProfit")
    ProfitMonth12 = GetNamedRangeValue("Eligibility12MonthProfit")
    ProfitMonthOOS = GetNamedRangeValue("EligibilityOOSMonthProfit")
    EfficiencyMonth1 = GetNamedRangeValue("Eligibility1MonthEff")
    EfficiencyMonth3 = GetNamedRangeValue("Eligibility3MonthEff")
    EfficiencyMonth6 = GetNamedRangeValue("Eligibility6MonthEff")
    EfficiencyMonth1Loss = GetNamedRangeValue("Eligibility1MonthEffLosses")
    EfficiencyMonth3Loss = GetNamedRangeValue("Eligibility3MonthEffLosses")
    EfficiencyMonth6Loss = GetNamedRangeValue("Eligibility6MonthEffLosses")
    EfficiencyMonth9 = GetNamedRangeValue("Eligibility9MonthEff")
    EfficiencyMonth12 = GetNamedRangeValue("Eligibility12MonthEff")
    EfficiencyMonthOOS = GetNamedRangeValue("EligibilityOOSMonthEff")
    efficiencyRatio = GetNamedRangeValue("EfficiencyRatio")
    IncubationStatus = GetNamedRangeValue("EligibilityIncubation")
    QuittingStatus = GetNamedRangeValue("EligibilityQuitting")
    
    PositiveMonthsCheck = GetNamedRangeValue("EligibilityCountMonthlyProfits")
    MinPositiveMonths = GetNamedRangeValue("EligibilityMinimumMonths")
    
    AddUserFilter = GetNamedRangeValue("AdditionalUserFilter")
    AddUserFilterValue = GetNamedRangeValue("AdditionalUserFilterValue")
    AdditionalUserFilterMin = GetNamedRangeValue("AdditionalUserFilterMin")
    
    
    
    EligibilityStatus = "Yes"
    
    Dim found As Integer
    found = 0
    
    
        ' In your eligibility check code
    If AddUserFilter = "Yes" Then
        For j = 1 To COL_SHARPE_ISOOS
            If wsSummary.Cells(1, j).value = AddUserFilterValue Then
                found = 1
                Exit For
            End If
        Next j
        
        If found = 1 Then
            ' Check if cell contains text
            If IsEmpty(wsSummary.Cells(i, j)) Or wsSummary.Cells(i, j).value = "" Then
                EligibilityStatus = "No"
            ElseIf VarType(wsSummary.Cells(i, j).value) = vbString Then
                ' Text comparison - you can customize this based on your needs
                If wsSummary.Cells(i, j).value <> AdditionalUserFilterMin Then
                    EligibilityStatus = "No"
                End If
            Else
                ' Numeric comparison
                If wsSummary.Cells(i, j).value < AdditionalUserFilterMin Then
                    EligibilityStatus = "No"
                End If
            End If
        End If
    End If
        
    
    ' Guard all numeric comparisons with IsNumeric: profit/efficiency cells are set to "" when a
    ' strategy lacks sufficient OOS history.  Comparing "" to a number raises Type Mismatch (Error 13)
    ' in VBA, which was silently swallowing the entire EligibilityTracking call via the outer
    ' On Error GoTo ErrorHandler / Resume Next, leaving the Eligibility column blank.
    ' Blank cells are treated as "no data yet" — the criterion is skipped (not disqualifying).

    Dim v1M As Variant, v3M As Variant, v6M As Variant, v9M As Variant, v12M As Variant, vOOS As Variant
    Dim e1M As Variant, e3M As Variant, e6M As Variant, e9M As Variant, e12M As Variant, eOOS As Variant
    v1M  = wsSummary.Cells(i, COL_PROFIT_LAST_1_MONTH).value
    v3M  = wsSummary.Cells(i, COL_PROFIT_LAST_3_MONTHS).value
    v6M  = wsSummary.Cells(i, COL_PROFIT_LAST_6_MONTHS).value
    v9M  = wsSummary.Cells(i, COL_PROFIT_LAST_9_MONTHS).value
    v12M = wsSummary.Cells(i, COL_PROFIT_LAST_12_MONTHS).value
    vOOS = wsSummary.Cells(i, COL_PROFIT_SINCE_OOS_START).value
    e1M  = wsSummary.Cells(i, COL_EFFICIENCY_LAST_1_MONTH).value
    e3M  = wsSummary.Cells(i, COL_EFFICIENCY_LAST_3_MONTHS).value
    e6M  = wsSummary.Cells(i, COL_EFFICIENCY_LAST_6_MONTHS).value
    e9M  = wsSummary.Cells(i, COL_EFFICIENCY_LAST_9_MONTHS).value
    e12M = wsSummary.Cells(i, COL_EFFICIENCY_LAST_12_MONTHS).value
    eOOS = wsSummary.Cells(i, COL_EFFICIENCY_SINCE_OOS_START).value

    If ProfitMonth1   = "Yes" And IsNumeric(v1M)  And CDbl(v1M)  <= 0 Then EligibilityStatus = "No"
    If ProfitMonth3   = "Yes" And IsNumeric(v3M)  And CDbl(v3M)  <= 0 Then EligibilityStatus = "No"
    If ProfitMonth6   = "Yes" And IsNumeric(v6M)  And CDbl(v6M)  <= 0 Then EligibilityStatus = "No"
    If ProfitMonth3or6 = "Yes" And IsNumeric(v3M) And IsNumeric(v6M) And CDbl(v3M) <= 0 And CDbl(v6M) < 0 Then EligibilityStatus = "No"
    If ProfitMonth9   = "Yes" And IsNumeric(v9M)  And CDbl(v9M)  <= 0 Then EligibilityStatus = "No"
    If ProfitMonth12  = "Yes" And IsNumeric(v12M) And CDbl(v12M) <= 0 Then EligibilityStatus = "No"
    If ProfitMonthOOS = "Yes" And IsNumeric(vOOS) And CDbl(vOOS) <= 0 Then EligibilityStatus = "No"

    If LossMonth1 = "Yes" And IsNumeric(v1M) And CDbl(v1M) >= 0 Then EligibilityStatus = "No"
    If LossMonth3 = "Yes" And IsNumeric(v3M) And CDbl(v3M) >= 0 Then EligibilityStatus = "No"
    If LossMonth6 = "Yes" And IsNumeric(v3M) And CDbl(v3M) >= 0 Then EligibilityStatus = "No"  ' Note: intentionally checks 3M (matches original logic)

    If EfficiencyMonth1   = "Yes" And IsNumeric(e1M)  And CDbl(e1M)  <= efficiencyRatio Then EligibilityStatus = "No"
    If EfficiencyMonth3   = "Yes" And IsNumeric(e3M)  And CDbl(e3M)  <= efficiencyRatio Then EligibilityStatus = "No"
    If EfficiencyMonth6   = "Yes" And IsNumeric(e6M)  And CDbl(e6M)  <= efficiencyRatio Then EligibilityStatus = "No"
    If EfficiencyMonth9   = "Yes" And IsNumeric(e9M)  And CDbl(e9M)  <= efficiencyRatio Then EligibilityStatus = "No"
    If EfficiencyMonth12  = "Yes" And IsNumeric(e12M) And CDbl(e12M) <= efficiencyRatio Then EligibilityStatus = "No"
    If EfficiencyMonthOOS = "Yes" And IsNumeric(eOOS) And CDbl(eOOS) <= efficiencyRatio Then EligibilityStatus = "No"

    If EfficiencyMonth1Loss = "Yes" And IsNumeric(e1M) And CDbl(e1M) >= efficiencyRatio Then EligibilityStatus = "No"
    If EfficiencyMonth3Loss = "Yes" And IsNumeric(e3M) And CDbl(e3M) >= efficiencyRatio Then EligibilityStatus = "No"
    If EfficiencyMonth6Loss = "Yes" And IsNumeric(e6M) And CDbl(e6M) >= efficiencyRatio Then EligibilityStatus = "No"
        
        
    If PositiveMonthsCheck = "Yes" And wsSummary.Cells(i, COL_COUNT_PROFIT_MONTHS).value < MinPositiveMonths Then EligibilityStatus = "No"
        
    If IncubationStatus = "Yes" And wsSummary.Cells(i, COL_INCUBATION_STATUS).value <> "Passed" Then EligibilityStatus = "No"
    If QuittingStatus = "Yes" And wsSummary.Cells(i, COL_QUITTING_STATUS).value = "Quit" Then EligibilityStatus = "No"
        
        
        
        
    wsSummary.Cells(i, COL_ELIGIBILITY).value = EligibilityStatus



End Sub



Sub CalculateProfitAndDrawdown(wsM2MEquity As Worksheet, strategyColumn As Long, OOSBeginDate As Date, OOSEndDate As Date, wsSummary As Worksheet, i As Long, ISAnnualizedProfit As Double, sd As Double)
    Dim lastRow As Long
   
    ' Variables for profit calculations
    Dim profitLast1Month As Variant, profitLast3Months As Variant, profitLast6Months As Variant
    Dim profitLast9Months As Variant, profitLast12Months As Variant, profitSinceOOSStart As Double
    Dim peakEquityLast12Months As Double, currentEquityLast12Months As Double
    Dim currentDrawdownLast12Months As Double, maxDrawdownLast12Months As Double
    Dim monthlyProfits As Collection
    Dim currentMonthProfit As Double
    Dim currentMonth As Long
    Dim currentYear As Long
    Dim lastMonth As Long
    Dim lastYear As Long
    Dim dailyProfit As Double
    Dim maxdrawdownpercent As Double
    Dim drawdownpercent As Double
    Dim quitting_method As String
    Dim SD_Multiple As Double
    Dim quitPercent As Double, quitDollar As Double, quittingPoint As Double
    Dim isMaxDrawdown As Double
    Dim quit_status As String
    Dim SDquitingpoint As Double
    Dim recoveryPoint As Double
    
    profitLast1Month = 0: profitLast3Months = 0: profitLast6Months = 0
    profitLast1Loss = 0: profitLast3Losses = 0: profitLast6Losses = 0
    profitLast9Months = 0: profitLast12Months = 0: profitSinceOOSStart = 0
    peakEquityLast12Months = 0: currentEquityLast12Months = 0
    currentDrawdownLast12Months = 0: maxDrawdownLast12Months = 0
    maxdrawdownpercent = 0
    ISAnnualizedProfit = IIf(ISAnnualizedProfit < 0, 0, ISAnnualizedProfit)
    
    lastRow = EndRowByCutoffSimple(wsM2MEquity, 1)

    Dim IncubationPeriod As Long
    Dim TotalMonthlyCount As Long
    
    Dim daysThreshold As Long
    daysThreshold = GetNamedRangeValue("EligibilityDaysThreshold")
    
    ' Initialize the collection before your main loop
    Set monthlyProfits = New Collection
    
    ' Variables for drawdown calculations
    Dim maxOOSDrawdown As Double, avgOOSDrawdown As Double, currentDrawdown As Double
    Dim peakEquity As Double, currentEquity As Double
    maxOOSDrawdown = 0: avgOOSDrawdown = 0: currentDrawdown = 0
    peakEquity = 0: currentEquity = 0

    ' Define date ranges for different periods
    Dim lastMonthDate As Date, last3MonthsDate As Date, last6MonthsDate As Date
    Dim last9MonthsDate As Date, last12MonthsDate As Date
    
    If daysThreshold > 0 Then
        ' Get the first day of the current month of OOSEndDate
        Dim currentMonthStart As Date
        currentMonthStart = DateSerial(Year(OOSEndDate), Month(OOSEndDate), 1)
        
        ' Calculate days in current month
        Dim daysInCurrentMonth As Long
        daysInCurrentMonth = DateDiff("d", currentMonthStart, OOSEndDate) + 1
        
        ' If we have less than threshold days in current month, use previous month as end point
        Dim effectiveEndDate As Date
        If daysInCurrentMonth < daysThreshold Then
            effectiveEndDate = DateSerial(Year(OOSEndDate), Month(OOSEndDate), 1) - 1
            ' Start dates when using previous month
            lastMonthDate = DateSerial(Year(effectiveEndDate), Month(effectiveEndDate), 1)
            last3MonthsDate = DateSerial(Year(effectiveEndDate), Month(effectiveEndDate) - 2, 1)
            last6MonthsDate = DateSerial(Year(effectiveEndDate), Month(effectiveEndDate) - 5, 1)
            last9MonthsDate = DateSerial(Year(effectiveEndDate), Month(effectiveEndDate) - 8, 1)
            last12MonthsDate = DateSerial(Year(effectiveEndDate), Month(effectiveEndDate) - 11, 1)
        Else
            effectiveEndDate = OOSEndDate
            ' Start dates when using current month
            lastMonthDate = currentMonthStart  ' Just current month from start
            last3MonthsDate = DateSerial(Year(OOSEndDate), Month(OOSEndDate) - 2, 1)  ' Current plus 2 previous
            last6MonthsDate = DateSerial(Year(OOSEndDate), Month(OOSEndDate) - 5, 1)  ' Current plus 5 previous
            last9MonthsDate = DateSerial(Year(OOSEndDate), Month(OOSEndDate) - 8, 1)  ' Current plus 8 previous
            last12MonthsDate = DateSerial(Year(OOSEndDate), Month(OOSEndDate) - 11, 1) ' Current plus 11 previous
        End If
    Else
        ' Use original rolling period logic
        effectiveEndDate = OOSEndDate
        lastMonthDate = DateAdd("m", -1, OOSEndDate) + 1
        last3MonthsDate = DateAdd("m", -3, OOSEndDate) + 1
        last6MonthsDate = DateAdd("m", -6, OOSEndDate) + 1
        last9MonthsDate = DateAdd("m", -9, OOSEndDate) + 1
        last12MonthsDate = DateAdd("m", -12, OOSEndDate) + 1
    End If
    
    
     


    TotalMonthlyCount = GetNamedRangeValue("EligibilityTotalMonths")
    IncubationPeriod = GetNamedRangeValue("Incubation_Period")
    minIncubationProfit = GetNamedRangeValue("Min_Incubation_Profit") ' Assuming Min Incubation Profit is in fixed cell
    
    quitting_method = GetNamedRangeValue("Quitting_Method")
    SD_Multiple = GetNamedRangeValue("Quitting_SD_Multiple")
    tradingdays = wsSummary.Cells(i, COL_TRADINGDAYS_IS).value
    isMaxDrawdown = wsSummary.Cells(i, COL_WORST_BACKTEST_DRAWDOWN).value
    quitPercent = GetNamedRangeValue("Quit_percent")
    quitDollar = GetNamedRangeValue("Quit_Dollar")
    SDquitingpoint = 0
    quittingPoint = Application.WorksheetFunction.Min(quitDollar, quitPercent * Abs(isMaxDrawdown))
    
    isMaxDrawdown = wsSummary.Cells(i, COL_WORST_BACKTEST_DRAWDOWN).value
    
    
    quit_status = "Continue"
    
    ' Find OOSBeginRow and OOSEndRow manually
    Dim OOSBeginRow As Long, OOSEndRow As Long
    OOSBeginRow = 0
    OOSEndRow = 0
    Set wsM2MEquity = ThisWorkbook.Sheets("DailyM2MEquity")

    Dim row As Long
    For row = 2 To lastRow
        Dim currentdate As Date
        currentdate = wsM2MEquity.Cells(row, 1).value * 1

        If OOSBeginRow = 0 And currentdate >= OOSBeginDate Then
            OOSBeginRow = row
        End If

        If currentdate <= OOSEndDate Then
            OOSEndRow = row
        ElseIf currentdate > OOSEndDate Then
            Exit For ' No need to check further
        End If
    Next row

    ' Ensure we found valid rows.
    ' NOTE: OOSBeginRow = OOSEndRow is intentionally allowed — Buy & Hold strategies
    ' often have a single-day OOS window or start on the very first data row.
    If OOSBeginRow = 0 Or OOSEndRow = 0 Then Exit Sub

    ' Loop through the rows within the OOS period
    For row = OOSBeginRow To OOSEndRow
        currentdate = wsM2MEquity.Cells(row, 1).value
        
        dailyProfit = wsM2MEquity.Cells(row, strategyColumn).value

      ' Get current month and year
       currentMonth = Month(currentdate)
       currentYear = Year(currentdate)

        ' If this is a new month, store the previous month's profit
        If currentMonth <> lastMonth Or currentYear <> lastYear Then
            If lastMonth <> 0 Then ' Skip the first iteration
                ' Store profit with a unique key (YYYYMM format)
                monthlyProfits.Add currentMonthProfit, Format(lastYear, "0000") & Format(lastMonth, "00")
            End If
            ' Reset for new month
            currentMonthProfit = dailyProfit
        Else
            ' Add to current month's profit
            currentMonthProfit = currentMonthProfit + dailyProfit
        End If
        
        lastMonth = currentMonth
        lastYear = currentYear


        ' Only calculate profits if the date is within our effective date range
        If currentdate <= effectiveEndDate Then
            ' Add profit to the appropriate period
            If currentdate >= lastMonthDate And lastMonthDate >= OOSBeginDate Then
                profitLast1Month = profitLast1Month + dailyProfit
            ElseIf lastMonthDate < OOSBeginDate Then
                profitLast1Month = ""
            End If
            If currentdate >= last3MonthsDate And last3MonthsDate >= OOSBeginDate Then
                profitLast3Months = profitLast3Months + dailyProfit
            ElseIf last3MonthsDate < OOSBeginDate Then
                profitLast3Months = ""
            End If
            If currentdate >= last6MonthsDate And last6MonthsDate >= OOSBeginDate Then
                profitLast6Months = profitLast6Months + dailyProfit
            ElseIf last6MonthsDate < OOSBeginDate Then
                profitLast6Months = ""
            End If
            If currentdate >= last9MonthsDate And last9MonthsDate >= OOSBeginDate Then
                profitLast9Months = profitLast9Months + dailyProfit
            ElseIf last9MonthsDate < OOSBeginDate Then
                profitLast9Months = ""
            End If
            If currentdate >= last12MonthsDate And last12MonthsDate >= OOSBeginDate Then
                profitLast12Months = profitLast12Months + dailyProfit
            ElseIf last12MonthsDate < OOSBeginDate Then
                profitLast12Months = ""
            End If
        End If
        
        ' Incubation logic: Accumulate profits and check if it exceeds the criteria
        incubationProfit = incubationProfit + dailyProfit
        If (row - OOSBeginRow + 1) >= IncubationPeriod * 30.5 Then
            If incubationProfit >= (ISAnnualizedProfit / 365.25) * (row - OOSBeginRow) * minIncubationProfit And hasPassedIncubation = False Then
             
                hasPassedIncubation = True
                incubationPassedDate = currentdate ' Store the date when incubation is passed
                wsSummary.Cells(i, COL_INCUBATION_STATUS).value = "Passed"
                wsSummary.Cells(i, COL_INCUBATION_DATE).value = incubationPassedDate ' Record the passed date
 
            End If
        End If
        
        
        ' Always add to profit since OOS start
        profitSinceOOSStart = profitSinceOOSStart + dailyProfit
        
        
        

        ' Update drawdown calculations
        currentEquity = currentEquity + dailyProfit
        If currentEquity > peakEquity Then
            peakEquity = currentEquity ' New peak
        End If

        ' Calculate drawdown from peak
        currentDrawdown = peakEquity - currentEquity
        If currentDrawdown > maxOOSDrawdown Then
            maxOOSDrawdown = currentDrawdown ' Update max drawdown
        End If
        avgOOSDrawdown = avgOOSDrawdown + currentDrawdown
        
        
           
        If drawdownpercent > maxdrawdownpercent Then
            maxdrawdownpercent = drawdownpercent ' Update max drawdown
        End If
    
        If quitting_method = "Drawdown" Then
            quitEquity = peakEquity - quittingPoint
        ElseIf quitting_method = "Standard Deviation" Then
            quitEquity = ((ISAnnualizedProfit / 365.25) * (row - OOSBeginRow) - Sqr(row - OOSBeginRow) * (sd / Sqr(365.25)) * SD_Multiple)
        Else
            quit_status = "N/A"
        End If
    
        If peakEquity < (ISAnnualizedProfit / 365.25) * (row - OOSBeginRow) * minIncubationProfit Then
            recoveryPoint = (ISAnnualizedProfit / 365.25) * (row - OOSBeginRow) * minIncubationProfit
        Else
            recoveryPoint = last_quit_equity_high
        End If
        
        
        ' Check if strategy has run for at least 21 days OOS
        If row - OOSBeginRow > 21 Then
            Select Case quit_status
                Case "Continue"
                    If currentEquity < quitEquity Then
                        quit_status = "Quit"
                        quit_date = currentdate
                        last_quit_equity_high = peakEquity ' Record the peak equity at quitting
                    End If
                Case "Quit"
                    
                    If currentEquity > (recoveryPoint + quitEquity) / 2 Then
                       quit_status = "Coming Back"
                       
                    End If
                    ' If strategy was profitable at quit, only restart when new equity high is reached
                    If currentEquity > recoveryPoint Then
                       quit_status = "Recovered"
                    End If
        
                Case "Coming Back"
                    ' Strategy has regained profitability but has not yet hit a new high
                    If currentEquity > recoveryPoint Then
                        quit_status = "Recovered"
                    End If
                    If currentEquity < (recoveryPoint + quitEquity) / 2 Then
                        quit_status = "Quit"
                        last_quit_equity_high = peakEquity
                    End If
                    
                    
                Case "Recovered"
                    ' Strategy is continuing after making a new high, check if it should quit again
                    If currentEquity < quitEquity Then
                        quit_status = "Quit"
                        quit_date = currentdate
                        last_quit_equity_high = peakEquity ' Update last quit high
                    End If
        
            End Select
        End If
        
        ' Update equity for the last 12 months only
        If currentdate >= last12MonthsDate Then
        currentEquityLast12Months = currentEquityLast12Months + dailyProfit
            
            ' Update peak equity within the last 12 months
            If currentEquityLast12Months > peakEquityLast12Months Then
                peakEquityLast12Months = currentEquityLast12Months
            End If
            
            ' Calculate drawdown for the last 12 months
            currentDrawdownLast12Months = peakEquityLast12Months - currentEquityLast12Months
            If currentDrawdownLast12Months > maxDrawdownLast12Months Then
                maxDrawdownLast12Months = currentDrawdownLast12Months
            End If
        End If
        
        
    Next row
    
    ' After the loop, add the last month's data
    If lastMonth <> 0 Then
        monthlyProfits.Add currentMonthProfit, Format(lastYear, "0000") & Format(lastMonth, "00")
    End If
    
    

    ' Final average drawdown calculation
    avgOOSDrawdown = avgOOSDrawdown / (OOSEndRow - OOSBeginRow + 1)

    ' Place the calculated drawdown values in the summary sheet
    wsSummary.Cells(i, COL_MAX_OOS_DRAWDOWN).value = maxOOSDrawdown
    wsSummary.Cells(i, COL_AVG_OOS_DRAWDOWN).value = avgOOSDrawdown
    wsSummary.Cells(i, COL_MAX_DRAWDOWN_LAST_12_MONTHS).value = maxDrawdownLast12Months
    'wsSummary.Cells(i, COL_MAX_DRAWDOWN_PERCENT).value = maxdrawdownpercent
    
    
    ' Place the calculated profits in the summary sheet
    wsSummary.Cells(i, COL_PROFIT_LAST_1_MONTH).value = profitLast1Month
    wsSummary.Cells(i, COL_PROFIT_LAST_3_MONTHS).value = profitLast3Months
    wsSummary.Cells(i, COL_PROFIT_LAST_6_MONTHS).value = profitLast6Months
    wsSummary.Cells(i, COL_PROFIT_LAST_9_MONTHS).value = profitLast9Months
    wsSummary.Cells(i, COL_PROFIT_LAST_12_MONTHS).value = profitLast12Months
    wsSummary.Cells(i, COL_PROFIT_SINCE_OOS_START).value = profitSinceOOSStart
    wsSummary.Cells(i, COL_COUNT_PROFIT_MONTHS).value = CountPositiveMonths(monthlyProfits, TotalMonthlyCount, OOSEndDate, daysThreshold)



    ' Calculate Efficiency against IS Expected Annualized Profit for each period
    ' Modify the efficiency calculations based on whether we're using monthly or rolling periods
    If ISAnnualizedProfit <> 0 Then
        If daysThreshold > 0 Then
            ' For monthly periods, use actual number of days in the period for more accurate calculation
            Dim days1M As Long, days3M As Long, days6M As Long, days9M As Long, days12M As Long
            
            ' Calculate actual days in each period
            days1M = DateDiff("d", lastMonthDate, effectiveEndDate) + 1
            days3M = DateDiff("d", last3MonthsDate, effectiveEndDate) + 1
            days6M = DateDiff("d", last6MonthsDate, effectiveEndDate) + 1
            days9M = DateDiff("d", last9MonthsDate, effectiveEndDate) + 1
            days12M = DateDiff("d", last12MonthsDate, effectiveEndDate) + 1
            
            ' Calculate efficiencies using actual days
            If profitLast1Month = "" Then
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_1_MONTH).value = ""
            Else
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_1_MONTH).value = profitLast1Month / (ISAnnualizedProfit * days1M / 365.25)
            End If
            
            If profitLast3Months = "" Then
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_3_MONTHS).value = ""
            Else
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_3_MONTHS).value = profitLast3Months / (ISAnnualizedProfit * days3M / 365.25)
            End If
            
            If profitLast6Months = "" Then
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_6_MONTHS).value = ""
            Else
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_6_MONTHS).value = profitLast6Months / (ISAnnualizedProfit * days6M / 365.25)
            End If
            
            If profitLast9Months = "" Then
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_9_MONTHS).value = ""
            Else
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_9_MONTHS).value = profitLast9Months / (ISAnnualizedProfit * days9M / 365.25)
            End If
            
            If profitLast12Months = "" Then
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_12_MONTHS).value = ""
            Else
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_12_MONTHS).value = profitLast12Months / (ISAnnualizedProfit * days12M / 365.25)
            End If
            
        Else
        
            ' Use original rolling period efficiency calculations
        
            If profitLast1Month = "" Then
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_1_MONTH).value = ""
            Else
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_1_MONTH).value = profitLast1Month / (ISAnnualizedProfit / 12)
            End If
            
            If profitLast3Months = "" Then
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_3_MONTHS).value = ""
            Else
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_3_MONTHS).value = profitLast3Months / (ISAnnualizedProfit / 4)
            End If
            
            If profitLast6Months = "" Then
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_6_MONTHS).value = ""
            Else
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_6_MONTHS).value = profitLast6Months / (ISAnnualizedProfit / 2)
            End If
            
            If profitLast9Months = "" Then
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_9_MONTHS).value = ""
            Else
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_9_MONTHS).value = profitLast9Months / (ISAnnualizedProfit / (3 / 4))
            End If
            
            If profitLast12Months = "" Then
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_12_MONTHS).value = ""
            Else
                wsSummary.Cells(i, COL_EFFICIENCY_LAST_12_MONTHS).value = profitLast12Months / (ISAnnualizedProfit)
            End If
        End If
        wsSummary.Cells(i, COL_EFFICIENCY_SINCE_OOS_START).value = profitSinceOOSStart / (ISAnnualizedProfit * (OOSEndDate - OOSBeginDate) / 365.25)
    End If

    Dim OOSmonths As Long
    OOSmonths = DateDiff("m", OOSBeginDate, OOSEndDate)

     ' Incubation Flag Logic: Default to "Not Passed" if not already set to "Passed"
    If (OOSEndDate - OOSBeginDate) < IncubationPeriod * 30.5 Then
        wsSummary.Cells(i, COL_INCUBATION_STATUS).value = "Not Passed Yet"
    End If
    If (OOSEndDate - OOSBeginDate) >= IncubationPeriod * 30.5 And wsSummary.Cells(i, COL_INCUBATION_STATUS).value <> "Passed" Then
        wsSummary.Cells(i, COL_INCUBATION_STATUS).value = "Not Passed"
    End If
    
    
    
    
    wsSummary.Cells(i, COL_QUITTING_STATUS).value = quit_status
    wsSummary.Cells(i, COL_QUITTING_DATE).value = quit_date
    
End Sub


' Function to count positive months in a range

Function CountPositiveMonths(ByRef monthlyProfits As Collection, ByVal numMonths As Long, ByVal endDate As Date, ByVal daysThreshold As Long) As Long
    Dim count As Long
    Dim i As Long
    Dim monthKey As String
    Dim profit As Double
    Dim eligiblitycompare As String
    
    
    Dim effectiveEndDate As Date
    
    ' Only apply monthly logic if daysThreshold > 0
    If daysThreshold > 0 Then
        ' Get the first day of the current month of endDate
        Dim currentMonthStart As Date
        currentMonthStart = DateSerial(Year(endDate), Month(endDate), 1)
        
        ' Calculate days in current month
        Dim daysInCurrentMonth As Long
        daysInCurrentMonth = DateDiff("d", currentMonthStart, endDate) + 1
        
        ' Set effective end date based on days threshold
        If daysInCurrentMonth < daysThreshold Then
            effectiveEndDate = DateSerial(Year(endDate), Month(endDate), 1) - 1
        Else
            effectiveEndDate = endDate
        End If
    Else
        ' Use original rolling period logic
        effectiveEndDate = endDate
    End If
    
    eligiblitycompare = GetNamedRangeValue("EligibilityGreaterThan")
    
    count = 0
    For i = 0 To numMonths - 1
        monthKey = Format(Year(DateAdd("m", -i, effectiveEndDate)), "0000") & Format(Month(DateAdd("m", -i, effectiveEndDate)), "00")
        
        ' Try to get the profit for this month
        On Error Resume Next
        profit = monthlyProfits(monthKey)
        If Err.Number = 0 Then ' If month exists in collection
            If eligiblitycompare = ">0" Then
                If profit > 0 Then
                    count = count + 1
                End If
            ElseIf eligiblitycompare = ">=0" Then
                If profit >= 0 Then
                    count = count + 1
                End If
            End If
        End If
        On Error GoTo 0
    Next i
    
    CountPositiveMonths = count
End Function



Sub CalculateTradeProfitFactors(wsSummary As Worksheet, rowIndex As Long, strategyName As String)
    ' This function calculates profit factors and related metrics from the Long_Trades and Short_Trades sheets
    ' and populates the corresponding columns in the Summary sheet
    '
    ' Parameters:
    '   wsSummary - The Summary worksheet
    '   rowIndex - The row index in the Summary worksheet to update
    '   strategyName - The name of the strategy to process
    
    Dim wsLongTrades As Worksheet
    Dim wsShortTrades As Worksheet
    Dim longColumn As Long, shortColumn As Long
    Dim longGrossProfit As Double, longGrossLoss As Double, longNetProfit As Double
    Dim shortGrossProfit As Double, shortGrossLoss As Double, shortNetProfit As Double
    Dim totalGrossProfit As Double, totalGrossLoss As Double, totalNetProfit As Double
    Dim longProfitFactor As Double, shortProfitFactor As Double, totalProfitFactor As Double
    Dim i As Long, lastRow As Long
    Dim tradeValue As Variant
    
    ' Initialize values
    longGrossProfit = 0
    longGrossLoss = 0
    shortGrossProfit = 0
    shortGrossLoss = 0
    
    ' Check if Long_Trades and Short_Trades sheets exist
    On Error Resume Next
    Set wsLongTrades = ThisWorkbook.Sheets("Long_Trades")
    Set wsShortTrades = ThisWorkbook.Sheets("Short_Trades")
    On Error GoTo 0
    
    If wsLongTrades Is Nothing Or wsShortTrades Is Nothing Then
        ' One or both sheets don't exist, exit
        Exit Sub
    End If
    
    ' Find the column for this strategy in Long_Trades
    longColumn = 0
    For i = 1 To wsLongTrades.Cells(1, wsLongTrades.Columns.count).End(xlToLeft).column
        If wsLongTrades.Cells(1, i).value = strategyName Then
            longColumn = i
            Exit For
        End If
    Next i
    
    ' Find the column for this strategy in Short_Trades
    shortColumn = 0
    For i = 1 To wsShortTrades.Cells(1, wsShortTrades.Columns.count).End(xlToLeft).column
        If wsShortTrades.Cells(1, i).value = strategyName Then
            shortColumn = i
            Exit For
        End If
    Next i
    
    ' If strategy not found in either sheet, exit
    If longColumn = 0 And shortColumn = 0 Then
        Exit Sub
    End If
    
    ' Process Long Trades
    If longColumn > 0 Then
        lastRow = wsLongTrades.Cells(wsLongTrades.rows.count, longColumn).End(xlUp).row
        
        For i = 2 To lastRow ' Start from row 2 (skip header)
            tradeValue = wsLongTrades.Cells(i, longColumn).value
            
            ' Skip empty cells
            If Not IsEmpty(tradeValue) Then
                ' Check if cell contains a numeric value
                If IsNumeric(tradeValue) Then
                    If CDbl(tradeValue) > 0 Then
                        longGrossProfit = longGrossProfit + CDbl(tradeValue)
                    ElseIf CDbl(tradeValue) < 0 Then
                        longGrossLoss = longGrossLoss + Abs(CDbl(tradeValue)) ' Store loss as positive value
                    End If
                End If
            End If
        Next i
    End If
    
    ' Process Short Trades
    If shortColumn > 0 Then
        lastRow = wsShortTrades.Cells(wsShortTrades.rows.count, shortColumn).End(xlUp).row
        
        For i = 2 To lastRow ' Start from row 2 (skip header)
            tradeValue = wsShortTrades.Cells(i, shortColumn).value
            
            ' Skip empty cells
            If Not IsEmpty(tradeValue) Then
                ' Check if cell contains a numeric value
                If IsNumeric(tradeValue) Then
                    If CDbl(tradeValue) > 0 Then
                        shortGrossProfit = shortGrossProfit + CDbl(tradeValue)
                    ElseIf CDbl(tradeValue) < 0 Then
                        shortGrossLoss = shortGrossLoss + Abs(CDbl(tradeValue)) ' Store loss as positive value
                    End If
                End If
            End If
        Next i
    End If
    
    ' Calculate net profits
    longNetProfit = longGrossProfit - longGrossLoss
    shortNetProfit = shortGrossProfit - shortGrossLoss
    totalNetProfit = longNetProfit + shortNetProfit
    totalGrossProfit = longGrossProfit + shortGrossProfit
    totalGrossLoss = longGrossLoss + shortGrossLoss
    
    ' Calculate profit factors (avoiding division by zero)
    If longGrossLoss > 0 Then
        longProfitFactor = longGrossProfit / longGrossLoss
    Else
        longProfitFactor = IIf(longGrossProfit > 0, 999, 0) ' If no losses but has profits, set to 999
    End If
    
    If shortGrossLoss > 0 Then
        shortProfitFactor = shortGrossProfit / shortGrossLoss
    Else
        shortProfitFactor = IIf(shortGrossProfit > 0, 999, 0) ' If no losses but has profits, set to 999
    End If
    
    If totalGrossLoss > 0 Then
        totalProfitFactor = totalGrossProfit / totalGrossLoss
    Else
        totalProfitFactor = IIf(totalGrossProfit > 0, 999, 0) ' If no losses but has profits, set to 999
    End If
    
    ' Update the Summary sheet with calculated values
    wsSummary.Cells(rowIndex, COL_PROFIT_FACTOR).value = totalProfitFactor
    wsSummary.Cells(rowIndex, COL_PROFIT_LONG_FACTOR).value = longProfitFactor
    wsSummary.Cells(rowIndex, COL_PROFIT_SHORT_FACTOR).value = shortProfitFactor
    wsSummary.Cells(rowIndex, COL_GROSS_LONG_NETPROFIT).value = longNetProfit
    wsSummary.Cells(rowIndex, COL_GROSS_LONG_PROFIT).value = longGrossProfit
    wsSummary.Cells(rowIndex, COL_GROSS_LONG_LOSS).value = longGrossLoss
    wsSummary.Cells(rowIndex, COL_GROSS_SHORT_NETPROFIT).value = shortNetProfit
    wsSummary.Cells(rowIndex, COL_GROSS_SHORT_PROFIT).value = shortGrossProfit
    wsSummary.Cells(rowIndex, COL_GROSS_SHORT_LOSS).value = shortGrossLoss
End Sub


Sub FormatSummaryTable()
    Dim wsSummary As Worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")

    Dim lastRow As Long, lastCol As Long
    Dim ExcelDateFormat As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    
    ExcelDateFormat = GetNamedRangeValue("DateFormat")
    
    lastRow = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row
    lastCol = wsSummary.Cells(1, wsSummary.Columns.count).End(xlToLeft).column

    ' Store the currently active sheet
    Dim currentSheet As Worksheet
    Set currentSheet = activeSheet
    
    ' Switch to wsSummary, freeze panes, and switch back
    
    Dim freezecol As String
    freezecol = GetNamedRangeValue("FreezePanesColumn")
    
    wsSummary.Activate
    
    wsSummary.Range(freezecol & "2").Select
    ActiveWindow.FreezePanes = True
    
    ' Apply formatting to the header row
    With wsSummary.rows(1)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True ' Enable wrap text for the header row
        .Interior.Color = RGB(220, 230, 241) ' Light blue header background
    End With
    
    ' Apply borders to the table
    With wsSummary.Range(wsSummary.Cells(1, 1), wsSummary.Cells(lastRow, lastCol))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    
    
    
    wsSummary.Columns(COL_SECTOR_RANK).NumberFormat = "0;-0;0"
    wsSummary.Columns(COL_SYMBOL_RANK).NumberFormat = "0;-0;0"
    
    ' Number formatting for certain columns (dollars and percentages)
    
    wsSummary.Columns(COL_MARGIN).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_EXPECTED_ANNUAL_PROFIT).NumberFormat = "$#,##0" ' Expected Annual Return
    wsSummary.Columns(COL_ACTUAL_ANNUAL_PROFIT).NumberFormat = "$#,##0" ' Actual Annual Return
    wsSummary.Columns(COL_EXPECTED_ANNUAL_RETURN).NumberFormat = "0%" ' Expected Annual Return
    wsSummary.Columns(COL_ACTUAL_ANNUAL_RETURN).NumberFormat = "0%" ' Actual Annual Return
    wsSummary.Columns(COL_NOTIONAL_CAPITAL).NumberFormat = "$#,##0" '
    
    
    wsSummary.Columns(COL_RETURN_EFFICIENCY).NumberFormat = "0%" ' Return Efficiency (as percentage)
    'wsSummary.Columns(COL_RISK_RUIN).NumberFormat = "0%"
    wsSummary.Columns(COL_BACKTEST_MC).NumberFormat = "0.00"
    wsSummary.Columns(COL_CLOSEDTRADEMC).NumberFormat = "0.00"

    wsSummary.Columns(COL_IS_MONTE_CARLO).NumberFormat = "0.00"
    wsSummary.Columns(COL_IS_OOS_MONTE_CARLO).NumberFormat = "0.00"
    wsSummary.Columns(COL_IS_ANNUAL_SD_IS).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_IS_ANNUAL_SD_ISOOS).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_OOS_WINRATE).NumberFormat = "0%"
    wsSummary.Columns(COL_OVERALL_WINRATE).NumberFormat = "0%"
    wsSummary.Columns(COL_TRADES_PER_YEAR).NumberFormat = "0;-0;0"
    wsSummary.Columns(COL_AVG_IS_OOS_TRADE).NumberFormat = "$#,##0"
    

    wsSummary.Columns(COL_AVG_PROFIT_IS_OOS_TRADE).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_AVG_LOSS_IS_OOS_TRADE).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_LARGEST_WIN_IS_OOS_TRADE).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_LARGEST_LOSS_IS_OOS_TRADE).NumberFormat = "$#,##0"
    
    
    
    
    wsSummary.Columns(COL_TOTAL_IS_PROFIT).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_TOTAL_IS_OSS_PROFIT).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_ANNUALIZED_NET_PROFIT_IS_OOS).NumberFormat = "$#,##0"
            
    wsSummary.Columns(COL_R_DD_12MONTH).NumberFormat = "0.00"
    wsSummary.Columns(COL_R_DD_OOS).NumberFormat = "0.00"
    wsSummary.Columns(COL_AVG_TRADE_LENGTH).NumberFormat = "0.00"
    
    wsSummary.Columns(COL_WORST_BACKTEST_DRAWDOWN).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_WORST_IS_OOS_DRAWDOWN).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_AVG_BACKTEST_DRAWDOWN).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_AVG_IS_OOS_DRAWDOWN).NumberFormat = "$#,##0"
    wsSummary.Columns(COL_PERCENT_TIME_IN_MARKET).NumberFormat = "0%"
    wsSummary.Columns(COL_BACKTEST_WINRATE).NumberFormat = "0%"
    wsSummary.Columns(COL_PROFIT_LAST_1_MONTH).NumberFormat = "$#,##0" ' Profit last 1 month
    wsSummary.Columns(COL_PROFIT_LAST_3_MONTHS).NumberFormat = "$#,##0" ' Profit last 3 months
    wsSummary.Columns(COL_PROFIT_LAST_6_MONTHS).NumberFormat = "$#,##0" ' Profit last 6 months
    wsSummary.Columns(COL_PROFIT_LAST_9_MONTHS).NumberFormat = "$#,##0" ' Profit last 9 months
    wsSummary.Columns(COL_PROFIT_LAST_12_MONTHS).NumberFormat = "$#,##0" ' Profit last 12 months
    wsSummary.Columns(COL_PROFIT_SINCE_OOS_START).NumberFormat = "$#,##0" ' Profit Since OOS Start
    wsSummary.Columns(COL_COUNT_PROFIT_MONTHS).NumberFormat = "0;-0;0"
    
    wsSummary.Columns(COL_MAX_OOS_DRAWDOWN).NumberFormat = "$#,##0" ' Max OOS Drawdown
    wsSummary.Columns(COL_AVG_OOS_DRAWDOWN).NumberFormat = "$#,##0" ' Avg OOS Drawdown
    wsSummary.Columns(COL_MAX_DRAWDOWN_LAST_12_MONTHS).NumberFormat = "$#,##0" ' Avg OOS Drawdown
   ' wsSummary.Columns(COL_MAX_DRAWDOWN_PERCENT).NumberFormat = "0%"
    wsSummary.Columns(COL_EFFICIENCY_LAST_1_MONTH).NumberFormat = "0%" ' Efficiency last 1 month
    wsSummary.Columns(COL_EFFICIENCY_LAST_3_MONTHS).NumberFormat = "0%" ' Efficiency last 3 months
    wsSummary.Columns(COL_EFFICIENCY_LAST_6_MONTHS).NumberFormat = "0%" ' Efficiency last 6 months
    wsSummary.Columns(COL_EFFICIENCY_LAST_9_MONTHS).NumberFormat = "0%" ' Efficiency last 9 months
    wsSummary.Columns(COL_EFFICIENCY_LAST_12_MONTHS).NumberFormat = "0%" ' Efficiency last 12 months
    wsSummary.Columns(COL_EFFICIENCY_SINCE_OOS_START).NumberFormat = "0%" ' Efficiency since OOS start
    
    wsSummary.Columns(COL_PROFIT_FACTOR).NumberFormat = "0.00"
    wsSummary.Columns(COL_PROFIT_LONG_FACTOR).NumberFormat = "0.00"
    wsSummary.Columns(COL_PROFIT_SHORT_FACTOR).NumberFormat = "0.00"
    
    wsSummary.Columns(COL_GROSS_LONG_NETPROFIT).NumberFormat = "$#,##0" ' Max OOS Drawdown
    wsSummary.Columns(COL_GROSS_LONG_PROFIT).NumberFormat = "$#,##0" ' Avg OOS Drawdown
    wsSummary.Columns(COL_GROSS_LONG_LOSS).NumberFormat = "$#,##0" ' Avg OOS Drawdown
      
    wsSummary.Columns(COL_GROSS_SHORT_NETPROFIT).NumberFormat = "$#,##0" ' Max OOS Drawdown
    wsSummary.Columns(COL_GROSS_SHORT_PROFIT).NumberFormat = "$#,##0" ' Avg OOS Drawdown
    wsSummary.Columns(COL_GROSS_SHORT_LOSS).NumberFormat = "$#,##0" ' Avg OOS Drawdown
    
    
    ' Date formatting for date columns
    
    If ExcelDateFormat <> "US" Then
        wsSummary.Columns(COL_NEXT_OPT_DATE).NumberFormat = "dd/mm/yyyy" ' Next Opt Date
        wsSummary.Columns(COL_LAST_OPT_DATE).NumberFormat = "dd/mm/yyyy" ' Last Opt Date
        wsSummary.Columns(COL_OOS_BEGIN_DATE).NumberFormat = "dd/mm/yyyy" ' OOS Begin Date
        wsSummary.Columns(COL_LAST_DATE_ON_FILE).NumberFormat = "dd/mm/yyyy" ' Last Date on File
        wsSummary.Columns(COL_START_DATE).NumberFormat = "dd/mm/yyyy" ' Last Date on File
        wsSummary.Columns(COL_INCUBATION_DATE).NumberFormat = "dd/mm/yyyy" ' Last Date on File
        wsSummary.Columns(COL_QUITTING_DATE).NumberFormat = "dd/mm/yyyy" ' Last Date on File
        
    Else
        wsSummary.Columns(COL_NEXT_OPT_DATE).NumberFormat = "mm/dd/yyyy" ' Next Opt Date
        wsSummary.Columns(COL_LAST_OPT_DATE).NumberFormat = "mm/dd/yyyy" ' Last Opt Date
        wsSummary.Columns(COL_OOS_BEGIN_DATE).NumberFormat = "mm/dd/yyyy" ' OOS Begin Date"
        wsSummary.Columns(COL_LAST_DATE_ON_FILE).NumberFormat = "mm/dd/yyyy" ' Last Date on File
        wsSummary.Columns(COL_START_DATE).NumberFormat = "mm/dd/yyyy" ' Last Date on File
        wsSummary.Columns(COL_INCUBATION_DATE).NumberFormat = "mm/dd/yyyy" ' Last Date on File
        wsSummary.Columns(COL_QUITTING_DATE).NumberFormat = "mm/dd/yyyy" ' Last Date on File
    End If
    
  
  
  
    ' Apply formatting
    With wsSummary.Range(wsSummary.Cells(2, COL_STATUS), wsSummary.Cells(lastRow, COL_STATUS))
        .Interior.Color = RGB(230, 230, 250) ' Light lavender color for input style
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With wsSummary.Range(wsSummary.Cells(2, COL_ELIGIBILITY), wsSummary.Cells(lastRow, COL_ELIGIBILITY))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    

    ' Conditional formatting for "Incubation Status"
    Dim incubationRange As Range
    Set incubationRange = wsSummary.Range(wsSummary.Cells(2, COL_INCUBATION_STATUS), wsSummary.Cells(lastRow, COL_INCUBATION_STATUS))
    With incubationRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Passed""")
        .Interior.Color = RGB(198, 239, 206) ' Green for Passed
        .Font.Bold = True
    End With
    With incubationRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Not Passed""")
        .Interior.Color = RGB(255, 199, 206) ' Red for Not Passed
        .Font.Bold = True
    End With
    With incubationRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Not Passed Yet""")
        .Interior.Color = RGB(255, 255, 0) ' Yellow for Not Passed Yet
        .Font.Bold = True
    End With

    ' Conditional formatting for "Quitting Status"
    Dim quittingRange As Range
    Set quittingRange = wsSummary.Range(wsSummary.Cells(2, COL_QUITTING_STATUS), wsSummary.Cells(lastRow, COL_QUITTING_STATUS))
    With quittingRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Quit""")
        .Interior.Color = RGB(255, 199, 206) ' Red for Quit
        .Font.Bold = True
    End With
    With quittingRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Continue""")
        .Interior.Color = RGB(198, 239, 206) ' Green for Continue
    End With
    With quittingRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Coming Back""")
        .Interior.Color = RGB(255, 255, 0) '  YEllow
        .Font.Bold = True
    End With
    With quittingRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Recovered""")
        .Interior.Color = RGB(180, 220, 200) ' Green for Continue
    End With
     
    

    ' Conditional formatting for "Eligibility"
    Dim EligibilityRange As Range
    Set EligibilityRange = wsSummary.Range(wsSummary.Cells(2, COL_ELIGIBILITY), wsSummary.Cells(lastRow, COL_QUITTING_STATUS))
    With EligibilityRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""No""")
        .Interior.Color = RGB(255, 199, 206) ' Red for Quit
        .Font.Bold = True
    End With
    With EligibilityRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Yes""")
        .Interior.Color = RGB(198, 239, 206) ' Green for Continue
        .Font.Bold = True
    End With


    ' Highlight negative profits in red
    Dim col As Long
    For col = COL_PROFIT_LAST_1_MONTH To COL_PROFIT_SINCE_OOS_START ' Columns containing profits
        wsSummary.Columns(col).FormatConditions.Delete ' Clear existing conditions
        wsSummary.Columns(col).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        wsSummary.Columns(col).FormatConditions(1).Font.Color = RGB(255, 0, 0) ' Red font for negative values
    Next col
    
     ' Auto-fit the columns to content
    wsSummary.Columns("A:BZ").AutoFit
    
    For col = COL_NEXT_OPT_DATE To COL_EFFICIENCY_SINCE_OOS_START ' Columns containing profits
        wsSummary.Columns(col).ColumnWidth = 12
    Next col
    For col = COL_STRATEGY_NUMBER To COL_OPEN_CODE_TAB ' Columns containing profits
        wsSummary.Columns(col).ColumnWidth = 10
    Next col
    For col = COL_SYMBOL To COL_STATUS ' Columns containing profits
        wsSummary.Columns(col).ColumnWidth = 9
    Next col

    For col = COL_INCUBATION_STATUS To COL_QUITTING_STATUS ' Columns containing profits
        wsSummary.Columns(col).ColumnWidth = 20
    Next col

    wsSummary.Columns(COL_CREATE_DETAILED_TAB).ColumnWidth = 11

    For col = COL_OPEN_CODE_TAB To COL_FOLDER ' Columns containing profits
        wsSummary.Columns(col).ColumnWidth = 7.5
    Next col

    wsSummary.Columns(COL_LAST_12_MONTHS).ColumnWidth = 15
    wsSummary.Columns(COL_ELIGIBILITY).ColumnWidth = 12
    wsSummary.Columns(COL_SECTOR_RANK).ColumnWidth = 6.5
    wsSummary.Columns(COL_SYMBOL_RANK).ColumnWidth = 6.5
    wsSummary.Columns(COL_STRATEGY_NAME).ColumnWidth = 87.5

    wsSummary.Columns(COL_STATUS).ColumnWidth = 18
    'wsSummary.Range(wsSummary.Columns(COL_WF_IN_OUT), wsSummary.Columns(199)).EntireColumn.Hidden = True
    
    
    
    
    With ThisWorkbook.Windows(1)
        .Zoom = 70 ' Set zoom level to 70%
    End With
    
    If GetNamedRangeValue("startcolwidthoverride") <> "" Then
        For col = COL_STRATEGY_NUMBER To COL_FOLDER ' Columns containing profits
            If GetNamedRangeValue("startcolwidthoverride") < 7.5 Then wsSummary.Cells(1, col).WrapText = False
            wsSummary.Columns(col).ColumnWidth = GetNamedRangeValue("startcolwidthoverride")
        Next col
    End If

End Sub




Sub ApplyConditionalFormatting()

    Dim wsSummary As Worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    Dim lastRow As Long
    lastRow = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row

    ' Status-based row shading (replaces alternate-row grey).
    ' Rows within the same status group are alternated between the base colour and a
    ' slightly darker shade so individual rows stay easy to read.
    Dim portSt As String, passSt As String, bnhSt As String
    portSt = GetNamedRangeValue("Port_Status")
    passSt = GetNamedRangeValue("Pass_Status")
    bnhSt  = GetNamedRangeValue("BuyandHoldStatus")

    Dim i As Long, sv As String, grpCount As Long, prevStatus As String
    grpCount = 0
    prevStatus = ""
    For i = 2 To lastRow
        sv = wsSummary.Cells(i, COL_STATUS).value
        If sv <> prevStatus Then
            grpCount = 0
            prevStatus = sv
        End If
        grpCount = grpCount + 1

        Dim baseR As Long, baseG As Long, baseB As Long
        Select Case sv
            Case portSt  ' Live / portfolio — green
                baseR = 198: baseG = 239: baseB = 206
            Case passSt  ' Passing — steel blue
                baseR = 189: baseG = 214: baseB = 238
            Case bnhSt   ' Buy & Hold — grey
                baseR = 220: baseG = 220: baseB = 220
            Case "Failed"
                baseR = 255: baseG = 199: baseB = 206  ' red
            Case Else
                baseR = 255: baseG = 243: baseB = 205  ' soft yellow for New / other
        End Select

        ' Alternate: even rows within a group get a slightly darker shade
        If grpCount Mod 2 = 0 Then
            baseR = Application.Max(0, baseR - 15)
            baseG = Application.Max(0, baseG - 15)
            baseB = Application.Max(0, baseB - 15)
        End If

        wsSummary.rows(i).Interior.Color = RGB(baseR, baseG, baseB)
    Next i

    ' Highlight if last re-opt happened within the last 6 weeks (42 days)
    Dim currentdate As Date
    currentdate = Date
    
    For i = 2 To lastRow
        Dim nextOptDate As Variant
        Dim lastOptDate As Variant
        Dim daysDiff As Long
        If wsSummary.Cells(i, COL_STATUS).value <> GetNamedRangeValue("BuyandHoldStatus") Then
            ' Check Next Option Date
            nextOptDate = wsSummary.Cells(i, COL_NEXT_OPT_DATE).value
            If IsDate(nextOptDate) Then
                daysDiff = Abs(DateDiff("d", nextOptDate, currentdate))
                If daysDiff <= 7 Then
                    wsSummary.Cells(i, COL_NEXT_OPT_DATE).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf daysDiff <= 15 Then
                    wsSummary.Cells(i, COL_NEXT_OPT_DATE).Interior.Color = RGB(255, 255, 0) ' Yellow
                End If
            End If
        
            ' Check Last Option Date
            lastOptDate = wsSummary.Cells(i, COL_LAST_OPT_DATE).value
            If IsDate(lastOptDate) Then
                daysDiff = Abs(DateDiff("d", lastOptDate, currentdate))
                If daysDiff <= 7 Then
                    wsSummary.Cells(i, COL_LAST_OPT_DATE).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf daysDiff <= 15 Then
                    wsSummary.Cells(i, COL_LAST_OPT_DATE).Interior.Color = RGB(255, 255, 0) ' Yellow
                End If
            End If
        End If
    Next i

    ' Conditional formatting for Actual Annual Return
    For i = 2 To lastRow
        Dim actualReturn As Double
        actualReturn = wsSummary.Cells(i, COL_ACTUAL_ANNUAL_PROFIT).value ' Actual Annual Return
        
        ' Red if the actual return is negative
        If actualReturn < 0 Then
            wsSummary.Cells(i, COL_ACTUAL_ANNUAL_PROFIT).Interior.Color = RGB(255, 99, 71) ' Red highlight
            wsSummary.Cells(i, COL_RETURN_EFFICIENCY).Interior.Color = RGB(255, 99, 71) ' Red highlight
        ' Yellow if return is positive but less than minimum incubation profit * expected annual profit
        ElseIf actualReturn > 0 Then
            Dim expectedProfit As Double, minIncubationProfit As Double
            expectedProfit = wsSummary.Cells(i, COL_EXPECTED_ANNUAL_PROFIT).value ' Expected Annual Profit
            minIncubationProfit = GetNamedRangeValue("Min_Incubation_Profit") ' Assuming Min Incubation Profit is in fixed cell
            
            If actualReturn < (expectedProfit * minIncubationProfit) Then
                wsSummary.Cells(i, COL_ACTUAL_ANNUAL_PROFIT).Interior.Color = RGB(255, 255, 102) ' Yellow highlight
                wsSummary.Cells(i, COL_RETURN_EFFICIENCY).Interior.Color = RGB(255, 255, 102) ' Yellow highlight
            End If
        End If
    Next i

End Sub



Sub AggregateWeeklyProfitsAndAddSparklines()
    Dim wsSummary As Worksheet
    Dim wsDailyM2M As Worksheet
    Dim lastRowSummary As Long
    Dim lastRowDailyM2M As Long
    Dim lastColDailyM2M As Long
    Dim strategyName As String
    Dim dailyProfitRange As Range
    Dim summaryRow As Long
    Dim startColumn As Long
    Dim weeklyProfits As Double
    Dim i As Long, j As Long
    Dim sparklineRange As Range
    Dim strategyCol As Long
    
    ' Set the worksheets
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    Set wsDailyM2M = ThisWorkbook.Sheets("DailyM2MEquity")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Set the starting column for weekly profits in the Summary sheet
    
    Dim mainstartstartColumn As Long
    mainstartstartColumn = COL_SHARPE_ISOOS + 20
    
     wsSummary.Cells(1, mainstartstartColumn).value = "Sparkline Inputs:"
    
    
    ' Find the last row with data in the Summary sheet
    lastRowSummary = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row
    
    ' Find the last row with data in the DailyM2MEquity sheet
    lastRowDailyM2M = EndRowByCutoffSimple(wsDailyM2M, 1)
    
    ' Find the last column in the DailyM2MEquity sheet (to search for strategy names)
    lastColDailyM2M = wsDailyM2M.Cells(1, wsDailyM2M.Columns.count).End(xlToLeft).column
    
    ' Loop through each strategy in the Summary sheet
    For summaryRow = 2 To lastRowSummary
        strategyName = wsSummary.Cells(summaryRow, COL_STRATEGY_NAME).value
        ' Reset start column for the next strategy
        startColumn = mainstartstartColumn
        ' Search for the strategy name in the first row of wsDailyM2M to find the corresponding column
        strategyCol = 0
        For i = 2 To lastColDailyM2M
            If wsDailyM2M.Cells(1, i).value = strategyName Then
                strategyCol = i
                Exit For
            End If
        Next i
        
        ' If the strategy is not found, skip to the next strategy
        If strategyCol = 0 Then
            MsgBox "Strategy " & strategyName & " not found in DailyM2MEquity sheet"
            GoTo NextStrategy
        End If
        
        ' Aggregate the daily returns into weekly cumulative profits (7-day intervals)
        j = 0
        weeklyProfits = 0
        For i = lastRowDailyM2M - 365 To lastRowDailyM2M Step 1
            weeklyProfits = weeklyProfits + wsDailyM2M.Cells(i, strategyCol).value ' Sum daily profits
            j = j + 1
            If j = 7 Then
                ' Store the weekly profit in the Summary sheet starting from column 50
                wsSummary.Cells(summaryRow, startColumn).value = weeklyProfits
                startColumn = startColumn + 1
                'weeklyProfits = 0
                j = 0
            End If
        Next i
        
               
        ' Add sparklines for the weekly profits in column 11 of the Summary sheet
        Set sparklineRange = wsSummary.Range(wsSummary.Cells(summaryRow, mainstartstartColumn), wsSummary.Cells(summaryRow, startColumn - 1))
        wsSummary.Cells(summaryRow, COL_LAST_12_MONTHS).SparklineGroups.Add Type:=xlSparkLine, sourceData:=sparklineRange.Address
        
NextStrategy:
    Next summaryRow
End Sub





Sub ButtonClickHandler()
    Dim clickedButton As Object
    Dim wsSummary As Worksheet
    Dim buttonRow As Long
    Dim gStrategyNumber As Long
    
    
     Call InitializeColumnConstantsManually
    ' Set the Summary worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    ' Get the button that was clicked
    Set clickedButton = wsSummary.Buttons(Application.Caller)
    
    ' Find the row of the button (based on the button's position)
    buttonRow = clickedButton.TopLeftCell.row
    
    ' Set the global variable to the strategy number in that row
    gStrategyNumber = wsSummary.Cells(buttonRow, COL_STRATEGY_NUMBER).value
    
    ' Call the macro to create the strategy tab
    CreateStrategyTab gStrategyNumber
End Sub

Sub ButtonClickHandlerPort()
    Dim clickedButton As Object
    Dim wsPortfolio As Worksheet
    Dim buttonRow As Long
    Dim gStrategyNumber As Long
    
    
     Call InitializeColumnConstantsManually
    ' Set the Summary worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    
    ' Get the button that was clicked
    Set clickedButton = wsPortfolio.Buttons(Application.Caller)
    
    ' Find the row of the button (based on the button's position)
    buttonRow = clickedButton.TopLeftCell.row
    
    ' Set the global variable to the strategy number in that row
    gStrategyNumber = wsPortfolio.Cells(buttonRow, COL_PORT_STRATEGY_NUMBER).value
    
    ' Call the macro to create the strategy tab
    CreateStrategyTab gStrategyNumber
End Sub



Sub ClearAllButtons()
    Dim wsSummary As Worksheet
    Dim btn As Object
    
    ' Set the Summary worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    ' Loop through all buttons in the worksheet and delete them
    For Each btn In wsSummary.Buttons
        btn.Delete
    Next btn
End Sub


Sub SetupButtonsAndStrategyTabCreation()
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
        buttonLeft = wsSummary.Cells(summaryRow, COL_CREATE_DETAILED_TAB).left + 2
        buttonTop = wsSummary.Cells(summaryRow, COL_CREATE_DETAILED_TAB).top + 1
        buttonWidth = wsSummary.Cells(summaryRow, COL_CREATE_DETAILED_TAB).Width - 2
        buttonHeight = wsSummary.Cells(summaryRow, COL_CREATE_DETAILED_TAB).Height - 1
        
        ' Add button for creating strategy tab
        Set btn = wsSummary.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
        With btn
            .OnAction = "ButtonClickHandler" ' Assign the macro to handle the button click
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




Sub SetupButtonsforMC()
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
    
    ' Loop through each strategy and add a button for running Monte Carlo in Column C
    For summaryRow = 2 To lastRowSummary
        ' Set the button's size and position to fit within the cell in Column C
        buttonLeft = wsSummary.Cells(summaryRow, COL_RUN_MC).left + 1
        buttonTop = wsSummary.Cells(summaryRow, COL_RUN_MC).top + 1
        buttonWidth = wsSummary.Cells(summaryRow, COL_RUN_MC).Width - 2
        buttonHeight = wsSummary.Cells(summaryRow, COL_RUN_MC).Height - 1
        
        ' Add button for running Monte Carlo
        Set btn = wsSummary.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
        With btn
            .OnAction = "ButtonClickHandlerMC" ' Assign the macro to handle the button click
            .Caption = "{}" ' Clearer caption
            .name = "RunMCBtn" & summaryRow ' Assign a unique name to each button
        End With
    Next summaryRow
End Sub


Sub CreateStatusDropdown(target As Range)
    On Error GoTo ErrorHandler
    
    ' Get status list
    Dim statusString As String
    statusString = GetStatusListAsString()
    
    ' Create validation
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=statusString
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in CreateStatusDropdown: " & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

          




Sub ButtonClickHandlerMC()
    Dim clickedButton As Object
    Dim wsSummary As Worksheet
    Dim buttonRow As Long
    Dim gStrategyNumber As Long

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually

    ' Set the Summary worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")

    ' Get the button that was clicked
    On Error Resume Next
    Set clickedButton = wsSummary.Buttons(Application.Caller)
    On Error GoTo 0
    
    If clickedButton Is Nothing Then
        MsgBox "Button not recognized. Please ensure the button is linked correctly.", vbExclamation
        Exit Sub
    End If

    ' Find the row of the button (based on the button's position)
    buttonRow = clickedButton.TopLeftCell.row
    ' Set the global variable to the strategy number in that row
    'gStrategyNumber = wsSummary.Cells(buttonRow, COL_STRATEGY_NUMBER).value

    ' Call the macro to create the strategy tab
    RunMonteCarloSimulation buttonRow
End Sub




Sub ButtonClickHandlerCode()
    Dim clickedButton As Object
    Dim wsSummary As Worksheet
    Dim buttonRow As Long
    Dim gStrategyName As String
    Dim gStrategyNumber As Long

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually

    ' Set the Summary worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")

    ' Get the button that was clicked
    On Error Resume Next
    Set clickedButton = wsSummary.Buttons(Application.Caller)
    On Error GoTo 0
    
    If clickedButton Is Nothing Then
        MsgBox "Button not recognized. Please ensure the button is linked correctly.", vbExclamation
        Exit Sub
    End If

    ' Find the row of the button (based on the button's position)
    buttonRow = clickedButton.TopLeftCell.row

    ' Set the global variable to the strategy number in that row
    gStrategyName = wsSummary.Cells(buttonRow, COL_STRATEGY_NAME).value
    gStrategyNumber = wsSummary.Cells(buttonRow, COL_STRATEGY_NUMBER).value
    ' Call the macro to create the strategy tab
    OpenStrategyCodeFile gStrategyName, gStrategyNumber, "tab"
    
End Sub




Sub ButtonClickHandlerCodeText()
    Dim clickedButton As Object
    Dim wsSummary As Worksheet
    Dim buttonRow As Long
    Dim gStrategyName As String
    Dim gStrategyNumber As Long

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually

    ' Set the Summary worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")

    ' Get the button that was clicked
    On Error Resume Next
    Set clickedButton = wsSummary.Buttons(Application.Caller)
    On Error GoTo 0
    
    If clickedButton Is Nothing Then
        MsgBox "Button not recognized. Please ensure the button is linked correctly.", vbExclamation
        Exit Sub
    End If

    ' Find the row of the button (based on the button's position)
    buttonRow = clickedButton.TopLeftCell.row

    ' Set the global variable to the strategy number in that row
    gStrategyName = wsSummary.Cells(buttonRow, COL_STRATEGY_NAME).value
    gStrategyNumber = wsSummary.Cells(buttonRow, COL_STRATEGY_NUMBER).value
    ' Call the macro to create the strategy tab
    OpenStrategyCodeFile gStrategyName, gStrategyNumber, "file"
    
End Sub

Sub ButtonClickHandlerCodeFolder()
    Dim clickedButton As Object
    Dim wsSummary As Worksheet
    Dim buttonRow As Long
    Dim gStrategyName As String
    Dim gStrategyNumber As Long

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually

    ' Set the Summary worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")

    ' Get the button that was clicked
    On Error Resume Next
    Set clickedButton = wsSummary.Buttons(Application.Caller)
    On Error GoTo 0
    
    If clickedButton Is Nothing Then
        MsgBox "Button not recognized. Please ensure the button is linked correctly.", vbExclamation
        Exit Sub
    End If

    ' Find the row of the button (based on the button's position)
    buttonRow = clickedButton.TopLeftCell.row

    ' Set the global variable to the strategy number in that row
    gStrategyName = wsSummary.Cells(buttonRow, COL_STRATEGY_NAME).value
    gStrategyNumber = wsSummary.Cells(buttonRow, COL_STRATEGY_NUMBER).value
    ' Call the macro to create the strategy tab
    OpenStrategyCodeFile gStrategyName, gStrategyNumber, "folder"
    
End Sub

Sub ButtonClickHandlerDetailedStrat()
    Dim clickedButton As Object
    Dim wsSummary As Worksheet
    Dim wsCurrentSheet As Worksheet
    Dim buttonRow As Long
    Dim gStrategyName As String
    Dim gStrategyNumber As Long

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    Set wsCurrentSheet = ThisWorkbook.activeSheet
    
    ' Set the Summary worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    ' Get the button that was clicked
    On Error Resume Next
    Set clickedButton = wsSummary.Buttons(Application.Caller)
    On Error GoTo 0
    
    If clickedButton Is Nothing Then
        MsgBox "Button not recognized. Please ensure the button is linked correctly.", vbExclamation
        Exit Sub
    End If


    
    ' Set the global variable to the strategy number in that row
    gStrategyName = wsCurrentSheet.Cells(1, 12).value
    gStrategyNumber = wsCurrentSheet.Cells(1, 11).value
    
    
    CreateStrategyTab (gStrategyNumber)
    
End Sub






Sub LookupMarginRequirements(wsSummary As Worksheet, wsDetails As Worksheet, summaryRow As Long, symbol As String, headerArray As Variant)

    Dim wsIBMargin As Worksheet
    Dim wsTSMargin As Worksheet
    Dim wsTSSymbolLookup As Worksheet
    Dim lastRowMargin As Long, lastRowLookup As Long
    Dim marginRow As Long
    Dim strippedSymbol As String
    Dim marginValue As Variant
    Dim SymbolLookup As String
    Dim LookupRow As Long
    Dim source As String, choice As String
    Dim initialmargin As Double
    Dim maintmargin As Double
    Dim defaultmargin As Double
    
     source = GetNamedRangeValue("Margin_Source")
     choice = GetNamedRangeValue("Margin_Choice")
       
     
    defaultmargin = wsDetails.Cells(summaryRow, FindColumnByHeader(headerArray, "Maint Overnight Margin")).value
     
    
    If source = "MutliWalk Margins" Then
     
        maintmargin = wsDetails.Cells(summaryRow, FindColumnByHeader(headerArray, "Maint Overnight Margin")).value
        initialmargin = wsDetails.Cells(summaryRow, FindColumnByHeader(headerArray, "Initial Overnight Margin")).value
    
    
    End If
   
         
    
    If source = "TradeStation Website" Then
        On Error Resume Next
        Set wsTSMargin = ThisWorkbook.Sheets("TradeStation Margins")
        On Error GoTo 0
    
        ' Exit and show error if the sheet doesn't exist
        If wsTSMargin Is Nothing Then
            MsgBox "Error: 'TS Margins' sheet does not exist.", vbExclamation
            Exit Sub
        End If
    
        lastRowMargin = wsTSMargin.Cells(wsTSMargin.rows.count, 1).End(xlUp).row
    
         ' Initialize margin value to be looked up
        marginValue = 0
        
        ' Search for the stripped symbol in the Margin Requirements sheet
        For marginRow = 2 To lastRowMargin
            If wsTSMargin.Cells(marginRow, 2).value = symbol Then ' Column 2 contains 'Symbol Root'
                
                
                
                maintmargin = (wsTSMargin.Cells(marginRow, 7).value * 1 + wsTSMargin.Cells(marginRow, 8).value * 1) / 2
                initialmargin = (wsTSMargin.Cells(marginRow, 5).value * 1 + wsTSMargin.Cells(marginRow, 6).value * 1) / 2
          
                
                           
                Exit For
                
            End If
        Next marginRow
        
         
    End If
    
    If source = "Interactive Brokers Website" Then
        On Error Resume Next
        Set wsIBMargin = ThisWorkbook.Sheets("InteractiveBrokers Margins")
        On Error GoTo 0
    
        ' Exit and show error if the sheet doesn't exist
        If wsIBMargin Is Nothing Then
            MsgBox "Error: 'IB Margins' sheet does not exist.", vbExclamation
            Exit Sub
        End If
            
        On Error Resume Next
        Set wsTSSymbolLookup = ThisWorkbook.Sheets("TS Symbol Lookup")
        On Error GoTo 0
    
        ' Exit and show error if the sheet doesn't exist
        If wsTSSymbolLookup Is Nothing Then
            MsgBox "Error: 'TS Symbol Lookup' sheet does not exist.", vbExclamation
            Exit Sub
        End If
    
        lastRowMargin = wsIBMargin.Cells(wsIBMargin.rows.count, 1).End(xlUp).row
        lastRowLookup = wsTSSymbolLookup.Cells(wsTSSymbolLookup.rows.count, 1).End(xlUp).row
        
        For LookupRow = 2 To lastRowLookup
            If wsTSSymbolLookup.Cells(LookupRow, 2).value = symbol Then ' Column 2 contains 'Symbol Root'
                SymbolLookup = wsTSSymbolLookup.Cells(LookupRow, 3).value
                Exit For
            End If
        Next LookupRow
        
         ' Initialize margin value to be looked up
        marginValue = 0
        
        ' Search for the stripped symbol in the Margin Requirements sheet
        For marginRow = 2 To lastRowMargin
            If wsIBMargin.Cells(marginRow, 1).value = SymbolLookup Then ' Column 2 contains 'Symbol Root'
                
                
                maintmargin = (wsIBMargin.Cells(marginRow, 4).value + wsIBMargin.Cells(marginRow, 6).value) / 2
                initialmargin = (wsIBMargin.Cells(marginRow, 3).value + wsIBMargin.Cells(marginRow, 5).value) / 2
              
                  Exit For
            End If
        Next marginRow
        
       
     
    End If

    
            
        If choice = "Overnight Maintenance" Then
                    marginValue = maintmargin
        ElseIf choice = "Overnight Initial" Then
                    marginValue = initialmargin
        ElseIf choice = "Average" Then
                    marginValue = (maintmargin + initialmargin) / 2
        End If
                 
         If marginValue <> 0 Then
            wsSummary.Cells(summaryRow, COL_MARGIN).value = marginValue
        Else
            wsSummary.Cells(summaryRow, COL_MARGIN).value = defaultmargin
        End If
    


End Sub

Function FindSectorValue(lookupValue As Variant) As Variant
    On Error Resume Next ' This suppresses errors if the value isn't found
    
    Dim result As Variant
    result = Application.VLookup(lookupValue, Range("SectorInput"), 2, False)
    
    ' If an error occurs (e.g., if not found), result will be an error value
    If IsError(result) Then
        FindSectorValue = "N/A"
    Else
        FindSectorValue = result
    End If
    
    On Error GoTo 0 ' Turn off error suppression
End Function




Sub CreateSummaryButtons(ws As Worksheet, colNumber As Long, currenttab As String)
    Dim btn As Object
    Dim captions As Variant
    Dim actions As Variant
    Dim i As Integer
    Dim leftOffset As Integer
    
    
    If currenttab = "Summary" Then
    ' Define button captions and their corresponding macros
    captions = Array("Control Tab", "Strategies Tab", "Inputs Tab", "Portfolio Tab", "Update Portfolio", "Save Status Changes")
    actions = Array("GoToControl", "GoToStrategies", "GoToInputs", "GoToPortfolio", "CreatePortfolioSummary", "UpdateStrategyStatuses")
    End If
    
    If currenttab = "Portfolio" Then
    ' Define button captions and their corresponding macros
    captions = Array("Control Tab", "Inputs Tab", "Check New Strats", "Summary Tab", "Update Portfolio", "Save Contract Changes")
    actions = Array("GoToControl", "GoToInputs", "IdentifyNewStrategiesAndContractChanges", "GoToSummary", "CreatePortfolioSummary", "UpdateStrategyContracts")
    End If
    
    
    ' Initial left offset
    leftOffset = 10
    
    ' Loop through captions and actions to create buttons dynamically
    For i = LBound(captions) To UBound(captions)
        Set btn = ws.Buttons.Add(left:=ws.Cells(1, colNumber).left + leftOffset, _
                                 top:=ws.Cells(1, 1).top, _
                                 Width:=70, Height:=35)
        With btn
            .Caption = captions(i)
            .OnAction = actions(i)
            
        End With
        ' Increment offset for the next button
        leftOffset = leftOffset + 75
    Next i
    
    ' Activate the specified worksheet
    ws.Activate
End Sub





' Sort Summary rows by: Status priority → Sector → Symbol → OOS Begin Date.
' Status order comes from GetStatusOrderNumber() / GetOrderedStatusList():
'   1. Port_Status (live portfolio)
'   2. Pass_Status (passing/backtest)
'   3+ StatusOptions (comma-separated named range — user-defined middle tiers)
'   last: BuyandHoldStatus  (always sorted to the bottom)
' Called automatically at the end of UpdateStrategySummaryWithArray.
' Can also be assigned to a button for on-demand re-sort after manual status changes.
Sub ReorderSummaryTab()
    On Error GoTo ErrorHandler

    Call InitializeColumnConstantsManually

    Dim wsSummary As Worksheet
    Dim wsStrategies As Worksheet
    Set wsSummary   = ThisWorkbook.Sheets("Summary")
    Set wsStrategies = ThisWorkbook.Sheets("Strategies")

    Application.ScreenUpdating = False

    Dim lastRow As Long, lastCol As Long
    lastRow = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row
    lastCol = wsSummary.Cells(1, wsSummary.Columns.count).End(xlToLeft).column

    If lastRow < 3 Then GoTo CleanExit  ' nothing to sort

    ' Write a temporary sort-key column (after lastCol) with numeric priority.
    ' GetStatusOrderNumber returns the position in GetOrderedStatusList:
    '   Port_Status=1, Pass_Status=2, ...StatusOptions..., BuyandHold=last, unknown=999.
    Dim helperCol As Long
    helperCol = lastCol + 1
    wsSummary.Cells(1, helperCol).value = "_SortPriority"
    Dim i As Long, sv As String
    For i = 2 To lastRow
        sv = wsSummary.Cells(i, COL_STATUS).value
        wsSummary.Cells(i, helperCol).value = GetStatusOrderNumber(sv)
    Next i

    ' 4-key sort using Excel's built-in engine
    With wsSummary.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsSummary.Range(wsSummary.Cells(1, helperCol), wsSummary.Cells(lastRow, helperCol)), _
                         Order:=xlAscending
        .SortFields.Add Key:=wsSummary.Range(wsSummary.Cells(1, COL_SECTOR), wsSummary.Cells(lastRow, COL_SECTOR)), _
                         Order:=xlAscending
        .SortFields.Add Key:=wsSummary.Range(wsSummary.Cells(1, COL_SYMBOL), wsSummary.Cells(lastRow, COL_SYMBOL)), _
                         Order:=xlAscending
        .SortFields.Add Key:=wsSummary.Range(wsSummary.Cells(1, COL_OOS_BEGIN_DATE), wsSummary.Cells(lastRow, COL_OOS_BEGIN_DATE)), _
                         Order:=xlAscending
        .SetRange wsSummary.Range(wsSummary.Cells(1, 1), wsSummary.Cells(lastRow, helperCol))
        .Header  = xlYes
        .MatchCase = False
        .Apply
    End With

    ' Remove the temporary helper column
    wsSummary.Columns(helperCol).Delete

    ' Renumber strategies sequentially (1, 2, 3 …) to match new sort order
    For i = 2 To lastRow
        wsSummary.Cells(i, COL_STRATEGY_NUMBER).value = i - 1
    Next i

    ' Sync symbol / timeframe back to Strategies tab so it stays consistent
    Dim lastRowStrat As Long
    lastRowStrat = wsStrategies.Cells(wsStrategies.rows.count, COL_STRAT_STRATEGY_NAME).End(xlUp).row
    Dim j As Long, stratName As String
    For i = 2 To lastRow
        stratName = wsSummary.Cells(i, COL_STRATEGY_NAME).value
        For j = 2 To lastRowStrat
            If StrComp(wsStrategies.Cells(j, COL_STRAT_STRATEGY_NAME).value, stratName, vbTextCompare) = 0 Then
                wsStrategies.Cells(j, COL_STRAT_SYMBOL).value    = wsSummary.Cells(i, COL_SYMBOL).value
                wsStrategies.Cells(j, COL_STRAT_TIMEFRAME).value = wsSummary.Cells(i, COL_TIMEFRAME).value
                Exit For
            End If
        Next j
    Next i

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub


Sub RankStrategiesInSummary()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rankColumn As Integer
    Dim sectorDict As Object, contractDict As Object
    Dim sectorData As Object, contractData As Object
    Dim i As Long, sector As Variant, contract As Variant, status As String
    Dim rankDirection As String, rankEligibility As String, Eligibility As String
    Dim namedRange As Range
    
    
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Summary")
    
    Call InitializeColumnConstantsManually  ' Uncomment if needed
    
    ' Find last row and last column
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).column
    
    ' Get the rank field column dynamically from the named range
    On Error Resume Next
    Set namedRange = ThisWorkbook.Names("rank_field").RefersToRange
    rankColumn = FindColumn(ws, namedRange.value)
    On Error GoTo 0
    
    If rankColumn = 0 Then
        MsgBox "Rank Field not found in Summary tab!", vbExclamation
        Exit Sub
    End If
    
    ' Get the ranking direction from named range
    On Error Resume Next
    rankDirection = ThisWorkbook.Names("rank_direction").RefersToRange.value
    On Error GoTo 0
    
    ' Get the ranking eligibility from named range
    On Error Resume Next
    rankEligibility = ThisWorkbook.Names("rank_eligibility").RefersToRange.value
    On Error GoTo 0
    
    ' Initialize dictionaries
    Set sectorDict = CreateObject("Scripting.Dictionary")
    Set contractDict = CreateObject("Scripting.Dictionary")
    
    ' Collect data by sector and contract
    For i = 2 To lastRow
        status = ws.Cells(i, COL_STATUS).value
        Eligibility = ws.Cells(i, COL_ELIGIBILITY).value
        sector = ws.Cells(i, COL_SECTOR).value
        contract = ws.Cells(i, COL_SYMBOL).value
        
        ' Check eligibility criteria
        If ShouldRank(status, Eligibility, rankEligibility) Then
            ' Collect sector data
            If Not sectorDict.Exists(sector) Then
                Set sectorDict(sector) = CreateObject("Scripting.Dictionary")
            End If
            sectorDict(sector)(i) = ws.Cells(i, rankColumn).value
            
            ' Collect contract data
            If Not contractDict.Exists(contract) Then
                Set contractDict(contract) = CreateObject("Scripting.Dictionary")
            End If
            contractDict(contract)(i) = ws.Cells(i, rankColumn).value
        End If
    Next i
    
    ' Rank by sector
    For Each sector In sectorDict.keys
        Set sectorData = sectorDict(sector)
        RankDictionary sectorData, ws, COL_SECTOR_RANK, rankDirection
    Next sector
    
    ' Rank by contract
    For Each contract In contractDict.keys
        Set contractData = contractDict(contract)
        RankDictionary contractData, ws, COL_SYMBOL_RANK, rankDirection
    Next contract
    
    ' Cleanup
    Set sectorDict = Nothing
    Set contractDict = Nothing
End Sub

' Function to find column index based on column name
Function FindColumn(ws As Worksheet, columnName As String) As Integer
    Dim cell As Range
    For Each cell In ws.rows(1).Cells
        If Trim(LCase(cell.value)) = Trim(LCase(columnName)) Then
            FindColumn = cell.column
            Exit Function
        End If
    Next cell
    FindColumn = 0
End Function

' Function to determine if a strategy should be ranked based on eligibility
Function ShouldRank(status As String, Eligibility As String, rankEligibility As String) As Boolean
    If rankEligibility = "Yes" Then
        ShouldRank = (Eligibility = "Yes") And (status = GetNamedRangeValue("Port_Status") Or status = GetNamedRangeValue("Pass_Status"))
    Else
        ShouldRank = (status = GetNamedRangeValue("Port_Status") Or status = GetNamedRangeValue("Pass_Status"))
    End If
End Function
Sub RankDictionary(dataDict As Object, ws As Worksheet, rankColumn As Long, rankDirection As String)
    Dim entries() As Variant
    Dim i As Long, n As Long
    Dim key As Variant
    Dim prevVal As Double
    Dim currentRank As Long, itemsAtRank As Long
    
    ' Count entries and create array
    n = dataDict.count
    ReDim entries(1 To n, 1 To 2) ' column 1 = rowIndex, column 2 = value
    
    ' Copy dictionary entries into array
    i = 0
    For Each key In dataDict.keys
        i = i + 1
        entries(i, 1) = CLng(key)        ' Row index
        entries(i, 2) = dataDict(key)    ' Value
    Next
    
    ' Sort the array using bubble sort (simple alternative to QuickSort for 2D array)
    ' Direction: "Increasing" means lowest value gets rank 1
    '            "Decreasing" means highest value gets rank 1
    Dim temp1 As Variant, temp2 As Variant
    Dim sorted As Boolean
    Dim j As Long
    
    For i = 1 To n - 1
        sorted = True
        For j = 1 To n - i
            If (rankDirection = "Increasing" And entries(j, 2) > entries(j + 1, 2)) Or _
               (rankDirection <> "Increasing" And entries(j, 2) < entries(j + 1, 2)) Then
                ' Swap values
                temp1 = entries(j, 1)
                temp2 = entries(j, 2)
                entries(j, 1) = entries(j + 1, 1)
                entries(j, 2) = entries(j + 1, 2)
                entries(j + 1, 1) = temp1
                entries(j + 1, 2) = temp2
                sorted = False
            End If
        Next j
        If sorted Then Exit For
    Next i
    
    ' Assign competition-style ranks
    currentRank = 1
    itemsAtRank = 1
    prevVal = entries(1, 2)
    ws.Cells(entries(1, 1), rankColumn).value = currentRank
    
    For i = 2 To n
        If entries(i, 2) = prevVal Then
            ' Same value gets same rank
            ws.Cells(entries(i, 1), rankColumn).value = currentRank
            itemsAtRank = itemsAtRank + 1
        Else
            ' New value gets next rank (competition style)
            currentRank = currentRank + itemsAtRank
            itemsAtRank = 1
            ws.Cells(entries(i, 1), rankColumn).value = currentRank
        End If
        prevVal = entries(i, 2)
    Next i
End Sub

' QuickSort algorithm for sorting arrays
Sub QuickSortdirection(arr As Variant, first As Long, last As Long, ascending As Boolean)
    Dim pivot As Double, i As Long, j As Long, temp As Double
    If first >= last Then Exit Sub
    pivot = arr((first + last) \ 2)
    i = first
    j = last
    Do
        If ascending Then
            Do While arr(i) < pivot: i = i + 1: Loop
            Do While arr(j) > pivot: j = j - 1: Loop
        Else
            Do While arr(i) > pivot: i = i + 1: Loop
            Do While arr(j) < pivot: j = j - 1: Loop
        End If
        
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop While i <= j
    QuickSortdirection arr, first, j, ascending
    QuickSortdirection arr, i, last, ascending
End Sub

Function DetermineStrategyType(longProfit As Double, longLoss As Double, shortProfit As Double, shortLoss As Double) As String
    ' Function to determine if a strategy is long-only, short-only, both, or has no trades
    ' Parameters:
    '   longProfit - Gross profit from long trades
    '   longLoss - Gross loss from long trades (as a positive number)
    '   shortProfit - Gross profit from short trades
    '   shortLoss - Gross loss from short trades (as a positive number)
    ' Returns:
    '   String indicating strategy type: "Long & Short", "Long Only", "Short Only", or "No Trades"
    
    ' Use Abs() to ensure we're comparing positive numbers
    Dim hasLongTrades As Boolean
    Dim hasShortTrades As Boolean
    
    hasLongTrades = (Abs(longProfit) > 0 Or Abs(longLoss) > 0)
    hasShortTrades = (Abs(shortProfit) > 0 Or Abs(shortLoss) > 0)
    
    If hasLongTrades And hasShortTrades Then
        DetermineStrategyType = "Long & Short"
    ElseIf hasLongTrades Then
        DetermineStrategyType = "Long Only"
    ElseIf hasShortTrades Then
        DetermineStrategyType = "Short Only"
    Else
        DetermineStrategyType = "No Trades"
    End If
End Function





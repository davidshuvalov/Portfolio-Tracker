Attribute VB_Name = "G_Create_Strategy_Tab"


Sub CreateStrategyTab(strategyNumber As Long)
    Dim wsSummary As Worksheet
    Dim newsheet As Worksheet
    Dim wsM2MEquity As Worksheet
    Dim summaryRow As Long
    Dim oosStartDate As Date
    Dim OOSEndDate As Date
    Dim expectedMonthlyProfit As Double
    Dim minMonthlyProfit As Double
    Dim cumulativeActualProfit As Double
    Dim currentdate As Date
    Dim endDate As Date
    Dim drawdown As Double
    Dim maxDrawdown As Double
    Dim monthCounter As Long
    Dim wasHidden As Boolean ' To check if the template was hidden
    Dim cumulativeExpectedProfit As Double
    Dim minCumulativeProfit As Double
    Dim actualMonthlyProfit As Double
    Dim monthlyDrawdown As Double
    Dim quittingPoint As Double
    Dim startdate As Date
    Dim longshort As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' Add near the beginning of the procedure after turning off screen updating
    Application.Calculation = xlCalculationManual
    
    ' Set the worksheets
         ' Check if "Summary" sheet exists and has data in row 2
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsSummary Is Nothing Then
        MsgBox "Error: 'Summary' sheet does not exist.", vbExclamation
        GoTo CleanExit
    End If

    ' Exit and show error if the sheet exists but has no data in row 2
    If wsSummary.Cells(2, COL_STRATEGY_NAME).value = "" Then
        MsgBox "Error: 'Summary' sheet exists but contains no data in row 2.", vbExclamation
        GoTo CleanExit
    End If
    
    Dim wsControl As Worksheet
        Set wsControl = ThisWorkbook.Sheets("Control")
    
    
    Set wsM2MEquity = ThisWorkbook.Sheets("DailyM2MEquity")
    
        ' Check if the template is hidden
    
    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    
    ' Find the row corresponding to the strategy number in the Summary sheet

    
    summaryRow = -99
        On Error Resume Next
        summaryRow = Application.Match(strategyNumber, wsSummary.Range(wsSummary.Cells(1, COL_STRATEGY_NUMBER), wsSummary.Cells(wsSummary.rows.count, COL_STRATEGY_NUMBER)), 0)
        On Error GoTo 0
        
            ' Exit and show error if the sheet doesn't exist
        If summaryRow = -99 Then
           MsgBox "Error: Cannot find Strategy Number " & strategyNumber & " in 'Summary' tab.", vbExclamation
           GoTo CleanExit
        End If
    
    Application.StatusBar = "Creating Strategy Tab: " & strategyNumber
    
    ' Check if the match was found
    If IsError(summaryRow) Then
        
        MsgBox "Strategy " & strategyNumber & " not found in Summary sheet."
        GoTo CleanExit
    End If
    
    
     ' Check if the strategy tab already exists, delete if it does
    On Error Resume Next
    Set newsheet = ThisWorkbook.Sheets("Strat - " & strategyNumber & " - Detail")
    On Error GoTo 0
    If Not newsheet Is Nothing Then
        Application.DisplayAlerts = False ' Disable alerts for sheet deletion
        newsheet.Delete ' Delete the existing strategy tab
        Application.DisplayAlerts = True ' Re-enable alerts
    End If
    
    ' Get the OOS start and end dates from the Summary sheet
    oosStartDate = wsSummary.Cells(summaryRow, COL_OOS_BEGIN_DATE).value ' OOS Begin Date from Summary
    OOSEndDate = wsSummary.Cells(summaryRow, COL_LAST_DATE_ON_FILE).value   ' OOS End Date from Summary
    
    ' Create a new worksheet for this strategy
    Set newsheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    
    ' Name the new sheet based on the StrategyNumber
    newsheet.name = "Strat - " & strategyNumber & " - Detail"
    newsheet.Tab.Color = RGB(242, 206, 239)
    
    
      ' Set white background color for the entire worksheet
    newsheet.Cells.Interior.Color = RGB(255, 255, 255)
    
   
    ' Set the zoom level of the active window to match the template's zoom
    ' Set the zoom level of the active window to 70%
    With ThisWorkbook.Windows(1)
        .Zoom = 70 ' Set zoom level to 70%
    End With
    ' Fill in the strategy-specific details from the Summary sheet
    
    
    With newsheet
    ' General formatting for all cells
    .Cells.Font.name = "Calibri"
    .Cells.Font.Size = 11
    .Columns("A:Z").ColumnWidth = 15

    ' Add headers and values with formatting
    .Cells(1, 1).value = "Strategy Name"
    .Cells(1, 2).value = wsSummary.Cells(summaryRow, COL_STRATEGY_NAME).value
    .Cells(2, 1).value = "Strategy Number"
    .Cells(2, 2).value = wsSummary.Cells(summaryRow, COL_STRATEGY_NUMBER).value
    .Cells(3, 1).value = "Symbol / Timeframe 1"
    .Cells(3, 2).value = wsSummary.Cells(summaryRow, COL_SYMBOL).value & " / " & wsSummary.Cells(summaryRow, COL_TIMEFRAME).value
    .Cells(4, 1).value = "Symbol / Timeframe 2"
    If Not IsEmpty(wsSummary.Cells(summaryRow, COL_DATA2_SYMBOL).value) And _
       Not IsEmpty(wsSummary.Cells(summaryRow, COL_DATA2_TIMEFRAME).value) Then
        .Cells(4, 2).value = wsSummary.Cells(summaryRow, COL_DATA2_SYMBOL).value & " / " & _
                             wsSummary.Cells(summaryRow, COL_DATA2_TIMEFRAME).value
    Else
        .Cells(4, 2).value = ""
    End If
    .Cells(5, 1).value = "Symbol / Timeframe 3"
    If Not IsEmpty(wsSummary.Cells(summaryRow, COL_DATA3_SYMBOL).value) And _
       Not IsEmpty(wsSummary.Cells(summaryRow, COL_DATA3_TIMEFRAME).value) Then
        .Cells(5, 2).value = wsSummary.Cells(summaryRow, COL_DATA3_SYMBOL).value & " / " & _
                             wsSummary.Cells(summaryRow, COL_DATA3_TIMEFRAME).value
    Else
        .Cells(5, 2).value = ""
    End If
    .Cells(6, 1).value = "Next Opt Date"
    .Cells(6, 2).value = wsSummary.Cells(summaryRow, COL_NEXT_OPT_DATE).value
    .Cells(6, 2).NumberFormat = "dd-mmm-yyyy"
    .Cells(7, 1).value = "Long / Short"
    .Cells(7, 2).value = wsSummary.Cells(summaryRow, COL_LONG_SHORT).value
    .Cells(8, 1).value = "Margin"
    .Cells(8, 2).value = wsSummary.Cells(summaryRow, COL_MARGIN).value
    .Cells(8, 2).NumberFormat = "$#,##0"
    .Cells(9, 1).value = "Notional Capital"
    .Cells(9, 2).value = wsSummary.Cells(summaryRow, COL_NOTIONAL_CAPITAL).value
    .Cells(9, 2).NumberFormat = "$#,##0"

    .Cells(10, 1).value = "Expected Annual Profit"
    .Cells(10, 2).value = wsSummary.Cells(summaryRow, COL_EXPECTED_ANNUAL_PROFIT).value
    .Cells(10, 2).NumberFormat = "$#,##0"
    .Cells(11, 1).value = "Actual Annual Profit"
    .Cells(11, 2).value = wsSummary.Cells(summaryRow, COL_ACTUAL_ANNUAL_PROFIT).value
    .Cells(11, 2).NumberFormat = "$#,##0"
    .Cells(12, 1).value = "Return Efficiency"
    .Cells(12, 2).value = wsSummary.Cells(summaryRow, COL_RETURN_EFFICIENCY).value
    .Cells(12, 2).NumberFormat = "$#%"
    .Cells(13, 1).value = "Expected Annual Return"
    .Cells(13, 2).value = wsSummary.Cells(summaryRow, COL_EXPECTED_ANNUAL_RETURN).value
    .Cells(13, 2).NumberFormat = "$#%"
    .Cells(14, 1).value = "Actual Annual Return"
    .Cells(14, 2).value = wsSummary.Cells(summaryRow, COL_ACTUAL_ANNUAL_RETURN).value
    .Cells(14, 2).NumberFormat = "$#%"

    .Cells(15, 1).value = "Profit Factor"
    .Cells(15, 2).value = wsSummary.Cells(summaryRow, COL_PROFIT_FACTOR).value
    .Cells(15, 2).NumberFormat = "0.00"
    .Cells(16, 1).value = "Long Profit Factor"
    .Cells(16, 2).value = wsSummary.Cells(summaryRow, COL_PROFIT_LONG_FACTOR).value
    .Cells(16, 2).NumberFormat = "0.00"
    .Cells(17, 1).value = "Short Profit Factor"
    .Cells(17, 2).value = wsSummary.Cells(summaryRow, COL_PROFIT_SHORT_FACTOR).value
    .Cells(17, 2).NumberFormat = "0.00"

    .Cells(18, 1).value = "Max Drawdown (IS)"
    .Cells(18, 2).value = wsSummary.Cells(summaryRow, COL_WORST_BACKTEST_DRAWDOWN).value
    .Cells(18, 2).NumberFormat = "$#,##0"
    .Cells(19, 1).value = "Max Drawdown (OOS)"
    .Cells(19, 2).value = wsSummary.Cells(summaryRow, COL_MAX_OOS_DRAWDOWN).value
    .Cells(19, 2).NumberFormat = "$#,##0"
    .Cells(20, 1).value = "Max Drawdown (IS + OOS)"
    .Cells(20, 2).value = wsSummary.Cells(summaryRow, COL_WORST_IS_OOS_DRAWDOWN).value
    .Cells(20, 2).NumberFormat = "$#,##0"

    .Cells(21, 1).value = "Average Drawdown (IS)"
    .Cells(21, 2).value = wsSummary.Cells(summaryRow, COL_AVG_BACKTEST_DRAWDOWN).value
    .Cells(21, 2).NumberFormat = "$#,##0"
    .Cells(22, 1).value = "Average Drawdown (OOS)"
    .Cells(22, 2).value = wsSummary.Cells(summaryRow, COL_AVG_OOS_DRAWDOWN).value
    .Cells(22, 2).NumberFormat = "$#,##0"
    .Cells(23, 1).value = "Avg Drawdown (IS + OOS)"
    .Cells(23, 2).value = wsSummary.Cells(summaryRow, COL_AVG_IS_OOS_DRAWDOWN).value
    .Cells(23, 2).NumberFormat = "$#,##0"

    .Cells(2, 3).value = "Out of Sample Period"
    .Cells(2, 4).value = wsSummary.Cells(summaryRow, COL_OOS_PERIOD).value
    .Cells(3, 3).value = "OOS Begin Date"
    .Cells(3, 4).value = wsSummary.Cells(summaryRow, COL_OOS_BEGIN_DATE).value
    .Cells(3, 4).NumberFormat = "dd-mmm-yyyy"
    .Cells(4, 3).value = "Last Date On File"
    .Cells(4, 4).value = wsSummary.Cells(summaryRow, COL_LAST_DATE_ON_FILE).value
    .Cells(4, 4).NumberFormat = "dd-mmm-yyyy"
    .Cells(5, 3).value = "Market Sector"
    .Cells(5, 4).value = wsSummary.Cells(summaryRow, COL_SECTOR).value
    .Cells(6, 3).value = "Session"
    .Cells(6, 4).value = wsSummary.Cells(summaryRow, COL_SESSION).value
    .Cells(7, 3).value = "Fitness"
    .Cells(7, 4).value = wsSummary.Cells(summaryRow, COL_FITNESS).value
    .Cells(8, 3).value = "WF Method"
    .Cells(8, 4).value = wsSummary.Cells(summaryRow, COL_ANCHORED).value
    .Cells(9, 3).value = "WF In/Out"
    .Cells(9, 4).value = wsSummary.Cells(summaryRow, COL_WF_IN_OUT).value
    .Cells(10, 3).value = "Status"
    .Cells(10, 4).value = wsSummary.Cells(summaryRow, COL_STATUS).value
    .Cells(11, 3).value = "Incubation Status"
    .Cells(11, 4).value = wsSummary.Cells(summaryRow, COL_INCUBATION_STATUS).value
    .Cells(12, 3).value = "Incubation Date"
    .Cells(12, 4).value = wsSummary.Cells(summaryRow, COL_INCUBATION_DATE).value
    .Cells(12, 4).NumberFormat = "dd-mmm-yyyy"
    .Cells(13, 3).value = "Quitting Status"
    .Cells(13, 4).value = wsSummary.Cells(summaryRow, COL_QUITTING_STATUS).value
    .Cells(14, 3).value = "Quitting Date"
    .Cells(14, 4).value = wsSummary.Cells(summaryRow, COL_QUITTING_DATE).value
    .Cells(14, 4).NumberFormat = "dd-mmm-yyyy"

    .Cells(15, 3).value = "Trades Per Year"
    .Cells(15, 4).value = wsSummary.Cells(summaryRow, COL_TRADES_PER_YEAR).value
    .Cells(16, 3).value = "Percent Time In Market"
    .Cells(16, 4).value = wsSummary.Cells(summaryRow, COL_PERCENT_TIME_IN_MARKET).value
    .Cells(16, 4).NumberFormat = "#%"
    .Cells(17, 3).value = "Winrate (OOS)"
    .Cells(17, 4).value = wsSummary.Cells(summaryRow, COL_OOS_WINRATE).value
    .Cells(17, 4).NumberFormat = "#%"
    .Cells(18, 3).value = "Winrate (IS + OOS)"
    .Cells(18, 4).value = wsSummary.Cells(summaryRow, COL_OVERALL_WINRATE).value
    .Cells(18, 4).NumberFormat = "#%"
    .Cells(19, 3).value = "Avg Trade (IS + OOS)"
    .Cells(19, 4).value = wsSummary.Cells(summaryRow, COL_AVG_IS_OOS_TRADE).value
    .Cells(19, 4).NumberFormat = "$#,##0"

    .Cells(20, 3).value = "Annual SD (IS)"
    .Cells(20, 4).value = wsSummary.Cells(summaryRow, COL_IS_ANNUAL_SD_IS).value
    .Cells(20, 4).NumberFormat = "$#,##0"
    .Cells(21, 3).value = "Annual SD (IS + OOS)"
    .Cells(21, 4).value = wsSummary.Cells(summaryRow, COL_IS_ANNUAL_SD_ISOOS).value
    .Cells(21, 4).NumberFormat = "$#,##0"

    .Cells(22, 3).value = "MW Monte Carlo (IS)"
    .Cells(22, 4).value = wsSummary.Cells(summaryRow, COL_IS_MONTE_CARLO).value
    .Cells(22, 4).NumberFormat = "0.0"
    .Cells(23, 3).value = "MW Monte Carlo (IS + OOS)"
    .Cells(23, 4).value = wsSummary.Cells(summaryRow, COL_IS_OOS_MONTE_CARLO).value
    .Cells(23, 4).NumberFormat = "0.0"
    
    .Cells(29, 10).value = wsSummary.Cells(summaryRow, COL_PROFIT_LAST_1_MONTH).value
        .Cells(30, 10).value = wsSummary.Cells(summaryRow, COL_PROFIT_LAST_3_MONTHS).value
        .Cells(31, 10).value = wsSummary.Cells(summaryRow, COL_PROFIT_LAST_6_MONTHS).value
        .Cells(32, 10).value = wsSummary.Cells(summaryRow, COL_PROFIT_LAST_9_MONTHS).value
        .Cells(33, 10).value = wsSummary.Cells(summaryRow, COL_PROFIT_LAST_12_MONTHS).value
        .Cells(34, 10).value = wsSummary.Cells(summaryRow, COL_PROFIT_SINCE_OOS_START).value
        
        .Cells(38, 10).value = wsSummary.Cells(summaryRow, COL_EFFICIENCY_LAST_1_MONTH).value
        .Cells(39, 10).value = wsSummary.Cells(summaryRow, COL_EFFICIENCY_LAST_3_MONTHS).value
        .Cells(40, 10).value = wsSummary.Cells(summaryRow, COL_EFFICIENCY_LAST_6_MONTHS).value
        .Cells(41, 10).value = wsSummary.Cells(summaryRow, COL_EFFICIENCY_LAST_9_MONTHS).value
        .Cells(42, 10).value = wsSummary.Cells(summaryRow, COL_EFFICIENCY_LAST_12_MONTHS).value
        .Cells(43, 10).value = wsSummary.Cells(summaryRow, COL_EFFICIENCY_SINCE_OOS_START).value
    
    
    startdate = wsSummary.Cells(summaryRow, COL_START_DATE).value

    
    
    ' Apply header formatting
    Dim headerRange As Range
    Set headerRange = .Range("A1:A23,C2:C23")
    With headerRange
        .Font.Bold = True
        .Interior.Color = RGB(185, 185, 185) ' Medium gray
        .HorizontalAlignment = xlLeft
    End With
    
    
    Set headerRange = .Range("B2:B23,D2:D23")
    With headerRange
        .Font.Bold = False
        .Interior.Color = RGB(220, 220, 220)
        .HorizontalAlignment = xlRight
    End With


     ' Define ranges as an array of addresses
    rngList = Array("A2:B23", "C2:D23", "I28:J34", "I37:J43", "A28:G28") ' Add your desired ranges here
    Dim rng As Range

    ' Loop through the range list and apply borders
    For i = LBound(rngList) To UBound(rngList)
        Set rng = newsheet.Range(rngList(i))
        
           With rng.Borders
        ' Top border
        With .Item(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
          
        End With
        
        ' Bottom border
        With .Item(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
      
        End With
        
        ' Left border
        With .Item(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
   
        End With
        
        ' Right border
        With .Item(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
  
        End With
        
        ' Inside vertical borders
        With .Item(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        ' Inside horizontal borders
        With .Item(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
       End With
    Next i

    




    ' Apply output formatting
End With


    'Monthly results header
    
    With newsheet

    ' Add headers and values with formatting
        .Cells(28, 1).value = "Month"
        .Cells(28, 2).value = "Cumulative Expected Monthly Profit"
        .Cells(28, 3).value = "Cumulative Min Monthly Profit"
        .Cells(28, 4).value = "Actual (Mark-To-Market)"
        .Cells(28, 5).value = "Cumulative Actual"
        .Cells(28, 6).value = "Quitting Point"
        .Cells(28, 7).value = "Drawdown"
        .Cells(28, 9).value = "Period"
        .Cells(28, 10).value = "Profit / Loss Dollars"
        .Cells(29, 9).value = "Last 1 Month"
        .Cells(30, 9).value = "Last 3 Months"
        .Cells(31, 9).value = "Last 6 Months"
        .Cells(32, 9).value = "Last 9 Months"
        .Cells(33, 9).value = "Last 12 Months"
        .Cells(34, 9).value = "Since OOS Start"
        
        .Cells(37, 9).value = "Period"
        .Cells(37, 10).value = "Efficiency"
        .Cells(38, 9).value = "Last 1 Month"
        .Cells(39, 9).value = "Last 3 Months"
        .Cells(40, 9).value = "Last 6 Months"
        .Cells(41, 9).value = "Last 9 Months"
        .Cells(42, 9).value = "Last 12 Months"
        .Cells(43, 9).value = "Since OOS Start"
        
        
        
        Set headerRange = .Range("A28:G28, I28:J28, I37:J37")
        With headerRange
            .Font.Bold = True
            .Interior.Color = RGB(185, 185, 185) ' Medium gray
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        
        Set headerRange = .Range("I29:I34, I38:I43")
        With headerRange
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220) ' Medium gray
            .HorizontalAlignment = xlLeft
            
        End With
        
        
        Set headerRange = .Range("A29:G1200")
        With headerRange
            .Font.Bold = False
            .Interior.Color = RGB(255, 255, 255) ' Medium gray
            .HorizontalAlignment = xlRight
            .NumberFormat = "$#,##0;[Red]-$#,##0"
            
            
        End With
        
        
        With headerRange.Borders
        ' Top border
        With .Item(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
          
        End With
        
        ' Bottom border
        With .Item(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
      
        End With
        
        ' Left border
        With .Item(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
   
        End With
        
        ' Right border
        With .Item(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
  
        End With
        
        ' Inside vertical borders
        With .Item(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        ' Inside horizontal borders
        With .Item(xlInsideHorizontal)
            .LineStyle = xlDash
            .Weight = xlThin
        End With
       End With
        
        
        Set headerRange = .Range("J29:J34")
        With headerRange
            .Font.Bold = False
            .Interior.Color = RGB(255, 255, 255) ' Medium gray
            .HorizontalAlignment = xlRight
            .NumberFormat = "$#,##0"
        End With
        
        Set headerRange = .Range("J38:J43")
        With headerRange
            .Font.Bold = False
            .Interior.Color = RGB(255, 255, 255) ' Medium gray
            .HorizontalAlignment = xlRight
            .NumberFormat = "#%"
        End With
        
        Set headerRange = .Range("A29:A1200")
        With headerRange
            .NumberFormat = "mmm-yyyy"
        End With
        
    End With
    
    Dim col As Integer
    
    newsheet.Columns(1).ColumnWidth = 34
    For col = 2 To 7 ' Columns containing profits
        newsheet.Columns(col).ColumnWidth = 30
    Next col
    
    newsheet.Columns(8).ColumnWidth = 5
    
    For col = 9 To 10 ' Columns containing profits
        newsheet.Columns(col).ColumnWidth = 20
    Next col

    ' Calculate expected monthly profit and min monthly profit
    expectedMonthlyProfit = wsSummary.Cells(summaryRow, COL_EXPECTED_ANNUAL_PROFIT).value / 12
    minMonthlyProfit = (wsSummary.Cells(summaryRow, COL_EXPECTED_ANNUAL_PROFIT).value * GetNamedRangeValue("Min_Incubation_Profit")) / 12

    ' Initialize cumulative profits and drawdown
    cumulativeActualProfit = 0
    maxDrawdown = 0
    monthCounter = 0
    LastColEquity = wsM2MEquity.Cells(1, wsM2MEquity.Columns.count).End(xlToLeft).column
    
    Dim strategyColumn As Long, counter As Long
    
     strategyColumn = -99
        On Error Resume Next
            For counter = 2 To LastColEquity
                If wsSummary.Cells(summaryRow, COL_STRATEGY_NAME).value = wsM2MEquity.Cells(1, counter).value Then
                    strategyColumn = counter
                    Exit For
                End If
                
            Next counter
        On Error GoTo 0
        
            ' Exit and show error if the sheet doesn't exist
        If strategyColumn = -99 Then
           
           MsgBox "Error: Cannot find  " & wsSummary.Cells(summaryRow, COL_STRATEGY_NAME).value & " in 'DailyM2MEquity' tab.", vbExclamation
           GoTo CleanExit
        End If
    
    
    
    
    
    ' Loop through each month from OOS Start Date to OOSEndDate
    currentdate = oosStartDate
    endDate = OOSEndDate
    
    ' Call the updated function with appropriate parameters
    monthCounter = GetMonthlyProfitAndMetrics(wsM2MEquity, currentdate, endDate, strategyColumn, expectedMonthlyProfit, _
                                        minMonthlyProfit, _
                                                      wsSummary.Cells(summaryRow, COL_WORST_BACKTEST_DRAWDOWN).value, _
                                                      monthlyDrawdown, quittingPoint, wsSummary, summaryRow, wsSummary.Cells(summaryRow, COL_IS_ANNUAL_SD_IS).value)



    ' Create a chart for Drawdown (using only columns N and Q)
     Set chartObj = newsheet.ChartObjects.Add(left:=newsheet.Cells(2, 5).left + (newsheet.Columns(5).Width / 4), top:=newsheet.Cells(2, 5).top, Width:=575, Height:=315)
With chartObj.chart
    Dim lastRow As Long
    lastRow = 28 + monthCounter ' Adjust the last row based on monthCounter

    .ChartType = xlLine
    .HasLegend = True
    .Legend.position = xlLegendPositionBottom ' Place legend at the bottom
    .HasTitle = True
    .chartTitle.text = newsheet.Cells(1, 2).value
    .chartTitle.Font.Size = 14 ' Set title font size

    ' Format x-axis
    With .Axes(xlCategory, xlPrimary)
        .TickLabels.NumberFormat = "dd-mmm-yyyy"
        .TickLabelPosition = xlLow ' Position x-axis labels at the bottom
        .TickLabels.Font.Size = 9 ' Set font size for date axis
        .Crosses = xlAutomatic ' Ensure x-axis crosses at default y-value (0)
    End With

    ' Format y-axis
    With .Axes(xlValue, xlPrimary)
        .TickLabels.NumberFormat = "$#,##0;-$#,##0" ' Format y-axis labels as currency
        .CrossesAt = 0 ' Set the x-axis to cross at y = 0
        .TickLabels.Font.Color = RGB(0, 0, 0) ' Default to black for positive numbers
        .TickLabels.Font.Size = 10 ' Set y-axis font size
        .HasMajorGridlines = True
    End With

    ' Series 1: Cumulative Expected Monthly Profit
    .SeriesCollection.NewSeries
    .SeriesCollection(1).name = "Cumulative Expected Monthly Profit"
    .SeriesCollection(1).XValues = newsheet.Range("A29:A" & lastRow)
    .SeriesCollection(1).values = newsheet.Range("B29:B" & lastRow)
    .SeriesCollection(1).Border.Weight = xlMedium
    .SeriesCollection(1).Border.Color = RGB(0, 0, 0) ' Black solid line

    ' Series 2: Cumulative Min Monthly Profit
    .SeriesCollection.NewSeries
    .SeriesCollection(2).name = "Cumulative Min Monthly Profit"
    .SeriesCollection(2).XValues = newsheet.Range("A29:A" & lastRow)
    .SeriesCollection(2).values = newsheet.Range("C29:C" & lastRow)
    .SeriesCollection(2).Border.Weight = xlMedium
    .SeriesCollection(2).Border.Color = RGB(0, 0, 0) ' Black dashed line
    .SeriesCollection(2).Border.LineStyle = xlDash

    ' Series 3: Cumulative Actual
    .SeriesCollection.NewSeries
    .SeriesCollection(3).name = "Cumulative Actual"
    .SeriesCollection(3).XValues = newsheet.Range("A29:A" & lastRow)
    .SeriesCollection(3).values = newsheet.Range("E29:E" & lastRow)
    .SeriesCollection(3).Border.Weight = xlMedium
    .SeriesCollection(3).Border.Color = RGB(0, 0, 255) ' Blue solid line

    ' Series 4: Quitting Point
    .SeriesCollection.NewSeries
    .SeriesCollection(4).name = "Quitting Point"
    .SeriesCollection(4).XValues = newsheet.Range("A29:A" & lastRow)
    .SeriesCollection(4).values = newsheet.Range("F29:F" & lastRow)
    .SeriesCollection(4).Border.Weight = xlMedium
    .SeriesCollection(4).Border.Color = RGB(255, 0, 0) ' Red dashed line
    .SeriesCollection(4).Border.LineStyle = xlDash
End With




    startdate = wsSummary.Cells(summaryRow, COL_START_DATE).value

    Dim m2mData As Variant
    Dim resultData() As Variant
    Dim row As Long
    
    Dim m2mProfit As Double
    Dim cumulativeProfit As Double
    Dim peakProfit As Double
    Dim drawdownpercent As Double
    Dim dataCollection As Collection
    Dim record As Variant
    Dim maxDD As Double
    Dim notional_capital As Double
    
    maxDD = wsSummary.Cells(summaryRow, COL_WORST_IS_OOS_DRAWDOWN).value
    notional_capital = wsSummary.Cells(summaryRow, COL_NOTIONAL_CAPITAL).value
    ' Initialize variables
    cumulativeProfit = 0
    peakProfit = 0
    Set dataCollection = New Collection
    
    ' Get the last row in wsM2MEquity
    lastRow = EndRowByCutoffSimple(wsM2MEquity, 1)
    
    ' Load wsM2MEquity data into an array for faster processing
    m2mData = wsM2MEquity.Range("A2", wsM2MEquity.Cells(lastRow, strategyColumn)).value
    
    ' Loop through each row in the m2mData array
    For row = 1 To UBound(m2mData, 1)
        ' Check if the date is within the specified range
        If m2mData(row, 1) >= startdate And m2mData(row, 1) <= OOSEndDate Then
            ' Calculate M2M profit, cumulative profit, and drawdown
            
            m2mProfit = m2mData(row, strategyColumn)
            cumulativeProfit = cumulativeProfit + m2mProfit
            peakProfit = WorksheetFunction.Max(peakProfit, cumulativeProfit)
            drawdown = peakProfit - cumulativeProfit
            drawdownpercent = drawdown / (peakProfit + notional_capital + 0.00001)
    
            ' Add a record to the collection
            record = Array(m2mData(row, 1), m2mProfit, cumulativeProfit, drawdown, drawdownpercent)
            dataCollection.Add record
        End If
    Next row
    
    ' If we have data, transfer the collection to a 2D array for writing
    If dataCollection.count > 0 Then
        ' Resize resultData array to fit the collection items
        ReDim resultData(1 To dataCollection.count, 1 To 5)
        
        ' Populate resultData from the collection
        For row = 1 To dataCollection.count
            resultData(row, 1) = dataCollection(row)(0) ' Date
            resultData(row, 2) = dataCollection(row)(1) ' M2M Profit
            resultData(row, 3) = dataCollection(row)(2) ' Cumulative Profit
            resultData(row, 4) = dataCollection(row)(3) ' Drawdown
            resultData(row, 5) = dataCollection(row)(4) ' DrawdownPercent
        Next row
        
        ' Write the resultData array to the new sheet starting from column N
        newsheet.Range("az2").Resize(UBound(resultData, 1), UBound(resultData, 2)).value = resultData
    Else
        
        Call Deletetab("Strat - " & strategyNumber & " - Detail")
        Call GoToSummary
        MsgBox "No out of sample data found.", vbInformation
        
        
        GoTo CleanExit
        
    End If
    ' Write the resultData array back to the new sheet starting from column P (14)
   ' newsheet.Range("az2").Resize(UBound(resultData, 1), UBound(resultData, 2)).value = resultData

  
    
    Dim chartProfit As ChartObject
    Dim chartDrawdown As ChartObject
    Dim chartRange As Range
    'Dim lastRow As Long
    Dim OOSStartRow As Long
    Dim oosStartDatePosition As Double
    
    Dim minScale As Double
    Dim maxScale As Double
    
    ' Determine the last row of data
    lastRow = newsheet.Cells(newsheet.rows.count, 52).End(xlUp).row  ' Column AZ is the 52th column
    
    ' Attempt to find the row where OOSStartDate matches in column N
    
    OOSStartRow = -99
        On Error Resume Next
        OOSStartRow = Application.Match(CLng(oosStartDate), newsheet.Range("az2:az" & lastRow), 0)
        On Error GoTo 0
        
            ' Exit and show error if the sheet doesn't exist
        If OOSStartRow = -99 Then
           
           MsgBox "Error: Cannot find out of sample start date in data", vbExclamation
           GoTo CleanExit
        End If
    
    
    
    ' Check if a match was found
    If IsError(OOSStartRow) Then
        MsgBox "OOS Start Date not found in the specified range.", vbExclamation
        GoTo CleanExit
    Else
        OOSStartRow = OOSStartRow + 1 ' Adjust for header if needed
    End If
    
    ' Create a chart for Cumulative Profits (using only columns N and P)
    Set chartProfit = newsheet.ChartObjects.Add(left:=newsheet.Cells(2, 12).left + newsheet.Cells(2, 12).Width / 3, top:=newsheet.Cells(2, 12).top, Width:=500, Height:=300)
    
    

With chartProfit.chart
    .ChartType = xlLine
    .HasLegend = False
    .HasTitle = True
    .chartTitle.text = "Cumulative Profit"

    ' Format x-axis and y-axis
    With .Axes(xlCategory, xlPrimary)
        .TickLabels.NumberFormat = "dd-mmm-yyyy"
        .TickLabelPosition = xlLow ' Position x-axis labels at the bottom
    End With

    With .Axes(xlValue, xlPrimary)
        .TickLabels.NumberFormat = "$#,##0;-$#,##0"
    End With

    ' Attempt to remove the secondary x-axis, if it exists
    On Error Resume Next
    .Axes(xlCategory, xlSecondary).Delete
    On Error GoTo 0


    ' Add the cumulative profit series explicitly
    .SeriesCollection.NewSeries
    .SeriesCollection(1).name = "Cumulative Profit"
    .SeriesCollection(1).XValues = newsheet.Range("az2:az" & lastRow)
    .SeriesCollection(1).values = newsheet.Range("bB2:bB" & lastRow)
    .SeriesCollection(1).Border.Weight = xlThin ' Set thinner line


    ' Shade the area after OOSStartDate
    With .PlotArea
        .Parent.Shapes.AddShape(msoShapeRectangle, .InsideLeft, _
                                .InsideTop, .InsideWidth * (OOSStartRow - 2) / (lastRow - 2), .InsideHeight).Select
        With Selection.ShapeRange.fill
            .ForeColor.RGB = RGB(211, 211, 211) ' Light gray shading
            .Transparency = 0.5
        End With
        Selection.ShapeRange.line.Visible = msoFalse ' Remove border from the shading box
        Selection.ShapeRange.ZOrder msoSendToBack
    End With
End With



' Create a chart for Drawdown (using only columns N and Q)
Set chartDrawdown = newsheet.ChartObjects.Add(left:=newsheet.Cells(2, 12).left + newsheet.Cells(2, 12).Width / 3, top:=newsheet.Cells(26, 12).top, Width:=500, Height:=300)
With chartDrawdown.chart
    .ChartType = xlLine
    .HasLegend = False
    .HasTitle = True
    .chartTitle.text = "Drawdown"

    ' Format x-axis and y-axis
    With .Axes(xlCategory, xlPrimary)
        .TickLabels.NumberFormat = "dd-mmm-yyyy"
        .TickLabelPosition = xlLow ' Position x-axis labels at the bottom
    End With

    With .Axes(xlValue, xlPrimary)
        .TickLabels.NumberFormat = "$#,##0;-$#,##0"
    End With

    ' Attempt to remove the secondary x-axis, if it exists
    On Error Resume Next
    .Axes(xlCategory, xlSecondary).Delete
    On Error GoTo 0

    ' Add the drawdown series explicitly
    .SeriesCollection.NewSeries
    .SeriesCollection(1).name = "Drawdown"
    .SeriesCollection(1).XValues = newsheet.Range("az2:az" & lastRow)
    .SeriesCollection(1).values = newsheet.Range("bC2:bC" & lastRow)
    .SeriesCollection(1).Border.Weight = xlThin ' Set thinner line
    .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(255, 0, 0) ' Red line for drawdown

    ' Shade the area after OOSStartDate
    With .PlotArea
        .Parent.Shapes.AddShape(msoShapeRectangle, .InsideLeft, _
                                .InsideTop, .InsideWidth * (OOSStartRow - 2) / (lastRow - 2), .InsideHeight).Select
        With Selection.ShapeRange.fill
            .ForeColor.RGB = RGB(211, 211, 211) ' Light gray shading
            .Transparency = 0.5
        End With
        Selection.ShapeRange.line.Visible = msoFalse ' Remove border from the shading box
        Selection.ShapeRange.ZOrder msoSendToBack
    End With
End With






' Create a chart for DrawdownPercent (using only columns Z and AD)
'Set chartDrawdown = newsheet.ChartObjects.Add(left:=newsheet.Cells(2, 19).left, top:=newsheet.Cells(48, 19).top, Width:=500, Height:=300)
'With chartDrawdown.chart
'    .ChartType = xlLine
'    .HasLegend = False
 '   .HasTitle = True
'    .chartTitle.Text = "Drawdown Percent"

    ' Format x-axis and y-axis
'    With .Axes(xlCategory, xlPrimary)
'        .TickLabels.NumberFormat = "dd-mmm-yyyy"
 '       .TickLabelPosition = xlLow ' Position x-axis labels at the bottom
 '   End With

'    With .Axes(xlValue, xlPrimary)
'        .TickLabels.NumberFormat = "0%" ' Format y-axis labels as currency
'    End With

    ' Attempt to remove the secondary x-axis, if it exists
'    On Error Resume Next
 '   .Axes(xlCategory, xlSecondary).Delete
'    On Error GoTo 0

    ' Add the drawdown series explicitly
'    .SeriesCollection.NewSeries
'    .SeriesCollection(1).name = "Drawdown"
'    .SeriesCollection(1).xValues = newsheet.Range("az2:az" & lastRow)
'    .SeriesCollection(1).values = newsheet.Range("bd2:bd" & lastRow)
'    .SeriesCollection(1).Border.Weight = xlThin ' Set thinner line
'    .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(255, 0, 0) ' Red line for drawdown

    ' Shade the area after OOSStartDate
'    With .PlotArea
'        .Parent.Shapes.AddShape(msoShapeRectangle, .InsideLeft, _
                                .InsideTop, .InsideWidth * (oosStartRow - 2) / (lastRow - 2), .InsideHeight).Select
'        With Selection.ShapeRange.fill
'            .ForeColor.RGB = RGB(211, 211, 211) ' Light gray shading
'            .Transparency = 0.5
'        End With
'        Selection.ShapeRange.line.Visible = msoFalse ' Remove border from the shading box
'        Selection.ShapeRange.ZOrder msoSendToBack
'    End With
'End With

' NEW CHART: Annual Profit/Loss
Set chartProfit = newsheet.ChartObjects.Add(left:=newsheet.Cells(2, 19).left, top:=newsheet.Cells(2, 19).top, Width:=500, Height:=300)
With chartProfit.chart
    .ChartType = xlColumnClustered
    .HasLegend = False
    .HasTitle = True
    .chartTitle.text = "Annual Profit/Loss"

    ' Format axes
    With .Axes(xlCategory, xlPrimary)
        .TickLabelPosition = xlLow
    End With

    With .Axes(xlValue, xlPrimary)
        .TickLabels.NumberFormat = "$#,##0;-$#,##0"
    End With

    ' Use the existing data from column Z (date) and column AA (m2mProfit)
    Dim yearlyData As Object
    Set yearlyData = CreateObject("Scripting.Dictionary")
    
    ' Determine the last row with data in column Z
    Dim dataLastRow As Long
    dataLastRow = newsheet.Cells(newsheet.rows.count, 52).End(xlUp).row
    
    ' Aggregate yearly profits from the data already in the sheet
    For i = 2 To dataLastRow
        Dim yearKey As String
        yearKey = Year(newsheet.Cells(i, 52).value) ' Column AZ (52) has the dates
        
        If Not yearlyData.Exists(yearKey) Then
            yearlyData.Add yearKey, newsheet.Cells(i, 53).value ' Column BA (53) has the M2M profits
        Else
            yearlyData(yearKey) = yearlyData(yearKey) + newsheet.Cells(i, 53).value
        End If
    Next i
    
    ' Create arrays for chart data
    Dim years() As String
    Dim profits() As Double
    ReDim years(1 To yearlyData.count)
    ReDim profits(1 To yearlyData.count)
    
    ' Fill arrays
    Dim idx As Integer, K As Variant
    idx = 1
    For Each K In yearlyData.keys
        years(idx) = K
        profits(idx) = yearlyData(K)
        idx = idx + 1
    Next K
    
    ' Add series to chart
    .SeriesCollection.NewSeries
    .SeriesCollection(1).name = "Annual P/L"
    .SeriesCollection(1).XValues = years
    .SeriesCollection(1).values = profits
    
    ' Set color for entire series
    .SeriesCollection(1).Format.fill.ForeColor.RGB = RGB(65, 105, 225) ' Royal Blue

   If Not IsError(OOSStartRow) And OOSStartRow > 0 Then
    ' Calculate fraction of year passed before OOS start date
    Dim yearFraction As Double
    Dim startOfYear As Date
    Dim endOfYear As Date
    
    startOfYear = DateSerial(Year(oosStartDate), 1, 1)
    endOfYear = DateSerial(Year(oosStartDate), 12, 31)
    
    ' Calculate what portion of the year has passed when OOS starts
    yearFraction = (oosStartDate - startOfYear) / (endOfYear - startOfYear)
    
    ' Find the oosYear position in the years array
    Dim oosYear As String
    oosYear = Year(oosStartDate)
    Dim oosYearIndex As Integer
    oosYearIndex = 0
    
    For i = 1 To UBound(years)
        If years(i) = oosYear Then
            oosYearIndex = i
            Exit For
        End If
    Next i
    
    ' Only add shading if we found the OOS year
    If oosYearIndex > 0 Then
        With .PlotArea
            ' Calculate the position for shading
            Dim shadeLeft As Double
            
            If oosYearIndex > 1 Then
                ' Factor in all years before OOS year, plus the fraction of OOS year
                shadeLeft = .InsideWidth * ((oosYearIndex - 1) + yearFraction) / UBound(years)
            Else
                ' If OOS starts in the first year of data, just use the year fraction
                shadeLeft = .InsideWidth * yearFraction / UBound(years)
            End If
            
            ' Create the shading shape
            .Parent.Shapes.AddShape(msoShapeRectangle, _
                                   .InsideLeft, .InsideTop, _
                                   shadeLeft, .InsideHeight).Select
            With Selection.ShapeRange.fill
                .ForeColor.RGB = RGB(211, 211, 211) ' Light gray shading
                .Transparency = 0.5
            End With
            Selection.ShapeRange.line.Visible = msoFalse ' Remove border
            Selection.ShapeRange.ZOrder msoSendToBack
        End With
    End If
End If
   
   
End With

' NEW CHART: Average Monthly Profit/Loss
Set chartDrawdown = newsheet.ChartObjects.Add(left:=newsheet.Cells(26, 19).left, top:=newsheet.Cells(26, 19).top, Width:=500, Height:=300)
With chartDrawdown.chart
    .ChartType = xlColumnClustered
    .HasLegend = False
    .HasTitle = True
    .chartTitle.text = "Average Monthly Profit/Loss by Month"
    ' Format axes
    With .Axes(xlCategory, xlPrimary)
        .TickLabelPosition = xlLow
    End With
    With .Axes(xlValue, xlPrimary)
        .TickLabels.NumberFormat = "$#,##0;-$#,##0"
    End With
    
    ' Prepare data structure to track yearly monthly data
    Dim monthlyData As Object
    Set monthlyData = CreateObject("Scripting.Dictionary")
    
    ' We'll need to track unique year-month combinations
    For i = 1 To 12
        Set monthlyData(i) = CreateObject("Scripting.Dictionary")
    Next i
    
    ' Collect monthly data by year
    'Dim currentDate As Date
    Dim currentMonth As Integer
    Dim currentYear As String
    Dim yearMonthKey As String
    
    ' First pass: Collect monthly totals by year
    Dim currentMonthTotal As Double
    Dim lastDate As Date
    lastDate = #1/1/1900#
    currentMonthTotal = 0
    
    ' Sort data by date if needed
    ' For this example, we'll assume data is already sorted by date
    
    For i = 2 To dataLastRow
        currentdate = newsheet.Cells(i, 52).value ' Column AZ (52) has the dates
        currentMonth = Month(currentdate)
        currentYear = Year(currentdate)
        
        ' If we're in a new month, save the previous month's total
        If (Month(lastDate) <> currentMonth Or Year(lastDate) <> currentYear) And lastDate <> #1/1/1900# Then
            yearMonthKey = Year(lastDate) & "-" & Month(lastDate)
            monthlyData(Month(lastDate))(yearMonthKey) = currentMonthTotal
            currentMonthTotal = 0
        End If
        
        ' Add current day's value to the running monthly total
        currentMonthTotal = currentMonthTotal + newsheet.Cells(i, 53).value ' Column BA (53) has the M2M profits
        lastDate = currentdate
    Next i
    
    ' Add the last month's data
    If lastDate <> #1/1/1900# Then
        yearMonthKey = Year(lastDate) & "-" & Month(lastDate)
        monthlyData(Month(lastDate))(yearMonthKey) = currentMonthTotal
    End If
    
    ' Month names for all 12 months
    Dim monthNames(1 To 12) As String
    monthNames(1) = "Jan": monthNames(2) = "Feb": monthNames(3) = "Mar"
    monthNames(4) = "Apr": monthNames(5) = "May": monthNames(6) = "Jun"
    monthNames(7) = "Jul": monthNames(8) = "Aug": monthNames(9) = "Sep"
    monthNames(10) = "Oct": monthNames(11) = "Nov": monthNames(12) = "Dec"
    
    ' Calculate averages of monthly totals
    Dim monthlyAvgs(1 To 12) As Double
    For i = 1 To 12
        If monthlyData(i).count > 0 Then
            Dim monthSum As Double
            monthSum = 0
            
            ' Sum up all years for this month
            Dim key As Variant
            For Each key In monthlyData(i).keys
                monthSum = monthSum + monthlyData(i)(key)
            Next key
            
            ' Calculate the average
            monthlyAvgs(i) = monthSum / monthlyData(i).count
        Else
            monthlyAvgs(i) = 0
        End If
    Next i
    
    ' Add to chart
    .SeriesCollection.NewSeries
    .SeriesCollection(1).name = "Avg Monthly P/L"
    .SeriesCollection(1).XValues = monthNames
    .SeriesCollection(1).values = monthlyAvgs
    
    ' Set color for entire series
    .SeriesCollection(1).Format.fill.ForeColor.RGB = RGB(65, 105, 225) ' Royal Blue
End With


    Call AddClosedTradeResultsChart(newsheet, wsSummary.Cells(summaryRow, COL_STRATEGY_NAME).value, startdate, OOSEndDate)
    

  ' Create delete button
    Dim btn As Object
    Set btn = newsheet.Buttons.Add(left:=newsheet.Cells(25, 1).left + 10, top:=newsheet.Cells(25, 1).top, Width:=140, Height:=25)
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteStrategyTab" ' Make sure to create this sub to handle deletion
    End With

 ' Create a button to return to the Summary page
    Set btn = newsheet.Buttons.Add(left:=newsheet.Cells(25, 2).left, _
                                    top:=newsheet.Cells(25, 2).top, _
                                    Width:=140, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary" ' Assign the macro to run when the button is clicked
    End With
 
 ' Create a button to return to the Portfolio page
    Set btn = newsheet.Buttons.Add(left:=newsheet.Cells(25, 3).left, _
                                    top:=newsheet.Cells(25, 3).top, _
                                    Width:=140, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio" ' Assign the macro to run when the button is clicked
    End With

 ' Create a button to return to the Portfolio page
    Set btn = newsheet.Buttons.Add(left:=newsheet.Cells(25, 4).left, _
                                    top:=newsheet.Cells(25, 4).top, _
                                    Width:=140, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl" ' Assign the macro to run when the button is clicked
    End With

 
    Set btn = newsheet.Buttons.Add(left:=newsheet.Cells(25, 5).left, _
                                    top:=newsheet.Cells(25, 5).top, _
                                    Width:=140, Height:=25)
    With btn
        .Caption = "Open Code"
        .OnAction = "ButtonClickHandlerCodeStrat" ' Assign the macro to run when the button is clicked
    End With
    
    Set btn = newsheet.Buttons.Add(left:=newsheet.Cells(25, 6).left, _
                                    top:=newsheet.Cells(25, 6).top, _
                                    Width:=140, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies" ' Assign the macro to run when the button is clicked
    End With
    
    Set btn = newsheet.Buttons.Add(left:=newsheet.Cells(25, 7).left, _
                                    top:=newsheet.Cells(25, 7).top, _
                                    Width:=140, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs" ' Assign the macro to run when the button is clicked
    End With



    With ThisWorkbook.Windows(1)
        .Zoom = 55 ' Set zoom level to 70%
    End With
    
    'hide data
    'newsheet.Columns("Z:AD").Hidden = True
    
    Call OrderVisibleTabsBasedOnList
    
    newsheet.Activate
    
    newsheet.Cells(1, 1).Select
    
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    Exit Sub

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    Call OrderVisibleTabsBasedOnList
    Call GoToControl
    
    Exit Sub

End Sub

Function GetMonthlyProfitAndMetrics(wsM2MEquity As Worksheet, startdate As Date, endDate As Date, strategyColumn As Long, _
                                     expectedMonthlyProfit As Double, minMonthlyProfit As Double, worstBacktestDrawdown As Double, _
                                     ByRef monthlyDrawdown As Double, ByRef quittingPoint As Double, _
                                     wsSummary As Worksheet, i As Long, sd As Long) As Double
    Dim lastRow As Long
    Dim totalProfit As Double
    Dim currentRow As Long
    Dim currentdate As Date
    Dim peakEquity As Double
    Dim currentEquity As Double
    Dim cumulativeActualProfit As Double
    Dim monthCounter As Long
    Dim currentDrawdown As Double
    Dim maxDrawdown As Double
    Dim quitting_method As String
    Dim SD_Multiple As Double
    Dim quitDollar As Double
    Dim quitPercent As Double
    
    ' Initialize variables
    totalProfit = 0
    peakEquity = 0
    currentEquity = 0
    monthlyDrawdown = 0
    cumulativeActualProfit = 0
    monthCounter = 0

    ' Find the last row in the DailyM2MEquity sheet
    lastRow = EndRowByCutoffSimple(wsM2MEquity, 1)


' Calculate the quitting point
    
        quitDollar = GetNamedRangeValue("Quit_Dollar")
        quitPercent = GetNamedRangeValue("Quit_percent")
        
    quitting_method = GetNamedRangeValue("Quitting_Method")
    SD_Multiple = GetNamedRangeValue("Quitting_SD_Multiple")
    quitPercent = GetNamedRangeValue("Quit_percent")
    quitDollar = GetNamedRangeValue("Quit_Dollar")

    expectedMonthlyProfit = IIf(expectedMonthlyProfit < 0, 0, expectedMonthlyProfit)

    ' Loop through each month from startDate to endDate
    currentdate = startdate
    Do While currentdate <= endDate
        Dim monthlyProfit As Double
        monthlyProfit = 0

        ' Loop through each row to find the matching month and year
        For currentRow = 2 To lastRow ' Assuming the first row is headers
            Dim rowDate As Date
            rowDate = wsM2MEquity.Cells(currentRow, 1).value ' Column 1 for date

            ' Check if the month and year match
            If Month(rowDate) = Month(currentdate) And Year(rowDate) = Year(currentdate) And rowDate >= startdate And rowDate <= endDate Then
                ' Sum the profit in the appropriate column (assumed to be in strategyColumn)
                monthlyProfit = monthlyProfit + wsM2MEquity.Cells(currentRow, strategyColumn).value * 1
                
                
                ' Update current equity
                currentEquity = currentEquity + wsM2MEquity.Cells(currentRow, strategyColumn).value * 1
                
                ' Update peak equity
                If currentEquity > peakEquity Then
                    peakEquity = currentEquity
                End If
                
                ' Calculate drawdown for the current month
                
                currentDrawdown = peakEquity - currentEquity
                
                
            End If
        Next currentRow

        ' Update total profit
        totalProfit = totalProfit + monthlyProfit

        
        If quitting_method = "Drawdown" Then
            quittingPoint = peakEquity - Application.Min(quitDollar, quitPercent * worstBacktestDrawdown)
        ElseIf quitting_method = "Standard Deviation" Then
            quittingPoint = ((expectedMonthlyProfit) * (monthCounter + 1) - Sqr(monthCounter + 1) * ((sd / Sqr(12)) * SD_Multiple))
            
        End If



        ' Record metrics in the summary sheet
        Cells(29 + monthCounter, 1).value = Format(currentdate, "mmm yyyy") ' Month label
        Cells(29 + monthCounter, 2).value = (expectedMonthlyProfit) * (monthCounter + 1) ' Cumulative Expected Profit
        Cells(29 + monthCounter, 3).value = (minMonthlyProfit) * (monthCounter + 1) ' Min Cumulative Profit
        Cells(29 + monthCounter, 4).value = monthlyProfit ' Actual Monthly Profit
        cumulativeActualProfit = cumulativeActualProfit + monthlyProfit
        Cells(29 + monthCounter, 5).value = cumulativeActualProfit ' Cumulative Actual Profit
        Cells(29 + monthCounter, 6).value = quittingPoint ' Quitting Point
        Cells(29 + monthCounter, 7).value = currentDrawdown ' Monthly Drawdown
        

        ' Move to the next month
        currentdate = DateAdd("m", 1, currentdate)
        monthCounter = monthCounter + 1
    Loop

    ' Return the monthCounter
    GetMonthlyProfitAndMetrics = monthCounter
End Function


   
   
 
    
   
   Sub AddClosedTradeResultsChart(newsheet As Worksheet, strategyName As String, startdate As Date, endDate As Date)
    ' This function adds a chart showing total, long, and short closed trade results to a strategy tab
    ' with enhanced error handling for cases with only long or only short trades
    '
    ' Parameters:
    '   newsheet - The strategy detail worksheet
    '   strategyName - The name of the strategy being analyzed
    '   startDate - The start date for data collection
    '   endDate - The end date for data collection
    
    Dim totalPnlDays As Variant
    Dim longTradesValues As Variant
    Dim shortTradesValues As Variant
    Dim chartResults As ChartObject
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim longTradeCount As Long, shortTradeCount As Long
    Dim hasLongTrades As Boolean, hasShortTrades As Boolean
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get the trade data using your existing function (which returns an array, not a collection)
    totalPnlDays = GetNonZeroPNLDays(strategyName, startdate, endDate)
    
    ' Check if we have data to display
    If UBound(totalPnlDays) < LBound(totalPnlDays) Then
        MsgBox "No closed trade data found for " & strategyName, vbInformation
        GoTo CleanExit
    End If
    
    ' Get long and short trade values with error handling
    On Error Resume Next
    longTradesValues = GetLongTradeValues(strategyName)
    If Err.Number <> 0 Then
        hasLongTrades = False
        ReDim longTradesValues(0)
    Else
        hasLongTrades = (UBound(longTradesValues) > LBound(longTradesValues))
    End If
    
    shortTradesValues = GetShortTradeValues(strategyName)
    If Err.Number <> 0 Then
        hasShortTrades = False
        ReDim shortTradesValues(0)
    Else
        hasShortTrades = (UBound(shortTradesValues) > LBound(shortTradesValues))
    End If
    On Error GoTo ErrorHandler
    
    ' Count actual trades in each category
    If hasLongTrades Then
        longTradeCount = UBound(longTradesValues) - LBound(longTradesValues) + 1
    Else
        longTradeCount = 0
    End If
    
    If hasShortTrades Then
        shortTradeCount = UBound(shortTradesValues) - LBound(shortTradesValues) + 1
    Else
        shortTradeCount = 0
    End If
    
    ' Set up column headers
    newsheet.Cells(1, 58).value = "Trade Count"
    newsheet.Cells(1, 59).value = "Total Trades"
    newsheet.Cells(1, 60).value = "Long Trades"
    newsheet.Cells(1, 61).value = "Short Trades"
    
    ' Format the headers
    newsheet.Range("BF1:BI1").Font.Bold = True
    newsheet.Range("BF1:BI1").Interior.Color = RGB(185, 185, 185)
    
    ' Populate the total trade data
    i = 2 ' Start from row 2
    For j = LBound(totalPnlDays) To UBound(totalPnlDays)
        newsheet.Cells(i, 58).value = i - 1 ' Trade count
        newsheet.Cells(i, 59).value = totalPnlDays(j) ' Total trade PNL
        i = i + 1
    Next j
    
    lastRow = i - 1
    
      ' Populate long trades data if available
    If hasShortTrades And shortTradeCount > 0 Then
        ' Fill in actual values for the range where we have long trades
        For i = 2 To Min(longTradeCount + 1, lastRow)
            newsheet.Cells(i, 60).value = longTradesValues(LBound(longTradesValues) + i - 2)
        Next i
        
        ' Fill in #N/A for the rest of the rows
        If longTradeCount + 1 < lastRow Then
            For i = longTradeCount + 2 To lastRow
                newsheet.Cells(i, 60).Formula = "=NA()"
            Next i
        End If
    Else
        ' If no long trades, duplicate the total closed trades
        For i = 2 To lastRow
            newsheet.Cells(i, 60).value = newsheet.Cells(i, 59).value
        Next i
    End If
    
    ' Populate short trades data if available
    If hasLongTrades And longTradeCount > 0 Then
        ' Fill in actual values for the range where we have short trades
        For i = 2 To Min(shortTradeCount + 1, lastRow)
            newsheet.Cells(i, 61).value = shortTradesValues(LBound(shortTradesValues) + i - 2)
        Next i
        
        ' Fill in #N/A for the rest of the rows
        If shortTradeCount + 1 < lastRow Then
            For i = shortTradeCount + 2 To lastRow
                newsheet.Cells(i, 61).Formula = "=NA()"
            Next i
        End If
    Else
        ' If no short trades, duplicate the total closed trades
        For i = 2 To lastRow
            newsheet.Cells(i, 61).value = newsheet.Cells(i, 59).value
        Next i
    End If
    
    ' Format the data columns
    newsheet.Range("BF2:BF" & lastRow).NumberFormat = "0" ' Trade count as integer
    newsheet.Range("BG2:BI" & lastRow).NumberFormat = "$#,##0.00;[Red]$#,##0.00"
    
    ' Create cumulative columns
    newsheet.Cells(1, 62).value = "Cumul Total"
    newsheet.Cells(1, 63).value = "Cumul Long"
    newsheet.Cells(1, 64).value = "Cumul Short"
    
    ' Format the cumulative headers
    newsheet.Range("BJ1:BL1").Font.Bold = True
    newsheet.Range("BJ1:BL1").Interior.Color = RGB(185, 185, 185)
    
    ' Calculate cumulative values
    Dim totalCumul As Double, longCumul As Double, shortCumul As Double
    totalCumul = 0
    longCumul = 0
    shortCumul = 0
    
    For i = 2 To lastRow
        ' Total trades cumulative
        If Not IsEmpty(newsheet.Cells(i, 59).value) Then
            totalCumul = totalCumul + newsheet.Cells(i, 59).value
            newsheet.Cells(i, 62).value = totalCumul
        Else
            newsheet.Cells(i, 62).value = totalCumul
        End If
        
        ' Long trades cumulative
        If Not IsEmpty(newsheet.Cells(i, 60).value) And Not IsError(newsheet.Cells(i, 60).value) Then
            longCumul = longCumul + newsheet.Cells(i, 60).value
            newsheet.Cells(i, 63).value = longCumul
        Else
            If hasLongTrades And i <= longTradeCount + 1 Then
                newsheet.Cells(i, 63).value = longCumul
            Else
                newsheet.Cells(i, 63).Formula = "=NA()"
            End If
        End If
        
        ' Short trades cumulative
        If Not IsEmpty(newsheet.Cells(i, 61).value) And Not IsError(newsheet.Cells(i, 61).value) Then
            shortCumul = shortCumul + newsheet.Cells(i, 61).value
            newsheet.Cells(i, 64).value = shortCumul
        Else
            If hasShortTrades And i <= shortTradeCount + 1 Then
                newsheet.Cells(i, 64).value = shortCumul
            Else
                newsheet.Cells(i, 64).Formula = "=NA()"
            End If
        End If
    Next i
    
    ' Format the cumulative columns
    newsheet.Range("BJ2:BL" & lastRow).NumberFormat = "$#,##0.00;[Red]$#,##0.00"
    
    ' Create a chart for the cumulative trade results at specified location
    Set chartResults = newsheet.ChartObjects.Add(left:=newsheet.Cells(2, 12).left + newsheet.Cells(2, 12).Width / 3, top:=newsheet.Cells(48, 12).top, Width:=500, Height:=300)
  
    
    With chartResults.chart
        .ChartType = xlLine
        .HasLegend = True
        .Legend.position = xlLegendPositionBottom
        .HasTitle = True
        .chartTitle.text = "Closed Trade Cumulative Profit"
        
        ' Format x-axis (trade count)
        With .Axes(xlCategory, xlPrimary)
            .TickLabels.NumberFormat = "0" ' Integer format for trade count
            .TickLabelPosition = xlLow
            .TickLabels.Font.Size = 9
            .Crosses = xlAutomatic
            .HasTitle = True
            .AxisTitle.text = "Trade Count"
            .AxisTitle.Font.Size = 10
        End With
        
        ' Format y-axis
        With .Axes(xlValue, xlPrimary)
            .TickLabels.NumberFormat = "$#,##0;[Red]-$#,##0"
            .CrossesAt = 0
            .TickLabels.Font.Color = RGB(0, 0, 0)
            .TickLabels.Font.Size = 10
            .HasMajorGridlines = True
            .HasTitle = True
            .AxisTitle.text = "Cumulative P&L"
            .AxisTitle.Font.Size = 10
        End With
        
        ' Add series for total cumulative trades
        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = "Total Closed Trades"
        .SeriesCollection(1).XValues = newsheet.Range("BF2:BF" & lastRow)
        .SeriesCollection(1).values = newsheet.Range("BJ2:BJ" & lastRow)
        .SeriesCollection(1).Border.Weight = xlMedium
        .SeriesCollection(1).Border.Color = RGB(0, 0, 255) ' Blue
        
        ' Add series for long cumulative trades if there are any
        If hasLongTrades Then
            .SeriesCollection.NewSeries
            .SeriesCollection(2).name = "Long Closed Trades"
            .SeriesCollection(2).XValues = newsheet.Range("BF2:BF" & lastRow)
            .SeriesCollection(2).values = newsheet.Range("BK2:BK" & lastRow)
            .SeriesCollection(2).Border.Weight = xlMedium
            .SeriesCollection(2).Border.Color = RGB(0, 128, 0) ' Green
        End If
        
        ' Add series for short cumulative trades if there are any
        If hasShortTrades Then
            Dim seriesIndex As Integer
            seriesIndex = IIf(hasLongTrades, 3, 2)
            
            .SeriesCollection.NewSeries
            .SeriesCollection(seriesIndex).name = "Short Closed Trades"
            .SeriesCollection(seriesIndex).XValues = newsheet.Range("BF2:BF" & lastRow)
            .SeriesCollection(seriesIndex).values = newsheet.Range("BL2:BL" & lastRow)
            .SeriesCollection(seriesIndex).Border.Weight = xlMedium
            .SeriesCollection(seriesIndex).Border.Color = RGB(255, 0, 0) ' Red
        End If
       
    End With
    
    ' No chart statistics labels
    
    ' Hide the data columns (optional)
    'newsheet.Columns("BF:BL").Hidden = True
    
CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error occurred while creating the trade chart: " & Err.Description, vbExclamation, "Chart Creation Error"
    Resume CleanExit
End Sub

Function GetLongTradeValues(strategyName As String) As Variant
    ' Function to get all long trade values for a strategy
    ' Returns an array of trade values
    
    Dim wsLongTrades As Worksheet
    Dim col As Long, lastRow As Long, i As Long
    Dim tradeValues() As Double
    
    On Error GoTo ErrorHandler
    
    ' Check if Long_Trades sheet exists
    On Error Resume Next
    Set wsLongTrades = ThisWorkbook.Sheets("Long_Trades")
    On Error GoTo 0
    
    If wsLongTrades Is Nothing Then
        ' Return an empty array if sheet doesn't exist
        ReDim tradeValues(0)
        GetLongTradeValues = tradeValues
        Exit Function
    End If
    
    ' Find the column for this strategy
    col = 0
    For i = 1 To wsLongTrades.Cells(1, wsLongTrades.Columns.count).End(xlToLeft).column
        If wsLongTrades.Cells(1, i).value = strategyName Then
            col = i
            Exit For
        End If
    Next i
    
    If col = 0 Then
        ' Return an empty array if strategy not found
        ReDim tradeValues(0)
        GetLongTradeValues = tradeValues
        Exit Function
    End If
    
    ' Find the last row with data
    lastRow = wsLongTrades.Cells(wsLongTrades.rows.count, col).End(xlUp).row
    
    ' Check if there's any data (if lastRow <= 1, there's only the header row or less)
    If lastRow <= 1 Then
        ReDim tradeValues(0)
        GetLongTradeValues = tradeValues
        Exit Function
    End If
    
    ' Initialize the array to hold trade values
    ReDim tradeValues(1 To lastRow - 1) ' Array for just the data rows (no header)
    
    ' Extract values (skip header row)
    For i = 2 To lastRow
        If IsNumeric(wsLongTrades.Cells(i, col).value) Then
            tradeValues(i - 1) = wsLongTrades.Cells(i, col).value
        Else
            tradeValues(i - 1) = 0
        End If
    Next i
    
    GetLongTradeValues = tradeValues
    Exit Function
    
ErrorHandler:
    ' Return an empty array if there's an error
    ReDim tradeValues(0)
    GetLongTradeValues = tradeValues
End Function

Function GetShortTradeValues(strategyName As String) As Variant
    ' Function to get all short trade values for a strategy
    ' Returns an array of trade values
    
    Dim wsShortTrades As Worksheet
    Dim col As Long, lastRow As Long, i As Long
    Dim tradeValues() As Double
    
    On Error GoTo ErrorHandler
    
    ' Check if Short_Trades sheet exists
    On Error Resume Next
    Set wsShortTrades = ThisWorkbook.Sheets("Short_Trades")
    On Error GoTo 0
    
    If wsShortTrades Is Nothing Then
        ' Return an empty array if sheet doesn't exist
        ReDim tradeValues(0)
        GetShortTradeValues = tradeValues
        Exit Function
    End If
    
    ' Find the column for this strategy
    col = 0
    For i = 1 To wsShortTrades.Cells(1, wsShortTrades.Columns.count).End(xlToLeft).column
        If wsShortTrades.Cells(1, i).value = strategyName Then
            col = i
            Exit For
        End If
    Next i
    
    If col = 0 Then
        ' Return an empty array if strategy not found
        ReDim tradeValues(0)
        GetShortTradeValues = tradeValues
        Exit Function
    End If
    
    ' Find the last row with data
    lastRow = wsShortTrades.Cells(wsShortTrades.rows.count, col).End(xlUp).row
    
    ' Check if there's any data (if lastRow <= 1, there's only the header row or less)
    If lastRow <= 1 Then
        ReDim tradeValues(0)
        GetShortTradeValues = tradeValues
        Exit Function
    End If
    
    ' Initialize the array to hold trade values
    ReDim tradeValues(1 To lastRow - 1) ' Array for just the data rows (no header)
    
    ' Extract values (skip header row)
    For i = 2 To lastRow
        If IsNumeric(wsShortTrades.Cells(i, col).value) Then
            tradeValues(i - 1) = wsShortTrades.Cells(i, col).value
        Else
            tradeValues(i - 1) = 0
        End If
    Next i
    
    GetShortTradeValues = tradeValues
    Exit Function
    
ErrorHandler:
    ' Return an empty array if there's an error
    ReDim tradeValues(0)
    GetShortTradeValues = tradeValues
End Function

Function Min(val1 As Long, val2 As Long) As Long
    ' Simple function to return the minimum of two values
    If val1 < val2 Then
        Min = val1
    Else
        Min = val2
    End If
End Function
   
   
   
   
   

' Helper function to check if a number is in the collection
Function IsInCollection(col As Collection, value As String) As Boolean
    Dim i As Long
    On Error Resume Next
    For i = 1 To col.count
        If col(i) = value Then
            IsInCollection = True
            Exit Function
        End If
    Next i
    IsInCollection = False
    On Error GoTo 0
End Function




Sub CreateAllLiveStrategiesTab()

    CreateStatusTab GetNamedRangeValue("Port_Status"), ""
End Sub

Sub CreateAllPassedTab()
   
    CreateStatusTab "Passed", "Passed (NT)"
    
End Sub

Sub CreateAllIncubationTab()
    
    CreateStatusTab "Incubation", ""
End Sub

Sub CreateAllFailedTab()
    
    CreateStatusTab "Failed", ""
End Sub

Sub CreateAllLiveStrategiesCodeTab()
  
    CreateStatusCodeTab GetNamedRangeValue("Port_Status"), ""
End Sub

Sub CreateAllPassedCodeTab()
    
    CreateStatusCodeTab "Passed", "Passed (NT)"
End Sub

Sub CreateAllIncubationCodeTab()
    
    CreateStatusCodeTab "Incubation", ""
End Sub

Sub CreateAllFailedCodeTab()
   
    CreateStatusCodeTab "Failed", ""
End Sub


Sub CreateStatusTab(status1 As String, status2 As String)
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tabsCreated As Long

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    
   Application.ScreenUpdating = False
    Application.EnableEvents = False
    
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

    ' Initialize counter for tabs created
    tabsCreated = 0

    ' Loop through the rows to create tabs based on status
    For i = 2 To lastRow
        If status1 = wsSummary.Cells(i, COL_STATUS).value Or status2 = wsSummary.Cells(i, COL_STATUS).value Then
            CreateStrategyTab wsSummary.Cells(i, COL_STRATEGY_NUMBER).value
            tabsCreated = tabsCreated + 1
              
            Application.StatusBar = "Creating Strategy Tab: " & wsSummary.Cells(i, COL_STRATEGY_NUMBER).value
        End If
    Next i
    Application.StatusBar = False
    ' Display a message based on the number of tabs created
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
     Dim status As String
    If status2 <> "" Then status = status1 & " or " & status2 Else status = status1
    
    If tabsCreated > 0 Then
        MsgBox tabsCreated & " tabs created successfully.", vbInformation
    Else
        MsgBox "No tabs were created. No entries found with status '" & status & "'.", vbExclamation
    End If
End Sub







Sub ButtonClickHandlerCodeStrat()
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
    gStrategyName = wsCurrentSheet.Cells(1, 2).value
    gStrategyNumber = wsCurrentSheet.Cells(2, 2).value
    ' Call the macro to create the strategy tab
    OpenStrategyCodeFile gStrategyName, gStrategyNumber, "tab"
    
End Sub


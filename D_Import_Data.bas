Attribute VB_Name = "D_Import_Data"
Private stratData As Variant
Private dataLoaded As Boolean


Sub OptimizeDataProcessing()
    ' Main procedure for optimizing date population and data handling

    Dim uniqueDatesDict As Object
    Dim combinedData As Object
    Dim tradeDataDict As Object
    Dim tradeLongDataDict As Object
    Dim tradeShortDataDict As Object
    Dim PNL() As Double
    Dim allDatesList() As Date
    Dim colHeaders() As String
    Dim colLongHeaders() As String    ' Headers for long trades
    Dim colShortHeaders() As String   ' Headers for short trades
    Dim colATRHeaders() As String
    Dim ws As Worksheet
    Dim folderLocations As Worksheet
    Dim wsStrategies As Worksheet
    Dim lastRow As Long, colIndex As Long, colLongIndex As Long, colShortIndex As Long, colATRIndex As Long, i As Long, j As Long
    Dim folderPath As String, csvFile As String, FileNameOnly As String
    Dim dataArray As Variant, ErrStr As String, numRows As Long, numCols As Long
    Dim minDate As Date, MaxDate As Date
    Dim ExcelDateFormat As String
    Dim headerName As String
    Dim ATR_Flag As Long
    Dim BuyAndHoldFilesFound As Boolean
    Dim errorMessages As String
    
     
    ATR_Flag = 0
    BuyAndHoldFilesFound = False
    errorMessages = ""
    
    
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        Exit Sub
    End If
    
    
    ' Check if BuyandHoldStatus is empty
    If GetNamedRangeValue("BuyandHoldStatus") = "" Then
        errorMessages = errorMessages & "Warning: BuyandHoldStatus is empty. ATR calculations will not be performed." & vbCrLf
    End If
    
    
    
    Call InitializeColumnConstantsManually
    
    ' Disable updates for performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Initializing data processing..."
    
    ' Get "MW Folder Locations" sheet
    On Error Resume Next
    Set folderLocations = ThisWorkbook.Sheets("MW Folder Locations")
    On Error GoTo 0
    If folderLocations Is Nothing Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Error: 'MW Folder Locations' sheet is missing.", vbExclamation
        Exit Sub
    End If
    
    If folderLocations.Cells(2, 1).value = "" Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Error: 'MW Folder Locations' sheet is empty.", vbExclamation
        Exit Sub
    End If
    
    If folderLocations.Cells(1, 1).value <> "Folder Count" Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Please correct issues highlighted in the 'MW Folder Locations' tab first before continuing...", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsStrategies = ThisWorkbook.Sheets("Strategies")
    On Error GoTo 0
    If wsStrategies Is Nothing Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Error: 'Strategies' sheet is missing.", vbExclamation
        Exit Sub
    End If
    
    ' 1?? Load once from sheet:
    InitializeStrategyCache Worksheets("Strategies")
    
    
    ' Initialize objects
    Set uniqueDatesDict = CreateObject("Scripting.Dictionary")
    Set combinedData = CreateObject("Scripting.Dictionary")
    Set tradeDataDict = CreateObject("Scripting.Dictionary")
    Set tradeLongDataDict = CreateObject("Scripting.Dictionary")   ' Initialize long trades dictionary
    Set tradeShortDataDict = CreateObject("Scripting.Dictionary")  ' Initialize short trades dictionary
    
    Dim latestPositionsDict As Object
    Set latestPositionsDict = CreateObject("Scripting.Dictionary")
    
    ' Get date format
    ExcelDateFormat = GetNamedRangeValue("DateFormat")
    
    
      
   
    Call Deletetab("TrueRanges")
    Call Deletetab("TradePNL")
    Call Deletetab("AverageTrueRange")
    Call Deletetab("InMarketLong")
    Call Deletetab("InMarketShort")
    Call Deletetab("DailyM2MEquity")
    Call Deletetab("ClosedTradePNL")
    Call Deletetab("Walkforward Details")
    Call Deletetab("ContractMultiples")
    Call Deletetab("LatestPositionData")
    
    ' Loop through each folder path
    lastRow = folderLocations.Cells(folderLocations.rows.count, 4).End(xlUp).row
    lastRowStrat = wsStrategies.Cells(wsStrategies.rows.count, 1).End(xlUp).row
    colIndex = 0
    colLSIndex = 0
    colATRIndex = 0
    ReDim colHeaders(1 To 1)
    ReDim colLSHeaders(1 To 1)
    ReDim colATRHeaders(1 To 1)
    i = 2
    While i <= lastRow
        
        folderPath = Trim(folderLocations.Cells(i, 4).value) & "\Walkforward Files\"
        If folderPath <> "" Then
            csvFile = Dir(GetShortPath(folderPath & "*EquityData.csv"))
            Do While csvFile <> ""
                FileNameOnly = left(csvFile, InStr(csvFile, " EquityData") - 1)
                
                ' Read data from file
                dataArray = getDataFromFile(GetShortPath(folderPath & csvFile), ",", ErrStr, numRows, numCols, ExcelDateFormat)
                If ErrStr <> "" Then
                    MsgBox ErrStr, vbExclamation
                    Exit Sub
                End If
                
                ' Process data
                ProcessEquityData dataArray, uniqueDatesDict, combinedData, FileNameOnly, numRows, colIndex, colHeaders
                i = i + 1
                csvFile = Dir
                Application.StatusBar = "Importing equity data in array: " & i - 1 & " of " & lastRow & "..."
            Loop
            
        End If
        
    Wend
    
    ' -------------------------------------------------------------------------
    ' Single TradeData pass — ALL strategies, including Buy & Hold.
    '
    ' Mirrors the EquityData loop: scan every *TradeData.csv in the folder,
    ' advance i once per file (matching the one-row-per-file structure that
    ' GetFolderData wrote to folderLocations).
    '
    ' Every strategy gets long/short splits (ProcessLSTradeData).
    ' BnH strategies additionally get ATR/PNL calculations (ProcessTradeData),
    ' using only the contract symbol as the header (ExtractContractName).
    '
    ' If a folder has no TradeData files (e.g. a strategy that only exports
    ' equity curves), i is advanced once so the outer loop keeps moving.
    ' -------------------------------------------------------------------------
    i = 2
    While i <= lastRow
        folderPath = Trim(folderLocations.Cells(i, 4).value) & "\Walkforward Files\"

        csvFile = Dir(GetShortPath(folderPath & "*TradeData.csv"))
        If csvFile <> "" Then
            Do While csvFile <> ""
                FileNameOnly = left(csvFile, InStr(csvFile, " TradeData") - 1)

                ' Read data from file
                dataArray = getDataFromFile(GetShortPath(folderPath & csvFile), ",", ErrStr, numRows, numCols, ExcelDateFormat)
                If ErrStr <> "" Then
                    MsgBox ErrStr, vbExclamation
                    Exit Sub
                End If

                ' Long/short splits for every strategy
                ProcessLSTradeData dataArray, tradeLongDataDict, tradeShortDataDict, FileNameOnly, numRows, colLongIndex, colShortIndex, colLongHeaders, colShortHeaders
                ProcessLatestPositions dataArray, latestPositionsDict, FileNameOnly, numRows

                ' ATR/PNL additionally for Buy & Hold strategies
                If folderLocations.Cells(i, 6).value = GetNamedRangeValue("BuyandHoldStatus") Then
                    headerName = ExtractContractName(FileNameOnly)
                    ProcessTradeData dataArray, tradeDataDict, headerName, numRows, colATRIndex, colATRHeaders
                    ATR_Flag = 1
                End If

                i = i + 1
                csvFile = Dir
                Application.StatusBar = "Importing trade data into arrays: " & i - 1 & " of " & lastRow & "..."
            Loop
        Else
            i = i + 1  ' no TradeData file in this folder — keep moving
        End If
    Wend
    
    ' Check if any Buy and Hold files were found
    If GetNamedRangeValue("BuyandHoldStatus") <> "" And ATR_Flag = 0 Then
        errorMessages = errorMessages & "Warning: No Buy and Hold files were found in any folder. ATR calculations will not be performed. Please check that the status in the 'Strategies' tab is updated to buy and hold for the appropriate strategies." & vbCrLf
    End If
    
 
    ' Determine min and max dates
    If uniqueDatesDict.count = 0 Then
        MsgBox "No dates found. Please check the input files.", vbExclamation
        Exit Sub
    End If
    
    Dim dateKey As Variant
    Dim isFirstDate As Boolean
    
    ' Initialize variables
    isFirstDate = True
    
    ' Loop through uniqueDatesDict keys to find min and max dates
    For Each dateKey In uniqueDatesDict.keys
        If isFirstDate Then
            minDate = dateKey
            MaxDate = dateKey
            isFirstDate = False
        Else
            If dateKey < minDate Then minDate = dateKey
            If dateKey > MaxDate Then MaxDate = dateKey
        End If
    Next dateKey
    
    ' Generate all dates from minDate to maxDate
    allDatesList = GenerateDateRange(minDate, MaxDate)
    
    ' Create data arrays
    Dim DailyM2MEquity() As Double, ClosedTradePNL() As Double, InMarketLong() As Double, InMarketShort() As Double, ATR() As Double
    ReDim DailyM2MEquity(1 To UBound(allDatesList), 1 To colIndex)
    ReDim ClosedTradePNL(1 To UBound(allDatesList), 1 To colIndex)
    ReDim InMarketLong(1 To UBound(allDatesList), 1 To colIndex)
    ReDim InMarketShort(1 To UBound(allDatesList), 1 To colIndex)
    
    Dim contractName As Variant
    
    
    If ATR_Flag = 1 Then
        Dim tradeDate As Variant
        Dim tradeDates As Object
        
        
        
        Set tradeDates = CreateObject("Scripting.Dictionary")
        
        ' Populate tradeDates with unique dates from tradeDataDict
        For Each contractName In tradeDataDict.keys
            For Each tradeDate In tradeDataDict(contractName).keys
                If Not tradeDates.Exists(tradeDate) Then
                    tradeDates.Add tradeDate, tradeDate
                End If
            Next tradeDate
        Next contractName
        
        
           
        ' Convert tradeDates to a sorted array
        Dim tradeDateArray() As Date
        tradeDateArray = GetSortedDatesArray(tradeDates)
        
        ' Resize ATR array based on unique trade dates
        ' When resizing ATR array:
        ReDim ATR(1 To UBound(tradeDateArray), 1 To colATRIndex)
        ReDim PNL(1 To UBound(tradeDateArray), 1 To colATRIndex)
    End If
    ' Populate data arrays
    
    
    ' Loop through dates and populate data for each strategy
    For i = 1 To UBound(allDatesList)
        For j = 1 To UBound(colHeaders)
            If combinedData(colHeaders(j)).Exists(allDatesList(i)) Then
                DailyM2MEquity(i, j) = combinedData(colHeaders(j))(allDatesList(i))(0) ' DailyM2MEquity
                ClosedTradePNL(i, j) = combinedData(colHeaders(j))(allDatesList(i))(3) ' ClosedTradePNL
                InMarketLong(i, j) = combinedData(colHeaders(j))(allDatesList(i))(1) ' InMarketLong
                InMarketShort(i, j) = combinedData(colHeaders(j))(allDatesList(i))(2) ' InMarketShort
                
            Else
                DailyM2MEquity(i, j) = 0
                ClosedTradePNL(i, j) = 0
                InMarketLong(i, j) = 0
                InMarketShort(i, j) = 0
            End If
        Next j
        
        
        Application.StatusBar = "Storing data in arrays: " & Format(i / UBound(allDatesList), "0%") & " completed"
    Next i
      
    If ATR_Flag = 1 Then
        ' Populate ATR array only for dates in tradeDates
        For i = 1 To UBound(tradeDateArray)
            For j = 1 To UBound(colATRHeaders)
                If tradeDataDict.Exists(colATRHeaders(j)) And tradeDataDict(colATRHeaders(j)).Exists(tradeDateArray(i)) Then
                    ' Extract both ATR and PNL from the stored array
                    ATR(i, j) = tradeDataDict(colATRHeaders(j))(tradeDateArray(i))(0)
                    PNL(i, j) = tradeDataDict(colATRHeaders(j))(tradeDateArray(i))(1)
                Else
                    ATR(i, j) = 0
                    PNL(i, j) = 0
                End If
            Next j
            
            Application.StatusBar = "Storing True Range data in arrays: " & Format((i) / (UBound(tradeDateArray)), "0%") & " completed"
        Next i
    End If
      
      
      
    ' Write results to sheets
    Application.StatusBar = "Writing data to DailyM2MEquity"
    WriteDataToSheet "DailyM2MEquity", allDatesList, colHeaders, DailyM2MEquity, ExcelDateFormat
    
    Application.StatusBar = "Writing data to ClosedTradePNL"
    WriteDataToSheet "ClosedTradePNL", allDatesList, colHeaders, ClosedTradePNL, ExcelDateFormat
    
    Application.StatusBar = "Writing data to InMarketLong"
    WriteDataToSheet "InMarketLong", allDatesList, colHeaders, InMarketLong, ExcelDateFormat
    
    Application.StatusBar = "Writing data to InMarketShort"
    WriteDataToSheet "InMarketShort", allDatesList, colHeaders, InMarketShort, ExcelDateFormat
    
    ' Write latest positions to sheet
    Application.StatusBar = "Writing latest positions to LatestPosition sheet..."
    Call WriteLatestPositionsToSheet(latestPositionsDict)
    
    ' After processing all trade data, write it to the sheets
    Application.StatusBar = "Writing Long trade data to Long_Trades sheet..."
    Call WriteTradeDataToSheet("Long_Trades", tradeLongDataDict, colLongHeaders)
    
    Application.StatusBar = "Writing Short trade data to Short_Trades sheet..."
    Call WriteTradeDataToSheet("Short_Trades", tradeShortDataDict, colShortHeaders)
    
    If ATR_Flag = 1 Then
        Application.StatusBar = "Writing True Range data to TrueRanges"
        WriteDataToSheet "TrueRanges", tradeDateArray, colATRHeaders, ATR, ExcelDateFormat
        
        Application.StatusBar = "Writing PNL data to TradePNL"
        WriteDataToSheet "TradePNL", tradeDateArray, colATRHeaders, PNL, ExcelDateFormat
    End If
    
    Application.StatusBar = "Writing data to WalkForward Details"
    Call ImportWalkforwardDetailsSimplified(ExcelDateFormat)
    
    
    Application.StatusBar = "Averge True Range calculations..."
    If ATR_Flag = 1 Then
        Call CalculateAverageATR
        Call CalculateAnnualATR
        Call CalculateYearlyContractMultiples
    End If
        
        ' Hide the specific sheets created
    ThisWorkbook.Sheets("DailyM2MEquity").Visible = xlSheetHidden
    ThisWorkbook.Sheets("ClosedTradePNL").Visible = xlSheetHidden
    ThisWorkbook.Sheets("InMarketShort").Visible = xlSheetHidden
    ThisWorkbook.Sheets("InMarketLong").Visible = xlSheetHidden
    ThisWorkbook.Sheets("Long_Trades").Visible = xlSheetHidden
    ThisWorkbook.Sheets("Short_Trades").Visible = xlSheetHidden
    ThisWorkbook.Sheets("LatestPositionData").Visible = xlSheetHidden
    ThisWorkbook.Sheets("Walkforward Details").Visible = xlSheetHidden
    
    If ATR_Flag = 1 Then
        ThisWorkbook.Sheets("AverageTrueRange").Visible = xlSheetHidden
        ThisWorkbook.Sheets("TrueRanges").Visible = xlSheetHidden
        ThisWorkbook.Sheets("TradePNL").Visible = xlSheetHidden  ' Add this line
        ThisWorkbook.Sheets("ContractMultiples").Visible = xlSheetHidden
    End If
    
    ThisWorkbook.Sheets("Control").Activate
    
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    
     ' Show completion message with any warnings
    If errorMessages <> "" Then
        MsgBox "Data processing completed with the following warnings:" & vbCrLf & vbCrLf & errorMessages, vbExclamation
    Else
        MsgBox "Data processing completed successfully.", vbInformation
    End If
    
    
End Sub

Function GenerateDateRange(minDate As Date, MaxDate As Date) As Date()
    Dim dateList() As Date, i As Long, totalDays As Long
    totalDays = DateDiff("d", minDate, MaxDate) + 1
    ReDim dateList(1 To totalDays)
    For i = 1 To totalDays
        dateList(i) = DateAdd("d", i - 1, minDate)
    Next i
    
    GenerateDateRange = dateList
End Function


Sub ProcessEquityData(dataArray As Variant, uniqueDatesDict As Object, combinedData As Object, FileNameOnly As String, numRows As Long, colIndex As Long, ByRef colHeaders() As String)
    ' Processes data from the equity file
    Dim i As Long, dateStr As Variant
    If Not combinedData.Exists(FileNameOnly) Then
        Set combinedData(FileNameOnly) = CreateObject("Scripting.Dictionary")
        colIndex = colIndex + 1
        ReDim Preserve colHeaders(1 To colIndex)
        colHeaders(colIndex) = FileNameOnly
    End If
    For i = 2 To numRows
        dateStr = dataArray(i, 1)
        If IsDate(dateStr) Then
            dateStr = CDate(dateStr)
            If Not uniqueDatesDict.Exists(dateStr) Then uniqueDatesDict.Add dateStr, dateStr
            If Not combinedData(FileNameOnly).Exists(dateStr) Then
                combinedData(FileNameOnly).Add dateStr, Array(dataArray(i, 2), dataArray(i, 3), dataArray(i, 4), dataArray(i, 6))
            End If
        End If
    Next i
End Sub


Sub ProcessTradeData(dataArray As Variant, tradeDataDict As Object, FileNameOnly As String, numRows As Long, colIndex As Long, ByRef colHeaders() As String)
    ' Processes trade data specifically for Exit trades and calculates MFE - MAE
    Dim i As Long, dateStr As Variant, tradeType As String
    Dim MFE As Double, MAE As Double, ATR As Double, PNL As Double

    ' Initialize the trade data dictionary for this contract if not already present
    If Not tradeDataDict.Exists(FileNameOnly) Then
        Set tradeDataDict(FileNameOnly) = CreateObject("Scripting.Dictionary")
        colIndex = colIndex + 1
        ReDim Preserve colHeaders(1 To colIndex)
        colHeaders(colIndex) = FileNameOnly
    End If

    ' Loop through the data array to extract Exit trades
    For i = 2 To numRows
        dateStr = dataArray(i, 1)
        tradeType = dataArray(i, 4)

        If IsDate(dateStr) And tradeType = "Exit" Then
            dateStr = CDate(dateStr)
            MFE = dataArray(i, 8)
            MAE = dataArray(i, 7)
            ATR = MFE - MAE
            PNL = dataArray(i, 6)
            
            ' Add the calculated values to the trade data dictionary
            ' Store both ATR and PNL as an array
            If Not tradeDataDict(FileNameOnly).Exists(dateStr) Then
                tradeDataDict(FileNameOnly).Add dateStr, Array(ATR, PNL)
            End If
        End If
    Next i
End Sub


Sub ProcessLSTradeData(dataArray As Variant, tradeLongDataDict As Object, tradeShortDataDict As Object, _
                     FileNameOnly As String, numRows As Long, ByRef colLongIndex As Long, ByRef colShortIndex As Long, _
                     ByRef colLongHeaders() As String, ByRef colShortHeaders() As String)
    Dim i As Long, dateStr As Variant, tradeType As String, TradePosition As String
    Dim pnlValue As Variant
    Dim currentdate As Date
    Dim longDateCount As Object, shortDateCount As Object ' Separate date counters
    
    ' Initialize the dictionaries for this contract if not already present
    If Not tradeLongDataDict.Exists(FileNameOnly) Then
        Set tradeLongDataDict(FileNameOnly) = CreateObject("Scripting.Dictionary")
        colLongIndex = colLongIndex + 1
        ReDim Preserve colLongHeaders(1 To colLongIndex)
        colLongHeaders(colLongIndex) = FileNameOnly
    End If
    
    If Not tradeShortDataDict.Exists(FileNameOnly) Then
        Set tradeShortDataDict(FileNameOnly) = CreateObject("Scripting.Dictionary")
        colShortIndex = colShortIndex + 1
        ReDim Preserve colShortHeaders(1 To colShortIndex)
        colShortHeaders(colShortIndex) = FileNameOnly
    End If
    
    ' Dictionaries to track trades per date for long and short separately
    Set longDateCount = CreateObject("Scripting.Dictionary")
    Set shortDateCount = CreateObject("Scripting.Dictionary")
    
    ' Loop through the data array to extract Exit trades
    For i = 2 To numRows
        dateStr = dataArray(i, 1)
        tradeType = dataArray(i, 4)
        TradePosition = dataArray(i, 5)
        
        If IsDate(dateStr) And tradeType = "Exit" Then
            currentdate = CDate(dateStr)
            pnlValue = dataArray(i, 6)
            
            ' Process only if PNL value is valid
            If IsNumeric(pnlValue) Then
                ' Determine if it's a long or short trade
                If left(TradePosition, 1) = "L" Then
                    ' Long trade processing
                    If Not longDateCount.Exists(currentdate) Then
                        longDateCount.Add currentdate, 1
                    Else
                        longDateCount(currentdate) = longDateCount(currentdate) + 1
                    End If
                    
                    ' Create a unique key using date and trade number for that date
                    Dim longTradeKey As String
                    longTradeKey = Format(currentdate, "yyyy-mm-dd") & "_" & longDateCount(currentdate)
                    
                    ' Add to long trades dictionary
                    tradeLongDataDict(FileNameOnly).Add longTradeKey, pnlValue
                    
                ElseIf left(TradePosition, 1) = "S" Then
                    ' Short trade processing
                    If Not shortDateCount.Exists(currentdate) Then
                        shortDateCount.Add currentdate, 1
                    Else
                        shortDateCount(currentdate) = shortDateCount(currentdate) + 1
                    End If
                    
                    ' Create a unique key using date and trade number for that date
                    Dim shortTradeKey As String
                    shortTradeKey = Format(currentdate, "yyyy-mm-dd") & "_" & shortDateCount(currentdate)
                    
                    ' Add to short trades dictionary
                    tradeShortDataDict(FileNameOnly).Add shortTradeKey, pnlValue
                End If
            End If
        End If
    Next i
End Sub


Sub ProcessLatestPositions(dataArray As Variant, latestPositionsDict As Object, FileNameOnly As String, numRows As Long)
    ' Processes the last row of trade data to determine current position
    Dim lastRowIndex As Long
    Dim tradeType As String
    Dim position As String
    Dim positionValue As Long
    
    ' Start from the last row and work backwards to find the most recent trade
    For lastRowIndex = numRows To 2 Step -1
        tradeType = Trim(UCase(dataArray(lastRowIndex, 4))) ' Column 4 is "Type"
        
        ' Check if this row has a valid trade type
        If tradeType = "ENTRY" Or tradeType = "EXIT" Then
            If tradeType = "EXIT" Then
                ' If last trade was an exit, position is flat (0)
                positionValue = 0
            ElseIf tradeType = "ENTRY" Then
                ' If last trade was an entry, check the position
                position = Trim(UCase(dataArray(lastRowIndex, 5))) ' Column 5 is "Position"
                If InStr(position, "LONG") > 0 Or left(position, 1) = "L" Then
                    positionValue = 1
                ElseIf InStr(position, "SHORT") > 0 Or left(position, 1) = "S" Then
                    positionValue = -1
                Else
                    positionValue = 0 ' Default to flat if position unclear
                End If
            End If
            
            ' Store the position for this strategy
            latestPositionsDict(FileNameOnly) = positionValue
            Exit For ' Exit once we found the most recent trade
        End If
    Next lastRowIndex
    
    ' If no valid trades found, default to flat position
    If Not latestPositionsDict.Exists(FileNameOnly) Then
        latestPositionsDict(FileNameOnly) = 0
    End If
End Sub


Sub WriteLatestPositionsToSheet(latestPositionsDict As Object)
    Dim ws As Worksheet
    Dim strategyName As Variant
    Dim rowIndex As Long
    
    ' Create or clear the LatestPosition sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LatestPositionData")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = "LatestPositionData"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    ' Write headers
    ws.Cells(1, 1).value = "Strategy Name"
    ws.Cells(1, 2).value = "Position"
    
    ' Write data
    rowIndex = 2
    For Each strategyName In latestPositionsDict.keys
        ws.Cells(rowIndex, 1).value = strategyName
        ws.Cells(rowIndex, 2).value = latestPositionsDict(strategyName)
        rowIndex = rowIndex + 1
    Next strategyName
    
    ' Format the sheet
    ws.Columns("A:B").AutoFit
    ws.rows(1).Font.Bold = True
    
End Sub

Function GetSortedDatesArray(uniqueDatesDict As Object) As Date()
    ' Converts dictionary keys to a sorted array
    Dim dateList() As Date, i As Long
    ReDim dateList(1 To uniqueDatesDict.count)
    i = 1
    For Each dateStr In uniqueDatesDict.keys
        dateList(i) = CDate(dateStr)
        i = i + 1
    Next dateStr
    QuickSortDates dateList, LBound(dateList), UBound(dateList)
    GetSortedDatesArray = dateList
End Function

Sub WriteDataToSheet(sheetName As String, allDatesList() As Date, colHeaders() As String, dataArray As Variant, ExcelDateFormat As String)
    Dim ws As Worksheet, i As Long, j As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = sheetName
    End If
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells(1, 1).value = "Date"

    ' Write headers
    For j = 1 To UBound(colHeaders)
        ws.Cells(1, j + 1).value = colHeaders(j)
    Next j

    ' Debugging date list
    
    'Debug.Print "Writing Dates to Sheet"
    For i = 1 To UBound(allDatesList)
        'Debug.Print "Date: "; allDatesList(i)
        ws.Cells(i + 1, 1).value = allDatesList(i) ' Write date
    Next i

    ' Write data
   ' Debug.Print "Writing Data Array to Sheet"
    ws.Cells(2, 1).Resize(UBound(allDatesList), 1).NumberFormat = IIf(ExcelDateFormat <> "US", "dd/mm/yyyy", "mm/dd/yyyy")
    ws.Cells(2, 2).Resize(UBound(allDatesList), UBound(colHeaders)).value = dataArray
End Sub

Sub WriteTradeDataToSheet(sheetName As String, tradeDataDict As Object, colHeaders() As String)
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim contractName As String
    Dim tradeDict As Object
    Dim tradeKeys As Variant
    Dim sortedKeys() As String
    
    ' Create or clear the sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = sheetName
    End If
    On Error GoTo 0
    ws.Cells.Clear
    
    ' Write column headers (strategy names)
    For j = 1 To UBound(colHeaders)
        ws.Cells(1, j).value = colHeaders(j)
    Next j
    
    ' Write trade data for each strategy
    For j = 1 To UBound(colHeaders)
        contractName = colHeaders(j)
        
        If tradeDataDict.Exists(contractName) Then
            Set tradeDict = tradeDataDict(contractName)
            
            If tradeDict.count > 0 Then
                ' Get all keys and sort them
                tradeKeys = tradeDict.keys
                
                ' Convert to array and sort
                ReDim sortedKeys(0 To tradeDict.count - 1)
                For i = 0 To UBound(sortedKeys)
                    sortedKeys(i) = tradeKeys(i)
                Next i
                
                ' Sort keys chronologically
                Call SortTradeKeys(sortedKeys)
                
                ' Write sorted trade data
                For i = 0 To UBound(sortedKeys)
                    ws.Cells(i + 2, j).value = tradeDict(sortedKeys(i))
                Next i
            End If
        End If
    Next j
    
    ' Format the worksheet
    'ws.Columns.AutoFit
    'ws.Rows(1).Font.Bold = True
    
    ' Make the sheet visible
    ws.Visible = xlSheetVisible
End Sub

' Helper function to sort trade keys by date and trade number
Sub SortTradeKeys(ByRef keys() As String)
    Dim i As Long, j As Long, temp As String
    
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(i) > keys(j) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
End Sub


Sub QuickSortDates(arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long
    Dim mid As Date, tmp As Date
    low = first
    high = last
    mid = arr((first + last) \ 2) ' Ensure mid is a Date

    Do While low <= high
        Do While arr(low) < mid
            low = low + 1
        Loop
        Do While arr(high) > mid
            high = high - 1
        Loop
        If low <= high Then
            tmp = arr(low)
            arr(low) = arr(high)
            arr(high) = tmp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortDates arr, first, high
    If low < last Then QuickSortDates arr, low, last
End Sub



Private Function getDataFromFile(parFileName As String, parDelimiter As String, ByRef ErrStr As String, ByRef numRows As Long, ByRef numCols As Long, ExcelDateFormat As String, Optional parExcludeCharacter As String = "") As Variant
    Dim ErrorFiles As Collection
    
    ' Initializes tracking of files that cannot be opened
    If ErrorFiles Is Nothing Then Set ErrorFiles = New Collection
    
    Dim locLinesList() As Variant
    Dim tempfile As String
    Dim locData As Variant
    Dim i As Long
    Dim j As Long
    Dim locNumRows As Long
    Dim locNumCols As Long
    Dim FSO As Variant
    Dim ts As Variant
    Const REDIM_STEP = 10000
    
    Dim decimalSep As String, thousandSep As String
    
    ' Set decimal and thousand separators based on ExcelDateFormat
    Select Case ExcelDateFormat
        Case "EU"
            decimalSep = ","
            thousandSep = "."
        Case "UK", "US", "AU"
            decimalSep = "."
            thousandSep = ","
        Case Else
            decimalSep = Application.International(xlDecimalSeparator)
            thousandSep = Application.International(xlThousandsSeparator)
    End Select
    
    ErrStr = ""
    numRows = 0
    numCols = 0
      
    If Right(parFileName, 4) <> ".csv" And Right(parFileName, 4) <> ".CSV" Then
        parFileName = parFileName & ".csv"
    End If
        
    Set FSO = CreateObject("Scripting.FileSystemObject")

    On Error GoTo error_open_file
    
    Const ForReading = 1
    Set ts = FSO.OpenTextFile(parFileName, ForReading)
    
    On Error GoTo unhandled_error

    ' Counts the number of lines and the largest number of columns
    ReDim locLinesList(1 To 1) As Variant
    i = 0
    Do While Not ts.AtEndOfStream
        If i Mod REDIM_STEP = 0 Then
            ReDim Preserve locLinesList(1 To UBound(locLinesList, 1) + REDIM_STEP) As Variant
        End If
        locLinesList(i + 1) = Split(ts.ReadLine, parDelimiter)
        j = UBound(locLinesList(i + 1), 1) ' number of columns
        If locNumCols < j Then locNumCols = j
        i = i + 1
    Loop

    ts.Close

    locNumRows = i

    If locNumRows = 0 Then Exit Function ' Empty file

    ReDim locData(1 To locNumRows, 1 To locNumCols + 1) As Variant

    ' Copies the file into an array
    If parExcludeCharacter <> "" Then
        For i = 1 To locNumRows
            For j = 0 To UBound(locLinesList(i), 1)
                locData(i, j + 1) = locLinesList(i)(j)
            Next j
        Next i
    Else
        Dim dateArray() As String
        
        ' Date and decimal separator correction based on format
        For i = 1 To locNumRows
            For j = 0 To UBound(locLinesList(i), 1)
                Dim tempValue As String
                tempValue = locLinesList(i)(j)
                
                ' Fix thousand and decimal separators
                'tempValue = Replace(tempValue, thousandSep, "") ' Remove thousand separator
                If "EU" = ExcelDateFormat Then
                    tempValue = Replace(tempValue, ".", decimalSep) ' Convert decimal separator to dot for uniformity
                End If
                
                ' Convert numbers correctly
                If IsNumeric(tempValue) Then
                    locData(i, j + 1) = CDbl(tempValue) ' Convert explicitly to double
                Else
                    locData(i, j + 1) = tempValue
                End If
                
                ' Handle date conversion
                If (i > 1 And j = 0 And (ExcelDateFormat <> "US")) Then
                    dateArray = Split(tempValue, "/")
                    If UBound(dateArray) = 2 Then
                        locData(i, j + 1) = CDate(dateArray(1) & "/" & dateArray(0) & "/" & dateArray(2))
                    Else
                        ErrStr = "Invalid date format detected in row " & i
                        Exit Function
                    End If
                End If
                
                ' Add integrity check for date issues
                If i > 2 Then
                    If CDate(locData(i, 1)) < CDate(locData(i - 1, 1)) Then
                        ErrStr = "Error in data - Check date format in Inputs tab. Date " & locLinesList(i)(0) & " on row " & i & " is less than previous date " & locLinesList(i - 1)(0)
                        Erase locData
                        Exit Function
                    End If
                End If
            Next j
        Next i
    End If

    getDataFromFile = locData

    numRows = locNumRows
    numCols = locNumCols
    
    Exit Function

error_open_file:
    ' Track the error and continue processing
    ErrorFiles.Add parFileName
    ErrStr = "Could not open " & parFileName
    Exit Function

unhandled_error:
    ErrStr = "Unknown error"
    Exit Function
End Function




Public Sub DisplayErrorFiles()
    If ErrorFiles.count > 0 Then
        Dim msg As String
        msg = "The following files could not be opened:" & vbCrLf
        Dim i As Integer
        For i = 1 To ErrorFiles.count
            msg = msg & "- " & ErrorFiles(i) & vbCrLf
        Next i
        MsgBox msg, vbExclamation, "File Open Errors"
    Else
        MsgBox "All files were processed successfully!", vbInformation
    End If
End Sub

Sub ImportWalkforwardDetailsSimplified(ExcelDateFormat As String)
    Dim ws As Worksheet
    Dim csvFile As String
    Dim folderPath As String
    Dim dataArray As Variant
    Dim ErrStr As String
    Dim numRows As Long, numCols As Long
    Dim lookupName As String
    Dim i As Long, j As Long
    Dim colELCodeFile As Long
    Dim dateArray() As String
    Dim folderLocations As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim processedFolders As Collection
    Dim processedFolderPath As String
    Dim startRow As Long
    Dim headersCopied As Boolean

    ' Set the worksheet containing the folder paths
    Set folderLocations = ThisWorkbook.Sheets("MW Folder Locations")

    ' Create or clear the "Walkforward Details" sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Walkforward Details")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = "Walkforward Details"
    End If
    ws.Cells.ClearContents

    ' Set the headers for Strategy Name and Folder Location
    ws.Cells(1, 1).value = "Strategy Name"
    ws.Cells(1, 2).value = "Folder Location"

    ' Initialize the collection to track processed folder paths
    Set processedFolders = New Collection

    ' Loop through each folder path in column 4 of "MW Folder Locations"
    lastRow = folderLocations.Cells(folderLocations.rows.count, 4).End(xlUp).row


    For row = 2 To lastRow ' Assuming the first row is headers
        folderPath = folderLocations.Cells(row, 4).value & "\Walkforward Files\"
        csvFile = folderPath & "Walkforward In-Out Periods Analysis Details.csv"

        ' Check if the file exists
        If Dir(GetShortPath(csvFile)) <> "" Then
            ' Check if this folder has already been processed
            On Error Resume Next
            processedFolderPath = Trim(folderPath) ' Remove trailing backslash for consistent checking
            processedFolders.Add processedFolderPath, processedFolderPath
            If Err.Number = 0 Then
                ' Use the optimized GetDataFromCSV function to read the CSV file into an array
                dataArray = GetDataFromCSV(GetShortPath(csvFile), ",", ErrStr, numRows, numCols)

                ' Check for errors from the CSV reading
                If ErrStr <> "" Then
                    MsgBox ErrStr, vbExclamation
                    GoTo NextFolder
                End If

                ' Determine where to start pasting data
                startRow = ws.Cells(ws.rows.count, 3).End(xlUp).row + 1 ' Next available row in column 3

                ' Copy headers from the first file opened to the third row
                If Not headersCopied Then
                    ws.Cells(1, 3).Resize(1, numCols).value = Application.index(dataArray, 1) ' Copy headers
                    headersCopied = True
                End If

                ' Paste data, starting from the second row of the data array
                For i = 2 To numRows
                    ws.Cells(startRow, 3).Resize(1, numCols).value = Application.index(dataArray, i)
                    startRow = startRow + 1 ' Move to the next row
                Next i

                ' Fill the folder location in column 2 for all newly added rows
                ws.Cells(startRow - numRows + 1, 2).Resize(numRows - 1, 1).value = processedFolderPath ' Full folder location for all new rows

                ' Find the column with "EL Code File" and adjust the first column
                colELCodeFile = Application.Match("EL Code File", ws.rows(1), 0)
                If Not IsError(colELCodeFile) Then
                    For i = startRow - numRows + 1 To startRow - 1
                        lookupName = left(ws.Cells(i, colELCodeFile).value, InStr(ws.Cells(i, colELCodeFile).value, " ELCode.txt") - 1)
                        ws.Cells(i, 1).value = lookupName ' Strategy Name
                    Next i
                End If

                If ExcelDateFormat <> "US" Then ' Adjust any date columns (any column header that includes the word "Date")
                    For j = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).column
                        If InStr(1, ws.Cells(1, j).value, "Date", vbTextCompare) > 0 Then
                            ' It's a date column, so convert each cell in this column to UK date format
                            For i = startRow - numRows + 1 To startRow - 1
                                If IsDate(ws.Cells(i, j).value) Then
                                    dateArray = Split(ws.Cells(i, j).value, "/")
                                    ws.Cells(i, j).value = CDate(dateArray(0) & "/" & dateArray(1) & "/" & dateArray(2))
                                End If
                            Next i
                        End If
                    Next j
                End If
            End If
            On Error GoTo 0 ' Reset error handling
        Else
        '    MsgBox "CSV file not found in: " & FolderPath, vbExclamation
        End If
NextFolder:
    Next row

    
   ' MsgBox "Data imported and adjusted in 'Walkforward Details' sheet!", vbInformation
End Sub

Sub CalculateAverageATR()
    Dim wsTrueRanges As Worksheet, wsAverageTrueRange As Worksheet
    Dim DataRange As Range, rowCount As Long, colCount As Long
    Dim contracts() As String, dates() As Date, ATRs As Variant
    Dim i As Long, j As Long, K As Long
    Dim periodHeaders As Variant
    Dim dateCutoffs As Variant
    Dim results() As Double
    Dim lastRow As Long, lastCol As Long
    Dim currentdate As Date
    
    ' Define period headers and durations
    periodHeaders = Array("1 Month", "3 Months", "6 Months", "12 Months", "24 Months", "60 Months", "All Time")
    dateCutoffs = Array(30, 90, 180, 365, 730, 1825, 999999999999#) ' Days for each period
    
    ' Load data from TrueRanges
    On Error Resume Next
    Set wsTrueRanges = ThisWorkbook.Sheets("TrueRanges")
    If wsTrueRanges Is Nothing Then
        MsgBox "Sheet 'TrueRanges' not found.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    lastRow = EndRowByCutoffSimple(wsTrueRanges, 1)
    lastCol = wsTrueRanges.Cells(1, wsTrueRanges.Columns.count).End(xlToLeft).column
    Set DataRange = wsTrueRanges.Range(wsTrueRanges.Cells(1, 1), wsTrueRanges.Cells(lastRow, lastCol))
    
    ' Extract the current date from the last row of the Date column
    currentdate = wsTrueRanges.Cells(lastRow, 1).value
    
    ' Extract contracts and dates
    colCount = lastCol - 1 ' Exclude date column
    ReDim contracts(1 To colCount)
    For j = 2 To lastCol
        contracts(j - 1) = wsTrueRanges.Cells(1, j).value
    Next j
    
    rowCount = lastRow - 1 ' Exclude header row
    ReDim dates(1 To rowCount)
    For i = 2 To lastRow
        dates(i - 1) = wsTrueRanges.Cells(i, 1).value
    Next i
    
    ' Extract ATR values
    ATRs = DataRange.Offset(1, 1).Resize(rowCount, colCount).value ' Exclude headers
    
    ' Initialize results array
    ReDim results(1 To colCount, 1 To UBound(periodHeaders) + 1)
    
    ' Calculate averages for each contract
    For j = 1 To colCount
        For K = LBound(periodHeaders) To UBound(periodHeaders)
            Dim total As Double, count As Long
            total = 0
            count = 0
            For i = 1 To rowCount
                If DateDiff("d", dates(i), currentdate) <= dateCutoffs(K) Or K = UBound(periodHeaders) Then
                    total = total + ATRs(i, j)
                    count = count + 1
                End If
            Next i
            If count > 0 Then
                results(j, K + 1) = total / count
            Else
                results(j, K + 1) = 0
            End If
        Next K
    Next j
    
    ' Create or clear AverageTrueRange sheet
    On Error Resume Next
    Set wsAverageTrueRange = ThisWorkbook.Sheets("AverageTrueRange")
    If wsAverageTrueRange Is Nothing Then
        Set wsAverageTrueRange = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsAverageTrueRange.name = "AverageTrueRange"
    Else
        wsAverageTrueRange.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Write headers to AverageTrueRange
    wsAverageTrueRange.Cells(1, 1).value = "Contract"
    For K = LBound(periodHeaders) To UBound(periodHeaders)
        wsAverageTrueRange.Cells(1, K + 2).value = periodHeaders(K)
    Next K
    
    ' Write results to AverageTrueRange
    For j = 1 To colCount
        wsAverageTrueRange.Cells(j + 1, 1).value = contracts(j)
        For K = LBound(periodHeaders) To UBound(periodHeaders)
            wsAverageTrueRange.Cells(j + 1, K + 2).value = results(j, K + 1)
        Next K
    Next j
    
    ' Format the table
    With wsAverageTrueRange
        .Columns("A:H").AutoFit
        .Cells(1, 1).EntireRow.Font.Bold = True
        .Columns("A:H").NumberFormat = "$#,##0"
    End With
    
End Sub


Sub CalculateAnnualATR()
    Dim wsTrueRanges As Worksheet, wsAverageTrueRange As Worksheet
    Dim DataRange As Range, rowCount As Long, colCount As Long
    Dim contracts() As String, dates() As Date, ATRs As Variant
    Dim i As Long, j As Long, K As Long
    Dim results() As Double
    Dim lastRow As Long, lastCol As Long
    Dim yearList As Collection, currentYear As Long
    Dim yearIndex As Long, yearATR() As Variant

    ' Load data from TrueRanges
    On Error Resume Next
    Set wsTrueRanges = ThisWorkbook.Sheets("TrueRanges")
    If wsTrueRanges Is Nothing Then
        MsgBox "Sheet 'TrueRanges' not found.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    lastRow = wsTrueRanges.Cells(wsTrueRanges.rows.count, 1).End(xlUp).row
    lastCol = wsTrueRanges.Cells(1, wsTrueRanges.Columns.count).End(xlToLeft).column
    Set DataRange = wsTrueRanges.Range(wsTrueRanges.Cells(1, 1), wsTrueRanges.Cells(lastRow, lastCol))

    ' Extract contracts and dates
    colCount = lastCol - 1 ' Exclude the date column
    ReDim contracts(1 To colCount)
    For j = 2 To lastCol
        contracts(j - 1) = wsTrueRanges.Cells(1, j).value
    Next j

    rowCount = lastRow - 1 ' Exclude the header row
    ReDim dates(1 To rowCount)
    For i = 2 To lastRow
        dates(i - 1) = wsTrueRanges.Cells(i, 1).value
    Next i

    ' Extract ATR values
    ATRs = DataRange.Offset(1, 1).Resize(rowCount, colCount).value ' Exclude headers

    ' Identify unique years
    Set yearList = New Collection
    On Error Resume Next
    For i = 1 To rowCount
        currentYear = Year(dates(i))
        yearList.Add currentYear, CStr(currentYear)
    Next i
    On Error GoTo 0

    ' Prepare results array
    ReDim results(1 To colCount, 1 To yearList.count)
    ReDim yearATR(1 To colCount, 1 To yearList.count)

    ' Calculate annual ATR for each contract
    For j = 1 To colCount
        For yearIndex = 1 To yearList.count
            Dim total As Double, count As Long
            total = 0
            count = 0
            For i = 1 To rowCount
                If Year(dates(i)) = yearList(yearIndex) Then
                    total = total + ATRs(i, j)
                    count = count + 1
                End If
            Next i
            If count > 0 Then
                results(j, yearIndex) = total / count
            Else
                results(j, yearIndex) = 0
            End If
        Next yearIndex
    Next j

    ' Create or append to the AverageTrueRange sheet
    On Error Resume Next
    Set wsAverageTrueRange = ThisWorkbook.Sheets("AverageTrueRange")
    If wsAverageTrueRange Is Nothing Then
        Set wsAverageTrueRange = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsAverageTrueRange.name = "AverageTrueRange"
    End If
    On Error GoTo 0

    ' Find the last column and paste results
    Dim startColumn As Long
    startColumn = wsAverageTrueRange.Cells(1, wsAverageTrueRange.Columns.count).End(xlToLeft).column + 2
    
    wsAverageTrueRange.Cells(1, startColumn).value = "Reference ATR"
    For j = 1 To colCount
        wsAverageTrueRange.Cells(j + 1, startColumn).value = (wsAverageTrueRange.Cells(j + 1, 2).value + wsAverageTrueRange.Cells(j + 1, 3).value + wsAverageTrueRange.Cells(j + 1, 4).value + wsAverageTrueRange.Cells(j + 1, 5).value) / 4
    Next j
    
    startColumn = startColumn + 1
    
    ' Write headers for years
    For yearIndex = 1 To yearList.count
        wsAverageTrueRange.Cells(1, startColumn + yearIndex - 1).value = yearList(yearIndex)
    Next yearIndex

    ' Write results for each contract
    For j = 1 To colCount
        For yearIndex = 1 To yearList.count
            wsAverageTrueRange.Cells(j + 1, startColumn + yearIndex - 1).value = results(j, yearIndex)
        Next yearIndex
    Next j

    ' Format the table
    With wsAverageTrueRange
        .Columns.AutoFit
        .Cells(1, 1).EntireRow.Font.Bold = True
        .Columns.NumberFormat = "$#,##0.00"
    End With
    wsAverageTrueRange.Cells(1, 1).EntireRow.NumberFormat = "####"
    
End Sub


Sub CalculateYearlyContractMultiples()
    Dim wsAverageTrueRange As Worksheet, wsContractMultiples As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim referenceATRCol As Long
    Dim yearStartCol As Long, yearEndCol As Long
    Dim contractCol As Long, i As Long, j As Long
    Dim yearHeaders() As String
    Dim contractName As String
    Dim referenceATR As Double, annualATR As Double
    Dim contractMultiple As Double

    ' Locate the AverageTrueRange sheet
    On Error Resume Next
    Set wsAverageTrueRange = ThisWorkbook.Sheets("AverageTrueRange")
    If wsAverageTrueRange Is Nothing Then
        MsgBox "Sheet 'AverageTrueRange' not found. Run the ATR calculation first.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Get the dimensions of the AverageTrueRange sheet
    lastRow = wsAverageTrueRange.Cells(wsAverageTrueRange.rows.count, 1).End(xlUp).row
    lastCol = wsAverageTrueRange.Cells(1, wsAverageTrueRange.Columns.count).End(xlToLeft).column

    ' Find the Reference ATR column
    For j = 1 To lastCol
        If wsAverageTrueRange.Cells(1, j).value = "Reference ATR" Then
            referenceATRCol = j
            Exit For
        End If
    Next j

    If referenceATRCol = 0 Then
        MsgBox "Reference ATR column not found in 'AverageTrueRange' sheet.", vbCritical
        Exit Sub
    End If

    ' Identify the start and end columns for years
    yearStartCol = referenceATRCol + 1
    yearEndCol = lastCol

    ' Extract year headers
    ReDim yearHeaders(yearStartCol To yearEndCol)
    For j = yearStartCol To yearEndCol
        yearHeaders(j) = wsAverageTrueRange.Cells(1, j).value
    Next j

    ' Create or clear the ContractMultiples sheet
    On Error Resume Next
    Set wsContractMultiples = ThisWorkbook.Sheets("ContractMultiples")
    If wsContractMultiples Is Nothing Then
        Set wsContractMultiples = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsContractMultiples.name = "ContractMultiples"
    Else
        wsContractMultiples.Cells.Clear
    End If
    On Error GoTo 0

    ' Write headers to ContractMultiples sheet
    wsContractMultiples.Cells(1, 1).value = "Contract"
    For j = yearStartCol To yearEndCol
        wsContractMultiples.Cells(1, j - yearStartCol + 2).value = yearHeaders(j)
    Next j

    ' Populate contract names and calculate multiples
    For i = 2 To lastRow
        contractName = wsAverageTrueRange.Cells(i, 1).value
        wsContractMultiples.Cells(i, 1).value = contractName
        referenceATR = wsAverageTrueRange.Cells(i, referenceATRCol).value

        For j = yearStartCol To yearEndCol
            annualATR = wsAverageTrueRange.Cells(i, j).value

            If referenceATR > 0 And annualATR > 0 Then
                contractMultiple = Application.Max(Application.Floor(referenceATR / annualATR, 1), 1)
            Else
                contractMultiple = 1 ' Default value if Reference ATR is zero
            End If

            wsContractMultiples.Cells(i, j - yearStartCol + 2).value = contractMultiple
        Next j
    Next i

    ' Format the ContractMultiples sheet
    With wsContractMultiples
        .Columns.AutoFit
        .Cells(1, 1).EntireRow.Font.Bold = True
        .Columns("B:ZZ").NumberFormat = "0.00"
    End With

    wsAverageTrueRange.Cells(1, 1).EntireRow.NumberFormat = "####"
End Sub


Private Function GetDataFromCSV(parFileName As String, parDelimiter As String, ByRef ErrStr As String, ByRef numRows As Long, ByRef numCols As Long, Optional parExcludeCharacter As String = "") As Variant
    Dim locData As Variant
    Dim locLinesList() As String
    Dim i As Long, j As Long
    Dim locNumRows As Long
    Dim locNumCols As Long
    Dim FSO As Object
    Dim ts As Object
    Dim line As String
    Const REDIM_STEP = 10000

    ErrStr = ""
    numRows = 0
    numCols = 0

    If Right(parFileName, 4) <> ".csv" Then
        parFileName = parFileName & ".csv"
    End If

    ' Create FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    On Error GoTo error_open_file

    ' Open the text file for reading
    Set ts = FSO.OpenTextFile(parFileName, 1)

    ' Read all lines into a list
    i = 0
    ReDim locLinesList(1 To 1)
    Do While Not ts.AtEndOfStream
        If i Mod REDIM_STEP = 0 Then
            ReDim Preserve locLinesList(1 To UBound(locLinesList) + REDIM_STEP)
        End If
        line = ts.ReadLine
        locLinesList(i + 1) = line
        i = i + 1
    Loop

    ts.Close
    locNumRows = i

    If locNumRows = 0 Then Exit Function ' Empty file

    ' Determine the number of columns by splitting the first line
    Dim headerrow() As String
    headerrow = SplitQuoted(locLinesList(1), parDelimiter)
    locNumCols = UBound(headerrow) + 1
    ReDim locData(1 To locNumRows, 1 To locNumCols)

    ' Populate locData array
    For i = 1 To locNumRows
        Dim rowData() As String
        rowData = SplitQuoted(locLinesList(i), parDelimiter) ' Use the splitting function

        ' Ensure locData can accommodate the row
        For j = 0 To locNumCols - 1
            If j < UBound(rowData) + 1 Then
                locData(i, j + 1) = rowData(j) ' 1-based index
            Else
                locData(i, j + 1) = "" ' Fill with empty string for missing fields
            End If

            ' Clean data if needed
            If parExcludeCharacter <> "" Then
                locData(i, j + 1) = Replace(Replace(locData(i, j + 1), parExcludeCharacter, ""), Chr(34), "") ' Remove quotes
            End If
        Next j
    Next i

    numRows = locNumRows
    numCols = locNumCols ' Ensure numCols reflects the header size
    GetDataFromCSV = locData
    Exit Function

error_open_file:
    ErrStr = "Could not open " & parFileName
    Exit Function
End Function


' Function to split a line into an array while respecting quotes
Private Function SplitQuoted(line As String, delimiter As String) As String()
    Dim result() As String
    Dim temp As String
    Dim inQuotes As Boolean
    Dim i As Long
    Dim count As Long

    inQuotes = False
    temp = ""
    count = 0

    For i = 1 To Len(line)
        Dim currentChar As String
        currentChar = mid(line, i, 1)

        If currentChar = Chr(34) Then ' Handle quotes
            inQuotes = Not inQuotes
        ElseIf currentChar = delimiter And Not inQuotes Then
            ' Add the temp string to the result array and reset temp
            If Len(temp) > 0 Or (Len(temp) = 0 And count > 0) Then
                If count = 0 Then
                    ReDim result(0)
                Else
                    ReDim Preserve result(0 To count)
                End If
                result(count) = temp ' Store the current temp
                temp = ""
                count = count + 1
            ElseIf Len(temp) = 0 Then
                ' Handle consecutive delimiters (blank fields)
                If count = 0 Then
                    ReDim result(0)
                Else
                    ReDim Preserve result(0 To count)
                End If
                result(count) = "" ' Store blank field
                count = count + 1
            End If
        Else
            temp = temp & currentChar
        End If
    Next i

    ' Add the last segment if there is any
    If Len(temp) > 0 Then
        If count = 0 Then
            ReDim result(0)
        Else
            ReDim Preserve result(0 To count)
        End If
        result(count) = temp
    End If

    SplitQuoted = result
End Function



Function ExtractContractName(fileName As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim prefix As String

    ' Locate the position of "[@" or "[$"
    startPos = InStr(fileName, "[@")
    If startPos = 0 Then
        startPos = InStr(fileName, "[$")
    End If
    
    ' If neither is found, return an empty string
    If startPos = 0 Then
        ExtractContractName = fileName
        Exit Function
    End If
    
    ' Move start position past the prefix
    startPos = startPos + 2
    
    ' Locate the position of the next "-"
    endPos = InStr(startPos, fileName, "-")
    
    ' Extract the substring
    If endPos > 0 Then
        ExtractContractName = mid(fileName, startPos, endPos - startPos)
    Else
        ExtractContractName = fileName
    End If
End Function




Public Sub InitializeStrategyCache(wsStrategies As Worksheet)
    Dim lastRow As Long
    lastRow = wsStrategies.Cells(wsStrategies.rows.count, 1).End(xlUp).row
    stratData = wsStrategies.Range("A2:E" & lastRow).value
    dataLoaded = True
End Sub

'�� Now this can be called repeatedly without hitting the sheet again
Public Function GetStrategyStatus(strategyName As String) As String
    Dim i As Long
    
    If Not dataLoaded Then
        Err.Raise vbObjectError + 1, , "Cache not initialized�run InitializeStrategyCache first."
    End If
    
    For i = 1 To UBound(stratData, 1)
        If Trim(stratData(i, COL_STRAT_STRATEGY_NAME)) = Trim(strategyName) Then
            GetStrategyStatus = Trim(stratData(i, COL_STRAT_STATUS))
            Exit Function
        End If
    Next i
    
    GetStrategyStatus = "Not Found"
End Function



Public Function LongFilePathname(ByVal fullFileSpec As String) As String
    Const MAX_PATH As Long = 260
    
    ' If path length is approaching MAX_PATH, we need to modify it
    If Len(fullFileSpec) >= (MAX_PATH - 15) Then
        ' Check if it's a network path (starts with \\)
        If left(fullFileSpec, 2) = "\\" Then
            ' Network path - convert to UNC format
            LongFilePathname = "\\?\UNC\" & mid(fullFileSpec, 3)
        Else
            ' Local path
            LongFilePathname = "\\?\" & fullFileSpec
        End If
    Else
        ' No modification needed for shorter paths
        LongFilePathname = fullFileSpec
    End If
End Function




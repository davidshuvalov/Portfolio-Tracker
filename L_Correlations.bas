Attribute VB_Name = "L_Correlations"
Option Explicit


Sub RunCorrelationAnalysis()
    ' Application settings
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Declare all variables
    Dim wsPortfolioDailyM2M As Worksheet
    Dim wsPortfolio As Worksheet
    Dim wsCorrelation As Worksheet
    Dim lastRow As Long, lastColumn As Long
    Dim i As Long, j As Long, K As Long
    Dim nonZeroCount As Long
    Dim correlationValue As Double
    Dim rowOffset As Long, rowOffset2 As Long
    Dim allData As Variant
    Dim drawdownData As Variant
    Dim portfolioData As Variant
    Dim startdate As Date, Current_Date As Date
    Dim dailyProfits1() As Variant, dailyProfits2() As Variant
    Dim correlations10Y() As Double
    Dim correlations1Y() As Double
    Dim combinedCorrelations() As Double
    Dim startRow As Long
    ' Cache threshold values
    Dim High_Threshold As Double
    Dim Low_Threshold As Double
    Dim Short_Period As Double
    Dim Long_Period As Double
    
    ' NEW: User selection variables
    Dim correlationType As String
    Dim sheetName As String
    Dim tabColor As Long
    
    ' Initialize column constants
    Call InitializeColumnConstantsManually
    
    ' Check license
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        GoTo CleanExit
    End If
    
    ' NEW: Ask user for correlation type
    correlationType = InputBox("Choose correlation analysis type:" & vbCrLf & _
                              "1 = Normal Correlations (profit/loss relationships)" & vbCrLf & _
                              "2 = Negative Correlations (excluding both positive days)" & vbCrLf & _
                              "3 = Drawdown Correlations (equity curve drawdowns)", _
                              "Correlation Analysis Type", "1")
    
    ' Check if user pressed Cancel
    If correlationType = "" Then GoTo CleanExit
    
    If correlationType <> "1" And correlationType <> "2" And correlationType <> "3" Then
        MsgBox "Invalid selection. Using Normal Correlations.", vbInformation
        correlationType = "1"
    End If
    
    ' Set sheet properties based on correlation type
    Select Case correlationType
        Case "1"
            sheetName = "Correlations"
            tabColor = RGB(117, 219, 255)
            High_Threshold = GetNamedRangeValue("Correl_High_Threshold")
            Low_Threshold = GetNamedRangeValue("Correl_Low_Threshold")
        Case "2"
            sheetName = "NegativeCorrelations"
            tabColor = RGB(255, 255, 255)
            High_Threshold = GetNamedRangeValue("NegativeCorrel_High_Threshold")
            Low_Threshold = GetNamedRangeValue("NegativeCorrel_Low_Threshold")
        Case "3"
            sheetName = "DrawdownCorrelations"
            tabColor = RGB(255, 150, 150)
            High_Threshold = GetNamedRangeValue("DDCorrel_High_Threshold")
            Low_Threshold = GetNamedRangeValue("DDCorrel_Low_Threshold")
    End Select
    
    ' Validate required sheets exist
    If Not SheetExists("Portfolio") Or Not SheetExists("PortfolioDailyM2M") Then
        MsgBox "Required sheets missing!", vbCritical
        GoTo CleanExit
    End If
    
    ' Set worksheet references
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set wsPortfolioDailyM2M = ThisWorkbook.Sheets("PortfolioDailyM2M")
    
    ' Validate data exists
    If wsPortfolio.Cells(2, COL_PORT_STRATEGY_NAME).value = "" Then
        MsgBox "Error: 'Portfolio' sheet exists but contains no data in row 2.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get common threshold values
    Short_Period = GetNamedRangeValue("Correl_Short_Period")
    Long_Period = GetNamedRangeValue("Correl_Long_Period")
    Current_Date = wsPortfolio.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    
    ' Validate threshold values
    If High_Threshold <= Low_Threshold Then
        MsgBox "Invalid threshold values. High threshold must be greater than low threshold.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get data dimensions
    lastRow = wsPortfolioDailyM2M.Cells(wsPortfolioDailyM2M.rows.count, 1).End(xlUp).row
    lastColumn = wsPortfolioDailyM2M.Cells(1, wsPortfolioDailyM2M.Columns.count).End(xlToLeft).column
    
    ' Read all data into arrays for faster processing
    allData = wsPortfolioDailyM2M.Range(wsPortfolioDailyM2M.Cells(1, 1), _
                                      wsPortfolioDailyM2M.Cells(lastRow, lastColumn)).value
    portfolioData = wsPortfolio.Range(wsPortfolio.Cells(2, 1), _
                                    wsPortfolio.Cells(lastColumn, COL_PORT_SYMBOL)).value
    
    ' For drawdown correlations, convert equity data to drawdown data
    If correlationType = "3" Then
        ReDim drawdownData(1 To UBound(allData, 1), 1 To UBound(allData, 2))
        
        ' Copy headers
        For j = 1 To UBound(allData, 2)
            drawdownData(1, j) = allData(1, j)
        Next j
        
        ' Calculate drawdowns for each strategy
        For j = 2 To UBound(allData, 2)
            Dim peak As Double: peak = 0
            For i = 2 To UBound(allData, 1)
                ' Copy date
                If j = 2 Then drawdownData(i, 1) = allData(i, 1)
                
                ' Find peak and calculate drawdown
                If CDbl(allData(i, j)) > peak Then
                    peak = CDbl(allData(i, j))
                    drawdownData(i, j) = 0 ' At peak, no drawdown
                Else
                    ' Calculate drawdown as percentage from peak
                    If peak > 0 Then
                        drawdownData(i, j) = (CDbl(allData(i, j)) - peak) / peak * 100
                    Else
                        drawdownData(i, j) = 0
                    End If
                End If
            Next i
            Application.StatusBar = "Converting to drawdown data: " & Format(j / UBound(allData, 2), "0%") & " completed"
        Next j
    End If
    
    ' Initialize correlation arrays
    ReDim correlations10Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim correlations1Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim combinedCorrelations(1 To lastColumn - 1, 1 To lastColumn - 1)
    
    ' Calculate Long Period Correlations
    startdate = DateAdd("yyyy", -Int(Long_Period), Current_Date)
    startdate = DateAdd("d", -Int((Long_Period - Int(Long_Period)) * 365.25), startdate)
    
    Application.StatusBar = "Calculating long-term correlations..."
    Select Case correlationType
        Case "1"
            Call CalculateCorrelationMatrix(correlations10Y, allData, startdate, Current_Date, lastRow, lastColumn, "Long")
        Case "2"
            Call CalculateNegativeCorrelationMatrix(correlations10Y, allData, startdate, Current_Date, lastRow, lastColumn, "Long")
        Case "3"
            Call CalculateDrawdownCorrelationMatrix(correlations10Y, drawdownData, startdate, Current_Date, lastRow, lastColumn, "Long")
    End Select
    
    ' Calculate Short Period Correlations
    startdate = DateAdd("yyyy", -Int(Short_Period), Current_Date)
    startdate = DateAdd("d", -Int((Short_Period - Int(Short_Period)) * 365.25), startdate)
    
    Application.StatusBar = "Calculating short-term correlations..."
    Select Case correlationType
        Case "1"
            Call CalculateCorrelationMatrix(correlations1Y, allData, startdate, Current_Date, lastRow, lastColumn, "Short")
        Case "2"
            Call CalculateNegativeCorrelationMatrix(correlations1Y, allData, startdate, Current_Date, lastRow, lastColumn, "Short")
        Case "3"
            Call CalculateDrawdownCorrelationMatrix(correlations1Y, drawdownData, startdate, Current_Date, lastRow, lastColumn, "Short")
    End Select
    
    ' Create new correlation sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(sheetName).Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    Set wsCorrelation = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsCorrelation.name = sheetName
    wsCorrelation.Tab.Color = tabColor
    
    ' Set background color based on type
    If correlationType = "2" Then
        wsCorrelation.Cells.Interior.Color = RGB(255, 255, 255)
    End If
    
    startRow = 4
    
    ' Use appropriate data source for formatting
    Dim dataToUse As Variant
    If correlationType = "3" Then
        dataToUse = drawdownData
    Else
        dataToUse = allData
    End If
    
    ' Format and populate Long Period correlation matrix
    Dim longTitle As String
    Select Case correlationType
        Case "1": longTitle = Long_Period & " Year(s) Correlations"
        Case "2": longTitle = Long_Period & " Year(s) Negative Correlations"
        Case "3": longTitle = Long_Period & " Year(s) Drawdown Correlations"
    End Select
    
    Call FormatCorrelationMatrix(wsCorrelation, correlations10Y, dataToUse, portfolioData, _
    (startRow + 1), longTitle, High_Threshold, Low_Threshold, lastColumn)
    
    ' Format and populate Short Period correlation matrix
    rowOffset = lastColumn + 3 + startRow
    Dim shortTitle As String
    Select Case correlationType
        Case "1": shortTitle = Short_Period & " Year(s) Correlations"
        Case "2": shortTitle = Short_Period & " Year(s) Negative Correlations"
        Case "3": shortTitle = Short_Period & " Year(s) Drawdown Correlations"
    End Select
    
    Call FormatCorrelationMatrix(wsCorrelation, correlations1Y, dataToUse, portfolioData, _
    rowOffset + 1, shortTitle, High_Threshold, Low_Threshold, lastColumn)
    
    ' Calculate and format Combined correlation matrix
    rowOffset2 = (lastColumn) * 2 + 7 + startRow
    For i = 1 To lastColumn - 1
        For j = 1 To lastColumn - 1
            combinedCorrelations(i, j) = Abs(correlations10Y(i, j)) + Abs(correlations1Y(i, j))
        Next j
    Next i
    
    Dim combinedTitle As String
    Select Case correlationType
        Case "1": combinedTitle = "Total Correlations"
        Case "2": combinedTitle = "Total Negative Correlations"
        Case "3": combinedTitle = "Total Drawdown Correlations"
    End Select
    
    Call FormatCorrelationMatrix(wsCorrelation, combinedCorrelations, dataToUse, portfolioData, _
    rowOffset2, combinedTitle, High_Threshold, Low_Threshold, lastColumn)
    
    ' Add appropriate navigation buttons based on correlation type
    Select Case correlationType
        Case "1"
            Call AddCorrNavigationButtons(wsCorrelation, "Correlations")
        Case "2"
            Call AddCorrNavigationButtons(wsCorrelation, "NegativeCorrelations")
        Case "3"
            Call AddDrawdownCorrNavigationButtons(wsCorrelation, "DrawdownCorrelations")
    End Select
    
    ' Final formatting
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
    
    Call OrderVisibleTabsBasedOnList
    wsCorrelation.Activate
    
    ' Show completion message
    Dim analysisTypeName As String
    Select Case correlationType
        Case "1": analysisTypeName = "Normal"
        Case "2": analysisTypeName = "Negative"
        Case "3": analysisTypeName = "Drawdown"
    End Select
    
    MsgBox analysisTypeName & " correlation analysis complete!", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub


Sub CreateCorrelationMatrices()
    ' Application settings
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Declare all variables
    Dim wsPortfolioDailyM2M As Worksheet
    Dim wsPortfolio As Worksheet
    Dim wsCorrelation As Worksheet
    Dim lastRow As Long, lastColumn As Long
    Dim i As Long, j As Long, K As Long
    Dim nonZeroCount As Long
    Dim correlationValue As Double
    Dim rowOffset As Long, rowOffset2 As Long
    Dim allData As Variant
    Dim portfolioData As Variant
    Dim startdate As Date, Current_Date As Date
    Dim dailyProfits1() As Variant, dailyProfits2() As Variant
    Dim correlations10Y() As Double
    Dim correlations1Y() As Double
    Dim combinedCorrelations() As Double
    Dim startRow As Long
    ' Cache threshold values
    Dim High_Threshold As Double
    Dim Low_Threshold As Double
    Dim Short_Period As Double
    Dim Long_Period As Double
    
    ' Initialize column constants
    Call InitializeColumnConstantsManually
    
    ' Check license
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        GoTo CleanExit
    End If
    
    ' Validate required sheets exist
    If Not SheetExists("Portfolio") Or Not SheetExists("PortfolioDailyM2M") Then
        MsgBox "Required sheets missing!", vbCritical
        GoTo CleanExit
    End If
    
    ' Set worksheet references
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set wsPortfolioDailyM2M = ThisWorkbook.Sheets("PortfolioDailyM2M")
    
    ' Validate data exists
    If wsPortfolio.Cells(2, COL_PORT_STRATEGY_NAME).value = "" Then
        MsgBox "Error: 'Portfolio' sheet exists but contains no data in row 2.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get threshold values
    High_Threshold = GetNamedRangeValue("Correl_High_Threshold")
    Low_Threshold = GetNamedRangeValue("Correl_Low_Threshold")
    Short_Period = GetNamedRangeValue("Correl_Short_Period")
    Long_Period = GetNamedRangeValue("Correl_Long_Period")
    Current_Date = wsPortfolio.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    
    ' Validate threshold values
    If High_Threshold <= Low_Threshold Then
        MsgBox "Invalid threshold values. High threshold must be greater than low threshold.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get data dimensions
    lastRow = wsPortfolioDailyM2M.Cells(wsPortfolioDailyM2M.rows.count, 1).End(xlUp).row
    lastColumn = wsPortfolioDailyM2M.Cells(1, wsPortfolioDailyM2M.Columns.count).End(xlToLeft).column
    
    ' Read all data into arrays for faster processing
    allData = wsPortfolioDailyM2M.Range(wsPortfolioDailyM2M.Cells(1, 1), _
                                      wsPortfolioDailyM2M.Cells(lastRow, lastColumn)).value
    portfolioData = wsPortfolio.Range(wsPortfolio.Cells(2, 1), _
                                    wsPortfolio.Cells(lastColumn, COL_PORT_SYMBOL)).value
    
    ' Initialize correlation arrays
    ReDim correlations10Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim correlations1Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim combinedCorrelations(1 To lastColumn - 1, 1 To lastColumn - 1)
    
    ' Calculate Long Period Correlations
    startdate = DateAdd("yyyy", -Int(Long_Period), Current_Date)
    startdate = DateAdd("d", -Int((Long_Period - Int(Long_Period)) * 365.25), startdate)
    
    Application.StatusBar = "Calculating long-term correlations..."
    Call CalculateCorrelationMatrix(correlations10Y, allData, startdate, Current_Date, lastRow, lastColumn, "Long")
    
    ' Calculate Short Period Correlations
    startdate = DateAdd("yyyy", -Int(Short_Period), Current_Date)
    startdate = DateAdd("d", -Int((Short_Period - Int(Short_Period)) * 365.25), startdate)
    
    Application.StatusBar = "My god that was fast, now calculating short-term correlations..."
    Call CalculateCorrelationMatrix(correlations1Y, allData, startdate, Current_Date, lastRow, lastColumn, "Short")
    
    ' Create new correlation sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Correlations").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    Set wsCorrelation = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsCorrelation.name = "Correlations"
    wsCorrelation.Tab.Color = RGB(117, 219, 255)
    
     ' Set white background color for the entire worksheet
    'wsCorrelation.Cells.Interior.Color = RGB(255, 255, 255)
    
    
    startRow = 4
    
    ' Format and populate Long Period correlation matrix
    Call FormatCorrelationMatrix(wsCorrelation, correlations10Y, allData, portfolioData, _
    (startRow + 1), Long_Period & " Year(s) Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Format and populate Short Period correlation matrix
    rowOffset = lastColumn + 3 + startRow
    Call FormatCorrelationMatrix(wsCorrelation, correlations1Y, allData, portfolioData, _
    rowOffset + 1, Short_Period & " Year(s) Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Calculate and format Combined correlation matrix
    rowOffset2 = (lastColumn) * 2 + 7 + startRow
    For i = 1 To lastColumn - 1
        For j = 1 To lastColumn - 1
            combinedCorrelations(i, j) = Abs(correlations10Y(i, j)) + Abs(correlations1Y(i, j))
        Next j
    Next i
    
    Call FormatCorrelationMatrix(wsCorrelation, combinedCorrelations, allData, portfolioData, _
    rowOffset2, "Total Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Add navigation buttons
    Call AddCorrNavigationButtons(wsCorrelation, "Correlations")
    
    
    ' Final formatting
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
    
    Call OrderVisibleTabsBasedOnList
    wsCorrelation.Activate

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub




Private Sub CalculateCorrelationMatrix(ByRef correlationMatrix() As Double, _
                                     ByRef data As Variant, _
                                     ByVal startdate As Date, _
                                     ByVal endDate As Date, _
                                     ByVal lastRow As Long, _
                                     ByVal lastColumn As Long, _
                                     ByVal periodType As String)
    Dim i As Long, j As Long, K As Long
    Dim nonZeroCount As Long
    Dim dailyProfits1() As Variant, dailyProfits2() As Variant
    Dim correlationValue As Double
    
    For i = 2 To lastColumn
        For j = i To lastColumn
            nonZeroCount = 0
            ReDim dailyProfits1(1 To lastRow - 1)
            ReDim dailyProfits2(1 To lastRow - 1)
            
            For K = 2 To lastRow
                If CDate(data(K, 1)) >= startdate And CDate(data(K, 1)) <= endDate Then
                    If data(K, i) <> 0 Or data(K, j) <> 0 Then
                        nonZeroCount = nonZeroCount + 1
                        dailyProfits1(nonZeroCount) = data(K, i)
                        dailyProfits2(nonZeroCount) = data(K, j)
                    End If
                End If
            Next K
            
            If nonZeroCount > 1 Then
                ReDim Preserve dailyProfits1(1 To nonZeroCount)
                ReDim Preserve dailyProfits2(1 To nonZeroCount)
                correlationValue = Round(CalculateCorrelation(dailyProfits1, dailyProfits2), 2)
                correlationMatrix(i - 1, j - 1) = correlationValue
                correlationMatrix(j - 1, i - 1) = correlationValue
            Else
                correlationMatrix(i - 1, j - 1) = 0
                correlationMatrix(j - 1, i - 1) = 0
            End If
        Next j
        Application.StatusBar = periodType & " Correlation Running: " & Format(i / lastColumn, "0%") & " completed"
    Next i
End Sub

Private Sub FormatCorrelationMatrix(ByRef ws As Worksheet, _
                                  ByRef correlationMatrix() As Double, _
                                  ByRef data As Variant, _
                                  ByRef portfolioData As Variant, _
                                  ByVal startRow As Long, _
                                  ByVal title As String, _
                                  ByVal highThreshold As Double, _
                                  ByVal lowThreshold As Double, _
                                  ByVal lastColumn As Long)
    Dim i As Long, j As Long
    Dim strategyName As String
    Dim symbolMap As Object
    
    ' Create dictionary to map strategy names to symbols
    Set symbolMap = CreateObject("Scripting.Dictionary")
    
    ' Populate symbol map from portfolio data
    ' Assuming portfolioData has strategy names in col COL_PORT_STRATEGY_NAME and symbols in col COL_PORT_SYMBOL
    For i = 1 To UBound(portfolioData)
        If Not symbolMap.Exists(CStr(portfolioData(i, COL_PORT_STRATEGY_NAME))) Then
            symbolMap.Add CStr(portfolioData(i, COL_PORT_STRATEGY_NAME)), CStr(portfolioData(i, COL_PORT_SYMBOL))
        End If
    Next i
    
    ' Format header
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + 1, 2))
        .Merge
        .value = title
        .WrapText = True
    End With
    
    ' Set up headers
    ws.Cells(startRow, 3).value = "Strategy Number"
    ws.Cells(startRow + 1, 3).value = "Strategy Name"
    ws.Cells(startRow + 2, 3).value = "Symbol"
    ws.Cells(startRow + 2, 1).value = "Strategy Number"
    ws.Cells(startRow + 2, 2).value = "Strategy Name"
    
    ' Populate headers
    For j = 1 To lastColumn - 1
        strategyName = CStr(data(1, j + 1))
        
        ' Column headers
        ws.Cells(startRow, j + 3).value = j
        ws.Cells(startRow + 1, j + 3).value = strategyName
        ws.Cells(startRow + 2, j + 3).value = IIf(symbolMap.Exists(strategyName), _
                                                 symbolMap(strategyName), _
                                                 "Symbol Not Found")
        
        ' Row headers
        ws.Cells(startRow + 2 + j, 1).value = j
        ws.Cells(startRow + 2 + j, 2).value = strategyName
        ws.Cells(startRow + 2 + j, 3).value = IIf(symbolMap.Exists(strategyName), _
                                                 symbolMap(strategyName), _
                                                 "Symbol Not Found")
    Next j
    
    ' Populate correlation values
    For i = 1 To lastColumn - 1
        For j = 1 To lastColumn - 1
            With ws.Cells(startRow + 2 + i, j + 3)
                If i = j Then
                    .value = 0
                    .Interior.Color = RGB(148, 220, 248)
                Else
                    .value = correlationMatrix(i, j)
                    .NumberFormat = "0.00"
                    If .value > highThreshold Or .value < -highThreshold Then
                        .Interior.Color = RGB(255, 0, 0)
                    ElseIf .value > lowThreshold Or .value < -lowThreshold Then
                        .Interior.Color = RGB(255, 255, 0)
                    End If
                End If
            End With
        Next j
        Application.StatusBar = "Formatting Matrix: " & Format(i / lastColumn, "0%") & " completed"
    Next i
    
    ' Optional: Add error highlighting for missing symbols
    For j = 1 To lastColumn - 1
        If ws.Cells(startRow + 2, j + 3).value = "Symbol Not Found" Then
            ws.Cells(startRow + 2, j + 3).Interior.Color = RGB(255, 200, 200)
        End If
        If ws.Cells(startRow + 2 + j, 3).value = "Symbol Not Found" Then
            ws.Cells(startRow + 2 + j, 3).Interior.Color = RGB(255, 200, 200)
        End If
    Next j
End Sub

Private Sub FormatATRPNLCorrelationMatrix(ByRef ws As Worksheet, _
                                        ByRef correlationMatrix() As Double, _
                                        ByRef data As Variant, _
                                        ByVal startRow As Long, _
                                        ByVal title As String, _
                                        ByVal highThreshold As Double, _
                                        ByVal lowThreshold As Double, _
                                        ByVal lastColumn As Long)
    Dim i As Long, j As Long
    Dim contractName As String
    
    ' Format header
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 2))
        .Merge
        .value = title
        .WrapText = True
    End With
    
    ' Set up headers
    ws.Cells(startRow + 1, 1).value = "Contract Number"
    ws.Cells(startRow + 1, 2).value = "Contract Name"
    
    ' Populate headers
    For j = 1 To lastColumn - 1
        contractName = CStr(data(1, j + 1))
        
        ' Column headers
        ws.Cells(startRow, j + 2).value = j
        ws.Cells(startRow + 1, j + 2).value = contractName
        
        ' Row headers
        ws.Cells(startRow + 1 + j, 1).value = j
        ws.Cells(startRow + 1 + j, 2).value = contractName
    Next j
    
    ' Populate correlation values
    For i = 1 To lastColumn - 1
        For j = 1 To lastColumn - 1
            With ws.Cells(startRow + 1 + i, j + 2)
                If i = j Then
                    .value = 0
                    .Interior.Color = RGB(148, 220, 248)
                Else
                    .value = correlationMatrix(i, j)
                    End If
                    .NumberFormat = "0.00"
                    If .value > highThreshold Or .value < -highThreshold Then
                        .Interior.Color = RGB(255, 0, 0)
                    ElseIf .value > lowThreshold Or .value < -lowThreshold Then
                        .Interior.Color = RGB(255, 255, 0)
                End If
            End With
        Next j
        Application.StatusBar = "Formatting Matrix: " & Format(i / lastColumn, "0%") & " completed"
    Next i
End Sub

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function


Sub CreateNegativeCorrelationMatrices()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Dim wsPortfolioDailyM2M As Worksheet
    Dim wsPortfolio As Worksheet
    Dim wsNegativeCorrelation As Worksheet
    Dim lastRow As Long, lastColumn As Long
    Dim allData As Variant
    Dim portfolioData As Variant
    Dim correlations10Y() As Double
    Dim correlations1Y() As Double
    Dim combinedCorrelations() As Double
    Dim startdate As Date
    Dim Current_Date As Date
    Dim rowOffset As Long, rowOffset2 As Long
    Dim i As Long, j As Long
    Dim startRow As Long
        
    ' Initialize constants and check license
    Call InitializeColumnConstantsManually
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        GoTo CleanExit
    End If
    
    ' Get threshold values
    Dim High_Threshold As Double: High_Threshold = GetNamedRangeValue("NegativeCorrel_High_Threshold")
    Dim Low_Threshold As Double: Low_Threshold = GetNamedRangeValue("NegativeCorrel_Low_Threshold")
    Dim Short_Period As Double: Short_Period = GetNamedRangeValue("Correl_Short_Period")
    Dim Long_Period As Double: Long_Period = GetNamedRangeValue("Correl_Long_Period")
    
    ' Set and validate worksheets
    If Not SheetExists("Portfolio") Or Not SheetExists("PortfolioDailyM2M") Then
        MsgBox "Required sheets missing!", vbCritical
        GoTo CleanExit
    End If
    
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set wsPortfolioDailyM2M = ThisWorkbook.Sheets("PortfolioDailyM2M")
    
    If wsPortfolio.Cells(2, COL_PORT_STRATEGY_NAME).value = "" Then
        MsgBox "Error: 'Portfolio' sheet exists but contains no data in row 2.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get dimensions and load data
    lastRow = wsPortfolioDailyM2M.Cells(wsPortfolioDailyM2M.rows.count, 1).End(xlUp).row
    lastColumn = wsPortfolioDailyM2M.Cells(1, wsPortfolioDailyM2M.Columns.count).End(xlToLeft).column
    
    allData = wsPortfolioDailyM2M.Range(wsPortfolioDailyM2M.Cells(1, 1), _
                                      wsPortfolioDailyM2M.Cells(lastRow, lastColumn)).value
    portfolioData = wsPortfolio.Range(wsPortfolio.Cells(2, 1), _
                                    wsPortfolio.Cells(lastColumn, COL_PORT_SYMBOL)).value
    
    ' Initialize arrays
    ReDim correlations10Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim correlations1Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim combinedCorrelations(1 To lastColumn - 1, 1 To lastColumn - 1)
    
    Current_Date = wsPortfolio.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    
    ' Calculate Long Period Correlations
    startdate = DateAdd("yyyy", -Int(Long_Period), Current_Date)
    startdate = DateAdd("d", -Int((Long_Period - Int(Long_Period)) * 365.25), startdate)
    CalculateNegativeCorrelationMatrix correlations10Y, allData, startdate, Current_Date, lastRow, lastColumn, "Long"
    
    ' Calculate Short Period Correlations
    startdate = DateAdd("yyyy", -Int(Short_Period), Current_Date)
    startdate = DateAdd("d", -Int((Short_Period - Int(Short_Period)) * 365.25), startdate)
    CalculateNegativeCorrelationMatrix correlations1Y, allData, startdate, Current_Date, lastRow, lastColumn, "Short"
    
    ' Create and format new sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("NegativeCorrelations").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    Set wsNegativeCorrelation = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsNegativeCorrelation.name = "NegativeCorrelations"
    'wsNegativeCorrelation.Tab.Color = RGB(117, 219, 255)
    
     ' Set white background color for the entire worksheet
    wsNegativeCorrelation.Cells.Interior.Color = RGB(255, 255, 255)
    
    startRow = 4
    
    ' Format Long Period correlation matrix
    FormatCorrelationMatrix wsNegativeCorrelation, correlations10Y, allData, portfolioData, _
                           (startRow + 1), Long_Period & " Year(s) Negative Correlations", _
                           High_Threshold, Low_Threshold, lastColumn
    
    ' Format Short Period correlation matrix
    rowOffset = lastColumn + 3 + startRow
    FormatCorrelationMatrix wsNegativeCorrelation, correlations1Y, allData, portfolioData, _
                           rowOffset + 1, Short_Period & " Year(s) Negative Correlations", _
                           High_Threshold, Low_Threshold, lastColumn
    
    ' Calculate and format combined correlations
    rowOffset2 = (lastColumn) * 2 + 7 + startRow
    For i = 1 To lastColumn - 1
        For j = 1 To lastColumn - 1
            combinedCorrelations(i, j) = Abs(correlations10Y(i, j)) + Abs(correlations1Y(i, j))
        Next j
    Next i
    
    FormatCorrelationMatrix wsNegativeCorrelation, combinedCorrelations, allData, portfolioData, _
                           rowOffset2, "Total Negative Correlations", _
                           High_Threshold, Low_Threshold, lastColumn
    
    
     ' Add navigation buttons
    Call AddCorrNavigationButtons(wsNegativeCorrelation, "NegativeCorrelations")
    
    ' Final formatting
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
    
    Call OrderVisibleTabsBasedOnList
    wsNegativeCorrelation.Activate
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub


Private Sub CalculateNegativeCorrelationMatrix(ByRef correlationMatrix() As Double, _
                                             ByRef data As Variant, _
                                             ByVal startdate As Date, _
                                             ByVal endDate As Date, _
                                             ByVal lastRow As Long, _
                                             ByVal lastColumn As Long, _
                                             ByVal periodType As String)
    Dim i As Long, j As Long, K As Long
    Dim nonZeroCount As Long
    Dim dailyProfits1() As Variant, dailyProfits2() As Variant
    Dim correlationValue As Double
    
    For i = 2 To lastColumn
        For j = i To lastColumn
            nonZeroCount = 0
            ReDim dailyProfits1(1 To lastRow - 1)
            ReDim dailyProfits2(1 To lastRow - 1)
            
            For K = 2 To lastRow
                If CDate(data(K, 1)) >= startdate And CDate(data(K, 1)) <= endDate Then
                    ' Changed condition to check for negative correlation pattern
                    If Not (data(K, i) > 0 And data(K, j) > 0) Then
                        nonZeroCount = nonZeroCount + 1
                        dailyProfits1(nonZeroCount) = data(K, i)
                        dailyProfits2(nonZeroCount) = data(K, j)
                    End If
                End If
            Next K
            
            If nonZeroCount > 1 Then
                ReDim Preserve dailyProfits1(1 To nonZeroCount)
                ReDim Preserve dailyProfits2(1 To nonZeroCount)
                correlationValue = Round(CalculateCorrelation(dailyProfits1, dailyProfits2), 2)
                correlationMatrix(i - 1, j - 1) = correlationValue
                correlationMatrix(j - 1, i - 1) = correlationValue
            Else
                correlationMatrix(i - 1, j - 1) = 0
                correlationMatrix(j - 1, i - 1) = 0
            End If
        Next j
        Application.StatusBar = periodType & " Negative Correlation Running: " & Format(i / lastColumn, "0%") & " completed"
    Next i
End Sub




Function CalculateCorrelation(dailyProfits1 As Variant, dailyProfits2 As Variant) As Double
    Dim minLength As Long, i As Long
    Dim tempArray1() As Double, tempArray2() As Double
    Dim cleanedCount As Long
    Dim sumArray1 As Double, sumArray2 As Double
    Dim corrResult As Variant
    
    ' Determine the minimum length of the input arrays
    minLength = Application.Min(UBound(dailyProfits1), UBound(dailyProfits2))
    
    ' Build cleaned arrays of that length
    ReDim tempArray1(1 To minLength)
    ReDim tempArray2(1 To minLength)
    
    For i = 1 To minLength
        If IsNumeric(dailyProfits1(i)) And IsNumeric(dailyProfits2(i)) Then
            cleanedCount = cleanedCount + 1
            tempArray1(cleanedCount) = CDbl(dailyProfits1(i))
            tempArray2(cleanedCount) = CDbl(dailyProfits2(i))
        End If
    Next i
    
    ' If too few points, bail
    If cleanedCount < 2 Then
        CalculateCorrelation = 0
        Exit Function
    End If
    
    ' Trim arrays down to the cleaned count
    ReDim Preserve tempArray1(1 To cleanedCount)
    ReDim Preserve tempArray2(1 To cleanedCount)
    
    ' Guard against zero-sum (zero variance) which also throws
    sumArray1 = Application.sum(tempArray1)
    sumArray2 = Application.sum(tempArray2)
    If sumArray1 = 0 Or sumArray2 = 0 Then
        CalculateCorrelation = 0
        Exit Function
    End If
    
    ' Use Application.Correl so that errors come back as Variants, not VBA exceptions
    corrResult = Application.Correl(tempArray1, tempArray2)
    If IsError(corrResult) Then
        CalculateCorrelation = 0
    Else
        CalculateCorrelation = Round(corrResult, 2)
    End If
End Function





Private Sub AddCorrNavigationButtons(ByRef wsSheet As Worksheet, tabname As String)
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    
    ' Delete Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 2).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    
    If tabname = "Correlations" Then
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteCorrelations"
    End With
    ElseIf tabname = "NegativeCorrelations" Then
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteNegativeCorrelations"
    End With
    Else
    With btn
        .Caption = "Error"
    End With
    End If
    ' Summary Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 5).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With
    
    ' Portfolio Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 8).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With
    
    ' Control Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 11).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With
    
    ' Strategies Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 14).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies"
    End With
    
    ' Inputs Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 17).left, _
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



Sub CreateDrawdownCorrelationMatrices()
    ' Application settings
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Declare all variables
    Dim wsPortfolioDailyM2M As Worksheet
    Dim wsPortfolio As Worksheet
    Dim wsDrawdownCorrelation As Worksheet
    Dim lastRow As Long, lastColumn As Long
    Dim i As Long, j As Long, K As Long
    Dim nonZeroCount As Long
    Dim correlationValue As Double
    Dim rowOffset As Long, rowOffset2 As Long
    Dim allData As Variant
    Dim drawdownData As Variant
    Dim portfolioData As Variant
    Dim startdate As Date, Current_Date As Date
    Dim dailyProfits1() As Variant, dailyProfits2() As Variant
    Dim correlations10Y() As Double
    Dim correlations1Y() As Double
    Dim combinedCorrelations() As Double
    Dim startRow As Long
    ' Cache threshold values
    Dim High_Threshold As Double
    Dim Low_Threshold As Double
    Dim Short_Period As Double
    Dim Long_Period As Double
    
    ' Initialize column constants
    Call InitializeColumnConstantsManually
    
    ' Check license
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        GoTo CleanExit
    End If
    
    ' Validate required sheets exist
    If Not SheetExists("Portfolio") Or Not SheetExists("PortfolioDailyM2M") Then
        MsgBox "Required sheets missing!", vbCritical
        GoTo CleanExit
    End If
    
    ' Set worksheet references
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set wsPortfolioDailyM2M = ThisWorkbook.Sheets("PortfolioDailyM2M")
    
    ' Validate data exists
    If wsPortfolio.Cells(2, COL_PORT_STRATEGY_NAME).value = "" Then
        MsgBox "Error: 'Portfolio' sheet exists but contains no data in row 2.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get threshold values
    High_Threshold = GetNamedRangeValue("DDCorrel_High_Threshold")
    Low_Threshold = GetNamedRangeValue("DDCorrel_Low_Threshold")
    Short_Period = GetNamedRangeValue("Correl_Short_Period")
    Long_Period = GetNamedRangeValue("Correl_Long_Period")
    Current_Date = wsPortfolio.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    
    ' Validate threshold values
    If High_Threshold <= Low_Threshold Then
        MsgBox "Invalid threshold values. High threshold must be greater than low threshold.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get data dimensions
    lastRow = wsPortfolioDailyM2M.Cells(wsPortfolioDailyM2M.rows.count, 1).End(xlUp).row
    lastColumn = wsPortfolioDailyM2M.Cells(1, wsPortfolioDailyM2M.Columns.count).End(xlToLeft).column
    
    ' Read all data into arrays for faster processing
    allData = wsPortfolioDailyM2M.Range(wsPortfolioDailyM2M.Cells(1, 1), _
                                      wsPortfolioDailyM2M.Cells(lastRow, lastColumn)).value
    portfolioData = wsPortfolio.Range(wsPortfolio.Cells(2, 1), _
                                    wsPortfolio.Cells(lastColumn, COL_PORT_SYMBOL)).value
    
    ' Convert equity data to drawdown data
    ReDim drawdownData(1 To UBound(allData, 1), 1 To UBound(allData, 2))
    
    ' Copy headers
    For j = 1 To UBound(allData, 2)
        drawdownData(1, j) = allData(1, j)
    Next j
    
    ' Calculate drawdowns for each strategy
    For j = 2 To UBound(allData, 2)
        Dim peak As Double: peak = 0
        For i = 2 To UBound(allData, 1)
            ' Copy date
            If j = 2 Then drawdownData(i, 1) = allData(i, 1)
            
            ' Find peak and calculate drawdown
            If CDbl(allData(i, j)) > peak Then
                peak = CDbl(allData(i, j))
                drawdownData(i, j) = 0 ' At peak, no drawdown
            Else
                ' Calculate drawdown as percentage from peak
                If peak > 0 Then
                    drawdownData(i, j) = (CDbl(allData(i, j)) - peak) / peak * 100
                Else
                    drawdownData(i, j) = 0
                End If
            End If
        Next i
        Application.StatusBar = "Converting to drawdown data: " & Format(j / UBound(allData, 2), "0%") & " completed"
    Next j
    
    ' Initialize correlation arrays
    ReDim correlations10Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim correlations1Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim combinedCorrelations(1 To lastColumn - 1, 1 To lastColumn - 1)
    
    ' Calculate Long Period Correlations
    startdate = DateAdd("yyyy", -Int(Long_Period), Current_Date)
    startdate = DateAdd("d", -Int((Long_Period - Int(Long_Period)) * 365.25), startdate)
    
    Application.StatusBar = "Calculating long-term drawdown correlations..."
    Call CalculateDrawdownCorrelationMatrix(correlations10Y, drawdownData, startdate, Current_Date, lastRow, lastColumn, "Long")
    
    ' Calculate Short Period Correlations
    startdate = DateAdd("yyyy", -Int(Short_Period), Current_Date)
    startdate = DateAdd("d", -Int((Short_Period - Int(Short_Period)) * 365.25), startdate)
    
    Application.StatusBar = "Calculating short-term drawdown correlations..."
    Call CalculateDrawdownCorrelationMatrix(correlations1Y, drawdownData, startdate, Current_Date, lastRow, lastColumn, "Short")
    
    ' Create new correlation sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("DrawdownCorrelations").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    Set wsDrawdownCorrelation = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsDrawdownCorrelation.name = "DrawdownCorrelations"
    wsDrawdownCorrelation.Tab.Color = RGB(255, 150, 150)
    
    startRow = 4
    
    ' Format and populate Long Period correlation matrix
    Call FormatCorrelationMatrix(wsDrawdownCorrelation, correlations10Y, drawdownData, portfolioData, _
    (startRow + 1), Long_Period & " Year(s) Drawdown Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Format and populate Short Period correlation matrix
    rowOffset = lastColumn + 3 + startRow
    Call FormatCorrelationMatrix(wsDrawdownCorrelation, correlations1Y, drawdownData, portfolioData, _
    rowOffset + 1, Short_Period & " Year(s) Drawdown Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Calculate and format Combined correlation matrix
    rowOffset2 = (lastColumn) * 2 + 7 + startRow
    For i = 1 To lastColumn - 1
        For j = 1 To lastColumn - 1
            combinedCorrelations(i, j) = Abs(correlations10Y(i, j)) + Abs(correlations1Y(i, j))
        Next j
    Next i
    
    Call FormatCorrelationMatrix(wsDrawdownCorrelation, combinedCorrelations, drawdownData, portfolioData, _
    rowOffset2, "Total Drawdown Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Add navigation buttons
    Call AddDrawdownCorrNavigationButtons(wsDrawdownCorrelation, "DrawdownCorrelations")
    
    ' Final formatting
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
    
    Call OrderVisibleTabsBasedOnList
    wsDrawdownCorrelation.Activate

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub CalculateDrawdownCorrelationMatrix(ByRef correlationMatrix() As Double, _
                                             ByRef drawdownData As Variant, _
                                             ByVal startdate As Date, _
                                             ByVal endDate As Date, _
                                             ByVal lastRow As Long, _
                                             ByVal lastColumn As Long, _
                                             ByVal periodType As String)
    Dim i As Long, j As Long, K As Long
    Dim nonZeroCount As Long
    Dim dailyProfits1() As Variant, dailyProfits2() As Variant
    Dim correlationValue As Double
    
    For i = 2 To lastColumn
        For j = i To lastColumn
            nonZeroCount = 0
            ReDim dailyProfits1(1 To lastRow - 1)
            ReDim dailyProfits2(1 To lastRow - 1)
            
            For K = 2 To lastRow
                If CDate(drawdownData(K, 1)) >= startdate And CDate(drawdownData(K, 1)) <= endDate Then
                    ' For drawdowns, we want to include all points where there is a drawdown in either strategy
                    If drawdownData(K, i) <> 0 Or drawdownData(K, j) <> 0 Then
                        nonZeroCount = nonZeroCount + 1
                        dailyProfits1(nonZeroCount) = drawdownData(K, i)
                        dailyProfits2(nonZeroCount) = drawdownData(K, j)
                    End If
                End If
            Next K
            
            If nonZeroCount > 1 Then
                ReDim Preserve dailyProfits1(1 To nonZeroCount)
                ReDim Preserve dailyProfits2(1 To nonZeroCount)
                correlationValue = Round(CalculateCorrelation(dailyProfits1, dailyProfits2), 2)
                correlationMatrix(i - 1, j - 1) = correlationValue
                correlationMatrix(j - 1, i - 1) = correlationValue
            Else
                correlationMatrix(i - 1, j - 1) = 0
                correlationMatrix(j - 1, i - 1) = 0
            End If
        Next j
        Application.StatusBar = periodType & " Drawdown Correlation Running: " & Format(i / lastColumn, "0%") & " completed"
    Next i
End Sub

Private Sub AddDrawdownCorrNavigationButtons(ByRef wsSheet As Worksheet, tabname As String)
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    
    ' Delete Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 2).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteDrawdownCorrelations"
    End With
    
    ' Summary Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 5).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With
    
    ' Portfolio Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 8).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With
    
    ' Control Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 11).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With
    
    ' Strategies Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 14).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies"
    End With
    
    ' Inputs Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 17).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs"
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddDrawdownCorrNavigationButtons: " & Err.Description
End Sub

' Delete button handler
Sub DeleteDrawdownCorrelations()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("DrawdownCorrelations").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Go to Summary tab
    Call GoToSummary
End Sub

Sub CreateTrueRangesCorrelationMatrices()
    ' Application settings
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Declare all variables
    Dim wsTrueRanges As Worksheet
    Dim wsTrueRangesCorrelation As Worksheet
    Dim lastRow As Long, lastColumn As Long
    Dim i As Long, j As Long, K As Long
    Dim nonZeroCount As Long
    Dim correlationValue As Double
    Dim rowOffset As Long, rowOffset2 As Long
    Dim allData As Variant
    Dim startdate As Date, Current_Date As Date
    Dim dailyValues1() As Variant, dailyValues2() As Variant
    Dim correlations10Y() As Double
    Dim correlations1Y() As Double
    Dim combinedCorrelations() As Double
    Dim startRow As Long
    ' Cache threshold values
    Dim High_Threshold As Double
    Dim Low_Threshold As Double
    Dim Short_Period As Double
    Dim Long_Period As Double
    
    ' Initialize column constants
    Call InitializeColumnConstantsManually
    
    ' Check license
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        GoTo CleanExit
    End If
    
    ' Validate required sheets exist
    If Not SheetExists("TrueRanges") Then
        MsgBox "TrueRanges sheet missing! Make sure you've run data processing first.", vbCritical
        GoTo CleanExit
    End If
    
    ' Set worksheet references
    Set wsTrueRanges = ThisWorkbook.Sheets("TrueRanges")
    
    ' Validate data exists
    If wsTrueRanges.Cells(2, 1).value = "" Then
        MsgBox "Error: 'TrueRanges' sheet exists but contains no data in row 2.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get threshold values
    High_Threshold = GetNamedRangeValue("Correl_High_Threshold")
    Low_Threshold = GetNamedRangeValue("Correl_Low_Threshold")
    Short_Period = GetNamedRangeValue("Correl_Short_Period")
    Long_Period = GetNamedRangeValue("Correl_Long_Period")
    Current_Date = wsTrueRanges.Cells(wsTrueRanges.rows.count, 1).End(xlUp).value
    
    ' Validate threshold values
    If High_Threshold <= Low_Threshold Then
        MsgBox "Invalid threshold values. High threshold must be greater than low threshold.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get data dimensions
    lastRow = EndRowByCutoffSimple(wsTrueRanges, 1)
    lastColumn = wsTrueRanges.Cells(1, wsTrueRanges.Columns.count).End(xlToLeft).column
    
    ' Read all data into arrays for faster processing
    allData = wsTrueRanges.Range(wsTrueRanges.Cells(1, 1), _
                                wsTrueRanges.Cells(lastRow, lastColumn)).value
    
    ' Initialize correlation arrays
    ReDim correlations10Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim correlations1Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim combinedCorrelations(1 To lastColumn - 1, 1 To lastColumn - 1)
    
    ' Calculate Long Period Correlations
    startdate = DateAdd("yyyy", -Int(Long_Period), Current_Date)
    startdate = DateAdd("d", -Int((Long_Period - Int(Long_Period)) * 365.25), startdate)
    
    Application.StatusBar = "Calculating long-term ATR correlations..."
    Call CalculateCorrelationMatrix(correlations10Y, allData, startdate, Current_Date, lastRow, lastColumn, "Long")
    
    ' Calculate Short Period Correlations
    startdate = DateAdd("yyyy", -Int(Short_Period), Current_Date)
    startdate = DateAdd("d", -Int((Short_Period - Int(Short_Period)) * 365.25), startdate)
    
    Application.StatusBar = "Calculating short-term ATR correlations..."
    Call CalculateCorrelationMatrix(correlations1Y, allData, startdate, Current_Date, lastRow, lastColumn, "Short")
    
    ' Create new correlation sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("ATRCorrelations").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    Set wsTrueRangesCorrelation = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("TrueRanges"))
    wsTrueRangesCorrelation.name = "ATRCorrelations"
    wsTrueRangesCorrelation.Tab.Color = RGB(255, 192, 0)
    
    startRow = 4
    
    ' Format and populate Long Period correlation matrix
    Call FormatATRPNLCorrelationMatrix(wsTrueRangesCorrelation, correlations10Y, allData, _
    (startRow + 1), Long_Period & " Year(s) ATR Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Format and populate Short Period correlation matrix
    rowOffset = lastColumn + 3 + startRow
    Call FormatATRPNLCorrelationMatrix(wsTrueRangesCorrelation, correlations1Y, allData, _
    rowOffset + 1, Short_Period & " Year(s) ATR Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Calculate and format Combined correlation matrix
    rowOffset2 = (lastColumn) * 2 + 7 + startRow
    For i = 1 To lastColumn - 1
        For j = 1 To lastColumn - 1
            combinedCorrelations(i, j) = correlations10Y(i, j) + correlations1Y(i, j)
        Next j
    Next i
    
    Call FormatATRPNLCorrelationMatrix(wsTrueRangesCorrelation, combinedCorrelations, allData, _
    rowOffset2, "Total ATR Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Add navigation buttons
    Call AddATRCorrNavigationButtons(wsTrueRangesCorrelation, "ATRCorrelations")
    
    ' Final formatting
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
    
    Call OrderVisibleTabsBasedOnList
    wsTrueRangesCorrelation.Activate

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub AddATRCorrNavigationButtons(ByRef wsSheet As Worksheet, tabname As String)
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    
    ' Delete Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 2).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteATRCorrelations"
    End With
    
    ' Summary Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 5).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With
    
    ' Portfolio Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 8).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With
    
    ' Control Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 11).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With
    
    ' Strategies Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 14).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies"
    End With
    
    ' Inputs Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 17).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs"
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddATRCorrNavigationButtons: " & Err.Description
End Sub

' Delete button handler
Sub DeleteATRCorrelations()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("ATRCorrelations").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Go to Summary tab
    Call GoToSummary
End Sub


Sub CreateTradePNLCorrelationMatrices()
    ' Application settings
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Declare all variables
    Dim wsTradePNL As Worksheet
    Dim wsPNLCorrelation As Worksheet
    Dim lastRow As Long, lastColumn As Long
    Dim i As Long, j As Long, K As Long
    Dim nonZeroCount As Long
    Dim correlationValue As Double
    Dim rowOffset As Long, rowOffset2 As Long
    Dim allData As Variant
    Dim startdate As Date, Current_Date As Date
    Dim dailyValues1() As Variant, dailyValues2() As Variant
    Dim correlations10Y() As Double
    Dim correlations1Y() As Double
    Dim combinedCorrelations() As Double
    Dim startRow As Long
    ' Cache threshold values
    Dim High_Threshold As Double
    Dim Low_Threshold As Double
    Dim Short_Period As Double
    Dim Long_Period As Double
    
    ' Initialize column constants
    Call InitializeColumnConstantsManually
    
    ' Check license
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        GoTo CleanExit
    End If
    
    ' Validate required sheets exist
    If Not SheetExists("TradePNL") Then
        MsgBox "TradePNL sheet missing! Make sure you've run data processing first.", vbCritical
        GoTo CleanExit
    End If
    
    ' Set worksheet references
    Set wsTradePNL = ThisWorkbook.Sheets("TradePNL")
    
    ' Validate data exists
    If wsTradePNL.Cells(2, 1).value = "" Then
        MsgBox "Error: 'TradePNL' sheet exists but contains no data in row 2.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get threshold values
    High_Threshold = GetNamedRangeValue("Correl_High_Threshold")
    Low_Threshold = GetNamedRangeValue("Correl_Low_Threshold")
    Short_Period = GetNamedRangeValue("Correl_Short_Period")
    Long_Period = GetNamedRangeValue("Correl_Long_Period")
    Current_Date = wsTradePNL.Cells(wsTradePNL.rows.count, 1).End(xlUp).value
    
    ' Validate threshold values
    If High_Threshold <= Low_Threshold Then
        MsgBox "Invalid threshold values. High threshold must be greater than low threshold.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get data dimensions
    lastRow = wsTradePNL.Cells(wsTradePNL.rows.count, 1).End(xlUp).row
    lastColumn = wsTradePNL.Cells(1, wsTradePNL.Columns.count).End(xlToLeft).column
    
    ' Read all data into arrays for faster processing
    allData = wsTradePNL.Range(wsTradePNL.Cells(1, 1), _
                              wsTradePNL.Cells(lastRow, lastColumn)).value
    
    ' Initialize correlation arrays
    ReDim correlations10Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim correlations1Y(1 To lastColumn - 1, 1 To lastColumn - 1)
    ReDim combinedCorrelations(1 To lastColumn - 1, 1 To lastColumn - 1)
    
    ' Calculate Long Period Correlations
    startdate = DateAdd("yyyy", -Int(Long_Period), Current_Date)
    startdate = DateAdd("d", -Int((Long_Period - Int(Long_Period)) * 365.25), startdate)
    
    Application.StatusBar = "Calculating long-term PNL correlations..."
    Call CalculateCorrelationMatrix(correlations10Y, allData, startdate, Current_Date, lastRow, lastColumn, "Long")
    
    ' Calculate Short Period Correlations
    startdate = DateAdd("yyyy", -Int(Short_Period), Current_Date)
    startdate = DateAdd("d", -Int((Short_Period - Int(Short_Period)) * 365.25), startdate)
    
    Application.StatusBar = "Calculating short-term PNL correlations..."
    Call CalculateCorrelationMatrix(correlations1Y, allData, startdate, Current_Date, lastRow, lastColumn, "Short")
    
    ' Create new correlation sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("PNLCorrelations").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    Set wsPNLCorrelation = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("TradePNL"))
    wsPNLCorrelation.name = "PNLCorrelations"
    wsPNLCorrelation.Tab.Color = RGB(255, 192, 0)
    
    startRow = 4
    
    ' Format and populate Long Period correlation matrix
    Call FormatATRPNLCorrelationMatrix(wsPNLCorrelation, correlations10Y, allData, _
    (startRow + 1), Long_Period & " Year(s) PNL Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Format and populate Short Period correlation matrix
    rowOffset = lastColumn + 3 + startRow
    Call FormatATRPNLCorrelationMatrix(wsPNLCorrelation, correlations1Y, allData, _
    rowOffset + 1, Short_Period & " Year(s) PNL Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Calculate and format Combined correlation matrix
    rowOffset2 = (lastColumn) * 2 + 7 + startRow
    For i = 1 To lastColumn - 1
        For j = 1 To lastColumn - 1
            combinedCorrelations(i, j) = correlations10Y(i, j) + correlations1Y(i, j)
        Next j
    Next i
    
    Call FormatATRPNLCorrelationMatrix(wsPNLCorrelation, combinedCorrelations, allData, _
    rowOffset2, "Total PNL Correlations", _
    High_Threshold, Low_Threshold, lastColumn)
    
    ' Add navigation buttons
    Call AddPNLCorrNavigationButtons(wsPNLCorrelation, "PNLCorrelations")
    
    ' Final formatting
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
    
    Call OrderVisibleTabsBasedOnList
    wsPNLCorrelation.Activate

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub AddPNLCorrNavigationButtons(ByRef wsSheet As Worksheet, tabname As String)
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    
    ' Delete Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 2).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeletePNLCorrelations"
    End With
    
    ' Summary Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 5).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With
    
    ' Portfolio Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 8).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With
    
    ' Control Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 11).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With
    
    ' Strategies Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 14).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies"
    End With
    
    ' Inputs Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 17).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs"
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddPNLCorrNavigationButtons: " & Err.Description
End Sub

' Delete button handler
Sub DeletePNLCorrelations()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("PNLCorrelations").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Go to Summary tab
    Call GoToSummary
End Sub








'-------------------------

Sub CreateCorrelationPeriodAnalysis()
    ' Application settings
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Declare all variables
    Dim wsPortfolioDailyM2M As Worksheet
    Dim wsPortfolio As Worksheet
    Dim wsPeriodAnalysis As Worksheet
    Dim rnPeriods As Range
    Dim lastRow As Long, lastColumn As Long
    Dim i As Long, j As Long, K As Long, p As Long
    Dim allData As Variant
    Dim portfolioData As Variant
    Dim periodTitle As String
    Dim fromDate As Date, toDate As Date
    Dim currentRow As Long
    Dim correlationMatrix() As Double
    Dim chartLeft As Double, chartTop As Double
    Dim chartWidth As Double, chartHeight As Double
    Dim correlationLeft As Long, equityCurveLeft As Long
    Dim matrixHeight As Long
    Dim seriesNames As Variant
    Dim periodCount As Long
    Dim validPeriod As Boolean
    Dim tempSheetName As String
    
    ' Cache threshold values
    Dim High_Threshold As Double
    Dim Low_Threshold As Double
    
    ' Initialize column constants
    Call InitializeColumnConstantsManually
    
    ' Check license
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        GoTo CleanExit
    End If
    
    ' Validate required sheets exist
    If Not SheetExists("Portfolio") Or Not SheetExists("PortfolioDailyM2M") Then
        MsgBox "Required sheets missing!", vbCritical
        GoTo CleanExit
    End If
    
    ' Use a random temporary sheet name to avoid conflicts
    tempSheetName = "Temp_" & Format(Now(), "yyyymmddhhmmss")
    
    ' Check if CorrelationPeriodAnalysis named range exists
    On Error Resume Next
    Set rnPeriods = ThisWorkbook.Names("CorrelationPeriodAnalysis").RefersToRange
    On Error GoTo ErrorHandler
    
    If rnPeriods Is Nothing Then
        MsgBox "CorrelationPeriodAnalysis named range not found!", vbCritical
        GoTo CleanExit
    End If
    
    ' Ensure CorrelationPeriodAnalysis has expected structure
    If rnPeriods.Columns.count < 3 Then
        MsgBox "CorrelationPeriodAnalysis range should have at least 3 columns: Title, From Date, To Date", vbCritical
        GoTo CleanExit
    End If
    
    ' Count valid periods (non-blank titles)
    periodCount = 0
    For p = 2 To rnPeriods.rows.count ' Assuming first row is headers
        If Not IsEmpty(rnPeriods.Cells(p, 1).value) Then
            periodCount = periodCount + 1
        End If
    Next p
    
    If periodCount = 0 Then
        MsgBox "No periods found in CorrelationPeriodAnalysis range!", vbCritical
        GoTo CleanExit
    End If
    
    ' Set worksheet references
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set wsPortfolioDailyM2M = ThisWorkbook.Sheets("PortfolioDailyM2M")
    
    ' Validate data exists
    If wsPortfolio.Cells(2, COL_PORT_STRATEGY_NAME).value = "" Then
        MsgBox "Error: 'Portfolio' sheet exists but contains no data in row 2.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get threshold values
    High_Threshold = GetNamedRangeValue("Correl_High_Threshold")
    Low_Threshold = GetNamedRangeValue("Correl_Low_Threshold")
    
    ' Validate threshold values
    If High_Threshold <= Low_Threshold Then
        MsgBox "Invalid threshold values. High threshold must be greater than low threshold.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get data dimensions
    lastRow = wsPortfolioDailyM2M.Cells(wsPortfolioDailyM2M.rows.count, 1).End(xlUp).row
    lastColumn = wsPortfolioDailyM2M.Cells(1, wsPortfolioDailyM2M.Columns.count).End(xlToLeft).column
    
    ' Read all data into arrays for faster processing
    allData = wsPortfolioDailyM2M.Range(wsPortfolioDailyM2M.Cells(1, 1), _
                                      wsPortfolioDailyM2M.Cells(lastRow, lastColumn)).value
    portfolioData = wsPortfolio.Range(wsPortfolio.Cells(2, 1), _
                                    wsPortfolio.Cells(lastColumn, COL_PORT_SYMBOL)).value
    
    ' Get series names for equity curves
    ReDim seriesNames(1 To lastColumn - 1)
    For i = 2 To lastColumn
        seriesNames(i - 1) = allData(1, i)
    Next i
    
    ' Create new period analysis sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    If SheetExists("CorrelationPeriodAnalysis") Then
        ThisWorkbook.Sheets("CorrelationPeriodAnalysis").Delete
    End If
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    ' Add the new worksheet
    Set wsPeriodAnalysis = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsPeriodAnalysis.name = "CorrelationPeriodAnalysis"
    wsPeriodAnalysis.Tab.Color = RGB(117, 219, 255)
    
    ' Create and format new sheet
    wsPeriodAnalysis.Cells.Interior.Color = RGB(255, 255, 255)
    
    ' Add title and header
    With wsPeriodAnalysis.Range("A1:F1")
        .Merge
        .value = "Correlation Period Analysis"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' Add navigation buttons
    Call AddPeriodAnalysisNavigationButtons(wsPeriodAnalysis)
    
    ' Set initial row for first period
    currentRow = 4
    
    ' Initialize dimensions for matrices and charts
    correlationLeft = 2  ' Column B
    Dim strategyChartLeft As Long, totalChartLeft As Long
    strategyChartLeft = lastColumn + 5  ' Give some space after correlation matrix
    totalChartLeft = lastColumn + 40 ' Further right for total chart
    matrixHeight = lastColumn + 5  ' Approx height of matrix
      chartWidth = 1500
    
    ' Process each period
    For p = 2 To rnPeriods.rows.count ' Assuming first row is headers
        ' Skip if title is blank
        If IsEmpty(rnPeriods.Cells(p, 1).value) Then
            GoTo NextPeriod
        End If
        
        ' Get period details
        periodTitle = rnPeriods.Cells(p, 1).value
        
        ' Validate dates - skip if not valid dates
        On Error Resume Next
        fromDate = CDate(rnPeriods.Cells(p, 2).value)
        toDate = CDate(rnPeriods.Cells(p, 3).value)
        
        validPeriod = Not (IsError(fromDate) Or IsError(toDate))
        On Error GoTo ErrorHandler
        
        If Not validPeriod Then
            MsgBox "Invalid date format in period: " & periodTitle, vbExclamation
            GoTo NextPeriod
        End If
        
        ' Check if dates are within data range
        Dim minDate As Date, MaxDate As Date
        minDate = CDate(allData(2, 1))
        MaxDate = CDate(allData(lastRow, 1))
        
        If fromDate > MaxDate Or toDate < minDate Then
            ' Skip this period - out of range
            GoTo NextPeriod
        End If
        
        ' Add period title
        With wsPeriodAnalysis.Cells(currentRow, correlationLeft)
            .value = periodTitle & " (" & Format(fromDate, "yyyy-mm-dd") & " to " & Format(toDate, "yyyy-mm-dd") & ")"
            .Font.Bold = True
            .Font.Size = 12
        End With
        
        ' Calculate correlation matrix for this period
        ReDim correlationMatrix(1 To lastColumn - 1, 1 To lastColumn - 1)
        Call CalculateCorrelationMatrix(correlationMatrix, allData, fromDate, toDate, lastRow, lastColumn, periodTitle)
        
        ' Format correlation matrix
        Call FormatCorrelationMatrix(wsPeriodAnalysis, correlationMatrix, allData, portfolioData, _
            currentRow + 1, periodTitle, High_Threshold, Low_Threshold, lastColumn)
        
        ' Create individual strategies equity curve chart
        chartLeft = wsPeriodAnalysis.Cells(currentRow + 1, strategyChartLeft).left
        chartTop = wsPeriodAnalysis.Cells(currentRow + 1, strategyChartLeft).top
        chartHeight = matrixHeight * 10  ' Approximate height to match matrix
        
        Call CreateStrategiesEquityCurveChart(wsPeriodAnalysis, allData, fromDate, toDate, lastRow, lastColumn, _
            chartLeft, chartTop, chartWidth, chartHeight, periodTitle)
        
        ' Create total portfolio equity curve chart
        chartLeft = wsPeriodAnalysis.Cells(currentRow + 1, totalChartLeft).left
        chartTop = wsPeriodAnalysis.Cells(currentRow + 1, totalChartLeft).top
        
        Call CreateTotalEquityCurveChart(wsPeriodAnalysis, allData, fromDate, toDate, lastRow, lastColumn, _
            chartLeft, chartTop, chartWidth, chartHeight, periodTitle)
        
        ' Update row for next period
        currentRow = currentRow + matrixHeight + 5  ' Add some spacing between periods
        
NextPeriod:
        ' Label for skipping to next iteration
    Next p
    
    ' Final formatting
    With ThisWorkbook.Windows(1)
        .Zoom = 70
    End With
    
    Call OrderVisibleTabsBasedOnList
    wsPeriodAnalysis.Activate

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub CreateStrategiesEquityCurveChart(ByRef ws As Worksheet, _
                                 ByRef data As Variant, _
                                 ByVal startdate As Date, _
                                 ByVal endDate As Date, _
                                 ByVal lastRow As Long, _
                                 ByVal lastColumn As Long, _
                                 ByVal chartLeft As Double, _
                                 ByVal chartTop As Double, _
                                 ByVal chartWidth As Double, _
                                 ByVal chartHeight As Double, _
                                 ByVal title As String)
    Dim cht As chart
    Dim i As Long, j As Long, K As Long
    Dim seriesCount As Long
    Dim strategyName As String
    Dim firstDataRow As Long, lastDataRow As Long
    Dim colorIndex As Long
    
    ' Use the specified chart width
    chartWidth = 1500
    
    ' Find date range within data
    firstDataRow = -1
    lastDataRow = -1
    
    For i = 2 To lastRow
        If CDate(data(i, 1)) >= startdate And firstDataRow = -1 Then
            firstDataRow = i
        End If
        
        If CDate(data(i, 1)) <= endDate Then
            lastDataRow = i
        Else
            If lastDataRow <> -1 Then
                Exit For
            End If
        End If
    Next i
    
    ' Check if we found valid data for this period
    If firstDataRow = -1 Or lastDataRow = -1 Or firstDataRow > lastDataRow Then
        Exit Sub
    End If
    
    ' Calculate row count for the period
    Dim rowCount As Long
    rowCount = lastDataRow - firstDataRow + 1
    
    ' Find the last used column for data - start after that
    Dim dataStartColumn As Long
    dataStartColumn = lastColumn + 80 ' Starting position for data areas
    
    ' Find the last used row in the data area
    Dim lastUsedRow As Long
    On Error Resume Next
    lastUsedRow = ws.Cells(ws.rows.count, dataStartColumn).End(xlUp).row
    If Err.Number <> 0 Or lastUsedRow < 3 Then lastUsedRow = 2 ' Header row
    On Error GoTo 0
    
    ' Add a blank row after the last used data
    If lastUsedRow > 2 Then lastUsedRow = lastUsedRow + 1
    
    ' Add header for this period's data
    ws.Cells(lastUsedRow + 1, dataStartColumn).value = "Data for " & title & " - Individual Strategies"
    ws.Cells(lastUsedRow + 1, dataStartColumn).Font.Bold = True
    
    ' Add column headers
    ws.Cells(lastUsedRow + 2, dataStartColumn).value = "Date"
    For j = 2 To lastColumn
        ws.Cells(lastUsedRow + 2, dataStartColumn + j - 1).value = data(1, j)
    Next j
    
    ' Remember the start of this dataset for chart references
    Dim dataStartRow As Long
    dataStartRow = lastUsedRow + 3 ' First data row
    
    ' Initialize starting values for each strategy to zero
    Dim cumulativePnL() As Double
    ReDim cumulativePnL(2 To lastColumn)
    For j = 2 To lastColumn
        cumulativePnL(j) = 0
    Next j
    
    ' Calculate cumulative values
    For i = 1 To rowCount
        ' Copy date
        ws.Cells(dataStartRow + i - 1, dataStartColumn).value = data(firstDataRow + i - 1, 1)
        
        ' Calculate cumulative P&L for each strategy
        For j = 2 To lastColumn
            Dim currValue As Double
            
            ' Get the value for the current day
            currValue = CDbl(data(firstDataRow + i - 1, j))
            
            ' Add to the cumulative total
            cumulativePnL(j) = cumulativePnL(j) + currValue
            
            ' Store the cumulative value
            ws.Cells(dataStartRow + i - 1, dataStartColumn + j - 1).value = cumulativePnL(j)
        Next j
    Next i
    
    ' Define the chart data as a named range for easier reference
    Dim dataRangeName As String
    dataRangeName = "StrategyData_" & Application.WorksheetFunction.Substitute(title, " ", "_")
    On Error Resume Next
    If ThisWorkbook.Names(dataRangeName).name <> "" Then ThisWorkbook.Names(dataRangeName).Delete
    ThisWorkbook.Names.Add name:=dataRangeName, _
        RefersTo:=ws.Range(ws.Cells(dataStartRow, dataStartColumn), _
                   ws.Cells(dataStartRow + rowCount - 1, dataStartColumn + lastColumn - 1))
    On Error GoTo 0
    
    ' Create chart
    Set cht = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight).chart
    
    ' Set chart type and title
    cht.ChartType = xlLine
    cht.HasTitle = True
    cht.chartTitle.text = "Individual Strategy Equity: " & title
    
    ' Set date axis format
    cht.Axes(xlCategory).TickLabels.NumberFormat = "yyyy-mm-dd"
    cht.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
    
    ' Add series for each strategy
    For j = 2 To lastColumn
        strategyName = data(1, j)
        
        ' Add series
        cht.SeriesCollection.NewSeries
        With cht.SeriesCollection(j - 1)
            .name = strategyName
            .XValues = ws.Range(ws.Cells(dataStartRow, dataStartColumn), _
                     ws.Cells(dataStartRow + rowCount - 1, dataStartColumn))
            .values = ws.Range(ws.Cells(dataStartRow, dataStartColumn + j - 1), _
                    ws.Cells(dataStartRow + rowCount - 1, dataStartColumn + j - 1))
            
            ' Assign colors cyclically to make them distinct
            colorIndex = ((j - 2) Mod 10) + 1
            Select Case colorIndex
                Case 1: .Format.line.ForeColor.RGB = RGB(0, 0, 255)    ' Blue
                Case 2: .Format.line.ForeColor.RGB = RGB(255, 0, 0)    ' Red
                Case 3: .Format.line.ForeColor.RGB = RGB(0, 128, 0)    ' Green
                Case 4: .Format.line.ForeColor.RGB = RGB(255, 165, 0)  ' Orange
                Case 5: .Format.line.ForeColor.RGB = RGB(128, 0, 128)  ' Purple
                Case 6: .Format.line.ForeColor.RGB = RGB(0, 128, 128)  ' Teal
                Case 7: .Format.line.ForeColor.RGB = RGB(255, 0, 255)  ' Magenta
                Case 8: .Format.line.ForeColor.RGB = RGB(128, 128, 0)  ' Olive
                Case 9: .Format.line.ForeColor.RGB = RGB(0, 0, 128)    ' Navy
                Case 10: .Format.line.ForeColor.RGB = RGB(128, 0, 0)   ' Maroon
            End Select
        End With
    Next j
    
    ' Format chart
    With cht
        .HasLegend = True
        .Legend.position = xlLegendPositionRight
        
        ' Format axes
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.text = "Date"
        
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.text = "Equity"
        .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
        
        ' Set common format
        .ChartArea.Format.line.Visible = msoTrue
        .ChartArea.Format.line.ForeColor.RGB = RGB(128, 128, 128)
        .PlotArea.Format.line.Visible = msoTrue
        .PlotArea.Format.line.ForeColor.RGB = RGB(128, 128, 128)
    End With
    
    ' Return the last used row for potential use by total curve
    lastUsedRow = dataStartRow + rowCount - 1
End Sub

Private Sub CreateTotalEquityCurveChart(ByRef ws As Worksheet, _
                                 ByRef data As Variant, _
                                 ByVal startdate As Date, _
                                 ByVal endDate As Date, _
                                 ByVal lastRow As Long, _
                                 ByVal lastColumn As Long, _
                                 ByVal chartLeft As Double, _
                                 ByVal chartTop As Double, _
                                 ByVal chartWidth As Double, _
                                 ByVal chartHeight As Double, _
                                 ByVal title As String)
    Dim cht As chart
    Dim i As Long, j As Long, K As Long
    Dim strategyName As String
    Dim firstDataRow As Long, lastDataRow As Long
    Dim totalEquity As Double
    Dim colorIndex As Long
    

    ' Find date range within data
    firstDataRow = -1
    lastDataRow = -1
    
    For i = 2 To lastRow
        If CDate(data(i, 1)) >= startdate And firstDataRow = -1 Then
            firstDataRow = i
        End If
        
        If CDate(data(i, 1)) <= endDate Then
            lastDataRow = i
        Else
            If lastDataRow <> -1 Then
                Exit For
            End If
        End If
    Next i
    
    ' Check if we found valid data for this period
    If firstDataRow = -1 Or lastDataRow = -1 Or firstDataRow > lastDataRow Then
        Exit Sub
    End If
    
    ' Calculate row count for the period
    Dim rowCount As Long
    rowCount = lastDataRow - firstDataRow + 1
    
    ' Find the data area for total portfolio
    Dim totalStartColumn As Long
    totalStartColumn = lastColumn + 80 + lastColumn  ' Position after strategy data area
    
    ' Find the last used row in the total data area
    Dim lastUsedRow As Long
    On Error Resume Next
    lastUsedRow = ws.Cells(ws.rows.count, totalStartColumn).End(xlUp).row
    If Err.Number <> 0 Or lastUsedRow < 3 Then lastUsedRow = 2 ' Header row
    On Error GoTo 0
    
    ' Add a blank row after the last used data
    If lastUsedRow > 2 Then lastUsedRow = lastUsedRow + 1
    
    ' Add header for this period's data
    ws.Cells(lastUsedRow + 1, totalStartColumn).value = "Data for " & title & " - Total Portfolio"
    ws.Cells(lastUsedRow + 1, totalStartColumn).Font.Bold = True
    
    ' Add column headers
    ws.Cells(lastUsedRow + 2, totalStartColumn).value = "Date"
    ws.Cells(lastUsedRow + 2, totalStartColumn + 1).value = "Total Portfolio"
    
    ' Remember the start of this dataset for chart references
    Dim dataStartRow As Long
    dataStartRow = lastUsedRow + 3 ' First data row
    
    ' Initialize cumulative total
    Dim cumulativeTotal As Double
    cumulativeTotal = 0
    
    ' Calculate and store cumulative P&L for total portfolio
    For i = 1 To rowCount
        ' Copy date
        ws.Cells(dataStartRow + i - 1, totalStartColumn).value = data(firstDataRow + i - 1, 1)
        
        ' Calculate today's total for all strategies
        totalEquity = 0
        For j = 2 To lastColumn
            totalEquity = totalEquity + CDbl(data(firstDataRow + i - 1, j))
        Next j
        
        ' Add to cumulative total
        cumulativeTotal = cumulativeTotal + totalEquity
        
        ' Store the cumulative total
        ws.Cells(dataStartRow + i - 1, totalStartColumn + 1).value = cumulativeTotal
    Next i
    
    ' Define the chart data as a named range for easier reference
    Dim dataRangeName As String
    dataRangeName = "TotalData_" & Application.WorksheetFunction.Substitute(title, " ", "_")
    On Error Resume Next
    If ThisWorkbook.Names(dataRangeName).name <> "" Then ThisWorkbook.Names(dataRangeName).Delete
    ThisWorkbook.Names.Add name:=dataRangeName, _
        RefersTo:=ws.Range(ws.Cells(dataStartRow, totalStartColumn), _
                   ws.Cells(dataStartRow + rowCount - 1, totalStartColumn + 1))
    On Error GoTo 0
    
    ' Create chart
    Set cht = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight).chart
    
    ' Set chart type and title
    cht.ChartType = xlLine
    cht.HasTitle = True
    cht.chartTitle.text = "Cumulative Portfolio Return: " & title
    
    ' Set date axis format
    cht.Axes(xlCategory).TickLabels.NumberFormat = "yyyy-mm-dd"
    cht.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
    
    ' Add series for total portfolio
    cht.SeriesCollection.NewSeries
    With cht.SeriesCollection(1)
        .name = "Total Portfolio"
        .XValues = ws.Range(ws.Cells(dataStartRow, totalStartColumn), _
                  ws.Cells(dataStartRow + rowCount - 1, totalStartColumn))
        .values = ws.Range(ws.Cells(dataStartRow, totalStartColumn + 1), _
                 ws.Cells(dataStartRow + rowCount - 1, totalStartColumn + 1))
        .Format.line.ForeColor.RGB = RGB(0, 112, 192)    ' Blue
        .Format.line.Weight = 3                          ' Thicker line
    End With
    
    ' Format chart
    With cht
        .HasLegend = True
        .Legend.position = xlLegendPositionBottom
        
        ' Format axes
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.text = "Date"
        
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.text = "Equity"
        .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
        ' Set common format
        .ChartArea.Format.line.Visible = msoTrue
        .ChartArea.Format.line.ForeColor.RGB = RGB(128, 128, 128)
        .PlotArea.Format.line.Visible = msoTrue
        .PlotArea.Format.line.ForeColor.RGB = RGB(128, 128, 128)
    End With
End Sub


Private Sub AddPeriodAnalysisNavigationButtons(ByRef wsSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    
    ' Delete Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 2).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeletePeriodAnalysis"
    End With
    
    ' Summary Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 5).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Summary Tab"
        .OnAction = "GoToSummary"
    End With
    
    ' Portfolio Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 8).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With
    
    ' Control Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 11).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With
    
    ' Strategies Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 14).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Strategies Tab"
        .OnAction = "GoToStrategies"
    End With
    
    ' Inputs Tab button
    Set btn = wsSheet.Buttons.Add(left:=wsSheet.Cells(2, 17).left, _
                                 top:=wsSheet.Cells(2, 2).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Inputs Tab"
        .OnAction = "GoToInputs"
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddPeriodAnalysisNavigationButtons: " & Err.Description
End Sub

' Delete button handler
Sub DeletePeriodAnalysis()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("CorrelationPeriodAnalysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Go to Summary tab
    Call GoToSummary
End Sub

Sub DeleteAllTempSheets()
    Dim ws As Worksheet
    Dim wsName As String
    Dim i As Long
    Dim failedDeletes As String
    Dim deleteCount As Long
    
    ' Disable alerts and events
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Loop through sheets backward (safer for deletion)
    deleteCount = 0
    failedDeletes = ""
    
    For i = ThisWorkbook.Sheets.count To 1 Step -1
        On Error Resume Next
        wsName = ThisWorkbook.Sheets(i).name
        
        ' Check if it's a temp sheet
        If left(wsName, 5) = "TempS" Or left(wsName, 5) = "TempT" Or left(wsName, 4) = "Temp" Then
            ' Try to unprotect first
            ThisWorkbook.Sheets(i).Unprotect
            
            ' Try to delete
            ThisWorkbook.Sheets(i).Delete
            
            If Err.Number <> 0 Then
                ' If deletion failed, record the error
                failedDeletes = failedDeletes & vbCrLf & "- " & wsName & " (" & Err.Description & ")"
                Err.Clear
            Else
                deleteCount = deleteCount + 1
            End If
        End If
        On Error GoTo 0
    Next i
    
    ' Restore settings
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Report results
    If deleteCount > 0 Then
        If Len(failedDeletes) > 0 Then
            MsgBox deleteCount & " temporary sheet(s) deleted successfully." & vbCrLf & vbCrLf & _
                   "The following sheets could not be deleted:" & failedDeletes, vbInformation, "Partial Success"
        Else
            MsgBox deleteCount & " temporary sheet(s) deleted successfully.", vbInformation, "Success"
        End If
    Else
        If Len(failedDeletes) > 0 Then
            MsgBox "Could not delete any temporary sheets." & vbCrLf & vbCrLf & _
                   "Sheets that could not be deleted:" & failedDeletes, vbExclamation, "Failed"
        Else
            MsgBox "No temporary sheets were found.", vbInformation, "No Action"
        End If
    End If
End Sub

' Alternative method - if above still fails, try this
Sub ForceDeleteTempSheets()
    Dim wsNames() As String
    Dim i As Long, j As Long
    Dim tempCount As Long
    
    ' First, count and collect names of temp sheets
    tempCount = 0
    For i = 1 To ThisWorkbook.Sheets.count
        If left(ThisWorkbook.Sheets(i).name, 5) = "TempS" Or _
           left(ThisWorkbook.Sheets(i).name, 5) = "TempT" Or _
           left(ThisWorkbook.Sheets(i).name, 4) = "Temp" Then
            tempCount = tempCount + 1
        End If
    Next i
    
    If tempCount = 0 Then
        MsgBox "No temporary sheets found.", vbInformation
        Exit Sub
    End If
    
    ' Collect names
    ReDim wsNames(1 To tempCount)
    j = 1
    For i = 1 To ThisWorkbook.Sheets.count
        If left(ThisWorkbook.Sheets(i).name, 5) = "TempS" Or _
           left(ThisWorkbook.Sheets(i).name, 5) = "TempT" Or _
           left(ThisWorkbook.Sheets(i).name, 4) = "Temp" Then
            wsNames(j) = ThisWorkbook.Sheets(i).name
            j = j + 1
        End If
    Next i
    
    ' Now attempt to delete each sheet
    Application.DisplayAlerts = False
    For i = 1 To tempCount
        On Error Resume Next
        ' Try to fully break any links first
        With ThisWorkbook.Sheets(wsNames(i))
            .DisplayPageBreaks = False
            .EnableCalculation = False
            .EnableFormatConditionsCalculation = False
            .EnablePivotTable = False
            .Protect False
            .Visible = xlSheetVisible  ' Make sure it's visible first
        End With
        ThisWorkbook.Sheets(wsNames(i)).Delete
        If Err.Number <> 0 Then
            MsgBox "Could not delete: " & wsNames(i) & vbCrLf & Err.Description, vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
    
    MsgBox "Deletion process completed. Please check if any temporary sheets remain.", vbInformation
End Sub

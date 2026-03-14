Attribute VB_Name = "V_PositionCheck"
Sub CreateLatestPositionReport()
    ' Create a report of latest positions for Live strategies using Portfolio tab lookup
    
    On Error GoTo ErrorHandler
    
    ' Declare variables
    Dim wsPortfolio As Worksheet, wsLatestPositions As Worksheet
    Dim uniqueSymbols As Object
    Dim lastRow As Long
    Dim i As Long, row As Long
    Dim strategyName As String
    Dim symbol As Variant
    Dim sector As String, status As String
    Dim currentPosition As Double
    Dim positionStatus As String
    Dim liveStatus As String
    
    ' Performance optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Debug Output
    Debug.Print "Starting position report generation..."
    
    ' Initialize column constants
    Call InitializeColumnConstantsManually
    Debug.Print "Column constants initialized"
    
    ' Get liveStatus value for filtering
    liveStatus = GetNamedRangeValue("Port_Status")
    Debug.Print "Live status = " & liveStatus
    
    ' Create or clear the Latest Positions sheet
    Call Deletetab("Latest Positions")
    Set wsLatestPositions = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsLatestPositions.name = "Latest Positions"
    wsLatestPositions.Tab.Color = RGB(146, 208, 80) ' Green color
    
    ' Check required worksheets exist
    On Error Resume Next
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    
    If Err.Number <> 0 Then
        Debug.Print "Error finding Portfolio sheet: " & Err.Description
        On Error GoTo ErrorHandler
        Call Deletetab("Latest Positions")
        MsgBox "Portfolio sheet not found. Please run portfolio analysis first.", vbExclamation
        GoTo CleanExit
    End If
    On Error GoTo ErrorHandler
    
    ' Verify we have data
    If wsPortfolio Is Nothing Then
        Debug.Print "Portfolio sheet not found"
        Call Deletetab("Latest Positions")
        MsgBox "Portfolio sheet not found. Please run portfolio analysis first.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get last row of Portfolio data
    On Error Resume Next
    lastRow = wsPortfolio.Cells(wsPortfolio.rows.count, 1).End(xlUp).row
    If Err.Number <> 0 Then
        Debug.Print "Error getting lastRow: " & Err.Description
        On Error GoTo ErrorHandler
        Call Deletetab("Latest Positions")
        MsgBox "Error determining last row of portfolio data.", vbExclamation
        GoTo CleanExit
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "Portfolio last row = " & lastRow
    
    If lastRow < 2 Then
        Debug.Print "No portfolio data rows found"
        Call Deletetab("Latest Positions")
        MsgBox "No portfolio data found.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Find the maximum "Last Date On File" from Live strategies
    Dim maxLastDate As Date, tempDate As Date
    maxLastDate = 0 ' Initialize to earliest possible date
    
    For i = 2 To lastRow
        On Error Resume Next
        status = wsPortfolio.Cells(i, COL_PORT_STATUS).value
        tempDate = wsPortfolio.Cells(i, COL_PORT_LAST_DATE_ON_FILE).value
        
        If Err.Number = 0 And status = liveStatus And IsDate(tempDate) Then
            If tempDate > maxLastDate Then
                maxLastDate = tempDate
            End If
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    Next i
    
    ' If no valid date found, use today's date as fallback
    If maxLastDate = 0 Then
        maxLastDate = Date
        Debug.Print "No valid Last Date On File found, using today's date"
    Else
        Debug.Print "Maximum Last Date On File found: " & maxLastDate
    End If
    
    ' Initialize dictionary for unique symbols
    Set uniqueSymbols = CreateObject("Scripting.Dictionary")
    
    ' Create header for detailed position report
    With wsLatestPositions
        ' Title section
        .Cells(1, 1).value = "LATEST POSITIONS REPORT"
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Font.Bold = True
        
        .Cells(2, 1).value = "As of: " & Format(maxLastDate, "mmm dd, yyyy")
        .Cells(2, 1).Font.Bold = True
        .Cells(2, 1).Font.Italic = True
        
        ' Headers for detailed positions
        .Cells(4, 1).value = "Strategy Name"
        .Cells(4, 2).value = "Symbol"
        .Cells(4, 3).value = "Sector"
        .Cells(4, 4).value = "Status"
        .Cells(4, 5).value = "Last Date On File"
        .Cells(4, 6).value = "Position"
        .Cells(4, 7).value = "Position Status"
        
        ' Format headers
        .Range("A4:G4").Font.Bold = True
        .Range("A4:G4").Interior.Color = RGB(217, 225, 242) ' Light blue header
        .Range("A4:G4").Borders.LineStyle = xlContinuous
        .Range("A4:G4").Borders.Weight = xlThin
        
        ' Create summary table headers (to the right of the main table)
        ' Position it at column I (leaving column H as a spacer)
        .Cells(4, 9).value = "Symbol"
        .Cells(4, 10).value = "Net Position"
        .Cells(4, 11).value = "Status"
        
        ' Format summary headers
        .Range("I4:K4").Font.Bold = True
        .Range("I4:K4").Interior.Color = RGB(217, 225, 242)
        .Range("I4:K4").Borders.LineStyle = xlContinuous
        .Range("I4:K4").Borders.Weight = xlThin
        
        ' Add a title for the summary table
        .Cells(2, 9).value = "POSITION SUMMARY BY SYMBOL"
        .Cells(2, 9).Font.Size = 12
        .Cells(2, 9).Font.Bold = True
    End With
    
    ' Process data from Portfolio tab
    row = 5 ' Start row for data
    Dim lastDataRow As Long
    lastDataRow = 4 ' Start with header row
    
    ' Process each strategy in Portfolio tab
    For i = 2 To lastRow
        ' Get strategy details from Portfolio tab
        On Error Resume Next
        strategyName = wsPortfolio.Cells(i, COL_PORT_STRATEGY_NAME).value
        symbol = wsPortfolio.Cells(i, COL_PORT_SYMBOL).value
        sector = wsPortfolio.Cells(i, COL_PORT_SECTOR).value
        status = wsPortfolio.Cells(i, COL_PORT_STATUS).value
        currentPosition = wsPortfolio.Cells(i, COL_PORT_CURRENT_POSITION).value
        Dim lastDateOnFile As Date
        lastDateOnFile = wsPortfolio.Cells(i, COL_PORT_LAST_DATE_ON_FILE).value
        
        If Err.Number <> 0 Then
            Debug.Print "Error reading portfolio data at row " & i & ": " & Err.Description
            On Error GoTo ErrorHandler
            GoTo NextStrategy
        End If
        On Error GoTo ErrorHandler
        
        Debug.Print "Processing strategy: " & strategyName & " with position: " & currentPosition
        
        ' Only include Live strategies
        If status = liveStatus And strategyName <> "" Then
            ' Add to unique symbols collection
            Dim symbolKey As Variant
            symbolKey = CStr(symbol)
            If Not uniqueSymbols.Exists(symbolKey) Then uniqueSymbols.Add symbolKey, 0
            
            ' Determine position status
            If currentPosition > 0 Then
                positionStatus = "Long"
            ElseIf currentPosition < 0 Then
                positionStatus = "Short"
            Else
                positionStatus = "Flat"
            End If
            
            ' Write strategy to report
            With wsLatestPositions
                .Cells(row, 1).value = strategyName
                .Cells(row, 2).value = symbol
                .Cells(row, 3).value = sector
                .Cells(row, 4).value = status
                .Cells(row, 5).value = lastDateOnFile
                .Cells(row, 6).value = currentPosition
                .Cells(row, 7).value = positionStatus
                
                ' Update symbol position count
                uniqueSymbols(symbolKey) = uniqueSymbols(symbolKey) + currentPosition
            End With
            
            ' Apply conditional formatting based on position
            Select Case positionStatus
                Case "Long"
                    wsLatestPositions.Cells(row, 7).Interior.Color = RGB(198, 239, 206) ' Light green
                Case "Short"
                    wsLatestPositions.Cells(row, 7).Interior.Color = RGB(255, 199, 206) ' Light red
                Case "Flat"
                    wsLatestPositions.Cells(row, 7).Interior.Color = RGB(255, 235, 156) ' Light yellow
            End Select
            
            ' Add borders to this data row only
            wsLatestPositions.Range("A" & row & ":G" & row).Borders.LineStyle = xlContinuous
            wsLatestPositions.Range("A" & row & ":G" & row).Borders.Weight = xlThin
            
            row = row + 1
            lastDataRow = row - 1 ' Update last data row
        End If
NextStrategy:
    Next i
    
    ' Check if we found any live strategies
    If row = 5 Then
        wsLatestPositions.Cells(row, 1).value = "No Live strategies found in Portfolio tab."
        wsLatestPositions.Cells(row, 1).Font.Italic = True
        row = row + 1
        lastDataRow = row - 1
    End If
    
    ' Create Summary Table by Symbol
    Dim summaryRow As Long
    summaryRow = 5 ' Start at same row as main data
    Dim lastSummaryRow As Long
    lastSummaryRow = 4 ' Start with header row
    Dim netPosition As Double
    
    With wsLatestPositions
        ' Check if we have any symbols
        If uniqueSymbols.count = 0 Then
            .Cells(summaryRow, 9).value = "No symbols found with Live strategies."
            .Cells(summaryRow, 9).Font.Italic = True
            summaryRow = summaryRow + 1
            lastSummaryRow = summaryRow - 1
        Else
            ' Process each symbol
            
            For Each symbolKey In uniqueSymbols.keys
                netPosition = uniqueSymbols(symbolKey)
                
                ' Determine position status
                If netPosition > 0 Then
                    positionStatus = "Long"
                ElseIf netPosition < 0 Then
                    positionStatus = "Short"
                Else
                    positionStatus = "Flat"
                End If
                
                ' Write data
                .Cells(summaryRow, 9).value = symbolKey
                .Cells(summaryRow, 10).value = netPosition
                .Cells(summaryRow, 11).value = positionStatus
                
                ' Apply conditional formatting
                Select Case positionStatus
                    Case "Long"
                        .Cells(summaryRow, 11).Interior.Color = RGB(198, 239, 206) ' Light green
                    Case "Short"
                        .Cells(summaryRow, 11).Interior.Color = RGB(255, 199, 206) ' Light red
                    Case "Flat"
                        .Cells(summaryRow, 11).Interior.Color = RGB(255, 235, 156) ' Light yellow
                End Select
                
                ' Add borders to this data row only
                .Range(.Cells(summaryRow, 9), .Cells(summaryRow, 11)).Borders.LineStyle = xlContinuous
                .Range(.Cells(summaryRow, 9), .Cells(summaryRow, 11)).Borders.Weight = xlThin
                
                summaryRow = summaryRow + 1
                lastSummaryRow = summaryRow - 1
            Next symbolKey
        End If
    End With
    
    ' Format and finalize the sheet
    With wsLatestPositions
        ' Set column widths
        .Columns("A").ColumnWidth = 30 ' Strategy Name
        .Columns("B").ColumnWidth = 15 ' Symbol
        .Columns("C").ColumnWidth = 20 ' Sector
        .Columns("D").ColumnWidth = 15 ' Status
        .Columns("E").ColumnWidth = 15 ' Last Date On File
        .Columns("F").ColumnWidth = 12 ' Position
        .Columns("G").ColumnWidth = 15 ' Position status
        .Columns("H").ColumnWidth = 4 ' Spacer column
        .Columns("I").ColumnWidth = 15 ' Symbol in summary
        .Columns("J").ColumnWidth = 12 ' Net position in summary
        .Columns("K").ColumnWidth = 12 ' Status in summary
        
        ' Add navigation buttons
        AddLatestPositionsNavigationButtons wsLatestPositions
        
        ' Apply filters to both tables
        If lastDataRow > 4 Then ' Only if we have data
            .Range("A4:G4").AutoFilter
        End If
        
        If lastSummaryRow > 4 Then ' Only if we have data
            .Range("I4:K4").AutoFilter
        End If
        
        ' Format numbers and dates
        If lastDataRow >= 5 Then
            .Range(.Cells(5, 5), .Cells(lastDataRow, 5)).NumberFormat = "mm/dd/yyyy" ' Last Date On File
            .Range(.Cells(5, 6), .Cells(lastDataRow, 6)).NumberFormat = "0.00" ' Positions
        End If
        
        If lastSummaryRow >= 5 Then
            .Range(.Cells(5, 10), .Cells(lastSummaryRow, 10)).NumberFormat = "0.00" ' Summary net positions
            .Range(.Cells(5, 10), .Cells(lastSummaryRow, 10)).Font.Bold = True ' Bold the numbers
        End If
        
        ' Add thin bottom border to last rows
        If lastDataRow >= 5 Then
            .Range(.Cells(lastDataRow, 1), .Cells(lastDataRow, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(lastDataRow, 1), .Cells(lastDataRow, 7)).Borders(xlEdgeBottom).Weight = xlMedium
        End If
        
        If lastSummaryRow >= 5 Then
            .Range(.Cells(lastSummaryRow, 9), .Cells(lastSummaryRow, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(lastSummaryRow, 9), .Cells(lastSummaryRow, 11)).Borders(xlEdgeBottom).Weight = xlMedium
        End If
        
        ' Zoom to better fit
        .Parent.Windows(1).Zoom = 80
    End With
    
    ' Select first cell
    wsLatestPositions.Range("A1").Select
    
    Debug.Print "Report generation complete"
    MsgBox "Latest positions report for Live strategies created successfully!", vbInformation
    
CleanExit:
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "An error occurred:" & vbNewLine & _
            "Error " & Err.Number & ": " & Err.Description & vbNewLine & _
            "Line: " & Erl
    Debug.Print errMsg
    MsgBox errMsg, vbCritical
    Resume CleanExit
End Sub

Private Sub AddLatestPositionsNavigationButtons(ByRef wsSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    
    ' Delete Tab button
    Set btn = wsSheet.Buttons.Add(left:=200, _
                                 top:=wsSheet.Cells(2, 3).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Delete Tab"
        .OnAction = "DeleteLatestPositions"
    End With
    
    ' Portfolio Tab button
    Set btn = wsSheet.Buttons.Add(left:=350, _
                                 top:=wsSheet.Cells(2, 5).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Portfolio Tab"
        .OnAction = "GoToPortfolio"
    End With
    
    ' Control Tab button
    Set btn = wsSheet.Buttons.Add(left:=500, _
                                 top:=wsSheet.Cells(2, 5).top, _
                                 Width:=100, Height:=25)
    With btn
        .Caption = "Control Tab"
        .OnAction = "GoToControl"
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddLatestPositionsNavigationButtons: " & Err.Description
End Sub

Sub DeleteLatestPositions()
    ' Delete the Latest Positions tab
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Latest Positions").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Go to Control tab
    ThisWorkbook.Sheets("Control").Activate
End Sub


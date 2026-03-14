Attribute VB_Name = "R_Check_New_Strategies"
Option Explicit

Sub IdentifyNewStrategiesAndContractChanges()
    ' This macro compares strategy names and contract quantities in the portfolio sheet to those in a selected file
    ' It highlights strategies that don't exist in the selected file or don't have "Live" status
    ' It also checks for contract quantity changes and lists missing strategies
    ' Uses column constants already defined by InitializeColumnConstantsManually()
    
    Dim wbPortfolio As Workbook
    Dim wsPortfolio As Worksheet
    Dim wbReference As Workbook
    Dim wsStrategies As Worksheet
    Dim rgPortfolioStrategies As Range
    Dim cell As Range
    Dim strategyName As String
    Dim fileName As Variant
    Dim strategyFound As Boolean
    Dim foundRow As Long
    Dim lastRow As Long
    Dim portfolioLastRow As Long
    Dim liveStrategyCount As Long
    Dim missingStrategyCount As Long
    Dim nonLiveStrategyCount As Long
    Dim contractChangedCount As Long
    Dim totalStrategyCount As Long
    Dim addedStrategyCount As Long
    Dim msgResult As VbMsgBoxResult
    Dim i As Long
    Dim j As Long
    Dim referenceStrategyName As String
    Dim referenceStrategyStatus As String
    Dim portfolioStrategies As Collection
    Dim insertRow As Long
    
    ' Variables for contract comparison
    Dim portfolioContracts As Double
    Dim referenceContracts As Double
    Dim contractDifference As Double
    Dim contractChangeDetails As String
    
    ' The column constants are already defined by InitializeColumnConstantsManually
    ' We will reference them directly
    
     ' Column indices for strategy names
     Call InitializeColumnConstantsManually

    
    On Error GoTo ErrorHandler
    
    ' Set reference to current workbook and Portfolio worksheet
    Set wbPortfolio = ThisWorkbook
    
    ' Check if Portfolio sheet exists
    On Error Resume Next
    Set wsPortfolio = wbPortfolio.Worksheets("Portfolio")
    On Error GoTo ErrorHandler
    
    If wsPortfolio Is Nothing Then
        MsgBox "Error: The 'Portfolio' worksheet does not exist in this workbook.", vbExclamation, "Missing Worksheet"
        Exit Sub
    End If
    
    ' Verify that the column constants are defined
    If COL_PORT_STRATEGY_NAME <= 0 Then
        MsgBox "Error: COL_PORT_STRATEGY_NAME is not defined or has an invalid value." & vbCrLf & _
               "Please ensure InitializeColumnConstantsManually has been run.", vbExclamation, "Missing Column Definition"
        Exit Sub
    End If
    
    ' Check if contract columns are defined
    If COL_PORT_CONTRACTS <= 0 Or COL_STRAT_CONTRACTS <= 0 Then
        MsgBox "Warning: Contract columns are not properly defined." & vbCrLf & _
               "Contract comparison will be skipped." & vbCrLf & _
               "COL_PORT_CONTRACTS: " & COL_PORT_CONTRACTS & vbCrLf & _
               "COL_STRAT_CONTRACTS: " & COL_STRAT_CONTRACTS, vbExclamation, "Missing Contract Columns"
    End If


    ' Display file picker dialog to select the reference file
    fileName = Application.GetOpenFilename("Excel Files (*.xlsx; *.xlsm; *.xls),*.xlsx;*.xlsm;*.xls", , _
                "Select 'PortfolioTrackerConfig' file to compare the portfolio tab to", , False)
       
    ' User canceled the file dialog
    If fileName = False Then
        MsgBox "Operation cancelled. No file was selected.", vbInformation, "Cancelled"
        Exit Sub
    End If
    

    ' CLEAR PREVIOUS ANALYSIS RESULTS FIRST
    Call ClearPreviousAnalysisResults(wsPortfolio)
    
    ' Determine the last row with data in the strategy name column
    portfolioLastRow = wsPortfolio.Cells(wsPortfolio.rows.count, COL_PORT_STRATEGY_NAME).End(xlUp).row
    
    ' Check if there's any data to process
    If portfolioLastRow <= 1 Then
        MsgBox "No strategy data found in the Portfolio sheet.", vbInformation, "No Data"
        Exit Sub
    End If
    
    ' Set range for all strategy names in the portfolio (starting from row 2 to skip header)
    Set rgPortfolioStrategies = wsPortfolio.Range(wsPortfolio.Cells(2, COL_PORT_STRATEGY_NAME), _
                                                 wsPortfolio.Cells(portfolioLastRow, COL_PORT_STRATEGY_NAME))
    
    ' Confirm total strategies to check
    totalStrategyCount = rgPortfolioStrategies.Cells.count
      

    ' Try to open the selected file
    Set wbReference = Workbooks.Open(fileName, ReadOnly:=True)
    
    ' Check if the selected file has Strategies sheet
    On Error Resume Next
    Set wsStrategies = wbReference.Worksheets("Strategies")
    On Error GoTo ErrorHandler_CloseWorkbook
    
    If wsStrategies Is Nothing Then
        MsgBox "Error: The selected file does not have a 'Strategies' worksheet.", vbExclamation, "Missing Worksheet"
        wbReference.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Verify the columns have the expected headers
    If InStr(1, wsStrategies.Cells(1, COL_STRAT_STRATEGY_NAME).value, "Strategy", vbTextCompare) = 0 Then
        MsgBox "Warning: Column " & ConvertToLetter(COL_STRAT_STRATEGY_NAME) & " doesn't appear to contain strategy names." & vbCrLf & _
               "Expected header containing 'Strategy', found '" & wsStrategies.Cells(1, COL_STRAT_STRATEGY_NAME).value & "'", _
               vbExclamation, "Column Verification"
    End If
    
    If InStr(1, wsStrategies.Cells(1, COL_STRAT_STATUS).value, "Status", vbTextCompare) = 0 Then
        MsgBox "Warning: Column " & ConvertToLetter(COL_STRAT_STATUS) & " doesn't appear to contain status information." & vbCrLf & _
               "Expected header containing 'Status', found '" & wsStrategies.Cells(1, COL_STRAT_STATUS).value & "'", _
               vbExclamation, "Column Verification"
    End If
    
    ' Initialize counters
    liveStrategyCount = 0
    missingStrategyCount = 0
    nonLiveStrategyCount = 0
    contractChangedCount = 0
    addedStrategyCount = 0
    contractChangeDetails = ""
    
    ' Reset any existing highlighting in the portfolio
    rgPortfolioStrategies.Interior.colorIndex = xlNone
    ' Also reset contract column highlighting if it exists
    If COL_PORT_CONTRACTS > 0 Then
        wsPortfolio.Range(wsPortfolio.Cells(2, COL_PORT_CONTRACTS), _
                         wsPortfolio.Cells(portfolioLastRow, COL_PORT_CONTRACTS)).Interior.colorIndex = xlNone
    End If
    
    ' Create a collection of current portfolio strategies for quick lookup
    Set portfolioStrategies = New Collection
    For Each cell In rgPortfolioStrategies.Cells
        If Not IsEmpty(cell.value) Then
            On Error Resume Next
            portfolioStrategies.Add Trim(cell.value), Trim(cell.value)
            On Error GoTo ErrorHandler
        End If
    Next cell
    
    ' Compare each strategy in the portfolio with the reference file
    For Each cell In rgPortfolioStrategies.Cells
        If Not IsEmpty(cell.value) Then
            strategyName = Trim(cell.value)
            strategyFound = False
            
            ' Find matching strategy in the reference file
            lastRow = wsStrategies.Cells(wsStrategies.rows.count, COL_STRAT_STRATEGY_NAME).End(xlUp).row
            For foundRow = 2 To lastRow ' Start from row 2 (skip header)
                If Trim(wsStrategies.Cells(foundRow, COL_STRAT_STRATEGY_NAME).value) = strategyName Then
                    strategyFound = True
                    
                    ' Check if the strategy has "Live" status
                    If Trim(wsStrategies.Cells(foundRow, COL_STRAT_STATUS).value) = "Live" Then
                        liveStrategyCount = liveStrategyCount + 1
                        
                        ' NEW: Check for contract changes if both columns are defined
                        If COL_PORT_CONTRACTS > 0 And COL_STRAT_CONTRACTS > 0 Then
                            On Error Resume Next
                            portfolioContracts = CDbl(wsPortfolio.Cells(cell.row, COL_PORT_CONTRACTS).value)
                            referenceContracts = CDbl(wsStrategies.Cells(foundRow, COL_STRAT_CONTRACTS).value)
                            On Error GoTo ErrorHandler
                            
                            ' Compare contract quantities (allow for small rounding differences)
                            contractDifference = portfolioContracts - referenceContracts
                            If Abs(contractDifference) > 0.001 Then
                                ' Highlight contract change with orange color
                                wsPortfolio.Cells(cell.row, COL_PORT_CONTRACTS).Interior.Color = RGB(255, 165, 0) ' Orange
                                cell.Interior.Color = RGB(255, 215, 0) ' Gold (lighter orange for strategy name)
                                contractChangedCount = contractChangedCount + 1
                                
                                ' Add to details string
                                contractChangeDetails = contractChangeDetails & strategyName & ": " & _
                                    referenceContracts & " -> " & portfolioContracts & _
                                    " (" & Format(contractDifference, "+0.00;-0.00") & ")" & vbCrLf
                            End If
                        End If
                        
                    Else
                        ' Highlight with yellow color (non-Live strategy)
                        cell.Interior.Color = RGB(255, 255, 0) ' Yellow
                        nonLiveStrategyCount = nonLiveStrategyCount + 1
                    End If
                    
                    Exit For
                End If
            Next foundRow
            
            ' If strategy not found, highlight with red color
            If Not strategyFound Then
                cell.Interior.Color = RGB(255, 0, 0) ' Red
                missingStrategyCount = missingStrategyCount + 1
            End If
        End If
    Next cell
    
    ' Find LIVE strategies in reference file that are missing from portfolio
    ' Add a separator row first
    insertRow = portfolioLastRow + 2
    wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).value = "--- LIVE STRATEGIES MISSING FROM PORTFOLIO ---"
    wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).Font.Bold = True
    wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).Interior.Color = RGB(200, 200, 200) ' Light gray
    insertRow = insertRow + 1
    
    ' Check each strategy in the reference file
    lastRow = wsStrategies.Cells(wsStrategies.rows.count, COL_STRAT_STRATEGY_NAME).End(xlUp).row
    For i = 2 To lastRow ' Start from row 2 (skip header)
        referenceStrategyName = Trim(wsStrategies.Cells(i, COL_STRAT_STRATEGY_NAME).value)
        referenceStrategyStatus = Trim(wsStrategies.Cells(i, COL_STRAT_STATUS).value)
        
        If referenceStrategyName <> "" And referenceStrategyStatus = "Live" Then
            ' Check if this strategy exists in our portfolio
            On Error Resume Next
            Dim tempItem As String
            tempItem = portfolioStrategies(referenceStrategyName)
            On Error GoTo ErrorHandler
            
            ' If the Live strategy is not found in portfolio, add it to the list
            If tempItem = "" Then
                wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).value = referenceStrategyName
                wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).Interior.Color = RGB(144, 238, 144) ' Light green - Live but missing
                
                ' Add contract information if available
                If COL_STRAT_CONTRACTS > 0 Then
                    On Error Resume Next
                    referenceContracts = CDbl(wsStrategies.Cells(i, COL_STRAT_CONTRACTS).value)
                    On Error GoTo ErrorHandler
                    
                    wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME + 1).value = "Live (Missing) - Contracts: " & referenceContracts
                    If COL_PORT_CONTRACTS > 0 Then
                        wsPortfolio.Cells(insertRow, COL_PORT_CONTRACTS).value = referenceContracts
                        wsPortfolio.Cells(insertRow, COL_PORT_CONTRACTS).Interior.Color = RGB(144, 238, 144) ' Light green
                    End If
                Else
                    wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME + 1).value = "Live (Missing from Portfolio)"
                End If
                
                addedStrategyCount = addedStrategyCount + 1
                insertRow = insertRow + 1
            End If
            
            ' Reset the temp variable for next iteration
            tempItem = ""
        End If
    Next i
    
    ' Add a summary section with updated color legend
    If addedStrategyCount > 0 Or contractChangedCount > 0 Then
        insertRow = insertRow + 1
        wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).value = "--- COLOR LEGEND ---"
        wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).Font.Bold = True
        wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).Interior.Color = RGB(200, 200, 200) ' Light gray
        insertRow = insertRow + 1
        
        wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).value = "Green = Live strategies missing from your portfolio"
        wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).Interior.Color = RGB(144, 238, 144) ' Light green
        insertRow = insertRow + 1
        
        If contractChangedCount > 0 Then
            wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).value = "Gold/Orange = Contract quantity changed"
            wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).Interior.Color = RGB(255, 215, 0) ' Gold
            insertRow = insertRow + 1
        End If
        
        wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).value = "Yellow = Your strategies that are not Live in reference"
        wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).Interior.Color = RGB(255, 255, 0) ' Yellow
        insertRow = insertRow + 1
        
        wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).value = "Red = Your strategies not found in reference file"
        wsPortfolio.Cells(insertRow, COL_PORT_STRATEGY_NAME).Interior.Color = RGB(255, 0, 0) ' Red
    End If
    
    ' Close the reference workbook
    wbReference.Close SaveChanges:=False
    
    ' Display detailed summary results
    Dim summaryMessage As String
    summaryMessage = "Comparison complete:" & vbCrLf & _
           "PORTFOLIO ANALYSIS:" & vbCrLf & _
           "- " & liveStrategyCount & " strategies confirmed as Live" & vbCrLf & _
           "- " & nonLiveStrategyCount & " strategies found but not Live (highlighted in yellow)" & vbCrLf & _
           "- " & missingStrategyCount & " strategies not found in reference (highlighted in red)" & vbCrLf
    
    ' Add contract change information if any
    If contractChangedCount > 0 Then
        summaryMessage = summaryMessage & "- " & contractChangedCount & " strategies with contract changes (highlighted in gold/orange)" & vbCrLf
    End If
    
    summaryMessage = summaryMessage & vbCrLf & _
           "REFERENCE FILE ANALYSIS:" & vbCrLf & _
           "- " & addedStrategyCount & " LIVE strategies in reference file but missing from portfolio" & vbCrLf & vbCrLf & _
           "Total strategies in portfolio: " & totalStrategyCount & vbCrLf & _
           "Check below the portfolio data for LIVE strategies missing from your portfolio."
    
    ' Add contract change details if any
    If contractChangedCount > 0 And contractChangeDetails <> "" Then
        summaryMessage = summaryMessage & vbCrLf & vbCrLf & "CONTRACT CHANGES DETECTED:" & vbCrLf & contractChangeDetails
    End If
    
    MsgBox summaryMessage, vbInformation, "Strategy & Contract Comparison Results"
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    Exit Sub
    
ErrorHandler_CloseWorkbook:
    If Not wbReference Is Nothing Then
        wbReference.Close SaveChanges:=False
    End If
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    Exit Sub
End Sub

Function ConvertToLetter(ColumnNumber As Long) As String
    ' Convert a column number to an Excel column letter (e.g., 1 = A, 2 = B, etc.)
    Dim n As Long
    Dim c As Byte
    Dim s As String
    
    n = ColumnNumber
    
    Do While n > 0
        c = ((n - 1) Mod 26) + 65
        s = Chr(c) & s
        n = (n - 1) \ 26
    Loop
    
    ConvertToLetter = s
End Function

Sub ClearPreviousAnalysisResults(ws As Worksheet)
    ' Clear any previous analysis results added below the portfolio data
    Dim lastRow As Long
    Dim currentRow As Long
    Dim cellValue As String
    Dim foundSeparator As Boolean
    
    ' Find the actual end of portfolio data by looking for separator rows
    lastRow = ws.Cells(ws.rows.count, COL_PORT_STRATEGY_NAME).End(xlUp).row
    foundSeparator = False
    
    ' Look for the separator row that indicates start of analysis results
    For currentRow = 2 To lastRow
        cellValue = Trim(ws.Cells(currentRow, COL_PORT_STRATEGY_NAME).value)
        
        ' Check if this row contains analysis results (starts with "---" or contains specific text)
        If left(cellValue, 3) = "---" Or _
           InStr(1, cellValue, "LIVE STRATEGIES MISSING", vbTextCompare) > 0 Or _
           InStr(1, cellValue, "COLOR LEGEND", vbTextCompare) > 0 Or _
           InStr(1, cellValue, "Green =", vbTextCompare) > 0 Or _
           InStr(1, cellValue, "Yellow =", vbTextCompare) > 0 Or _
           InStr(1, cellValue, "Red =", vbTextCompare) > 0 Or _
           InStr(1, cellValue, "Gold/Orange =", vbTextCompare) > 0 Or _
           InStr(1, cellValue, "Live (Missing", vbTextCompare) > 0 Then
            
            foundSeparator = True
            ' Clear from this row to the end
            If currentRow <= lastRow Then
                ws.Range(ws.Cells(currentRow, 1), ws.Cells(lastRow, ws.Columns.count)).Clear
                ws.Range(ws.Cells(currentRow, 1), ws.Cells(lastRow, ws.Columns.count)).Interior.colorIndex = xlNone
            End If
            Exit For
        End If
    Next currentRow
    
    ' Alternative approach: if no separator found, look for rows with specific formatting
    If Not foundSeparator Then
        For currentRow = lastRow To 2 Step -1
            cellValue = Trim(ws.Cells(currentRow, COL_PORT_STRATEGY_NAME).value)
            
            ' Check if this looks like an added strategy (has specific background colors)
            If ws.Cells(currentRow, COL_PORT_STRATEGY_NAME).Interior.Color = RGB(144, 238, 144) Or _
               ws.Cells(currentRow, COL_PORT_STRATEGY_NAME).Interior.Color = RGB(200, 200, 200) Or _
               InStr(1, cellValue, "Live (Missing", vbTextCompare) > 0 Or _
               cellValue = "" Then
                ' This looks like an added row, clear it
                ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, ws.Columns.count)).Clear
                ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, ws.Columns.count)).Interior.colorIndex = xlNone
            Else
                ' This looks like real portfolio data, stop here
                Exit For
            End If
        Next currentRow
    End If

    ' —— NEW: Clear any interior colour in the two key columns ——
    Dim dataLastRow As Long
    dataLastRow = ws.Cells(ws.rows.count, COL_PORT_STRATEGY_NAME).End(xlUp).row
    If dataLastRow >= 2 Then
        ws.Range(ws.Cells(2, COL_PORT_STRATEGY_NAME), ws.Cells(dataLastRow, COL_PORT_STRATEGY_NAME)).Interior.colorIndex = xlNone
        ws.Range(ws.Cells(2, COL_PORT_CONTRACTS), ws.Cells(dataLastRow, COL_PORT_CONTRACTS)).Interior.colorIndex = xlNone
    End If
End Sub


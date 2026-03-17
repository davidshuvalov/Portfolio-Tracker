Attribute VB_Name = "O_Strategies_Tab"



Sub FormatStrategiesTab(wsStrategies As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long, j As Long

    
    Application.ScreenUpdating = False
    
    ' Get last row
    lastRow = wsStrategies.Cells(wsStrategies.rows.count, COL_STRAT_STRATEGY_NAME).End(xlUp).row
    
    ' Apply status dropdown to entire status column
    Dim statusRange As Range
    Set statusRange = wsStrategies.Range(wsStrategies.Cells(2, COL_STRAT_STATUS), _
                                       wsStrategies.Cells(lastRow, COL_STRAT_STATUS))
    Call CreateStatusDropdown(statusRange)
    
    ' Color coding based on status
    Dim status As String
    Dim portStatus As String
    portStatus = Trim(GetNamedRangeValue("Port_Status"))
    
    For i = 2 To lastRow
        status = Trim(wsStrategies.Cells(i, COL_STRAT_STATUS).value)
        
        ' Remove "Not Loaded - " prefix if present
        If InStr(1, status, "Not Loaded - ", vbTextCompare) > 0 Then
            status = Replace(status, "Not Loaded - ", "")
        End If
        
        ' Check for "Not Loaded" prefix first
        If InStr(1, wsStrategies.Cells(i, COL_STRAT_STATUS).value, "Not Loaded - ", vbTextCompare) > 0 Then
            wsStrategies.rows(i).Interior.Color = RGB(198, 156, 109)  ' Brown
        Else
            
            statusOptions = GetNamedRangeValue("StatusOptions")
            If Len(Trim(statusOptions)) > 0 Then
                statusArray = Split(statusOptions, ",")
                For j = 0 To UBound(statusArray)
                    If StrComp(status, Trim(statusArray(j)), vbTextCompare) = 0 Then
                        With wsStrategies.rows(i).Interior
                          Select Case j
                            Case 0
                                .Color = RGB(230, 210, 200)  ' Pale Rose - warm neutral
                            Case 1
                                .Color = RGB(200, 220, 190)  ' Sage Green - soft natural green
                            Case 2
                                .Color = RGB(215, 190, 215)  ' Dusty Mauve - muted purple
                            Case 3
                                .Color = RGB(245, 220, 180)  ' Wheat - warm beige
                            Case 4
                                .Color = RGB(170, 210, 215)  ' Duck Egg Blue - soft cyan
                            Case 5
                                .Color = RGB(225, 198, 170)  ' Tan - neutral brown
                            Case 6
                                .Color = RGB(190, 210, 200)  ' Sea Foam - light green-grey
                            Case 7
                                .Color = RGB(210, 200, 220)  ' Lilac Grey - soft purple-grey
                            Case 8
                                .Color = RGB(220, 210, 190)  ' Sand - light warm grey
                            Case Else
                                .Color = RGB(240, 240, 240)  ' Light Grey - default fallback
                            End Select
                         End With
                    End If
                Next j
            End If
            
            ' Color code based on status
            With wsStrategies.rows(i).Interior
                Select Case status
                Case GetNamedRangeValue("Port_Status")
                        .Color = RGB(144, 238, 144)  ' Green
                Case GetNamedRangeValue("Pass_Status")
                        .Color = RGB(176, 196, 222)  ' Light Steel Blue - soft blue that's easy to read text on
                
                Case GetNamedRangeValue("BuyandHoldStatus")
                    .Color = RGB(220, 220, 220)  ' Grey
                Case "New"
                    
                    .Color = RGB(100, 245, 5)  ' Lavender
                Case "Failed"
                    .Color = RGB(200, 0, 0)  ' red
                
                Case "Duplicate Strategy"
                    .Color = RGB(255, 0, 0)  ' red
                
                
                End Select
            End With
            
            
            
        End If

    Next i
    
    If GetNamedRangeValue("Strat_Type") <> "" Then AddDropdown wsStrategies, COL_STRAT_TYPE, StringToArray(GetNamedRangeValue("Strat_Type")), 2, lastRow
    If GetNamedRangeValue("Strat_Horizon") <> "" Then AddDropdown wsStrategies, COL_STRAT_HORIZON, StringToArray(GetNamedRangeValue("Strat_Horizon")), 2, lastRow
    
    If GetNamedRangeValue("OtherInput") <> "" Then AddDropdown wsStrategies, COL_STRAT_OTHER, StringToArray(GetNamedRangeValue("OtherInput")), 2, lastRow
    

    
    
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred in FormatStrategiesTab: " & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub



Function RemoveTableGaps() As Boolean
    On Error GoTo ErrorHandler
    Dim wsStrategies As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tableRange As Range
    Dim i As Long
    Dim isRowEmpty As Boolean
    
    
    Set wsStrategies = ThisWorkbook.Sheets("Strategies")
    
    ' Find last used row and column
    lastRow = wsStrategies.Cells(wsStrategies.rows.count, "A").End(xlUp).row
    lastCol = wsStrategies.Cells(1, wsStrategies.Columns.count).End(xlToLeft).column
    
    ' Define table range (excluding header row)
    Set tableRange = wsStrategies.Range(wsStrategies.Cells(2, 1), wsStrategies.Cells(lastRow, lastCol))
    
    Application.ScreenUpdating = False
    
    For i = tableRange.rows.count To 1 Step -1
        isRowEmpty = WorksheetFunction.CountA(tableRange.rows(i)) = 0
        
        If isRowEmpty Then
            tableRange.rows(i).Delete Shift:=xlUp
        End If
    Next i
    
    Application.ScreenUpdating = True
    RemoveTableGaps = True
    Exit Function

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description
    RemoveTableGaps = False
End Function


Private Function GetStatusPriority(status As String, wsInputs As Worksheet) As Integer
    ' Check if status starts with "Not Loaded - "
    If InStr(1, status, "Not Loaded - ", vbTextCompare) > 0 Then
        GetStatusPriority = 10000  ' Very high number to ensure it's at the end
        Exit Function
    End If
    
    ' Remove "Not Loaded - " prefix if it exists
    Dim cleanStatus As String
    cleanStatus = Replace(Trim(status), "Not Loaded - ", "")
    cleanStatus = Trim(cleanStatus)
    
    ' Check if it matches Port_Status (highest priority)
    If StrComp(cleanStatus, Trim(GetNamedRangeValue("Port_Status")), vbTextCompare) = 0 Then
        GetStatusPriority = 2
        Exit Function
    End If
    
    ' Check if it matches Port_Status (highest priority)
    If StrComp(cleanStatus, Trim(GetNamedRangeValue("Pass_Status")), vbTextCompare) = 0 Then
        GetStatusPriority = 3
        Exit Function
    End If
    
     If StrComp(cleanStatus, "Duplicate Strategy", vbTextCompare) = 0 Then
        GetStatusPriority = 1
        Exit Function
    End If
    
    
    ' Check if it's Buy and Hold (second lowest priority)
    If StrComp(cleanStatus, Trim(GetNamedRangeValue("BuyandHoldStatus")), vbTextCompare) = 0 Then
        GetStatusPriority = 9999
        Exit Function
    End If
    
    
    ' Check if it's Buy and Hold (second lowest priority)
    If StrComp(cleanStatus, "New", vbTextCompare) = 0 Then
        GetStatusPriority = 9998
        Exit Function
    End If
    
    
    
    ' Check against StatusOptions
    Dim statusOptions As String
    Dim statusArray() As String
    Dim i As Long
    
    statusOptions = Range("StatusOptions").value
    If Len(Trim(statusOptions)) > 0 Then
        statusArray = Split(statusOptions, ",")
        For i = 0 To UBound(statusArray)
            If StrComp(cleanStatus, Trim(statusArray(i)), vbTextCompare) = 0 Then
                GetStatusPriority = i + 4  ' +2 because Port_Status is 2 and pass is 3
                Exit Function
            End If
        Next i
    End If
    
    ' If not found in any of the defined statuses, put after defined ones but before Buy and Hold
    GetStatusPriority = 999
End Function


Function ReorderStrategies() As Boolean
    On Error GoTo ErrorHandler
    
    Dim wsStrategies As Worksheet
    Dim wsInputs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim helperCol As Long
    
    ' Initialize worksheets
    Set wsStrategies = ThisWorkbook.Sheets("Strategies")
    Set wsInputs = ThisWorkbook.Sheets("Inputs")
    
    
    Call RemoveTableGaps
    
    ' Get last row and column
    lastRow = wsStrategies.Cells(wsStrategies.rows.count, 3).End(xlUp).row  ' Column C for Strategy Name
    If lastRow < 2 Then Exit Function
    
    lastCol = wsStrategies.Cells(1, wsStrategies.Columns.count).End(xlToLeft).column
    
    ' Add helper column at the end
    helperCol = lastCol + 1
    wsStrategies.Columns(helperCol).Insert
    wsStrategies.Cells(1, helperCol).value = "SortHelper"
    
    ' Populate helper column with status priority
    Dim statusCell As Range
    For Each statusCell In wsStrategies.Range(wsStrategies.Cells(2, 2), wsStrategies.Cells(lastRow, 2))
        wsStrategies.Cells(statusCell.row, helperCol).value = GetStatusPriority(statusCell.value, wsInputs)
    Next statusCell
    
    ' Basic sort using the helper column
    wsStrategies.Range("A2:Z" & lastRow).Sort _
        Key1:=wsStrategies.Cells(2, helperCol), _
        Order1:=xlAscending, _
        key2:=wsStrategies.Cells(2, 3), _
        Order2:=xlAscending, _
        Header:=xlNo
    
    ' Delete any SortHelper columns
    Call DeleteSortHelperColumns
    
    
    ' Renumber strategy numbers
    For i = 2 To lastRow
        wsStrategies.Cells(i, 1).value = i - 1
    Next i
    
    ReorderStrategies = True
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred in ReorderStrategies: " & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    
    ' Make sure to clean up SortHelper columns even if there was an error
    Call DeleteSortHelperColumns
    
    wsStrategies.Cells(1, 1).Select
    
    ReorderStrategies = False
End Function

Private Sub DeleteSortHelperColumns()
    On Error Resume Next
    
    Dim lastCol As Long
    lastCol = ThisWorkbook.Sheets("Strategies").Cells(1, Columns.count).End(xlToLeft).column
    
    ' Loop through columns backward
    Dim i As Long
    For i = lastCol To 1 Step -1
        If ThisWorkbook.Sheets("Strategies").Cells(1, i).value = "SortHelper" Then
            ThisWorkbook.Sheets("Strategies").Columns(i).Delete
        End If
    Next i
    
    On Error GoTo 0
End Sub



Sub OrganizeStrategiesTab()
    On Error GoTo ErrorHandler
    
    ' Turn off screen updating for performance
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ' Reorder strategies
    If Not ReorderStrategies() Then
        MsgBox "Error reordering strategies.", vbCritical
        GoTo CleanExit
    End If
    
CleanExit:
    ' Restore Excel settings
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred in OrganizeStrategiesTab: " & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub








Sub UpdateStrategyStatuses()
    On Error GoTo ErrorHandler
    
    Dim wsSummary As Worksheet
    Dim wsStrategies As Worksheet
    Dim stratNameCol As Range
    Dim stratNumCol As Range
    Dim stratStatusCol As Range
    Dim sumStatusCol As Range
    Dim sumStratNameCol As Range
    Dim sumStratNumCol As Range
    Dim lastRowStrat As Long
    Dim lastRowSum As Long
    Dim i As Long
    Dim foundCell As Range
    Dim updatedCount As Long
    Dim notFoundCount As Long
    Dim strategyList As String
    
    ' Initialize worksheets
    On Error Resume Next
    Set wsSummary = GetWorksheetSafe(ThisWorkbook, "Summary")
    Set wsStrategies = GetWorksheetSafe(ThisWorkbook, "Strategies")
    On Error GoTo ErrorHandler
    
    ' Verify both worksheets exist
    If wsSummary Is Nothing Then
        MsgBox "Summary worksheet not found!", vbCritical
        Exit Sub
    End If
    
    If wsStrategies Is Nothing Then
        MsgBox "Strategy worksheet not found!", vbCritical
        Exit Sub
    End If
     
     
     Call InitializeColumnConstantsManually
    
    ' Find last rows in both sheets
    lastRowStrat = wsStrategies.Cells(wsStrategies.rows.count, COL_STRAT_STRATEGY_NUMBER).End(xlUp).row
    lastRowSum = wsSummary.Cells(wsSummary.rows.count, COL_STRATEGY_NUMBER).End(xlUp).row
    
    ' Set up range references using the constants
    Set stratNameCol = wsStrategies.Range(wsStrategies.Cells(2, COL_STRAT_STRATEGY_NAME), _
                                   wsStrategies.Cells(lastRowStrat, COL_STRAT_STRATEGY_NAME))
    Set stratNumCol = wsStrategies.Range(wsStrategies.Cells(2, COL_STRAT_STRATEGY_NUMBER), _
                                  wsStrategies.Cells(lastRowStrat, COL_STRAT_STRATEGY_NUMBER))
    Set stratStatusCol = wsStrategies.Range(wsStrategies.Cells(2, COL_STRAT_STATUS), _
                                     wsStrategies.Cells(lastRowStrat, COL_STRAT_STATUS))
    
    Set sumStratNameCol = wsSummary.Range(wsSummary.Cells(2, COL_STRATEGY_NAME), _
                                    wsSummary.Cells(lastRowSum, COL_STRATEGY_NAME))
    Set sumStratNumCol = wsSummary.Range(wsSummary.Cells(2, COL_STRATEGY_NUMBER), _
                                   wsSummary.Cells(lastRowSum, COL_STRATEGY_NUMBER))
    Set sumStatusCol = wsSummary.Range(wsSummary.Cells(2, COL_STATUS), _
                                 wsSummary.Cells(lastRowSum, COL_STATUS))
    
    ' Initialize counters
    updatedCount = 0
    notFoundCount = 0
    strategyList = ""
    
    ' Begin update process
    Application.ScreenUpdating = False
    
    For i = 2 To lastRowSum
        If Not IsEmpty(wsSummary.Cells(i, COL_STRATEGY_NUMBER)) Then
            ' Look for matching strategy number in Strategy tab
            Set foundCell = stratNumCol.Find(What:=wsSummary.Cells(i, COL_STRATEGY_NUMBER).value, _
                                          LookIn:=xlValues, _
                                          LookAt:=xlWhole, _
                                          MatchCase:=False)
            
            If Not foundCell Is Nothing Then
                ' Get the corresponding status cell in Strategy tab
                Dim stratStatusCell As Range
                Set stratStatusCell = wsStrategies.Cells(foundCell.row, COL_STRAT_STATUS)
                
                ' Update status if it's different
                If stratStatusCell.value <> wsSummary.Cells(i, COL_STATUS).value Then
                    stratStatusCell.value = wsSummary.Cells(i, COL_STATUS).value
                    updatedCount = updatedCount + 1
                End If
            Else
                ' Add to list of not found strategies
                notFoundCount = notFoundCount + 1
                strategyList = strategyList & "Strategy #" & wsSummary.Cells(i, COL_STRATEGY_NUMBER).value & _
                             " (" & wsSummary.Cells(i, COL_STRATEGY_NAME).value & ")" & vbNewLine
            End If
        End If
    Next i
    

    If updatedCount > 0 Then
        Call OrganizeStrategiesTab

        Call FormatStrategiesTab(wsStrategies)

        ' Performance recalculation is intentionally NOT triggered here.
        ' Run "Recalculate Performance" separately to rebuild Summary metrics
        ' once you are happy with all status changes.
    End If


    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.EnableEvents = True

    ' Report results
    If updatedCount > 0 Or notFoundCount > 0 Then
        Dim msg As String
        msg = "Status changes saved:" & vbNewLine & _
              updatedCount & " strategy status(es) updated in Strategies sheet."

        If updatedCount > 0 Then
            msg = msg & vbNewLine & vbNewLine & _
                  "Next step: click 'Recalculate Performance' to update Summary metrics."
        End If

        If notFoundCount > 0 Then
            msg = msg & vbNewLine & vbNewLine & _
                  notFoundCount & " strategies not found in Strategies sheet:" & vbNewLine & _
                  strategyList
        End If

        MsgBox msg, vbInformation
    Else
        MsgBox "No updates needed. All statuses are current.", vbInformation
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number, vbCritical
    Exit Sub
End Sub


Sub ResetAndMoveSummaryTab()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    ' Create a temporary copy
    ThisWorkbook.Sheets("Summary").Copy After:=ThisWorkbook.Sheets("MW Folder Locations")
    
    ' Delete the original Summary sheet
    ThisWorkbook.Sheets("Summary").Delete
    
    ' Rename the copy back to Summary
    activeSheet.name = "Summary"
    
ExitSub:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
    Resume ExitSub
End Sub




Sub AddDropdown(ws As Worksheet, columnIndex As Long, listItems As Variant, startRow As Long, endRow As Long)
    Dim listString As String
    Dim validationRange As Range

    ' Build the list as a comma-separated string
    listString = Join(listItems, ",")

    ' Apply data validation to the specified range
    Set validationRange = ws.Range(ws.Cells(startRow, columnIndex), ws.Cells(endRow, columnIndex))
    With validationRange.Validation
        .Delete ' Clear existing validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listString
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub


Function StringToArray(ByVal InputString As String) As Variant
    Dim result() As String
    ' Split the string into an array using the Split function
    result = Split(InputString, ",")
    
    ' Trim whitespace from each element in the array
    Dim i As Integer
    For i = LBound(result) To UBound(result)
        result(i) = Trim(result(i))
    Next i
    
    ' Return the array
    StringToArray = result
End Function









Sub UpdateStrategyContracts()
    On Error GoTo ErrorHandler
    
    Dim wsPortfolio As Worksheet
    Dim wsStrategies As Worksheet
    Dim stratNameCol As Range
    Dim stratNumCol As Range
    Dim stratContractCol As Range
    Dim sumContractCol As Range
    Dim sumStratNameCol As Range
    Dim sumStratNumCol As Range
    Dim lastRowStrat As Long
    Dim lastRowSum As Long
    Dim i As Long
    Dim foundCell As Range
    Dim updatedCount As Long
    Dim notFoundCount As Long
    Dim strategyList As String
    
    ' Initialize worksheets
    On Error Resume Next
    Set wsPortfolio = ThisWorkbook.Worksheets("Portfolio")
    Set wsStrategies = ThisWorkbook.Worksheets("Strategies")
    On Error GoTo ErrorHandler
    
    ' Verify both worksheets exist
    If wsPortfolio Is Nothing Then
        MsgBox "Portfolio worksheet not found!", vbCritical
        Exit Sub
    End If
    
    If wsStrategies Is Nothing Then
        MsgBox "Strategy worksheet not found!", vbCritical
        Exit Sub
    End If
     
     
     Call InitializeColumnConstantsManually
    
    ' Find last rows in both sheets
    lastRowStrat = wsStrategies.Cells(wsStrategies.rows.count, COL_STRAT_STRATEGY_NUMBER).End(xlUp).row
    lastRowSum = wsPortfolio.Cells(wsPortfolio.rows.count, COL_PORT_STRATEGY_NUMBER).End(xlUp).row
    
    ' Set up range references using the constants
    Set stratNameCol = wsStrategies.Range(wsStrategies.Cells(2, COL_STRAT_STRATEGY_NAME), _
                                   wsStrategies.Cells(lastRowStrat, COL_STRAT_STRATEGY_NAME))
    Set stratNumCol = wsStrategies.Range(wsStrategies.Cells(2, COL_STRAT_STRATEGY_NUMBER), _
                                  wsStrategies.Cells(lastRowStrat, COL_STRAT_STRATEGY_NUMBER))
    Set stratContractCol = wsStrategies.Range(wsStrategies.Cells(2, COL_STRAT_CONTRACTS), _
                                     wsStrategies.Cells(lastRowStrat, COL_STRAT_CONTRACTS))
    
    Set sumStratNameCol = wsPortfolio.Range(wsPortfolio.Cells(2, COL_PORT_STRATEGY_NAME), _
                                    wsPortfolio.Cells(lastRowSum, COL_PORT_STRATEGY_NAME))
    Set sumStratNumCol = wsPortfolio.Range(wsPortfolio.Cells(2, COL_PORT_STRATEGY_NUMBER), _
                                   wsPortfolio.Cells(lastRowSum, COL_PORT_STRATEGY_NUMBER))
    Set sumContractCol = wsPortfolio.Range(wsPortfolio.Cells(2, COL_PORT_CONTRACTS), _
                                 wsPortfolio.Cells(lastRowSum, COL_PORT_CONTRACTS))
    
    ' Initialize counters
    updatedCount = 0
    notFoundCount = 0
    strategyList = ""
    
    ' Begin update process
    Application.ScreenUpdating = False
    
    For i = 2 To lastRowSum
        If Not IsEmpty(wsPortfolio.Cells(i, COL_PORT_STRATEGY_NUMBER)) Then
            ' Look for matching strategy number in Strategy tab
            Set foundCell = stratNumCol.Find(What:=wsPortfolio.Cells(i, COL_PORT_STRATEGY_NUMBER).value, _
                                          LookIn:=xlValues, _
                                          LookAt:=xlWhole, _
                                          MatchCase:=False)
            
            If Not foundCell Is Nothing Then
                ' Get the corresponding contract cell in Strategy tab
                Dim stratContractCell As Range
                Set stratContractCell = wsStrategies.Cells(foundCell.row, COL_STRAT_CONTRACTS)
                
                ' Update status if it's different
                If stratContractCell.value <> wsPortfolio.Cells(i, COL_PORT_CONTRACTS).value Then
                    stratContractCell.value = wsPortfolio.Cells(i, COL_PORT_CONTRACTS).value
                    updatedCount = updatedCount + 1
                End If
            Else
                ' Add to list of not found strategies
                notFoundCount = notFoundCount + 1
                strategyList = strategyList & "Strategy #" & wsPortfolio.Cells(i, COL_PORT_STRATEGY_NUMBER).value & _
                             " (" & wsPortfolio.Cells(i, COL_PORT_STRATEGY_NAME).value & ")" & vbNewLine
            End If
        End If
    Next i
    
    
    If updatedCount > 0 Then
       Call CreatePortfolioSummary
    End If
        
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.EnableEvents = True
    
    ' Report results
    If updatedCount > 0 Or notFoundCount > 0 Then
        Dim msg As String
        msg = "Update Complete:" & vbNewLine & _
              updatedCount & " contracts updated"
        
        If notFoundCount > 0 Then
            msg = msg & vbNewLine & vbNewLine & _
                  notFoundCount & " strategies not found:" & vbNewLine & _
                  strategyList
        End If
        
        MsgBox msg, vbInformation
    Else
        MsgBox "No updates needed. All contracts are current.", vbInformation
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number, vbCritical
    Exit Sub
End Sub



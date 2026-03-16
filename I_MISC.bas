Attribute VB_Name = "I_MISC"


Sub DeleteStrategyTab()
    Dim strategyTabName As String
    strategyTabName = "Strat - " & activeSheet.Cells(2, 2).value & " - Detail" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(strategyTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Sub DeleteStrategyCodeTab()
    Dim strategyTabName As String
    strategyTabName = "Strat - " & activeSheet.Cells(1, 11).value & " - Code" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(strategyTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub



Sub DeleteLeaveOneOut()
    Dim DeleteLeaveOneOutTabName As String
    DeleteLeaveOneOutTabName = "LeaveOneOut" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(DeleteLeaveOneOutTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub


Sub DeletePortfolioGraphs()
    Dim PortfolioGraphsTabName As String
    PortfolioGraphsTabName = "PortfolioGraphs" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(PortfolioGraphsTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Sub DeleteContractMarginTracking()
    Dim ContractMarginTrackingTabName As String
    ContractMarginTrackingTabName = "ContractMarginTracking" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(ContractMarginTrackingTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Sub DeleteCorrelations()
    Dim CorrelationsTabName As String
    CorrelationsTabName = "Correlations" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(CorrelationsTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Sub DeleteNegativeCorrelations()
    Dim NegativeCorrelationsTabName As String
    NegativeCorrelationsTabName = "NegativeCorrelations" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(NegativeCorrelationsTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub


Sub DeleteBackTestGraphs()
    Dim BackTestGraphsTabName As String
    BackTestGraphsTabName = "BackTestGraphs" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(BackTestGraphsTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Sub DeleteBacktestDetails()
    Dim BackTestStrategyTabName As String
    BackTestStrategyTabName = "BacktestDetails" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(BackTestStrategyTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub



Sub DeleteSizingGraphs()
    Dim SizingGraphsTabName As String
    SizingGraphsTabName = "SizingGraphs" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(SizingGraphsTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Sub DeleteDiversificator()
    Dim DiversificatorTabName As String
    DiversificatorTabName = "Diversificator" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(DiversificatorTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Sub DeleteSectorTypeGraphs()
    Dim SectorTypeGraphsTabName As String
    SectorTypeGraphsTabName = "SectorTypeGraphs" ' Adjust based on where strategy number is located
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(SectorTypeGraphsTabName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub




Sub Deletetab(tabname As String)
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(tabname).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub



Sub GoToSummary()
    Dim wsSummary As Worksheet
    ' Check if "Summary" sheet exists and has data in row 2
    
    Call InitializeColumnConstantsManually
    
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsSummary Is Nothing Then
        MsgBox "Error: 'Summary' sheet does not exist.", vbExclamation
        Exit Sub
    End If

   
    
    wsSummary.Activate ' Activate the Summary sheet
End Sub

Sub GoToPortfolio()
    Dim wsPortfolio As Worksheet
        ' Check if "Summary" sheet exists and has data in row 2


    Call InitializeColumnConstantsManually
    
    On Error Resume Next
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsPortfolio Is Nothing Then
        MsgBox "Error: 'Portfolio' sheet does not exist.", vbExclamation
        Exit Sub
    End If

    
    wsPortfolio.Activate ' Activate the Summary sheet
End Sub

Sub GoToControl()
    Dim wsControl As Worksheet
    Set wsControl = ThisWorkbook.Sheets("Control")
    wsControl.Activate ' Activate the Summary sheet
End Sub


Sub GoToStrategies()
    Dim wsStrategies As Worksheet
    Set wsStrategies = ThisWorkbook.Sheets("Strategies")
    wsStrategies.Activate ' Activate the Summary sheet
End Sub

Sub GoToInputs()
    Dim wsInputs As Worksheet
    Set wsInputs = ThisWorkbook.Sheets("Inputs")
    wsInputs.Activate ' Activate the Summary sheet
End Sub

Sub GoToMarkets()
    Call W_Markets.GoToMarkets
End Sub








Sub DeleteAllTabs()
   
   ' Show warning message and get user confirmation
    Dim response As VbMsgBoxResult

    
    response = MsgBox("WARNING: This will delete all data tabs and strategy tabs." & vbNewLine & _
                     "This action cannot be undone." & vbNewLine & vbNewLine & _
                     "Are you sure you want to continue?", _
                     vbExclamation + vbYesNo + vbDefaultButton2, _
                     "Confirm Delete All Tabs")
    
    ' Exit if user clicks No
    If response = vbNo Then
        Exit Sub
    End If
    
    ' Proceed with deletion if user clicks Yes
    'Application.DisplayAlerts = False
    
    ' Array of tabs to delete
    Dim tabsToDelete As Variant
    tabsToDelete = Array("Rules_Analysis", "Symbol_Analysis", "Sector_Analysis", "ContractMarginTracking", "NegativeCorrelations", "Correlations", _
                        "Portfolio", "PortfolioGraphs", "Summary", _
                        "SizingGraphs", "SectorTypeGraphs", "DailyM2MEquity", _
                        "ClosedTradePNL", "InMarketLong", "InMarketShort", _
                        "Long_Trades", "Short_Trades", _
                        "Walkforward Details", "PortfolioDailyM2M", "TotalPortfolioM2M", _
                        "PortInMarketShort", "PortInMarketLong", "PortfolioMC", "LeaveOneOut", "Diversificator", _
                        "MW Folder Locations", "StrategiesOld", "TotalBackTest", _
                        "BackTestGraphs", "BacktestDetails", "BackTestM2MEquity", "TrueRanges", "AverageTrueRange", _
                        "PortClosedTrade", "ContractMultiples", "CorrelationPeriodAnalysis", "DrawdownCorrelations", "TradePNL", "ATRCorrelations", "PNLCorrelations", "LatestPositionData", "Latest Positions", _
                        "Markets", "MarketCorrelations", "MarketVolatility")
    
    ' Delete standard tabs
    Dim tabname As Variant
    For Each tabname In tabsToDelete
        Call Deletetab(CStr(tabname))
    Next tabname
    
  
    ' Handle strategy tabs
    Dim ws As Worksheet
    Dim strategyCount As Long
    Dim strategyTabs() As String
    Dim i As Long
    
    ' Initialize count of strategy tabs
    strategyCount = 0
    
    ' Loop through all sheets to count strategy tabs
    For Each ws In ThisWorkbook.Sheets
        If left(ws.name, 8) = "Strat - " Then
            strategyCount = strategyCount + 1
            ReDim Preserve strategyTabs(1 To strategyCount)
            strategyTabs(strategyCount) = ws.name ' Store the strategy tab name
        End If
    Next ws
    
    ' Delete the strategy tabs
    If strategyCount > 0 Then
        For i = 1 To strategyCount
            Call Deletetab(CStr(strategyTabs(i)))
        Next i
    End If
    
    'Application.DisplayAlerts = True
    
    ' Show completion message
    MsgBox "All Data, Strategy, Summary and Portfolio Tabs Have Been Deleted", vbInformation, "Delete Complete"
End Sub


Sub DeleteAllStrategyTabs()
    Dim ws As Worksheet
    Dim strategyCount As Long
    Dim strategyTabs() As String
    Dim i As Long


    ' Initialize count of strategy tabs
    strategyCount = 0

    ' Loop through all sheets to count strategy tabs
    For Each ws In ThisWorkbook.Sheets
        If left(ws.name, 8) = "Strat - " Then
            strategyCount = strategyCount + 1
            ReDim Preserve strategyTabs(1 To strategyCount)
            strategyTabs(strategyCount) = ws.name ' Store the strategy tab name
        End If
    Next ws

    ' Delete the strategy tabs
    If strategyCount > 0 Then
        Application.DisplayAlerts = False ' Disable alerts for sheet deletion
        For i = 1 To strategyCount
            On Error Resume Next ' Ignore errors if the sheet does not exist
            ThisWorkbook.Sheets(strategyTabs(i)).Delete ' Delete the tab based on the stored name
            On Error GoTo 0 ' Resume error handling
        Next i
        Application.DisplayAlerts = True ' Re-enable alerts
        MsgBox strategyCount & " strategy tabs have been deleted.", vbInformation
    Else
        MsgBox "No strategy tabs found.", vbInformation
    End If
End Sub




Sub DeleteAllStrategyTabsNoPrompt()
    Dim ws As Worksheet
    Dim strategyCount As Long
    Dim strategyTabs() As String
    Dim i As Long


    ' Initialize count of strategy tabs
    strategyCount = 0

    ' Loop through all sheets to count strategy tabs
    For Each ws In ThisWorkbook.Sheets
        If left(ws.name, 8) = "Strat - " Then
             Call Deletetab(ws.name)
        End If
    Next ws
End Sub






Function ExtractNumericPart(sheetName As String) As Long
    Dim i As Long
    Dim numPart As String
    numPart = ""
    
    ' Extract numeric part
    For i = 1 To Len(sheetName)
        If IsNumeric(mid(sheetName, i, 1)) Then
            numPart = numPart & mid(sheetName, i, 1)
        ElseIf Len(numPart) > 0 Then
            Exit For
        End If
    Next i
    
    ' Return numeric part or 0 if none
    If numPart = "" Then
        ExtractNumericPart = 0
    Else
        ExtractNumericPart = CLng(numPart)
    End If
End Function

Function IsNumericPart(sheetName As String) As Boolean
    Dim i As Long
    IsNumericPart = False
    
    ' Check if the name contains any numbers
    For i = 1 To Len(sheetName)
        If IsNumeric(mid(sheetName, i, 1)) Then
            IsNumericPart = True
            Exit Function
        End If
    Next i
End Function

Sub QuickSort(arr As Variant, first As Long, last As Long)
    ' QuickSort algorithm to sort sheet names
    Dim low As Long, high As Long
    Dim temp As String
    Dim mid As String
    low = first
    high = last
    mid = arr((first + last) \ 2)
    Do While low <= high
        Do While arr(low) < mid
            low = low + 1
        Loop
        Do While arr(high) > mid
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop
    If first < high Then QuickSort arr, first, high
    If low < last Then QuickSort arr, low, last
End Sub





Function IsNonTradingDay(checkDate As Date) As Boolean
    ' Returns True if the date is a weekend or a U.S. public holiday when CME trading is closed.
    
    Dim yearInput As Integer
    Dim holiday As Variant
    Dim holidays As Collection

    ' Initialize
    yearInput = Year(checkDate)
    Set holidays = New Collection

    ' Add U.S. public holidays for the CME
    ' New Year's Day (observed)
    If Weekday(DateSerial(yearInput, 1, 1), vbMonday) = 6 Then
        holidays.Add DateSerial(yearInput, 1, 2) ' Observed on Monday
    Else
        holidays.Add DateSerial(yearInput, 1, 1)
    End If
    
    ' Martin Luther King Jr. Day (Third Monday in January)
    holidays.Add DateSerial(yearInput, 1, 1 + (15 - Weekday(DateSerial(yearInput, 1, 1), vbMonday)) Mod 7 + 14)
    
    ' Presidents' Day (Third Monday in February)
    holidays.Add DateSerial(yearInput, 2, 1 + (15 - Weekday(DateSerial(yearInput, 2, 1), vbMonday)) Mod 7 + 14)
    
    ' Good Friday (if you have a custom Easter function)
    On Error Resume Next
    Dim easterSunday As Date
    easterSunday = Application.WorksheetFunction.easterSunday(yearInput)
    If Not IsError(easterSunday) Then
        holidays.Add easterSunday - 2
    End If
    On Error GoTo 0

    ' Memorial Day (Last Monday in May)
    holidays.Add DateSerial(yearInput, 5, 31 - Weekday(DateSerial(yearInput, 5, 31), vbMonday))
    
    ' Independence Day (observed)
    If Weekday(DateSerial(yearInput, 7, 4), vbMonday) = 6 Then
        holidays.Add DateSerial(yearInput, 7, 5)
    ElseIf Weekday(DateSerial(yearInput, 7, 4), vbMonday) = 7 Then
        holidays.Add DateSerial(yearInput, 7, 3)
    Else
        holidays.Add DateSerial(yearInput, 7, 4)
    End If
    
    ' Labor Day (First Monday in September)
    holidays.Add DateSerial(yearInput, 9, 1 + (8 - Weekday(DateSerial(yearInput, 9, 1), vbMonday)) Mod 7)
    
    ' Thanksgiving (Fourth Thursday in November)
    holidays.Add DateSerial(yearInput, 11, 1 + (22 - Weekday(DateSerial(yearInput, 11, 1), vbMonday)) Mod 7 + 21)
    
    ' Christmas Day (observed)
    If Weekday(DateSerial(yearInput, 12, 25), vbMonday) = 6 Then
        holidays.Add DateSerial(yearInput, 12, 26)
    ElseIf Weekday(DateSerial(yearInput, 12, 25), vbMonday) = 7 Then
        holidays.Add DateSerial(yearInput, 12, 24)
    Else
        holidays.Add DateSerial(yearInput, 12, 25)
    End If

    ' Check if the date is a weekend
    If Weekday(checkDate, vbMonday) = 6 Or Weekday(checkDate, vbMonday) = 7 Then
        IsNonTradingDay = True
        Exit Function
    End If

    ' Check if the date is a public holiday
    For Each holiday In holidays
        If checkDate = holiday Then
            IsNonTradingDay = True
            Exit Function
        End If
    Next holiday

    ' Not a weekend or holiday
    IsNonTradingDay = False
End Function



Sub SortMonthsArray(arr As Variant)
    Dim i As Long, j As Long, temp As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                ' Swap elements
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub



Function GetDesiredTabOrder() As Variant
    Dim initialOrder As Variant
    Dim strategyTabs() As String
    Dim ws As Worksheet
    Dim i As Long, stratCount As Long
    
    ' Tab order grouped by workflow stage:
    '   Settings → Folder → Strategies → Portfolio → Backtest/WhatIf → Markets
    initialOrder = Array( _
        "Disclosure", "ReadMe", "Inputs", "Control", _
        "MW Folder Locations", "Status Changes", _
        "Strategies", "Backtest", _
        "Summary", "Portfolio", "PortfolioGraphs", "SectorTypeGraphs", "SizingGraphs", _
        "PortfolioMC", "Diversificator", "LeaveOneOut", _
        "Correlations", "NegativeCorrelations", "DrawdownCorrelations", "CorrelationPeriodAnalysis", _
        "PNLCorrelations", "ATRCorrelations", "ContractMarginTracking", "Latest Positions", _
        "TotalBackTest", "BackTestGraphs", "BacktestDetails", _
        "Markets", "MarketCorrelations", "MarketVolatility")
    
    ' Count and collect strategy tabs
    stratCount = 0
    For Each ws In ThisWorkbook.Sheets
        If left(ws.name, 8) = "Strat - " Then
            stratCount = stratCount + 1
            ReDim Preserve strategyTabs(1 To stratCount)
            strategyTabs(stratCount) = ws.name
        End If
    Next ws
    
    ' Sort strategy tabs numerically if any exist
    If stratCount > 0 Then
        Call NumericQuickSort(strategyTabs, 1, stratCount)
        
        ' Combine initial order with strategy tabs
        Dim finalOrder() As Variant
        ReDim finalOrder(0 To UBound(initialOrder) + stratCount)
        
        ' Copy initial order
        For i = 0 To UBound(initialOrder)
            finalOrder(i) = initialOrder(i)
        Next i
        
        ' Add sorted strategy tabs
        For i = 1 To stratCount
            finalOrder(UBound(initialOrder) + i) = strategyTabs(i)
        Next i
        
        GetDesiredTabOrder = finalOrder
    Else
        GetDesiredTabOrder = initialOrder
    End If
End Function

' Helper function to extract number from "Strategy - X" format
Private Function ExtractStrategyNumber(ByVal strategyName As String) As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim numStr As String
    
    ' Find position after first "- "
    startPos = InStr(1, strategyName, "- ") + 2
    
    ' Find position of " -" after the number
    endPos = InStr(startPos, strategyName, " -") - 1
    
    ' Extract the number between the delimiters
    numStr = mid(strategyName, startPos, endPos - startPos + 1)
    
    ' Convert to long
    ExtractStrategyNumber = CLng(numStr)
End Function

' QuickSort implementation for sorting strategy tabs numerically
Private Sub NumericQuickSort(ByRef arr() As String, ByVal low As Long, ByVal high As Long)
    Dim pivot As Long
    Dim tmp As String
    Dim i As Long
    Dim j As Long
    
    If low < high Then
        i = low
        j = high
        pivot = ExtractStrategyNumber(arr((low + high) \ 2))
        
        Do
            Do While ExtractStrategyNumber(arr(i)) < pivot
                i = i + 1
            Loop
            
            Do While ExtractStrategyNumber(arr(j)) > pivot
                j = j - 1
            Loop
            
            If i <= j Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        
        If low < j Then NumericQuickSort arr, low, j
        If i < high Then NumericQuickSort arr, i, high
    End If
End Sub

Function GetExistingTabs(desiredOrder As Variant) As Variant
    Dim existingTabs() As String
    Dim i As Long, j As Long
    Dim tabname As String
    
    j = 0
    For i = LBound(desiredOrder) To UBound(desiredOrder)
        tabname = desiredOrder(i)
        If SheetExists(tabname) Then
            j = j + 1
            ReDim Preserve existingTabs(1 To j)
            existingTabs(j) = tabname
        End If
    Next i
    
    GetExistingTabs = existingTabs
End Function

Function SheetExists(sheetName As Variant) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub OrderVisibleTabsBasedOnList()
    Dim desiredOrder As Variant
    Dim existingTabs As Variant
    Dim i As Long, j As Long, K As Long
    Dim tabname As String
    Dim ws As Worksheet
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    desiredOrder = GetDesiredTabOrder()
    existingTabs = GetExistingTabs(desiredOrder)
    
    For i = LBound(existingTabs) To UBound(existingTabs)
        tabname = existingTabs(i)
        Set ws = ThisWorkbook.Sheets(tabname)
        
        If ws.Visible = xlSheetVisible Then
            ' Find the desired position of the sheet
            For j = LBound(desiredOrder) To UBound(desiredOrder)
                If desiredOrder(j) = tabname Then
                    ' Move the sheet to its desired position
                    If j = LBound(desiredOrder) Then
                        ws.Move Before:=ThisWorkbook.Sheets(1)
                    Else
                        ' Find the next existing sheet in the desired order
                        For K = j - 1 To LBound(desiredOrder) Step -1
                            If SheetExists(desiredOrder(K)) Then
                                ws.Move After:=ThisWorkbook.Sheets(desiredOrder(K))
                                Exit For
                            End If
                        Next K
                        
                        ' If no existing sheet is found, move the sheet to the beginning
                        If K < LBound(desiredOrder) Then
                            ws.Move Before:=ThisWorkbook.Sheets(1)
                        End If
                    End If
                    Exit For
                End If
            Next j
        End If
    Next i
    
ExitSub:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Sub RemoveFilter(ws As Worksheet)
    On Error Resume Next
    ws.ShowAllData
    If Err.Number <> 0 Then
        ws.AutoFilterMode = False
    End If
    On Error GoTo 0
    
    
End Sub



Public Sub ResolveOOSDates( _
    ByVal storedOOSBegin As Variant, _
    ByVal storedOOSEnd As Variant, _
    ByVal useCutoff As Boolean, _
    ByVal cutoffDate As Variant, _
    ByRef outOOSBegin As Variant, _
    ByRef outOOSEnd As Variant)

    outOOSBegin = storedOOSBegin
    outOOSEnd = storedOOSEnd

    If Not useCutoff Or Not IsDate(cutoffDate) Then Exit Sub
    If Not IsDate(storedOOSBegin) Then Exit Sub  ' no OOS begin ? caller should skip calcs

    Dim beg As Date: beg = CDate(storedOOSBegin)
    Dim cut As Date: cut = CDate(cutoffDate)

    ' Case 1: cutoff before OOS begin ? clamp end to begin
    If cut < beg Then
        outOOSEnd = beg
        Exit Sub
    End If

    ' Case 2/3: cutoff >= begin
    If IsDate(storedOOSEnd) Then
        Dim en As Date: en = CDate(storedOOSEnd)
        If cut < en Then
            ' cutoff between begin and end ? cap end at cutoff
            outOOSEnd = cut
        Else
            ' cutoff > end ? keep stored end
            outOOSEnd = en
        End If
    Else
        ' stored OOS end is blank ? per your rule, keep it blank (unchanged)
        ' (do nothing)
    End If
End Sub



Private Function ReadNamedValue(ByVal nm As String) As Variant
    On Error GoTo EH
    Dim n As name
    For Each n In ThisWorkbook.Names
        If StrComp(n.name, nm, vbTextCompare) = 0 Then
            ReadNamedValue = n.RefersToRange.value
            Exit Function
        End If
    Next
EH:
    ReadNamedValue = Empty
End Function

Private Function UseCutoffSetting() As Boolean
    Dim v As Variant: v = ReadNamedValue(NAME_USE_CUTOFF)
    If IsError(v) Or IsEmpty(v) Then Exit Function
    UseCutoffSetting = (UCase$(Trim$(CStr(v))) = "YES")
End Function

Private Function CutoffDateSetting() As Variant
    Dim v As Variant: v = ReadNamedValue(NAME_CUTOFF_DATE)
    If IsDate(v) Then CutoffDateSetting = CDate(v)
End Function

' ===== The helper you asked about =====
' Returns last row whose date (in dateCol) <= cutoff if toggle is ON;
' otherwise returns the true last row. Assumes dates ascend in dateCol.
Public Function EndRowByCutoffSimple(ws As Worksheet, dateCol As Long) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, dateCol).End(xlUp).row
    If lastRow < 2 Then EndRowByCutoffSimple = lastRow: Exit Function

    Dim useCut As Boolean: useCut = UseCutoffSetting()
    Dim cut As Variant:    cut = CutoffDateSetting()

    If Not useCut Or Not IsDate(cut) Then
        EndRowByCutoffSimple = lastRow
        Exit Function
    End If

    Dim rng As Range, idx As Variant
    Set rng = ws.Range(ws.Cells(2, dateCol), ws.Cells(lastRow, dateCol)) ' data only
    idx = Application.Match(CDbl(CDate(cut)), rng, 1) ' last position <= cutoff

    If IsError(idx) Or IsEmpty(idx) Then
        EndRowByCutoffSimple = 1        ' none = cutoff (header only)
    Else
        EndRowByCutoffSimple = 1 + CLng(idx)  ' +1 because data starts at row 2
    End If
End Function


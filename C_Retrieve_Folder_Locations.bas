Attribute VB_Name = "C_Retrieve_Folder_Locations"
Sub RetrieveAllFolderData(Optional ByVal resetfoldertab As String = "No")
    Dim wsFolderLocations As Worksheet
    Dim wsStrategies  As Worksheet
    Dim dict As Object
    Dim duplicatesFound As Boolean
    Dim duplicateList As String
    Dim missingEquityDataList As String
    Dim equityDataMissing As Boolean
    Dim missingDetailsDataList As String
    Dim detailsDataMissing As Boolean
    Dim newfolderlist As String
    Dim newfolderflag As Boolean
    Dim strategyduplist As String
    Dim strategydupflag As Boolean
    Dim buyAndHoldFound As Boolean
    Dim folderCount As Long
    
    
    On Error GoTo ErrorHandler
    
    
    If resetfoldertab = "No" Then
        If Not IsLicenseValid() Then
            MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
            Exit Sub
        End If
    End If
        
    Call InitializeColumnConstantsManually
    
    ' Initialize variables
    Set dict = CreateObject("Scripting.Dictionary")
    duplicatesFound = False
    equityDataMissing = False
    detailsDataMissing = False
    newfolderflag = False
    strategydupflag = False
    buyAndHoldFound = False
    folderCount = 0
    
    duplicateList = ""
    missingEquityDataList = ""
    missingDetailsDataList = ""
    newfolderlist = ""
    strategyduplist = ""
    
    ' Create a new sheet for summary table if it doesn't exist
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("MW Folder Locations").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    
    ' Initialize worksheets
    On Error Resume Next
    Set wsStrategies = GetWorksheetSafe(ThisWorkbook, "Strategies")
    On Error GoTo ErrorHandler
    
    
    If wsStrategies Is Nothing Then
        MsgBox "Strategy worksheet not found!", vbCritical
        Exit Sub
    End If
    
    Call RemoveFilter(wsStrategies)
    
    If COL_STRAT_STRATEGY_NAME = 0 Or COL_STRAT_STATUS = 0 Then
        MsgBox "Column constants not properly initialized. Call InitializeColumnConstantsManually first.", vbCritical
        Exit Sub
    End If
        
    On Error Resume Next
    Set wsFolderLocations = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Control"))
    If Err.Number <> 0 Then
        MsgBox "Error creating MW Folder Locations sheet: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    wsFolderLocations.name = "MW Folder Locations"
    wsFolderLocations.Tab.Color = RGB(71, 211, 89)
    
      ' Set white background color for the entire worksheet
    wsFolderLocations.Cells.Interior.Color = RGB(255, 255, 255)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Clear previous data in the sheet (if any)
    wsFolderLocations.Columns("A:G").ClearContents

    ' Add headers to the table
    With wsFolderLocations
    .Cells(1, 1).value = "Folder Count"
    .Cells(1, 2).value = "Full Strategy Name"
    .Cells(1, 3).value = "Folder Name"
    .Cells(1, 4).value = "Full Folder Path"
    .Cells(1, 5).value = "Base Folder Path"
    .Cells(1, 6).value = "Status"
    
        ' Apply basic formatting to headers
        With .Range("A1:F1")
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(169, 208, 142) ' Light green fill
        End With
    End With

    ' Initialize starting row for data
    Dim startRow As Long
    startRow = 2
    
    Dim FoldersToImport As Variant
    FoldersToImport = Array("Folder1", "Folder2", "Folder3", _
                           "Folder4", "Folder5", "Folder6", "Folder7", "Folder8", "Folder9", "Folder10", "FolderBH")
    
    ' Process each folder
    Dim folder As Variant
     For Each folder In FoldersToImport
        
        ' Safely get folder path
        Dim folderPath As String
        Dim processFolderFlag As Boolean
        processFolderFlag = False
        
        On Error Resume Next
        folderPath = Range(folder).value
        If Err.Number <> 0 Or folderPath = "" Then
            Debug.Print "Warning: Named range '" & folder & "' not found or empty"
            Err.Clear
            processFolderFlag = False
        Else
            processFolderFlag = True
        End If
        On Error GoTo ErrorHandler
            
        ' Check if this is the Buy and Hold folder
        ' Only process if we got a valid folder path
        If processFolderFlag Then
            ' Check if this is the Buy and Hold folder
            If folder = "FolderBH" And GetNamedRangeValue("BuyandHoldStatus") <> "" Then
                Dim tempStartRow As Long
                tempStartRow = startRow

                ' Pass forcedStatus = BuyandHoldStatus so every strategy found in FolderBH
                ' is automatically classified as Buy & Hold — even when the folder contains
                ' multiple period files whose names differ from the Strategies tab entry.
                startRow = GetFolderData(folderPath, _
                                        wsFolderLocations, _
                                        startRow, _
                                        dict, _
                                        duplicatesFound, _
                                        duplicateList, _
                                        equityDataMissing, _
                                        missingEquityDataList, _
                                        detailsDataMissing, _
                                        missingDetailsDataList, _
                                        newfolderflag, _
                                        newfolderlist, _
                                        strategydupflag, _
                                        strategyduplist, _
                                        GetNamedRangeValue("BuyandHoldStatus"))
                
                ' Check if any Buy and Hold strategies were actually found
                If startRow > tempStartRow Then
                    buyAndHoldFound = True
                End If
            Else
                startRow = GetFolderData(folderPath, _
                                        wsFolderLocations, _
                                        startRow, _
                                        dict, _
                                        duplicatesFound, _
                                        duplicateList, _
                                        equityDataMissing, _
                                        missingEquityDataList, _
                                        detailsDataMissing, _
                                        missingDetailsDataList, _
                                        newfolderflag, _
                                        newfolderlist, _
                                        strategydupflag, _
                                        strategyduplist)
            End If
            
            ' Count total folders processed
            folderCount = startRow - 2
        End If
    Next folder


    ' After processing all folders, check for missing strategies
    If Not CheckMissingStrategies(wsFolderLocations) Then
        MsgBox "Error occurred while checking for missing strategies.", vbCritical
    End If
    
    
    Call OrganizeStrategiesTab

    Call FormatStrategiesTab(wsStrategies)
    
    
    
    
    ' Auto-fit the columns to content
    wsFolderLocations.Columns("A:F").AutoFit
    
    ' Apply AutoFilter to the header row
    wsFolderLocations.Range("A1:F1").AutoFilter
    
    
    If resetfoldertab = "No" Then
    
    
        Call OrderVisibleTabsBasedOnList
        
        ' Cleanup
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        
        
        
        
        ' Prepare message
        Dim message As String
        Dim errorCount As Integer
        errorCount = 0
        
        ' Start building the message
        If duplicatesFound Or equityDataMissing Or detailsDataMissing Or strategydupflag Or startRow <= 2 Then
            ThisWorkbook.Sheets("MW Folder Locations").Activate
            Range("A1").value = "Oh no ERRORS!"
            Range("A1").Select
            message = "PLEASE FIX THESE ISSUES BEFORE CONTINUING:" & vbNewLine & vbNewLine
            errorCount = 1  ' At least one error exists
        Else
            ThisWorkbook.Sheets("Control").Activate
            Range("A1").Select
            message = "SUCCESS: Folder names, paths, and strategy names retrieved successfully!" & vbNewLine & vbNewLine
            message = message & "Total folders processed: " & folderCount & vbNewLine & vbNewLine
        End If
        
        ' Add specific error messages if needed
        If duplicatesFound Then
            message = message & "ERROR #" & errorCount & ": Duplicate strategy names found:" & vbNewLine
            message = message & "----------------------------------------------" & vbNewLine
            message = message & duplicateList & vbNewLine & vbNewLine
            errorCount = errorCount + 1
        End If
        
        If equityDataMissing Then
            message = message & "ERROR #" & errorCount & ": Equity or Trade Files Missing:" & vbNewLine
            message = message & "----------------------------------------------" & vbNewLine
            message = message & missingEquityDataList & vbNewLine & vbNewLine
            errorCount = errorCount + 1
        End If
        
        If detailsDataMissing Then
            message = message & "ERROR #" & errorCount & ": Walkforward In-Out Periods Analysis Details file missing:" & vbNewLine
            message = message & "----------------------------------------------" & vbNewLine
            message = message & missingDetailsDataList & vbNewLine & vbNewLine
            errorCount = errorCount + 1
        End If
        
        If strategydupflag Then
            message = message & "ERROR #" & errorCount & ": Duplicate strategies in the Strategies tab:" & vbNewLine
            message = message & "----------------------------------------------" & vbNewLine
            message = message & strategyduplist & vbNewLine & vbNewLine
            errorCount = errorCount + 1
        End If
        
        If startRow = 2 Then
            message = message & "ERROR #" & errorCount & ": No Folder Locations Loaded" & vbNewLine & vbNewLine
            errorCount = errorCount + 1
        End If
        
        If newfolderflag Then
            message = message & "Friendly Informative Note: New Strategies found, update the status for:" & vbNewLine
            message = message & "----------------------------------------------" & vbNewLine
            If Len(newfolderlist) > 500 Then
                message = message & left(newfolderlist, 500) & vbNewLine & vbNewLine & "Plus a few more... You get the drift" & vbNewLine & vbNewLine
            Else
                message = message & newfolderlist & vbNewLine & vbNewLine
            End If
        End If
        
        ' Check for Buy and Hold status
        Dim buyHoldStatus As String
        buyHoldStatus = GetNamedRangeValue("BuyandHoldStatus")
        
        If buyHoldStatus = "" Then
            message = message & "WARNING: BuyandHoldStatus is empty. ATR calculations will not be performed." & vbNewLine & vbNewLine
        ElseIf Not buyAndHoldFound Then
            message = message & "WARNING: No Buy and Hold strategies were found. ATR calculations will not be performed. Please check that the status in the 'Strategies' tab is updated to buy and hold for the appropriate strategies." & vbNewLine & vbNewLine
        End If
        
        
        ' Determine message box type based on if there are errors
        Dim msgType As VbMsgBoxStyle
        If errorCount > 0 Then
            msgType = vbExclamation  ' Yellow exclamation mark for errors
        Else
            msgType = vbInformation  ' Blue information icon for success
        End If
        
        ' Display the message
        MsgBox message, msgType, "Folder Data Retrieval Report"
        
    End If
        
    Exit Sub

ErrorHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
    Debug.Print "Error occurred at line: " & Erl
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "An error occurred in RetrieveAllFolderData:" & vbNewLine & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description & vbNewLine & vbNewLine & _
           IIf(Erl <> 0, "Line: " & Erl & vbNewLine, "") & _
           "Please check:" & vbNewLine & _
           "1. All named ranges exist (Folder1-10, FolderBH, etc.)" & vbNewLine & _
           "2. Strategies worksheet exists" & vbNewLine & _
           "3. Column constants are initialized", vbCritical, "Error in Folder Data Retrieval"
End Sub



Function GetFolderData(baseFolderPath As String, ByRef wsFolderLocations As Worksheet, _
                      ByRef startRow As Long, ByRef dict As Object, _
                      ByRef duplicatesFound As Boolean, ByRef duplicateList As String, _
                      ByRef equityDataMissing As Boolean, ByRef missingEquityDataList As String, _
                      ByRef detailsDataMissing As Boolean, ByRef missingDetailsDataList As String, _
                      ByRef newfolderflag As Boolean, ByRef newfolderlist As String, _
                      ByRef strategydupflag As Boolean, ByRef strategyduplist As String, _
                      Optional forcedStatus As String = "") As Integer
' forcedStatus: when non-empty every strategy discovered in this folder is assigned
' that status in both folderLocations and the Strategies tab, overriding whatever
' the Strategies tab currently contains.  Pass GetNamedRangeValue("BuyandHoldStatus")
' when scanning FolderBH so that all BnH files are auto-classified regardless of
' how many period/sub-files exist for a single strategy.
    On Error GoTo ErrorHandler
    
    Dim foldername As String
    Dim FullFolderPath As String
    Dim folder As Object
    Dim FSO As Object
    Dim csvFile As String
    Dim csvFileTrades As String
    Dim FileNameOnly As String
    Dim CSV_Found As Boolean
    Dim CSV_Details_Found As Boolean
    Dim CSV_Found_Trades As Boolean
    Dim CSVGDetailsFile As String
    Dim equityFolderPath As String
    Dim status As String
     
    ' Input parameter validation
    If Trim(baseFolderPath) = "" Then
        'MsgBox "Error: Base folder path is empty", vbCritical
        GetFolderData = startRow
        Exit Function
    End If
    
    If wsFolderLocations Is Nothing Then
        MsgBox "Error: Folder locations worksheet not provided", vbCritical
        GetFolderData = startRow
        Exit Function
    End If
    
    If dict Is Nothing Then
        MsgBox "Error: Dictionary object not provided", vbCritical
        GetFolderData = startRow
        Exit Function
    End If
    
    ' Create FileSystemObject with error handling
    On Error Resume Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Or FSO Is Nothing Then
        MsgBox "Error creating FileSystemObject: " & Err.Description, vbCritical
        GetFolderData = startRow
        Exit Function
    End If
    On Error GoTo ErrorHandler

    ' Check if the folder path exists
    If Not FSO.FolderExists(baseFolderPath) Then
        MsgBox "The folder path does not exist: " & baseFolderPath, vbExclamation
        GetFolderData = startRow
        Exit Function
    End If
    
    ' Initialize Application properties for better performance
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    foldername = GetFolderName(baseFolderPath)

    ' Loop through each subfolder in the directory
    For Each folder In FSO.GetFolder(baseFolderPath).SubFolders
        ' Basic folder information
        
        FullFolderPath = folder.Path
        
        ' Reset flags for each folder
        CSV_Found = False
        CSV_Details_Found = False
        CSV_Found_Trades = False
        
        ' Check Walkforward Files folder
        equityFolderPath = BuildPath(FullFolderPath, "Walkforward Files")
        
        If FSO.FolderExists(equityFolderPath) Then
            ' Look for Details file
            CSVGDetailsFile = Dir(GetShortPath(BuildPath(equityFolderPath, "*Walkforward In-Out Periods Analysis Details.csv")))
            CSV_Details_Found = (CSVGDetailsFile <> "")
            
            ' Look for EquityData file
            
            
            On Error Resume Next
            csvFile = Dir(GetShortPath(BuildPath(equityFolderPath, "*EquityData.csv")))
            If Err.Number <> 0 Then
                Debug.Print "Error accessing directory: " & Err.Description
                csvFile = ""
                Err.Clear
            End If
            On Error GoTo ErrorHandler
                        
            ' Process each CSV file found
            Do While csvFile <> ""
                CSV_Found = True
                FileNameOnly = Replace(csvFile, "EquityData.csv", "")
                
                ' Get strategy status with error handling
                On Error Resume Next
                status = GetStrategyStatus(Trim(FileNameOnly), newfolderflag, newfolderlist, strategydupflag, strategyduplist)

                If Err.Number <> 0 Then
                    status = "Unknown"
                    Debug.Print "Error getting strategy status: " & Err.Description
                End If
                On Error GoTo ErrorHandler

                ' If a forced status is supplied (e.g. BuyandHold for FolderBH), override
                ' whatever status the Strategies tab currently stores and write it back.
                ' This ensures that all files discovered in the BnH folder — including
                ' multi-period files whose names don't exactly match the Strategies tab —
                ' are correctly classified without manual intervention.
                If Len(forcedStatus) > 0 And status <> forcedStatus Then
                    status = forcedStatus
                    Dim wsStratForce As Worksheet
                    Set wsStratForce = ThisWorkbook.Sheets("Strategies")
                    Dim forceRow As Long
                    For forceRow = 2 To wsStratForce.Cells(wsStratForce.rows.count, COL_STRAT_STRATEGY_NAME).End(xlUp).row
                        If StrComp(Trim(wsStratForce.Cells(forceRow, COL_STRAT_STRATEGY_NAME).value), _
                                   Trim(FileNameOnly), vbTextCompare) = 0 Then
                            wsStratForce.Cells(forceRow, COL_STRAT_STATUS).value = forcedStatus
                            Exit For
                        End If
                    Next forceRow
                    Set wsStratForce = Nothing
                End If
                
                ' Check for duplicates with proper string handling
                If dict.Exists(Trim(FileNameOnly)) Then
                    wsFolderLocations.rows(startRow).Interior.Color = RGB(255, 0, 0)
                    duplicatesFound = True
                    If InStr(1, duplicateList, FileNameOnly, vbTextCompare) = 0 Then
                        duplicateList = duplicateList & FileNameOnly & vbNewLine
                    End If
                Else
                    dict.Add Trim(FileNameOnly), True
                End If
                
                ' Write data to worksheet with error handling
                On Error Resume Next
                With wsFolderLocations
                    .Cells(startRow, 1).value = startRow - 1
                    .Cells(startRow, 2).value = IIf(CSV_Found, Trim(FileNameOnly), "No Equity or Trade Data Found")
                    .Cells(startRow, 3).value = foldername
                    .Cells(startRow, 4).value = FullFolderPath
                    .Cells(startRow, 5).value = baseFolderPath
                    .Cells(startRow, 6).value = status
                End With
                
                If Err.Number <> 0 Then
                    Debug.Print "Error writing to worksheet: " & Err.Description
                    ' Continue processing despite write error
                End If
                On Error GoTo ErrorHandler
                
                ' Get next CSV file
                csvFile = Dir
                startRow = startRow + 1
            Loop
            
            ' Handle missing equity data
            If Not CSV_Found Then
                On Error Resume Next
                With wsFolderLocations
                    .rows(startRow).Interior.Color = RGB(255, 255, 0)
                    .Cells(startRow, 1).value = startRow - 1
                    .Cells(startRow, 2).value = "No Equity or Trade Data Found"
                    .Cells(startRow, 3).value = foldername
                    .Cells(startRow, 4).value = FullFolderPath
                    .Cells(startRow, 5).value = baseFolderPath
                    .Cells(startRow, 6).value = "No Data"
                End With
                
                equityDataMissing = True
                missingEquityDataList = missingEquityDataList & foldername & " (Path: " & FullFolderPath & ")" & vbNewLine
                startRow = startRow + 1
            End If
            
            ' Handle missing details file
            If Not CSV_Details_Found Then
                If CSV_Found Then
                    On Error Resume Next
                    With wsFolderLocations
                        .rows(startRow - 1).Interior.Color = RGB(255, 255, 0)
                        .Cells(startRow - 1, 2).value = "No Walkforward In-Out Periods Analysis Details file found"
                    End With
                End If
                
                detailsDataMissing = True
                missingDetailsDataList = missingDetailsDataList & foldername & " (Path: " & FullFolderPath & ")" & vbNewLine
            End If
        Else
            ' Handle missing Walkforward Files folder
            On Error Resume Next
            With wsFolderLocations
                .rows(startRow).Interior.Color = RGB(255, 255, 0)
                .Cells(startRow, 1).value = startRow - 1
                .Cells(startRow, 2).value = "Walkforward Files folder not found"
                .Cells(startRow, 3).value = foldername
                .Cells(startRow, 4).value = FullFolderPath
                .Cells(startRow, 5).value = baseFolderPath
                .Cells(startRow, 6).value = "No Data"
            End With
            
            equityDataMissing = True
            missingEquityDataList = missingEquityDataList & foldername & " (Path: " & FullFolderPath & ")" & vbNewLine
            startRow = startRow + 1
        End If
    Next folder

CleanExit:
    ' Restore Application properties
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    GetFolderData = startRow
    Exit Function

ErrorHandler:
    Dim errorMsg As String
    errorMsg = "An error occurred in GetFolderData:" & vbNewLine & _
               "Error " & Err.Number & ": " & Err.Description & vbNewLine & _
               "Folder being processed: " & foldername & vbNewLine & _
               "Path: " & FullFolderPath
    
    Debug.Print errorMsg
    MsgBox errorMsg, vbCritical
    
    Resume CleanExit
End Function

Private Function BuildPath(path1 As String, path2 As String) As String
    ' Safely combines path components
    If Right(path1, 1) = "\" Then
        BuildPath = path1 & path2
    Else
        BuildPath = path1 & "\" & path2
    End If
End Function

Function GetStrategyStatus(strategyName As String, ByRef newfolderflag As Boolean, ByRef newfolderlist As String, _
                                                   ByRef strategydupflag As Boolean, ByRef strategyduplist As String) As String
    On Error GoTo ErrorHandler
    
    Dim wsStrategies As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim strategyFound As Boolean
    Dim stratCount As Long
    
    
    stratCount = 0
    
    ' Check if we're skipping all strategies first
    If skipAllStrategies Then
        GetStrategyStatus = "Unknown"
        Exit Function
    End If
    
    Set wsStrategies = ThisWorkbook.Sheets("Strategies")
    lastRow = wsStrategies.Cells(wsStrategies.rows.count, COL_STRAT_STRATEGY_NAME).End(xlUp).row
    
    strategyFound = False
    ' Manually loop through each row to find the strategy
    For i = 2 To lastRow  ' Start from row 2 to skip header
        If StrComp(Trim(wsStrategies.Cells(i, COL_STRAT_STRATEGY_NAME).value), _
                  Trim(strategyName), vbTextCompare) = 0 Then
            
            If stratCount = 0 Then
            ' Get the current status
                Dim currentStatus As String
                currentStatus = wsStrategies.Cells(i, COL_STRAT_STATUS).value
                
                ' Check if status starts with "Not Loaded - " and remove it
                If InStr(1, currentStatus, "Not Loaded - ", vbTextCompare) > 0 Then
                    currentStatus = Replace(currentStatus, "Not Loaded - ", "")
                    wsStrategies.Cells(i, COL_STRAT_STATUS).value = currentStatus
                End If
                
                GetStrategyStatus = currentStatus
                strategyFound = True
            End If
            
            If stratCount > 0 Then
                wsStrategies.Cells(i, COL_STRAT_STATUS).value = "Duplicate Strategy"
           
           End If
            stratCount = stratCount + 1
            
        End If
    Next i
    
    
    
    
    If Not strategyFound Then
        ' If we're adding all strategies, skip the prompt
        ' Add new row at the bottom with error handling
        On Error Resume Next
        lastRow = lastRow + 1
        
                
        If Err.Number <> 0 Then
            MsgBox "Error adding data validation: " & Err.Description, vbCritical
            GetStrategyStatus = "Unknown"
            Exit Function
        End If
        On Error GoTo ErrorHandler
        
        ' Add the new strategy
        wsStrategies.Cells(lastRow, COL_STRAT_STRATEGY_NUMBER).value = "-" 'dummy
        wsStrategies.Cells(lastRow, COL_STRAT_STATUS).value = "New" ' Default status
        wsStrategies.Cells(lastRow, COL_STRAT_STRATEGY_NAME).value = strategyName
        wsStrategies.Cells(lastRow, COL_STRAT_CONTRACTS).value = 1 ' Default contract size
        wsStrategies.Cells(lastRow, COL_STRAT_SYMBOL).value = "-" 'dummy
        wsStrategies.Cells(lastRow, COL_STRAT_TIMEFRAME).value = "-" 'dummy
        wsStrategies.Cells(lastRow, COL_STRAT_TYPE).value = "-" 'dummy
        wsStrategies.Cells(lastRow, COL_STRAT_HORIZON).value = "-" 'dummy
        wsStrategies.Cells(lastRow, COL_STRAT_OTHER).value = "-" 'dummy
        wsStrategies.Cells(lastRow, COL_STRAT_CLOSEDTRADEMC).value = "-" 'dummy
        
        GetStrategyStatus = "New"
        
        newfolderflag = True
        newfolderlist = newfolderlist & strategyName & vbNewLine
        
    End If
    
    If stratCount > 1 Then
        strategydupflag = True
        strategyduplist = strategyduplist & strategyName & " - " & stratCount & " duplicates" & vbNewLine
    End If
    
    
    Exit Function

ErrorHandler:
    GetStrategyStatus = "Unknown"
    MsgBox "An error occurred in GetStrategyStatus: " & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
End Function

Private Function ColumnExists(ws As Worksheet, colNum As Long) As Boolean
    On Error Resume Next
    ColumnExists = (ws.Cells(1, colNum).column = colNum)
    On Error GoTo 0
End Function

Private Function GetLastRow(ws As Worksheet, colNum As Long) As Long
    On Error Resume Next
    GetLastRow = ws.Cells(ws.rows.count, colNum).End(xlUp).row
    If Err.Number <> 0 Then GetLastRow = 1
    On Error GoTo 0
End Function


Function GetFolderName(ByVal folderPath As String) As String
    On Error GoTo ErrorHandler
    
    ' Initialize return value
    GetFolderName = ""
    
    ' Input validation
    If Len(Trim(folderPath)) = 0 Then
        Err.Raise vbObjectError + 1, "GetFolderName", "Folder path is empty"
    End If
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Remove any trailing backslash
    If Right(folderPath, 1) = "\" Then
        folderPath = left(folderPath, Len(folderPath) - 1)
    End If
    
    ' Get just the folder name (everything after the last backslash)
    GetFolderName = FSO.GetFolder(folderPath).name
    
    Exit Function

ErrorHandler:
    Select Case Err.Number
        Case 76 ' Path not found
            GetFolderName = "ERROR: Path not found"
        Case vbObjectError + 1 ' Empty path
            GetFolderName = "ERROR: Empty path"
        Case 424 ' Object required
            GetFolderName = "ERROR: Invalid path format"
        Case Else
            GetFolderName = "ERROR: " & Err.Description
    End Select
    
    ' Optional: Log the error
    Debug.Print "Error in GetFolderName: " & Err.Description & " (Path: " & folderPath & ")"
End Function




Function CheckMissingStrategies(ByRef wsFolderLocations As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wsStrategies As Worksheet
    Dim lastRowStrategies As Long
    Dim lastRowFolders As Long
    Dim i As Long, j As Long
    Dim strategyFound As Boolean
    Dim strategyName As String
    
    ' For reporting
    Dim missingList As String
    Dim missingCount As Long
    Dim message As String
    
    missingList = ""
    missingCount = 0
    
    ' Get Strategies worksheet
    Set wsStrategies = ThisWorkbook.Sheets("Strategies")
    
    ' Get last rows for both sheets
    lastRowStrategies = wsStrategies.Cells(wsStrategies.rows.count, COL_STRAT_STRATEGY_NAME).End(xlUp).row
    lastRowFolders = wsFolderLocations.Cells(wsFolderLocations.rows.count, "B").End(xlUp).row
    
    ' Loop through each strategy in Strategies tab
    For i = 2 To lastRowStrategies
        strategyName = Trim(wsStrategies.Cells(i, COL_STRAT_STRATEGY_NAME).value)
        If strategyName <> "" Then
            strategyFound = False
            
            ' Check if this strategy exists in the folder locations
            For j = 2 To lastRowFolders
                If StrComp(Trim(wsFolderLocations.Cells(j, "B").value), strategyName, vbTextCompare) = 0 Then
                    wsStrategies.Cells(i, COL_STRAT_FOLDER).value = wsFolderLocations.Cells(j, "C").value
                    strategyFound = True
                    Exit For
                End If
            Next j
            
            ' If strategy not found in folders, update the Strategies tab
            If Not strategyFound Then
                missingCount = missingCount + 1
                missingList = missingList & vbNewLine & missingCount & ". " & strategyName
                
                With wsStrategies
                    ' Check if "Not Loaded" is already in the status
                    If InStr(.Cells(i, COL_STRAT_STATUS).value, "Not Loaded") = 0 Then
                        ' Update the status to include "Not Loaded"
                        .Cells(i, COL_STRAT_STATUS).value = "Not Loaded - " & .Cells(i, COL_STRAT_STATUS).value
                    End If
                    
                End With
            End If
        End If
    Next i
    
  
    CheckMissingStrategies = True
    Exit Function

ErrorHandler:
    MsgBox "An error occurred in CheckMissingStrategies: " & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    CheckMissingStrategies = False
End Function



Private Function GetStatusOrder(status As String, statusOrder As Collection) As Integer
    Dim i As Integer
    
    ' Loop through status order collection
    For i = 1 To statusOrder.count
        If StrComp(status, statusOrder(i), vbTextCompare) = 0 Then
            GetStatusOrder = i
            Exit Function
        End If
    Next i
    
    ' If status not found in order list, put at end
    GetStatusOrder = 999
End Function



Function GetOrderedStatusList() As Collection
    On Error GoTo ErrorHandler
    
    Dim statusList As Collection
    Dim statusArr As Variant
    Dim status As Variant
    Dim portStatus As String
    Dim passStatus As String
    Dim statusOptions As String
    Dim buyHoldStatus As String
    
    Set statusList = New Collection
    
    ' Add Portfolio Status (typically "Live")
    portStatus = GetNamedRangeValue("Port_Status")
    If portStatus <> "" Then statusList.Add Trim(portStatus)
    
    ' Add Pass Status
    passStatus = GetNamedRangeValue("Pass_Status")
    If passStatus <> "" Then statusList.Add Trim(passStatus)
    
    ' Add statuses from StatusOptions
    statusOptions = GetNamedRangeValue("StatusOptions")
    If Len(Trim(statusOptions)) > 0 Then
        statusArr = Split(statusOptions, ",")
        For Each status In statusArr
            If Len(Trim(status)) > 0 Then
                statusList.Add Trim(status)
            End If
        Next status
    End If
    
    ' Add Buy and Hold Status
    buyHoldStatus = GetNamedRangeValue("BuyandHoldStatus")
    If buyHoldStatus <> "" Then statusList.Add Trim(buyHoldStatus)
    
    Set GetOrderedStatusList = statusList
    Exit Function
    
ErrorHandler:
    MsgBox "Error in GetOrderedStatusList: " & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    Set GetOrderedStatusList = New Collection ' Return empty collection instead of Nothing
End Function


Function GetStatusListAsString(Optional delimiter As String = ",") As String
    On Error GoTo ErrorHandler
    
    Dim statusList As Collection
    Dim status As Variant
    Dim result As String
    
    Set statusList = GetOrderedStatusList()
    
    ' Build delimited string
    For Each status In statusList
        If Len(result) > 0 Then
            result = result & delimiter
        End If
        result = result & CStr(status)
    Next status
    
    GetStatusListAsString = result
    Exit Function
    
ErrorHandler:
    MsgBox "Error in GetStatusListAsString: " & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    GetStatusListAsString = ""
End Function



Function GetStatusOrderNumber(status As String) As Integer
    On Error GoTo ErrorHandler
    
    Dim statusList As Collection
    Dim i As Integer
    
    Set statusList = GetOrderedStatusList()
    
    ' Look for exact match first
    For i = 1 To statusList.count
        If StrComp(status, statusList(i), vbTextCompare) = 0 Then
            GetStatusOrderNumber = i
            Exit Function
        End If
    Next i
    
    ' Check for "Not Loaded" prefix
    If InStr(1, status, "Not Loaded - ", vbTextCompare) > 0 Then
        status = Replace(status, "Not Loaded - ", "")
        For i = 1 To statusList.count
            If StrComp(status, statusList(i), vbTextCompare) = 0 Then
                GetStatusOrderNumber = i
                Exit Function
            End If
        Next i
    End If
    
    ' If not found, return high number to sort at end
    GetStatusOrderNumber = 999
    Exit Function
    
ErrorHandler:
    MsgBox "Error in GetStatusOrderNumber: " & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    GetStatusOrderNumber = 999
End Function



Function GetNamedRangeValue(rangeName As String, Optional defaultValue As Variant = "") As Variant
    Dim nm As name
    Dim rng As Range
    
    ' First check if the named range exists
    On Error Resume Next
    Set nm = ThisWorkbook.Names(rangeName)
    If Err.Number <> 0 Then
        Debug.Print "Warning: Named range '" & rangeName & "' does not exist"
        GetNamedRangeValue = defaultValue
        Err.Clear
        Exit Function
    End If
    
    ' Try to get the range it refers to
    Set rng = nm.RefersToRange
    If Err.Number <> 0 Or rng Is Nothing Then
        Debug.Print "Warning: Named range '" & rangeName & "' has invalid reference"
        GetNamedRangeValue = defaultValue
        Err.Clear
        Exit Function
    End If
    
    ' Get the value
    GetNamedRangeValue = rng.value
    If Err.Number <> 0 Then
        Debug.Print "Warning: Error reading value from named range '" & rangeName & "'"
        GetNamedRangeValue = defaultValue
        Err.Clear
    End If
    
    Set rng = Nothing
    Set nm = Nothing
    On Error GoTo 0
End Function



Function GetWorksheetSafe(workbookRef As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetSafe = workbookRef.Worksheets(sheetName)
    If Err.Number <> 0 Or GetWorksheetSafe Is Nothing Then
        Debug.Print "Warning: Worksheet '" & sheetName & "' not found"
        Set GetWorksheetSafe = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End Function

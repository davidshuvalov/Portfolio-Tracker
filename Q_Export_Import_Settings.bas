Attribute VB_Name = "Q_Export_Import_Settings"
Sub ExportStrategyBacktestInputs()
    Dim wbNew As Workbook
    Dim wsStrategies As Worksheet, wsBacktest As Worksheet, wsInputs As Worksheet, wsConfig As Worksheet
    Dim wsInputsSource As Worksheet
    Dim rng As name
    Dim savePath As String
    Dim rowIndex As Integer
    Dim versionNumber As String
    
    ' Define the version number of the tool
    versionNumber = Range("version").value ' Update as needed

    Dim defaultFileName As String
    defaultFileName = "PortfolioTrackerConfig_" & Format(Now, "yyyy-mm-dd") & ".xlsx"

    ' Ask user to select folder and enter file name
    Set FileDialog = Application.FileDialog(msoFileDialogSaveAs)
    With FileDialog
        .title = "Save Configuration File"
        .InitialFileName = defaultFileName
        .FilterIndex = 1
        If .Show = -1 Then
            savePath = .SelectedItems(1) ' Get user-selected file path
        Else
            MsgBox "Export cancelled.", vbInformation
            Exit Sub
        End If
    End With
    ' Create new workbook
    Set wbNew = Workbooks.Add
    
    ' Create Config Sheet
    Set wsConfig = wbNew.Sheets.Add
    wsConfig.name = "Config"
    wsConfig.Cells(1, 1).value = "Portfolio Tracker Configuration File"
    wsConfig.Cells(2, 1).value = "Version:"
    wsConfig.Cells(2, 2).value = versionNumber
    wsConfig.Cells(3, 1).value = "Generated On:"
    wsConfig.Cells(3, 2).value = Format(Now, "yyyy-mm-dd HH:MM:SS")
    wsConfig.Cells(4, 1).value = "Required Sheets:"
    wsConfig.Cells(5, 1).value = "Strategies"
    wsConfig.Cells(6, 1).value = "Backtest"
    wsConfig.Cells(7, 1).value = "Inputs"
    wsConfig.Columns("A:B").AutoFit ' Format the sheet

   
    ' Copy Strategies Sheet - Values with full formatting and column widths
    On Error Resume Next
    Set wsStrategies = ThisWorkbook.Sheets("Strategies")
    On Error GoTo 0
    If Not wsStrategies Is Nothing Then
        ' Create new sheet
        Set wsStrategies = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.count))
        wsStrategies.name = "Strategies"
        
        ' Copy column widths first
        Dim col As Integer
        For col = 1 To ThisWorkbook.Sheets("Strategies").UsedRange.Columns.count
            wsStrategies.Columns(col).ColumnWidth = ThisWorkbook.Sheets("Strategies").Columns(col).ColumnWidth
        Next col
        
        ' Copy all formatting including borders
        ThisWorkbook.Sheets("Strategies").UsedRange.Copy
        wsStrategies.Range("A1").PasteSpecial xlPasteFormats
        
        ' Then paste values
        ThisWorkbook.Sheets("Strategies").UsedRange.Copy
        wsStrategies.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End If
    
    ' Copy Backtest Sheet - Values with full formatting and column widths
    On Error Resume Next
    Set wsBacktest = ThisWorkbook.Sheets("Backtest")
    On Error GoTo 0
    If Not wsBacktest Is Nothing Then
        ' Create new sheet
        Set wsBacktest = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.count))
        wsBacktest.name = "Backtest"
        
        ' Copy column widths first
        Dim bcol As Integer
        For bcol = 1 To ThisWorkbook.Sheets("Backtest").UsedRange.Columns.count
            wsBacktest.Columns(bcol).ColumnWidth = ThisWorkbook.Sheets("Backtest").Columns(bcol).ColumnWidth
        Next bcol
        
        ' Copy all formatting including borders
        ThisWorkbook.Sheets("Backtest").UsedRange.Copy
        wsBacktest.Range("A1").PasteSpecial xlPasteFormats
        
        ' Then paste values
        ThisWorkbook.Sheets("Backtest").UsedRange.Copy
        wsBacktest.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End If
   
    ' Create Inputs Sheet
    Set wsInputs = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.count))
    wsInputs.name = "Inputs"
    wsInputs.Cells(1, 1).value = "Named Range"
    wsInputs.Cells(1, 2).value = "Type"
    wsInputs.Cells(1, 3).value = "Values"
    wsInputs.Cells(1, 1).Font.Bold = True
    wsInputs.Cells(1, 2).Font.Bold = True
    wsInputs.Cells(1, 3).Font.Bold = True

    ' Get Named Ranges from "Inputs" Sheet
    On Error Resume Next
    Set wsInputsSource = ThisWorkbook.Sheets("Inputs")
    On Error GoTo 0
    If wsInputsSource Is Nothing Then
        MsgBox "Warning: 'Inputs' sheet not found. Named ranges will not be exported.", vbExclamation
    Else
        rowIndex = 2 ' Start from second row
        For Each rng In ThisWorkbook.Names
            ' Check if named range refers to "Inputs" sheet
            If InStr(1, rng.RefersTo, wsInputsSource.name, vbTextCompare) > 0 Then
                wsInputs.Cells(rowIndex, 1).value = rng.name ' Named range name
                
                On Error Resume Next
                Dim rngRef As Range
                Set rngRef = rng.RefersToRange
                
                If Not rngRef Is Nothing Then
                    ' Check if it's a multi-cell range
                    If rngRef.Cells.count > 1 Then
                        wsInputs.Cells(rowIndex, 2).value = "Table"
                        
                        ' For tables, store as JSON-like format for easier parsing
                        Dim dataString As String
                        Dim i As Long, j As Long
                        dataString = ""
                        
                        For i = 1 To rngRef.rows.count
                            For j = 1 To rngRef.Columns.count
                                dataString = dataString & rngRef.Cells(i, j).value
                                If j < rngRef.Columns.count Then
                                    dataString = dataString & vbTab ' Tab separator between columns
                                End If
                            Next j
                            If i < rngRef.rows.count Then
                                dataString = dataString & vbLf ' Line feed between rows
                            End If
                        Next i
                        
                        ' Store dimensions for table reconstruction
                        wsInputs.Cells(rowIndex, 3).value = rngRef.rows.count & "|" & rngRef.Columns.count & "|" & dataString
                    Else
                        wsInputs.Cells(rowIndex, 2).value = "Single"
                        wsInputs.Cells(rowIndex, 3).value = rngRef.value
                    End If
                End If
                On Error GoTo 0
                rowIndex = rowIndex + 1
            End If
        Next rng
    End If
    
    ' Autofit columns
    wsInputs.Columns("A:C").AutoFit
    
    ' Turn off alerts to prevent confirmation prompt
    Application.DisplayAlerts = False
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wbNew.Sheets("Sheet1")
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
  
    ' Turn alerts back on
    Application.DisplayAlerts = True

    ' Ensure savePath is valid before saving
    If Trim(savePath) = "" Then
        MsgBox "Error: No valid folder selected. Export aborted.", vbCritical
        wbNew.Close False
        Exit Sub
    End If

    ' Save and close only if savePath is valid
    Application.DisplayAlerts = False
    wbNew.SaveAs savePath, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close False
    Application.DisplayAlerts = True

    MsgBox "Export complete: " & vbNewLine & savePath, vbInformation
End Sub

Sub ImportConfigurationFile()
    Dim wbImport As Workbook
    Dim wsImport As Worksheet
    Dim wsPortfolio As Worksheet
    Dim wsBacktest As Worksheet
    Dim wsInputs As Worksheet
    Dim wsImportInputs As Worksheet
    Dim rng As Range
    Dim lastRow As Long, lastCol As Long
    Dim filePath As String
    Dim namedRange As name
    Dim namedRangeDict As Object
    Dim missingRanges As String
    Dim rowIndex As Long
    Dim rangeName As String
    Dim rangeValue As Variant
    Dim rangeType As String
    
    Set namedRangeDict = CreateObject("Scripting.Dictionary")

    ' Prompt user to select file
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Select Configuration File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsb; *.xlsm", 1
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "Import cancelled.", vbInformation
            Exit Sub
        End If
    End With

    ' Open the selected workbook
    On Error Resume Next
    Set wbImport = Workbooks.Open(filePath, ReadOnly:=True)
    On Error GoTo 0
    If wbImport Is Nothing Then
        MsgBox "Error opening file.", vbCritical
        Exit Sub
    End If

    ' Check if 'Config' sheet exists using loop
    Dim SheetExists As Boolean
    SheetExists = False
    Dim ws As Worksheet
    For Each ws In wbImport.Sheets
        If ws.name = "Config" Then
            SheetExists = True
            Exit For
        End If
    Next ws

    ' If 'Config' does not exist, exit with error
    If Not SheetExists Then
        wbImport.Close False
        MsgBox "Error: This file does not appear to be a valid configuration file. No 'Config' sheet found.", vbCritical
        Exit Sub
    End If

    ' Set reference to 'Config' sheet (now safe)
    Dim wsConfig As Worksheet
    Set wsConfig = wbImport.Sheets("Config")

    ' Validate that A1 contains the correct identifier
    Dim configCheck As String
    configCheck = Trim(wsConfig.Cells(1, 1).value)
    If configCheck <> "Portfolio Tracker Configuration File" Then
        wbImport.Close False
        MsgBox "Error: The 'Config' sheet does not contain a valid configuration file identifier in cell A1.", vbCritical
        Exit Sub
    End If
    
    ' Set references to target sheets in Portfolio Tracker
    On Error Resume Next
    Set wsPortfolio = ThisWorkbook.Sheets("Strategies")
    Set wsBacktest = ThisWorkbook.Sheets("Backtest")
    Set wsInputs = ThisWorkbook.Sheets("Inputs")
    On Error GoTo 0

    If wsPortfolio Is Nothing Or wsBacktest Is Nothing Or wsInputs Is Nothing Then
        MsgBox "Error: One or more required sheets are missing.", vbExclamation
        wbImport.Close False
        Exit Sub
    End If

    ' Copy Strategies Data with Formatting
    On Error Resume Next
    Set wsImport = wbImport.Sheets("Strategies")
    On Error GoTo 0
    If Not wsImport Is Nothing Then
        ' Clear existing content but preserve column structure
        wsPortfolio.Cells.Clear ' This clears both content and formatting
        
        ' Copy everything including all formatting
        wsImport.Cells.Copy
        wsPortfolio.Cells.PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        ' Fix date values that may have been converted to text

        lastRow = wsImport.Cells(wsImport.rows.count, 1).End(xlUp).row
        lastCol = wsImport.Cells(1, wsImport.Columns.count).End(xlToLeft).column
        
        Dim i As Long, j As Long
        For i = 1 To lastRow
            For j = 2 To lastCol ' Check columns 2 and 3
                If Not IsEmpty(wsImport.Cells(i, j).value) Then
                    ' Copy the number format and font specifically for date cells
                    wsPortfolio.Cells(i, j).NumberFormat = wsImport.Cells(i, j).NumberFormat
                    With wsPortfolio.Cells(i, j).Font
                        .name = wsImport.Cells(i, j).Font.name
                        .Size = wsImport.Cells(i, j).Font.Size
                        .Bold = wsImport.Cells(i, j).Font.Bold
                        .Italic = wsImport.Cells(i, j).Font.Italic
                        .Color = wsImport.Cells(i, j).Font.Color
                    End With
                    
                    ' If the source is a date but target is text, convert it properly
                    If IsDate(wsImport.Cells(i, j).value) And Not IsDate(wsPortfolio.Cells(i, j).value) Then
                        ' Convert text date back to proper date value
                        Dim dateVal As Date
                        dateVal = CDate(wsImport.Cells(i, j).value)
                        wsPortfolio.Cells(i, j).value = dateVal
                        wsPortfolio.Cells(i, j).NumberFormat = wsImport.Cells(i, j).NumberFormat
                    End If
                End If
            Next j
        Next i
    End If

    ' Copy Backtest Data with Formatting
    On Error Resume Next
    Set wsImport = wbImport.Sheets("Backtest")
    On Error GoTo 0
    If Not wsImport Is Nothing Then
        ' Clear existing content but preserve column structure
        wsBacktest.Cells.Clear ' This clears both content and formatting
        
        ' Copy everything including all formatting
        wsImport.Cells.Copy
        wsBacktest.Cells.PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        ' Fix date values that may have been converted to text
        lastRow = wsImport.Cells(wsImport.rows.count, 1).End(xlUp).row
        lastCol = wsImport.Cells(1, wsImport.Columns.count).End(xlToLeft).column
        
        For i = 1 To lastRow
            For j = 2 To lastCol ' Check columns 2 and 3
                If Not IsEmpty(wsImport.Cells(i, j).value) Then
                    ' Copy the number format and font specifically for date cells
                    wsBacktest.Cells(i, j).NumberFormat = wsImport.Cells(i, j).NumberFormat
                    With wsBacktest.Cells(i, j).Font
                        .name = wsImport.Cells(i, j).Font.name
                        .Size = wsImport.Cells(i, j).Font.Size
                        .Bold = wsImport.Cells(i, j).Font.Bold
                        .Italic = wsImport.Cells(i, j).Font.Italic
                        .Color = wsImport.Cells(i, j).Font.Color
                    End With
                    
                    ' If the source is a date but target is text, convert it properly
                    If IsDate(wsImport.Cells(i, j).value) And Not IsDate(wsBacktest.Cells(i, j).value) Then
                        ' Convert text date back to proper date value
     
                        dateVal = CDate(wsImport.Cells(i, j).value)
                        wsBacktest.Cells(i, j).value = dateVal
                        wsBacktest.Cells(i, j).NumberFormat = wsImport.Cells(i, j).NumberFormat
                    End If
                End If
            Next j
        Next i
    End If

    ' Process Inputs (Named Ranges)
    On Error Resume Next
    Set wsImportInputs = wbImport.Sheets("Inputs")
    On Error GoTo 0
    If wsImportInputs Is Nothing Then
        MsgBox "Warning: 'Inputs' sheet not found in the configuration file.", vbExclamation
    Else
        ' Get last row in Inputs sheet
        lastRow = wsImportInputs.Cells(wsImportInputs.rows.count, 1).End(xlUp).row

        ' Loop through named ranges and update values in Portfolio Tracker
        For rowIndex = 2 To lastRow ' Start from row 2 to skip headers
            ' Check if this is the old format (2 columns) or new format (3 columns)
            ' Old format: Column A = Name, Column B = Value
            ' New format: Column A = Name, Column B = Type, Column C = Value
            
            rangeName = wsImportInputs.Cells(rowIndex, 1).value ' Column A: Named range name
            
            ' Check if Column B contains "Single" or "Table" (new format)
            Dim tempValue As Variant
            tempValue = wsImportInputs.Cells(rowIndex, 2).value
            
            If tempValue = "Single" Or tempValue = "Table" Then
                ' New format with Type column
                rangeType = wsImportInputs.Cells(rowIndex, 2).value ' Column B: Type
                rangeValue = wsImportInputs.Cells(rowIndex, 3).value ' Column C: Values
            Else
                ' Old format without Type column
                rangeType = "Single" ' Default to single value
                rangeValue = wsImportInputs.Cells(rowIndex, 2).value ' Column B: Value
            End If

            ' Skip empty rows
            If rangeName <> "" Then
                ' Check if named range exists in the Portfolio Tracker
                On Error Resume Next
                Set namedRange = ThisWorkbook.Names(rangeName)
                On Error GoTo 0

                If Not namedRange Is Nothing Then
                    ' Handle different types
                    If rangeType = "Single" Then
                        ' Update single value named range
                        On Error Resume Next
                        namedRange.RefersToRange.value = rangeValue
                        On Error GoTo 0
                    ElseIf rangeType = "Table" Then
                        ' Parse and update table named range
                        Call ImportTableNamedRange(namedRange, rangeValue)
                    Else
                        ' Handle legacy format (no type specified) - assume single value
                        On Error Resume Next
                        namedRange.RefersToRange.value = rangeValue
                        On Error GoTo 0
                    End If
                Else
                    ' If named range does not exist, add to missing list
                    missingRanges = missingRanges & rangeName & vbNewLine
                End If
            End If
        Next rowIndex
    End If

    ' Close imported workbook
    wbImport.Close False
    
    ThisWorkbook.Sheets("Control").Activate

    ' Notify user of missing named ranges
    If missingRanges <> "" Then
        MsgBox "The following named ranges could not be updated because they do not exist in the Portfolio Tracker:" & vbNewLine & missingRanges, vbExclamation, "Missing Named Ranges"
    Else
        MsgBox "Configuration file successfully imported!", vbInformation
    End If
End Sub

Private Sub ImportTableNamedRange(namedRange As name, ByVal tableData As Variant)
    ' Parse table data and update the named range
    Dim parts() As String
    Dim dimensions() As String
    Dim rows() As String
    Dim cols() As String
    Dim targetRange As Range
    Dim numRows As Long, numCols As Long
    Dim i As Long, j As Long
    Dim dataContent As String
    
    ' Convert to string if needed
    If Not IsEmpty(tableData) Then
        tableData = CStr(tableData)
    Else
        Exit Sub
    End If
    
    ' Split the format: numRows|numCols|data
    parts = Split(tableData, "|", 3)
    If UBound(parts) < 2 Then Exit Sub
    
    ' Add error handling for numeric conversions
    On Error Resume Next
    numRows = CLng(parts(0))
    numCols = CLng(parts(1))
    On Error GoTo 0
    
    If numRows <= 0 Or numCols <= 0 Then Exit Sub
    
    dataContent = parts(2)
    
    ' Get the starting cell of the named range
    On Error Resume Next
    Set targetRange = namedRange.RefersToRange
    On Error GoTo 0
    
    If targetRange Is Nothing Then Exit Sub
    
    ' Expand the range to match the table dimensions
    Set targetRange = targetRange.Cells(1, 1).Resize(numRows, numCols)
    
    ' Parse the data
    rows = Split(dataContent, vbLf)
    
    For i = 0 To UBound(rows)
        If i < numRows Then
            cols = Split(rows(i), vbTab)
            For j = 0 To UBound(cols)
                If j < numCols Then
                    targetRange.Cells(i + 1, j + 1).value = cols(j)
                End If
            Next j
        End If
    Next i
    
    ' Update the named range to point to the filled table
    namedRange.RefersTo = targetRange
End Sub

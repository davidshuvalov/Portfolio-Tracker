Attribute VB_Name = "H_Open_Code_Tab"

Sub OpenStrategyCodeFile(strategyName As String, strategyNumber As Long, Optional action As String = "tab")
    ' action can be: "tab" (default), "file", or "folder"
    
    Dim ws As Worksheet
    Dim newsheet As Worksheet
    Dim wsMWFolderLocations As Worksheet
    Dim filePath As String
    Dim targetSheetName As String
    Dim fileName As String
    Dim strategyRow As Long
     Dim folderPath As String
     
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Check if "MW Folder Locations" sheet exists and has data in row 2
    On Error Resume Next
    Set wsMWFolderLocations = ThisWorkbook.Sheets("MW Folder Locations")
    On Error GoTo 0

    ' Exit and show error if the sheet doesn't exist
    If wsMWFolderLocations Is Nothing Then
        MsgBox "Error: 'MW Folder Locations' sheet does not exist.", vbExclamation
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Exit and show error if the sheet exists but has no data in row 2
    If wsMWFolderLocations.Cells(2, 1).value = "" Then
        MsgBox "Error: 'MW Folder Locations' sheet exists but contains no data in row 2.", vbExclamation
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If
    
    targetSheetName = "Strat - " & strategyNumber & " - Code"
    
    ' Look up the filename in the StrategyList tab based on the strategy name
    Dim lastRowFolder As Long, counter As Long
    lastRowFolder = wsMWFolderLocations.Cells(wsMWFolderLocations.rows.count, 1).End(xlUp).row
    strategyRow = -99
    
    On Error Resume Next
    For counter = 2 To lastRowFolder
         If strategyName = wsMWFolderLocations.Cells(counter, 2).value Then strategyRow = counter
    Next counter
    On Error GoTo 0
    
    ' Check if strategy name is found
    If strategyRow = -99 Then
        MsgBox "Strategy name not found in folder location.", vbExclamation
        Exit Sub
    End If

    ' Get the corresponding filename from the table
    fileName = wsMWFolderLocations.Cells(strategyRow, 4).value
    filePath = GetShortPath(fileName & "\Walkforward Files\" & strategyName & " ELCode.txt")
    folderPath = GetShortPath(fileName & "\Walkforward Files")
    
    
    Select Case LCase(action)
        Case "file"
            ' Open the text file in notepad
            If Dir(filePath) = "" Then
                MsgBox "File not found: " & filePath, vbExclamation
                Exit Sub
            End If
            
            On Error Resume Next
            Shell "notepad.exe """ & filePath & """", vbNormalFocus
            If Err.Number <> 0 Then
                Shell "C:\WINDOWS\explorer.exe """ & filePath & """", vbNormalFocus
            End If
            On Error GoTo 0
            
        Case "folder"
            ' Open the folder in Windows Explorer
           
            
            
            If Dir(folderPath, vbDirectory) = "" Then
                MsgBox "Folder not found: " & folderPath, vbExclamation
                Exit Sub
            End If
            
            Shell "explorer.exe """ & folderPath & """", vbNormalFocus
            
        Case "tab"
            ' Original functionality to create tab
            If Dir(filePath) = "" Then
                MsgBox "File not found: " & filePath, vbExclamation
                Exit Sub
            End If
            
            ' Delete existing sheet if it already exists
            Application.DisplayAlerts = False
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(targetSheetName)
            If Not ws Is Nothing Then ws.Delete
            On Error GoTo 0
            Application.DisplayAlerts = True

            ' Create new sheet with target sheet name
            Set newsheet = ThisWorkbook.Sheets.Add
            newsheet.name = targetSheetName
            newsheet.Tab.Color = RGB(228, 158, 221)
            
            ' Set white background color for the entire worksheet
            newsheet.Cells.Interior.Color = RGB(255, 255, 255)
            
            ' Open the text file and read its contents
            Dim fileNumber As Integer
            fileNumber = FreeFile
            Open filePath For Input As #fileNumber

            ' Insert each line into the new sheet
            Dim row As Integer
            Dim fileLine As String
            row = 1
            Do While Not EOF(fileNumber)
                Line Input #fileNumber, fileLine
                newsheet.Cells(row, 1).value = fileLine
                row = row + 1
            Loop

            Close #fileNumber
            
            newsheet.Cells(1, 11).value = strategyNumber
            newsheet.Cells(1, 12).value = strategyName
            
            ' Add all the buttons and formatting
            Call AddStrategySheetButtons(newsheet, strategyName, strategyNumber)
            
            With ThisWorkbook.Windows(1)
                .Zoom = 70 ' Set zoom level to 70%
            End With
            
             
            
            Call OrderVisibleTabsBasedOnList
            newsheet.Activate
            
        Case Else
            MsgBox "Invalid action specified. Use 'tab', 'file', or 'folder'.", vbExclamation
    End Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub AddStrategySheetButtons(newsheet As Worksheet, strategyName As String, strategyNumber As Long)
    Dim btn As Object
    Dim Mainlogicrow As Integer
    Dim targetCell As Range
    
    ' Find the "MAIN STRATEGY LOGIC" row
    For Mainlogicrow = 1 To 10000
        If InStr(1, newsheet.Cells(Mainlogicrow, 1).value, "MAIN STRATEGY LOGIC", vbTextCompare) > 0 Then
            ' Add buttons at MAIN STRATEGY LOGIC location
            Call AddButtonSet(newsheet, Mainlogicrow)
            
            newsheet.Cells(Mainlogicrow - 1, 11).value = strategyNumber
            newsheet.Cells(Mainlogicrow - 1, 12).value = strategyName
            
            Set targetCell = newsheet.Cells(Mainlogicrow - 1, 1)
            targetCell.Select
            Application.GoTo targetCell, True
            
            Exit For
        End If
    Next Mainlogicrow
    
    ' Add buttons at the top
    Call AddButtonSet(newsheet, 2)
End Sub

Private Sub AddButtonSet(sheet As Worksheet, startRow As Integer)
    Dim btn As Object
    Dim buttonConfigs As Variant
    Dim i As Integer
    
    ' Define button configurations: Caption, OnAction pairs
    buttonConfigs = Array( _
        Array("Delete Tab", "DeleteStrategyCodeTab"), _
        Array("Summary Tab", "GoToSummary"), _
        Array("Portfolio Tab", "GoToPortfolio"), _
        Array("Control Tab", "GoToControl"), _
        Array("Open Detailed Strategy Tab", "ButtonClickHandlerDetailedStrat"), _
        Array("Strategies Tab", "GoToStrategies"), _
        Array("Inputs Tab", "GoToInputs"))
    
    ' Create buttons
    For i = 0 To UBound(buttonConfigs)
        Set btn = sheet.Buttons.Add( _
            left:=sheet.Cells(1, 16).left, _
            top:=sheet.Cells(startRow + (i * 2), 1).top, _
            Width:=100, _
            Height:=25)
        
        With btn
            .Caption = buttonConfigs(i)(0)
            .OnAction = buttonConfigs(i)(1)
        End With
    Next i
End Sub























Sub CreateStatusCodeTab(status1 As String, status2 As String)
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tabsCreated As Long

    ' Initialize column constants manually
    Call InitializeColumnConstantsManually
    
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
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' Find the last row in the Summary sheet
    lastRow = wsSummary.Cells(wsSummary.rows.count, 1).End(xlUp).row

    ' Initialize counter for tabs created
    tabsCreated = 0

    ' Loop through the rows to create tabs based on status
    For i = 2 To lastRow
        If status1 = wsSummary.Cells(i, COL_STATUS).value Or status2 = wsSummary.Cells(i, COL_STATUS).value Then
            OpenStrategyCodeFile wsSummary.Cells(i, COL_STRATEGY_NAME).value, wsSummary.Cells(i, COL_STRATEGY_NUMBER).value
            tabsCreated = tabsCreated + 1
              
            Application.StatusBar = "Creating Strategy Code Tab: " & wsSummary.Cells(i, COL_STRATEGY_NUMBER).value
        End If
    Next i
    Application.StatusBar = False
     Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Dim status As String
    If status2 <> "" Then status = status1 & " or " & status2 Else status = status1
    
    ' Display a message based on the number of tabs created
    If tabsCreated > 0 Then
        MsgBox tabsCreated & " code tabs created successfully.", vbInformation
    Else
        MsgBox "No code tabs were created. No entries found with status '" & status & "'.", vbExclamation
    End If
End Sub





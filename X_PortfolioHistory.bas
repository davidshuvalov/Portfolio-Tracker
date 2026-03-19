Attribute VB_Name = "X_PortfolioHistory"
Option Explicit

' ============================================================
' X_PortfolioHistory — Portfolio snapshot management and what-if analysis
'
' Allows users to save the current portfolio (strategies + contracts)
' as a named snapshot, then run performance analysis on any saved
' snapshot using the existing DailyM2MEquity data.
'
' Answers the question: "What if we hadn't made changes to the portfolio?"
'
' Entry points:
'   SaveCurrentPortfolioSnapshot  — button on PortfolioHistory sheet
'   RunSnapshotFromButton         — [Run] buttons in the snapshot list
'   DeleteSnapshotFromButton      — [Delete] buttons in the snapshot list
'   GoToPortfolioHistory          — navigation from Control / Portfolio
'   RefreshPortfolioHistorySheet  — rebuild the sheet list (no re-run)
'
' Storage (no external files needed):
'   PortfolioSnapshots (hidden sheet) — one row per strategy per snapshot
'   PortfolioHistory   (visible sheet) — snapshot list + last performance results
'
' Error handling for missing strategies:
'   If a historical strategy is not found in DailyM2MEquity, it is
'   flagged visually (red row, "NOT FOUND" label) and excluded from
'   the performance calculation. All other strategies still run.
' ============================================================

' ---- Storage sheet column layout ----------------------------
Private Const SNAP_SHEET_NAME    As String = "PortfolioSnapshots"
Private Const HISTORY_SHEET_NAME As String = "PortfolioHistory"

Private Const SN_ID      As Long = 1   ' SnapshotID (auto-increment integer)
Private Const SN_NAME    As Long = 2   ' SnapshotName (user text)
Private Const SN_DATE    As Long = 3   ' SavedDate (Now())
Private Const SN_STRAT   As Long = 4   ' StrategyName
Private Const SN_CONTR   As Long = 5   ' Contracts
Private Const SN_SYMBOL  As Long = 6   ' Symbol
Private Const SN_SECTOR  As Long = 7   ' Sector

' ---- PortfolioHistory sheet fixed row positions -------------
Private Const H_TITLE_ROW  As Long = 1
Private Const H_BTN_ROW    As Long = 2
Private Const H_HDR_ROW    As Long = 4   ' column headers for snapshot list
Private Const H_LIST_START As Long = 5   ' first snapshot data row


' =====================================================================
' ENTRY POINTS
' =====================================================================

Sub SaveCurrentPortfolioSnapshot()
    Call InitializeColumnConstantsManually

    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license.", vbCritical
        Exit Sub
    End If

    ' Guard: Portfolio sheet must exist and have data
    Dim wsPort As Worksheet
    On Error Resume Next
    Set wsPort = ThisWorkbook.Sheets("Portfolio")
    On Error GoTo 0
    If wsPort Is Nothing Then
        MsgBox "Portfolio sheet not found. Please run 'Create Portfolio Summary' first.", _
               vbExclamation, "Save Snapshot"
        Exit Sub
    End If

    ' Prompt for snapshot name
    Dim snapName As String
    snapName = Trim(InputBox( _
        "Enter a name for this portfolio snapshot:" & vbCrLf & _
        "(e.g.  ""Q1 2025"",  ""Pre-rebalance Mar 2025"")", _
        "Save Portfolio Snapshot"))
    If snapName = "" Then Exit Sub

    ' Handle duplicate name
    Dim wsSnap As Worksheet
    Set wsSnap = GetOrCreateSnapshotsSheet()
    If SnapshotExists(wsSnap, snapName) Then
        Dim overwrite As Long
        overwrite = MsgBox( _
            "A snapshot named """ & snapName & """ already exists." & vbCrLf & _
            "Overwrite it?", vbYesNo + vbQuestion, "Duplicate Snapshot")
        If overwrite = vbNo Then Exit Sub
        Call DeleteSnapshotRows(wsSnap, snapName)
    End If

    ' Read Portfolio sheet rows
    Dim lastPortRow As Long
    lastPortRow = wsPort.Cells(wsPort.Rows.Count, COL_PORT_STRATEGY_NAME).End(xlUp).Row
    If lastPortRow < 2 Then
        MsgBox "No strategies found in Portfolio sheet.", vbExclamation, "Save Snapshot"
        Exit Sub
    End If

    Dim nextID As Long
    nextID = NextSnapshotID(wsSnap)

    Dim lastSnapRow As Long
    lastSnapRow = wsSnap.Cells(wsSnap.Rows.Count, SN_ID).End(xlUp).Row
    If lastSnapRow < 1 Then lastSnapRow = 1  ' safety: at least the header row

    Dim stratCount As Long
    stratCount = 0
    Dim savedDate As Date
    savedDate = Now()

    Dim i As Long
    For i = 2 To lastPortRow
        Dim stratName As String
        stratName = Trim(CStr(wsPort.Cells(i, COL_PORT_STRATEGY_NAME).Value))
        If stratName = "" Then GoTo NextPortRow

        lastSnapRow = lastSnapRow + 1
        wsSnap.Cells(lastSnapRow, SN_ID).Value     = nextID
        wsSnap.Cells(lastSnapRow, SN_NAME).Value   = snapName
        wsSnap.Cells(lastSnapRow, SN_DATE).Value   = savedDate
        wsSnap.Cells(lastSnapRow, SN_STRAT).Value  = stratName
        wsSnap.Cells(lastSnapRow, SN_CONTR).Value  = wsPort.Cells(i, COL_PORT_CONTRACTS).Value
        wsSnap.Cells(lastSnapRow, SN_SYMBOL).Value = wsPort.Cells(i, COL_PORT_SYMBOL).Value
        wsSnap.Cells(lastSnapRow, SN_SECTOR).Value = wsPort.Cells(i, COL_PORT_SECTOR).Value
        stratCount = stratCount + 1
NextPortRow:
    Next i

    If stratCount = 0 Then
        MsgBox "No strategies were saved. Check that the Portfolio sheet has data.", _
               vbExclamation, "Save Snapshot"
        Exit Sub
    End If

    Call RefreshPortfolioHistorySheet

    MsgBox "Snapshot """ & snapName & """ saved with " & stratCount & " strategies.", _
           vbInformation, "Snapshot Saved"
End Sub


Sub RunSnapshotFromButton()
    ' Called by every [Run] FormControl button in the snapshot list.
    ' Application.Caller returns the shape name; TopLeftCell gives the row.
    On Error GoTo ErrHandler
    Dim btn As Shape
    Set btn = ActiveSheet.Shapes(Application.Caller)
    Dim snapName As String
    snapName = Trim(CStr(ActiveSheet.Cells(btn.TopLeftCell.Row, 1).Value))
    If snapName <> "" Then Call RunSnapshotAnalysis(snapName)
    Exit Sub
ErrHandler:
    MsgBox "Could not determine which snapshot to run. Please try again.", vbExclamation
End Sub


Sub DeleteSnapshotFromButton()
    ' Called by every [Delete] FormControl button in the snapshot list.
    On Error GoTo ErrHandler
    Dim btn As Shape
    Set btn = ActiveSheet.Shapes(Application.Caller)
    Dim snapName As String
    snapName = Trim(CStr(ActiveSheet.Cells(btn.TopLeftCell.Row, 1).Value))
    If snapName = "" Then Exit Sub

    Dim confirm As Long
    confirm = MsgBox( _
        "Delete snapshot """ & snapName & """?" & vbCrLf & "This cannot be undone.", _
        vbYesNo + vbQuestion, "Delete Snapshot")
    If confirm = vbNo Then Exit Sub

    Dim wsSnap As Worksheet
    Set wsSnap = GetOrCreateSnapshotsSheet()
    Call DeleteSnapshotRows(wsSnap, snapName)
    Call RefreshPortfolioHistorySheet
    Exit Sub
ErrHandler:
    MsgBox "Could not determine which snapshot to delete. Please try again.", vbExclamation
End Sub


Sub GoToPortfolioHistory()
    Dim ws As Worksheet
    Set ws = GetPortfolioHistorySheet()
    ws.Activate
End Sub


Sub RefreshPortfolioHistorySheet()
    Dim wsHist As Worksheet
    Set wsHist = GetPortfolioHistorySheet()

    Dim wsSnap As Worksheet
    Set wsSnap = GetOrCreateSnapshotsSheet()

    Dim snapshots As Variant
    snapshots = ListAllSnapshots(wsSnap)

    Call WriteSnapshotListSection(wsHist, snapshots)
End Sub


' =====================================================================
' STORAGE HELPERS
' =====================================================================

Function GetOrCreateSnapshotsSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SNAP_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SNAP_SHEET_NAME
        ws.Cells(1, SN_ID).Value     = "SnapshotID"
        ws.Cells(1, SN_NAME).Value   = "SnapshotName"
        ws.Cells(1, SN_DATE).Value   = "SavedDate"
        ws.Cells(1, SN_STRAT).Value  = "StrategyName"
        ws.Cells(1, SN_CONTR).Value  = "Contracts"
        ws.Cells(1, SN_SYMBOL).Value = "Symbol"
        ws.Cells(1, SN_SECTOR).Value = "Sector"
        ws.Rows(1).Font.Bold = True
    End If

    ws.Visible = xlSheetHidden
    Set GetOrCreateSnapshotsSheet = ws
End Function


Function GetPortfolioHistorySheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(HISTORY_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Insert after Portfolio sheet if it exists, else at the end
        Dim wsPort As Worksheet
        On Error Resume Next
        Set wsPort = ThisWorkbook.Sheets("Portfolio")
        On Error GoTo 0

        If Not wsPort Is Nothing Then
            Set ws = ThisWorkbook.Sheets.Add(After:=wsPort)
        Else
            Set ws = ThisWorkbook.Sheets.Add( _
                After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        End If
        ws.Name = HISTORY_SHEET_NAME
        Call InitialiseHistorySheetLayout(ws)
    End If

    Set GetPortfolioHistorySheet = ws
End Function


Private Sub InitialiseHistorySheetLayout(ws As Worksheet)
    Application.ScreenUpdating = False
    ws.Cells.Clear

    ' Title bar
    With ws.Cells(H_TITLE_ROW, 1)
        .Value = "Portfolio History  —  Saved Portfolio Snapshots"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
    End With
    ws.Range(ws.Cells(H_TITLE_ROW, 1), ws.Cells(H_TITLE_ROW, 8)).Merge

    ' Instruction text (row 2)
    With ws.Cells(H_BTN_ROW, 1)
        .Value = "Save a snapshot of the current portfolio at any time, then Run to see historical performance."
        .Font.Italic = True
        .Font.Color = RGB(80, 80, 80)
    End With
    ws.Range(ws.Cells(H_BTN_ROW, 1), ws.Cells(H_BTN_ROW, 5)).Merge

    ' Save button (row 2, col 7)
    Dim saveBtn As Shape
    Dim btnLeft As Double:  btnLeft = ws.Columns(7).Left
    Dim btnTop As Double:   btnTop  = ws.Rows(H_BTN_ROW).Top + 1
    Set saveBtn = ws.Shapes.AddFormControl(xlButtonControl, btnLeft, btnTop, 160, 20)
    saveBtn.TextFrame.Characters.Text = "Save Current Portfolio"
    saveBtn.OnAction = "SaveCurrentPortfolioSnapshot"
    saveBtn.Name = "SaveSnapshotBtn"

    ' Row 3: blank spacer (nothing to write)

    ' Column headers for snapshot list (row 4)
    With ws.Rows(H_HDR_ROW)
        .Interior.Color = RGB(68, 114, 196)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .RowHeight = 18
    End With
    ws.Cells(H_HDR_ROW, 1).Value = "Snapshot Name"
    ws.Cells(H_HDR_ROW, 2).Value = "Saved Date"
    ws.Cells(H_HDR_ROW, 3).Value = "Strategies"
    ws.Cells(H_HDR_ROW, 4).Value = "Run"
    ws.Cells(H_HDR_ROW, 5).Value = "Delete"

    ' Column widths
    ws.Columns(1).ColumnWidth = 38
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 12
    ws.Columns(4).ColumnWidth = 10
    ws.Columns(5).ColumnWidth = 10
    ws.Columns(6).ColumnWidth = 20
    ws.Columns(7).ColumnWidth = 20
    ws.Columns(8).ColumnWidth = 18

    ' Row heights
    ws.Rows(H_TITLE_ROW).RowHeight = 26
    ws.Rows(H_BTN_ROW).RowHeight = 24

    Application.ScreenUpdating = True
End Sub


Function NextSnapshotID(wsSnap As Worksheet) As Long
    Dim lastRow As Long
    lastRow = wsSnap.Cells(wsSnap.Rows.Count, SN_ID).End(xlUp).Row
    If lastRow <= 1 Then
        NextSnapshotID = 1
        Exit Function
    End If

    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lastRow
        Dim v As Long
        On Error Resume Next
        v = CLng(wsSnap.Cells(i, SN_ID).Value)
        On Error GoTo 0
        If v > maxID Then maxID = v
    Next i
    NextSnapshotID = maxID + 1
End Function


Function ListAllSnapshots(wsSnap As Worksheet) As Variant
    ' Returns a 2D Variant array (1-based, rows = snapshots, cols = 3):
    '   col 1 = SnapshotName, col 2 = SavedDate, col 3 = StrategyCount
    ' Ordered by SnapshotID ascending. Returns Array() if none found.
    Dim lastRow As Long
    lastRow = wsSnap.Cells(wsSnap.Rows.Count, SN_ID).End(xlUp).Row
    If lastRow < 2 Then
        ListAllSnapshots = Array()
        Exit Function
    End If

    Dim idArr()   As Long
    Dim nameArr() As String
    Dim dateArr() As Variant
    Dim cntArr()  As Long
    Dim snapCount As Long
    snapCount = 0

    Dim i As Long, j As Long
    Dim thisID As Long, thisName As String

    For i = 2 To lastRow
        On Error Resume Next
        thisID = CLng(wsSnap.Cells(i, SN_ID).Value)
        On Error GoTo 0
        thisName = CStr(wsSnap.Cells(i, SN_NAME).Value)

        Dim found As Boolean
        found = False
        For j = 1 To snapCount
            If idArr(j) = thisID Then
                cntArr(j) = cntArr(j) + 1
                found = True
                Exit For
            End If
        Next j

        If Not found Then
            snapCount = snapCount + 1
            ReDim Preserve idArr(1 To snapCount)
            ReDim Preserve nameArr(1 To snapCount)
            ReDim Preserve dateArr(1 To snapCount)
            ReDim Preserve cntArr(1 To snapCount)
            idArr(snapCount)  = thisID
            nameArr(snapCount) = thisName
            dateArr(snapCount) = wsSnap.Cells(i, SN_DATE).Value
            cntArr(snapCount) = 1
        End If
    Next i

    If snapCount = 0 Then
        ListAllSnapshots = Array()
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To snapCount, 1 To 3)
    For i = 1 To snapCount
        result(i, 1) = nameArr(i)
        result(i, 2) = dateArr(i)
        result(i, 3) = cntArr(i)
    Next i
    ListAllSnapshots = result
End Function


Function GetSnapshotRows(wsSnap As Worksheet, snapName As String) As Variant
    ' Returns a 2D array (4 rows × stratCount cols):
    '   row 1 = StrategyName, row 2 = Contracts, row 3 = Symbol, row 4 = Sector
    Dim lastRow As Long
    lastRow = wsSnap.Cells(wsSnap.Rows.Count, SN_ID).End(xlUp).Row

    Dim props() As Variant
    Dim rowCount As Long
    rowCount = 0

    Dim i As Long
    For i = 2 To lastRow
        If StrComp(CStr(wsSnap.Cells(i, SN_NAME).Value), snapName, vbTextCompare) = 0 Then
            rowCount = rowCount + 1
            ReDim Preserve props(1 To 4, 1 To rowCount)
            props(1, rowCount) = CStr(wsSnap.Cells(i, SN_STRAT).Value)
            props(2, rowCount) = CDbl(wsSnap.Cells(i, SN_CONTR).Value)
            props(3, rowCount) = CStr(wsSnap.Cells(i, SN_SYMBOL).Value)
            props(4, rowCount) = CStr(wsSnap.Cells(i, SN_SECTOR).Value)
        End If
    Next i

    If rowCount = 0 Then
        GetSnapshotRows = Array()
    Else
        GetSnapshotRows = props
    End If
End Function


Function SnapshotExists(wsSnap As Worksheet, snapName As String) As Boolean
    Dim lastRow As Long
    lastRow = wsSnap.Cells(wsSnap.Rows.Count, SN_NAME).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If StrComp(CStr(wsSnap.Cells(i, SN_NAME).Value), snapName, vbTextCompare) = 0 Then
            SnapshotExists = True
            Exit Function
        End If
    Next i
    SnapshotExists = False
End Function


Private Sub DeleteSnapshotRows(wsSnap As Worksheet, snapName As String)
    ' Delete bottom-up to preserve row indices during deletion
    Dim lastRow As Long
    lastRow = wsSnap.Cells(wsSnap.Rows.Count, SN_NAME).End(xlUp).Row
    Dim i As Long
    For i = lastRow To 2 Step -1
        If StrComp(CStr(wsSnap.Cells(i, SN_NAME).Value), snapName, vbTextCompare) = 0 Then
            wsSnap.Rows(i).Delete
        End If
    Next i
End Sub


' =====================================================================
' PERFORMANCE COMPUTATION
' =====================================================================

Sub RunSnapshotAnalysis(snapName As String)
    Call InitializeColumnConstantsManually
    Application.ScreenUpdating = False
    Application.StatusBar = "Portfolio History: loading """ & snapName & """..."

    On Error GoTo CleanExit

    ' Load DailyM2MEquity
    Dim wsDM2M As Worksheet
    On Error Resume Next
    Set wsDM2M = ThisWorkbook.Sheets("DailyM2MEquity")
    On Error GoTo CleanExit
    If wsDM2M Is Nothing Then
        MsgBox "DailyM2MEquity sheet not found. Please run data import first.", _
               vbExclamation, "Portfolio History"
        GoTo CleanExit
    End If

    ' Load snapshot strategy rows
    Dim wsSnap As Worksheet
    Set wsSnap = GetOrCreateSnapshotsSheet()
    Dim snapRows As Variant
    snapRows = GetSnapshotRows(wsSnap, snapName)

    If Not IsArray(snapRows) Then
        MsgBox "Snapshot """ & snapName & """ has no strategies.", vbExclamation
        GoTo CleanExit
    End If
    Dim hasData As Boolean
    hasData = False
    On Error Resume Next
    hasData = (UBound(snapRows, 2) >= 1)
    On Error GoTo CleanExit
    If Not hasData Then
        MsgBox "Snapshot """ & snapName & """ has no strategies.", vbExclamation
        GoTo CleanExit
    End If

    Dim stratCount As Long
    stratCount = UBound(snapRows, 2)

    ' Build case-insensitive column lookup from DailyM2MEquity header row
    Dim colDict As Object
    Set colDict = CreateObject("Scripting.Dictionary")
    Dim lastDMCol As Long
    lastDMCol = wsDM2M.Cells(1, wsDM2M.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 2 To lastDMCol
        Dim hdr As String
        hdr = Trim(CStr(wsDM2M.Cells(1, c).Value))
        If hdr <> "" Then colDict(LCase(hdr)) = c
    Next c

    ' Match each snapshot strategy to its DailyM2MEquity column
    Dim foundCols()      As Long
    Dim foundContracts() As Double
    Dim missingStrats()  As String
    Dim missingCount     As Long
    missingCount = 0
    ReDim foundCols(1 To stratCount)
    ReDim foundContracts(1 To stratCount)

    Dim k As Long
    For k = 1 To stratCount
        Dim sName As String
        sName = Trim(snapRows(1, k))
        foundCols(k)      = FindStrategyInDailyM2M(sName, colDict)
        foundContracts(k) = snapRows(2, k)
        If foundCols(k) = 0 Then
            missingCount = missingCount + 1
            ReDim Preserve missingStrats(1 To missingCount)
            missingStrats(missingCount) = sName
        End If
    Next k

    ' Load DailyM2MEquity into a Variant array for speed
    Dim lastDMRow As Long
    lastDMRow = EndRowByCutoffSimple(wsDM2M, 1)
    If lastDMRow < 2 Then
        MsgBox "DailyM2MEquity has no data rows.", vbExclamation
        GoTo CleanExit
    End If

    Application.StatusBar = "Portfolio History: computing performance for """ & snapName & """..."

    Dim dmData As Variant
    dmData = wsDM2M.Range(wsDM2M.Cells(2, 1), wsDM2M.Cells(lastDMRow, lastDMCol)).Value
    ' dmData(r, 1) = date; dmData(r, c) = daily P&L at 1 contract for sheet column c

    ' Core performance loop  — mirrors N_BackTest.bas logic
    Dim cumPnL As Double, peakPnL As Double, maxDD As Double, maxDDPct As Double
    Dim totalDailyProfit As Double
    Dim positiveDays As Long, totalActiveDays As Long
    Dim fromDate As Date, toDate As Date
    Dim stratCumPnL() As Double
    ReDim stratCumPnL(1 To stratCount)
    Dim firstDate As Boolean
    firstDate = True

    Dim r As Long
    For r = 1 To lastDMRow - 1   ' dmData is 1-based; row 1 = sheet row 2
        If Not IsDate(dmData(r, 1)) Then GoTo NextDMRow

        Dim rowDate As Date
        rowDate = CDate(dmData(r, 1))

        totalDailyProfit = 0
        For k = 1 To stratCount
            If foundCols(k) > 0 Then
                Dim raw As Double
                raw = 0
                On Error Resume Next
                raw = CDbl(dmData(r, foundCols(k)))   ' handles CVErr cells gracefully
                On Error GoTo CleanExit
                Dim contrib As Double
                contrib = raw * foundContracts(k)
                totalDailyProfit  = totalDailyProfit + contrib
                stratCumPnL(k)    = stratCumPnL(k) + contrib
            End If
        Next k

        cumPnL = cumPnL + totalDailyProfit
        If cumPnL > peakPnL Then peakPnL = cumPnL
        Dim currentDD As Double
        currentDD = peakPnL - cumPnL
        If currentDD > maxDD Then maxDD = currentDD
        If peakPnL > 0 Then
            Dim thisDDPct As Double
            thisDDPct = currentDD / peakPnL
            If thisDDPct > maxDDPct Then maxDDPct = thisDDPct
        End If

        totalActiveDays = totalActiveDays + 1
        If totalDailyProfit > 0 Then positiveDays = positiveDays + 1

        If firstDate Then fromDate = rowDate: firstDate = False
        toDate = rowDate
NextDMRow:
    Next r

    Dim winRate As Double
    If totalActiveDays > 0 Then winRate = CDbl(positiveDays) / CDbl(totalActiveDays)

    Dim avgDailyProfit As Double
    If totalActiveDays > 0 Then avgDailyProfit = cumPnL / totalActiveDays

    ' Write results to PortfolioHistory sheet
    Dim wsHist As Worksheet
    Set wsHist = GetPortfolioHistorySheet()

    Call WritePerformanceSection( _
        wsHist, snapName, snapRows, foundCols, foundContracts, _
        missingStrats, missingCount, stratCumPnL, _
        cumPnL, maxDD, maxDDPct, winRate, avgDailyProfit, _
        fromDate, toDate)

    wsHist.Activate
    wsHist.Range("A1").Select

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub


Function FindStrategyInDailyM2M(stratName As String, colDict As Object) As Long
    ' Returns the DailyM2MEquity sheet column index for stratName, or 0 if not found.
    ' Uses case-insensitive matching via the pre-built lower-case dictionary.
    Dim key As String
    key = LCase(Trim(stratName))
    If colDict.Exists(key) Then
        FindStrategyInDailyM2M = colDict(key)
    Else
        FindStrategyInDailyM2M = 0
    End If
End Function


' =====================================================================
' DISPLAY — SNAPSHOT LIST
' =====================================================================

Sub WriteSnapshotListSection(wsHist As Worksheet, snapshots As Variant)
    Application.ScreenUpdating = False

    ' Remove existing Run/Delete buttons (collect first, then delete to avoid iteration issues)
    Dim shapesToDel() As String
    Dim delCount As Long
    delCount = 0
    Dim shp As Shape
    For Each shp In wsHist.Shapes
        If shp.Name Like "RunBtn_*" Or shp.Name Like "DelBtn_*" Then
            delCount = delCount + 1
            ReDim Preserve shapesToDel(1 To delCount)
            shapesToDel(delCount) = shp.Name
        End If
    Next shp
    Dim si As Long
    For si = 1 To delCount
        On Error Resume Next
        wsHist.Shapes(shapesToDel(si)).Delete
        On Error GoTo 0
    Next si

    ' Clear old list area  (from H_LIST_START to the first large empty block)
    Dim lastUsedRow As Long
    lastUsedRow = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row
    Dim clearTo As Long
    clearTo = Application.Max(H_LIST_START + 100, lastUsedRow)
    wsHist.Range(wsHist.Cells(H_LIST_START, 1), wsHist.Cells(clearTo, 8)).Clear

    ' Empty state message
    Dim hasSnaps As Boolean
    hasSnaps = False
    Dim snapCount As Long
    snapCount = 0
    On Error Resume Next
    snapCount = UBound(snapshots, 1)
    If Err.Number = 0 And snapCount > 0 Then hasSnaps = True
    On Error GoTo 0

    If Not hasSnaps Then
        With wsHist.Cells(H_LIST_START, 1)
            .Value = "(No snapshots saved yet — click 'Save Current Portfolio' to create one.)"
            .Font.Italic = True
            .Font.Color = RGB(150, 150, 150)
        End With
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Write snapshot rows
    Dim listRow As Long
    Dim r As Long
    For r = 1 To snapCount
        listRow = H_LIST_START + r - 1

        ' Alternating row colour
        Dim rowClr As Long
        If r Mod 2 = 0 Then rowClr = RGB(235, 241, 250) Else rowClr = RGB(255, 255, 255)
        wsHist.Rows(listRow).Interior.Color = rowClr
        wsHist.Rows(listRow).RowHeight = 20

        wsHist.Cells(listRow, 1).Value = snapshots(r, 1)   ' Snapshot name
        wsHist.Cells(listRow, 2).Value = snapshots(r, 2)   ' Saved date
        wsHist.Cells(listRow, 2).NumberFormat = "dd/mm/yyyy hh:mm"
        wsHist.Cells(listRow, 3).Value = snapshots(r, 3)   ' Strategy count
        wsHist.Cells(listRow, 3).HorizontalAlignment = xlCenter

        ' [Run] button
        Dim rowTop As Double:  rowTop  = wsHist.Rows(listRow).Top
        Dim rowHt  As Double:  rowHt   = wsHist.Rows(listRow).Height
        Dim c4L    As Double:  c4L     = wsHist.Columns(4).Left
        Dim c5L    As Double:  c5L     = wsHist.Columns(5).Left
        Dim c4W    As Double:  c4W     = wsHist.Columns(4).Width
        Dim c5W    As Double:  c5W     = wsHist.Columns(5).Width

        Dim runBtn As Shape
        Set runBtn = wsHist.Shapes.AddFormControl( _
            xlButtonControl, c4L + 1, rowTop + 1, c4W - 2, rowHt - 2)
        runBtn.TextFrame.Characters.Text = "Run"
        runBtn.OnAction = "RunSnapshotFromButton"
        runBtn.Name = "RunBtn_" & r

        ' [Delete] button
        Dim delBtn As Shape
        Set delBtn = wsHist.Shapes.AddFormControl( _
            xlButtonControl, c5L + 1, rowTop + 1, c5W - 2, rowHt - 2)
        delBtn.TextFrame.Characters.Text = "Delete"
        delBtn.OnAction = "DeleteSnapshotFromButton"
        delBtn.Name = "DelBtn_" & r
    Next r

    ' Thin border around the table
    With wsHist.Range( _
            wsHist.Cells(H_HDR_ROW, 1), _
            wsHist.Cells(H_LIST_START + snapCount - 1, 5)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(180, 180, 180)
        .Weight = xlThin
    End With

    Application.ScreenUpdating = True
End Sub


' =====================================================================
' DISPLAY — PERFORMANCE RESULTS
' =====================================================================

Sub WritePerformanceSection( _
    wsHist As Worksheet, snapName As String, _
    snapRows As Variant, foundCols() As Long, foundContracts() As Double, _
    missingStrats() As String, missingCount As Long, stratCumPnL() As Double, _
    cumPnL As Double, maxDD As Double, maxDDPct As Double, _
    winRate As Double, avgDailyProfit As Double, _
    fromDate As Date, toDate As Date)

    Application.ScreenUpdating = False

    ' Find where the performance section should start (below list + 2 blank rows)
    Dim wsSnap As Worksheet
    Set wsSnap = GetOrCreateSnapshotsSheet()
    Dim snapshots As Variant
    snapshots = ListAllSnapshots(wsSnap)

    Dim snapCount As Long
    snapCount = 0
    On Error Resume Next
    snapCount = UBound(snapshots, 1)
    On Error GoTo 0

    Dim listBottom As Long
    If snapCount > 0 Then
        listBottom = H_LIST_START + snapCount - 1
    Else
        listBottom = H_LIST_START
    End If

    Dim perfStart As Long
    perfStart = listBottom + 3   ' 2 blank separator rows + section heading

    ' Clear old performance section
    Dim lastClr As Long
    lastClr = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row
    If lastClr < perfStart + 50 Then lastClr = perfStart + 200
    wsHist.Range(wsHist.Cells(perfStart, 1), wsHist.Cells(lastClr, 8)).Clear

    Dim r As Long
    r = perfStart

    ' ---- Section heading --------------------------------------------------------
    With wsHist.Cells(r, 1)
        .Value = "PERFORMANCE RESULTS  —  " & snapName
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
    wsHist.Range(wsHist.Cells(r, 1), wsHist.Cells(r, 8)).Merge
    wsHist.Rows(r).RowHeight = 22
    r = r + 1

    ' ---- Warning banner (missing strategies) ------------------------------------
    If missingCount > 0 Then
        Dim missingList As String
        Dim m As Long
        For m = 1 To missingCount
            missingList = missingList & missingStrats(m)
            If m < missingCount Then missingList = missingList & ";  "
        Next m
        With wsHist.Cells(r, 1)
            .Value = "[!]  " & missingCount & " strategy/strategies not found in current data" & _
                     " — EXCLUDED from calculations:  " & missingList
            .Font.Bold = True
            .Font.Color = RGB(180, 60, 0)
            .Interior.Color = RGB(255, 245, 190)
        End With
        wsHist.Range(wsHist.Cells(r, 1), wsHist.Cells(r, 8)).Merge
        wsHist.Rows(r).RowHeight = 18
        r = r + 1
    End If

    r = r + 1   ' blank spacer

    ' ---- Strategy table header --------------------------------------------------
    With wsHist.Rows(r)
        .Interior.Color = RGB(68, 114, 196)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .RowHeight = 18
    End With
    wsHist.Cells(r, 1).Value = "Strategy Name"
    wsHist.Cells(r, 2).Value = "Contracts"
    wsHist.Cells(r, 3).Value = "Status"
    wsHist.Cells(r, 4).Value = "Cumulative P&L"
    wsHist.Cells(r, 5).Value = "% of Portfolio Total"
    r = r + 1

    ' ---- Strategy detail rows ---------------------------------------------------
    Dim stratCount As Long
    stratCount = UBound(snapRows, 2)

    Dim k As Long
    For k = 1 To stratCount
        Dim sName As String
        sName = snapRows(1, k)

        Dim rowClr As Long
        If foundCols(k) = 0 Then
            rowClr = RGB(255, 210, 210)     ' light red  — missing
        ElseIf k Mod 2 = 0 Then
            rowClr = RGB(235, 241, 250)     ' light blue — found, even
        Else
            rowClr = RGB(255, 255, 255)     ' white      — found, odd
        End If

        wsHist.Rows(r).Interior.Color = rowClr
        wsHist.Rows(r).RowHeight = 18

        wsHist.Cells(r, 1).Value = sName
        wsHist.Cells(r, 2).Value = foundContracts(k)
        wsHist.Cells(r, 2).HorizontalAlignment = xlCenter

        If foundCols(k) = 0 Then
            With wsHist.Cells(r, 3)
                .Value = "[!] NOT FOUND"
                .Font.Bold = True
                .Font.Color = RGB(180, 0, 0)
            End With
            wsHist.Cells(r, 4).Value = ""
            wsHist.Cells(r, 5).Value = ""
        Else
            wsHist.Cells(r, 3).Value = "Found"
            wsHist.Cells(r, 3).Font.Color = RGB(0, 128, 0)
            wsHist.Cells(r, 4).Value = stratCumPnL(k)
            wsHist.Cells(r, 4).NumberFormat = "$#,##0"
            If cumPnL <> 0 Then
                wsHist.Cells(r, 5).Value = stratCumPnL(k) / cumPnL
                wsHist.Cells(r, 5).NumberFormat = "0.0%"
            End If
        End If
        r = r + 1
    Next k

    ' ---- Total row --------------------------------------------------------------
    With wsHist.Rows(r)
        .Interior.Color = RGB(210, 225, 245)
        .Font.Bold = True
        .RowHeight = 18
    End With
    wsHist.Cells(r, 1).Value = "PORTFOLIO TOTAL"
    wsHist.Cells(r, 4).Value = cumPnL
    wsHist.Cells(r, 4).NumberFormat = "$#,##0"
    wsHist.Cells(r, 5).Value = 1
    wsHist.Cells(r, 5).NumberFormat = "0.0%"
    r = r + 2

    ' ---- Portfolio metrics block ------------------------------------------------
    With wsHist.Cells(r, 1)
        .Value = "Snapshot Metrics"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
    End With
    wsHist.Range(wsHist.Cells(r, 1), wsHist.Cells(r, 2)).Merge
    wsHist.Rows(r).RowHeight = 18
    r = r + 1

    Dim metricsLabel(1 To 6) As String
    Dim metricsValue(1 To 6) As String
    metricsLabel(1) = "Date Range"
    metricsValue(1) = Format(fromDate, "dd/mm/yyyy") & "  to  " & Format(toDate, "dd/mm/yyyy")
    metricsLabel(2) = "Cumulative P&L"
    metricsValue(2) = Format(cumPnL, "$#,##0")
    metricsLabel(3) = "Max Drawdown ($)"
    metricsValue(3) = Format(maxDD, "$#,##0")
    metricsLabel(4) = "Max Drawdown (%)"
    metricsValue(4) = Format(maxDDPct, "0.0%")
    metricsLabel(5) = "Win Rate (by day)"
    metricsValue(5) = Format(winRate, "0.0%")
    metricsLabel(6) = "Avg Daily P&L"
    metricsValue(6) = Format(avgDailyProfit, "$#,##0")

    Dim mi As Long
    For mi = 1 To 6
        wsHist.Cells(r, 1).Value = metricsLabel(mi)
        wsHist.Cells(r, 1).Font.Bold = True
        wsHist.Cells(r, 2).Value = metricsValue(mi)
        If mi Mod 2 = 0 Then wsHist.Rows(r).Interior.Color = RGB(235, 241, 250)
        wsHist.Rows(r).RowHeight = 17
        r = r + 1
    Next mi

    r = r + 1

    ' ---- Current portfolio comparison over same period --------------------------
    Call WriteCurrentPortfolioComparison(wsHist, r, fromDate, toDate)

    Application.ScreenUpdating = True
End Sub


Sub WriteCurrentPortfolioComparison( _
    wsHist As Worksheet, startRow As Long, fromDate As Date, toDate As Date)

    Dim wsTotalM2M As Worksheet
    On Error Resume Next
    Set wsTotalM2M = ThisWorkbook.Sheets("TotalPortfolioM2M")
    On Error GoTo 0

    Dim r As Long
    r = startRow

    ' Section heading
    With wsHist.Cells(r, 1)
        .Value = "COMPARISON  —  Current Portfolio (same period: " & _
                 Format(fromDate, "dd/mm/yyyy") & " to " & Format(toDate, "dd/mm/yyyy") & ")"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(56, 107, 44)   ' dark green
    End With
    wsHist.Range(wsHist.Cells(r, 1), wsHist.Cells(r, 8)).Merge
    wsHist.Rows(r).RowHeight = 22
    r = r + 1

    If wsTotalM2M Is Nothing Then
        wsHist.Cells(r, 1).Value = _
            "(TotalPortfolioM2M not available — run 'Create Portfolio Summary' first.)"
        wsHist.Cells(r, 1).Font.Italic = True
        wsHist.Cells(r, 1).Font.Color = RGB(150, 150, 150)
        Exit Sub
    End If

    ' Read TotalPortfolioM2M — col 1 = Date, col 2 = TotalDailyProfit
    Dim lastTMRow As Long
    lastTMRow = wsTotalM2M.Cells(wsTotalM2M.Rows.Count, 1).End(xlUp).Row
    If lastTMRow < 2 Then
        wsHist.Cells(r, 1).Value = "(TotalPortfolioM2M has no data.)"
        wsHist.Cells(r, 1).Font.Italic = True
        Exit Sub
    End If

    Dim currCumPnL As Double, currPeak As Double
    Dim currMaxDD As Double, currMaxDDPct As Double
    Dim currPosDays As Long, currTotDays As Long

    Dim i As Long
    For i = 2 To lastTMRow
        If Not IsDate(wsTotalM2M.Cells(i, 1).Value) Then GoTo NextTMRow
        Dim d As Date
        d = CDate(wsTotalM2M.Cells(i, 1).Value)
        If d < fromDate Or d > toDate Then GoTo NextTMRow

        Dim dayProfit As Double
        On Error Resume Next
        dayProfit = CDbl(wsTotalM2M.Cells(i, 2).Value)
        On Error GoTo 0
        currCumPnL = currCumPnL + dayProfit
        If currCumPnL > currPeak Then currPeak = currCumPnL
        Dim dd As Double
        dd = currPeak - currCumPnL
        If dd > currMaxDD Then currMaxDD = dd
        If currPeak > 0 Then
            Dim ddPct As Double
            ddPct = dd / currPeak
            If ddPct > currMaxDDPct Then currMaxDDPct = ddPct
        End If
        currTotDays = currTotDays + 1
        If dayProfit > 0 Then currPosDays = currPosDays + 1
NextTMRow:
    Next i

    Dim currWinRate As Double
    If currTotDays > 0 Then currWinRate = CDbl(currPosDays) / CDbl(currTotDays)
    Dim currAvgDay As Double
    If currTotDays > 0 Then currAvgDay = currCumPnL / currTotDays

    ' Metrics block header
    With wsHist.Cells(r, 1)
        .Value = "Current Portfolio Metrics"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(84, 130, 53)  ' sage green
    End With
    wsHist.Range(wsHist.Cells(r, 1), wsHist.Cells(r, 2)).Merge
    wsHist.Rows(r).RowHeight = 18
    r = r + 1

    Dim cLabel(1 To 5) As String
    Dim cValue(1 To 5) As String
    cLabel(1) = "Cumulative P&L"
    cValue(1) = Format(currCumPnL, "$#,##0")
    cLabel(2) = "Max Drawdown ($)"
    cValue(2) = Format(currMaxDD, "$#,##0")
    cLabel(3) = "Max Drawdown (%)"
    cValue(3) = Format(currMaxDDPct, "0.0%")
    cLabel(4) = "Win Rate (by day)"
    cValue(4) = Format(currWinRate, "0.0%")
    cLabel(5) = "Avg Daily P&L"
    cValue(5) = Format(currAvgDay, "$#,##0")

    Dim ci As Long
    For ci = 1 To 5
        wsHist.Cells(r, 1).Value = cLabel(ci)
        wsHist.Cells(r, 1).Font.Bold = True
        wsHist.Cells(r, 2).Value = cValue(ci)
        If ci Mod 2 = 0 Then wsHist.Rows(r).Interior.Color = RGB(220, 236, 217)
        wsHist.Rows(r).RowHeight = 17
        r = r + 1
    Next ci
End Sub

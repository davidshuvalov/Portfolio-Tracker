Attribute VB_Name = "W_Markets"
Option Explicit

' ============================================================
' W_Markets — Market deep-dive analysis using Buy & Hold ATR data
'
' Entry point : CreateMarketsSummary()
'
' Sheets created:
'   Markets            — hub: market ATR summary + sector summary
'   MarketCorrelations — pairwise ATR correlation matrix (heat-map)
'   MarketVolatility   — rolling 90-day ATR time series per contract
'
' Data sources:
'   AverageTrueRange   — pre-computed ATR averages written by D_Import_Data
'   TrueRanges         — full ATR history (exit-trade MFE-MAE) written by D_Import_Data
'   Summary            — sector & symbol per strategy (for labelling)
' ============================================================

Sub CreateMarketsSummary()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrorHandler

    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license. The tool will not function without a valid license.", vbCritical
        GoTo CleanExit
    End If

    Call InitializeColumnConstantsManually

    ' ---- Load source sheets ----------------------------------------
    Dim wsATR As Worksheet, wsTR As Worksheet, wsSummary As Worksheet
    On Error Resume Next
    Set wsATR = ThisWorkbook.Sheets("AverageTrueRange")
    Set wsTR = ThisWorkbook.Sheets("TrueRanges")
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0

    If wsATR Is Nothing Or wsTR Is Nothing Then
        MsgBox "Market data not found. Please run data import with Buy & Hold strategies first.", vbExclamation
        GoTo CleanExit
    End If

    ' ---- Read AverageTrueRange: one row per contract ---------------
    ' Layout: col 1 = Contract, cols 2-8 = ATR for 1M/3M/6M/12M/24M/60M/All
    Dim lastATRRow As Long
    lastATRRow = wsATR.Cells(wsATR.Rows.Count, 1).End(xlUp).Row
    Dim contractCount As Long: contractCount = lastATRRow - 1

    If contractCount <= 0 Then
        MsgBox "No market data found in AverageTrueRange. Please re-run data import.", vbExclamation
        GoTo CleanExit
    End If

    Dim contracts() As String
    Dim atr1M() As Double, atr3M() As Double, atr6M() As Double
    Dim atr12M() As Double, atr24M() As Double, atr60M() As Double, atrAll() As Double
    ReDim contracts(1 To contractCount)
    ReDim atr1M(1 To contractCount)
    ReDim atr3M(1 To contractCount)
    ReDim atr6M(1 To contractCount)
    ReDim atr12M(1 To contractCount)
    ReDim atr24M(1 To contractCount)
    ReDim atr60M(1 To contractCount)
    ReDim atrAll(1 To contractCount)

    Dim i As Long
    For i = 1 To contractCount
        contracts(i) = CStr(wsATR.Cells(i + 1, 1).Value)
        atr1M(i) = Val(CStr(wsATR.Cells(i + 1, 2).Value))
        atr3M(i) = Val(CStr(wsATR.Cells(i + 1, 3).Value))
        atr6M(i) = Val(CStr(wsATR.Cells(i + 1, 4).Value))
        atr12M(i) = Val(CStr(wsATR.Cells(i + 1, 5).Value))
        atr24M(i) = Val(CStr(wsATR.Cells(i + 1, 6).Value))
        atr60M(i) = Val(CStr(wsATR.Cells(i + 1, 7).Value))
        atrAll(i) = Val(CStr(wsATR.Cells(i + 1, 8).Value))
    Next i

    ' ---- Read TrueRanges into memory for percentile & correlation --
    ' Layout: col 1 = Date, col 2+ = ATR per contract (exit-trade MFE-MAE)
    Dim lastTRRow As Long, lastTRCol As Long
    lastTRRow = EndRowByCutoffSimple(wsTR, 1)
    lastTRCol = wsTR.Cells(1, wsTR.Columns.Count).End(xlToLeft).Column
    Dim trRowCount As Long: trRowCount = lastTRRow - 1

    If trRowCount < 1 Then
        MsgBox "TrueRanges sheet has no data rows.", vbExclamation
        GoTo CleanExit
    End If

    Dim trData As Variant
    trData = wsTR.Range(wsTR.Cells(2, 1), wsTR.Cells(lastTRRow, lastTRCol)).Value
    ' trData(row, 1) = date  |  trData(row, j) = ATR for sheet col j (j >= 2)

    ' Map each contract to its column index in trData (same as sheet column)
    Dim trColMap() As Long
    ReDim trColMap(1 To contractCount)
    Dim j As Long
    For i = 1 To contractCount
        trColMap(i) = 0
        For j = 2 To lastTRCol
            If CStr(wsTR.Cells(1, j).Value) = contracts(i) Then
                trColMap(i) = j   ' trData(row, j) maps directly to sheet col j
                Exit For
            End If
        Next j
    Next i

    ' ---- Compute ATR percentile, trend, and volatility regime ------
    Dim atrPct() As Double
    Dim atrTrend() As String
    Dim volRegime() As String
    ReDim atrPct(1 To contractCount)
    ReDim atrTrend(1 To contractCount)
    ReDim volRegime(1 To contractCount)

    Dim col As Long
    Dim v As Double
    Dim belowCount As Long, totalCount As Long
    Dim ratio As Double
    Dim rollAvg As Double, rollSum As Double, rollN As Long
    Dim k As Long, dateJ As Date, dateK As Date

    For i = 1 To contractCount
        ' Trend: current 3M ATR relative to 12M average
        If atr12M(i) > 0 Then
            ratio = atr3M(i) / atr12M(i)
            If ratio > 1.1 Then
                atrTrend(i) = "Rising"
            ElseIf ratio < 0.9 Then
                atrTrend(i) = "Falling"
            Else
                atrTrend(i) = "Stable"
            End If
        Else
            atrTrend(i) = "N/A"
        End If

        ' Percentile: rank of current 90-day avg (atr3M) among all historical 90-day rolling averages.
        ' This matches the calendar-based window used by CalculateAverageATR (dateCutoff = 90 days).
        col = trColMap(i)
        If col > 0 Then
            belowCount = 0
            totalCount = 0
            For j = 1 To trRowCount
                If IsDate(trData(j, 1)) Then
                    dateJ = CDate(trData(j, 1))
                    rollSum = 0: rollN = 0
                    For k = j To 1 Step -1
                        If IsDate(trData(k, 1)) Then
                            dateK = CDate(trData(k, 1))
                            If DateDiff("d", dateK, dateJ) > 90 Then Exit For
                            If trData(k, col) > 0 Then
                                rollSum = rollSum + trData(k, col)
                                rollN = rollN + 1
                            End If
                        End If
                    Next k
                    If rollN > 0 Then
                        rollAvg = rollSum / rollN
                        totalCount = totalCount + 1
                        If rollAvg <= atr3M(i) Then belowCount = belowCount + 1
                    End If
                End If
            Next j
            If totalCount > 0 Then
                atrPct(i) = CDbl(belowCount) / CDbl(totalCount) * 100
            End If
        End If

        ' Regime bucket
        Select Case True
            Case atrPct(i) >= 66: volRegime(i) = "High"
            Case atrPct(i) >= 33: volRegime(i) = "Normal"
            Case Else:             volRegime(i) = "Low"
        End Select
    Next i

    ' ---- Sector and strategy counts from Summary sheet -------------
    Dim sectors() As String
    Dim stratCounts() As Long
    ReDim sectors(1 To contractCount)
    ReDim stratCounts(1 To contractCount)

    If Not wsSummary Is Nothing Then
        Dim lastSumRow As Long
        lastSumRow = wsSummary.Cells(wsSummary.Rows.Count, COL_SYMBOL).End(xlUp).Row
        Dim sym As String, sec As String
        For j = 2 To lastSumRow
            sym = CStr(wsSummary.Cells(j, COL_SYMBOL).Value)
            sec = CStr(wsSummary.Cells(j, COL_SECTOR).Value)
            If sym = "" Then GoTo NextSumRow
            For i = 1 To contractCount
                ' Exact case-insensitive match only — avoids false positives (e.g. "ES" matching "ESET")
                If StrComp(sym, contracts(i), vbTextCompare) = 0 Then
                    stratCounts(i) = stratCounts(i) + 1
                    If sectors(i) = "" And sec <> "" Then sectors(i) = sec
                    Exit For
                End If
            Next i
NextSumRow:
        Next j
    End If

    ' ---- Build sector groupings ------------------------------------
    Dim uniqueSectors() As String
    Dim sectorCount As Long: sectorCount = 0
    Dim thisSec As String, found As Boolean

    For i = 1 To contractCount
        thisSec = sectors(i)
        If thisSec = "" Then thisSec = "Unknown"
        found = False
        For j = 1 To sectorCount
            If uniqueSectors(j) = thisSec Then found = True: Exit For
        Next j
        If Not found Then
            sectorCount = sectorCount + 1
            ReDim Preserve uniqueSectors(1 To sectorCount)
            uniqueSectors(sectorCount) = thisSec
        End If
    Next i

    ' Aggregate per sector
    Dim secMarkets() As String
    Dim secMarkCount() As Long
    Dim secStratCount() As Long
    Dim secSumPct() As Double
    Dim secSumATR3M() As Double
    Dim secSumATR12M() As Double
    ReDim secMarkets(1 To sectorCount)
    ReDim secMarkCount(1 To sectorCount)
    ReDim secStratCount(1 To sectorCount)
    ReDim secSumPct(1 To sectorCount)
    ReDim secSumATR3M(1 To sectorCount)
    ReDim secSumATR12M(1 To sectorCount)

    For i = 1 To contractCount
        thisSec = sectors(i)
        If thisSec = "" Then thisSec = "Unknown"
        For j = 1 To sectorCount
            If uniqueSectors(j) = thisSec Then
                secMarkCount(j) = secMarkCount(j) + 1
                secStratCount(j) = secStratCount(j) + stratCounts(i)
                secSumPct(j) = secSumPct(j) + atrPct(i)
                secSumATR3M(j) = secSumATR3M(j) + atr3M(i)
                secSumATR12M(j) = secSumATR12M(j) + atr12M(i)
                If secMarkets(j) = "" Then
                    secMarkets(j) = contracts(i)
                Else
                    secMarkets(j) = secMarkets(j) & ", " & contracts(i)
                End If
                Exit For
            End If
        Next j
    Next i

    Dim secAvgPct() As Double, secAvgATR3M() As Double, secAvgATR12M() As Double
    ReDim secAvgPct(1 To sectorCount)
    ReDim secAvgATR3M(1 To sectorCount)
    ReDim secAvgATR12M(1 To sectorCount)
    For j = 1 To sectorCount
        If secMarkCount(j) > 0 Then
            secAvgPct(j) = secSumPct(j) / secMarkCount(j)
            secAvgATR3M(j) = secSumATR3M(j) / secMarkCount(j)
            secAvgATR12M(j) = secSumATR12M(j) / secMarkCount(j)
        End If
    Next j

    ' ---- Write Markets hub sheet -----------------------------------
    Application.StatusBar = "Creating Markets sheet..."
    Dim wsMarkets As Worksheet
    Call RecreateMarketSheet("Markets", wsMarkets)
    wsMarkets.Tab.Color = RGB(0, 150, 60)

    Dim dataRow As Long, c As Long

    With wsMarkets
        ' ---- Header ------------------------------------------------
        .Cells(1, 1).Value = "Markets Overview"
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Last Updated: " & Format(Now(), "dd-mmm-yyyy")
        .Cells(2, 1).Font.Italic = True

        ' ---- Market ATR Summary table ------------------------------
        .Cells(4, 1).Value = "MARKET ATR SUMMARY"
        .Cells(4, 1).Font.Bold = True
        .Cells(4, 1).Font.Size = 12

        Dim mHdrRow As Long: mHdrRow = 5
        Dim mHeaders As Variant
        mHeaders = Array("Market", "Sector", "Strategies", _
                         "ATR 1M", "ATR 3M", "ATR 6M", "ATR 12M", "ATR 24M", "ATR 60M", "ATR All Time", _
                         "3M/12M Trend", "ATR Percentile", "Volatility Regime")

        For c = 0 To UBound(mHeaders)
            With .Cells(mHdrRow, c + 1)
                .Value = mHeaders(c)
                .Font.Bold = True
                .Interior.Color = RGB(0, 70, 127)
                .Font.Color = RGB(255, 255, 255)
            End With
        Next c

        For i = 1 To contractCount
            dataRow = mHdrRow + i

            ' Alternate row background first (cols 1-10)
            If i Mod 2 = 0 Then
                .Range(.Cells(dataRow, 1), .Cells(dataRow, 10)).Interior.Color = RGB(242, 242, 242)
            End If

            .Cells(dataRow, 1).Value = contracts(i)
            .Cells(dataRow, 2).Value = IIf(sectors(i) <> "", sectors(i), "Unknown")
            .Cells(dataRow, 3).Value = stratCounts(i)
            .Cells(dataRow, 4).Value = atr1M(i)
            .Cells(dataRow, 5).Value = atr3M(i)
            .Cells(dataRow, 6).Value = atr6M(i)
            .Cells(dataRow, 7).Value = atr12M(i)
            .Cells(dataRow, 8).Value = atr24M(i)
            .Cells(dataRow, 9).Value = atr60M(i)
            .Cells(dataRow, 10).Value = atrAll(i)
            .Cells(dataRow, 11).Value = atrTrend(i)
            .Cells(dataRow, 12).Value = atrPct(i)
            .Cells(dataRow, 13).Value = volRegime(i)

            ' Trend colour (overrides alternating)
            Select Case atrTrend(i)
                Case "Rising":  .Cells(dataRow, 11).Interior.Color = RGB(255, 180, 180)
                Case "Falling": .Cells(dataRow, 11).Interior.Color = RGB(180, 220, 255)
                Case "Stable":  .Cells(dataRow, 11).Interior.Color = RGB(220, 255, 220)
            End Select

            ' Regime colour (overrides alternating)
            Select Case volRegime(i)
                Case "High":   .Cells(dataRow, 13).Interior.Color = RGB(255, 100, 100)
                Case "Normal": .Cells(dataRow, 13).Interior.Color = RGB(255, 255, 150)
                Case "Low":    .Cells(dataRow, 13).Interior.Color = RGB(150, 230, 150)
            End Select
        Next i

        ' Number formats
        .Range(.Cells(mHdrRow + 1, 4), .Cells(mHdrRow + contractCount, 10)).NumberFormat = "$#,##0"
        .Range(.Cells(mHdrRow + 1, 12), .Cells(mHdrRow + contractCount, 12)).NumberFormat = "0.0"

        ' ---- Sector Summary table ----------------------------------
        Dim sSumHdr As Long: sSumHdr = mHdrRow + contractCount + 3
        .Cells(sSumHdr - 1, 1).Value = "SECTOR SUMMARY"
        .Cells(sSumHdr - 1, 1).Font.Bold = True
        .Cells(sSumHdr - 1, 1).Font.Size = 12

        Dim sHeaders As Variant
        sHeaders = Array("Sector", "Markets", "Market Count", "Strategy Count", _
                         "Avg ATR 3M", "Avg ATR 12M", "Avg ATR Percentile", "Volatility Regime")
        For c = 0 To UBound(sHeaders)
            With .Cells(sSumHdr, c + 1)
                .Value = sHeaders(c)
                .Font.Bold = True
                .Interior.Color = RGB(0, 70, 127)
                .Font.Color = RGB(255, 255, 255)
            End With
        Next c

        Dim secRegime As String
        For j = 1 To sectorCount
            Dim secRow As Long: secRow = sSumHdr + j

            If j Mod 2 = 0 Then
                .Range(.Cells(secRow, 1), .Cells(secRow, 7)).Interior.Color = RGB(242, 242, 242)
            End If

            If secAvgPct(j) >= 66 Then
                secRegime = "High"
            ElseIf secAvgPct(j) >= 33 Then
                secRegime = "Normal"
            Else
                secRegime = "Low"
            End If

            .Cells(secRow, 1).Value = uniqueSectors(j)
            .Cells(secRow, 2).Value = secMarkets(j)
            .Cells(secRow, 3).Value = secMarkCount(j)
            .Cells(secRow, 4).Value = secStratCount(j)
            .Cells(secRow, 5).Value = secAvgATR3M(j)
            .Cells(secRow, 6).Value = secAvgATR12M(j)
            .Cells(secRow, 7).Value = secAvgPct(j)
            .Cells(secRow, 8).Value = secRegime

            Select Case secRegime
                Case "High":   .Cells(secRow, 8).Interior.Color = RGB(255, 100, 100)
                Case "Normal": .Cells(secRow, 8).Interior.Color = RGB(255, 255, 150)
                Case "Low":    .Cells(secRow, 8).Interior.Color = RGB(150, 230, 150)
            End Select
        Next j

        .Range(.Cells(sSumHdr + 1, 5), .Cells(sSumHdr + sectorCount, 6)).NumberFormat = "$#,##0"
        .Range(.Cells(sSumHdr + 1, 7), .Cells(sSumHdr + sectorCount, 7)).NumberFormat = "0.0"

        .Columns("A:M").AutoFit
        .Columns("B").ColumnWidth = 30   ' Markets list in sector table can be wide
    End With

    wsMarkets.Visible = xlSheetVisible

    ' ---- Create sub-sheets -----------------------------------------
    Application.StatusBar = "Creating Market Correlations..."
    Call CreateMarketCorrelations(wsTR, contracts, contractCount, trData, trRowCount, lastTRCol)

    Application.StatusBar = "Creating Market Volatility..."
    Call CreateMarketVolatility(wsTR, contracts, contractCount, trData, trRowCount, lastTRCol)

    Application.StatusBar = "Creating Market Seasonality..."
    Call CreateMarketSeasonality(wsTR, contracts, contractCount, trData, trRowCount, lastTRCol)

    Application.StatusBar = "Creating Market Regime Analysis..."
    Call CreateMarketRegimes(wsTR, contracts, contractCount, trData, trRowCount, lastTRCol)

    wsMarkets.Activate
    Application.StatusBar = False

    MsgBox "Markets analysis complete." & vbCrLf & _
           "Sheets created: Markets, MarketCorrelations, MarketVolatility, MarketSeasonality, MarketRegimes.", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Markets analysis error (line " & Erl & "): " & Err.Description, vbCritical
    Resume CleanExit
End Sub


' ============================================================
' ATR Correlation Matrix
' ============================================================
Private Sub CreateMarketCorrelations(wsTR As Worksheet, contracts() As String, _
    contractCount As Long, trData As Variant, trRowCount As Long, lastTRCol As Long)

    Dim wsCorr As Worksheet
    Call RecreateMarketSheet("MarketCorrelations", wsCorr)
    wsCorr.Tab.Color = RGB(0, 150, 60)

    ' Map contract → sheet column (= trData column, same index)
    Dim colIdx() As Long
    ReDim colIdx(1 To contractCount)
    Dim i As Long, j As Long
    For i = 1 To contractCount
        colIdx(i) = 0
        For j = 2 To lastTRCol
            If CStr(wsTR.Cells(1, j).Value) = contracts(i) Then
                colIdx(i) = j: Exit For
            End If
        Next j
    Next i

    With wsCorr
        .Cells(1, 1).Value = "Market ATR Correlations"
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Pearson correlation of exit-trade ATR series (MFE-MAE) per contract. " & _
                             "Only dates where both contracts have non-zero ATR are included."
        .Cells(2, 1).Font.Italic = True

        Dim hdrRow As Long: hdrRow = 4

        ' Row and column labels
        For i = 1 To contractCount
            With .Cells(hdrRow, i + 1)
                .Value = contracts(i)
                .Font.Bold = True
                .Interior.Color = RGB(0, 70, 127)
                .Font.Color = RGB(255, 255, 255)
                .Orientation = 45
            End With
            With .Cells(hdrRow + i, 1)
                .Value = contracts(i)
                .Font.Bold = True
                .Interior.Color = RGB(0, 70, 127)
                .Font.Color = RGB(255, 255, 255)
            End With
        Next i

        ' Compute upper triangle, mirror to lower
        Dim corr As Double
        Dim r As Long, g As Long, b As Long
        For i = 1 To contractCount
            For j = 1 To contractCount
                If i = j Then
                    .Cells(hdrRow + i, j + 1).Value = 1
                    .Cells(hdrRow + i, j + 1).Interior.Color = RGB(180, 180, 180)
                    .Cells(hdrRow + i, j + 1).NumberFormat = "0.00"
                ElseIf j > i Then
                    corr = PearsonATR(trData, trRowCount, colIdx(i), colIdx(j))
                    .Cells(hdrRow + i, j + 1).Value = corr
                    .Cells(hdrRow + j, i + 1).Value = corr

                    ' Colour: green = positive, red = negative, white = zero
                    Dim absC As Double: absC = Abs(corr)
                    If corr >= 0 Then
                        r = CLng(255 - absC * 80)
                        g = 255
                        b = CLng(255 - absC * 80)
                    Else
                        r = 255
                        g = CLng(255 - absC * 80)
                        b = CLng(255 - absC * 80)
                    End If
                    .Cells(hdrRow + i, j + 1).Interior.Color = RGB(r, g, b)
                    .Cells(hdrRow + j, i + 1).Interior.Color = RGB(r, g, b)
                    .Cells(hdrRow + i, j + 1).NumberFormat = "0.00"
                    .Cells(hdrRow + j, i + 1).NumberFormat = "0.00"
                End If
            Next j
        Next i

        ' Column sizing
        .Columns(1).AutoFit
        Dim k As Long
        For k = 2 To contractCount + 1
            .Columns(k).ColumnWidth = 7
        Next k
        .Rows(hdrRow).RowHeight = 55
    End With

    wsCorr.Visible = xlSheetVisible
End Sub


' ============================================================
' Rolling ATR time series — market volatility over time
' ============================================================
Private Sub CreateMarketVolatility(wsTR As Worksheet, contracts() As String, _
    contractCount As Long, trData As Variant, trRowCount As Long, lastTRCol As Long)

    Dim wsVol As Worksheet
    Call RecreateMarketSheet("MarketVolatility", wsVol)
    wsVol.Tab.Color = RGB(0, 150, 60)

    ' Map contract → trData column
    Dim colIdx() As Long
    ReDim colIdx(1 To contractCount)
    Dim i As Long, j As Long

    For i = 1 To contractCount
        colIdx(i) = 0
        For j = 2 To lastTRCol
            If CStr(wsTR.Cells(1, j).Value) = contracts(i) Then
                colIdx(i) = j: Exit For
            End If
        Next j
    Next i

    With wsVol
        .Cells(1, 1).Value = "Market Volatility — Rolling 90-Day ATR"
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Rolling 90-day average of exit-trade ATR (MFE-MAE) per contract. " & _
                             "Each row is a date on which at least one contract had an exit."
        .Cells(2, 1).Font.Italic = True

        ' Headers
        Dim hdrRow As Long: hdrRow = 4
        With .Cells(hdrRow, 1)
            .Value = "Date"
            .Font.Bold = True
            .Interior.Color = RGB(0, 70, 127)
            .Font.Color = RGB(255, 255, 255)
        End With
        For i = 1 To contractCount
            With .Cells(hdrRow, i + 1)
                .Value = contracts(i)
                .Font.Bold = True
                .Interior.Color = RGB(0, 70, 127)
                .Font.Color = RGB(255, 255, 255)
            End With
        Next i

        ' Rolling 90-day average: for each exit date k, look back 90 days
        ' and average all non-zero ATR values for each contract.
        ' Use a start pointer per-iteration to avoid redundant scanning
        ' (since dates in TrueRanges are sorted ascending).
        Dim windowDays As Long: windowDays = 90
        Dim outputRow As Long: outputRow = hdrRow + 1
        Dim currentDate As Date
        Dim startPtr As Long: startPtr = 1
        Dim total As Double, cnt As Long

        Dim k As Long
        For k = 1 To trRowCount
            If Not IsDate(trData(k, 1)) Then GoTo SkipVolRow

            currentDate = CDate(trData(k, 1))

            ' Advance startPtr past rows older than 90 days
            Do While startPtr < k
                If IsDate(trData(startPtr, 1)) Then
                    If DateDiff("d", CDate(trData(startPtr, 1)), currentDate) > windowDays Then
                        startPtr = startPtr + 1
                    Else
                        Exit Do
                    End If
                Else
                    startPtr = startPtr + 1
                End If
            Loop

            ' Check that at least one contract has data on this date
            Dim hasData As Boolean: hasData = False
            For i = 1 To contractCount
                If colIdx(i) > 0 Then
                    If trData(k, colIdx(i)) > 0 Then hasData = True: Exit For
                End If
            Next i
            If Not hasData Then GoTo SkipVolRow

            .Cells(outputRow, 1).Value = currentDate

            For i = 1 To contractCount
                If colIdx(i) = 0 Then
                    .Cells(outputRow, i + 1).Value = 0
                Else
                    total = 0: cnt = 0
                    Dim ptr As Long
                    For ptr = startPtr To k
                        If IsDate(trData(ptr, 1)) Then
                            Dim av As Double: av = trData(ptr, colIdx(i))
                            If av > 0 Then
                                total = total + av
                                cnt = cnt + 1
                            End If
                        End If
                    Next ptr
                    .Cells(outputRow, i + 1).Value = IIf(cnt > 0, total / cnt, 0)
                End If
            Next i

            outputRow = outputRow + 1
SkipVolRow:
        Next k

        ' Format
        If outputRow > hdrRow + 1 Then
            .Range(.Cells(hdrRow + 1, 1), .Cells(outputRow - 1, 1)).NumberFormat = "dd-mmm-yyyy"
            .Range(.Cells(hdrRow + 1, 2), _
                   .Cells(outputRow - 1, contractCount + 1)).NumberFormat = "$#,##0"
        End If
        .Columns(1).ColumnWidth = 13
        .Columns("B:" & Chr(65 + contractCount)).AutoFit
    End With

    wsVol.Visible = xlSheetVisible
End Sub


' ============================================================
' Pearson correlation of two ATR columns in trData
' Only rows where BOTH columns have a non-zero value are included.
' ============================================================
Private Function PearsonATR(data As Variant, rowCount As Long, _
    col1 As Long, col2 As Long) As Double

    If col1 = 0 Or col2 = 0 Then PearsonATR = 0: Exit Function

    Dim i As Long, n As Long
    Dim sum1 As Double, sum2 As Double
    Dim sum1Sq As Double, sum2Sq As Double, sumProd As Double
    Dim v1 As Double, v2 As Double

    For i = 1 To rowCount
        v1 = data(i, col1)
        v2 = data(i, col2)
        If v1 > 0 And v2 > 0 Then
            n = n + 1
            sum1 = sum1 + v1
            sum2 = sum2 + v2
            sum1Sq = sum1Sq + v1 * v1
            sum2Sq = sum2Sq + v2 * v2
            sumProd = sumProd + v1 * v2
        End If
    Next i

    If n < 2 Then PearsonATR = 0: Exit Function

    Dim denom As Double
    denom = Sqr((sum1Sq - sum1 * sum1 / n) * (sum2Sq - sum2 * sum2 / n))
    PearsonATR = IIf(denom = 0, 0, (sumProd - sum1 * sum2 / n) / denom)
End Function


' ============================================================
' Helper: create or clear a sheet, placing it after the last sheet
' ============================================================
Private Sub RecreateMarketSheet(sheetName As String, ByRef ws As Worksheet)
    Application.DisplayAlerts = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
        On Error Resume Next
        ws.DrawingObjects.Delete
        On Error GoTo 0
    End If
    Application.DisplayAlerts = True
End Sub


' ============================================================
' Calendar / Seasonality Analysis
' Creates MarketSeasonality sheet with monthly average ATR
' and a seasonality index (monthly avg / annual avg × 100).
' ============================================================
Private Sub CreateMarketSeasonality(wsTR As Worksheet, contracts() As String, _
    contractCount As Long, trData As Variant, trRowCount As Long, lastTRCol As Long)

    Dim wsS As Worksheet
    Call RecreateMarketSheet("MarketSeasonality", wsS)
    wsS.Tab.Color = RGB(0, 150, 60)

    Dim colIdx() As Long
    ReDim colIdx(1 To contractCount)
    Dim i As Long, j As Long
    For i = 1 To contractCount
        For j = 2 To lastTRCol
            If CStr(wsTR.Cells(1, j).Value) = contracts(i) Then
                colIdx(i) = j: Exit For
            End If
        Next j
    Next i

    ' Accumulate ATR sums per contract per month
    Dim monthSum() As Double, monthCnt() As Long
    ReDim monthSum(1 To contractCount, 1 To 12)
    ReDim monthCnt(1 To contractCount, 1 To 12)

    Dim k As Long, mo As Integer, atrVal As Double
    For k = 1 To trRowCount
        If IsDate(trData(k, 1)) Then
            mo = Month(CDate(trData(k, 1)))
            For i = 1 To contractCount
                If colIdx(i) > 0 Then
                    If IsNumeric(trData(k, colIdx(i))) Then
                        atrVal = CDbl(trData(k, colIdx(i)))
                        If atrVal > 0 Then
                            monthSum(i, mo) = monthSum(i, mo) + atrVal
                            monthCnt(i, mo) = monthCnt(i, mo) + 1
                        End If
                    End If
                End If
            Next i
        End If
    Next k

    Dim monthAvg() As Double, annualAvg() As Double
    ReDim monthAvg(1 To contractCount, 1 To 12)
    ReDim annualAvg(1 To contractCount)

    For i = 1 To contractCount
        Dim totSum As Double: totSum = 0
        Dim totCnt As Long: totCnt = 0
        For mo = 1 To 12
            If monthCnt(i, mo) > 0 Then
                monthAvg(i, mo) = monthSum(i, mo) / monthCnt(i, mo)
                totSum = totSum + monthSum(i, mo)
                totCnt = totCnt + monthCnt(i, mo)
            End If
        Next mo
        If totCnt > 0 Then annualAvg(i) = totSum / totCnt
    Next i

    Dim monthNames As Variant
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    Dim c As Long
    With wsS
        .Cells(1, 1).Value = "Market Seasonality — Monthly ATR Patterns"
        .Cells(1, 1).Font.Size = 14: .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Monthly average ATR (|MFE| + |MAE|) per contract.  Seasonality Index = (month avg / annual avg) x 100."
        .Cells(2, 1).Font.Italic = True

        Dim hdrRow As Long: hdrRow = 4
        ' Fixed headers
        With .Cells(hdrRow, 1)
            .Value = "Market": .Font.Bold = True
            .Interior.Color = RGB(0, 70, 127): .Font.Color = RGB(255, 255, 255)
        End With
        With .Cells(hdrRow, 2)
            .Value = "Annual Avg ATR": .Font.Bold = True
            .Interior.Color = RGB(0, 70, 127): .Font.Color = RGB(255, 255, 255)
        End With
        ' Monthly Avg columns (cols 3-14) and Index columns (cols 15-26)
        For c = 1 To 12
            With .Cells(hdrRow, c + 2)
                .Value = monthNames(c - 1) & " Avg"
                .Font.Bold = True: .Interior.Color = RGB(0, 70, 127): .Font.Color = RGB(255, 255, 255)
            End With
            With .Cells(hdrRow, c + 14)
                .Value = monthNames(c - 1) & " Index"
                .Font.Bold = True: .Interior.Color = RGB(100, 0, 127): .Font.Color = RGB(255, 255, 255)
            End With
        Next c

        Dim dataRow As Long: dataRow = hdrRow + 1
        Dim idx As Double
        For i = 1 To contractCount
            If i Mod 2 = 0 Then .Range(.Cells(dataRow, 1), .Cells(dataRow, 26)).Interior.Color = RGB(242, 242, 242)
            .Cells(dataRow, 1).Value = contracts(i)
            .Cells(dataRow, 2).Value = annualAvg(i)
            For c = 1 To 12
                .Cells(dataRow, c + 2).Value = monthAvg(i, c)
                If annualAvg(i) > 0 And monthAvg(i, c) > 0 Then
                    idx = monthAvg(i, c) / annualAvg(i) * 100
                Else
                    idx = 0
                End If
                .Cells(dataRow, c + 14).Value = idx
                ' Colour: red = high season (>= 120), blue = low season (<= 80)
                If idx >= 120 Then
                    .Cells(dataRow, c + 14).Interior.Color = RGB(255, 100, 100)
                ElseIf idx >= 110 Then
                    .Cells(dataRow, c + 14).Interior.Color = RGB(255, 180, 180)
                ElseIf idx > 0 And idx <= 80 Then
                    .Cells(dataRow, c + 14).Interior.Color = RGB(150, 200, 255)
                ElseIf idx > 0 And idx <= 90 Then
                    .Cells(dataRow, c + 14).Interior.Color = RGB(200, 230, 255)
                End If
            Next c
            dataRow = dataRow + 1
        Next i

        .Range(.Cells(hdrRow + 1, 2), .Cells(dataRow - 1, 14)).NumberFormat = "$#,##0"
        .Range(.Cells(hdrRow + 1, 15), .Cells(dataRow - 1, 26)).NumberFormat = "0.0"
        .Columns("A:Z").AutoFit
        .Columns(1).ColumnWidth = IIf(.Columns(1).ColumnWidth < 10, 10, .Columns(1).ColumnWidth)
    End With
    wsS.Visible = xlSheetVisible
End Sub


' ============================================================
' Regime Analysis
' Creates MarketRegimes sheet with vol × trend regime stats.
'
' Vol regime  (30-day ATR vs 90-day ATR):
'   High    — 30d > 90d × 1.15
'   Normal  — within 15%
'   Low     — 30d < 90d × 0.85
'
' Trend regime (same ratio, looser bands):
'   Rising  — 30d > 90d × 1.10  (expanding / trending)
'   Sideways — within 10%
'   Falling — 30d < 90d × 0.90  (contracting / mean-rev)
' ============================================================
Private Sub CreateMarketRegimes(wsTR As Worksheet, contracts() As String, _
    contractCount As Long, trData As Variant, trRowCount As Long, lastTRCol As Long)

    Dim wsR As Worksheet
    Call RecreateMarketSheet("MarketRegimes", wsR)
    wsR.Tab.Color = RGB(0, 150, 60)

    ' Map contracts to columns
    Dim colIdx() As Long
    ReDim colIdx(1 To contractCount)
    Dim i As Long, j As Long, k As Long
    For i = 1 To contractCount
        For j = 2 To lastTRCol
            If CStr(wsTR.Cells(1, j).Value) = contracts(i) Then
                colIdx(i) = j: Exit For
            End If
        Next j
    Next i

    ' Rolling sums for 30-day and 90-day windows per contract
    Dim rs30() As Double, rc30() As Long
    Dim rs90() As Double, rc90() As Long
    ReDim rs30(1 To contractCount): ReDim rc30(1 To contractCount)
    ReDim rs90(1 To contractCount): ReDim rc90(1 To contractCount)

    ' Regime counters: volCnt(contract, 0=Low/1=Normal/2=High)
    '                  trendCnt(contract, 0=Falling/1=Sideways/2=Rising)
    '                  vtCnt(contract, vol, trend) — combined 3×3 grid
    Dim volCnt() As Long, trendCnt() As Long, vtCnt() As Long
    ReDim volCnt(1 To contractCount, 0 To 2)
    ReDim trendCnt(1 To contractCount, 0 To 2)
    ReDim vtCnt(1 To contractCount, 0 To 2, 0 To 2)
    Dim totalDays() As Long
    ReDim totalDays(1 To contractCount)

    Dim sp30 As Long: sp30 = 1
    Dim sp90 As Long: sp90 = 1
    Dim curDate As Date, oldDate As Date
    Dim atrV As Double, avg30 As Double, avg90 As Double
    Dim vReg As Integer, tReg As Integer

    For k = 1 To trRowCount
        If Not IsDate(trData(k, 1)) Then GoTo NextRR
        curDate = CDate(trData(k, 1))

        ' Remove stale rows from 30-day window
        Do While sp30 < k
            If IsDate(trData(sp30, 1)) Then
                oldDate = CDate(trData(sp30, 1))
                If DateDiff("d", oldDate, curDate) > 30 Then
                    For i = 1 To contractCount
                        If colIdx(i) > 0 Then
                            If IsNumeric(trData(sp30, colIdx(i))) Then
                                atrV = CDbl(trData(sp30, colIdx(i)))
                                If atrV > 0 Then
                                    rs30(i) = rs30(i) - atrV
                                    If rc30(i) > 0 Then rc30(i) = rc30(i) - 1
                                End If
                            End If
                        End If
                    Next i
                    sp30 = sp30 + 1
                Else
                    Exit Do
                End If
            Else
                sp30 = sp30 + 1
            End If
        Loop

        ' Remove stale rows from 90-day window
        Do While sp90 < k
            If IsDate(trData(sp90, 1)) Then
                oldDate = CDate(trData(sp90, 1))
                If DateDiff("d", oldDate, curDate) > 90 Then
                    For i = 1 To contractCount
                        If colIdx(i) > 0 Then
                            If IsNumeric(trData(sp90, colIdx(i))) Then
                                atrV = CDbl(trData(sp90, colIdx(i)))
                                If atrV > 0 Then
                                    rs90(i) = rs90(i) - atrV
                                    If rc90(i) > 0 Then rc90(i) = rc90(i) - 1
                                End If
                            End If
                        End If
                    Next i
                    sp90 = sp90 + 1
                Else
                    Exit Do
                End If
            Else
                sp90 = sp90 + 1
            End If
        Loop

        ' Add current row to windows and classify
        For i = 1 To contractCount
            If colIdx(i) > 0 Then
                If IsNumeric(trData(k, colIdx(i))) Then
                    atrV = CDbl(trData(k, colIdx(i)))
                    If atrV > 0 Then
                        rs30(i) = rs30(i) + atrV: rc30(i) = rc30(i) + 1
                        rs90(i) = rs90(i) + atrV: rc90(i) = rc90(i) + 1
                    End If
                End If
                ' Classify once we have enough history (10 days min)
                If rc30(i) >= 5 And rc90(i) >= 10 Then
                    avg30 = rs30(i) / rc30(i)
                    avg90 = rs90(i) / rc90(i)
                    If avg90 = 0 Then GoTo NextCI

                    If avg30 >= avg90 * 1.15 Then
                        vReg = 2
                    ElseIf avg30 <= avg90 * 0.85 Then
                        vReg = 0
                    Else
                        vReg = 1
                    End If

                    If avg30 >= avg90 * 1.1 Then
                        tReg = 2
                    ElseIf avg30 <= avg90 * 0.9 Then
                        tReg = 0
                    Else
                        tReg = 1
                    End If

                    volCnt(i, vReg) = volCnt(i, vReg) + 1
                    trendCnt(i, tReg) = trendCnt(i, tReg) + 1
                    vtCnt(i, vReg, tReg) = vtCnt(i, vReg, tReg) + 1
                    totalDays(i) = totalDays(i) + 1
                End If
            End If
NextCI:
        Next i
NextRR:
    Next k

    Dim volLbls As Variant:   volLbls = Array("Low Vol", "Normal Vol", "High Vol")
    Dim trendLbls As Variant: trendLbls = Array("Falling (Contracting)", "Sideways", "Rising (Trending)")
    Dim c As Long, dataRow As Long, vr As Integer, tr As Integer

    With wsR
        .Cells(1, 1).Value = "Market Regime Analysis"
        .Cells(1, 1).Font.Size = 14: .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Regime classification using rolling 30-day vs 90-day ATR comparison per contract."
        .Cells(3, 1).Value = "Vol: High = 30d ATR > 90d x 1.15 | Normal = within 15% | Low = 30d ATR < 90d x 0.85"
        .Cells(4, 1).Value = "Trend: Rising = 30d > 90d x 1.10 (trending/expanding) | Sideways = within 10% | Falling = 30d < 90d x 0.90 (contracting)"
        For j = 2 To 4: .Cells(j, 1).Font.Italic = True: Next j

        ' --- Volatility Regime Summary ---
        Dim vHdr As Long: vHdr = 6
        .Cells(vHdr, 1).Value = "VOLATILITY REGIME DISTRIBUTION"
        .Cells(vHdr, 1).Font.Bold = True: .Cells(vHdr, 1).Font.Size = 12

        Dim vHdrs As Variant
        vHdrs = Array("Market", "Total Days", "Low Vol Days", "Low Vol %", "Normal Vol Days", "Normal Vol %", "High Vol Days", "High Vol %")
        For c = 0 To UBound(vHdrs)
            With .Cells(vHdr + 1, c + 1)
                .Value = vHdrs(c): .Font.Bold = True
                .Interior.Color = RGB(0, 70, 127): .Font.Color = RGB(255, 255, 255): .WrapText = True
            End With
        Next c

        dataRow = vHdr + 2
        For i = 1 To contractCount
            If totalDays(i) = 0 Then GoTo SkipVol
            If i Mod 2 = 0 Then .Range(.Cells(dataRow, 1), .Cells(dataRow, 8)).Interior.Color = RGB(242, 242, 242)
            .Cells(dataRow, 1).Value = contracts(i)
            .Cells(dataRow, 2).Value = totalDays(i)
            .Cells(dataRow, 3).Value = volCnt(i, 0)
            .Cells(dataRow, 4).Value = IIf(totalDays(i) > 0, volCnt(i, 0) / totalDays(i) * 100, 0)
            .Cells(dataRow, 5).Value = volCnt(i, 1)
            .Cells(dataRow, 6).Value = IIf(totalDays(i) > 0, volCnt(i, 1) / totalDays(i) * 100, 0)
            .Cells(dataRow, 7).Value = volCnt(i, 2)
            .Cells(dataRow, 8).Value = IIf(totalDays(i) > 0, volCnt(i, 2) / totalDays(i) * 100, 0)
            If totalDays(i) > 0 Then
                If volCnt(i, 2) / totalDays(i) > 0.33 Then .Cells(dataRow, 8).Interior.Color = RGB(255, 150, 150)
                If volCnt(i, 0) / totalDays(i) > 0.33 Then .Cells(dataRow, 4).Interior.Color = RGB(150, 200, 255)
            End If
            dataRow = dataRow + 1
SkipVol:
        Next i
        .Range(.Cells(vHdr + 2, 4), .Cells(dataRow - 1, 4)).NumberFormat = "0.0"
        .Range(.Cells(vHdr + 2, 6), .Cells(dataRow - 1, 6)).NumberFormat = "0.0"
        .Range(.Cells(vHdr + 2, 8), .Cells(dataRow - 1, 8)).NumberFormat = "0.0"

        ' --- Trend Regime Summary ---
        Dim tHdr As Long: tHdr = dataRow + 2
        .Cells(tHdr, 1).Value = "TREND / DIRECTION REGIME DISTRIBUTION"
        .Cells(tHdr, 1).Font.Bold = True: .Cells(tHdr, 1).Font.Size = 12

        Dim tHdrs As Variant
        tHdrs = Array("Market", "Total Days", "Falling Days", "Falling %", "Sideways Days", "Sideways %", "Rising Days", "Rising %")
        For c = 0 To UBound(tHdrs)
            With .Cells(tHdr + 1, c + 1)
                .Value = tHdrs(c): .Font.Bold = True
                .Interior.Color = RGB(0, 70, 127): .Font.Color = RGB(255, 255, 255): .WrapText = True
            End With
        Next c

        dataRow = tHdr + 2
        For i = 1 To contractCount
            If totalDays(i) = 0 Then GoTo SkipTrend
            If i Mod 2 = 0 Then .Range(.Cells(dataRow, 1), .Cells(dataRow, 8)).Interior.Color = RGB(242, 242, 242)
            .Cells(dataRow, 1).Value = contracts(i)
            .Cells(dataRow, 2).Value = totalDays(i)
            .Cells(dataRow, 3).Value = trendCnt(i, 0)
            .Cells(dataRow, 4).Value = IIf(totalDays(i) > 0, trendCnt(i, 0) / totalDays(i) * 100, 0)
            .Cells(dataRow, 5).Value = trendCnt(i, 1)
            .Cells(dataRow, 6).Value = IIf(totalDays(i) > 0, trendCnt(i, 1) / totalDays(i) * 100, 0)
            .Cells(dataRow, 7).Value = trendCnt(i, 2)
            .Cells(dataRow, 8).Value = IIf(totalDays(i) > 0, trendCnt(i, 2) / totalDays(i) * 100, 0)
            If totalDays(i) > 0 Then
                If trendCnt(i, 2) / totalDays(i) > 0.33 Then .Cells(dataRow, 8).Interior.Color = RGB(150, 230, 150)
                If trendCnt(i, 0) / totalDays(i) > 0.33 Then .Cells(dataRow, 4).Interior.Color = RGB(255, 180, 150)
            End If
            dataRow = dataRow + 1
SkipTrend:
        Next i
        .Range(.Cells(tHdr + 2, 4), .Cells(dataRow - 1, 4)).NumberFormat = "0.0"
        .Range(.Cells(tHdr + 2, 6), .Cells(dataRow - 1, 6)).NumberFormat = "0.0"
        .Range(.Cells(tHdr + 2, 8), .Cells(dataRow - 1, 8)).NumberFormat = "0.0"

        ' --- Combined Vol × Trend Grid (aggregate across all contracts) ---
        Dim gHdr As Long: gHdr = dataRow + 2
        .Cells(gHdr, 1).Value = "COMBINED REGIME GRID — % of days in each state (aggregate all contracts)"
        .Cells(gHdr, 1).Font.Bold = True: .Cells(gHdr, 1).Font.Size = 12

        .Cells(gHdr + 1, 1).Value = "Vol \ Trend"
        .Cells(gHdr + 1, 1).Font.Bold = True
        For tr = 0 To 2
            With .Cells(gHdr + 1, tr + 2)
                .Value = trendLbls(tr): .Font.Bold = True
                .Interior.Color = RGB(0, 70, 127): .Font.Color = RGB(255, 255, 255): .WrapText = True
            End With
        Next tr

        Dim totalAgg As Long: totalAgg = 0
        Dim aggGrid(0 To 2, 0 To 2) As Long
        For i = 1 To contractCount
            For vr = 0 To 2
                For tr = 0 To 2
                    aggGrid(vr, tr) = aggGrid(vr, tr) + vtCnt(i, vr, tr)
                    totalAgg = totalAgg + vtCnt(i, vr, tr)
                Next tr
            Next vr
        Next i

        Dim volColors As Variant: volColors = Array(RGB(150, 200, 255), RGB(200, 230, 200), RGB(255, 150, 150))
        Dim pct As Double
        For vr = 0 To 2
            With .Cells(gHdr + 2 + vr, 1)
                .Value = volLbls(vr): .Font.Bold = True
                .Interior.Color = volColors(vr)
            End With
            For tr = 0 To 2
                pct = IIf(totalAgg > 0, aggGrid(vr, tr) / totalAgg * 100, 0)
                With .Cells(gHdr + 2 + vr, tr + 2)
                    .Value = pct: .NumberFormat = "0.0"
                    If pct > 20 Then .Interior.Color = RGB(200, 255, 200)
                    If pct > 30 Then .Interior.Color = RGB(150, 240, 150)
                End With
            Next tr
        Next vr

        .Columns("A:J").AutoFit
        .Columns(1).ColumnWidth = IIf(.Columns(1).ColumnWidth < 14, 14, .Columns(1).ColumnWidth)
    End With
    wsR.Visible = xlSheetVisible
End Sub


' ============================================================
' Trade Analytics — MAE/MFE & Strategy Performance
' Creates TradeAnalytics sheet with:
'   Section 1: Strategy-level closed-trade statistics (from PortClosedTrade)
'   Section 2: Contract-level ATR efficiency (TrueRanges vs TradePNL)
'              Efficiency = Avg Winner PNL / Avg ATR × 100%
' ============================================================
Sub CreateTradeAnalytics()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrorHandler

    Call InitializeColumnConstantsManually

    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license.", vbCritical
        GoTo CleanExit
    End If

    ' Require at least one data source
    Dim hasClosed As Boolean, hasTR As Boolean, hasTP As Boolean
    Dim wsClosedTrade As Worksheet, wsTR As Worksheet, wsTP As Worksheet

    On Error Resume Next
    Set wsClosedTrade = ThisWorkbook.Sheets("PortClosedTrade")
    Set wsTR = ThisWorkbook.Sheets("TrueRanges")
    Set wsTP = ThisWorkbook.Sheets("TradePNL")
    On Error GoTo ErrorHandler
    hasClosed = Not wsClosedTrade Is Nothing
    hasTR = Not wsTR Is Nothing
    hasTP = Not wsTP Is Nothing

    If Not hasClosed And Not hasTR Then
        MsgBox "PortClosedTrade and TrueRanges sheets not found. Run portfolio setup and data import first.", vbExclamation
        GoTo CleanExit
    End If

    ' Make hidden sheets readable
    If hasClosed Then wsClosedTrade.Visible = xlSheetVisible
    If hasTR Then wsTR.Visible = xlSheetVisible
    If hasTP Then wsTP.Visible = xlSheetVisible

    ' Create output sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("TradeAnalytics").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    Dim wsTA As Worksheet
    Dim wsPort As Worksheet
    On Error Resume Next
    Set wsPort = ThisWorkbook.Sheets("Portfolio")
    On Error GoTo ErrorHandler

    If wsPort Is Nothing Then
        Set wsTA = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    Else
        Set wsTA = ThisWorkbook.Sheets.Add(After:=wsPort)
    End If
    wsTA.Name = "TradeAnalytics"
    wsTA.Tab.Color = RGB(180, 100, 0)

    With wsTA
        .Cells(1, 1).Value = "Trade Analytics — MAE/MFE & Strategy Performance"
        .Cells(1, 1).Font.Size = 16: .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Last updated: " & Format(Now(), "dd-mmm-yyyy")
        .Cells(3, 1).Value = "ATR = |MFE| + |MAE| (daily exit-trade range).  Efficiency % = Avg Winner PNL / Avg ATR x 100 (range captured on winning days)."
        .Cells(2, 1).Font.Italic = True: .Cells(3, 1).Font.Italic = True
    End With

    Dim nextRow As Long: nextRow = 5
    Dim c As Long, dataRow As Long

    ' ================================================================
    ' Section 1 — Strategy closed-trade statistics (PortClosedTrade)
    ' ================================================================
    If hasClosed Then
        Application.StatusBar = "Trade Analytics: Reading PortClosedTrade..."
        Dim lastRowCT As Long, lastColCT As Long
        lastRowCT = wsClosedTrade.Cells(wsClosedTrade.Rows.Count, 1).End(xlUp).Row
        lastColCT = wsClosedTrade.Cells(1, wsClosedTrade.Columns.Count).End(xlToLeft).Column

        If lastRowCT > 1 And lastColCT > 1 Then
            Dim ctData As Variant
            ctData = wsClosedTrade.Range(wsClosedTrade.Cells(1, 1), wsClosedTrade.Cells(lastRowCT, lastColCT)).Value
            Dim nStrats As Long: nStrats = lastColCT - 1

            Dim stratNames() As String: ReDim stratNames(1 To nStrats)
            Dim winD() As Long, lossD() As Long, flatD() As Long
            Dim sumW() As Double, sumL() As Double
            Dim bestD() As Double, worstD() As Double
            ReDim winD(1 To nStrats): ReDim lossD(1 To nStrats): ReDim flatD(1 To nStrats)
            ReDim sumW(1 To nStrats): ReDim sumL(1 To nStrats)
            ReDim bestD(1 To nStrats): ReDim worstD(1 To nStrats)

            Dim s As Long
            For s = 1 To nStrats
                stratNames(s) = CStr(ctData(1, s + 1))
                bestD(s) = -1E+20: worstD(s) = 1E+20
            Next s

            Dim r As Long, pnl As Double
            For r = 2 To lastRowCT
                For s = 1 To nStrats
                    If IsNumeric(ctData(r, s + 1)) Then
                        pnl = CDbl(ctData(r, s + 1))
                        If pnl > 0 Then
                            winD(s) = winD(s) + 1: sumW(s) = sumW(s) + pnl
                        ElseIf pnl < 0 Then
                            lossD(s) = lossD(s) + 1: sumL(s) = sumL(s) + pnl
                        Else
                            flatD(s) = flatD(s) + 1
                        End If
                        If pnl > bestD(s) Then bestD(s) = pnl
                        If pnl < worstD(s) Then worstD(s) = pnl
                    End If
                Next s
            Next r

            With wsTA
                .Cells(nextRow, 1).Value = "STRATEGY CLOSED-TRADE STATISTICS  (source: PortClosedTrade)"
                .Cells(nextRow, 1).Font.Bold = True: .Cells(nextRow, 1).Font.Size = 12

                Dim ctHdrs As Variant
                ctHdrs = Array("Strategy", "Win Days", "Loss Days", "Flat Days", "Win Rate %", _
                               "Avg Win/Day", "Avg Loss/Day", "Profit Factor", "Best Day", "Worst Day", "Net P&L")
                For c = 0 To UBound(ctHdrs)
                    With .Cells(nextRow + 1, c + 1)
                        .Value = ctHdrs(c): .Font.Bold = True
                        .Interior.Color = RGB(0, 70, 127): .Font.Color = RGB(255, 255, 255): .WrapText = True
                    End With
                Next c

                dataRow = nextRow + 2
                Dim tot As Long, wr As Double, avgW As Double, avgL As Double, pf As Double
                For s = 1 To nStrats
                    tot = winD(s) + lossD(s)
                    If tot = 0 Then GoTo SkipStrat
                    If s Mod 2 = 0 Then .Range(.Cells(dataRow, 1), .Cells(dataRow, 11)).Interior.Color = RGB(242, 242, 242)
                    wr = winD(s) / tot * 100
                    avgW = IIf(winD(s) > 0, sumW(s) / winD(s), 0)
                    avgL = IIf(lossD(s) > 0, sumL(s) / lossD(s), 0)
                    pf = IIf(sumL(s) <> 0, -sumW(s) / sumL(s), 0)
                    .Cells(dataRow, 1).Value = stratNames(s)
                    .Cells(dataRow, 2).Value = winD(s)
                    .Cells(dataRow, 3).Value = lossD(s)
                    .Cells(dataRow, 4).Value = flatD(s)
                    .Cells(dataRow, 5).Value = wr
                    .Cells(dataRow, 6).Value = avgW
                    .Cells(dataRow, 7).Value = avgL
                    .Cells(dataRow, 8).Value = pf
                    .Cells(dataRow, 9).Value = IIf(bestD(s) > -1E+19, bestD(s), 0)
                    .Cells(dataRow, 10).Value = IIf(worstD(s) < 1E+19, worstD(s), 0)
                    .Cells(dataRow, 11).Value = sumW(s) + sumL(s)
                    ' Colour win rate
                    Select Case True
                        Case wr >= 60: .Cells(dataRow, 5).Interior.Color = RGB(150, 230, 150)
                        Case wr >= 45: .Cells(dataRow, 5).Interior.Color = RGB(255, 255, 150)
                        Case Else:     .Cells(dataRow, 5).Interior.Color = RGB(255, 150, 150)
                    End Select
                    ' Colour profit factor
                    Select Case True
                        Case pf >= 2:  .Cells(dataRow, 8).Interior.Color = RGB(150, 230, 150)
                        Case pf >= 1:  .Cells(dataRow, 8).Interior.Color = RGB(255, 255, 150)
                        Case Else:     .Cells(dataRow, 8).Interior.Color = RGB(255, 150, 150)
                    End Select
                    dataRow = dataRow + 1
SkipStrat:
                Next s
                .Range(.Cells(nextRow + 2, 5), .Cells(dataRow - 1, 5)).NumberFormat = "0.0"
                .Range(.Cells(nextRow + 2, 6), .Cells(dataRow - 1, 11)).NumberFormat = "$#,##0"
                .Range(.Cells(nextRow + 2, 8), .Cells(dataRow - 1, 8)).NumberFormat = "0.00"
                nextRow = dataRow + 2
            End With
        End If
    End If

    ' ================================================================
    ' Section 2 — ATR Efficiency (TrueRanges vs TradePNL)
    ' ================================================================
    If hasTR And hasTP Then
        Application.StatusBar = "Trade Analytics: ATR efficiency analysis..."
        Dim lastRowTR As Long, lastColTR As Long
        lastRowTR = wsTR.Cells(wsTR.Rows.Count, 1).End(xlUp).Row
        lastColTR = wsTR.Cells(1, wsTR.Columns.Count).End(xlToLeft).Column
        Dim lastRowTP As Long, lastColTP As Long
        lastRowTP = wsTP.Cells(wsTP.Rows.Count, 1).End(xlUp).Row
        lastColTP = wsTP.Cells(1, wsTP.Columns.Count).End(xlToLeft).Column

        If lastRowTR > 1 And lastColTR > 1 And lastRowTP > 1 Then
            Dim trData As Variant, tpData As Variant
            trData = wsTR.Range(wsTR.Cells(1, 1), wsTR.Cells(lastRowTR, lastColTR)).Value
            tpData = wsTP.Range(wsTP.Cells(1, 1), wsTP.Cells(lastRowTP, lastColTP)).Value

            Dim nContracts As Long: nContracts = lastColTR - 1
            Dim contractNames() As String: ReDim contractNames(1 To nContracts)
            Dim trColM() As Long: ReDim trColM(1 To nContracts)
            Dim tpColM() As Long: ReDim tpColM(1 To nContracts)

            Dim co As Long, jj As Long
            For co = 1 To nContracts
                contractNames(co) = CStr(trData(1, co + 1))
                trColM(co) = co + 1
                For jj = 2 To lastColTP
                    If CStr(tpData(1, jj)) = contractNames(co) Then tpColM(co) = jj: Exit For
                Next jj
            Next co

            Dim sumATR() As Double, cntATR() As Long
            Dim sumPW() As Double, cntW() As Long
            Dim sumPL() As Double, cntL() As Long
            Dim maxATR() As Double
            ReDim sumATR(1 To nContracts): ReDim cntATR(1 To nContracts)
            ReDim sumPW(1 To nContracts): ReDim cntW(1 To nContracts)
            ReDim sumPL(1 To nContracts): ReDim cntL(1 To nContracts)
            ReDim maxATR(1 To nContracts)

            ' Build date→row index for TradePNL
            Dim tpDateMap As Object
            Set tpDateMap = CreateObject("Scripting.Dictionary")
            Dim rr As Long
            For rr = 2 To lastRowTP
                If IsDate(tpData(rr, 1)) Then
                    Dim tpKey As String: tpKey = CStr(CDate(tpData(rr, 1)))
                    If Not tpDateMap.Exists(tpKey) Then tpDateMap.Add tpKey, rr
                End If
            Next rr

            Dim atrV As Double, pnlV As Double, tpRow As Long, trKey As String
            For r = 2 To lastRowTR
                If Not IsDate(trData(r, 1)) Then GoTo NextTR
                trKey = CStr(CDate(trData(r, 1)))
                tpRow = 0
                If tpDateMap.Exists(trKey) Then tpRow = tpDateMap(trKey)
                For co = 1 To nContracts
                    atrV = 0: pnlV = 0
                    If IsNumeric(trData(r, trColM(co))) Then atrV = CDbl(trData(r, trColM(co)))
                    If tpRow > 0 And tpColM(co) > 0 Then
                        If IsNumeric(tpData(tpRow, tpColM(co))) Then pnlV = CDbl(tpData(tpRow, tpColM(co)))
                    End If
                    If atrV > 0 Then
                        sumATR(co) = sumATR(co) + atrV: cntATR(co) = cntATR(co) + 1
                        If atrV > maxATR(co) Then maxATR(co) = atrV
                        If pnlV > 0 Then
                            sumPW(co) = sumPW(co) + pnlV: cntW(co) = cntW(co) + 1
                        ElseIf pnlV < 0 Then
                            sumPL(co) = sumPL(co) + pnlV: cntL(co) = cntL(co) + 1
                        End If
                    End If
NextTR:
                Next co
            Next r

            With wsTA
                .Cells(nextRow, 1).Value = "ATR EFFICIENCY ANALYSIS — Contract Level  (source: TrueRanges + TradePNL)"
                .Cells(nextRow, 1).Font.Bold = True: .Cells(nextRow, 1).Font.Size = 12
                .Cells(nextRow + 1, 1).Value = "Efficiency % measures how much of the daily ATR (MFE-MAE range) is captured as profit on winning days."
                .Cells(nextRow + 1, 1).Font.Italic = True

                Dim atrHdrs As Variant
                atrHdrs = Array("Contract", "Trade Days", "Avg ATR", "Max ATR Day", "Win Days", "Avg Winner PNL", _
                                "Loss Days", "Avg Loser PNL", "Win/Loss Ratio", "Efficiency %")
                For c = 0 To UBound(atrHdrs)
                    With .Cells(nextRow + 2, c + 1)
                        .Value = atrHdrs(c): .Font.Bold = True
                        .Interior.Color = RGB(100, 60, 0): .Font.Color = RGB(255, 255, 255): .WrapText = True
                    End With
                Next c

                dataRow = nextRow + 3
                Dim avgATR As Double, avgW2 As Double, avgL2 As Double, eff As Double, wlr As Double
                For co = 1 To nContracts
                    If cntATR(co) = 0 Then GoTo SkipCo
                    If co Mod 2 = 0 Then .Range(.Cells(dataRow, 1), .Cells(dataRow, 10)).Interior.Color = RGB(242, 242, 242)
                    avgATR = sumATR(co) / cntATR(co)
                    avgW2 = IIf(cntW(co) > 0, sumPW(co) / cntW(co), 0)
                    avgL2 = IIf(cntL(co) > 0, sumPL(co) / cntL(co), 0)
                    eff = IIf(avgATR > 0 And avgW2 > 0, avgW2 / avgATR * 100, 0)
                    wlr = IIf(avgL2 <> 0, -avgW2 / avgL2, 0)
                    .Cells(dataRow, 1).Value = contractNames(co)
                    .Cells(dataRow, 2).Value = cntATR(co)
                    .Cells(dataRow, 3).Value = avgATR
                    .Cells(dataRow, 4).Value = maxATR(co)
                    .Cells(dataRow, 5).Value = cntW(co)
                    .Cells(dataRow, 6).Value = avgW2
                    .Cells(dataRow, 7).Value = cntL(co)
                    .Cells(dataRow, 8).Value = avgL2
                    .Cells(dataRow, 9).Value = wlr
                    .Cells(dataRow, 10).Value = eff
                    ' Colour efficiency
                    Select Case True
                        Case eff >= 70: .Cells(dataRow, 10).Interior.Color = RGB(150, 230, 150)
                        Case eff >= 40: .Cells(dataRow, 10).Interior.Color = RGB(255, 255, 150)
                        Case eff > 0:   .Cells(dataRow, 10).Interior.Color = RGB(255, 150, 150)
                    End Select
                    dataRow = dataRow + 1
SkipCo:
                Next co
                .Range(.Cells(nextRow + 3, 3), .Cells(dataRow - 1, 4)).NumberFormat = "$#,##0"
                .Range(.Cells(nextRow + 3, 6), .Cells(dataRow - 1, 8)).NumberFormat = "$#,##0"
                .Range(.Cells(nextRow + 3, 9), .Cells(dataRow - 1, 9)).NumberFormat = "0.00"
                .Range(.Cells(nextRow + 3, 10), .Cells(dataRow - 1, 10)).NumberFormat = "0.0"
                nextRow = dataRow + 2
            End With
        End If
    End If

    With wsTA
        .Columns("A:K").AutoFit
        .Columns(1).ColumnWidth = IIf(.Columns(1).ColumnWidth < 20, 20, .Columns(1).ColumnWidth)
    End With

    wsTA.Activate
    MsgBox "Trade Analytics complete!", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "Trade Analytics error (line " & Erl & "): " & Err.Description, vbCritical
    Resume CleanExit
End Sub


' ============================================================
' Automated Report — runs all market analyses in one pass:
'   1. Markets summary (ATR, sector, correlations, volatility,
'      seasonality, regime analysis)
'   2. Rolling correlations (strategy P&L correlations over time)
'   3. Trade analytics (MAE/MFE efficiency per strategy/contract)
' ============================================================
Sub RunAutomatedReport()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrorHandler

    Call InitializeColumnConstantsManually

    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license.", vbCritical
        GoTo CleanExit
    End If

    Dim startTime As Double: startTime = Timer

    Application.StatusBar = "Automated Report: Step 1/3 — Markets analysis..."
    Call CreateMarketsSummary

    Application.StatusBar = "Automated Report: Step 2/3 — Rolling correlations..."
    Call CreateRollingCorrelationSheet

    Application.StatusBar = "Automated Report: Step 3/3 — Trade analytics..."
    Call CreateTradeAnalytics

    Application.StatusBar = False
    Dim elapsed As Long: elapsed = CLng(Timer - startTime)
    MsgBox "Automated report complete in " & elapsed & " seconds." & vbCrLf & _
           "Sheets created: Markets, MarketCorrelations, MarketVolatility," & vbCrLf & _
           "MarketSeasonality, MarketRegimes, RollingCorrelations, TradeAnalytics.", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "Automated report error (line " & Erl & "): " & Err.Description, vbCritical
    Resume CleanExit
End Sub


' ============================================================
' Navigation helper called from buttons / I_MISC
' ============================================================
Sub GoToMarkets()
    Dim wsMarkets As Worksheet
    On Error Resume Next
    Set wsMarkets = ThisWorkbook.Sheets("Markets")
    On Error GoTo 0
    If wsMarkets Is Nothing Then
        MsgBox "Markets sheet not found. Please run 'Create Markets Summary' first.", vbExclamation
        Exit Sub
    End If
    wsMarkets.Activate
End Sub

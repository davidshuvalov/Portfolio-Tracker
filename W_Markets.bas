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

    wsMarkets.Activate
    Application.StatusBar = False

    MsgBox "Markets analysis complete." & vbCrLf & _
           "Sheets created: Markets, MarketCorrelations, MarketVolatility.", vbInformation

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

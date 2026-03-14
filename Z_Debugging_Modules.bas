Attribute VB_Name = "Z_Debugging_Modules"


Sub DebugWeeklyVsDailyPnL()
    Dim wsPort As Worksheet, wsDebug As Worksheet
    Dim yearsToConsider As Double
    Dim currentdate As Date, startdate As Date, endDate As Date
    Dim dailyPnL As Variant, weeklyPnL As Variant, originaldailyPnL As Variant
    Dim numDailyRows As Long, numWeeklyRows As Long, numOrigRows As Long
    Dim numStrategies As Long
    Dim dailySum As Double, weeklySum As Double, OrigSum As Double
    Dim i As Long, j As Long

    Call InitializeColumnConstantsManually

    '— get date range—
    Set wsPort = ThisWorkbook.Sheets("Portfolio")
    yearsToConsider = GetNamedRangeValue("PortfolioPeriod")
    currentdate = wsPort.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    startdate = DateAdd("yyyy", -Int(yearsToConsider), currentdate)
    startdate = DateAdd("m", -(yearsToConsider - Int(yearsToConsider)) * 12, startdate)
    endDate = currentdate

    '— fetch arrays—
    dailyPnL = CleanPortfolioDailyPnL(startdate, endDate)
    weeklyPnL = ConvertDailyToWeeklyPnL(startdate, endDate)
    originaldailyPnL = GetNonZeroDailyPortfolioPNLDays(startdate, endDate)

    '— create/clear debug sheet—
    On Error Resume Next
    Set wsDebug = ThisWorkbook.Sheets("PnLDebug")
    If Not wsDebug Is Nothing Then
        wsDebug.Cells.Clear
    Else
        Set wsDebug = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsDebug.name = "PnLDebug"
    End If
    On Error GoTo 0

    '— if no data, note and exit —
    If IsEmpty(dailyPnL) Or IsEmpty(weeklyPnL) Then
        wsDebug.Range("A1").value = "No PnL data found between " & Format(startdate, "yyyy-mm-dd") & " and " & Format(endDate, "yyyy-mm-dd")
        wsDebug.Activate
        Exit Sub
    End If

    '— dimensions & init sums—
    numDailyRows = UBound(dailyPnL, 1)
    numStrategies = UBound(dailyPnL, 2)
    numWeeklyRows = UBound(weeklyPnL, 1)

    dailySum = 0: weeklySum = 0: OrigSum = 0

    '— compute totals—
    For i = 1 To numDailyRows: For j = 1 To numStrategies: dailySum = dailySum + dailyPnL(i, j): Next j: Next i
    For i = 1 To numWeeklyRows: For j = 1 To numStrategies: weeklySum = weeklySum + weeklyPnL(i, j): Next j: Next i


    '— print summary—
    With wsDebug
        .Range("A1").value = "PnL Comparison (Daily vs Weekly)"
        .Range("A3").value = "Date Range:": .Range("B3").value = Format(startdate, "yyyy-mm-dd") & " to " & Format(endDate, "yyyy-mm-dd")
        .Range("A4").value = "Daily Periods:": .Range("B4").value = numDailyRows
        .Range("A5").value = "Weekly Periods:": .Range("B5").value = numWeeklyRows
        .Range("A6").value = "Strategies:": .Range("B6").value = numStrategies

        .Range("A8").value = "Total Daily PnL:": .Range("B8").value = dailySum
        .Range("A9").value = "Total Weekly PnL:": .Range("B9").value = weeklySum
        .Range("A10").value = "Difference (W–D):": .Range("B10").value = weeklySum - dailySum
        .Range("A11").value = "Pct Difference:": .Range("B11").value = IIf(dailySum <> 0, (weeklySum - dailySum) / Abs(dailySum), 0)


        .Range("B8:B10").NumberFormat = "$#,##0.00"
        .Range("B11").NumberFormat = "0.00%"

        .Range("A13:E13").value = Array("Strat", "Daily Sum", "Weekly Sum", "Diff", "% Diff")
        For j = 1 To numStrategies
            Dim sd As Double, sw As Double
            sd = 0: sw = 0
            For i = 1 To numDailyRows: sd = sd + dailyPnL(i, j): Next i
            For i = 1 To numWeeklyRows: sw = sw + weeklyPnL(i, j): Next i
            .Cells(13 + j, 1).value = j
            .Cells(13 + j, 2).value = sd
            .Cells(13 + j, 3).value = sw
            .Cells(13 + j, 4).value = sw - sd
            .Cells(13 + j, 5).value = IIf(sd <> 0, (sw - sd) / Abs(sd), 0)
            .Range(.Cells(13 + j, 2), .Cells(13 + j, 4)).NumberFormat = "$#,##0.00"
            .Cells(13 + j, 5).NumberFormat = "0.00%"
        Next j
        .Columns.AutoFit
    End With

Call DumpDailyPnLByDate


    wsDebug.Activate
    MsgBox "PnLDebug sheet created.", vbInformation
End Sub


Sub DumpDailyPnLByDate()
    Dim wsSrc      As Worksheet
    Dim wsDump     As Worksheet
    Dim dailyArr   As Variant
    Dim lastCol    As Long
    Dim i As Long, j As Long

    ' 1) define your window here
    Dim startdate  As Date
    Dim endDate    As Date
    startdate = DateSerial(2006, 12, 29)    ' ? adjust as needed
    endDate = DateSerial(2025, 4, 25)     ' ? adjust as needed

    ' 2) call your cleaning function
    dailyArr = CleanPortfolioDailyPnL(startdate, endDate)
    If IsEmpty(dailyArr) Then
        MsgBox "No data returned by CleanPortfolioDailyPnL.", vbExclamation
        Exit Sub
    End If

    ' 3) get a reference to the source sheet & header count
    Set wsSrc = ThisWorkbook.Sheets("PortfolioDailyM2M")
    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).column

    ' 4) create or clear the dump sheet
    On Error Resume Next
      Set wsDump = ThisWorkbook.Sheets("DailyPnLDump")
    On Error GoTo 0
    If wsDump Is Nothing Then
        Set wsDump = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsDump.name = "DailyPnLDump"
    Else
        wsDump.Cells.Clear
    End If

    ' 5) write headers: Date + strategy names from row 1, cols 2…lastCol
    wsDump.Cells(1, 1).value = "Date"
    For j = 2 To lastCol
        wsDump.Cells(1, j).value = wsSrc.Cells(1, j).value
    Next j

    ' 6) dump the array (dailyArr is 1…N rows × 0…(stratCount) cols)
    Dim numRows As Long, stratCount As Long
    numRows = UBound(dailyArr, 1)
    stratCount = UBound(dailyArr, 2)
    For i = 1 To numRows
        ' column A = date
        wsDump.Cells(i + 1, 1).value = dailyArr(i, 0)
        ' columns B… = PnL values
        For j = 1 To stratCount
            wsDump.Cells(i + 1, j + 1).value = dailyArr(i, j)
        Next j
    Next i

    wsDump.Columns.AutoFit
    MsgBox "DailyPnLDump sheet populated.", vbInformation
End Sub


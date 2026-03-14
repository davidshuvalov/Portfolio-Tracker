Attribute VB_Name = "U_BackTest_Eligibility"
Option Explicit

' ===== Sheets & headers (adjust if yours differ) =====
Private Const SHEET_M2M As String = "DailyM2MEquity"   ' Date in col 1; strategies in col 2..N (daily PnL)
Private Const SHEET_SUM As String = "Summary"
Private Const HDR_STRAT As String = "Strategy Name"
Private Const HDR_OOSBEG As String = "OOS Begin Date"
Private Const HDR_INCUBATION_PASSED As String = "Incubation Passed Date"
Private Const HDR_ISDD_CANDIDATES As String = "IS Max Drawdown|In Sample Max Drawdown|IS Max DD"
Private Const HDR_STATUS As String = "Status (Input Column)"
Private Const HDR_SECTOR As String = "Sector"
Private Const HDR_SYMBOL As String = "Symbol"
Private Const COL_EXPECTED_ANNUAL_PROFIT As String = "Expected Annual Return"
' ===== 1. UPDATED CONSTANTS - Add these at the top =====
Private Const NM_ENABLE_SECTOR As String = "EnableSectorAnalysis"    ' TRUE/FALSE
Private Const NM_ENABLE_SYMBOL As String = "EnableSymbolAnalysis"    ' TRUE/FALSE


' ===== Named ranges (optional; defaults used if missing) =====
Private Const NM_INCUBATE As String = "EligibilityDaysThresholdOOS"  ' >=0 (default 0)
Private Const NM_DDCAP    As String = "OOSDDvsISCap"              ' 0 disables (default 1.5)
Private Const NM_STATUS_INCLUDE As String = "EligibilityStatusInclude" ' list/range of statuses to include
Private Const NM_EFFICIENCY_RATIO As String = "EfficiencyRatio"
Private Const NM_ELIGIBILITY_DATE_TYPE As String = "EligibilityDateType"  ' "OOS Start Date" or "Incubation Pass Date"

' Default if EligibilityStatusInclude not provided
Private Const DEFAULT_STATUS_LIST As String = "Live"

' ===== Horizons config =====
Public Const MAX_HORIZON As Long = 12   ' evaluate next 1..12 months

' Module-level variable to store current rule count
Private m_CurrentRuleCount As Long

' Module-level storage for base case averages
Private baseCaseAvgMonthly() As Double


' Rule definition structure
Private Type EligibilityRule
    ID As Long
    label As String
    RuleType As String  ' "CONSECUTIVE", "COUNT", "THRESHOLD", "MOMENTUM", etc.
    Param1 As Double    ' First parameter (e.g., number of months)
    Param2 As Double    ' Second parameter (e.g., threshold multiplier)
    Param3 As Double    ' Third parameter if needed
    IsActive As Boolean
End Type

' Per-condition tallies across horizons (dynamic arrays)
' ===== 2. ENHANCED CondStat TYPE - Replace existing type =====
Private Type CondStat
    label As String
    n() As Long
    W() As Long
    SumPnL() As Double
    SumPct() As Double     ' sum of percentage change vs BASE CASE
    CntPct() As Long       ' count of valid ratio terms
    IsBaseCase As Boolean  ' Flag to identify the baseline rule
End Type

' ===== ENTRY POINT =====
Public Sub Study_Eligibility_Rules_Horizons_1to12_Combos()
    Dim wsM As Worksheet, wsS As Worksheet
        Dim strategyMonthlyData As Object
    Set strategyMonthlyData = CreateObject("Scripting.Dictionary")
    Dim currentIteration As Long, totalIterations As Long
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets(SHEET_M2M)
    Set wsS = ThisWorkbook.Worksheets(SHEET_SUM)
    On Error GoTo 0
    If wsM Is Nothing Or wsS Is Nothing Then
        MsgBox "Can't find '" & SHEET_M2M & "' or '" & SHEET_SUM & "' sheet.", vbCritical
        Exit Sub
    End If

       ' ===== DELETE ALL EXISTING ANALYSIS SHEETS =====
    DeleteAnalysisSheets
    

    ' Read ALL of DailyM2MEquity
    Dim lastCol As Long, lastRow As Long
    lastCol = wsM.Cells(1, wsM.Columns.count).End(xlToLeft).column
    lastRow = wsM.Cells(wsM.rows.count, 1).End(xlUp).row
    If lastCol < 2 Or lastRow < 2 Then
        MsgBox "DailyM2MEquity has no strategy columns or no data.", vbExclamation
        Exit Sub
    End If
    Dim data As Variant
    data = wsM.Range(wsM.Cells(1, 1), wsM.Cells(lastRow, lastCol)).value2  ' 1-based 2D array

    ' Map strategies -> column
    Dim colMap As Object: Set colMap = CreateObject("Scripting.Dictionary")
    Dim c As Long, sName As String
    For c = 2 To lastCol
        sName = Trim$(CStr(data(1, c)))
        If Len(sName) > 0 Then colMap(sName) = c
    Next

    ' Read Summary: columns & maps
    Dim cStrat As Long, cOOS As Long, cIncubation As Long, cIS As Long, cStatus As Long, cSector As Long, cSymbol As Long, cExpected As Long
    cStrat = FindHeaderCol(wsS, HDR_STRAT)
    cOOS = FindHeaderCol(wsS, HDR_OOSBEG)
    cIncubation = FindHeaderCol(wsS, HDR_INCUBATION_PASSED)
    cIS = FindHeaderColAlt(wsS, Split(HDR_ISDD_CANDIDATES, "|"))
    cStatus = FindHeaderCol(wsS, HDR_STATUS)
    cSector = FindHeaderCol(wsS, HDR_SECTOR)
    cSymbol = FindHeaderCol(wsS, HDR_SYMBOL)
    cExpected = FindHeaderCol(wsS, COL_EXPECTED_ANNUAL_PROFIT)
    
    ' Determine which date type to use
    Dim dateTypeChoice As String, useDateCol As Long, dateColName As String
    dateTypeChoice = Trim$(CStr(NzNamed(NM_ELIGIBILITY_DATE_TYPE, "OOS Start Date")))
    
    If InStr(1, dateTypeChoice, "Incubation", vbTextCompare) > 0 Then
        useDateCol = cIncubation
        dateColName = HDR_INCUBATION_PASSED
    Else
        useDateCol = cOOS
        dateColName = HDR_OOSBEG
    End If
    
    If cStrat = 0 Or useDateCol = 0 Or cStatus = 0 Then
        MsgBox "Missing one or more required Summary headers: '" & HDR_STRAT & "', '" & dateColName & "', '" & HDR_STATUS & "'.", vbCritical
        Exit Sub
    End If

    ' Status include list (case-insensitive)
    Dim includeStatuses As Collection
    Set includeStatuses = ReadIncludeList(NM_STATUS_INCLUDE, DEFAULT_STATUS_LIST)

    Dim eligibilityBeg As Object: Set eligibilityBeg = CreateObject("Scripting.Dictionary")
    Dim isDD As Object:   Set isDD = CreateObject("Scripting.Dictionary")
    Dim secMap As Object: Set secMap = CreateObject("Scripting.Dictionary")
    Dim symMap As Object: Set symMap = CreateObject("Scripting.Dictionary")
    Dim expectedReturns As Object: Set expectedReturns = CreateObject("Scripting.Dictionary")

    Dim r As Long, lastS As Long
    lastS = wsS.Cells(wsS.rows.count, cStrat).End(xlUp).row
    For r = 2 To lastS
        sName = Trim$(CStr(wsS.Cells(r, cStrat).value))
        If Len(sName) > 0 Then
            Dim statusVal As String
            statusVal = Trim$(CStr(wsS.Cells(r, cStatus).value))
            If StatusIsIncluded(statusVal, includeStatuses) Then
                ' Use the selected date column and skip if blank
                Dim b As Variant: b = wsS.Cells(r, useDateCol).value
                If IsDate(b) And Not IsEmpty(b) Then
                    eligibilityBeg(sName) = CDate(b)
                    
                    If cIS > 0 Then
                        Dim v As Variant: v = wsS.Cells(r, cIS).value
                        If IsNumeric(v) Then isDD(sName) = Abs(CDbl(v))
                    End If
                    If cSector > 0 Then secMap(sName) = Trim$(CStr(wsS.Cells(r, cSector).value))
                    If cSymbol > 0 Then symMap(sName) = Trim$(CStr(wsS.Cells(r, cSymbol).value))
                    If cExpected > 0 Then
                        Dim expVal As Variant: expVal = wsS.Cells(r, cExpected).value
                        If IsNumeric(expVal) Then expectedReturns(sName) = CDbl(expVal)
                    End If
                End If
            End If
        End If
    Next

    If eligibilityBeg.count = 0 Then
        MsgBox "No strategies matched the included statuses with valid " & dateColName & " dates.", vbExclamation
        Exit Sub
    End If

    ' Build month windows from Date column
    Dim mStart() As Long, mEnd() As Long, mCount As Long
    BuildMonthWindows data, mStart, mEnd, mCount
    If mCount < 2 Then
        MsgBox "Not enough full months in DailyM2MEquity.", vbExclamation
        Exit Sub
    End If

    ' Parameters
    Dim incubDays As Long, ddCapX As Double
    incubDays = CLng(NzNamed(NM_INCUBATE, 0))
    ddCapX = CDbl(NzNamed(NM_DDCAP, 1.5))   ' 0 disables the DD cap

    ' ===== Initialize flexible rules system =====
    Dim rules() As EligibilityRule
    InitializeEligibilityRules rules
    Dim totalRules As Long: totalRules = UBound(rules)
    m_CurrentRuleCount = totalRules  ' Store for global access
    
    ' Initialize condition arrays with dynamic size
    Dim conds() As CondStat
    ' ===== Initialize segment analysis dictionaries =====
        Dim enableSectorAnalysis As Boolean, enableSymbolAnalysis As Boolean
        enableSectorAnalysis = CBool(NzNamed(NM_ENABLE_SECTOR, "Yes") = "Yes")  ' Default True
        enableSymbolAnalysis = CBool(NzNamed(NM_ENABLE_SYMBOL, "Yes") = "Yes")  ' Default True
        
        Dim secN As Object, secW As Object, secSum As Object, secSumPct As Object, secCntPct As Object
        Dim symN As Object, symW As Object, symSum As Object, symSumPct As Object, symCntPct As Object
        
        If cSector > 0 And enableSectorAnalysis Then
            Set secN = CreateObject("Scripting.Dictionary")
            Set secW = CreateObject("Scripting.Dictionary")
            Set secSum = CreateObject("Scripting.Dictionary")
            Set secSumPct = CreateObject("Scripting.Dictionary")
            Set secCntPct = CreateObject("Scripting.Dictionary")
        End If
        
        If cSymbol > 0 And enableSymbolAnalysis Then
            Set symN = CreateObject("Scripting.Dictionary")
            Set symW = CreateObject("Scripting.Dictionary")
            Set symSum = CreateObject("Scripting.Dictionary")
            Set symSumPct = CreateObject("Scripting.Dictionary")
            Set symCntPct = CreateObject("Scripting.Dictionary")
        End If
    
    ReDim conds(1 To totalRules)
    
    ' Set labels from rules and size dynamic arrays
    Dim ci As Long
    For ci = 1 To totalRules
        conds(ci).label = rules(ci).label
        conds(ci).IsBaseCase = (rules(ci).RuleType = "BASELINE")  ' Mark base case
        ReDim conds(ci).n(1 To MAX_HORIZON)
        ReDim conds(ci).W(1 To MAX_HORIZON)
        ReDim conds(ci).SumPnL(1 To MAX_HORIZON)
        ReDim conds(ci).SumPct(1 To MAX_HORIZON)
        ReDim conds(ci).CntPct(1 To MAX_HORIZON)
    Next ci
    
    ' Initialize base case tracking
    ' Insert this section right AFTER your data collection loops end but BEFORE output generation:

    Application.StatusBar = "Calculating base case comparisons..."
    
    ' Initialize base case tracking
    ReDim baseCaseAvgMonthly(1 To MAX_HORIZON)
    
   

    ' Segment processing disabled for performance

    ' ===== Main evaluation with pre-calculation =====
    Application.StatusBar = "Pre-calculating monthly PnL data..."
    
    ' Pre-calculate all monthly data (THIS IS THE KEY PERFORMANCE IMPROVEMENT)
    BuildStrategyMonthlyData data, colMap, eligibilityBeg, mStart, mEnd, mCount, strategyMonthlyData
    
    ' Initialize progress tracking - FIXED
    totalIterations = (mCount - 1) * eligibilityBeg.count
    currentIteration = 0
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Starting eligibility analysis..."
    
    Dim j As Long
    For j = 1 To mCount - 1
        Dim rEvalS As Long, rEvalE As Long
        rEvalS = mStart(j): rEvalE = mEnd(j)
        Dim dEval As Date: dEval = CDate(data(rEvalE, 1))
    
        ' Month-aligned trailing window starts (same as before)
        Dim d1 As Date, d3 As Date, d6 As Date, d9 As Date, d12 As Date
        d1 = DateSerial(Year(dEval), Month(dEval), 1)
        d3 = DateSerial(Year(dEval), Month(dEval) - 2, 1)
        d6 = DateSerial(Year(dEval), Month(dEval) - 5, 1)
        d9 = DateSerial(Year(dEval), Month(dEval) - 8, 1)
        d12 = DateSerial(Year(dEval), Month(dEval) - 11, 1)
    
        Dim key As Variant
        For Each key In eligibilityBeg.keys
            ' INCREMENT COUNTER FIRST - FIXED
            currentIteration = currentIteration + 1
            
            ' UPDATE PROGRESS - MOVED TO CORRECT POSITION
            If currentIteration Mod 50 = 0 Or currentIteration = 1 Then  ' Show progress every 50 iterations and at start
                Dim progressPct As Double: progressPct = currentIteration / totalIterations
                Application.StatusBar = "Processing " & currentIteration & " of " & totalIterations & " (" & Format(progressPct, "0%") & ") - Month " & j & " of " & (mCount - 1)
                DoEvents
            End If
            
            sName = CStr(key)
            If Not colMap.Exists(sName) Then GoTo NextStrat
            Dim col As Long: col = colMap(sName)
            Dim beg As Date: beg = eligibilityBeg(sName)
            If dEval < beg Then GoTo NextStrat

        ' Use pre-calculated data (much faster!)
        If Not strategyMonthlyData.Exists(sName) Then GoTo NextStrat
        Dim monthlyPnL As Variant: monthlyPnL = strategyMonthlyData(sName)

        ' Incubation check (same as before)
        If incubDays > 0 Then
            Dim firstRowThisMo As Long
            firstRowThisMo = FirstRowOnOrAfterDate(data, rEvalS, rEvalE, beg)
            If firstRowThisMo = 0 Or (rEvalE - firstRowThisMo + 1) < incubDays Then GoTo NextStrat
        End If

        ' OOS Max DD cap check (same as before)
        If ddCapX > 0 And isDD.Exists(sName) And InStr(1, dateTypeChoice, "OOS", vbTextCompare) > 0 Then
            Dim isAbsDD As Double: isAbsDD = Abs(CDbl(isDD(sName)))
            If isAbsDD > 0 Then
                Dim oosAbsDD As Double
                oosAbsDD = MaxDrawdownUpToRow_Eligible(data, col, beg, rEvalE)
                If oosAbsDD > ddCapX * isAbsDD Then GoTo NextStrat
            End If
        End If

        ' ULTRA-FAST trailing PnL calculation using pre-calculated monthly data
        Dim p1 As Double, p3 As Double, p6 As Double, p9 As Double, p12v As Double
        p1 = GetTrailingSumFromArray(monthlyPnL, j, 1)
        p3 = GetTrailingSumFromArray(monthlyPnL, j, 3)
        p6 = GetTrailingSumFromArray(monthlyPnL, j, 6)
        p9 = GetTrailingSumFromArray(monthlyPnL, j, 9)
        p12v = GetTrailingSumFromArray(monthlyPnL, j, 12)

        ' Continue with horizon evaluation (same as before)
        Dim maxH As Long, hIdx As Long
        maxH = MAX_HORIZON
        If j + maxH > mCount Then maxH = mCount - j
        If maxH < 1 Then GoTo NextStrat

        ' OPTIMIZATION: Evaluate rules ONCE per strategy-month using pre-calculated data
        Dim ruleResults() As Boolean
        ReDim ruleResults(1 To totalRules)
        
        ' *** MISSING SECTION: Actually evaluate the rules and populate ruleResults ***
        Dim ruleIdx As Long  ' <-- ADD THIS DECLARATION
        For ruleIdx = 1 To totalRules
            If rules(ruleIdx).IsActive Then
                ' Use the fast evaluation with pre-calculated data
                ruleResults(ruleIdx) = EvaluateRuleFast(rules(ruleIdx), p1, p3, p6, p9, p12v, _
                                                       sName, monthlyPnL, j, expectedReturns)
            Else
                ruleResults(ruleIdx) = False
            End If
        Next ruleIdx

        ' Apply results to all horizons (same as before) - KEEP ONLY ONE OF THESE LOOPS
        For hIdx = 1 To maxH
            Dim rFwdS As Long, rFwdE As Long
            rFwdS = mStart(j + 1)
            rFwdE = mEnd(j + hIdx)

            Dim pfwd As Double
            pfwd = GetForwardSumFromArray(monthlyPnL, j + 1, hIdx)
            Dim win As Boolean: win = (pfwd > 0#)

             ' Update statistics for all rules that passed
            For ruleIdx = 1 To totalRules
                If ruleResults(ruleIdx) Then
                    ' Overall statistics
                    BumpOverall conds(ruleIdx), hIdx, win, pfwd, p12v
                    
                    ' Sector statistics (only if enabled)
                    If enableSectorAnalysis And cSector > 0 And secMap.Exists(sName) Then
                        Dim sectorName As String: sectorName = secMap(sName)
                        BumpSegment secN, secW, secSum, secSumPct, secCntPct, sectorName, ruleIdx, totalRules, hIdx, win, pfwd, p12v
                    End If
                    
                    ' Symbol statistics (only if enabled)
                    If enableSymbolAnalysis And cSymbol > 0 And symMap.Exists(sName) Then
                        Dim symbolName As String: symbolName = symMap(sName)
                        BumpSegment symN, symW, symSum, symSumPct, symCntPct, symbolName, ruleIdx, totalRules, hIdx, win, pfwd, p12v
                    End If
                End If
            Next ruleIdx
        Next hIdx

NextStrat:
    Next key
Next j

' Calculate percentage comparisons against final base case
CalculateBaseComparisons conds
    

    ' ===== Output: Overall with enhanced formatting =====
Application.StatusBar = "Generating formatted output..."

Dim outOverall As String: outOverall = "Rules_Analysis"  ' Simplified name
Dim wsOut As Worksheet
Set wsOut = ThisWorkbook.Worksheets.Add(After:=wsM)
wsOut.name = outOverall
    
    ' Color the tab pink
    wsOut.Tab.Color = RGB(255, 182, 193)  ' Light pink

    ' Create formatted header and summary info
    CreateFormattedHeader wsOut, totalRules, eligibilityBeg.count, dateTypeChoice, mCount - 1
    
    ' Write data starting from row 8
    Dim rowOut As Long, condIndex As Long
    rowOut = 8
    For condIndex = LBound(conds) To UBound(conds)
        WriteConditionRow wsOut, rowOut, conds(condIndex)
        rowOut = rowOut + 1
    Next condIndex
    
    ' Apply formatting
    FormatResultsTable wsOut, 8, rowOut - 1
    
    ' Add delete button
    AddDeleteButton wsOut
    
    wsOut.Columns.AutoFit
    wsOut.Activate
    wsOut.Range("A1").Select

        ' ===== Segment Output Generation =====
    If enableSectorAnalysis And cSector > 0 And Not secN Is Nothing Then
        If secN.count > 0 Then
            OutputSegmentResults "Sector_Analysis", "Sector", secN, secW, secSum, secSumPct, secCntPct, wsOut, totalRules, rules
        End If
    End If
    
    If enableSymbolAnalysis And cSymbol > 0 And Not symN Is Nothing Then
        If symN.count > 0 Then
            OutputSegmentResults "Symbol_Analysis", "Symbol", symN, symW, symSum, symSumPct, symCntPct, wsOut, totalRules, rules
        End If
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Eligibility rules analysis complete!" & vbCrLf & _
           "Rules evaluated: " & totalRules & vbCrLf & _
           "Strategies: " & eligibilityBeg.count & vbCrLf & _
           "Months processed: " & (mCount - 1) & vbCrLf & _
           "Worksheets created: " & GetWorksheetCount() & vbCrLf & _
           "Segment analysis enabled", vbInformation
End Sub

' ===== 3. UPDATED RULE INITIALIZATION - Replace entire InitializeEligibilityRules =====
Private Sub InitializeEligibilityRules(ByRef rules() As EligibilityRule)
    Dim ruleCount As Long: ruleCount = 160  ' Doubled for OOS variants
    ReDim rules(1 To ruleCount)
    Dim i As Long: i = 1
    
    ' =====================================
    ' SECTION A: BASELINE RULES
    ' =====================================
    rules(i).ID = i: rules(i).label = "Baseline (All Eligible)": rules(i).RuleType = "BASELINE": rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Out-of-Sample Profitable": rules(i).RuleType = "OOS_PROFITABLE": rules(i).IsActive = True: i = i + 1
    
    ' =====================================
    ' SECTION B: SIMPLE PERIOD RULES
    ' =====================================
    ' Original rules
    rules(i).ID = i: rules(i).label = "Last 1M > 0": rules(i).RuleType = "SIMPLE_POSITIVE": rules(i).Param1 = 1: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 3M > 0": rules(i).RuleType = "SIMPLE_POSITIVE": rules(i).Param1 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M > 0": rules(i).RuleType = "SIMPLE_POSITIVE": rules(i).Param1 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 9M > 0": rules(i).RuleType = "SIMPLE_POSITIVE": rules(i).Param1 = 9: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 12M > 0": rules(i).RuleType = "SIMPLE_POSITIVE": rules(i).Param1 = 12: rules(i).IsActive = True: i = i + 1
    
    ' OOS variants
    rules(i).ID = i: rules(i).label = "Last 1M > 0 AND OOS > 0": rules(i).RuleType = "SIMPLE_POSITIVE_OOS": rules(i).Param1 = 1: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 3M > 0 AND OOS > 0": rules(i).RuleType = "SIMPLE_POSITIVE_OOS": rules(i).Param1 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M > 0 AND OOS > 0": rules(i).RuleType = "SIMPLE_POSITIVE_OOS": rules(i).Param1 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 9M > 0 AND OOS > 0": rules(i).RuleType = "SIMPLE_POSITIVE_OOS": rules(i).Param1 = 9: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 12M > 0 AND OOS > 0": rules(i).RuleType = "SIMPLE_POSITIVE_OOS": rules(i).Param1 = 12: rules(i).IsActive = True: i = i + 1
    
    ' Original rules
    rules(i).ID = i: rules(i).label = "Last 1M < 0": rules(i).RuleType = "SIMPLE_NEGATIVE": rules(i).Param1 = 1: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 3M < 0": rules(i).RuleType = "SIMPLE_NEGATIVE": rules(i).Param1 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M < 0": rules(i).RuleType = "SIMPLE_NEGATIVE": rules(i).Param1 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 9M < 0": rules(i).RuleType = "SIMPLE_NEGATIVE": rules(i).Param1 = 9: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 12M < 0": rules(i).RuleType = "SIMPLE_NEGATIVE": rules(i).Param1 = 12: rules(i).IsActive = True: i = i + 1
    
    ' OOS variants
    rules(i).ID = i: rules(i).label = "Last 1M < 0 AND OOS > 0": rules(i).RuleType = "SIMPLE_NEGATIVE_OOS": rules(i).Param1 = 1: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 3M < 0 AND OOS > 0": rules(i).RuleType = "SIMPLE_NEGATIVE_OOS": rules(i).Param1 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M < 0 AND OOS > 0": rules(i).RuleType = "SIMPLE_NEGATIVE_OOS": rules(i).Param1 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 9M < 0 AND OOS > 0": rules(i).RuleType = "SIMPLE_NEGATIVE_OOS": rules(i).Param1 = 9: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 12M < 0 AND OOS > 0": rules(i).RuleType = "SIMPLE_NEGATIVE_OOS": rules(i).Param1 = 12: rules(i).IsActive = True: i = i + 1
    
    
    
    ' =====================================
    ' SECTION C: CONSECUTIVE MONTH RULES
    ' =====================================
    ' Original rules
    rules(i).ID = i: rules(i).label = "Last 3M All Positive": rules(i).RuleType = "CONSECUTIVE": rules(i).Param1 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 4M All Positive": rules(i).RuleType = "CONSECUTIVE": rules(i).Param1 = 4: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 5M All Positive": rules(i).RuleType = "CONSECUTIVE": rules(i).Param1 = 5: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M All Positive": rules(i).RuleType = "CONSECUTIVE": rules(i).Param1 = 6: rules(i).IsActive = True: i = i + 1
    
    ' OOS variants
    rules(i).ID = i: rules(i).label = "Last 3M All Positive AND OOS > 0": rules(i).RuleType = "CONSECUTIVE_OOS": rules(i).Param1 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 4M All Positive AND OOS > 0": rules(i).RuleType = "CONSECUTIVE_OOS": rules(i).Param1 = 4: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 5M All Positive AND OOS > 0": rules(i).RuleType = "CONSECUTIVE_OOS": rules(i).Param1 = 5: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M All Positive AND OOS > 0": rules(i).RuleType = "CONSECUTIVE_OOS": rules(i).Param1 = 6: rules(i).IsActive = True: i = i + 1
    
    ' =====================================
    ' SECTION D: COUNT-BASED RULES
    ' =====================================
    ' Original rules
    rules(i).ID = i: rules(i).label = "4+ of Last 6M Positive": rules(i).RuleType = "COUNT_POSITIVE": rules(i).Param1 = 4: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "5+ of Last 6M Positive": rules(i).RuleType = "COUNT_POSITIVE": rules(i).Param1 = 5: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "7+ of Last 9M Positive": rules(i).RuleType = "COUNT_POSITIVE": rules(i).Param1 = 7: rules(i).Param2 = 9: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "8+ of Last 12M Positive": rules(i).RuleType = "COUNT_POSITIVE": rules(i).Param1 = 8: rules(i).Param2 = 12: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "9+ of Last 12M Positive": rules(i).RuleType = "COUNT_POSITIVE": rules(i).Param1 = 9: rules(i).Param2 = 12: rules(i).IsActive = True: i = i + 1
    
    ' OOS variants
    rules(i).ID = i: rules(i).label = "4+ of Last 6M Positive AND OOS > 0": rules(i).RuleType = "COUNT_POSITIVE_OOS": rules(i).Param1 = 4: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "5+ of Last 6M Positive AND OOS > 0": rules(i).RuleType = "COUNT_POSITIVE_OOS": rules(i).Param1 = 5: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "7+ of Last 9M Positive AND OOS > 0": rules(i).RuleType = "COUNT_POSITIVE_OOS": rules(i).Param1 = 7: rules(i).Param2 = 9: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "8+ of Last 12M Positive AND OOS > 0": rules(i).RuleType = "COUNT_POSITIVE_OOS": rules(i).Param1 = 8: rules(i).Param2 = 12: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "9+ of Last 12M Positive AND OOS > 0": rules(i).RuleType = "COUNT_POSITIVE_OOS": rules(i).Param1 = 9: rules(i).Param2 = 12: rules(i).IsActive = True: i = i + 1
    
    ' =====================================
    ' SECTION E: MOMENTUM & ACCELERATION
    ' =====================================
    ' Original rules
    rules(i).ID = i: rules(i).label = "Last 1M > Prev 3M": rules(i).RuleType = "MOMENTUM": rules(i).Param1 = 1: rules(i).Param2 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 3M > Prev 3M": rules(i).RuleType = "MOMENTUM": rules(i).Param1 = 3: rules(i).Param2 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M > Prev 6M": rules(i).RuleType = "MOMENTUM": rules(i).Param1 = 6: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Accelerating (3M>6M Ann.)": rules(i).RuleType = "ACCELERATION": rules(i).Param1 = 3: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    
    ' OOS variants
    rules(i).ID = i: rules(i).label = "Last 1M > Prev 3M AND OOS > 0": rules(i).RuleType = "MOMENTUM_OOS": rules(i).Param1 = 1: rules(i).Param2 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 3M > Prev 3M AND OOS > 0": rules(i).RuleType = "MOMENTUM_OOS": rules(i).Param1 = 3: rules(i).Param2 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M > Prev 6M AND OOS > 0": rules(i).RuleType = "MOMENTUM_OOS": rules(i).Param1 = 6: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Accelerating AND OOS > 0": rules(i).RuleType = "ACCELERATION_OOS": rules(i).Param1 = 3: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    
    ' =====================================
    ' SECTION F: COMBINATION RULES
    ' =====================================
    ' Original rules
    rules(i).ID = i: rules(i).label = "Last 3M AND 6M > 0": rules(i).RuleType = "AND_COMBO": rules(i).Param1 = 3: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 3M AND 12M > 0": rules(i).RuleType = "AND_COMBO": rules(i).Param1 = 3: rules(i).Param2 = 12: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M AND 12M > 0": rules(i).RuleType = "AND_COMBO": rules(i).Param1 = 6: rules(i).Param2 = 12: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Any of {1M,3M,6M} > 0": rules(i).RuleType = "ANY_OF_3": rules(i).Param1 = 1: rules(i).Param2 = 3: rules(i).Param3 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "All of {1M,3M,6M} > 0": rules(i).RuleType = "ALL_OF_3": rules(i).Param1 = 1: rules(i).Param2 = 3: rules(i).Param3 = 6: rules(i).IsActive = True: i = i + 1
    
    ' OOS variants
    rules(i).ID = i: rules(i).label = "Last 3M AND 6M > 0 AND OOS > 0": rules(i).RuleType = "AND_COMBO_OOS": rules(i).Param1 = 3: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 3M AND 12M > 0 AND OOS > 0": rules(i).RuleType = "AND_COMBO_OOS": rules(i).Param1 = 3: rules(i).Param2 = 12: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M AND 12M > 0 AND OOS > 0": rules(i).RuleType = "AND_COMBO_OOS": rules(i).Param1 = 6: rules(i).Param2 = 12: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Any of {1M,3M,6M} > 0 AND OOS > 0": rules(i).RuleType = "ANY_OF_3_OOS": rules(i).Param1 = 1: rules(i).Param2 = 3: rules(i).Param3 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "All of {1M,3M,6M} > 0 AND OOS > 0": rules(i).RuleType = "ALL_OF_3_OOS": rules(i).Param1 = 1: rules(i).Param2 = 3: rules(i).Param3 = 6: rules(i).IsActive = True: i = i + 1
    
    ' =====================================
    ' SECTION G: THRESHOLD RULES
    ' =====================================
    ' Original rules
    rules(i).ID = i: rules(i).label = "Last 3M Ann. > Efficiency*Expected": rules(i).RuleType = "THRESHOLD_ANNUAL": rules(i).Param1 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M Ann. > Efficiency*Expected": rules(i).RuleType = "THRESHOLD_ANNUAL": rules(i).Param1 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 12M > Efficiency*Expected": rules(i).RuleType = "THRESHOLD_ANNUAL": rules(i).Param1 = 12: rules(i).IsActive = True: i = i + 1
    
    ' OOS variants
    rules(i).ID = i: rules(i).label = "Last 3M Ann. > Efficiency*Expected AND OOS > 0": rules(i).RuleType = "THRESHOLD_ANNUAL_OOS": rules(i).Param1 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 6M Ann. > Efficiency*Expected AND OOS > 0": rules(i).RuleType = "THRESHOLD_ANNUAL_OOS": rules(i).Param1 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Last 12M > Efficiency*Expected AND OOS > 0": rules(i).RuleType = "THRESHOLD_ANNUAL_OOS": rules(i).Param1 = 12: rules(i).IsActive = True: i = i + 1
    
    ' =====================================
    ' SECTION H: RECOVERY RULES
    ' =====================================
    ' Original rules
    rules(i).ID = i: rules(i).label = "Recovery: 1M>0 after 3M<0": rules(i).RuleType = "RECOVERY": rules(i).Param1 = 1: rules(i).Param2 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Recovery: 3M>0 after 6M<0": rules(i).RuleType = "RECOVERY": rules(i).Param1 = 3: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Recovery: 6M>0 after 12M<0": rules(i).RuleType = "RECOVERY": rules(i).Param1 = 6: rules(i).Param2 = 12: rules(i).IsActive = True: i = i + 1
    
    ' OOS variants
    rules(i).ID = i: rules(i).label = "Recovery: 1M>0 after 3M<0 AND OOS > 0": rules(i).RuleType = "RECOVERY_OOS": rules(i).Param1 = 1: rules(i).Param2 = 3: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Recovery: 3M>0 after 6M<0 AND OOS > 0": rules(i).RuleType = "RECOVERY_OOS": rules(i).Param1 = 3: rules(i).Param2 = 6: rules(i).IsActive = True: i = i + 1
    rules(i).ID = i: rules(i).label = "Recovery: 6M>0 after 12M<0 AND OOS > 0": rules(i).RuleType = "RECOVERY_OOS": rules(i).Param1 = 6: rules(i).Param2 = 12: rules(i).IsActive = True: i = i + 1
    
    ' Truncate array to actual number of rules used
    Dim actualCount As Long: actualCount = i - 1
    ReDim Preserve rules(1 To actualCount)
End Sub



' Helper functions for complex rule types
Private Function EvaluateConsecutivePositive(ByRef data As Variant, ByVal col As Long, _
                                             ByVal eligibleBegin As Date, ByVal evalDate As Date, _
                                             ByVal numMonths As Long) As Boolean
    ' Check if last numMonths were all positive
    Dim i As Long, monthStart As Date, monthEnd As Date, monthPnL As Double
    
    For i = 0 To numMonths - 1
        monthStart = DateSerial(Year(evalDate), Month(evalDate) - i, 1)
        monthEnd = DateSerial(Year(evalDate), Month(evalDate) - i + 1, 0)
        monthPnL = SumPnLBetweenDates_Eligible(data, col, eligibleBegin, monthStart, monthEnd)
        If monthPnL <= 0 Then
            EvaluateConsecutivePositive = False
            Exit Function
        End If
    Next i
    EvaluateConsecutivePositive = True
End Function

Private Function EvaluateCountPositive(ByRef data As Variant, ByVal col As Long, _
                                      ByVal eligibleBegin As Date, ByVal evalDate As Date, _
                                      ByVal minPositive As Long, ByVal totalMonths As Long) As Boolean
    ' Check if at least minPositive out of last totalMonths were positive
    Dim i As Long, monthStart As Date, monthEnd As Date, monthPnL As Double, positiveCount As Long
    
    positiveCount = 0
    For i = 0 To totalMonths - 1
        monthStart = DateSerial(Year(evalDate), Month(evalDate) - i, 1)
        monthEnd = DateSerial(Year(evalDate), Month(evalDate) - i + 1, 0)
        monthPnL = SumPnLBetweenDates_Eligible(data, col, eligibleBegin, monthStart, monthEnd)
        If monthPnL > 0 Then positiveCount = positiveCount + 1
    Next i
    EvaluateCountPositive = (positiveCount >= minPositive)
End Function

Private Function EvaluateMomentum(ByRef data As Variant, ByVal col As Long, _
                                  ByVal eligibleBegin As Date, ByVal evalDate As Date, _
                                  ByVal recent As Long, ByVal previous As Long) As Boolean
    ' Compare recent period vs previous period
    Dim recentStart As Date, recentEnd As Date, prevStart As Date, prevEnd As Date
    Dim recentPnL As Double, prevPnL As Double
    
    recentEnd = evalDate
    recentStart = DateSerial(Year(evalDate), Month(evalDate) - recent + 1, 1)
    prevEnd = DateSerial(Year(evalDate), Month(evalDate) - recent, 0)
    prevStart = DateSerial(Year(evalDate), Month(evalDate) - recent - previous + 1, 1)
    
    recentPnL = SumPnLBetweenDates_Eligible(data, col, eligibleBegin, recentStart, recentEnd)
    prevPnL = SumPnLBetweenDates_Eligible(data, col, eligibleBegin, prevStart, prevEnd)
    
    EvaluateMomentum = (recentPnL > prevPnL)
End Function

' ===== AGGREGATION HELPERS =====



' PASS 1: Just collect all raw statistics (no percentage calculations)
Private Sub BumpOverall(ByRef cs As CondStat, ByVal h As Long, ByVal win As Boolean, ByVal pfwd As Double, ByVal prev12 As Double)
    If h < LBound(cs.n) Or h > UBound(cs.n) Then Exit Sub
    cs.n(h) = cs.n(h) + 1
    If win Then cs.W(h) = cs.W(h) + 1
    cs.SumPnL(h) = cs.SumPnL(h) + pfwd
End Sub

' ===== 2. NEW FUNCTION: Calculate base case comparisons AFTER data collection =====
Private Sub CalculateBaseComparisons(ByRef conds() As CondStat)
    ' Find the baseline rule
    Dim baseIdx As Long: baseIdx = -1
    Dim ci As Long
    
    For ci = LBound(conds) To UBound(conds)
        If InStr(1, conds(ci).label, "Baseline (All Eligible)", vbTextCompare) > 0 Then
            baseIdx = ci
            Exit For
        End If
    Next ci
    
    If baseIdx = -1 Then
        Debug.Print "ERROR: Baseline rule not found!"
        Exit Sub
    End If
    
    ' Initialize base case array if needed
    If Not IsArrayInitialized(baseCaseAvgMonthly) Then
        ReDim baseCaseAvgMonthly(1 To MAX_HORIZON)
    End If
    
    ' Calculate base case averages
    Dim h As Long
    For h = 1 To MAX_HORIZON
        If h <= UBound(conds(baseIdx).n) And conds(baseIdx).n(h) > 0 Then
            baseCaseAvgMonthly(h) = SafeDivD(conds(baseIdx).SumPnL(h), conds(baseIdx).n(h) * h)
            Debug.Print "Base case H" & h & ": " & Format(baseCaseAvgMonthly(h), "$#,##0")
        End If
    Next h
    
    ' Calculate percentage differences for all rules
    For ci = LBound(conds) To UBound(conds)
        For h = 1 To MAX_HORIZON
            If h <= UBound(conds(ci).n) And conds(ci).n(h) > 0 Then
                If ci = baseIdx Then
                    ' Base case is always 0% different from itself
                    conds(ci).SumPct(h) = 0
                    conds(ci).CntPct(h) = 1
                Else
                    ' Calculate difference from base case
                    If baseCaseAvgMonthly(h) <> 0 Then
                        Dim ruleAvg As Double
                        ruleAvg = SafeDivD(conds(ci).SumPnL(h), conds(ci).n(h) * h)
                        
                        Dim pctDiff As Double
                        pctDiff = (ruleAvg / baseCaseAvgMonthly(h)) - 1
                        
                        ' Store the percentage (as accumulated value for compatibility)
                        conds(ci).SumPct(h) = pctDiff * conds(ci).n(h)
                        conds(ci).CntPct(h) = conds(ci).n(h)
                        
                        ' Debug output for verification
                        If h = 1 Then  ' Only show 1-month for brevity
                            Debug.Print conds(ci).label & " H1: $" & Format(ruleAvg, "#,##0") & " vs Base $" & Format(baseCaseAvgMonthly(h), "#,##0") & " = " & Format(pctDiff, "+0.0%;-0.0%")
                        End If
                    Else
                        conds(ci).SumPct(h) = 0
                        conds(ci).CntPct(h) = 1
                    End If
                End If
            End If
        Next h
    Next ci
End Sub



' Helper function to check if array is initialized
Private Function IsArrayInitialized(arr As Variant) As Boolean
    On Error GoTo NotInitialized
    Dim ub As Long: ub = UBound(arr)
    IsArrayInitialized = True
    Exit Function
NotInitialized:
    IsArrayInitialized = False
End Function



' Segment processing functions removed for performance

' ===== ENHANCED OUTPUT HELPERS =====

Private Sub CreateFormattedHeader(ws As Worksheet, ByVal ruleCount As Long, ByVal strategyCount As Long, _
                                  ByVal dateType As String, ByVal monthCount As Long)
    ' Enhanced title and summary information
    With ws
        .Range("A1").value = "ELIGIBILITY RULES BACKTESTING ANALYSIS"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(68, 114, 196)
        .Range("A1:M1").Merge
        .Range("A1").HorizontalAlignment = xlCenter
        
        .Range("A2").value = "Forward-Looking Performance Analysis by Rule"
        .Range("A2").Font.Size = 12
        .Range("A2").Font.Italic = True
        .Range("A2").Font.Color = RGB(89, 89, 89)
        .Range("A2:M2").Merge
        .Range("A2").HorizontalAlignment = xlCenter
        
        .Range("A3").value = "Generated: " & Format(Now(), "yyyy-mm-dd hh:mm") & " | " & _
                            "Eligibility Date: " & dateType & " | " & _
                            "Rules Tested: " & ruleCount & " | " & _
                            "Strategies Analyzed: " & strategyCount & " | " & _
                            "Evaluation Periods: " & monthCount
        .Range("A3").Font.Size = 10
        .Range("A3").Font.Color = RGB(89, 89, 89)
        .Range("A3:M3").Merge
        .Range("A3").HorizontalAlignment = xlCenter
        
        ' Add explanatory note
        .Range("A4").value = "Note: Each rule is evaluated at month-end. Forward performance measured over next 1-12 months."
        .Range("A4").Font.Size = 9
        .Range("A4").Font.Italic = True
        .Range("A4").Font.Color = RGB(112, 112, 112)
        .Range("A4:M4").Merge
        .Range("A4").HorizontalAlignment = xlCenter
        
        ' Column headers starting at row 5
        WriteDataHeader ws
        
        ' Add separator line above headers
        .Range("A5:M5").Borders(xlEdgeTop).Weight = xlMedium
        .Range("A5:M5").Borders(xlEdgeTop).Color = RGB(68, 114, 196)
    End With
End Sub

' ===== 8. UPDATED WriteDataHeader with new column label =====
Private Sub WriteDataHeader(ws As Worksheet)
    Dim h As Long, col As Long
    
    ' Clear any existing content first
    ws.Range("A5:ZZ7").Clear
    
    ' Write rule column headers (no merging here)
    ws.Cells(5, 1).value = "Eligibility Rule"
    ws.Cells(6, 1).value = "(Applied at Month-End)"
    ws.Cells(7, 1).value = ""
    
    ' Write horizon headers without merging first
    col = 2
    For h = 1 To MAX_HORIZON
        ' Row 7: Individual column headers (write these first)
        ws.Cells(7, col).value = "N"
        ws.Cells(7, col + 1).value = "Win%"
        ws.Cells(7, col + 2).value = "$/Month"
        ws.Cells(7, col + 3).value = "vs Base"
        
        col = col + 4
    Next h
    
    ' Now add the merged headers on top
    Application.DisplayAlerts = False
    On Error Resume Next
    
    col = 2
    For h = 1 To MAX_HORIZON
        ' Only merge if cells are not already merged
        If Not ws.Range(ws.Cells(5, col), ws.Cells(5, col + 3)).MergeCells Then
            ws.Range(ws.Cells(5, col), ws.Cells(5, col + 3)).Merge
            If h = 1 Then
                ws.Cells(5, col).value = h & " Month Forward"
            Else
                ws.Cells(5, col).value = h & " Months Forward"
            End If
            ws.Cells(5, col).HorizontalAlignment = xlCenter
        End If
        
        If Not ws.Range(ws.Cells(6, col), ws.Cells(6, col + 3)).MergeCells Then
            ws.Range(ws.Cells(6, col), ws.Cells(6, col + 3)).Merge
            ws.Cells(6, col).value = "Forward Performance Statistics"
            ws.Cells(6, col).HorizontalAlignment = xlCenter
        End If
        
        col = col + 4
    Next h
    
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Format headers
    With ws.Range("A5:A7")
        .Font.Bold = True
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Format other header rows
    Dim lastCol As Long: lastCol = 1 + 4 * MAX_HORIZON
    With ws.Range(ws.Cells(5, 2), ws.Cells(5, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
    End With
    
    With ws.Range(ws.Cells(6, 2), ws.Cells(6, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(149, 179, 215)
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .Font.Size = 10
    End With
    
    With ws.Range(ws.Cells(7, 2), ws.Cells(7, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .Font.Size = 9
    End With
End Sub

' ===== 7. UPDATED FormatResultsTable for base case freeze =====
Private Sub FormatResultsTable(ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    Dim lastCol As Long: lastCol = 1 + 4 * MAX_HORIZON
    
    ' Add professional borders
    With ws.Range(ws.Cells(5, 1), ws.Cells(endRow, lastCol))
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders.Color = RGB(89, 89, 89)
    End With
    
    ' Format data columns with appropriate number formats
    Dim col As Long, dataStartRow As Long
    dataStartRow = 8  ' Data starts at row 8 now (after 3-row header)
    
    For col = 2 To lastCol Step 4
        ' Sample Count columns (N) - whole numbers
        ws.Range(ws.Cells(dataStartRow, col), ws.Cells(endRow, col)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(dataStartRow, col), ws.Cells(endRow, col)).HorizontalAlignment = xlCenter
    Next col
    
    For col = 3 To lastCol Step 4
        ' Success Rate columns (Win%) - percentage with color coding
        With ws.Range(ws.Cells(dataStartRow, col), ws.Cells(endRow, col))
            .NumberFormat = "0.0%"
            .HorizontalAlignment = xlCenter
            
            ' Enhanced color scale for win rates
            With .FormatConditions.AddColorScale(3)
                .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
                .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 69, 58)  ' Red
                .ColorScaleCriteria(2).Type = xlConditionValueNumber
                .ColorScaleCriteria(2).value = 0.5
                .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 214, 10)  ' Yellow
                .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
                .ColorScaleCriteria(3).FormatColor.Color = RGB(52, 199, 89)   ' Green
            End With
        End With
    Next col
    
    For col = 4 To lastCol Step 4
        ' Average Return columns ($/Month) - currency format
        With ws.Range(ws.Cells(dataStartRow, col), ws.Cells(endRow, col))
            .NumberFormat = "$#,##0_);($#,##0)"
            .HorizontalAlignment = xlRight
            
            ' Color code positive/negative values
            With .FormatConditions.AddColorScale(2)
                .ColorScaleCriteria(1).Type = xlConditionValueNumber
                .ColorScaleCriteria(1).value = 0
                .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 69, 58)  ' Red for negative
                .ColorScaleCriteria(2).Type = xlConditionValueHighestValue
                .ColorScaleCriteria(2).FormatColor.Color = RGB(52, 199, 89)   ' Green for positive
            End With
        End With
    Next col
    
    For col = 5 To lastCol Step 4
        ' vs Base Case columns (% Change) - percentage
        With ws.Range(ws.Cells(dataStartRow, col), ws.Cells(endRow, col))
            .NumberFormat = "+0.0%;-0.0%;0.0%"
            .HorizontalAlignment = xlCenter
            
            ' Color code relative performance vs base case
            With .FormatConditions.AddColorScale(2)
                .ColorScaleCriteria(1).Type = xlConditionValueNumber
                .ColorScaleCriteria(1).value = 0
                .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 69, 58)  ' Red for underperformance
                .ColorScaleCriteria(2).Type = xlConditionValueHighestValue
                .ColorScaleCriteria(2).FormatColor.Color = RGB(52, 199, 89)   ' Green for outperformance
            End With
        End With
    Next col
    
    ' Alternate row colors for better readability
    Dim r As Long
    For r = dataStartRow To endRow Step 2
        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(248, 248, 248)
    Next r
    
    ' Highlight rule name column
    With ws.Range(ws.Cells(dataStartRow, 1), ws.Cells(endRow, 1))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    
    ' *** FREEZE PANES ON BASE CASE (row 9 instead of B8) ***
    ws.Range("B9").Select  ' Base case is at row 8, so freeze after it
    ActiveWindow.FreezePanes = True
    
    ' Auto-fit and set column widths appropriately
    ws.Columns("A:A").ColumnWidth = 50  ' Wider for longer rule names
    
    ' Set consistent widths for data columns
    For col = 2 To lastCol Step 4
        ws.Columns(col).ColumnWidth = 8    ' Sample Count
        ws.Columns(col + 1).ColumnWidth = 8  ' Win%
        ws.Columns(col + 2).ColumnWidth = 12 ' Avg Return (currency)
        ws.Columns(col + 3).ColumnWidth = 12 ' vs Base Case (wider for new label)
    Next col
End Sub


Private Sub AddDeleteButton(ws As Worksheet)
    ' Add a delete button to the worksheet
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("O3").left, ws.Range("O3").top, 100, 20)
    btn.Caption = "Delete This Sheet"
    btn.Font.Size = 9
    btn.OnAction = "DeleteCurrentSheet"
End Sub

' Button click handler
Public Sub DeleteCurrentSheet()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to delete the current worksheet '" & activeSheet.name & "'?", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete")
    
    If response = vbYes Then
        Dim sheetName As String: sheetName = activeSheet.name
        Application.DisplayAlerts = False
        activeSheet.Delete
        Application.DisplayAlerts = True
        MsgBox "Worksheet '" & sheetName & "' has been deleted.", vbInformation
    End If
End Sub

Private Sub OutputSegments(sheetName As String, segLabel As String, _
                           dictN As Object, dictW As Object, dictSum As Object, dictSumPct As Object, dictCntPct As Object, _
                           wsAfter As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=wsAfter): ws.name = sheetName

    WriteHeader ws
    Dim rowOut As Long: rowOut = 2
    Dim key As Variant
    For Each key In dictN.keys
        ws.Cells(rowOut, 1).value = "*** " & segLabel & ": " & CStr(key) & " ***"
        ws.rows(rowOut).Font.Bold = True
        rowOut = rowOut + 1

        Dim arrN As Variant, arrW As Variant, arrSum As Variant, arrSumPct As Variant, arrCnt As Variant
        arrN = dictN(key): arrW = dictW(key): arrSum = dictSum(key): arrSumPct = dictSumPct(key): arrCnt = dictCntPct(key)

        Dim condIdx As Long
        For condIdx = 1 To UBound(arrN, 1)
            WriteConditionRowFromArrays ws, rowOut, GetCondLabel(condIdx), _
                                        arrN, arrW, arrSum, arrSumPct, arrCnt, condIdx
            rowOut = rowOut + 1
        Next condIdx
    Next key

    FormatPercentColumns ws, 2, rowOut - 1
    ws.Columns.AutoFit
End Sub

' ===== 6. UPDATED WriteConditionRow for better formatting =====
Private Sub WriteConditionRow(ws As Worksheet, ByVal r As Long, ByRef cs As CondStat)
    ' Enhanced formatting for base case
    If cs.label = "Baseline (All Eligible)" Then
        ws.Cells(r, 1).value = "* " & cs.label & " *"  ' Use stars instead
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Interior.Color = RGB(255, 235, 156)  ' Light gold background
        ws.Cells(r, 1).Font.Color = RGB(0, 0, 0)
    Else
        ws.Cells(r, 1).value = cs.label
        ' Indent OOS rules slightly
        If InStr(1, cs.label, "AND OOS > 0", vbTextCompare) > 0 Then
            ws.Cells(r, 1).value = " | " & cs.label
            ws.Cells(r, 1).Font.Italic = True
        End If
    End If
    
    Dim c As Long, h As Long, n As Long
    Dim wr As Double, avgMonthlyP As Double, avgPct As Double

    c = 2
    For h = LBound(cs.n) To UBound(cs.n)
        n = cs.n(h)
        wr = SafeRate(cs.W(h), n)
        avgMonthlyP = SafeDivD(cs.SumPnL(h), n * h)
        
        ' For base case, show 0% change; for others, show change vs base case
        If cs.label = "Baseline (All Eligible)" Then
            avgPct = 0#  ' Base case is always 0% change from itself
        Else
            avgPct = SafeDivD(cs.SumPct(h), cs.CntPct(h))
        End If

        ws.Cells(r, c).value = n: c = c + 1
        ws.Cells(r, c).value = wr: c = c + 1
        ws.Cells(r, c).value = avgMonthlyP: c = c + 1
        ws.Cells(r, c).value = avgPct: c = c + 1
    Next
End Sub

Private Sub WriteConditionRowFromArrays(ws As Worksheet, ByVal r As Long, ByVal label As String, _
                                        ByRef aN As Variant, ByRef aW As Variant, _
                                        ByRef aSum As Variant, ByRef aSumPct As Variant, ByRef aCnt As Variant, _
                                        ByVal condIdx As Long)
    ws.Cells(r, 1).value = label
    Dim c As Long, h As Long, n As Long
    Dim wr As Double, avgMonthlyP As Double, avgPct As Double

    c = 2
    For h = 1 To MAX_HORIZON
        n = aN(condIdx, h)
        wr = SafeRate(aW(condIdx, h), n)
        
        ' Fixed: Average monthly profit = total profit / (count * horizon)
        avgMonthlyP = SafeDivD(aSum(condIdx, h), n * h)
        
        avgPct = SafeDivD(aSumPct(condIdx, h), aCnt(condIdx, h))

        ws.Cells(r, c).value = n:              c = c + 1
        ws.Cells(r, c).value = wr:             c = c + 1
        ws.Cells(r, c).value = avgMonthlyP:    c = c + 1
        ws.Cells(r, c).value = avgPct:         c = c + 1
    Next
End Sub

Private Function GetCondLabel(ByVal condIdx As Long) As String
    ' This function won't be used in the flexible system since labels come from rules
    ' But keeping it for compatibility
    GetCondLabel = "Condition " & condIdx
End Function

' ===== STATUS INCLUDE PARSING/MATCH =====

Private Function ReadIncludeList(nm As String, fallbackList As String) As Collection
    Dim result As New Collection
    On Error GoTo UseFallback
    Dim rg As Range: Set rg = ThisWorkbook.Names(nm).RefersToRange
    If rg.Cells.count > 1 Then
        Dim cell As Range
        For Each cell In rg.Cells
            AddIfNotEmpty result, LCase$(Trim$(CStr(cell.value)))
        Next
    Else
        Dim raw As String: raw = CStr(rg.value)
        AddDelimitedItems result, raw
    End If
    If result.count = 0 Then GoTo UseFallback
    Set ReadIncludeList = result
    Exit Function
UseFallback:
    Dim tmp As New Collection
    AddDelimitedItems tmp, fallbackList
    Set ReadIncludeList = tmp
End Function

Private Sub AddDelimitedItems(ByRef coll As Collection, ByVal text As String)
    Dim s As String: s = LCase$(text)
    Dim parts() As String: parts = SplitMulti(s, ",;|")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        AddIfNotEmpty coll, Trim$(parts(i))
    Next
End Sub

Private Sub AddIfNotEmpty(ByRef coll As Collection, ByVal itemText As String)
    If Len(itemText) > 0 Then coll.Add itemText
End Sub

Private Function SplitMulti(ByVal s As String, ByVal delims As String) As Variant
    Dim i As Long, ch As String
    For i = 1 To Len(delims)
        ch = mid$(delims, i, 1)
        s = Replace$(s, ch, vbLf)
    Next
    SplitMulti = Split(s, vbLf)
End Function

Private Function StatusIsIncluded(ByVal statusText As String, ByRef includeList As Collection) As Boolean
    Dim s As String: s = LCase$(Trim$(statusText))
    If Len(s) = 0 Then Exit Function
    Dim itm As Variant
    ' exact match or substring match
    For Each itm In includeList
        If s = itm Then StatusIsIncluded = True: Exit Function
        If InStr(1, s, CStr(itm), vbTextCompare) > 0 Then StatusIsIncluded = True: Exit Function
    Next
End Function

' ===== GENERIC HELPERS =====

Private Function FindHeaderCol(ws As Worksheet, headerText As String) As Long
    Dim lc As Long: lc = ws.Cells(1, ws.Columns.count).End(xlToLeft).column
    Dim c As Long
    For c = 1 To lc
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), Trim$(headerText), vbTextCompare) = 0 Then
            FindHeaderCol = c: Exit Function
        End If
    Next
End Function

Private Function FindHeaderColAlt(ws As Worksheet, ByVal candidates As Variant) As Long
    Dim i As Long, col As Long
    If IsArray(candidates) Then
        For i = LBound(candidates) To UBound(candidates)
            col = FindHeaderCol(ws, CStr(candidates(i)))
            If col > 0 Then FindHeaderColAlt = col: Exit Function
        Next
    End If
    FindHeaderColAlt = 0
End Function

Private Sub BuildMonthWindows(ByRef data As Variant, _
                              ByRef monthStart() As Long, ByRef monthEnd() As Long, ByRef mCount As Long)
    Dim lastRow As Long, r As Long, curStart As Long
    lastRow = UBound(data, 1): curStart = 2
    Dim bufS() As Long, bufE() As Long, n As Long
    ReDim bufS(1 To 1024): ReDim bufE(1 To 1024): n = 0

    For r = 2 To lastRow - 1
        Dim d As Date, d2 As Date
        d = CDate(data(r, 1)): d2 = CDate(data(r + 1, 1))
        If Month(d) <> Month(d2) Or Year(d) <> Year(d2) Then
            n = n + 1
            If n > UBound(bufS) Then
                ReDim Preserve bufS(1 To n * 2)
                ReDim Preserve bufE(1 To n * 2)
            End If
            bufS(n) = curStart: bufE(n) = r
            curStart = r + 1
        End If
    Next
    n = n + 1
    If n > UBound(bufS) Then
        ReDim Preserve bufS(1 To n)
        ReDim Preserve bufE(1 To n)
    End If
    bufS(n) = curStart: bufE(n) = lastRow

    ReDim monthStart(1 To n): ReDim monthEnd(1 To n)
    Dim i As Long
    For i = 1 To n
        monthStart(i) = bufS(i): monthEnd(i) = bufE(i)
    Next
    mCount = n
End Sub

Private Function FirstRowOnOrAfterDate(ByRef data As Variant, ByVal r1 As Long, ByVal r2 As Long, ByVal target As Date) As Long
    Dim r As Long
    For r = r1 To r2
        If CDate(data(r, 1)) >= target Then FirstRowOnOrAfterDate = r: Exit Function
    Next
End Function

Private Function SumPnLBetweenRows_Eligible(ByRef data As Variant, ByVal col As Long, ByVal eligibleBegin As Date, _
                                       ByVal r1 As Long, ByVal r2 As Long) As Double
    Dim s As Double, r As Long
    For r = r1 To r2
        If CDate(data(r, 1)) >= eligibleBegin Then
            If IsNumeric(data(r, col)) Then s = s + CDbl(data(r, col))
        End If
    Next
    SumPnLBetweenRows_Eligible = s
End Function

Private Function SumPnLBetweenDates_Eligible(ByRef data As Variant, ByVal col As Long, ByVal eligibleBegin As Date, _
                                        ByVal dStart As Date, ByVal dEnd As Date) As Double
    Dim s As Double, r As Long, d As Date
    If dStart < eligibleBegin Then dStart = eligibleBegin
    For r = 2 To UBound(data, 1)
        d = CDate(data(r, 1))
        If d >= dStart Then
            If d > dEnd Then Exit For
            If IsNumeric(data(r, col)) Then s = s + CDbl(data(r, col))
        End If
    Next
    SumPnLBetweenDates_Eligible = s
End Function

Private Function MaxDrawdownUpToRow_Eligible(ByRef data As Variant, ByVal col As Long, ByVal eligibleBegin As Date, ByVal endRow As Long) As Double
    Dim eq As Double, peak As Double, dd As Double, r As Long, d As Date
    eq = 0#: peak = 0#: dd = 0#
    For r = 2 To endRow
        d = CDate(data(r, 1))
        If d >= eligibleBegin Then
            If IsNumeric(data(r, col)) Then eq = eq + CDbl(data(r, col))
            If eq > peak Then peak = eq
            If (peak - eq) > dd Then dd = (peak - eq)
        End If
    Next
    MaxDrawdownUpToRow_Eligible = dd
End Function

Private Function NzNamed(ByVal nm As String, ByVal defVal As Variant) As Variant
    On Error GoTo EH
    NzNamed = ThisWorkbook.Names(nm).RefersToRange.value
    Exit Function
EH:
    NzNamed = defVal
End Function

Private Function SafeRate(ByVal wins As Long, ByVal n As Long) As Double
    If n <= 0 Then SafeRate = 0# Else SafeRate = wins / n
End Function

Private Function SafeDivD(ByVal num As Double, ByVal denom As Double) As Double
    If denom = 0# Then
        SafeDivD = 0#
    Else
        SafeDivD = num / denom
    End If
End Function

Private Function BToI(ByVal b As Boolean) As Long
    If b Then BToI = 1 Else BToI = 0
End Function

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function


' ===== PRE-CALCULATION SYSTEM =====

Private Sub BuildStrategyMonthlyData(ByRef data As Variant, ByRef colMap As Object, _
                                   ByRef eligibilityBeg As Object, ByRef mStart() As Long, _
                                   ByRef mEnd() As Long, ByVal mCount As Long, _
                                   ByRef strategyMonthlyData As Object)
    
    Application.StatusBar = "Pre-calculating monthly data..."
    
    ' Initialize the dictionary
    Set strategyMonthlyData = CreateObject("Scripting.Dictionary")
    
    Dim sName As Variant, col As Long, beg As Date
    For Each sName In eligibilityBeg.keys
        If colMap.Exists(CStr(sName)) Then
            col = colMap(CStr(sName))
            beg = eligibilityBeg(CStr(sName))
            
            ' Create array for this strategy (1-based indexing to match month numbers)
            Dim monthlyPnL() As Double
            ReDim monthlyPnL(1 To mCount)
            
            ' Calculate PnL for each month
            Dim j As Long
            For j = 1 To mCount
                monthlyPnL(j) = SumPnLBetweenRows_Eligible(data, col, beg, mStart(j), mEnd(j))
            Next j
            
            ' Store in dictionary
            strategyMonthlyData(CStr(sName)) = monthlyPnL
        End If
    Next sName
End Sub

Private Function GetTrailingSumFromArray(ByRef monthlyArray As Variant, ByVal currentMonth As Long, _
                                       ByVal trailingMonths As Long) As Double
    ' Sum the last N months ending at currentMonth
    Dim sum As Double, i As Long, monthIdx As Long
    sum = 0#
    
    For i = 0 To trailingMonths - 1
        monthIdx = currentMonth - i
        If monthIdx >= LBound(monthlyArray) And monthIdx <= UBound(monthlyArray) Then
            sum = sum + monthlyArray(monthIdx)
        End If
    Next i
    
    GetTrailingSumFromArray = sum
End Function

Private Function GetForwardSumFromArray(ByRef monthlyArray As Variant, ByVal startMonth As Long, _
                                      ByVal forwardMonths As Long) As Double
    ' Sum the next N months starting from startMonth
    Dim sum As Double, i As Long, monthIdx As Long
    sum = 0#
    
    For i = 0 To forwardMonths - 1
        monthIdx = startMonth + i
        If monthIdx >= LBound(monthlyArray) And monthIdx <= UBound(monthlyArray) Then
            sum = sum + monthlyArray(monthIdx)
        End If
    Next i
    
    GetForwardSumFromArray = sum
End Function

Private Function EvaluateRuleFast(ByRef rule As EligibilityRule, _
                                 ByVal p1 As Double, ByVal p3 As Double, ByVal p6 As Double, _
                                 ByVal p9 As Double, ByVal p12 As Double, _
                                 ByVal sName As String, ByRef monthlyPnL As Variant, _
                                 ByVal currentMonth As Long, ByRef expectedReturns As Object) As Boolean
    
    Dim efficiencyRatio As Double
    efficiencyRatio = CDbl(NzNamed(NM_EFFICIENCY_RATIO, 1#))
    Dim oosCheck1 As Double: oosCheck1 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
    Select Case rule.RuleType
        ' =====================================
        ' BASELINE RULES
        ' =====================================
        Case "BASELINE"
            EvaluateRuleFast = True
            
        Case "OOS_PROFITABLE"
            Dim oosTotal As Double
            oosTotal = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            EvaluateRuleFast = (oosTotal > 0)
            
        Case "OOS_PROFITABLE_AND_12M"
            Dim oosTotal12 As Double
            oosTotal12 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            EvaluateRuleFast = (oosTotal12 > 0 And p12 > 0)
            
        Case "OOS_PROFITABLE_AND_6M"
            Dim oosTotal6 As Double
            oosTotal6 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            EvaluateRuleFast = (oosTotal6 > 0 And p6 > 0)
            
        ' =====================================
        ' SIMPLE PERIOD RULES (Original)
        ' =====================================
        Case "SIMPLE_POSITIVE"
            Select Case CLng(rule.Param1)
                Case 1: EvaluateRuleFast = (p1 > 0)
                Case 3: EvaluateRuleFast = (p3 > 0)
                Case 6: EvaluateRuleFast = (p6 > 0)
                Case 9: EvaluateRuleFast = (p9 > 0)
                Case 12: EvaluateRuleFast = (p12 > 0)
            End Select
            
        Case "SIMPLE_NEGATIVE"
            Select Case CLng(rule.Param1)
                Case 1: EvaluateRuleFast = (p1 < 0)
                Case 3: EvaluateRuleFast = (p3 < 0)
                Case 6: EvaluateRuleFast = (p6 < 0)
                Case 9: EvaluateRuleFast = (p9 < 0)
                Case 12: EvaluateRuleFast = (p12 < 0)
            End Select
            
        ' =====================================
        ' SIMPLE PERIOD RULES (OOS Variants)
        ' =====================================
        Case "SIMPLE_POSITIVE_OOS"
            
            Select Case CLng(rule.Param1)
                Case 1: EvaluateRuleFast = (p1 > 0 And oosCheck1 > 0)
                Case 3: EvaluateRuleFast = (p3 > 0 And oosCheck1 > 0)
                Case 6: EvaluateRuleFast = (p6 > 0 And oosCheck1 > 0)
                Case 9: EvaluateRuleFast = (p9 > 0 And oosCheck1 > 0)
                Case 12: EvaluateRuleFast = (p12 > 0 And oosCheck1 > 0)
            End Select
        
        Case "SIMPLE_NEGATIVE_OOS"

            Select Case CLng(rule.Param1)
                Case 1: EvaluateRuleFast = (p1 < 0 And oosCheck1 > 0)
                Case 3: EvaluateRuleFast = (p3 < 0 And oosCheck1 > 0)
                Case 6: EvaluateRuleFast = (p6 < 0 And oosCheck1 > 0)
                Case 9: EvaluateRuleFast = (p9 < 0 And oosCheck1 > 0)
                Case 12: EvaluateRuleFast = (p12 < 0 And oosCheck1 > 0)
            End Select
        
        ' =====================================
        ' CONSECUTIVE RULES (Original)
        ' =====================================
        Case "CONSECUTIVE"
            EvaluateRuleFast = EvaluateConsecutiveFromArray(monthlyPnL, currentMonth, CLng(rule.Param1))
            
        ' =====================================
        ' CONSECUTIVE RULES (OOS Variants)
        ' =====================================
        Case "CONSECUTIVE_OOS"
            Dim oosCheck2 As Double: oosCheck2 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            EvaluateRuleFast = (EvaluateConsecutiveFromArray(monthlyPnL, currentMonth, CLng(rule.Param1)) And oosCheck2 > 0)
            
        ' =====================================
        ' COUNT-BASED RULES (Original)
        ' =====================================
        Case "COUNT_POSITIVE"
            EvaluateRuleFast = EvaluateCountFromArray(monthlyPnL, currentMonth, CLng(rule.Param1), CLng(rule.Param2))
            
        ' =====================================
        ' COUNT-BASED RULES (OOS Variants)
        ' =====================================
        Case "COUNT_POSITIVE_OOS"
            Dim oosCheck3 As Double: oosCheck3 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            EvaluateRuleFast = (EvaluateCountFromArray(monthlyPnL, currentMonth, CLng(rule.Param1), CLng(rule.Param2)) And oosCheck3 > 0)
            
        ' =====================================
        ' THRESHOLD RULES (Original)
        ' =====================================
        Case "THRESHOLD_ANNUAL"
            EvaluateRuleFast = EvaluateThresholdRule(sName, rule.Param1, p1, p3, p6, p9, p12, expectedReturns, efficiencyRatio)
            
        ' =====================================
        ' THRESHOLD RULES (OOS Variants)
        ' =====================================
        Case "THRESHOLD_ANNUAL_OOS"
            Dim oosCheck4 As Double: oosCheck4 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            EvaluateRuleFast = (EvaluateThresholdRule(sName, rule.Param1, p1, p3, p6, p9, p12, expectedReturns, efficiencyRatio) And oosCheck4 > 0)
            
        ' =====================================
        ' MOMENTUM & ACCELERATION (Original)
        ' =====================================
        Case "MOMENTUM"
            EvaluateRuleFast = EvaluateMomentumFromArray(monthlyPnL, currentMonth, CLng(rule.Param1), CLng(rule.Param2))
            
        Case "MOMENTUM_AND_12M"
            EvaluateRuleFast = (EvaluateMomentumFromArray(monthlyPnL, currentMonth, CLng(rule.Param1), CLng(rule.Param2)) And p12 > 0)
            
        Case "ACCELERATION"
            EvaluateRuleFast = (p3 * 4 > p6 * 2)
            
        Case "ACCELERATION_AND_12M"
            EvaluateRuleFast = (p3 * 4 > p6 * 2 And p12 > 0)
            
        ' =====================================
        ' MOMENTUM & ACCELERATION (OOS Variants)
        ' =====================================
        Case "MOMENTUM_OOS"
            Dim oosCheck5 As Double: oosCheck5 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            EvaluateRuleFast = (EvaluateMomentumFromArray(monthlyPnL, currentMonth, CLng(rule.Param1), CLng(rule.Param2)) And oosCheck5 > 0)
            
        Case "ACCELERATION_OOS"
            Dim oosCheck6 As Double: oosCheck6 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            EvaluateRuleFast = (p3 * 4 > p6 * 2 And oosCheck6 > 0)
            
        ' =====================================
        ' RECOVERY RULES (Original)
        ' =====================================
        Case "RECOVERY"
            If rule.Param1 = 1 And rule.Param2 = 3 Then
                EvaluateRuleFast = (p1 > 0 And p3 <= 0)
            ElseIf rule.Param1 = 3 And rule.Param2 = 6 Then
                EvaluateRuleFast = (p3 > 0 And p6 <= 0)
            ElseIf rule.Param1 = 6 And rule.Param2 = 12 Then
                EvaluateRuleFast = (p6 > 0 And p12 <= 0)
            End If
            
        Case "RECOVERY_AND_12M"
            EvaluateRuleFast = (p3 > 0 And p6 <= 0 And p12 > 0)
            
        ' =====================================
        ' RECOVERY RULES (OOS Variants)
        ' =====================================
        Case "RECOVERY_OOS"
            Dim oosCheck7 As Double: oosCheck7 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            If rule.Param1 = 1 And rule.Param2 = 3 Then
                EvaluateRuleFast = (p1 > 0 And p3 <= 0 And oosCheck7 > 0)
            ElseIf rule.Param1 = 3 And rule.Param2 = 6 Then
                EvaluateRuleFast = (p3 > 0 And p6 <= 0 And oosCheck7 > 0)
            ElseIf rule.Param1 = 6 And rule.Param2 = 12 Then
                EvaluateRuleFast = (p6 > 0 And p12 <= 0 And oosCheck7 > 0)
            End If
            
        ' =====================================
        ' OR COMBINATION RULES (Original)
        ' =====================================
        Case "OR_COMBO"
            If rule.Param1 = 3 And rule.Param2 = 6 Then
                EvaluateRuleFast = (p3 > 0 Or p6 > 0)
            End If
            
        Case "ANY_OF_3"
            EvaluateRuleFast = (p1 > 0 Or p3 > 0 Or p6 > 0)
            
        Case "ANY_OF_3_AND_12M"
            EvaluateRuleFast = ((p1 > 0 Or p3 > 0 Or p6 > 0) And p12 > 0)
            
        Case "ANY_OF_3_AND_12M_THRESHOLD"
            Dim threshold12 As Double
            If expectedReturns.Exists(sName) Then
                threshold12 = efficiencyRatio * CDbl(expectedReturns(sName))
                EvaluateRuleFast = ((p1 > 0 Or p3 > 0 Or p6 > 0) And p12 > threshold12)
            Else
                EvaluateRuleFast = False
            End If
            
        ' =====================================
        ' OR COMBINATION RULES (OOS Variants)
        ' =====================================
        Case "ANY_OF_3_OOS"
            Dim oosCheck8 As Double: oosCheck8 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            EvaluateRuleFast = ((p1 > 0 Or p3 > 0 Or p6 > 0) And oosCheck8 > 0)
            
        ' =====================================
        ' AND COMBINATION RULES (Original)
        ' =====================================
        Case "AND_COMBO"
            If rule.Param1 = 1 And rule.Param2 = 12 Then
                EvaluateRuleFast = (p1 > 0 And p12 > 0)
            ElseIf rule.Param1 = 3 And rule.Param2 = 6 Then
                EvaluateRuleFast = (p3 > 0 And p6 > 0)
            ElseIf rule.Param1 = 3 And rule.Param2 = 9 Then
                EvaluateRuleFast = (p3 > 0 And p9 > 0)
            ElseIf rule.Param1 = 3 And rule.Param2 = 12 Then
                EvaluateRuleFast = (p3 > 0 And p12 > 0)
            ElseIf rule.Param1 = 6 And rule.Param2 = 12 Then
                EvaluateRuleFast = (p6 > 0 And p12 > 0)
            End If
            
        Case "ALL_OF_3"
            EvaluateRuleFast = (p1 > 0 And p3 > 0 And p6 > 0)
            
        Case "AT_LEAST_2_OF_3"
            Dim count As Long: count = 0
            If p3 > 0 Then count = count + 1
            If p6 > 0 Then count = count + 1
            If p12 > 0 Then count = count + 1
            EvaluateRuleFast = (count >= 2)
            
        Case "AT_LEAST_3_OF_4"
            Dim count4 As Long: count4 = 0
            If p3 > 0 Then count4 = count4 + 1
            If p6 > 0 Then count4 = count4 + 1
            If p9 > 0 Then count4 = count4 + 1
            If p12 > 0 Then count4 = count4 + 1
            EvaluateRuleFast = (count4 >= 3)
            
        ' =====================================
        ' AND COMBINATION RULES (OOS Variants)
        ' =====================================
        Case "AND_COMBO_OOS"
            Dim oosCheck9 As Double: oosCheck9 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            If rule.Param1 = 3 And rule.Param2 = 6 Then
                EvaluateRuleFast = (p3 > 0 And p6 > 0 And oosCheck9 > 0)
            ElseIf rule.Param1 = 3 And rule.Param2 = 12 Then
                EvaluateRuleFast = (p3 > 0 And p12 > 0 And oosCheck9 > 0)
            ElseIf rule.Param1 = 6 And rule.Param2 = 12 Then
                EvaluateRuleFast = (p6 > 0 And p12 > 0 And oosCheck9 > 0)
            End If
            
        Case "ALL_OF_3_OOS"
            Dim oosCheck10 As Double: oosCheck10 = CalculateOOSTotal(sName, currentMonth, monthlyPnL)
            EvaluateRuleFast = (p1 > 0 And p3 > 0 And p6 > 0 And oosCheck10 > 0)
            
        ' =====================================
        ' COMPLEX COMBINATION RULES
        ' =====================================
        Case "COMPLEX_COMBO"
            If rule.Param1 = 1 Then ' (3M OR 6M) AND 12M
                EvaluateRuleFast = ((p3 > 0 Or p6 > 0) And p12 > 0)
            End If
            
        ' =====================================
        ' DUAL THRESHOLD RULES
        ' =====================================
        Case "DUAL_THRESHOLD"
            If expectedReturns.Exists(sName) Then
                Dim expected As Double: expected = CDbl(expectedReturns(sName))
                Dim threshold1 As Double: threshold1 = efficiencyRatio * expected
                Dim actual1 As Double, actual2 As Double
                
                If rule.Param1 = 12 Then actual1 = p12
                If rule.Param2 = 6 Then actual2 = p6 * 2
                If rule.Param2 = 3 Then actual2 = p3 * 4
                
                EvaluateRuleFast = (actual1 > threshold1 And actual2 > threshold1)
            Else
                EvaluateRuleFast = False
            End If
            
        Case "THRESHOLD_OR_COMBO"
            If expectedReturns.Exists(sName) Then
                Dim expectedVal As Double: expectedVal = CDbl(expectedReturns(sName))
                Dim thresholdVal As Double: thresholdVal = efficiencyRatio * expectedVal
                Dim actual12 As Double: actual12 = p12
                Dim actual6Ann As Double: actual6Ann = p6 * 2
                Dim actual3Ann As Double: actual3Ann = p3 * 4
                
                EvaluateRuleFast = (actual12 > thresholdVal And (actual6Ann > thresholdVal Or actual3Ann > thresholdVal))
            Else
                EvaluateRuleFast = False
            End If
            
        ' =====================================
        ' SPECIAL CASE RULES
        ' =====================================
        Case "SPECIAL_6_NOT_12"
            EvaluateRuleFast = (p6 > 0 And p12 <= 0)
            
        Case "SPECIAL_12_NOT_3"
            EvaluateRuleFast = (p12 > 0 And p3 <= 0)
            
        ' =====================================
        ' DEFAULT CASE
        ' =====================================
        Case Else
            EvaluateRuleFast = False
    End Select
End Function



' Helper functions for array-based evaluations
Private Function EvaluateConsecutiveFromArray(ByRef monthlyPnL As Variant, ByVal currentMonth As Long, ByVal numMonths As Long) As Boolean
    Dim consecutive As Long, i As Long, monthIdx As Long
    consecutive = 0
    
    For i = 0 To numMonths - 1
        monthIdx = currentMonth - i
        If monthIdx >= LBound(monthlyPnL) And monthIdx <= UBound(monthlyPnL) Then
            If monthlyPnL(monthIdx) > 0 Then
                consecutive = consecutive + 1
            Else
                Exit For
            End If
        Else
            Exit For
        End If
    Next i
    EvaluateConsecutiveFromArray = (consecutive >= numMonths)
End Function

Private Function EvaluateCountFromArray(ByRef monthlyPnL As Variant, ByVal currentMonth As Long, ByVal minPositive As Long, ByVal totalMonths As Long) As Boolean
    Dim positiveCount As Long, i As Long, monthIdx As Long
    positiveCount = 0
    
    For i = 0 To totalMonths - 1
        monthIdx = currentMonth - i
        If monthIdx >= LBound(monthlyPnL) And monthIdx <= UBound(monthlyPnL) Then
            If monthlyPnL(monthIdx) > 0 Then positiveCount = positiveCount + 1
        End If
    Next i
    EvaluateCountFromArray = (positiveCount >= minPositive)
End Function


Private Function EvaluateMomentumFromArray(ByRef monthlyPnL As Variant, ByVal currentMonth As Long, ByVal recent As Long, ByVal previous As Long) As Boolean
    Dim recentSum As Double, prevSum As Double, i As Long, monthIdx As Long
    
    ' Recent period sum
    For i = 0 To recent - 1
        monthIdx = currentMonth - i
        If monthIdx >= LBound(monthlyPnL) And monthIdx <= UBound(monthlyPnL) Then
            recentSum = recentSum + monthlyPnL(monthIdx)
        End If
    Next i
    
    ' Previous period sum
    For i = recent To recent + previous - 1
        monthIdx = currentMonth - i
        If monthIdx >= LBound(monthlyPnL) And monthIdx <= UBound(monthlyPnL) Then
            prevSum = prevSum + monthlyPnL(monthIdx)
        End If
    Next i
    
    EvaluateMomentumFromArray = (recentSum > prevSum)
End Function
Private Function EvaluateThresholdRule(ByVal sName As String, ByVal period As Double, ByVal p1 As Double, ByVal p3 As Double, ByVal p6 As Double, ByVal p9 As Double, ByVal p12 As Double, ByRef expectedReturns As Object, ByVal efficiencyRatio As Double) As Boolean
    If expectedReturns.Exists(sName) Then
        Dim expectedAnnual As Double, threshold As Double, actualAnnual As Double
        expectedAnnual = CDbl(expectedReturns(sName))
        threshold = efficiencyRatio * expectedAnnual
        
        Select Case CLng(period)
            Case 3: actualAnnual = p3 * 4
            Case 6: actualAnnual = p6 * 2
            Case 12: actualAnnual = p12
        End Select
        EvaluateThresholdRule = (actualAnnual > threshold)
    Else
        EvaluateThresholdRule = False
    End If
End Function


Private Function CalculateOOSTotal(ByVal strategyName As String, ByVal currentMonth As Long, ByRef monthlyPnL As Variant) As Double
    Dim sum As Double, i As Long
    sum = 0#
    
    For i = LBound(monthlyPnL) To currentMonth
        If i <= UBound(monthlyPnL) Then
            sum = sum + monthlyPnL(i)
        End If
    Next i
    
    CalculateOOSTotal = sum
End Function


' ===== STEP 6: Add the BumpSegment helper function =====

Private Sub BumpSegment(ByRef dictN As Object, ByRef dictW As Object, ByRef dictSum As Object, _
                       ByRef dictSumPct As Object, ByRef dictCntPct As Object, _
                       ByVal segmentName As String, ByVal ruleIdx As Long, ByVal totalRules As Long, _
                       ByVal h As Long, ByVal win As Boolean, ByVal pfwd As Double, ByVal prev12 As Double)
    
    ' Initialize segment if it doesn't exist
    If Not dictN.Exists(segmentName) Then
        Dim arrN() As Long, arrW() As Long, arrSum() As Double, arrSumPct() As Double, arrCnt() As Long
        ReDim arrN(1 To totalRules, 1 To MAX_HORIZON)
        ReDim arrW(1 To totalRules, 1 To MAX_HORIZON)
        ReDim arrSum(1 To totalRules, 1 To MAX_HORIZON)
        ReDim arrSumPct(1 To totalRules, 1 To MAX_HORIZON)
        ReDim arrCnt(1 To totalRules, 1 To MAX_HORIZON)
        
        dictN(segmentName) = arrN
        dictW(segmentName) = arrW
        dictSum(segmentName) = arrSum
        dictSumPct(segmentName) = arrSumPct
        dictCntPct(segmentName) = arrCnt
    End If
    
    ' Get arrays and update statistics
    Dim aN As Variant, aW As Variant, aSum As Variant, aSumPct As Variant, aCnt As Variant
    aN = dictN(segmentName): aW = dictW(segmentName): aSum = dictSum(segmentName)
    aSumPct = dictSumPct(segmentName): aCnt = dictCntPct(segmentName)
    
    ' Update counters
    aN(ruleIdx, h) = aN(ruleIdx, h) + 1
    If win Then aW(ruleIdx, h) = aW(ruleIdx, h) + 1
    aSum(ruleIdx, h) = aSum(ruleIdx, h) + pfwd
    
    ' Update percentage tracking
    If Abs(prev12) > 0.0001 Then
        Dim fwdMonthlyAvg As Double, trailingMonthlyAvg As Double
        fwdMonthlyAvg = pfwd / h
        trailingMonthlyAvg = prev12 / 12
        aSumPct(ruleIdx, h) = aSumPct(ruleIdx, h) + (fwdMonthlyAvg / trailingMonthlyAvg - 1)
        aCnt(ruleIdx, h) = aCnt(ruleIdx, h) + 1
    End If
    
    ' Store back updated arrays
    dictN(segmentName) = aN: dictW(segmentName) = aW: dictSum(segmentName) = aSum
    dictSumPct(segmentName) = aSumPct: dictCntPct(segmentName) = aCnt
End Sub


' ===== FIX 2: Update OutputSegmentResults function =====

Private Sub OutputSegmentResults(ByVal sheetName As String, ByVal segLabel As String, _
                                ByRef dictN As Object, ByRef dictW As Object, ByRef dictSum As Object, _
                                ByRef dictSumPct As Object, ByRef dictCntPct As Object, _
                                ByVal wsAfter As Worksheet, ByVal totalRules As Long, ByRef rules() As EligibilityRule)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=wsAfter)
    ws.name = sheetName
    
    ' Color the tab pink
    ws.Tab.Color = RGB(255, 182, 193)  ' Light pink

    ' Create enhanced header for segment analysis
    CreateSegmentHeader ws, segLabel, dictN.count, totalRules
    
    Dim rowOut As Long: rowOut = 8  ' Start after header
    Dim key As Variant
    
    ' Sort segments by name for consistent output
    Dim sortedKeys() As String, keyCount As Long, i As Long, j As Long
    keyCount = dictN.count
    ReDim sortedKeys(1 To keyCount)
    
    i = 1
    For Each key In dictN.keys
        sortedKeys(i) = CStr(key)
        i = i + 1
    Next key
    
    ' Simple bubble sort for segment names
    For i = 1 To keyCount - 1
        For j = i + 1 To keyCount
            If sortedKeys(i) > sortedKeys(j) Then
                Dim temp As String: temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j): sortedKeys(j) = temp
            End If
        Next j
    Next i
    
    ' Output each segment
    For i = 1 To keyCount
        key = sortedKeys(i)
        
        ' Segment header - FIXED SYMBOL
        ws.Cells(rowOut, 1).value = "" & segLabel & ": " & CStr(key)
        ws.Cells(rowOut, 1).Font.Bold = True
        ws.Cells(rowOut, 1).Font.Size = 12
        ws.Cells(rowOut, 1).Interior.Color = RGB(79, 129, 189)
        ws.Cells(rowOut, 1).Font.Color = RGB(0, 32, 96)
        
        ' Merge across visible columns
        Dim lastCol As Long: lastCol = 1 + 4 * MAX_HORIZON
        Application.DisplayAlerts = False
        ws.Range(ws.Cells(rowOut, 1), ws.Cells(rowOut, lastCol)).Merge
        ws.Range(ws.Cells(rowOut, 1), ws.Cells(rowOut, lastCol)).HorizontalAlignment = xlCenter
        Application.DisplayAlerts = True
        rowOut = rowOut + 1

        ' Get arrays for this segment
        Dim arrN As Variant, arrW As Variant, arrSum As Variant, arrSumPct As Variant, arrCnt As Variant
        arrN = dictN(key): arrW = dictW(key): arrSum = dictSum(key)
        arrSumPct = dictSumPct(key): arrCnt = dictCntPct(key)

        ' Output each rule's results for this segment
        Dim condIdx As Long
        For condIdx = 1 To totalRules
            WriteSegmentConditionRow ws, rowOut, rules(condIdx).label, _
                                   arrN, arrW, arrSum, arrSumPct, arrCnt, condIdx
            rowOut = rowOut + 1
        Next condIdx
        
        ' Add separator row
        rowOut = rowOut + 1
    Next i

    ' Apply formatting
    FormatSegmentTable ws, 8, rowOut - 1
    ws.Columns.AutoFit
    ws.Activate
    ws.Range("A1").Select
End Sub

Private Sub WriteSegmentConditionRow(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                                    ByRef aN As Variant, ByRef aW As Variant, _
                                    ByRef aSum As Variant, ByRef aSumPct As Variant, ByRef aCnt As Variant, _
                                    ByVal condIdx As Long)
    ws.Cells(r, 1).value = "  " & label  ' Indent rule names
    Dim c As Long, h As Long, n As Long
    Dim wr As Double, avgMonthlyP As Double, avgPct As Double

    c = 2
    For h = 1 To MAX_HORIZON
        n = aN(condIdx, h)
        wr = SafeRate(aW(condIdx, h), n)
        avgMonthlyP = SafeDivD(aSum(condIdx, h), n * h)
        avgPct = SafeDivD(aSumPct(condIdx, h), aCnt(condIdx, h))

        ws.Cells(r, c).value = n: c = c + 1
        ws.Cells(r, c).value = wr: c = c + 1
        ws.Cells(r, c).value = avgMonthlyP: c = c + 1
        ws.Cells(r, c).value = avgPct: c = c + 1
    Next h
End Sub

Private Sub CreateSegmentHeader(ByVal ws As Worksheet, ByVal segmentType As String, _
                               ByVal segmentCount As Long, ByVal ruleCount As Long)
    With ws
        .Range("A1").value = segmentType & " ANALYSIS - ELIGIBILITY RULES"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(68, 114, 196)
        .Range("A1:M1").Merge
        .Range("A1").HorizontalAlignment = xlCenter
        
        .Range("A2").value = "Performance Breakdown by " & segmentType
        .Range("A2").Font.Size = 12
        .Range("A2").Font.Italic = True
        .Range("A2:M2").Merge
        .Range("A2").HorizontalAlignment = xlCenter
        
        .Range("A3").value = "Analysis Date: " & Format(Now(), "yyyy-mm-dd hh:mm") & " | " & _
                            segmentType & "s: " & segmentCount & " | Rules: " & ruleCount
        .Range("A3").Font.Size = 10
        .Range("A3:M3").Merge
        .Range("A3").HorizontalAlignment = xlCenter
    End With
    
    ' Use the same header structure as main analysis
    WriteDataHeader ws
End Sub


Private Sub FormatPercentColumns(ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    ' Apply percentage formatting to appropriate columns
    Dim lastCol As Long: lastCol = 1 + 4 * MAX_HORIZON
    Dim col As Long
    
    ' Format Win% columns
    For col = 3 To lastCol Step 4
        ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col)).NumberFormat = "0.0%"
    Next col
    
    ' Format % Change columns
    For col = 5 To lastCol Step 4
        ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col)).NumberFormat = "+0.0%;-0.0%;0.0%"
    Next col
    
    ' Format currency columns
    For col = 4 To lastCol Step 4
        ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col)).NumberFormat = "$#,##0_);($#,##0)"
    Next col
End Sub


' ===== FIX 4: Add missing WriteHeader function =====

Private Sub WriteHeader(ws As Worksheet)
    ' Simple header for legacy OutputSegments function
    ws.Cells(1, 1).value = "Eligibility Rules Analysis"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
End Sub


Public Sub ColorAnalysisTabsPink()
    ' Call this function to color all existing analysis tabs pink
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Rules_Analysis" Or _
           ws.name = "Sector_Analysis" Or _
           ws.name = "Symbol_Analysis" Or _
           InStr(1, ws.name, "EligRule", vbTextCompare) > 0 Then
            ws.Tab.Color = RGB(255, 182, 193)  ' Light pink
        End If
    Next ws
    
    MsgBox "Analysis tabs colored pink!", vbInformation
End Sub


' ===== FIX 4: Add helper function for worksheet count =====

Private Function GetWorksheetCount() As String
    Dim count As Long: count = 1  ' Main analysis sheet (Rules_Analysis)
    
    ' Check if Sector_Analysis sheet exists
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Sector_Analysis")
    On Error GoTo 0
    If Not ws Is Nothing Then count = count + 1
    
    ' Check if Symbol_Analysis sheet exists
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Symbol_Analysis")
    On Error GoTo 0
    If Not ws Is Nothing Then count = count + 1
    
    GetWorksheetCount = CStr(count)
End Function


' ===== FIX 3: Update FormatSegmentTable function =====

Private Sub FormatSegmentTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    ' Apply similar formatting as main results table but with segment-specific styling
    Dim lastCol As Long: lastCol = 1 + 4 * MAX_HORIZON
    
    ' Basic borders and formatting
    FormatResultsTable ws, startRow, endRow
     ' Unfreeze for segment sheets
    ActiveWindow.FreezePanes = False
    
    ' Additional formatting for segment headers - FIXED SYMBOL CHECK
    Dim r As Long
    For r = startRow To endRow
        If InStr(1, ws.Cells(r, 1).value, "?") > 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(79, 129, 189)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Font.Color = RGB(255, 255, 255)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Font.Bold = True
        End If
    Next r
End Sub



Private Sub DeleteAnalysisSheets()
    ' Delete all analysis sheets if they exist
    Dim sheetNames As Variant
    sheetNames = Array("Rules_Analysis", "Sector_Analysis", "Symbol_Analysis")
    
    Dim i As Long, ws As Worksheet
    Application.DisplayAlerts = False ' Suppress delete confirmation dialogs
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(sheetNames(i)))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ws.Delete
            Debug.Print "Deleted sheet: " & CStr(sheetNames(i))
        End If
    Next i
    
    Application.DisplayAlerts = True
End Sub

Attribute VB_Name = "Z_ExportFixtures"
Option Explicit

' ============================================================
' Export Golden Dataset Fixtures for Python v2 Regression Tests
'
' Run this ONCE after a full data import to capture the current
' state of all key sheets as CSV files.
'
' Output folder: Desktop\pt_fixtures\
' ============================================================

Public Sub ExportGoldenDataset(Optional Silent As Boolean = False)

    Dim exportPath As String
    exportPath = Environ("USERPROFILE") & "\Desktop\pt_fixtures\"

    ' Create output folder
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim exported As Long
    Dim skipped As Long
    exported = 0
    skipped = 0

    ' ── Sheets we want to export ──────────────────────────────
    ' Format: Array(sheetName, wasHidden)
    ' wasHidden is filled in by the macro — don't set it here
    Dim targets(8) As String
    targets(0) = "Summary"
    targets(1) = "DailyM2MEquity"
    targets(2) = "ClosedTradePNL"
    targets(3) = "Portfolio"
    targets(4) = "Walkforward Details"
    targets(5) = "PortfolioDailyM2M"
    targets(6) = "TotalPortfolioM2M"
    targets(7) = "LatestPositionData"
    targets(8) = "Strategies"

    Dim i As Long
    For i = 0 To UBound(targets)
        Dim sheetName As String
        sheetName = targets(i)

        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sheetName)
        On Error GoTo 0

        If ws Is Nothing Then
            ' Sheet doesn't exist — skip silently
            skipped = skipped + 1
            GoTo NextSheet
        End If

        ' Remember original visibility state
        Dim wasHidden As Boolean
        wasHidden = (ws.Visible <> xlSheetVisible)

        ' Unhide if needed
        If wasHidden Then
            ws.Visible = xlSheetVisible
        End If

        ' Copy sheet to a new temporary workbook and save as CSV
        On Error Resume Next
        ws.Copy
        If Err.Number = 0 Then
            ActiveWorkbook.SaveAs _
                Filename:=exportPath & sheetName & ".csv", _
                FileFormat:=xlCSV, _
                CreateBackup:=False
            ActiveWorkbook.Close SaveChanges:=False
            exported = exported + 1
        Else
            ' Copy failed (e.g. sheet has unsupported objects) — skip
            Err.Clear
            skipped = skipped + 1
        End If
        On Error GoTo 0

        ' Re-hide if it was hidden before
        If wasHidden Then
            ws.Visible = xlSheetHidden
        End If

NextSheet:
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not Silent Then
        MsgBox "Export complete!" & vbCrLf & vbCrLf & _
               "Exported: " & exported & " sheets" & vbCrLf & _
               "Skipped:  " & skipped & " sheets (not found or error)" & vbCrLf & vbCrLf & _
               "Output folder:" & vbCrLf & exportPath & vbCrLf & vbCrLf & _
               "Copy this folder to:" & vbCrLf & "  tests\fixtures\sample_data\", _
               vbInformation, "Golden Dataset Export"
    End If

End Sub

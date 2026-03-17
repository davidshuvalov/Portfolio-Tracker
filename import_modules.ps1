# import_modules.ps1
# Imports changed .bas modules into the Portfolio Tracker .xlsb workbook.
# Run from the repo root on Windows:
#   powershell -ExecutionPolicy Bypass -File import_modules.ps1

param(
    [string]$WorkbookPath = "$PSScriptRoot\Portfolio Tracker - A Tool for MultiWalk v1.24.xlsb"
)

# Modules to import (in dependency order)
$modules = @(
    "E_ColumnConstants",
    "I_MISC",
    "W_Markets",
    "O_Strategies_Tab",
    "F_Summary_Tab_Setup",
    "X_PortfolioHistory"
)

Write-Host "Portfolio Tracker — Module Importer" -ForegroundColor Cyan
Write-Host "Workbook: $WorkbookPath"
Write-Host ""

if (-not (Test-Path $WorkbookPath)) {
    Write-Host "ERROR: Workbook not found at: $WorkbookPath" -ForegroundColor Red
    exit 1
}

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening workbook..." -NoNewline
    $wb = $excel.Workbooks.Open($WorkbookPath)
    Write-Host " OK" -ForegroundColor Green

    $vbProject = $wb.VBProject

    foreach ($moduleName in $modules) {
        $basFile = Join-Path $PSScriptRoot "$moduleName.bas"

        if (-not (Test-Path $basFile)) {
            Write-Host "  SKIP  $moduleName.bas (file not found)" -ForegroundColor Yellow
            continue
        }

        # Remove existing module if present
        $existing = $null
        try {
            foreach ($comp in $vbProject.VBComponents) {
                if ($comp.Name -eq $moduleName) {
                    $existing = $comp
                    break
                }
            }
        } catch {}

        if ($null -ne $existing) {
            $vbProject.VBComponents.Remove($existing)
            Write-Host "  REMOVE  $moduleName" -ForegroundColor DarkGray
        }

        # Import the .bas file
        $vbProject.VBComponents.Import($basFile) | Out-Null
        Write-Host "  IMPORT  $moduleName" -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "Saving workbook..." -NoNewline
    $wb.Save()
    Write-Host " OK" -ForegroundColor Green

    $wb.Close($false)
    Write-Host ""
    Write-Host "Done. Open the workbook in Excel to test." -ForegroundColor Cyan

} catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "If you see 'Programmatic access to VBA is disabled':" -ForegroundColor Yellow
    Write-Host "  Excel Options → Trust Center → Trust Center Settings" -ForegroundColor Yellow
    Write-Host "  → Macro Settings → Enable 'Trust access to the VBA project object model'" -ForegroundColor Yellow
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

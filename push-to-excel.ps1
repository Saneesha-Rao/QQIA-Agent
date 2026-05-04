<#
  QQIA Daily Excel Sync
  
  Reads bot data from localhost:3978/api/steps/json and pushes any changes
  to the OneDrive Excel file via Excel Desktop COM (co-authoring safe).
  
  Usage: Right-click → Run with PowerShell, or:
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "push-to-excel.ps1"
  
  Prerequisites: Excel Desktop installed, bot server running on localhost:3978
#>

$ErrorActionPreference = "Stop"
$BotUrl = "http://localhost:3978/api/steps/json"
$ExcelPath = "C:\Users\salingal\OneDrive - Microsoft\Seller Incentives\QQIA\FY27_Mint_RolloverTimeline.xlsx"
$SheetName = "FY27_Rollover"

# Column index for Corp Status (1-based) — adjust if your sheet layout differs
$ColStepId = 1
$ColCorpStatus = 6
$ColCorpCompletedDate = 7

Write-Host "=== QQIA Daily Excel Sync ===" -ForegroundColor Cyan
Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host ""

# Step 1: Fetch bot data
Write-Host "[1/4] Fetching bot data from $BotUrl ..."
try {
    $response = Invoke-RestMethod -Uri $BotUrl -Method GET -TimeoutSec 10
    $steps = $response.steps
    Write-Host "  Got $($steps.Count) steps from bot" -ForegroundColor Green
} catch {
    Write-Host "  ERROR: Bot not reachable. Is the server running?" -ForegroundColor Red
    Write-Host "  Start it with: cd qqia-agent; npx tsc; node dist/index.js" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Step 2: Open Excel
Write-Host "[2/4] Opening Excel file ..."
$excel = $null
$wb = $null
$createdExcel = $false

try {
    # Try to attach to a running Excel instance
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    Write-Host "  Attached to running Excel" -ForegroundColor Green
    
    # Check if our file is already open
    $wb = $null
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -like "*FY27_Mint_RolloverTimeline*") {
            $wb = $w
            break
        }
    }
    
    if (-not $wb) {
        Write-Host "  Opening file ..."
        $wb = $excel.Workbooks.Open($ExcelPath, 0, $false)
    } else {
        Write-Host "  File already open: $($wb.Name)" -ForegroundColor Green
    }
} catch {
    # No Excel running — start one
    Write-Host "  Starting Excel Desktop ..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $createdExcel = $true
    $wb = $excel.Workbooks.Open($ExcelPath, 0, $false)
}

$ws = $wb.Sheets.Item($SheetName)
if (-not $ws) {
    Write-Host "  ERROR: Sheet '$SheetName' not found!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

$lastRow = $ws.UsedRange.Rows.Count
Write-Host "  Sheet '$SheetName' — $lastRow rows" -ForegroundColor Green

# Step 3: Build a lookup of Excel rows by Step ID
Write-Host "[3/4] Comparing bot data with Excel ..."
$excelData = @{}
for ($i = 4; $i -le $lastRow; $i++) {
    $id = $ws.Cells.Item($i, $ColStepId).Text.Trim()
    if ($id) {
        $excelData[$id] = @{
            Row = $i
            CorpStatus = $ws.Cells.Item($i, $ColCorpStatus).Text.Trim()
            CorpCompletedDate = $ws.Cells.Item($i, $ColCorpCompletedDate).Text.Trim()
        }
    }
}

# Step 4: Push changes
Write-Host "[4/4] Pushing changes to Excel ..."
$changeCount = 0
$skipCount = 0

foreach ($step in $steps) {
    $id = $step.id
    if (-not $excelData.ContainsKey($id)) { continue }
    
    $row = $excelData[$id].Row
    $excelStatus = $excelData[$id].CorpStatus
    $botStatus = $step.corpStatus
    
    if ($botStatus -and $botStatus -ne $excelStatus) {
        $ws.Cells.Item($row, $ColCorpStatus).Value2 = $botStatus
        Write-Host "  $id : '$excelStatus' -> '$botStatus'" -ForegroundColor Yellow
        
        # If completed, also set the completed date
        if ($botStatus -eq "Completed" -and $step.corpCompletedDate) {
            $ws.Cells.Item($row, $ColCorpCompletedDate).Value2 = $step.corpCompletedDate
        }
        
        $changeCount++
    } else {
        $skipCount++
    }
}

if ($changeCount -gt 0) {
    Write-Host ""
    Write-Host "Saving $changeCount change(s) ..." -ForegroundColor Cyan
    $wb.Save()
    Write-Host "DONE — $changeCount cell(s) updated, $skipCount unchanged" -ForegroundColor Green
    Write-Host "Excel will auto-sync to SharePoint via co-authoring." -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "No changes needed — Excel is already up to date." -ForegroundColor Green
}

Write-Host ""
Write-Host "Sync complete at $(Get-Date -Format 'HH:mm:ss')"
Read-Host "Press Enter to exit"

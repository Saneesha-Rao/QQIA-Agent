<# 
  QQIA Excel Sync Script
  Called by the bot to update the SharePoint Excel file via the running Excel COM instance.
  
  Usage:
    powershell.exe -NoProfile -File excel-sync.ps1 -Action read
    powershell.exe -NoProfile -File excel-sync.ps1 -Action write -StepId "1.C" -CorpStatus "Completed"
#>

param(
  [Parameter(Mandatory=$true)]
  [ValidateSet("read","write")]
  [string]$Action,
  
  [string]$StepId,
  [string]$CorpStatus,
  [string]$FedStatus,
  [string]$CorpCompletedDate,
  [string]$ReferenceNotes
)

$ErrorActionPreference = "Stop"

function Get-ExcelWorkbook {
  try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
  } catch {
    # Excel not running - start it and open the file
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $path = "C:\Users\salingal\OneDrive - Microsoft\Seller Incentives\QQIA\FY27_Mint_RolloverTimeline.xlsx"
    $wb = $excel.Workbooks.Open($path, 0, $false)
    return $wb
  }
  
  foreach ($wb in $excel.Workbooks) {
    if ($wb.Name -like "*FY27_Mint*") {
      return $wb
    }
  }
  
  # File not open - open it
  $path = "C:\Users\salingal\OneDrive - Microsoft\Seller Incentives\QQIA\FY27_Mint_RolloverTimeline.xlsx"
  return $excel.Workbooks.Open($path, 0, $false)
}

try {
  $wb = Get-ExcelWorkbook
  $ws = $wb.Sheets.Item("FY27_Rollover")
  $lastRow = $ws.UsedRange.Rows.Count

  if ($Action -eq "read") {
    # Read all step rows as JSON
    $rows = @()
    for ($i = 4; $i -le $lastRow; $i++) {
      $id = $ws.Cells.Item($i, 1).Text
      if ([string]::IsNullOrWhiteSpace($id)) { continue }
      
      $row = @{
        id = $id
        workstream = $ws.Cells.Item($i, 2).Text
        description = $ws.Cells.Item($i, 3).Text
        corpStartDate = $ws.Cells.Item($i, 4).Text
        corpEndDate = $ws.Cells.Item($i, 5).Text
        corpStatus = $ws.Cells.Item($i, 6).Text
        corpCompletedDate = $ws.Cells.Item($i, 7).Text
        fedStartDate = $ws.Cells.Item($i, 8).Text
        fedEndDate = $ws.Cells.Item($i, 9).Text
        fedStatus = $ws.Cells.Item($i, 10).Text
        setupValidation = $ws.Cells.Item($i, 11).Text
        engineeringDependent = $ws.Cells.Item($i, 12).Text
        wwicPoc = $ws.Cells.Item($i, 13).Text
        fedPoc = $ws.Cells.Item($i, 14).Text
        engineeringDri = $ws.Cells.Item($i, 15).Text
        engineeringLead = $ws.Cells.Item($i, 16).Text
        dependencies = $ws.Cells.Item($i, 17).Text
        referenceNotes = $ws.Cells.Item($i, 19).Text
        row = $i
      }
      $rows += $row
    }
    
    $result = @{ action = "read"; count = $rows.Count; rows = $rows }
    $result | ConvertTo-Json -Depth 3 -Compress
  }
  elseif ($Action -eq "write") {
    if (-not $StepId) {
      Write-Output '{"error":"stepId required for write"}'
      exit 1
    }
    
    $updated = $false
    for ($i = 4; $i -le $lastRow; $i++) {
      if ($ws.Cells.Item($i, 1).Text -eq $StepId) {
        if ($CorpStatus) { $ws.Cells.Item($i, 6).Value2 = $CorpStatus }
        if ($FedStatus) { $ws.Cells.Item($i, 10).Value2 = $FedStatus }
        if ($CorpCompletedDate) { $ws.Cells.Item($i, 7).Value2 = $CorpCompletedDate }
        if ($ReferenceNotes) { $ws.Cells.Item($i, 19).Value2 = $ReferenceNotes }
        
        $wb.Save()
        
        $result = @{
          action = "write"
          updated = 1
          stepId = $StepId
          row = $i
          corpStatus = $ws.Cells.Item($i, 6).Text
        }
        $result | ConvertTo-Json -Compress
        $updated = $true
        break
      }
    }
    
    if (-not $updated) {
      Write-Output "{`"error`":`"Step $StepId not found`",`"updated`":0}"
    }
  }
} catch {
  $err = @{ error = $_.Exception.Message; action = $Action }
  $err | ConvertTo-Json -Compress
  exit 1
}

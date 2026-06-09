<#
  firewall-tracker.ps1
  Starts a local HTTP server with a web form for tracking firewall requests.
  Data is saved to fw-requests.xlsx on your desktop.
  
  Usage: run in PowerShell (no admin needed)
    cd to this folder, then: .\firewall-tracker.ps1
  Then open: http://localhost:18080
#>

param([int]$Port = 18080)

# ── Excel file location ──
$desktop = [Environment]::GetFolderPath("Desktop")
$excelFile = Join-Path $desktop "firewall-requests.xlsx"

# ── Ensure Excel file exists with headers ──
function Ensure-ExcelFile {
    if (-not (Test-Path $excelFile)) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Sheets.Item(1)
        $ws.Name = "Firewall Requests"
        $headers = @(
            "Application", "Requester", "Priority", "Status",
            "Source IP", "Dest IP", "Dest Port",
            "Protocol", "Direction", "Justification",
            "Date Submitted", "Date Closed", "Notes", "Ticket Ref"
        );
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $ws.Cells.Item(1, $i + 1) = $headers[$i]
        }
        $ws.Range("A1:O1").Font.Bold = $true
        $ws.Range("A1:O1").Interior.ColorIndex = 44
        $ws.Range("A1:O1").Font.ColorIndex = 2
        $ws.Range("A1:O1").HorizontalAlignment = -4108 # xlCenter
        $ws.Columns.AutoFit() | Out-Null
        $wb.SaveAs($excelFile, 51) # 51 = xlsx format
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Write-Host "Created: $excelFile" -ForegroundColor Green
    }
}

# ── Read all rows from Excel ──
function Get-ExcelData {
    $excel = New-Object -ComObject Excel.Application
    $wb = $excel.Workbooks.Open($excelFile)
    $ws = $wb.Sheets.Item(1)
    $usedRows = $ws.UsedRange.Rows.Count
    $columns = @("Application", "Requester", "Priority", "Status", "SourceIP", "DestIP", "DestPort", "Protocol", "Direction", "Justification", "DateSubmitted", "DateClosed", "Notes", "TicketRef", "")
    $data = @()
    if ($usedRows -gt 1) {
        for ($r = 2; $r -le $usedRows; $r++) {
            $row = @{}
            for ($c = 1; $c -le 15; $c++) {
                $val = $ws.Cells.Item($r, $c).Value2
                $key = $columns[$c - 1]
                $row[$key] = if ($val -eq $null) { "" } else { $val.ToString() }
            }
            $data += $row
        }
    }
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    return $data
}

# ── Append row to Excel ──
function Add-ExcelRow($obj) {
    $excel = New-Object -ComObject Excel.Application
    $wb = $excel.Workbooks.Open($excelFile)
    $ws = $wb.Sheets.Item(1)
    $lastRow = $ws.UsedRange.Rows.Count + 1
    $ws.Cells.Item($lastRow, 1) = $obj.Application
    $ws.Cells.Item($lastRow, 2) = $obj.Requester
    $ws.Cells.Item($lastRow, 3) = $obj.Priority
    $ws.Cells.Item($lastRow, 4) = $obj.Status
    $ws.Cells.Item($lastRow, 5) = $obj.SourceIP
    $ws.Cells.Item($lastRow, 6) = $obj.DestIP
    $ws.Cells.Item($lastRow, 7) = $obj.DestPort
    $ws.Cells.Item($lastRow, 8) = $obj.Protocol
    $ws.Cells.Item($lastRow, 9) = $obj.Direction
    $ws.Cells.Item($lastRow, 10) = $obj.Justification
    $ws.Cells.Item($lastRow, 11) = Get-Date -Format "yyyy-MM-dd HH:mm"
    $ws.Cells.Item($lastRow, 12) = $obj.DateClosed
    $ws.Cells.Item($lastRow, 13) = $obj.Notes
    $ws.Cells.Item($lastRow, 14) = $obj.TicketRef
    $wb.Save()
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# ── HTTP Server ──
Ensure-ExcelFile

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Firewall Request Tracker" -ForegroundColor Cyan
Write-Host "  Open: http://localhost:$Port" -ForegroundColor Yellow
Write-Host "  Data: $excelFile" -ForegroundColor Gray
Write-Host "  Press Ctrl+C to stop" -ForegroundColor Gray
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
                $url = if ($request.Url.AbsolutePath -eq "/") { "/" } else { $request.Url.AbsolutePath.TrimEnd("/") }
        $method = $request.HttpMethod
        $response = $context.Response

        # ── Serve HTML form ──
        if ($url -eq "/" -or $url -eq "/index.html" -or $url -eq "/form.html") {
            # Read the form file from the same directory
            $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
            $htmlPath = Join-Path $scriptDir "form.html"
            if (Test-Path $htmlPath) {
                $body = [System.IO.File]::ReadAllText($htmlPath)
            } else {
                $body = "<h1>form.html not found</h1><p>Expected at: " + $scriptDir + "</p><p>Place form.html in the same directory as firewall-tracker.ps1</p>"
            }
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($body)
            $response.ContentType = "text/html; charset=utf-8"
            $response.ContentLength64 = $buffer.Length
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.Close()
        }
        # ── API: GET all records ──
        elseif ($url -eq "/api/records" -and $method -eq "GET") {
            $data = Get-ExcelData
            $json = $data | ConvertTo-Json -Depth 5
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
            $response.ContentType = "application/json"
            $response.ContentLength64 = $buffer.Length
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.Close()
        }
        # ── API: POST new record ──
        elseif ($url -eq "/api/records" -and $method -eq "POST") {
            $reader = New-Object System.IO.StreamReader($request.InputStream)
            $raw = $reader.ReadToEnd()
            $obj = $raw | ConvertFrom-Json
            Add-ExcelRow $obj
            $respBody = '{"ok":true}'
            $buf = [System.Text.Encoding]::UTF8.GetBytes($respBody)
            $response.ContentType = "application/json"
            $response.ContentLength64 = $buf.Length
            $response.OutputStream.Write($buf, 0, $buf.Length)
            $response.Close()
        }
        # ── API: DELETE record ──
        elseif ($url -match "^/api/records/([^/]+)$" -and $method -eq "DELETE") {
            $key = $Matches[1]
            $excel = New-Object -ComObject Excel.Application
            $wb = $excel.Workbooks.Open($excelFile)
            $ws = $wb.Sheets.Item(1)
            $usedRows = $ws.UsedRange.Rows.Count
            $found = $false
            for ($r = 2; $r -le $usedRows; $r++) {
                if ($ws.Cells.Item($r, 1).Value2 -eq $key) {
                    $ws.Rows.Item($r).Delete()
                    $found = $true
                    break
                }
            }
            if ($found) {
                $wb.Save()
            }
            $wb.Close($false)
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            $status = if ($found) { "deleted" } else { "not_found" }
            $resp = '{"ok":true,"status":"' + $status + '"}'
            $buf = [System.Text.Encoding]::UTF8.GetBytes($resp)
            $response.ContentType = "application/json"
            $response.ContentLength64 = $buf.Length
            $response.OutputStream.Write($buf, 0, $buf.Length)
            $response.Close()
        }
        # ── API: UPDATE record ──
        elseif ($url -match "^/api/records/([^/]+)$" -and $method -eq "PUT") {
            $key = $Matches[1]
            $reader = New-Object System.IO.StreamReader($request.InputStream)
            $raw = $reader.ReadToEnd()
            $obj = $raw | ConvertFrom-Json
            $excel = New-Object -ComObject Excel.Application
            $wb = $excel.Workbooks.Open($excelFile)
            $ws = $wb.Sheets.Item(1)
            $usedRows = $ws.UsedRange.Rows.Count
            $found = $false
            for ($r = 2; $r -le $usedRows; $r++) {
                if ($ws.Cells.Item($r, 1).Value2 -eq $key) {
                    if ($obj.PSObject.Properties["Status"])        { $ws.Cells.Item($r, 4) = $obj.Status }
                    if ($obj.PSObject.Properties["SourceIP"])       { $ws.Cells.Item($r, 5) = $obj.SourceIP }
                    if ($obj.PSObject.Properties["DestIP"])         { $ws.Cells.Item($r, 6) = $obj.DestIP }
                    if ($obj.PSObject.Properties["DestPort"])       { $ws.Cells.Item($r, 7) = $obj.DestPort }
                    if ($obj.PSObject.Properties["Protocol"])       { $ws.Cells.Item($r, 8) = $obj.Protocol }
                    if ($obj.PSObject.Properties["Direction"])      { $ws.Cells.Item($r, 9) = $obj.Direction }
                    if ($obj.PSObject.Properties["Priority"])       { $ws.Cells.Item($r, 3) = $obj.Priority }
                    if ($obj.PSObject.Properties["Justification"])  { $ws.Cells.Item($r, 10) = $obj.Justification }
                    if ($obj.PSObject.Properties["DateClosed"])     { $ws.Cells.Item($r, 12) = $obj.DateClosed }
                    if ($obj.PSObject.Properties["Notes"])          { $ws.Cells.Item($r, 13) = $obj.Notes }
                    $found = $true
                    break
                }
            }
            if ($found) {
                $wb.Save()
            }
            $wb.Close($false)
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            $status = if ($found) { "updated" } else { "not_found" }
            $resp = '{"ok":true,"status":"' + $status + '"}'
            $buf = [System.Text.Encoding]::UTF8.GetBytes($resp)
            $response.ContentType = "application/json"
            $response.ContentLength64 = $buf.Length
            $response.OutputStream.Write($buf, 0, $buf.Length)
            $response.Close()
        }
        # ── 404 ──
        else {
            $body404 = "<h1>404</h1>"
            $buf404 = [System.Text.Encoding]::UTF8.GetBytes($body404)
            $response.ContentType = "text/html"
            $response.ContentLength64 = $buf404.Length
            $response.OutputStream.Write($buf404, 0, $buf404.Length)
            $response.Close()
        }
    }
}
finally {
    $listener.Stop()
    $listener.Close()
}

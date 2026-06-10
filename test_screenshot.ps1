# test_screenshot.ps1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$excelPath = "C:\Users\opera\.gemini\antigravity\scratch\Escala 10.06.26.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($excelPath)

# Find scale sheet
$escalaSheet = $null
foreach ($sheet in $workbook.Sheets) {
    if ($sheet.Name.Trim().ToLower() -like "escala di*ria seg a sex") {
        $escalaSheet = $sheet
        break
    }
}

if (-not $escalaSheet) {
    Write-Host "Escala sheet not found!"
    $workbook.Close($false)
    $excel.Quit()
    exit
}

$lastRow = $escalaSheet.UsedRange.Rows.Count
Write-Host "Last row: ${lastRow}"

# Filter by "ADEMAR DOS SANTOS"
# Row 2 is headers, filter field 6 (MOTORISTA)
$filterRange = $escalaSheet.Range("A2:H${lastRow}")
$filterRange.AutoFilter(6, "ADEMAR DOS SANTOS")

# Define the full range to copy (including Row 1 title)
$copyRange = $escalaSheet.Range("A1:H${lastRow}")

# Copy as picture
# xlScreen = 1, xlBitmap = 2
$copyRange.CopyPicture(1, 2) | Out-Null

# Wait a second for clipboard
Start-Sleep -Milliseconds 500

# Save image from clipboard
if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
    $img = [System.Windows.Forms.Clipboard]::GetImage()
    $savePath = "C:\Users\opera\.gemini\antigravity\scratch\test_ademar.png"
    $img.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Host "Image saved to ${savePath}"
    $img.Dispose()
} else {
    Write-Host "Clipboard does not contain image!"
}

# Clear filter
if ($escalaSheet.AutoFilterMode) {
    $escalaSheet.AutoFilter.ShowAllData()
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

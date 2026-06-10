# generate_prints.ps1
param (
    [string]$ExcelPath = "C:\Users\opera\.gemini\antigravity\scratch\Escala 10.06.26.xlsx",
    [string]$OutputDir = "C:\Users\opera\.gemini\antigravity\scratch\temp_prints"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create output dir if it doesn't exist, or clear it if it does to avoid stale images
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
} else {
    Remove-Item -Path (Join-Path $OutputDir "*") -Force -ErrorAction SilentlyContinue
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.WindowState = -4140 # xlMinimized

$workbook = $null

try {
    $workbook = $excel.Workbooks.Open($ExcelPath)
    
    # Locate sheets
    $dadosSheet = $null
    $escalaSheet = $null
    foreach ($sheet in $workbook.Sheets) {
        $nameTrim = $sheet.Name.Trim().ToLower()
        if ($nameTrim -eq "dados") {
            $dadosSheet = $sheet
        }
        if ($nameTrim -like "escala di*ria seg a sex") {
            $escalaSheet = $sheet
        }
    }
    
    if (-not $dadosSheet -or -not $escalaSheet) {
        Write-Error "Required sheets ('Dados' or 'Escala') not found in workbook."
        exit 1
    }
    
    # Get last row of "Dados" sheet
    $xlUp = -4162
    $dadosLastRow = $dadosSheet.Cells.Item($dadosSheet.Rows.Count, 1).End($xlUp).Row
    
    # Get last row of "Escala" sheet
    $escalaLastRow = $escalaSheet.Cells.Item($escalaSheet.Rows.Count, 6).End($xlUp).Row
    
    $driversList = @()
    
    # Read drivers (ignore any driver whose name cell has a red background)
    for ($row = 2; $row -le $dadosLastRow; $row++) {
        $cell = $dadosSheet.Cells.Item($row, 1)
        $name = $cell.Text.Trim()
        $phone = $dadosSheet.Cells.Item($row, 2).Text.Trim()
        
        # Check if cell background is red (Color 255 or ColorIndex 3)
        $isRed = ($cell.Interior.Color -eq 255) -or ($cell.Interior.ColorIndex -eq 3)
        
        if ($name -and -not $isRed) {
            $driversList += [PSCustomObject]@{
                Name = $name
                Phone = $phone
                HasScale = $false
                ImagePath = $null
            }
        }
    }
    
    # Get all motoristas in the scale sheet in one single COM call to optimize performance!
    $escalaMotoristas = @()
    if ($escalaLastRow -ge 3) {
        $rangeValues = $escalaSheet.Range("F3:F${escalaLastRow}").Value2
        if ($rangeValues -is [array]) {
            # Convert 2D COM array to flat 1D array of lowercase strings
            foreach ($val in $rangeValues) {
                if ($val) {
                    $escalaMotoristas += $val.ToString().Trim().ToLower()
                }
            }
        } elseif ($rangeValues) {
            $escalaMotoristas += $rangeValues.ToString().Trim().ToLower()
        }
    }

    # Process each driver
    foreach ($driver in $driversList) {
        try {
            # Check if driver has scale (thousands of times faster in-memory check!)
            $hasAny = $escalaMotoristas -contains $driver.Name.ToLower()
            
            if ($hasAny) {
                $driver.HasScale = $true
                
                # Apply filter (columns A-H)
                $filterRange = $escalaSheet.Range("A2:H${escalaLastRow}")
                $filterRange.AutoFilter(6, $driver.Name) | Out-Null
                
                # Copy range A1:F as picture (excludes columns G and H with car numbers)
                $copyRange = $escalaSheet.Range("A1:F${escalaLastRow}")
                
                # Activate sheet and select range to stabilize CopyPicture
                $escalaSheet.Activate()
                $copyRange.Select() | Out-Null
                
                # Clear and verify clipboard is empty to prevent stale images
                $cleared = $false
                for ($c = 1; $c -le 3; $c++) {
                    [System.Windows.Forms.Clipboard]::Clear()
                    Start-Sleep -Milliseconds 50
                    if (-not [System.Windows.Forms.Clipboard]::ContainsImage()) {
                        $cleared = $true
                        break
                    }
                }
                
                $copied = $false
                # Retry loop for copying and retrieving the range image from clipboard
                for ($attempt = 1; $attempt -le 5; $attempt++) {
                    # Clear clipboard each attempt
                    [System.Windows.Forms.Clipboard]::Clear()
                    Start-Sleep -Milliseconds 50
                    
                    $copyRange.CopyPicture(1, 2) | Out-Null
                    Start-Sleep -Milliseconds 150
                    
                    if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
                        $img = [System.Windows.Forms.Clipboard]::GetImage()
                        if ($img) {
                            $sanitizedName = $driver.Name -replace '[^a-zA-Z0-9_]', '_'
                            $savePath = Join-Path $OutputDir "${sanitizedName}.png"
                            $webPath = $savePath -replace '\\', '/'
                            $img.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)
                            $img.Dispose()
                            
                            $driver.ImagePath = $webPath
                            $copied = $true
                            break
                        }
                    }
                    Start-Sleep -Milliseconds 200
                }
                
                if (-not $copied) {
                    Write-Warning "Could not copy scale screenshot for $($driver.Name) after 5 attempts."
                }
                
                # Clear filter
                if ($escalaSheet.AutoFilterMode) {
                    $escalaSheet.AutoFilter.ShowAllData() | Out-Null
                }
            }
        } catch {
            # Log warning but do not crash the script, allowing other drivers to be processed successfully!
            Write-Warning "Erro ao processar motorista $($driver.Name): $($_.Exception.Message)"
            if ($escalaSheet -and $escalaSheet.AutoFilterMode) {
                try { $escalaSheet.AutoFilter.ShowAllData() } catch {}
            }
        }
    }
    
    # Output JSON to stdout
    $jsonOutput = $driversList | ConvertTo-Json -Compress
    Write-Output "JSON_START"
    Write-Output $jsonOutput
    Write-Output "JSON_END"
    
} catch {
    Write-Error $_.Exception.Message
    exit 1
} finally {
    if ($workbook) {
        $workbook.Close($false)
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    # Ensure processes are freed
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

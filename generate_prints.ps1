# generate_prints.ps1
param (
    [string]$ExcelPath = "C:\Users\opera\.gemini\antigravity\scratch\Escala 10.06.26.xlsx",
    [string]$OutputDir = "C:\Users\opera\.gemini\antigravity\scratch\temp_prints"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create output dir if it doesn't exist
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

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
    
    # Process each driver
    foreach ($driver in $driversList) {
        try {
            # Check if driver has scale in columns 6 (MOTORISTA)
            $hasAny = $false
            for ($r = 3; $r -le $escalaLastRow; $r++) {
                $mName = $escalaSheet.Cells.Item($r, 6).Text.Trim()
                if ($mName -eq $driver.Name) {
                    $hasAny = $true
                    break
                }
            }
            
            if ($hasAny) {
                $driver.HasScale = $true
                
                # Apply filter (columns A-H)
                $filterRange = $escalaSheet.Range("A2:H${escalaLastRow}")
                $filterRange.AutoFilter(6, $driver.Name)
                
                # Copy range A1:F as picture (excludes columns G and H with car numbers)
                $copyRange = $escalaSheet.Range("A1:F${escalaLastRow}")
                
                # Activate sheet and select range to stabilize CopyPicture
                $escalaSheet.Activate()
                $copyRange.Select() | Out-Null
                
                # Clean clipboard first
                [System.Windows.Forms.Clipboard]::Clear()
                
                $copyRange.CopyPicture(1, 2) | Out-Null
                Start-Sleep -Milliseconds 400
                
                # Save from clipboard
                if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
                    $img = [System.Windows.Forms.Clipboard]::GetImage()
                    $sanitizedName = $driver.Name -replace '[^a-zA-Z0-9_]', '_'
                    $savePath = Join-Path $OutputDir "${sanitizedName}.png"
                    # Convert path to forward slashes for Javascript friendliness
                    $webPath = $savePath -replace '\\', '/'
                    $img.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)
                    $img.Dispose()
                    
                    $driver.ImagePath = $webPath
                } else {
                    # Retry once
                    Start-Sleep -Milliseconds 400
                    $copyRange.CopyPicture(1, 2) | Out-Null
                    Start-Sleep -Milliseconds 400
                    if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
                        $img = [System.Windows.Forms.Clipboard]::GetImage()
                        $sanitizedName = $driver.Name -replace '[^a-zA-Z0-9_]', '_'
                        $savePath = Join-Path $OutputDir "${sanitizedName}.png"
                        $webPath = $savePath -replace '\\', '/'
                        $img.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)
                        $img.Dispose()
                        $driver.ImagePath = $webPath
                    }
                }
                
                # Clear filter
                if ($escalaSheet.AutoFilterMode) {
                    $escalaSheet.AutoFilter.ShowAllData()
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

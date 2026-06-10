# inspect_excel_v2.ps1
$excelPath = "C:\Users\opera\.gemini\antigravity\scratch\Escala 10.06.26.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)

Write-Host "Sheets in this workbook:"
$index = 1
$dadosSheet = $null
$escalaSheet = $null

foreach ($sheet in $workbook.Sheets) {
    $name = $sheet.Name
    $len = $name.Length
    Write-Host "Index ${index} - Name='${name}' Length=${len}"
    if ($name.Trim().ToLower() -eq "dados") {
        $dadosSheet = $sheet
    }
    if ($name.Trim().ToLower() -like "escala di*ria seg a sex") {
        $escalaSheet = $sheet
    }
    $index++
}

# Inspect "Dados" sheet
if ($dadosSheet) {
    Write-Host "`nFirst 20 rows of 'Dados' sheet:"
    for ($row = 1; $row -le 20; $row++) {
        $val1 = $dadosSheet.Cells.Item($row, 1).Text
        $val2 = $dadosSheet.Cells.Item($row, 2).Text
        $val3 = $dadosSheet.Cells.Item($row, 3).Text
        $val4 = $dadosSheet.Cells.Item($row, 4).Text
        if ($val1 -or $val2 -or $val3 -or $val4) {
            Write-Host "Row ${row} - | ${val1} | ${val2} | ${val3} | ${val4} |"
        }
    }
} else {
    Write-Host "`n'Dados' sheet not found."
}

# Inspect active sheet or "Escala Diária seg a sex"
if ($escalaSheet) {
    Write-Host "`nFirst 20 rows of '$($escalaSheet.Name)' sheet:"
    for ($row = 1; $row -le 20; $row++) {
        $val1 = $escalaSheet.Cells.Item($row, 1).Text
        $val2 = $escalaSheet.Cells.Item($row, 2).Text
        $val3 = $escalaSheet.Cells.Item($row, 3).Text
        $val4 = $escalaSheet.Cells.Item($row, 4).Text
        $val5 = $escalaSheet.Cells.Item($row, 5).Text
        $val6 = $escalaSheet.Cells.Item($row, 6).Text
        $val7 = $escalaSheet.Cells.Item($row, 7).Text
        $val8 = $escalaSheet.Cells.Item($row, 8).Text
        if ($val1 -or $val2 -or $val3 -or $val4 -or $val5 -or $val6 -or $val7 -or $val8) {
            Write-Host "Row ${row} - | ${val1} | ${val2} | ${val3} | ${val4} | ${val5} | ${val6} | ${val7} | ${val8} |"
        }
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

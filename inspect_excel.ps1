# inspect_excel.ps1
$excelPath = "C:\Users\opera\.gemini\antigravity\scratch\Escala 10.06.26.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)

Write-Host "Sheets in this workbook:"
foreach ($sheet in $workbook.Sheets) {
    Write-Host "- $($sheet.Name)"
}

# Inspect "Dados" sheet
$dadosSheet = $workbook.Sheets.Item("Dados")
if ($dadosSheet) {
    Write-Host "`nFirst 15 rows of 'Dados' sheet:"
    for ($row = 1; $row -le 15; $row++) {
        $val1 = $dadosSheet.Cells.Item($row, 1).Text
        $val2 = $dadosSheet.Cells.Item($row, 2).Text
        $val3 = $dadosSheet.Cells.Item($row, 3).Text
        if ($val1 -or $val2 -or $val3) {
            Write-Host "Row $row - | $val1 | $val2 | $val3 |"
        }
    }
} else {
    Write-Host "`n'Dados' sheet not found."
}

# Inspect active sheet or "Escala Diária seg a sex"
$escalaSheet = $workbook.Sheets.Item("Escala Diária seg a sex")
if ($escalaSheet) {
    Write-Host "`nFirst 15 rows of 'Escala Diária seg a sex' sheet:"
    for ($row = 1; $row -le 15; $row++) {
        $val1 = $escalaSheet.Cells.Item($row, 1).Text
        $val2 = $escalaSheet.Cells.Item($row, 2).Text
        $val3 = $escalaSheet.Cells.Item($row, 3).Text
        $val4 = $escalaSheet.Cells.Item($row, 4).Text
        $val5 = $escalaSheet.Cells.Item($row, 5).Text
        $val6 = $escalaSheet.Cells.Item($row, 6).Text
        if ($val1 -or $val2 -or $val3 -or $val4 -or $val5 -or $val6) {
            Write-Host "Row $row - | $val1 | $val2 | $val3 | $val4 | $val5 | $val6 |"
        }
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

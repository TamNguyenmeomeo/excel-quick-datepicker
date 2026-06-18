$regSecurity = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$oldVbom = (Get-ItemProperty $regSecurity -Name "AccessVBOM" -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty $regSecurity -Name "AccessVBOM" -Value 1 -Force | Out-Null

try {
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open("C:\Users\lenovo\Downloads\CanlendarExcel\CalendarForm v1.5.2.xlsm")
    
    # Inspect ThisWorkbook
    $comp = $workbook.VBProject.VBComponents.Item("ThisWorkbook")
    $lines = $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
    Write-Output "--- ThisWorkbook ---"
    Write-Output $lines
    
    # Inspect Sheet1
    $comp = $workbook.VBProject.VBComponents.Item("Sheet1")
    $lines = $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
    Write-Output "--- Sheet1 ---"
    Write-Output $lines
    
    # Inspect Sheet3
    $comp = $workbook.VBProject.VBComponents.Item("Sheet3")
    $lines = $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
    Write-Output "--- Sheet3 ---"
    Write-Output $lines

    # Inspect Sheet4
    $comp = $workbook.VBProject.VBComponents.Item("Sheet4")
    $lines = $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
    Write-Output "--- Sheet4 ---"
    Write-Output $lines

    $excel.Quit()
} catch {
    Write-Output $_.Exception.Message
} finally {
    if ($oldVbom -ne $null) {
        Set-ItemProperty $regSecurity -Name "AccessVBOM" -Value $oldVbom -Force | Out-Null
    } else {
        Remove-ItemProperty $regSecurity -Name "AccessVBOM" -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

$outputFile = "merged.xlsx"

$exportFolder = "export"

$exportExcelPath = Join-Path $PSScriptRoot $exportFolder

if (Test-Path -Path $exportExcelPath) {
    Write-Host "export folder found! Processing..."
} else {
    New-Item $exportFolder  -ItemType Directory
    Write-Host "created export folder. Processing..."
}

$exportFullPath = Join-Path $exportExcelPath $outputFile

$excelFiles = Get-ChildItem -Path $PSScriptRoot -Filter "*.xlsx"
foreach ($excelFile in $excelFiles)
{
    $theWorksheetName = [io.path]::GetFileNameWithoutExtension($excelFile) 
    $tabName = $theWorksheetName.Split(" ")
    Import-Excel -Path $excelFile.FullName | Export-Excel -Path $exportFullPath -WorkSheetname $tabName[0]
}
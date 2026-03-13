param(
    [string]$ExcelPath = '.\biljart moyenne.xlsx',
    [string]$OutputPath = '.\site-data.js'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$culture = [System.Globalization.CultureInfo]::GetCultureInfo('nl-NL')

function Format-DutchDate {
    param([datetime]$Date)

    return $Date.ToString('d MMMM yyyy', $culture)
}

function Get-WeekDataFromWorksheet {
    param($Worksheet)

    $rows = $Worksheet.UsedRange.Rows.Count
    $cols = $Worksheet.UsedRange.Columns.Count

    if ($cols -lt 5) {
        throw 'Werkblad heeft te weinig kolommen. Verwacht: datum + 4 spelers.'
    }

    $weeks = [System.Collections.Generic.List[object]]::new()
    $currentWeek = $null

    for ($row = 2; $row -le $rows; $row++) {
        $dateText = [string]$Worksheet.Cells.Item($row, 1).Text
        if (-not [string]::IsNullOrWhiteSpace($dateText)) {
            $date = [datetime]::Parse($dateText, $culture)
            $currentWeek = [pscustomobject]@{
                datumSort = $date
                datum = Format-DutchDate -Date $date
                partijen = [System.Collections.Generic.List[object]]::new()
            }
            $weeks.Add($currentWeek)
        }

        if ($null -eq $currentWeek) {
            continue
        }

        $partij = [System.Collections.Generic.List[double]]::new()
        $hasAllValues = $true
        for ($col = 2; $col -le 5; $col++) {
            $value = $Worksheet.Cells.Item($row, $col).Value2
            if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) {
                $hasAllValues = $false
                break
            }

            $partij.Add([Math]::Round([double]$value, 9))
        }

        if ($hasAllValues) {
            $currentWeek.partijen.Add($partij.ToArray())
        }
    }

    return $weeks |
        Where-Object { $_.partijen.Count -gt 0 } |
        Sort-Object datumSort -Descending |
        ForEach-Object {
            [ordered]@{
                datum = $_.datum
                partijen = @($_.partijen)
            }
        }
}

$resolvedExcelPath = Resolve-Path -Path $ExcelPath
$resolvedOutputPath = Join-Path -Path (Split-Path -Parent $resolvedExcelPath) -ChildPath (Split-Path -Leaf $OutputPath)

$excel = $null
$workbook = $null
$worksheet = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($resolvedExcelPath, $null, $true)
    $worksheet = $workbook.Worksheets.Item(1)

    $weeks = Get-WeekDataFromWorksheet -Worksheet $worksheet

    $payload = [ordered]@{
        generatedAt = (Get-Date).ToString('dd-MM-yyyy HH:mm', $culture)
        sourceFile = [System.IO.Path]::GetFileName($resolvedExcelPath)
        weeks = $weeks
    }

    $json = $payload | ConvertTo-Json -Depth 6
    $jsContent = "window.BILJART_SITE_DATA = $json;"
    Set-Content -Path $resolvedOutputPath -Value $jsContent -Encoding UTF8

    Write-Host "site-data.js bijgewerkt op basis van $([System.IO.Path]::GetFileName($resolvedExcelPath))."
    Write-Host "Datums verwerkt: $($weeks.Count)"
}
finally {
    if ($workbook) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($worksheet) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
    }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
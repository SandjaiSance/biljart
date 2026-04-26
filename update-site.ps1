param(
    [string[]]$ExcelPaths = @('.\biljart moyenne.xlsx'),
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
        $dateValue = $Worksheet.Cells.Item($row, 1).Value2
        $dateText = [string]$Worksheet.Cells.Item($row, 1).Text
        $hasDate = $false
        $date = $null
        if ($null -ne $dateValue -and $dateValue -is [double] -and $dateValue -gt 1) {
            # Excel stores dates as OA date serial numbers
            $date = [DateTime]::FromOADate($dateValue)
            $hasDate = $true
        } elseif (-not [string]::IsNullOrWhiteSpace($dateText)) {
            try {
                $date = [datetime]::Parse($dateText, $culture)
                $hasDate = $true
            } catch {
                # skip unparseable text
            }
        }
        if ($hasDate) {
            if ($date.Year -eq 2025 -and $date.Month -eq 3 -and $date.Day -eq 27) {
                # Known source typo in workbook: this match day belongs to season 2026.
                $date = $date.AddYears(1)
            }
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

        $partij = [System.Collections.Generic.List[object]]::new()
        $filledCount = 0
        for ($col = 2; $col -le 5; $col++) {
            $value = $Worksheet.Cells.Item($row, $col).Value2
            if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) {
                $partij.Add($null)
                continue
            }

            $partij.Add([Math]::Round([double]$value, 9))
            $filledCount++
        }

        # Include rows where at least 3 spelers have a value.
        if ($filledCount -ge 3) {
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

function Update-IndexCacheBuster {
    param(
        [string]$IndexPath,
        [string]$VersionToken
    )

    if (-not (Test-Path -Path $IndexPath)) {
        return
    }

    $indexContent = Get-Content -Path $IndexPath -Raw
    $updatedContent = [regex]::Replace(
        $indexContent,
        '<script\s+src="site-data\.js(?:\?v=[^"]*)?"\s*></script>',
        ('<script src="site-data.js?v=' + $VersionToken + '"></script>'),
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    if ($updatedContent -ne $indexContent) {
        Set-Content -Path $IndexPath -Value $updatedContent -Encoding UTF8
    }
}

$resolvedOutputPath = $null
$resolvedIndexPath = $null
$allWeeks = [System.Collections.Generic.List[object]]::new()
$sourceFiles = [System.Collections.Generic.List[string]]::new()

$excel = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($ExcelPath in $ExcelPaths) {
        if (-not (Test-Path -Path $ExcelPath)) {
            Write-Warning "Bestand niet gevonden, overgeslagen: $ExcelPath"
            continue
        }

        $resolvedExcelPath = Resolve-Path -Path $ExcelPath
        if ($null -eq $resolvedOutputPath) {
            $resolvedOutputPath = Join-Path -Path (Split-Path -Parent $resolvedExcelPath) -ChildPath (Split-Path -Leaf $OutputPath)
            $resolvedIndexPath  = Join-Path -Path (Split-Path -Parent $resolvedExcelPath) -ChildPath 'index.html'
        }

        $workbook  = $excel.Workbooks.Open($resolvedExcelPath, $null, $true)
        $worksheet = $workbook.Worksheets.Item(1)

        try {
            $weeks = Get-WeekDataFromWorksheet -Worksheet $worksheet
            foreach ($w in $weeks) { $allWeeks.Add($w) }
            $sourceFiles.Add([System.IO.Path]::GetFileName($resolvedExcelPath))
        } finally {
            $workbook.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
    }

    # Merge: deduplicate on datum, sort descending
    $mergedWeeks = $allWeeks |
        Group-Object { $_['datum'] } |
        ForEach-Object {
            # If same date in both files, take the one with most partijen
            $_.Group | Sort-Object { @($_['partijen']).Count } -Descending | Select-Object -First 1
        } |
        Sort-Object { [datetime]::ParseExact($_['datum'], 'd MMMM yyyy', $culture) } -Descending

    $payload = [ordered]@{
        generatedAt = (Get-Date).ToString('dd-MM-yyyy HH:mm', $culture)
        sourceFile  = ($sourceFiles -join ', ')
        weeks       = @($mergedWeeks)
    }

    $json      = $payload | ConvertTo-Json -Depth 6
    $jsContent = "window.BILJART_SITE_DATA = $json;"
    Set-Content -Path $resolvedOutputPath -Value $jsContent -Encoding UTF8

    $cacheBustVersion = (Get-Date).ToString('yyyyMMddHHmmss')
    Update-IndexCacheBuster -IndexPath $resolvedIndexPath -VersionToken $cacheBustVersion

    Write-Host "site-data.js bijgewerkt op basis van: $($sourceFiles -join ', ')."
    Write-Host "Datums verwerkt: $($mergedWeeks.Count)"
}
finally {
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$basePath = Split-Path $MyInvocation.MyCommand.Path
$srcPath  = Join-Path $basePath 'Driebanden moyenne.xlsx'
$dstPath  = Join-Path $basePath 'biljart moyenne.xlsx'

$excel = New-Object -ComObject Excel.Application
$excel.Visible         = $false
$excel.DisplayAlerts   = $false
$excel.AskToUpdateLinks = $false

try {
    # --- Lees bron: Driebanden moyenne.xlsx ---
    $wbSrc  = $excel.Workbooks.Open($srcPath, $null, $true)
    $wsSrc  = $wbSrc.Worksheets.Item(1)
    $srcRows = $wsSrc.UsedRange.Rows.Count

    # Verzamel alle blokken (datum + bijbehorende partij-rijen) die NIET al in doel staan
    # We weten dat ontbrekende datums zijn: 27/03/2026 en 03/04/2026

    # Bepaal bestaande datums in doel
    $wbDst  = $excel.Workbooks.Open($dstPath, $null, $false)
    if ($wbDst.ReadOnly) { throw 'biljart moyenne.xlsx is read-only!' }
    $wsDst  = $wbDst.Worksheets.Item(1)
    $dstRows = $wsDst.UsedRange.Rows.Count

    $existingDates = @{}
    for ($r = 2; $r -le $dstRows; $r++) {
        $v = $wsDst.Cells.Item($r,1).Value2
        if ($null -ne $v -and $v -is [double] -and $v -gt 1) {
            $d = [DateTime]::FromOADate($v).Date
            $existingDates[$d] = $true
        }
    }

    # Loop door bron en kopieer blokken waarvan de datum ontbreekt in doel
    $appendRow   = $dstRows + 1
    $copiedDates = [System.Collections.Generic.List[string]]::new()
    $currentDate = $null

    for ($r = 2; $r -le $srcRows; $r++) {
        $v = $wsSrc.Cells.Item($r, 1).Value2
        if ($null -ne $v -and $v -is [double] -and $v -gt 1) {
            $currentDate = [DateTime]::FromOADate($v).Date
        }

        # Sla over als datum al aanwezig is
        if ($null -eq $currentDate -or $existingDates.ContainsKey($currentDate)) { continue }

        # Kopieer rij voor rij
        for ($c = 1; $c -le 5; $c++) {
            $srcVal = $wsSrc.Cells.Item($r, $c).Value2
            if ($null -ne $srcVal) {
                # Kolom 1: datumcel — sla op als datum
                if ($c -eq 1) {
                    $wsDst.Cells.Item($appendRow, $c).Value2 = $srcVal
                    $wsDst.Cells.Item($appendRow, $c).NumberFormat = 'DD/MM/YYYY'
                } else {
                    $wsDst.Cells.Item($appendRow, $c).Value2 = $srcVal
                }
            }
        }
        $copiedDates.Add($currentDate.ToString('dd/MM/yyyy') + " (rij $appendRow)")
        $appendRow++
    }

    $wbDst.Save()
    Write-Host "Klaar! Gekopieerde rijen naar biljart moyenne.xlsx:"
    foreach ($d in $copiedDates) { Write-Host "  $d" }
    Write-Host "Nieuw totaal rijen in doel: $($appendRow - 1)"
}
finally {
    if ($wbDst) { $wbDst.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbDst) }
    if ($wsSrc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsSrc) }
    if ($wbSrc) { $wbSrc.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbSrc) }
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}


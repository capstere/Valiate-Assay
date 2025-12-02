param()

function Invoke-HelperLogger {
    param(
        [scriptblock]$Logger,
        [string]$Message,
        [string]$Severity = 'Info'
    )

    if ($Logger) {
        & $Logger -Message $Message -Severity $Severity
    }
}

function Write-SealTestSignatures {
    param(
        [OfficeOpenXml.ExcelPackage]$PkgNeg,
        [OfficeOpenXml.ExcelPackage]$PkgPos,
        [string]$Signature,
        [bool]$Overwrite = $false,
        [bool]$NegWritable = $true,
        [bool]$PosWritable = $true,
        [scriptblock]$Logger = $null
    )

    $result = [pscustomobject]@{
        NegWritten = 0
        PosWritten = 0
        NegSkipped = 0
        PosSkipped = 0
    }

    if (-not $Signature) {
        Invoke-HelperLogger -Logger $Logger -Message "Ingen signatur angiven (B47)." -Severity 'Error'
        return $result
    }

    foreach ($ws in $PkgNeg.Workbook.Worksheets) {
        if ($ws.Name -eq 'Worksheet Instructions') { continue }
        $h3 = ($ws.Cells['H3'].Text + '').Trim()
        if ($h3 -match '^[0-9]') {
            $existing = ($ws.Cells['B47'].Text + '').Trim()
            if ($existing -and -not $Overwrite) { $result.NegSkipped++; continue }
            $ws.Cells['B47'].Style.Numberformat.Format = '@'
            $ws.Cells['B47'].Value = $Signature
            $result.NegWritten++
        } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/\?A|NA|Tomt( innehåll)?)$') {
            break
        }
    }

    foreach ($ws in $PkgPos.Workbook.Worksheets) {
        if ($ws.Name -eq 'Worksheet Instructions') { continue }
        $h3 = ($ws.Cells['H3'].Text + '').Trim()
        if ($h3 -match '^[0-9]') {
            $existing = ($ws.Cells['B47'].Text + '').Trim()
            if ($existing -and -not $Overwrite) { $result.PosSkipped++; continue }
            $ws.Cells['B47'].Style.Numberformat.Format = '@'
            $ws.Cells['B47'].Value = $Signature
            $result.PosWritten++
        } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/\?A|NA|Tomt( innehåll)?)$') {
            break
        }
    }

    try {
        if ($result.NegWritten -gt 0 -and $NegWritable) {
            $PkgNeg.Save()
        } elseif ($result.NegWritten -gt 0) {
            Invoke-HelperLogger -Logger $Logger -Message "NEG kunde inte sparas (låst)." -Severity 'Warn'
        }

        if ($result.PosWritten -gt 0 -and $PosWritable) {
            $PkgPos.Save()
        } elseif ($result.PosWritten -gt 0) {
            Invoke-HelperLogger -Logger $Logger -Message "POS kunde inte sparas (låst)." -Severity 'Warn'
        }
    } catch {
        Invoke-HelperLogger -Logger $Logger -Message ("Kunde inte spara signatur i NEG/POS: {0}" -f $_.Exception.Message) -Severity 'Warn'
    }

    $msg = "Signatur satt: NEG $($result.NegWritten) blad (överhoppade $($result.NegSkipped)), POS $($result.PosWritten) blad (överhoppade $($result.PosSkipped))."
    Invoke-HelperLogger -Logger $Logger -Message $msg -Severity 'Info'
    return $result
}

function Get-SealTestViolations {
    param(
        [OfficeOpenXml.ExcelPackage]$PkgNeg,
        [OfficeOpenXml.ExcelPackage]$PkgPos
    )

    $violationsNeg = @()
    $violationsPos = @()
    $failNegCount = 0
    $failPosCount = 0

    foreach ($ws in $PkgNeg.Workbook.Worksheets) {
        if ($ws.Name -eq "Worksheet Instructions") { continue }
        if (-not $ws.Dimension) { continue }
        $obsC = Find-ObservationCol $ws
        for ($r = 3; $r -le 45; $r++) {
            $valK = $ws.Cells["K$r"].Value
            $textL = $ws.Cells["L$r"].Text
            if ($valK -ne $null -and $valK -is [double]) {
                if ($textL -eq "FAIL" -or $valK -le -3.0) {
                    $obsTxt = $ws.Cells[$r, $obsC].Text
                    $violationsNeg += [PSCustomObject]@{
                        Sheet      = $ws.Name
                        Cartridge  = $ws.Cells["H$r"].Text
                        InitialW   = $ws.Cells["I$r"].Value
                        FinalW     = $ws.Cells["J$r"].Value
                        WeightLoss = $valK
                        Status     = if ($textL -eq "FAIL") { "FAIL" } else { "Minusvärde" }
                        Obs        = $obsTxt
                    }
                    if ($textL -eq "FAIL") { $failNegCount++ }
                }
            }
        }
    }

    foreach ($ws in $PkgPos.Workbook.Worksheets) {
        if ($ws.Name -eq "Worksheet Instructions") { continue }
        if (-not $ws.Dimension) { continue }
        $obsC = Find-ObservationCol $ws
        for ($r = 3; $r -le 45; $r++) {
            $valK = $ws.Cells["K$r"].Value
            $textL = $ws.Cells["L$r"].Text
            if ($valK -ne $null -and $valK -is [double]) {
                if ($textL -eq "FAIL" -or $valK -le -3.0) {
                    $obsTxt = $ws.Cells[$r, $obsC].Text
                    $violationsPos += [PSCustomObject]@{
                        Sheet      = $ws.Name
                        Cartridge  = $ws.Cells["H$r"].Text
                        InitialW   = $ws.Cells["I$r"].Value
                        FinalW     = $ws.Cells["J$r"].Value
                        WeightLoss = $valK
                        Status     = if ($textL -eq "FAIL") { "FAIL" } else { "Minusvärde" }
                        Obs        = $obsTxt
                    }
                    if ($textL -eq "FAIL") { $failPosCount++ }
                }
            }
        }
    }

    return [pscustomobject]@{
        ViolationsNeg = $violationsNeg
        ViolationsPos = $violationsPos
        FailNegCount  = $failNegCount
        FailPosCount  = $failPosCount
    }
}

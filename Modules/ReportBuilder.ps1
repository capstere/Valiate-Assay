param()

function Invoke-ReportLogger {
    param(
        [scriptblock]$Logger,
        [string]$Message,
        [string]$Severity = 'Info'
    )

    if ($Logger) {
        & $Logger -Message $Message -Severity $Severity
    } else {
        Write-BackendLog -Message $Message -Severity $Severity
    }
}

function Write-SharePointSheet {
    param(
        [OfficeOpenXml.ExcelPackage]$OutputPkg,
        [pscustomobject]$RunConfig,
        [pscustomobject]$BatchInfo,
        [scriptblock]$Logger = $null
    )

    if (-not $OutputPkg) { return }

    try {
        if ($RunConfig.IncludeSharePointInfo -eq $false) {
            Invoke-ReportLogger -Logger $Logger -Message "SharePoint Info ej valt – hoppar över." -Severity 'Info'
            try { $old = $OutputPkg.Workbook.Worksheets["SharePoint Info"]; if ($old) { $OutputPkg.Workbook.Worksheets.Delete($old) } } catch {}
            return
        }

        $spOk = $false
        if ($global:SpConnected) { $spOk = $true }
        elseif (Get-Command Get-PnPConnection -ErrorAction SilentlyContinue) {
            try { $null = Get-PnPConnection; $spOk = $true } catch { $spOk = $false }
        }

        if (-not $spOk) {
            $errMsg = if ($global:SpError) { $global:SpError } else { 'Okänt fel' }
            Invoke-ReportLogger -Logger $Logger -Message ("SharePoint ej tillgängligt: {0}" -f $errMsg) -Severity 'Warn'
        }

        $batch = $BatchInfo.Batch
        $selPos = $RunConfig.SealPosPath
        $selNeg = $RunConfig.SealNegPath
        $lsp    = $RunConfig.Lsp

        if (-not $batch) {
            Invoke-ReportLogger -Logger $Logger -Message "Inget Batch # i POS/NEG – skriver tom SharePoint Info." -Severity 'Info'
            if (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue) {
                [void](Write-SPSheet-Safe -Pkg $OutputPkg -Rows @() -DesiredOrder @() -Batch '—')
            } else {
                $wsSp = $OutputPkg.Workbook.Worksheets["SharePoint Info"]; if ($wsSp) { $OutputPkg.Workbook.Worksheets.Delete($wsSp) }
                $wsSp = $OutputPkg.Workbook.Worksheets.Add("SharePoint Info")
                $wsSp.Cells[1,1].Value = "Rubrik"; $wsSp.Cells[1,2].Value = "Värde"
                $wsSp.Cells[2,1].Value = "Batch";  $wsSp.Cells[2,2].Value = "—"
                try { $wsSp.Cells[$wsSp.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
            }
            return
        }

        Invoke-ReportLogger -Logger $Logger -Message ("Batch hittad: {0}" -f $batch) -Severity 'Info'

        $fields = @(
            'Work_x0020_Center','Title','Batch_x0023_','SAP_x0020_Batch_x0023__x0020_2',
            'LSP','Material','BBD_x002f_SLED','Actual_x0020_startdate_x002f__x0',
            'PAL_x0020__x002d__x0020_Sample_x','Sample_x0020_Reagent_x0020_P_x00',
            'Order_x0020_quantity','Total_x0020_good','ITP_x0020_Test_x0020_results',
            'IPT_x0020__x002d__x0020_Testing_0','MES_x0020__x002d__x0020_Order_x0'
        )
        $renameMap = @{
            'Work Center'            = 'Work Center'
            'Title'                  = 'Order#'
            'Batch#'                 = 'SAP Batch#'
            'SAP Batch# 2'           = 'SAP Batch# 2'
            'LSP'                    = 'LSP'
            'Material'               = 'Material'
            'BBD/SLED'               = 'BBD/SLED'
            'Actual startdate/_x0'   = 'ROBAL - Actual start date/time'
            'PAL - Sample_x'         = 'Sample Reagent use'
            'Sample Reagent P'       = 'Sample Reagent P/N'
            'Order quantity'         = 'Order quantity'
            'Total good'             = 'ROBAL - Till Packning'
            'IPT Test results'       = 'IPT Test results'
            'IPT - Testing_0'        = 'IPT - Testing Finalized'
            'MES - Order_x0'         = 'MES Order'
        }

        $desiredOrder = @(
            'Work Center','Order#','SAP Batch#','SAP Batch# 2','LSP','Material','BBD/SLED',
            'ROBAL - Actual start date/time','Sample Reagent use','Sample Reagent P/N',
            'Order quantity','ROBAL - Till Packning','IPT Test results',
            'IPT - Testing Finalized','MES Order'
        )

        $dateFields      = @('BBD/SLED','ROBAL - Actual start date/time','IPT - Testing Finalized')
        $shortDateFields = @('BBD/SLED')
        $rows = @()
        if ($spOk) {
            try {
                $items = Get-PnPListItem -List "Cepheid | Production orders" -Fields $fields -PageSize 2000 -ErrorAction Stop
                $match = $items | Where-Object {
                    $v1 = $_['Batch_x0023_']; $v2 = $_['SAP_x0020_Batch_x0023__x0020_2']
                    $s1 = if ($null -ne $v1) { ([string]$v1).Trim() } else { '' }
                    $s2 = if ($null -ne $v2) { ([string]$v2).Trim() } else { '' }
                    $s1 -eq $batch -or $s2 -eq $batch
                } | Select-Object -First 1
                if ($match) {
                    foreach ($f in $fields) {
                        $val = $match[$f]
                        $label = $f -replace '_x0020_', ' ' `
                                     -replace '_x002d_', '-' `
                                     -replace '_x0023_', '#' `
                                     -replace '_x002f_', '/' `
                                     -replace '_x2013_', '–' `
                                     -replace '_x00',''
                        $label = $label.Trim()
                        if ($renameMap.ContainsKey($label)) { $label = $renameMap[$label] }
                        if ($null -ne $val -and $val -ne '') {
                            if ($val -eq $true) { $val = 'JA' }
                            elseif ($val -eq $false) { $val = 'NEJ' }

                            $dt = $null
                            if ($val -is [datetime]) { $dt = [datetime]$val }
                            else { try { $dt = [datetime]::Parse($val) } catch { $dt = $null } }
                            if ($dt -ne $null -and ($dateFields -contains $label)) {
                                $fmt = if ($shortDateFields -contains $label) { 'yyyy-MM-dd' } else { 'yyyy-MM-dd HH:mm' }
                                $val = $dt.ToString($fmt)
                            }
                            $rows += [pscustomobject]@{ Rubrik = $label; 'Värde' = $val }
                        }
                    }

                    if ($rows.Count -gt 0) {
                        $ordered = @()
                        foreach ($label in $desiredOrder) {
                            $hit = $rows | Where-Object { $_.Rubrik -eq $label } | Select-Object -First 1
                            if ($hit) { $ordered += $hit }
                        }
                        if ($ordered.Count -gt 0) { $rows = $ordered }
                    }
                    Invoke-ReportLogger -Logger $Logger -Message "SharePoint-post hittad – skriver blad." -Severity 'Info'
                } else {
                    Invoke-ReportLogger -Logger $Logger -Message "Ingen post i SharePoint för Batch=$batch." -Severity 'Info'
                }
            } catch {
                Invoke-ReportLogger -Logger $Logger -Message ("SP: Get-PnPListItem misslyckades: {0}" -f $_.Exception.Message) -Severity 'Warn'
            }
        }

        if (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue) {
            [void](Write-SPSheet-Safe -Pkg $OutputPkg -Rows $rows -DesiredOrder $desiredOrder -Batch $batch)
        } else {
            $wsSp = $OutputPkg.Workbook.Worksheets["SharePoint Info"]; if ($wsSp) { $OutputPkg.Workbook.Worksheets.Delete($wsSp) }
            $wsSp = $OutputPkg.Workbook.Worksheets.Add("SharePoint Info")
            $wsSp.Cells[1,1].Value = "Rubrik"; $wsSp.Cells[1,2].Value = "Värde"
            if ($rows.Count -gt 0) {
                $r=2; foreach($rowObj in $rows) { $wsSp.Cells[$r,1].Value = $rowObj.Rubrik; $wsSp.Cells[$r,2].Value = $rowObj.'Värde'; $r++ }
            } else {
                $wsSp.Cells[2,1].Value = "Batch";  $wsSp.Cells[2,2].Value = $batch
                $wsSp.Cells[3,1].Value = "Info";   $wsSp.Cells[3,2].Value = "No matching SharePoint row"
            }
            try { $wsSp.Cells[$wsSp.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
        }

        try {
            $wsSP = $OutputPkg.Workbook.Worksheets['SharePoint Info']
            if ($wsSP -and $wsSP.Dimension) {
                $labelCol = 1; $valueCol = 2
                for ($r = 1; $r -le $wsSP.Dimension.End.Row; $r++) {
                    if (($wsSP.Cells[$r,$labelCol].Text).Trim() -eq 'Sample Reagent use') {
                        $wsSP.Cells[$r,$valueCol].Style.WrapText = $true
                        $wsSP.Cells[$r,$valueCol].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
                        try { $wsSP.Column($valueCol).Width = 55 } catch {}
                        $wsSP.Row($r).CustomHeight = $true
                        break
                    }
                }
            }
        } catch {
            Invoke-ReportLogger -Logger $Logger -Message ("WrapText på 'Sample Reagent use' misslyckades: {0}" -f $_.Exception.Message) -Severity 'Warn'
        }
    } catch {
        Invoke-ReportLogger -Logger $Logger -Message ("SP-blad: {0}" -f $_.Exception.Message) -Severity 'Warn'
    }
}

function Apply-Watermarks {
    param(
        [OfficeOpenXml.ExcelPackage]$OutputPkg
    )

    try {
        foreach ($ws in $OutputPkg.Workbook.Worksheets) {
            try {
                $ws.HeaderFooter.OddHeader.CenteredText   = '&"Arial,Bold"&14 UNCONTROLLED'
                $ws.HeaderFooter.EvenHeader.CenteredText  = '&"Arial,Bold"&14 UNCONTROLLED'
                $ws.HeaderFooter.FirstHeader.CenteredText = '&"Arial,Bold"&14 UNCONTROLLED'
            } catch { Write-Warning "Kunde inte sätta header på blad: $($ws.Name)" }
        }
    } catch { Write-Warning "Fel vid vattenstämpling av rapporten." }
}

function Apply-TabColors {
    param(
        [OfficeOpenXml.ExcelPackage]$OutputPkg
    )

    try {
        $wsT = $OutputPkg.Workbook.Worksheets['Information'];            if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 52, 152, 219) }
        $wsT = $OutputPkg.Workbook.Worksheets['Infinity/GX'];            if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 33, 115, 70) }
        $wsT = $OutputPkg.Workbook.Worksheets['SharePoint Info'];        if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 0, 120, 212) }
    } catch {
        Write-Warning ("Kunde inte sätta tab-färg: {0}" -f $_.Exception.Message)
    }
}

function Build-AssayReport {
    param(
        [OfficeOpenXml.ExcelPackage]$OutputPkg,
        [OfficeOpenXml.ExcelPackage]$PkgNeg,
        [OfficeOpenXml.ExcelPackage]$PkgPos,
        [pscustomobject]$RunConfig,
        [pscustomobject]$SealAnalysis,
        [pscustomobject]$SignatureInfo,
        [pscustomobject]$InfoContext,
        [string]$ControlTab,
        [string]$RawDataPath,
        [string]$UtrustningListPath,
        [scriptblock]$Logger = $null
    )

    $sigSummary = Write-SealTestInfoSheet -OutputPkg $OutputPkg -PkgNeg $PkgNeg -PkgPos $PkgPos -SignatureInfo $SignatureInfo -Logger $Logger
    Write-StfSumSheet -OutputPkg $OutputPkg -SealAnalysis $SealAnalysis -Logger $Logger
    $csvStats = Write-InformationSheet -OutputPkg $OutputPkg -InfoContext $InfoContext -Logger $Logger
    Write-EquipmentSheet -OutputPkg $OutputPkg -InfoContext $InfoContext -UtrustningListPath $UtrustningListPath -Logger $Logger
    Copy-ControlMaterialSheet -OutputPkg $OutputPkg -ControlTab $ControlTab -RawDataPath $RawDataPath -Logger $Logger
    Write-SharePointSheet -OutputPkg $OutputPkg -RunConfig $RunConfig -BatchInfo $InfoContext.BatchInfo -Logger $Logger
    Apply-Watermarks -OutputPkg $OutputPkg
    Apply-TabColors  -OutputPkg $OutputPkg

    return [pscustomobject]@{
        SignatureSummary = $sigSummary
        CsvStats         = $csvStats
    }
}

function Copy-ControlMaterialSheet {
    param(
        [OfficeOpenXml.ExcelPackage]$OutputPkg,
        [string]$ControlTab,
        [string]$RawDataPath,
        [scriptblock]$Logger = $null
    )

    if (-not $OutputPkg) { return }

    try {
        if ($ControlTab -and (Test-Path -LiteralPath $RawDataPath)) {
            $srcPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($RawDataPath))
            try { $srcPkg.Workbook.Calculate() } catch {}
            $candidates = if ($ControlTab -match '\|') { $ControlTab -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @($ControlTab) }
            $srcWs = $null
            foreach ($cand in $candidates) {
                $srcWs = $srcPkg.Workbook.Worksheets | Where-Object { $_.Name -eq $cand } | Select-Object -First 1
                if ($srcWs) { break }
                $srcWs = $srcPkg.Workbook.Worksheets | Where-Object { $_.Name -like "*$cand*" } | Select-Object -First 1
                if ($srcWs) { break }
            }
            if ($srcWs) {
                $safeName = if ($srcWs.Name.Length -gt 31) { $srcWs.Name.Substring(0,31) } else { $srcWs.Name }
                $destName = $safeName; $n=1
                while ($OutputPkg.Workbook.Worksheets[$destName]) { $base = if ($safeName.Length -gt 27) { $safeName.Substring(0,27) } else { $safeName }; $destName = "$base($n)"; $n++ }
                $wsCM = $OutputPkg.Workbook.Worksheets.Add($destName, $srcWs)
                if ($wsCM.Dimension) {
                    foreach ($cell in $wsCM.Cells[$wsCM.Dimension.Address]) {
                        if ($cell.Formula -or $cell.FormulaR1C1) { $v=$cell.Value; $cell.Formula=$null; $cell.FormulaR1C1=$null; $cell.Value=$v }
                    }
                    try { $wsCM.Cells[$wsCM.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
                }
                Invoke-ReportLogger -Logger $Logger -Message ("Control Material kopierad: '{0}' → '{1}'" -f $srcWs.Name, $destName) -Severity 'Info'
            } else {
                Invoke-ReportLogger -Logger $Logger -Message "Hittade inget blad i kontrollfilen som matchar '$ControlTab'." -Severity 'Info'
            }
            $srcPkg.Dispose()
        } else {
            Invoke-ReportLogger -Logger $Logger -Message "Ingen Control-flik skapad (saknar mappning eller kontrollfil)." -Severity 'Info'
        }
    } catch {
        Invoke-ReportLogger -Logger $Logger -Message ("Control Material-fel: {0}" -f $_.Exception.Message) -Severity 'Warn'
    }
}

function Write-SealTestInfoSheet {
    param(
        [OfficeOpenXml.ExcelPackage]$OutputPkg,
        [OfficeOpenXml.ExcelPackage]$PkgNeg,
        [OfficeOpenXml.ExcelPackage]$PkgPos,
        [pscustomobject]$SignatureInfo,
        [scriptblock]$Logger = $null
    )

    $wsOut1 = $OutputPkg.Workbook.Worksheets["Seal Test Info"]
    if (-not $wsOut1) {
        Invoke-ReportLogger -Logger $Logger -Message "Fliken 'Seal Test Info' saknas i mallen" -Severity 'Error'
        return $null
    }

    for ($row = 3; $row -le 15; $row++) {
        $wsOut1.Cells["D$row"].Value = $null
        try { $wsOut1.Cells["D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None } catch {}
    }

    $fields = @(
        @{ Label = "ROBAL";                         Cell = "F2"  }
        @{ Label = "Part Number";                   Cell = "B2"  }
        @{ Label = "Batch Number";                  Cell = "D2"  }
        @{ Label = "Cartridge Number (LSP)";        Cell = "B6"  }
        @{ Label = "PO Number";                     Cell = "B10" }
        @{ Label = "Assay Family";                  Cell = "D10" }
        @{ Label = "Weight Loss Spec";              Cell = "F10" }
        @{ Label = "Balance ID Number";             Cell = "B14" }
        @{ Label = "Balance Cal Due Date";          Cell = "D14" }
        @{ Label = "Vacuum Oven ID Number";         Cell = "B20" }
        @{ Label = "Vacuum Oven Cal Due Date";      Cell = "D20" }
        @{ Label = "Timer ID Number";               Cell = "B25" }
        @{ Label = "Timer Cal Due Date";            Cell = "D25" }
    )

    $forceText = @("ROBAL","Part Number","Batch Number","Cartridge Number (LSP)","PO Number","Assay Family","Balance ID Number","Vacuum Oven ID Number","Timer ID Number")
    $mismatchFields = $fields[0..6] | ForEach-Object { $_.Label }

    $row = 3
    foreach ($f in $fields) {
        $valNeg=''; $valPos=''
        foreach ($wsN in $PkgNeg.Workbook.Worksheets) {
            if ($wsN.Name -eq "Worksheet Instructions") { continue }
            $cell = $wsN.Cells[$f.Cell]
            if ($cell.Value -ne $null) { if ($cell.Value -is [datetime]) { $valNeg = $cell.Value.ToString('MMM-yy') } else { $valNeg = $cell.Text }; break }
        }
        foreach ($wsP in $PkgPos.Workbook.Worksheets) {
            if ($wsP.Name -eq "Worksheet Instructions") { continue }
            $cell = $wsP.Cells[$f.Cell]
            if ($cell.Value -ne $null) { if ($cell.Value -is [datetime]) { $valPos = $cell.Value.ToString('MMM-yy') } else { $valPos = $cell.Text }; break }
        }

        if ($forceText -contains $f.Label) {
            $wsOut1.Cells["B$row"].Style.Numberformat.Format = '@'
            $wsOut1.Cells["C$row"].Style.Numberformat.Format = '@'
        }

        $wsOut1.Cells["B$row"].Value = $valNeg
        $wsOut1.Cells["C$row"].Value = $valPos
        $wsOut1.Cells["B$row"].Style.Border.Right.Style = "Medium"
        $wsOut1.Cells["C$row"].Style.Border.Left.Style  = "Medium"

        if ($mismatchFields -contains $f.Label -and $valNeg -ne $valPos) {
            $wsOut1.Cells["D$row"].Value = "Mismatch"

            Style-Cell $wsOut1.Cells["D$row"] $true "FF0000" "Medium" "FFFFFF"
            Invoke-ReportLogger -Logger $Logger -Message ("Mismatch: {0} ({1} vs {2})" -f $f.Label, $valNeg, $valPos) -Severity 'Warn'
        }
        $row++
    }

    $testersNeg = @(); $testersPos = @()
    foreach ($s in $PkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) { $t=$s.Cells["B43"].Text; if ($t) { $testersNeg += ($t -split ",") } }
    foreach ($s in $PkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) { $t=$s.Cells["B43"].Text; if ($t) { $testersPos += ($t -split ",") } }
    $testersNeg = $testersNeg | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
    $testersPos = $testersPos | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique

    $wsOut1.Cells["B16"].Value = "Name of Tester"
    $wsOut1.Cells["B16:C16"].Merge = $true
    $wsOut1.Cells["B16"].Style.HorizontalAlignment = "Center"

    $maxTesters = [Math]::Max($testersNeg.Count, $testersPos.Count)
    $initialRows = 11
    if ($maxTesters -lt $initialRows) { $wsOut1.DeleteRow(17 + $maxTesters, $initialRows - $maxTesters) }
    if ($maxTesters -gt $initialRows) {
        $rowsToAdd = $maxTesters - $initialRows
        $lastRow = 16 + $initialRows
        for ($i = 1; $i -le $rowsToAdd; $i++) { $wsOut1.InsertRow($lastRow + 1, 1, $lastRow) }
    }
    for ($i = 0; $i -lt $maxTesters; $i++) {
        $rowIndex = 17 + $i
        $wsOut1.Cells["A$rowIndex"].Value = $null
        $wsOut1.Cells["B$rowIndex"].Value = if ($i -lt $testersNeg.Count) { $testersNeg[$i] } else { "N/A" }
        $wsOut1.Cells["C$rowIndex"].Value = if ($i -lt $testersPos.Count) { $testersPos[$i] } else { "N/A" }

        $topStyle    = if ($i -eq 0) { "Medium" } else { "Thin" }
        $bottomStyle = if ($i -eq $maxTesters - 1) { "Medium" } else { "Thin" }
        foreach ($col in @("B","C")) {
            $cell = $wsOut1.Cells["$col$rowIndex"]
            $cell.Style.Border.Top.Style    = $topStyle
            $cell.Style.Border.Bottom.Style = $bottomStyle
            $cell.Style.Border.Left.Style   = "Medium"
            $cell.Style.Border.Right.Style  = "Medium"
            $cell.Style.Fill.PatternType = "Solid"
            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
        }
    }

    function Set-MergedWrapAutoHeight {
        param([OfficeOpenXml.ExcelWorksheet]$Sheet,[int]$RowIndex,[int]$ColStart=2,[int]$ColEnd=3,[string]$Text)
        $rng = $Sheet.Cells[$RowIndex, $ColStart, $RowIndex, $ColEnd]
        $rng.Style.WrapText = $true
        $rng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None
        $Sheet.Row($RowIndex).CustomHeight = $false

        try {
            $wChars = [Math]::Floor(($Sheet.Column($ColStart).Width + $Sheet.Column($ColEnd).Width) - 2); if ($wChars -lt 1) { $wChars = 1 }
            $segments = $Text -split "(\r\n|\n|\r)"; $lineCount = 0
            foreach ($seg in $segments) { if (-not $seg) { $lineCount++ } else { $lineCount += [Math]::Ceiling($seg.Length / $wChars) } }
            if ($lineCount -lt 1) { $lineCount = 1 }
            $targetHeight = [Math]::Max(15, [Math]::Ceiling(15 * $lineCount * 2.15))
            if ($Sheet.Row($RowIndex).Height -lt $targetHeight) {
                $Sheet.Row($RowIndex).Height = $targetHeight
                $Sheet.Row($RowIndex).CustomHeight = $true
            }
        } catch { $Sheet.Row($RowIndex).CustomHeight = $false }
    }

    $negSigSet = $SignatureInfo.NegSigSet
    $posSigSet = $SignatureInfo.PosSigSet
    $signToWrite = $SignatureInfo.SignatureToWrite

    $negSet = New-Object 'System.Collections.Generic.HashSet[string]'
    $posSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($n in $negSigSet.NormSet) { [void]$negSet.Add($n) }
    foreach ($p in $posSigSet.NormSet) { [void]$posSet.Add($p) }
    $hasNeg = ($negSet.Count -gt 0)
    $hasPos = ($posSet.Count -gt 0)
    $onlyNeg = @(); $onlyPos = @(); $sigMismatch = $false
    if ($hasNeg -and $hasPos) {
        foreach ($n in $negSet) { if (-not $posSet.Contains($n)) { $onlyNeg += $n } }
        foreach ($p in $posSet) { if (-not $negSet.Contains($p)) { $onlyPos += $p } }
        $sigMismatch = ($onlyNeg.Count -gt 0 -or $onlyPos.Count -gt 0)
    } else {
        $sigMismatch = $false
    }

    $mismatchSheets = @()
    if ($sigMismatch) {
        foreach ($k in $onlyNeg) {
            $raw = if ($negSigSet.RawByNorm.ContainsKey($k)) { $negSigSet.RawByNorm[$k] } else { $k }
            $where = if ($negSigSet.Occ.ContainsKey($k)) { ($negSigSet.Occ[$k] -join ', ') } else { '—' }
            $mismatchSheets += ("NEG: " + $raw + "  [Blad: " + $where + "]")
        }

        foreach ($k in $onlyPos) {
            $raw = if ($posSigSet.RawByNorm.ContainsKey($k)) { $posSigSet.RawByNorm[$k] } else { $k }
            $where = if ($posSigSet.Occ.ContainsKey($k)) { ($posSigSet.Occ[$k] -join ', ') } else { '—' }
            $mismatchSheets += ("POS: " + $raw + "  [Blad: " + $where + "]")
        }
        Invoke-ReportLogger -Logger $Logger -Message "Mismatch: Print Full Name, Sign, and Date (NEG vs POS)" -Severity 'Warn'
    }

    $signRow = 17 + $maxTesters + 3
    $displaySignNeg = $null; $displaySignPos = $null
    if ($signToWrite) { $displaySignNeg = $signToWrite; $displaySignPos = $signToWrite }
    else {
        $displaySignNeg = if ($negSigSet.RawFirst) { $negSigSet.RawFirst } else { '—' }
        $displaySignPos = if ($posSigSet.RawFirst) { $posSigSet.RawFirst } else { '—' }
    }

    $wsOut1.Cells["B$signRow"].Style.Numberformat.Format = '@'
    $wsOut1.Cells["C$signRow"].Style.Numberformat.Format = '@'
    $wsOut1.Cells["B$signRow"].Value = $displaySignNeg
    $wsOut1.Cells["C$signRow"].Value = $displaySignPos
    foreach ($col in @('B','C')) {
        $cell = $wsOut1.Cells["${col}$signRow"]
        Style-Cell $cell $false 'CCFFFF' 'Medium' $null
        $cell.Style.HorizontalAlignment = 'Center'
    }

    try { $wsOut1.Column(2).Width = 40; $wsOut1.Column(3).Width = 40 } catch {}
    try { $wsOut1.Column(4).Width = 10 } catch {}

    if ($sigMismatch) {
        $mismatchCell = $wsOut1.Cells["D$signRow"]
        $mismatchCell.Value = 'Mismatch'
        Style-Cell $mismatchCell $true 'FF0000' 'Medium' 'FFFFFF'

        if ($mismatchSheets.Count -gt 0) {
            for ($j = 0; $j -lt $mismatchSheets.Count; $j++) {
                $rowIdx = $signRow + 1 + $j
                try { $wsOut1.Cells["B$rowIdx:C$rowIdx"].Merge = $true } catch {}
                $text = $mismatchSheets[$j]
                $wsOut1.Cells["B$rowIdx"].Value = $text
                foreach ($mc in $wsOut1.Cells["B$rowIdx:C$rowIdx"]) { Style-Cell $mc $false 'CCFFFF' 'Medium' $null }
                $wsOut1.Cells["B$rowIdx:C$rowIdx"].Style.HorizontalAlignment = 'Center'

                if ($text -like 'NEG:*' -or $text -like 'POS:*') {
                    Set-MergedWrapAutoHeight -Sheet $wsOut1 -RowIndex $rowIdx -ColStart 2 -ColEnd 3 -Text $text
                }
            }
        }
    }

    return [pscustomobject]@{
        SigMismatch    = $sigMismatch
        MismatchSheets = $mismatchSheets
    }
}

# Write-StfSumSheet, Write-InformationSheet, Write-EquipmentSheet, Copy-ControlMaterialSheet,
# Write-SharePointSheet, Apply-Watermarks, Apply-TabColors, Build-AssayReport will be implemented below.

function Write-StfSumSheet {
    param(
        [OfficeOpenXml.ExcelPackage]$OutputPkg,
        [pscustomobject]$SealAnalysis,
        [scriptblock]$Logger = $null
    )

    $wsOut2 = $OutputPkg.Workbook.Worksheets["STF Sum"]
    if (-not $wsOut2) {
        Invoke-ReportLogger -Logger $Logger -Message "Fliken 'STF Sum' saknas i mallen!" -Severity 'Error'
        return
    }

    $violationsNeg = $SealAnalysis.ViolationsNeg
    $violationsPos = $SealAnalysis.ViolationsPos
    $failNegCount  = $SealAnalysis.FailNegCount
    $failPosCount  = $SealAnalysis.FailPosCount

    $totalRows = $violationsNeg.Count + $violationsPos.Count
    $currentRow = 2

    if ($totalRows -eq 0) {
        Invoke-ReportLogger -Logger $Logger -Message "Seal Test hittades" -Severity 'Info'
        $wsOut2.Cells["B1:H1"].Value = $null
        $wsOut2.Cells["A1"].Value = "Inga STF hittades!"
        Style-Cell $wsOut2.Cells["A1"] $true "D9EAD3" "Medium" "006100"
        $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
        if ($wsOut2.Dimension -and $wsOut2.Dimension.End.Row -gt 1) { $wsOut2.DeleteRow(2, $wsOut2.Dimension.End.Row - 1) }

        return
    }

    Invoke-ReportLogger -Logger $Logger -Message ("{0} avvikelser i NEG, {1} i POS" -f $failNegCount, $failPosCount) -Severity 'Warn'
    $oldDataRows = 0
    if ($wsOut2.Dimension) { $oldDataRows = $wsOut2.Dimension.End.Row - 1; if ($oldDataRows -lt 0) { $oldDataRows = 0 } }
    if ($totalRows -lt $oldDataRows) { $wsOut2.DeleteRow(2 + $totalRows, $oldDataRows - $totalRows) }
    elseif ($totalRows -gt $oldDataRows) { $wsOut2.InsertRow(2 + $oldDataRows, $totalRows - $oldDataRows, 1 + $oldDataRows) }

    $currentRow = 2
    foreach ($v in $violationsNeg) {
        $wsOut2.Cells["A$currentRow"].Value = "NEG"
        $wsOut2.Cells["B$currentRow"].Value = $v.Sheet
        $wsOut2.Cells["C$currentRow"].Value = $v.Cartridge
        $wsOut2.Cells["D$currentRow"].Value = $v.InitialW
        $wsOut2.Cells["E$currentRow"].Value = $v.FinalW
        $wsOut2.Cells["F$currentRow"].Value = [Math]::Round($v.WeightLoss, 1)
        $wsOut2.Cells["G$currentRow"].Value = $v.Status
        $wsOut2.Cells["H$currentRow"].Value = if ([string]::IsNullOrWhiteSpace($v.Obs)) { 'NA' } else { $v.Obs }
        Style-Cell $wsOut2.Cells["A$currentRow"] $true "B5E6A2" "Medium" $null
        $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.PatternType = "Solid"
        $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
        $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.PatternType = "Solid"
        $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF99"))
        $wsOut2.Cells["H$currentRow"].Style.Fill.PatternType = "Solid"
        $wsOut2.Cells["H$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))

        if ($v.Status -in @("FAIL","Minusvärde")) {
            $wsOut2.Cells["F$currentRow"].Style.Font.Bold = $true
            $wsOut2.Cells["F$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
            $wsOut2.Cells["G$currentRow"].Style.Font.Bold = $true
            $wsOut2.Cells["G$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
        }
        Set-RowBorder -ws $wsOut2 -row $currentRow -firstRow 2 -lastRow ($totalRows + 1)
        $currentRow++
    }
    foreach ($v in $violationsPos) {
        $wsOut2.Cells["A$currentRow"].Value = "POS"
        $wsOut2.Cells["B$currentRow"].Value = $v.Sheet
        $wsOut2.Cells["C$currentRow"].Value = $v.Cartridge
        $wsOut2.Cells["D$currentRow"].Value = $v.InitialW
        $wsOut2.Cells["E$currentRow"].Value = $v.FinalW
        $wsOut2.Cells["F$currentRow"].Value = [Math]::Round($v.WeightLoss, 1)
        $wsOut2.Cells["G$currentRow"].Value = $v.Status
        $wsOut2.Cells["H$currentRow"].Value = if ($v.Obs) { $v.Obs } else { 'NA' }

        Style-Cell $wsOut2.Cells["A$currentRow"] $true "FFB3B3" "Medium" $null
        $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.PatternType = "Solid"
        $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
        $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.PatternType = "Solid"
        $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF99"))
        $wsOut2.Cells["H$currentRow"].Style.Fill.PatternType = "Solid"
        $wsOut2.Cells["H$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))

        if ($v.Status -in @("FAIL","Minusvärde")) {
            $wsOut2.Cells["F$currentRow"].Style.Font.Bold = $true
            $wsOut2.Cells["F$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
            $wsOut2.Cells["G$currentRow"].Style.Font.Bold = $true
            $wsOut2.Cells["G$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
        }
        Set-RowBorder -ws $wsOut2 -row $currentRow -firstRow 2 -lastRow ($totalRows + 1)
        $currentRow++
    }

    $wsOut2.Cells.Style.WrapText = $false
    $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
    try { $wsOut2.Cells[2,6,([Math]::Max($currentRow-1,2)),6].Style.Numberformat.Format = '0.0' } catch {}
    if ($wsOut2.Dimension) { $wsOut2.Cells[$wsOut2.Dimension.Address].AutoFitColumns() }
}

function Write-InformationSheet {
    param(
        [OfficeOpenXml.ExcelPackage]$OutputPkg,
        [pscustomobject]$InfoContext,
        [scriptblock]$Logger = $null
    )

    function Log {
        param(
            [string]$Message,
            [string]$Severity = 'Info'
        )
        Invoke-ReportLogger -Logger $Logger -Message $Message -Severity $Severity
    }

    $selCsv       = $InfoContext.CsvPath
    $compilingSummary = $InfoContext.CompilingSummary
    $compilingData    = $InfoContext.CompilingData
    $runAssay     = $InfoContext.RunAssay
    $assayName    = $InfoContext.AssayName
    $assayVersion = $InfoContext.AssayVersion
    $lsp          = $InfoContext.Lsp
    $lspSummary   = $InfoContext.LspSummary
    $testCount    = $InfoContext.TestCount
    $miniVal      = $InfoContext.MiniVal
    $selLsp       = $InfoContext.SelLsp
    $selPos       = $InfoContext.SelPos
    $selNeg       = $InfoContext.SelNeg
    $headerWs     = $InfoContext.HeaderWs
    $headerPos    = $InfoContext.HeaderPos
    $headerNeg    = $InfoContext.HeaderNeg
    $wsHeaderCheck= $InfoContext.WsHeaderCheck
    $eqInfo       = $InfoContext.EqInfo
    $tsControls   = $InfoContext.TsControls
    $worksheetControlMaterials = $InfoContext.WorksheetControlMaterials
    $controlMap   = $InfoContext.ControlMap
    $batchInfo    = $InfoContext.BatchInfo
    $batch        = $InfoContext.Batch
    $reportOptions= $InfoContext.ReportOptions

    if (-not $compilingSummary) { $compilingSummary = Get-CompilingAnalysis -Data @() }

    if (-not $tsControls -or $tsControls.Count -eq 0) {
        Log "Kontrollmaterial kunde inte hittas i Test Summary – visar CSV-baserade kontroller." 'Warn'
    }

    if (-not (Get-Command Add-Hyperlink -ErrorAction SilentlyContinue)) {
        function Add-Hyperlink {
            param([OfficeOpenXml.ExcelRange]$Cell,[string]$Text,[string]$Url)
            try {
                $Cell.Value = $Text
                $Cell.Hyperlink = [Uri]$Url
                $Cell.Style.Font.UnderLine = $true
                $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0,102,204))
            } catch {}
        }
    }
    if (-not (Get-Command Find-RegexCell -ErrorAction SilentlyContinue)) {
        function Find-RegexCell {
            param([OfficeOpenXml.ExcelWorksheet]$Ws,[regex]$Rx,[int]$MaxRows=1000,[int]$MaxCols=200)
            if (-not $Ws -or -not $Ws.Dimension) { return $null }
            $rMax = [Math]::Min($Ws.Dimension.End.Row, $MaxRows)
            $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
            for ($r=1; $r -le $rMax; $r++) {
                for ($c=1; $c -le $cMax; $c++) {
                    $t = Normalize-HeaderText ($Ws.Cells[$r,$c].Text + '')
                    if ($t -and $Rx.IsMatch($t)) { return @{Row=$r;Col=$c;Text=$t} }
                }
            }
            return $null
        }
    }
    if (-not (Get-Command Get-SealHeaderDocInfo -ErrorAction SilentlyContinue)) {
        function Get-SealHeaderDocInfo {
            param([OfficeOpenXml.ExcelPackage]$Pkg)
            $result = [pscustomobject]@{ Raw=''; DocNo=''; Rev='' }
            if (-not $Pkg) { return $result }
            $ws = $Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
            if (-not $ws) { return $result }
            try {
                $lt = ($ws.HeaderFooter.OddHeader.LeftAlignedText + '').Trim()
                if (-not $lt) { $lt = ($ws.HeaderFooter.EvenHeader.LeftAlignedText + '').Trim() }
                $result.Raw = $lt
                $rx = [regex]'(?i)(?:document\s*(?:no|nr|#|number)\s*[:#]?\s*([A-Z0-9\-_\.\/]+))?.*?(?:rev(?:ision)?\.?\s*[:#]?\s*([A-Z0-9\-_\.]+))?'
                $m = $rx.Match($lt)
                if ($m.Success) {
                    if ($m.Groups[1].Value) { $result.DocNo = $m.Groups[1].Value.Trim() }
                    if ($m.Groups[2].Value) { $result.Rev   = $m.Groups[2].Value.Trim() }
                }
            } catch {}
            return $result
        }
    }

    $wsInfo = $OutputPkg.Workbook.Worksheets['Information']
    if (-not $wsInfo) { $wsInfo = $OutputPkg.Workbook.Worksheets.Add('Information') }
    try { $wsInfo.Cells.Clear() | Out-Null } catch {}
    try { $wsInfo.Cells.Style.Font.Name='Arial'; $wsInfo.Cells.Style.Font.Size=11 } catch {}

    $csvLines = @()
    try {
        if ($selCsv -and (Test-Path -LiteralPath $selCsv)) {
            try { $csvLines = Get-Content -LiteralPath $selCsv } catch { Log ("Kunde inte läsa CSV: " + $_.Exception.Message) 'Warn' }
            if (-not $compilingData -or $compilingData.Count -eq 0) {
                try { $compilingData = Convert-XpertSummaryCsvToData -Path $selCsv } catch { Log ("Convert-XpertSummaryCsvToData: " + $_.Exception.Message) 'Warn' }
            }
            if (-not $compilingSummary -or -not $compilingSummary.TestCount) {
                try { $compilingSummary = Get-CompilingAnalysis -Data $compilingData } catch { Log ("Get-CompilingAnalysis: " + $_.Exception.Message) 'Warn' }
            }
        }
    } catch {}
    if (-not $compilingSummary) { $compilingSummary = Get-CompilingAnalysis -Data @() }

    $assayName    = if ($compilingSummary.AssayName) { $compilingSummary.AssayName } elseif ($runAssay) { $runAssay } else { '' }
    $assayVersion = $compilingSummary.AssayVersion
    $lspSummary   = if ($compilingSummary.LspSummary) { $compilingSummary.LspSummary } else { $lspSummary }
    $testCount    = if ($compilingSummary.TestCount) { $compilingSummary.TestCount } else { $testCount }
    $csvStats     = [pscustomobject]@{ TestCount = $testCount }

    $assayForMacro = ''
    if ($runAssay) {
        $assayForMacro = $runAssay
    } elseif ($assayName) {
        $assayForMacro = $assayName
    } elseif ($lspSummary) {
        $assayForMacro = $lspSummary
    }

    $wsInfo.Column(1).Width = 18
    $wsInfo.Column(2).Width = 28
    $wsInfo.Column(3).Width = 32
    $wsInfo.Column(4).Width = 28
    $wsInfo.Column(5).Width = 45

    $row = 3

    $wsInfo.Cells['A1'].Value = 'Information'
    $wsInfo.Cells['A1'].Style.Font.Bold = $true
    $wsInfo.Cells['A1'].Style.Font.Size = 14

    $infoHeaderColor = [System.Drawing.Color]::LightGray

    function Convert-ToInfoValue {
        param([object]$Value)
        if ($null -eq $Value) { return '' }
        if ($Value -is [System.Array]) { return ($Value -join '; ') }
        return $Value
    }

    function Join-LimitedText {
        param(
            [object[]]$Items,
            [int]$Max = 6
        )
        if (-not $Items -or $Items.Count -eq 0) { return '' }
        $arr = @($Items)
        if ($arr.Count -gt $Max) {
            $shown = $arr[0..($Max-1)]
            $rest  = $arr.Count - $shown.Count
            return (($shown -join '; ') + " (+$rest till)")
        }
        return ($arr -join '; ')
    }

    function Get-ReportOption {
        param(
            [string]$Name,
            [bool]$Default = $true
        )
        if (-not $Name) { return $Default }
        try {
            if ($reportOptions -and $reportOptions.ContainsKey($Name)) {
                return [bool]$reportOptions[$Name]
            }
        } catch {}
        return $Default
    }

    function Set-InfoHeaderRow {
        param(
            [int]$Row,
            [int]$FromCol,
            [int]$ToCol,
            [string]$Text,
            [System.Drawing.Color]$Color
        )

        $cell = $wsInfo.Cells[$Row,$FromCol]
        $cell.Value = $Text

        $range = $wsInfo.Cells[$Row,$FromCol,$Row,$ToCol]

        if ($ToCol -gt $FromCol) {
            try { $range.Merge = $true } catch {}
        }

        try {
            $range.Style.Font.Bold = $true
            $range.Style.Fill.PatternType = 'Solid'
            $range.Style.Fill.BackgroundColor.SetColor($Color)
        } catch {}
    }

    function Add-InfoBorder {
        param(
            [int]$FromRow,
            [int]$FromCol,
            [int]$ToRow,
            [int]$ToCol
        )
        if ($ToRow -lt $FromRow -or $ToCol -lt $FromCol) { return }
        $rng = $wsInfo.Cells[$FromRow,$FromCol,$ToRow,$ToCol]
        $rng.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $rng.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $rng.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $rng.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    }

    # A. Körningsöversikt
    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Körningsöversikt' -Color $infoHeaderColor
    $row++
    $wsInfo.Cells[$row,1].Value = 'Fält'
    $wsInfo.Cells[$row,2].Value = 'Värde'
    $wsInfo.Cells[$row,1,$row,2].Style.Font.Bold = $true
    $wsInfo.Cells[$row,1,$row,2].Style.Fill.PatternType = 'Solid'
    $wsInfo.Cells[$row,1,$row,2].Style.Fill.BackgroundColor.SetColor($infoHeaderColor)
    $row++
    $overviewRows = @(
        @{K='Skriptversion'; V=$InfoContext.ScriptVersion},
        @{K='Användare';     V=[Environment]::UserName},
        @{K='Datum Tid';     V=(Get-Date).ToString('yyyy-MM-dd HH:mm')},
        @{K='Minitab Macro'; V=$miniVal},
        @{K='CSV';           V=if ($selCsv) { Split-Path $selCsv -Leaf } else { 'N/A' }},
        @{K='Assay';         V=$assayName},
        @{K='Assay Version'; V=$assayVersion},
        @{K='Reagent Lot (LSP)'; V=if ($lspSummary) { $lspSummary } else { $lsp }},
        @{K='Antal tester';  V=$testCount}
    )
    foreach ($ov in $overviewRows) {
        $wsInfo.Cells[$row,1].Value = Convert-ToInfoValue $ov.K
        $wsInfo.Cells[$row,1].Style.Font.Bold = $true
        $wsInfo.Cells[$row,2].Value = if ($ov.V) { Convert-ToInfoValue $ov.V } else { 'N/A' }
        $row++
    }
    $ovStart = $row - $overviewRows.Count - 1
    Add-InfoBorder -FromRow $ovStart -FromCol 1 -ToRow ($row-1) -ToCol 2
    $row += 1

    # B. Kontrollmaterial
    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Kontrollmaterial' -Color $infoHeaderColor
    $row++
    $showDetails = Get-ReportOption -Name 'IncludeControlDetails' -Default $true

    [string[]]$cmHeader = if ($showDetails) {
        'Kod','Namn','Kategori','Källa','Antal'
    } else {
        'Kod','Namn','Källa','Antal'
    }

    for ($i = 0; $i -lt $cmHeader.Length; $i++) {
        $cell = $wsInfo.Cells.Get_Item($row, ($i + 1))
        if ($cell) {
            $cell.Value = $cmHeader[$i]
        }
    }

    $wsInfo.Cells[$row,1,$row,$cmHeader.Length].Style.Font.Bold = $true
    $wsInfo.Cells[$row,1,$row,$cmHeader.Length].Style.Fill.PatternType = 'Solid'
    $wsInfo.Cells[$row,1,$row,$cmHeader.Length].Style.Fill.BackgroundColor.SetColor($infoHeaderColor)
    $row++

    $csvCounts = @{}
    if ($compilingSummary -and $compilingSummary.ControlMaterialCount) {
        foreach ($code in $compilingSummary.ControlMaterialCount.Keys) {
            $key = ($code + '').ToUpper()
            $csvCounts[$key] = [int]$compilingSummary.ControlMaterialCount[$code]
        }
    }

    $tsCounts = @{}
    if ($tsControls -and $tsControls.Count -gt 0) {
        foreach ($ctrl in $tsControls) {
            if (-not $ctrl.PartNos) { continue }
            foreach ($pn in $ctrl.PartNos) {
                $key = ($pn + '').ToUpper()
                if (-not $tsCounts.ContainsKey($key)) { $tsCounts[$key] = 0 }
                $tsCounts[$key]++
            }
        }
    }

    $allKeys = @()
    $allKeys += $csvCounts.Keys
    $allKeys += $tsCounts.Keys
    $allKeys = @($allKeys | Sort-Object -Unique)

    $cmEntries = @()
    foreach ($key in $allKeys) {
        $csvCount = if ($csvCounts.ContainsKey($key)) { [int]$csvCounts[$key] } else { 0 }
        $tsCount  = if ($tsCounts.ContainsKey($key))  { [int]$tsCounts[$key] }  else { 0 }
        $source = if ($csvCount -gt 0 -and $tsCount -gt 0) { 'CSV+TS' } elseif ($csvCount -gt 0) { 'CSV' } else { 'TS' }
        $countCombined = $csvCount + $tsCount

        $name = ''
        $cat  = ''
        try {
            if ($controlMap -and $controlMap.PartNoIndex -and $controlMap.PartNoIndex.ContainsKey($key)) {
                $info = $controlMap.PartNoIndex[$key]
                $name = if ($info.NameOfficial) { $info.NameOfficial } else { '' }
                $cat  = if ($info.Category)     { $info.Category }     else { '' }
            }
        } catch {}

        if (-not $name) { $name = 'Okänd' }

        $cmEntries += [pscustomobject]@{
            Code      = $key
            Name      = $name
            Category  = $cat
            Source    = $source
            Count     = $countCombined
            CsvCount  = $csvCount
            TsCount   = $tsCount
        }
    }

    $showOnlyMismatches = Get-ReportOption -Name 'HighlightMismatchesOnly' -Default $false
    if ($showOnlyMismatches) {
        $cmEntries = $cmEntries | Where-Object { $_.Source -ne 'CSV+TS' -or $_.Name -eq 'Okänd' }
    }

    $cmEntries = $cmEntries | Sort-Object @{Expression={-($_.Count)}}, @{Expression={$_.Code}}
    $cmStart = $row - 1

    if ($cmEntries.Count -gt 0) {
        foreach ($ent in $cmEntries) {
            $wsInfo.Cells.Get_Item($row,1).Value = $ent.Code
            $wsInfo.Cells.Get_Item($row,2).Value = $ent.Name
            $colOffset = 2
            if ($showDetails) {
                $wsInfo.Cells.Get_Item($row,3).Value = if ($ent.Category) { $ent.Category } else { '' }
                $colOffset = 3
            }
            $wsInfo.Cells.Get_Item($row,$colOffset+1).Value = $ent.Source
            $wsInfo.Cells.Get_Item($row,$colOffset+2).Value = $ent.Count
            $row++
        }
    } else {
        $wsInfo.Cells.Get_Item($row,1).Value = 'Inga kontrollmaterial hittades'
        $wsInfo.Cells[$row,1,$row,$cmHeader.Length].Merge = $true
        $row++
    }
    Add-InfoBorder -FromRow $cmStart -FromCol 1 -ToRow ($row-1) -ToCol $cmHeader.Length
    $row += 1

    # C. Kontrolltyp-design
    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Kontrolltyp-design' -Color $infoHeaderColor
    $row++

    [string[]]$ctHeader = @(
        'ControlType',
        'Label',
        'Förväntat',
        'Bag-intervall',
        'Replikat-intervall',
        'Faktiskt',
        'Saknas',
        'Status'
    )

    for ($i = 0; $i -lt $ctHeader.Length; $i++) {
        $cell = $wsInfo.Cells.Get_Item($row, ($i + 1))
        if ($cell) {
            $cell.Value = $ctHeader[$i]
        }
    }

    $hdrRange = $wsInfo.Cells[$row,1,$row,$ctHeader.Length]
    $hdrRange.Style.Font.Bold = $true
    $hdrRange.Style.Fill.PatternType = 'Solid'
    $hdrRange.Style.Fill.BackgroundColor.SetColor($infoHeaderColor)

    $row++
    $ctStart = $row - 1

    if ($compilingSummary.ControlTypeStats.Count -gt 0) {
        foreach ($ct in ($compilingSummary.ControlTypeStats | Sort-Object { [int]($_.ControlType) })) {
            $expectedVal = $ct.ExpectedCount
            if ($null -eq $expectedVal -and $ct.DesignExpected -ne $null) { $expectedVal = $ct.DesignExpected }
            $bagText  = if ($ct.BagRangeText) { $ct.BagRangeText } else { '' }
            $repText  = if ($ct.ReplicateRangeText) { $ct.ReplicateRangeText } else { '' }

            $wsInfo.Cells[$row,1].Value = $ct.ControlType
            $wsInfo.Cells[$row,2].Value = $ct.Label
            $wsInfo.Cells[$row,3].Value = $expectedVal
            $wsInfo.Cells[$row,4].Value = $bagText
            $wsInfo.Cells[$row,5].Value = $repText
            $wsInfo.Cells[$row,6].Value = $ct.ActualCount
            $wsInfo.Cells[$row,7].Value = $ct.MissingCount
            $wsInfo.Cells[$row,8].Value = if ($ct.Ok -eq $true) { '✔' } else { '⚠' }
            $row++
        }
    } else {
        $wsInfo.Cells[$row,1].Value = 'Ingen Compiling-analys tillgänglig'
        $wsInfo.Cells[$row,1,$row,$ctHeader.Length].Merge = $true
        $row++
    }

    Add-InfoBorder -FromRow $ctStart -FromCol 1 -ToRow ($row-1) -ToCol $ctHeader.Length
    $row += 1

    # D. Saknade replikat
    $includeMissing = Get-ReportOption -Name 'IncludeMissingReplicates' -Default $true
    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Saknade replikat' -Color $infoHeaderColor
    $row++
    if (-not $includeMissing) {
        $wsInfo.Cells[$row,1].Value = 'Sektionen ej aktiverad'
        $wsInfo.Cells[$row,1,$row,5].Merge = $true
        $row++
    } else {
        $wsInfo.Cells[$row,1].Value = 'Bag'
        $wsInfo.Cells[$row,2].Value = 'ControlType'
        $wsInfo.Cells[$row,3].Value = 'Replikat'
        $wsInfo.Cells[$row,1,$row,3].Style.Font.Bold = $true
        $wsInfo.Cells[$row,1,$row,3].Style.Fill.PatternType = 'Solid'
        $wsInfo.Cells[$row,1,$row,3].Style.Fill.BackgroundColor.SetColor($infoHeaderColor)
        $row++
        $missStart = $row - 1
        $maxMissingToShow = 10
        if ($compilingSummary.MissingReplicates.Count -gt 0) {
            $missItems = @($compilingSummary.MissingReplicates | Sort-Object Bag, ControlType, ReplicateNumber)
            $shown = $missItems | Select-Object -First $maxMissingToShow
            foreach ($m in $shown) {
                $wsInfo.Cells[$row,1].Value = $m.Bag
                $wsInfo.Cells[$row,2].Value = $m.ControlType
                $wsInfo.Cells[$row,3].Value = $m.ReplicateNumber
                $row++
            }
            $extra = $missItems.Count - $shown.Count
            if ($extra -gt 0) {
                $wsInfo.Cells[$row,1].Value = "(+$extra till)"
                $wsInfo.Cells[$row,1,$row,3].Merge = $true
                $row++
            }
        } else {
            $wsInfo.Cells[$row,1].Value = 'Inga saknade replikat'
            $wsInfo.Cells[$row,1,$row,3].Merge = $true
            $row++
        }
        Add-InfoBorder -FromRow $missStart -FromCol 1 -ToRow ($row-1) -ToCol 3
    }
    $row += 1

    # E. Ersättningar & Delaminations
    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Ersättningar & Delaminations' -Color $infoHeaderColor
    $row++
    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Ersättningar (A/AA/AAA)' -Color $infoHeaderColor
    $row++
    $wsInfo.Cells[$row,1].Value = 'Sample ID'
    $wsInfo.Cells[$row,2].Value = 'Typ'
    $wsInfo.Cells[$row,1,$row,2].Style.Font.Bold = $true
    $wsInfo.Cells[$row,1,$row,2].Style.Fill.PatternType = 'Solid'
    $wsInfo.Cells[$row,1,$row,2].Style.Fill.BackgroundColor.SetColor($infoHeaderColor)
    $row++
    $repStart = $row - 1
    if ($compilingSummary.Replacements.Count -gt 0) {
        foreach ($rpl in $compilingSummary.Replacements) {
            $wsInfo.Cells[$row,1].Value = $rpl.SampleId
            $wsInfo.Cells[$row,2].Value = $rpl.ReplacementType
            $row++
        }
    } else {
        $wsInfo.Cells[$row,1].Value = 'Inga ersättningar'
        $wsInfo.Cells[$row,1,$row,3].Merge = $true
        $row++
    }
    Add-InfoBorder -FromRow $repStart -FromCol 1 -ToRow ($row-1) -ToCol 3
    $row++

    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Delaminations (D)' -Color $infoHeaderColor
    $row++
    $wsInfo.Cells[$row,1].Value = 'Sample ID'
    $wsInfo.Cells[$row,1].Style.Font.Bold = $true
    $wsInfo.Cells[$row,1].Style.Fill.PatternType = 'Solid'
    $wsInfo.Cells[$row,1].Style.Fill.BackgroundColor.SetColor($infoHeaderColor)
    $row++
    $delamStart = $row - 1
    if ($compilingSummary.Delaminations.Count -gt 0) {
        foreach ($d in $compilingSummary.Delaminations) {
            $wsInfo.Cells[$row,1].Value = $d
            $row++
        }
    } else {
        $wsInfo.Cells[$row,1].Value = 'Inga delaminations'
        $wsInfo.Cells[$row,1,$row,3].Merge = $true
        $row++
    }
    Add-InfoBorder -FromRow $delamStart -FromCol 1 -ToRow ($row-1) -ToCol 3
    $row += 1

    # F. Dubbletter
    $includeDup = Get-ReportOption -Name 'IncludeDuplicates' -Default $true
    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Dublett Sample ID' -Color $infoHeaderColor
    $row++
    if (-not $includeDup) {
        $wsInfo.Cells[$row,1].Value = 'Sektionen ej aktiverad'
        $wsInfo.Cells[$row,1,$row,5].Merge = $true
        $row++
    } else {
        $wsInfo.Cells[$row,1].Value = 'Sample ID'
        $wsInfo.Cells[$row,2].Value = 'Antal'
        $wsInfo.Cells[$row,1,$row,2].Style.Font.Bold = $true
        $wsInfo.Cells[$row,1,$row,2].Style.Fill.PatternType = 'Solid'
        $wsInfo.Cells[$row,1,$row,2].Style.Fill.BackgroundColor.SetColor($infoHeaderColor)
        $row++
        $dupSStart = $row - 1
        if ($compilingSummary.DuplicateSampleIDs.Count -gt 0) {
            foreach ($ds in $compilingSummary.DuplicateSampleIDs) {
                $wsInfo.Cells[$row,1].Value = $ds.SampleId
                $wsInfo.Cells[$row,2].Value = $ds.Count
                $row++
            }
        } else {
            $wsInfo.Cells[$row,1].Value = 'Inga dubbletter'
            $wsInfo.Cells[$row,1,$row,3].Merge = $true
            $row++
        }
        Add-InfoBorder -FromRow $dupSStart -FromCol 1 -ToRow ($row-1) -ToCol 3
        $row++

        Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Dublett Cartridge S/N' -Color $infoHeaderColor
        $row++
        $wsInfo.Cells[$row,1].Value = 'Cartridge S/N'
        $wsInfo.Cells[$row,2].Value = 'Antal'
        $wsInfo.Cells[$row,1,$row,2].Style.Font.Bold = $true
        $wsInfo.Cells[$row,1,$row,2].Style.Fill.PatternType = 'Solid'
        $wsInfo.Cells[$row,1,$row,2].Style.Fill.BackgroundColor.SetColor($infoHeaderColor)
        $row++
        $dupCStart = $row - 1
        if ($compilingSummary.DuplicateCartridgeSN.Count -gt 0) {
            foreach ($dc in $compilingSummary.DuplicateCartridgeSN) {
                $wsInfo.Cells[$row,1].Value = $dc.Cartridge
                $wsInfo.Cells[$row,2].Value = $dc.Count
                $row++
            }
        } else {
            $wsInfo.Cells[$row,1].Value = 'Inga dubbletter'
            $wsInfo.Cells[$row,1,$row,3].Merge = $true
            $row++
        }
        Add-InfoBorder -FromRow $dupCStart -FromCol 1 -ToRow ($row-1) -ToCol 3
        $row += 1
    }

    # G. Instrumentfel
    $includeInstr = Get-ReportOption -Name 'IncludeInstrumentErrors' -Default $true
    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Instrumentfel' -Color $infoHeaderColor
    $row++
    if ($includeInstr) {
        Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Instrumentfel (sammanfattning)' -Color $infoHeaderColor
        $row++
        $wsInfo.Cells[$row,1].Value = 'Errorkod'
        $wsInfo.Cells[$row,2].Value = 'Antal'
        $wsInfo.Cells[$row,1,$row,2].Style.Font.Bold = $true
        $wsInfo.Cells[$row,1,$row,2].Style.Fill.PatternType = 'Solid'
        $wsInfo.Cells[$row,1,$row,2].Style.Fill.BackgroundColor.SetColor($infoHeaderColor)
        $row++
        $errSumStart = $row - 1
        if ($compilingSummary.ErrorCodes.Keys.Count -gt 0) {
            foreach ($code in ($compilingSummary.ErrorCodes.Keys | Sort-Object)) {
                $vals = $compilingSummary.ErrorCodes[$code]
                $wsInfo.Cells[$row,1].Value = $code
                $wsInfo.Cells[$row,2].Value = if ($vals) { $vals.Count } else { 0 }
                $row++
            }
        } else {
            $wsInfo.Cells[$row,1].Value = 'Inga instrumentfel'
            $wsInfo.Cells[$row,1,$row,3].Merge = $true
            $row++
        }
        Add-InfoBorder -FromRow $errSumStart -FromCol 1 -ToRow ($row-1) -ToCol 3
        $row++

        Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Instrumentfel (detaljer)' -Color $infoHeaderColor
        $row++
        $wsInfo.Cells[$row,1].Value = 'Errorkod'
        $wsInfo.Cells[$row,2].Value = 'Sample ID'
        $wsInfo.Cells[$row,1,$row,2].Style.Font.Bold = $true
        $wsInfo.Cells[$row,1,$row,2].Style.Fill.PatternType = 'Solid'
        $wsInfo.Cells[$row,1,$row,2].Style.Fill.BackgroundColor.SetColor($infoHeaderColor)
        $row++
        $errDetStart = $row - 1
        if ($compilingSummary.ErrorList.Count -gt 0) {
            $detailMax = 10
            $shownErr = $compilingSummary.ErrorList | Select-Object -First $detailMax
            foreach ($e in $shownErr) {
                $parts = $e -split ':',2
                $wsInfo.Cells[$row,1].Value = ($parts[0] + '').Trim()
                if ($parts.Count -gt 1) { $wsInfo.Cells[$row,2].Value = ($parts[1] + '').Trim() }
                $row++
            }
            $extra = $compilingSummary.ErrorList.Count - $shownErr.Count
            if ($extra -gt 0) {
                $wsInfo.Cells[$row,1].Value = "(+$extra till)"
                $wsInfo.Cells[$row,1,$row,3].Merge = $true
                $row++
            }
        } else {
            $wsInfo.Cells[$row,1].Value = 'Inga detaljer'
            $wsInfo.Cells[$row,1,$row,3].Merge = $true
            $row++
        }
        Add-InfoBorder -FromRow $errDetStart -FromCol 1 -ToRow ($row-1) -ToCol 3
        try { $wsInfo.Cells[$errDetStart,2,($row-1),2].Style.WrapText = $true } catch {}
        $row += 1
    } else {
        $wsInfo.Cells[$row,1].Value = 'Sektionen ej aktiverad'
        $wsInfo.Cells[$row,1,$row,5].Merge = $true
        $row++
    }

    # H. Dokumentinformation
    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Dokumentinformation (Worksheet / Seal Test)' -Color $infoHeaderColor
    $row++
    $wsInfo.Cells[$row,1].Value = ''
    $wsInfo.Cells[$row,2].Value = 'Worksheet'
    $wsInfo.Cells[$row,3].Value = 'Seal Test POS'
    $wsInfo.Cells[$row,4].Value = 'Seal Test NEG'
    $wsInfo.Cells[$row,1,$row,4].Style.Font.Bold = $true
    $wsInfo.Cells[$row,1,$row,4].Style.Fill.PatternType = 'Solid'
    $wsInfo.Cells[$row,1,$row,4].Style.Fill.BackgroundColor.SetColor($infoHeaderColor)
    $row++

    function Set-DocRow {
        param(
            [int]$RowIndex,
            [string]$Label,
            [string]$WsVal,
            [string]$PosVal,
            [string]$NegVal
        )
        $wsInfo.Cells[$RowIndex,1].Value = Convert-ToInfoValue $Label
        $wsInfo.Cells[$RowIndex,2].Value = if ($WsVal) { Convert-ToInfoValue $WsVal } else { 'N/A' }
        $wsInfo.Cells[$RowIndex,3].Value = if ($PosVal) { Convert-ToInfoValue $PosVal } else { 'N/A' }
        $wsInfo.Cells[$RowIndex,4].Value = if ($NegVal) { Convert-ToInfoValue $NegVal } else { 'N/A' }

        $mismatch = ($PosVal -and $NegVal -and ($PosVal -ne $NegVal))
        if ($mismatch) {
            $cells = $wsInfo.Cells[$RowIndex,3,$RowIndex,4]
            $cells.Style.Fill.PatternType = 'Solid'
            $cells.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightCoral)
        }
        $wsInfo.Cells[$RowIndex,1].Style.Font.Bold = $true
    }

    $wsFileName  = if ($selLsp) { Split-Path $selLsp -Leaf } else { '' }
    $posFileName = if ($selPos) { Split-Path $selPos -Leaf } else { '' }
    $negFileName = if ($selNeg) { Split-Path $selNeg -Leaf } else { '' }

    $wsDocClean = $null
    if ($headerWs) {
        $wsDocClean = $headerWs.DocumentNumber
        if ($wsDocClean) { $wsDocClean = ($wsDocClean -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$', '').Trim() }
        if ($headerWs.Attachment -and ($wsDocClean -notmatch '(?i)\bAttachment\s+\w+\b')) { $wsDocClean = "$wsDocClean Attachment $($headerWs.Attachment)" }
    }
    $docPosClean = $null
    if ($headerPos) {
        $docPosClean = $headerPos.DocumentNumber
        if ($docPosClean) { $docPosClean = ($docPosClean -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$','').Trim() }
    }
    $docNegClean = $null
    if ($headerNeg) {
        $docNegClean = $headerNeg.DocumentNumber
        if ($docNegClean) { $docNegClean = ($docNegClean -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$','').Trim() }
    }

    $docRows = @(
        @{Label='Filnamn';        Ws=$wsFileName;  Pos=$posFileName; Neg=$negFileName},
        @{Label='Dokumentnummer'; Ws=$wsDocClean;  Pos=$docPosClean; Neg=$docNegClean},
        @{Label='Revision';       Ws=if ($headerWs) { $headerWs.Rev } else { $null }; Pos=if ($headerPos) { $headerPos.Rev } else { $null }; Neg=if ($headerNeg) { $headerNeg.Rev } else { $null }},
        @{Label='Giltig fr.o.m.'; Ws=if ($headerWs) { $headerWs.Effective } else { $null }; Pos=if ($headerPos) { $headerPos.Effective } else { $null }; Neg=if ($headerNeg) { $headerNeg.Effective } else { $null }}
    )
    foreach ($dr in $docRows) {
        Set-DocRow -RowIndex $row -Label $dr.Label -WsVal $dr.Ws -PosVal $dr.Pos -NegVal $dr.Neg
        $row++
    }
    Add-InfoBorder -FromRow ($row - $docRows.Count - 1) -FromCol 1 -ToRow ($row-1) -ToCol 4
    $row += 1

    # I. Länkar
    Set-InfoHeaderRow -Row $row -FromCol 1 -ToCol 5 -Text 'Länkar' -Color $infoHeaderColor
    $row++
    $wsInfo.Cells[$row,1].Value = 'SharePoint Batch-länk'
    $wsInfo.Cells[$row,1].Style.Font.Bold = $true
    if ($batchInfo -and $batchInfo.Url) {
        $linkText = if ($batchInfo.LinkText) { $batchInfo.LinkText } else { 'LÄNK' }
        Add-Hyperlink -Cell $wsInfo.Cells[$row,2] -Text $linkText -Url $batchInfo.Url
    } else {
        $wsInfo.Cells[$row,2].Value = 'Ingen batchlänk tillgänglig'
    }
    $row++

    $linkMap = [ordered]@{
        'IPT App'      = 'https://apps.powerapps.com/play/e/default-771c9c47-7f24-44dc-958e-34f8713a8394/a/fd340dbd-bbbf-470b-b043-d2af4cb62c83'
        'MES Login'    = 'http://mes.cepheid.pri/camstarportal/?domain=CEPHEID.COM'
        'CSV Uploader' = 'http://auw2wgxtpap01.cepaws.com/Welcome.aspx'
        'BMRAM'        = 'https://cepheid62468.coolbluecloud.com/'
        'Agile'        = 'https://agileprod.cepheid.com/Agile/default/login-cms.jsp'
    }
    foreach ($key in $linkMap.Keys) {
        $wsInfo.Cells[$row,1].Value = $key
        Add-Hyperlink -Cell $wsInfo.Cells[$row,2] -Text 'LÄNK' -Url $linkMap[$key]
        $row++
    }

    try {
        $wsInfo.Cells[1,2,($row-1),5].Style.WrapText = $true
        $wsInfo.Cells[1,1,($row-1),5].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
    } catch {}

    return $csvStats
}

function Write-EquipmentSheet {
    param(
        [OfficeOpenXml.ExcelPackage]$OutputPkg,
        [pscustomobject]$InfoContext,
        [string]$UtrustningListPath,
        [scriptblock]$Logger = $null
    )

    $eqInfo = $InfoContext.EqInfo
    $selLsp = $InfoContext.SelLsp
    $headerWs = $InfoContext.HeaderWs

    if (-not $OutputPkg) { return }

    try {
        if (Test-Path -LiteralPath $UtrustningListPath) {
            $srcPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($UtrustningListPath))
            try {
                $srcWs = $srcPkg.Workbook.Worksheets['Sheet1']
                if (-not $srcWs) {
                    $srcWs = $srcPkg.Workbook.Worksheets[1]
                }

                if ($srcWs) {
                    $wsEq = $OutputPkg.Workbook.Worksheets['Infinity/GX']
                    if ($wsEq) {
                        $OutputPkg.Workbook.Worksheets.Delete($wsEq)
                    }
                    $wsEq = $OutputPkg.Workbook.Worksheets.Add('Infinity/GX', $srcWs)

                    if ($wsEq.Dimension) {
                        foreach ($cell in $wsEq.Cells[$wsEq.Dimension.Address]) {
                            if ($cell.Formula -or $cell.FormulaR1C1) {
                                $val = $cell.Value
                                $cell.Formula     = $null
                                $cell.FormulaR1C1 = $null
                                $cell.Value       = $val
                            }
                        }
                        $colCount = $srcWs.Dimension.End.Column
                        for ($c = 1; $c -le $colCount; $c++) {
                            try {
                                $wsEq.Column($c).Width = $srcWs.Column($c).Width
                            } catch {}
                        }
                    }

                    if ($eqInfo) {
                        $wsName = $null
                        if ($selLsp) {
                            $wsName = Split-Path $selLsp -Leaf
                        } elseif ($eqInfo.PSObject.Properties['WorksheetName'] -and $eqInfo.WorksheetName) {
                            $wsName = $eqInfo.WorksheetName
                        } elseif ($headerWs -and $headerWs.PSObject.Properties['WorksheetName'] -and $headerWs.WorksheetName) {
                            $wsName = $headerWs.WorksheetName
                        } else {
                            $wsName = 'Test Summary'
                        }

                        $cellHeaderPip = $wsEq.Cells['A14']
                        $cellHeaderPip.Value = "PIPETTER hämtade från $wsName"
                        $cellHeaderPip.Style.Font.Bold = $true
                        $cellHeaderPip.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                        $cellHeaderPip.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

                        $cellHeaderInst = $wsEq.Cells['A18']
                        $cellHeaderInst.Value = "INSTRUMENT hämtade från $wsName"
                        $cellHeaderInst.Style.Font.Bold = $true
                        $cellHeaderInst.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                        $cellHeaderInst.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

                        function Convert-ToEqDate {
                            param(
                                [Parameter(Mandatory = $false)]
                                $Value
                            )

                            if (-not $Value -or $Value -eq 'N/A') {
                                return $null
                            }

                            if ($Value -is [datetime]) {
                                return $Value
                            }

                            if ($Value -is [double] -or $Value -is [int]) {
                                try {
                                    $base = Get-Date '1899-12-30'
                                    return $base.AddDays([double]$Value)
                                } catch {
                                    return $Value
                                }
                            }

                            try {
                                return (Get-Date -Date $Value -ErrorAction Stop)
                            } catch {
                                return $Value
                            }
                        }

                        $pipetteIdCells  = @('B15','D15','F15','H15','J15','L15')
                        $pipetteDueCells = @('B16','D16','F16','H16','J16','L16')

                        $pipettes = @()
                        if ($eqInfo.PSObject.Properties['Pipettes'] -and $eqInfo.Pipettes) {
                            $pipettes = @($eqInfo.Pipettes)
                        }

                        for ($i = 0; $i -lt $pipetteIdCells.Count; $i++) {
                            $cellId  = $wsEq.Cells[$pipetteIdCells[$i]]
                            $cellDue = $wsEq.Cells[$pipetteDueCells[$i]]

                            if ($i -lt $pipettes.Count) {
                                $p = $pipettes[$i]

                                $id  = $null
                                $due = $null

                                if ($p -is [string]) {
                                    $id = $p
                                } else {
                                    $idCandidates = @()
                                    foreach ($propName in 'Id','CepheidId','Name','PipetteId') {
                                        if ($p.PSObject.Properties[$propName]) {
                                            $idCandidates += $p.$propName
                                        }
                                    }
                                    $id = $idCandidates | Where-Object { $_ } | Select-Object -First 1

                                    $dueCandidates = @()
                                    foreach ($propName in 'CalibrationDueDate','DueDate','CalDue') {
                                        if ($p.PSObject.Properties[$propName]) {
                                            $dueCandidates += $p.$propName
                                        }
                                    }
                                    $due = $dueCandidates | Where-Object { $_ } | Select-Object -First 1
                                }

                                if ([string]::IsNullOrWhiteSpace($id) -or $id -eq 'N/A') {
                                    $cellId.Value  = 'N/A'
                                    $cellDue.Value = 'N/A'
                                } else {
                                    $cellId.Value = $id

                                    $dt = Convert-ToEqDate -Value $due
                                    if ($dt -is [datetime]) {
                                        $cellDue.Value = $dt
                                        $cellDue.Style.Numberformat.Format = 'mmm-yy'
                                    } elseif ($dt) {
                                        $cellDue.Value = $dt
                                    } else {
                                        $cellDue.Value = 'N/A'
                                    }
                                }
                            } else {
                                $cellId.Value  = 'N/A'
                                $cellDue.Value = 'N/A'
                            }

                            foreach ($c in @($cellId,$cellDue)) {
                                $c.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                                $c.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
                            }
                        }

                        $instIdCells  = @(
                            'B19','D19','F19','H19','J19','L19',
                            'B21','D21','F21','H21','J21','L21'
                        )
                        $instDueCells = @(
                            'B20','D20','F20','H20','J20','L20',
                            'B22','D22','F22','H22','J22','L22'
                        )

                        $instruments = @()
                        if ($eqInfo.PSObject.Properties['Instruments'] -and $eqInfo.Instruments) {
                            $instruments = @($eqInfo.Instruments)
                        }

                        for ($i = 0; $i -lt $instIdCells.Count; $i++) {
                            $cellId  = $wsEq.Cells[$instIdCells[$i]]
                            $cellDue = $wsEq.Cells[$instDueCells[$i]]

                            if ($i -lt $instruments.Count) {
                                $inst = $instruments[$i]

                                $id  = $null
                                $due = $null

                                if ($inst -is [string]) {
                                    $id = $inst
                                } else {
                                    $idCandidates = @()
                                    foreach ($propName in 'Id','CepheidId','Name','InstrumentId') {
                                        if ($inst.PSObject.Properties[$propName]) {
                                            $idCandidates += $inst.$propName
                                        }
                                    }
                                    $id = $idCandidates | Where-Object { $_ } | Select-Object -First 1

                                    $dueCandidates = @()
                                    foreach ($propName in 'CalibrationDueDate','DueDate','CalDue') {
                                        if ($inst.PSObject.Properties[$propName]) {
                                            $dueCandidates += $inst.$propName
                                        }
                                    }
                                    $due = $dueCandidates | Where-Object { $_ } | Select-Object -First 1
                                }

                                if ([string]::IsNullOrWhiteSpace($id) -or $id -eq 'N/A') {
                                    $cellId.Value  = 'N/A'
                                    $cellDue.Value = 'N/A'
                                } else {
                                    $cellId.Value = $id

                                    $dt = Convert-ToEqDate -Value $due
                                    if ($dt -is [datetime]) {
                                        $cellDue.Value = $dt
                                        $cellDue.Style.Numberformat.Format = 'mmm-yy'
                                    } elseif ($dt) {
                                        $cellDue.Value = $dt
                                    } else {
                                        $cellDue.Value = 'N/A'
                                    }
                                }
                            } else {
                                $cellId.Value  = 'N/A'
                                $cellDue.Value = 'N/A'
                            }

                            foreach ($c in @($cellId,$cellDue)) {
                                $c.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                                $c.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
                            }
                        }

                        if ($pipettes.Count -gt $pipetteIdCells.Count -or
                            $instruments.Count -gt $instIdCells.Count) {
                            Log ("Infinity/GX: allt får inte plats i mallen (pipetter={0}, instrument={1})" -f $pipettes.Count, $instruments.Count) 'Info'
                        }
                    }
                } else {
                    Log "Utrustning saknas – Infinity/GX lämnas som mall." 'Info'
                }
            }
            finally {
                if ($srcPkg) { $srcPkg.Dispose() }
            }
        } else {
            Log "Infinity/GX mall saknas eller hittades inte." 'Info'
        }
    }
    catch {
        Log ("Kunde inte skapa 'Infinity/GX'-flik: {0}" -f $_.Exception.Message) 'Warn'
    }
}

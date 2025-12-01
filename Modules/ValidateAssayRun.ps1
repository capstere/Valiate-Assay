param()

function Invoke-ValidateAssayRun {
    param(
        [pscustomobject]$RunConfig,
        [scriptblock]$Logger = $null
    )

    $warnings = New-Object System.Collections.Generic.List[string]
    $errors   = New-Object System.Collections.Generic.List[string]

    function Log {
        param(
            [string]$Message,
            [string]$Severity = 'Info'
        )

        if ($Severity -in @('Warn','Warning')) { $warnings.Add($Message) }
        if ($Severity -eq 'Error') { $errors.Add($Message) }

        if ($Logger) {
            & $Logger -Message $Message -Severity $Severity
        } else {
            Write-BackendLog -Message $Message -Severity $Severity
        }
    }

    if (-not $RunConfig) {
        $errors.Add("Ingen körningskonfiguration angavs.")
        return New-Result -Ok $false -Errors $errors.ToArray() -Warnings $warnings.ToArray()
    }

    $selCsv = $RunConfig.CsvPath
    $selNeg = $RunConfig.SealNegPath
    $selPos = $RunConfig.SealPosPath
    $selLsp = $RunConfig.WorksheetPath
    $lsp    = $RunConfig.Lsp
    $templatePath = if ($RunConfig.TemplatePath) { $RunConfig.TemplatePath } else { Join-Path $RunConfig.ScriptRoot "output_template-v4.xlsx" }

    if (-not $selNeg -or -not $selPos) {
        $errors.Add("Du måste välja en Seal NEG och en Seal POS.")
        return New-Result -Ok $false -Errors $errors.ToArray() -Warnings $warnings.ToArray()
    }
    if (-not $lsp) {
        $warnings.Add("LSP saknas.")
        return New-Result -Ok $false -Errors $errors.ToArray() -Warnings $warnings.ToArray()
    }

    try {
        if (-not (Load-EPPlus)) {
            $errors.Add("EPPlus kunde inte laddas.")
            Log "EPPlus kunde inte laddas – avbryter." 'Error'
            return New-Result -Ok $false -Errors $errors.ToArray() -Warnings $warnings.ToArray()
        }
    } catch {
        $errors.Add("EPPlus-laddning misslyckades: $($_.Exception.Message)")
        Log ("EPPlus-laddning misslyckades: {0}" -f $_.Exception.Message) 'Error'
        return New-Result -Ok $false -Errors $errors.ToArray() -Warnings $warnings.ToArray()
    }

    $negWritable = $true; $posWritable = $true
    if ($RunConfig.WriteSignatures) {
        $negWritable = -not (Test-FileLocked $selNeg); if (-not $negWritable) { Log "NEG är låst (öppen i Excel?)." 'Warn' }
        $posWritable = -not (Test-FileLocked $selPos); if (-not $posWritable) { Log "POS är låst (öppen i Excel?)." 'Warn' }
    }

    $pkgNeg = $null; $pkgPos = $null; $pkgOut = $null; $tmpPkg = $null
    $auditData = [pscustomobject]@{}

    try {
        try {
            $pkgNeg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selNeg))
            $pkgPos = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selPos))
        } catch {
            $errors.Add("Kunde inte öppna NEG/POS: $($_.Exception.Message)")
            Log ("Kunde inte öppna NEG/POS: {0}" -f $_.Exception.Message) 'Error'
            return New-Result -Ok $false -Errors $errors.ToArray() -Warnings $warnings.ToArray()
        }

        if (-not (Test-Path -LiteralPath $templatePath)) {
            $errors.Add("Mallfilen 'output_template-v4.xlsx' saknas!")
            Log "Mallfilen 'output_template-v4.xlsx' saknas!" 'Error'
            return New-Result -Ok $false -Errors $errors.ToArray() -Warnings $warnings.ToArray()
        }
        try {
            $pkgOut = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($templatePath))
        } catch {
            $errors.Add("Kunde inte läsa mall: $($_.Exception.Message)")
            Log ("Kunde inte läsa mall: {0}" -f $_.Exception.Message) 'Error'
            return New-Result -Ok $false -Errors $errors.ToArray() -Warnings $warnings.ToArray()
        }

        $signToWrite = ($RunConfig.SignatureText + '').Trim()
        if ($RunConfig.WriteSignatures) {
            if (-not $signToWrite) {
                $errors.Add("Ingen signatur angiven (B47).")
                Log "Ingen signatur angiven (B47)." 'Error'
                return New-Result -Ok $false -Errors $errors.ToArray() -Warnings $warnings.ToArray()
            }
            $sigResult = Write-SealTestSignatures -PkgNeg $pkgNeg -PkgPos $pkgPos -Signature $signToWrite -Overwrite $RunConfig.OverwriteSignature -NegWritable $negWritable -PosWritable $posWritable -Logger $Logger
        }

        $csvRows = @(); $runAssay = $null
        if ($selCsv) {
            try { $csvRows = Import-CsvRows -Path $selCsv -StartRow 10 } catch {}
            try { $runAssay = Get-AssayFromCsv -Path $selCsv -StartRow 10 } catch {}
            if ($runAssay) { Log ("Assay från CSV: {0}" -f $runAssay) 'Info' }
        }
        $controlTab = $null
        if ($runAssay) { $controlTab = Get-ControlTabName -AssayName $runAssay }
        if ($controlTab) { Log ("Control Material-flik: {0}" -f $controlTab) 'Info' } else { Log "Ingen control-mappning (fortsätter utan)." 'Info' }

        $sealAnalysis = Get-SealTestViolations -PkgNeg $pkgNeg -PkgPos $pkgPos

        $negSigSet = Get-SignatureSetForDataSheets -Pkg $pkgNeg
        $posSigSet = Get-SignatureSetForDataSheets -Pkg $pkgPos

        if (-not $selLsp) {
            $probeDir = $null
            if ($selPos) { $probeDir = Split-Path -Parent $selPos }
            if (-not $probeDir -and $selNeg) { $probeDir = Split-Path -Parent $selNeg }
            if ($probeDir -and (Test-Path -LiteralPath $probeDir)) {
                $cand = Get-ChildItem -LiteralPath $probeDir -File -ErrorAction SilentlyContinue |
                        Where-Object {
                            ($_.Name -match '(?i)worksheet') -and ($_.Extension -match '^\.(xlsx|xlsm|xls)$')
                        } |
                        Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($cand) { $selLsp = $cand.FullName }
            }
        }

        $headerWs  = $null
        $headerNeg = $null
        $headerPos = $null
        $wsHeaderCheck = $null
        $eqInfo = $null
        $tsControls = $null
        $worksheetControlMaterials = @()

        try {
            if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
                $tmpPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selLsp))
                try {
                    $eqInfo = Get-TestSummaryEquipment -Pkg $tmpPkg
                    if ($eqInfo) {
                        Log ("Utrustning hittad i WS '{0}': Pipetter={1}, Instrument={2}" -f $eqInfo.WorksheetName, ($eqInfo.Pipettes.Count), ($eqInfo.Instruments.Count)) 'Info'
                    }
                    try {
                        $tsControls = Get-TestSummaryControls -Pkg $tmpPkg
                        if ($tsControls -and $tsControls.Controls) {
                            $worksheetControlMaterials = @($tsControls.Controls)
                            Log ("Kontrollmaterial hittade i WS '{0}': {1}" -f $tsControls.WorksheetName, ($tsControls.Controls.Count)) 'Info'
                        }
                    } catch {
                        Log ("Kunde inte extrahera kontrollmaterial från Test Summary: " + $_.Exception.Message) 'Warn'
                    }
                } catch {
                    Log ("Kunde inte extrahera utrustning från Test Summary: " + $_.Exception.Message) 'Warn'
                }

                $headerWs = Extract-WorksheetHeader -Pkg $tmpPkg
                $wsHeaderRows  = Get-WorksheetHeaderPerSheet -Pkg $tmpPkg
                $wsHeaderCheck = Compare-WorksheetHeaderSet   -Rows $wsHeaderRows
                try {
                    if ($wsHeaderCheck.Issues -gt 0 -and $wsHeaderCheck.Summary) {
                        Log ("Worksheet header-avvikelser: {0} – se Information!" -f $wsHeaderCheck.Summary) 'Warn'
                    } else {
                        Log "Worksheet header korrekt" 'Info'
                    }
                } catch {}

                function Find-LabelValueRightward {
                    param(
                        [OfficeOpenXml.ExcelWorksheet]$Ws,
                        [string]$Label,
                        [int]$MaxRows = 1000,
                        [int]$MaxCols = 200
                    )
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
                    $normLbl = Normalize-HeaderText $Label
                    $pat = '^(?i)\s*' + [regex]::Escape($normLbl).Replace('\ ', '\s*') + '\s*[:\.,-]*\s*$'
                    $rx  = [regex]::new($pat, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    $hit = Find-RegexCell -Ws $Ws -Rx $rx -MaxRows $MaxRows -MaxCols $MaxCols
                    if (-not $hit) { return $null }
                    $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
                    for ($c = $hit.Col + 1; $c -le $cMax; $c++) {
                        $t = Normalize-HeaderText ($Ws.Cells[$hit.Row,$c].Text + '')
                        if ($t) { return $t }
                    }
                    return $null
                }

                $wsLsp   = $tmpPkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                if ($wsLsp) {
                    if (-not $headerWs -or -not $headerWs.PartNo) {
                        $val = $null
                        foreach ($lbl in @('Part No.','Part No.:','Part No','Part Number','Part Number:','Part Number.','Part Number.:')) {
                            $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                            if ($val) { break }
                        }
                        if ($val) { $headerWs.PartNo = $val }
                    }
                    if (-not $headerWs -or -not $headerWs.BatchNo) {
                        $val = $null
                        foreach ($lbl in @('Batch No(s)','Batch No(s).','Batch No(s):','Batch No(s).:','Batch No','Batch No.','Batch No:','Batch No.:','Batch Number','Batch Number.','Batch Number:','Batch Number.:')) {
                            $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                            if ($val) { break }
                        }
                        if ($val) { $headerWs.BatchNo = $val }
                    }
                    if (-not $headerWs -or -not $headerWs.CartridgeNo -or $headerWs.CartridgeNo -eq '.') {
                        $val = $null
                        foreach ($lbl in @('Cartridge No. (LSP)','Cartridge No. (LSP):','Cartridge No. (LSP) :','Cartridge No (LSP)','Cartridge No (LSP):','Cartridge No (LSP) :','Cartridge Number (LSP)','Cartridge Number (LSP):','Cartridge Number (LSP) :','Cartridge No.','Cartridge No.:','Cartridge No. :','Cartridge No :','Cartridge Number','Cartridge Number:','Cartridge Number :','Cartridge No','Cartridge No:','Cartridge No :')) {
                            $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                            if ($val) { break }
                        }
                        if (-not $val) {
                            $rxCart = [regex]::new('(?i)Cartridge.*\(LSP\)')
                            $hitCart = Find-RegexCell -Ws $wsLsp -Rx $rxCart -MaxRows 200 -MaxCols ([Math]::Min($wsLsp.Dimension.End.Column, 100))
                            if ($hitCart) {
                                for ($c = $hitCart.Col + 1; $c -le $wsLsp.Dimension.End.Column; $c++) {
                                    $cellVal = ($wsLsp.Cells[$hitCart.Row, $c].Text + '').Trim()
                                    if ($cellVal) { $val = $cellVal; break }
                                }
                            }
                        }
                        if ($val) { $headerWs.CartridgeNo = $val }
                    }
                    if (-not $headerWs -or -not $headerWs.Effective) {
                        $val = Find-LabelValueRightward -Ws $wsLsp -Label 'Effective'
                        if (-not $val) { $val = Find-LabelValueRightward -Ws $wsLsp -Label 'Effective Date' }
                        if ($val) { $headerWs.Effective = $val }
                    }
                }
                if ($selLsp -and (-not $headerWs -or -not $headerWs.CartridgeNo -or $headerWs.CartridgeNo -eq '.' -or $headerWs.CartridgeNo -eq '')) {
                    $fn = Split-Path $selLsp -Leaf
                    $m = [regex]::Matches($fn, '(?<!\d)(\d{5,7})(?!\d)')
                    if ($m.Count -gt 0) { $headerWs.CartridgeNo = $m[0].Groups[1].Value }
                }
            } else {
                Log "Ingen WS-fil vald/hittad (LSP Worksheet). Hoppar över WS-extraktion." 'Info'
            }
        } catch {
            Log ("WS-block fel: " + $_.Exception.Message) 'Warn'
        } finally {
            if ($tmpPkg) { $tmpPkg.Dispose() }
        }

        try { $headerNeg = Extract-SealTestHeader -Pkg $pkgNeg } catch {}
        try { $headerPos = Extract-SealTestHeader -Pkg $pkgPos } catch {}

        try {
            if ($pkgPos -and -not $headerPos.Effective) {
                $wsPos = $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                if ($wsPos) {
                    $val = Find-LabelValueRightward -Ws $wsPos -Label 'Effective'
                    if (-not $val) { $val = Find-LabelValueRightward -Ws $wsPos -Label 'Effective Date' }
                    if ($val) { $headerPos.Effective = $val }
                }
            }
        } catch {}
        try {
            if ($pkgNeg -and -not $headerNeg.Effective) {
                $wsNeg = $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                if ($wsNeg) {
                    $val = Find-LabelValueRightward -Ws $wsNeg -Label 'Effective'
                    if (-not $val) { $val = Find-LabelValueRightward -Ws $wsNeg -Label 'Effective Date' }
                    if ($val) { $headerNeg.Effective = $val }
                }
            }
        } catch {}

        $batchInfo = Get-BatchLinkInfo -SealPosPath $selPos -SealNegPath $selNeg -Lsp $lsp
        $batch = $batchInfo.Batch

        $compilingData = @()
        $compilingSummary = $null
        try {
            if ($selCsv -and (Test-Path -LiteralPath $selCsv)) {
                try { $compilingData = Convert-XpertSummaryCsvToData -Path $selCsv } catch { Log ("Convert-XpertSummaryCsvToData: " + $_.Exception.Message) 'Warn' }
                try { $compilingSummary = Get-CompilingAnalysis -Data $compilingData } catch { Log ("Get-CompilingAnalysis: " + $_.Exception.Message) 'Warn' }
            }
        } catch {}
        if (-not $compilingSummary) { $compilingSummary = Get-CompilingAnalysis -Data @() }

        $assayName    = if ($compilingSummary.AssayName) { $compilingSummary.AssayName } elseif ($runAssay) { $runAssay } else { '' }
        $assayVersion = $compilingSummary.AssayVersion
        $lspSummary   = if ($compilingSummary.LspSummary) { $compilingSummary.LspSummary } else { '' }
        $testCount    = $compilingSummary.TestCount

        $miniVal = $RunConfig.MiniTabMacro
        if (-not $miniVal) {
            $assayForMacro = ''
            if ($runAssay) { $assayForMacro = $runAssay }
            elseif ($assayName) { $assayForMacro = $assayName }
            elseif ($headerWs -and $headerWs.CartridgeNo) { $assayForMacro = $headerWs.CartridgeNo }
            if (-not $assayForMacro) { $assayForMacro = $assayName }
            if (Get-Command Get-MinitabMacro -ErrorAction SilentlyContinue) {
                $miniVal = Get-MinitabMacro -AssayName $assayForMacro
            }
            if (-not $miniVal) { $miniVal = 'N/A' }
        }

        $controlMap = $global:ControlMaterialData
        if (-not $controlMap -and $global:ControlMaterialMapPath -and (Test-Path -LiteralPath $global:ControlMaterialMapPath)) {
            try { $controlMap = Get-ControlMaterialMap } catch {}
        }

        if (-not $tsControls -and $worksheetControlMaterials) {
            $tsControls = [pscustomobject]@{ Controls = $worksheetControlMaterials }
        } elseif ($tsControls -and $tsControls.Controls) {
            $worksheetControlMaterials = @($tsControls.Controls)
        }

        $reportOptions = if ($RunConfig.ReportOptions) { $RunConfig.ReportOptions } elseif ($global:ReportOptions) { $global:ReportOptions } else { @{} }

        $infoContext = [pscustomobject]@{
            CsvPath                  = $selCsv
            CompilingSummary         = $compilingSummary
            CompilingData            = $compilingData
            RunAssay                 = $runAssay
            AssayName                = $assayName
            AssayVersion             = $assayVersion
            Lsp                      = $lsp
            LspSummary               = $lspSummary
            TestCount                = $testCount
            MiniVal                  = $miniVal
            SelLsp                   = $selLsp
            SelPos                   = $selPos
            SelNeg                   = $selNeg
            HeaderWs                 = $headerWs
            HeaderPos                = $headerPos
            HeaderNeg                = $headerNeg
            WsHeaderCheck            = $wsHeaderCheck
            EqInfo                   = $eqInfo
            TsControls               = if ($tsControls -and $tsControls.Controls) { @($tsControls.Controls) } else { $null }
            WorksheetControlMaterials= $worksheetControlMaterials
            ControlMap               = $controlMap
            BatchInfo                = $batchInfo
            Batch                    = $batch
            ReportOptions            = $reportOptions
            ScriptVersion            = $RunConfig.ScriptVersion
        }

        $signatureInfo = [pscustomobject]@{
            NegSigSet       = $negSigSet
            PosSigSet       = $posSigSet
            SignatureToWrite= $signToWrite
        }

        $buildResult = Build-AssayReport -OutputPkg $pkgOut -PkgNeg $pkgNeg -PkgPos $pkgPos -RunConfig $RunConfig -SealAnalysis $sealAnalysis -SignatureInfo $signatureInfo -InfoContext $infoContext -ControlTab $controlTab -RawDataPath $RunConfig.RawDataPath -UtrustningListPath $RunConfig.UtrustningListPath -Logger $Logger

        $nowTs   = Get-Date -Format "yyyyMMdd_HHmmss"
        $baseName = "$($env:USERNAME)_output_${lsp}_$nowTs.xlsx"
        if ($RunConfig.SaveInLsp) {
            $saveDir = Split-Path -Parent $selNeg
            $SavePath = Join-Path $saveDir $baseName
            Log ("Sparläge: LSP-mapp → {0}" -f $saveDir) 'Info'
        } else {
            $saveDir = $env:TEMP
            $SavePath = Join-Path $saveDir $baseName
            Log ("Sparläge: Temporärt → {0}" -f $SavePath) 'Info'
        }

        try {
            $pkgOut.Workbook.View.ActiveTab = 0
            $wsInitial = $pkgOut.Workbook.Worksheets["Information"]
            if ($wsInitial) { $wsInitial.View.TabSelected = $true }
            $pkgOut.SaveAs($SavePath)
            Log ("Rapport sparad: {0}" -f $SavePath) 'Info'
            $global:LastReportPath = $SavePath
        } catch {
            $errors.Add("Kunde inte spara/öppna: $($_.Exception.Message)")
            Log ("Kunde inte spara/öppna: {0}" -f $_.Exception.Message) 'Warn'
            return New-Result -Ok $false -Errors $errors.ToArray() -Warnings $warnings.ToArray()
        }

        try {
            $auditDir = Join-Path $RunConfig.ScriptRoot 'audit'
            if (-not (Test-Path $auditDir)) { New-Item -ItemType Directory -Path $auditDir -Force | Out-Null }
            $auditObj = [pscustomobject]@{
                DatumTid        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                Användare       = $env:USERNAME
                LSP             = $lsp
                ValdCSV         = if ($selCsv) { Split-Path $selCsv -Leaf } else { '' }
                ValdSealNEG     = Split-Path $selNeg -Leaf
                ValdSealPOS     = Split-Path $selPos -Leaf
                SignaturSkriven = if ($RunConfig.WriteSignatures) { 'Ja' } else { 'Nej' }
                OverwroteSign   = if ($RunConfig.OverwriteSignature) { 'Ja' } else { 'Nej' }
                SigMismatch     = if ($buildResult.SignatureSummary -and $buildResult.SignatureSummary.SigMismatch) { 'Ja' } else { 'Nej' }
                MismatchSheets  = if ($buildResult.SignatureSummary -and $buildResult.SignatureSummary.MismatchSheets -and $buildResult.SignatureSummary.MismatchSheets.Count -gt 0) { ($buildResult.SignatureSummary.MismatchSheets -join ';') } else { '' }
                ViolationsNEG   = $sealAnalysis.ViolationsNeg.Count
                ViolationsPOS   = $sealAnalysis.ViolationsPos.Count
                Violations      = ($sealAnalysis.ViolationsNeg.Count + $sealAnalysis.ViolationsPos.Count)
                Sparläge        = if ($RunConfig.SaveInLsp) { 'LSP-mapp' } else { 'Temporärt' }
                OutputFile      = $SavePath
                Kommentar       = 'UNCONTROLLED rapport, ingen källfil ändrades automatiskt.'
                ScriptVersion   = $RunConfig.ScriptVersion
            }

            $auditFile = Join-Path $auditDir ("$($env:USERNAME)_audit_${nowTs}.csv")
            $auditObj | Export-Csv -Path $auditFile -NoTypeInformation -Encoding UTF8
            try {
                $statusText = 'OK'
                if (($sealAnalysis.ViolationsNeg.Count + $sealAnalysis.ViolationsPos.Count) -gt 0 -or ($buildResult.SignatureSummary -and $buildResult.SignatureSummary.SigMismatch)) {
                    $statusText = 'Warnings'
                }
                $auditTests = $null
                try { if ($buildResult.CsvStats) { $auditTests = $buildResult.CsvStats.TestCount } } catch {}
                Add-AuditEntry -Lsp $lsp -Assay $runAssay -BatchNumber $batch -TestCount $auditTests -Status $statusText -ReportPath $SavePath
            } catch { Log ("Kunde inte skriva audit-CSV: {0}" -f $_.Exception.Message) 'Warn' }
        } catch {
            Log ("Kunde inte skriva revisionsfil: {0}" -f $_.Exception.Message) 'Warn'
        }

        return New-Result -Ok $true -Data ([pscustomobject]@{
            ReportPath        = $SavePath
            SealAnalysis      = $sealAnalysis
            CompilingSummary  = $compilingSummary
            SignatureSummary  = $buildResult.SignatureSummary
            CsvStats          = $buildResult.CsvStats
            BatchInfo         = $batchInfo
        }) -Warnings $warnings.ToArray() -Errors $errors.ToArray()
    } finally {
        try { if ($pkgNeg) { $pkgNeg.Dispose() } } catch {}
        try { if ($pkgPos) { $pkgPos.Dispose() } } catch {}
        try { if ($pkgOut) { $pkgOut.Dispose() } } catch {}
        try { if ($tmpPkg) { $tmpPkg.Dispose() } } catch {}
    }
}

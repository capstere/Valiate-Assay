<#
    Module: UsedEquipment

    This module encapsulates functions for extracting equipment and control
    information from various input sources (CSV, worksheet, seal files).  The
    intent is to separate the parsing logic from the main workflow and make
    the code easier to unit test.  The functions exposed here should not
    perform any validation beyond basic sanity checks; all assay‑specific
    validation belongs in the AssayValidation module.

    Implementations are intentionally lightweight to satisfy the
    demonstration requirements.  In a production environment you should
    handle exceptions, character encodings, and data type conversions with
    greater care.
#>

function Get-UsedEquipmentFromCsv {
    <#
        .SYNOPSIS
            Extracts a list of instruments and modules used in a CSV file.

        .PARAMETER CsvRows
            An array of rows imported via Import-Csv.

        .PARAMETER ColumnMap
            A mapping of canonical column names to actual property names.

        .OUTPUTS
            An array of PSCustomObject with Instrument and Module serial numbers.

        This function assumes that the CSV contains columns named
        'Instrument S/N' and 'Module S/N'.  If additional columns are used
        in the future, update the canonical names accordingly.
    #>
    param(
        [array]$CsvRows,
        [hashtable]$ColumnMap
    )
    $equipment = @()
    foreach ($row in $CsvRows) {
        $inst = $row.$($ColumnMap['instrument s/n'])
        $mod  = $row.$($ColumnMap['module s/n'])
        if ($inst -and $mod) {
            $equipment += [PSCustomObject]@{ InstrumentSN = $inst; ModuleSN = $mod }
        }
    }
    return $equipment | Sort-Object -Unique
}

function Get-UsedPipettesFromWorksheet {
    <#
        .SYNOPSIS
            Extracts pipette identifiers and calibration dates from a worksheet.

        .DESCRIPTION
            This function opens the provided Excel worksheet and searches for a row
            containing the text "Cepheid ID".  It then reads the pipette IDs in
            the cells to the right of this header and the calibration dates in
            the row below.  Dates are normalized to MMM-yy format.  If the
            worksheet is missing or does not contain pipette information, an
            empty result is returned.

        .PARAMETER WorksheetPath
            Full path to the worksheet Excel file (.xlsx).  If null or the
            file does not exist, an empty result is returned.

        .OUTPUTS
            A PSCustomObject with two properties: 'Pipettes Used:' and
            'Calibration Due Date:'.  Each is an array of strings.
    #>
    param([string]$WorksheetPath)
    if ([string]::IsNullOrEmpty($WorksheetPath) -or -not (Test-Path -LiteralPath $WorksheetPath)) {
        return [PSCustomObject]@{ 'Pipettes Used:' = @(); 'Calibration Due Date:' = @() }
    }
    $used = @(); $cal = @()
    try {
        $package = New-Object OfficeOpenXml.ExcelPackage -ArgumentList (New-Object System.IO.FileInfo($WorksheetPath))
        $ws = $package.Workbook.Worksheets['Test Summary']
        if ($null -eq $ws) { throw "Worksheet 'Test Summary' not found" }
        $rowCount = $ws.Dimension.End.Row
        $colCount = $ws.Dimension.End.Column
        $found = $false
        for ($r = 1; $r -le $rowCount -and -not $found; $r++) {
            for ($c = 1; $c -le $colCount; $c++) {
                $val = $ws.Cells[$r,$c].Text
                if ($val -match 'Cepheid ID') {
                    $found = $true
                    $headerRow = $r
                    $startCol = $c + 1
                    for ($cc = $startCol; $cc -le $colCount; $cc++) {
                        $pipette = $ws.Cells[$headerRow,$cc].Text
                        $calVal  = $ws.Cells[$headerRow + 1,$cc].Text
                        if ($pipette -match '^Nr\.\s*\d+') {
                            $used += $pipette
                            $parsed = $null
                            $calText = if ([datetime]::TryParse($calVal,[ref]$parsed)) {
                                $parsed.ToString('MMM-yy',[System.Globalization.CultureInfo]::InvariantCulture)
                            } else { $calVal }
                            $cal += $calText
                        }
                    }
                    break
                }
            }
        }
        $package.Dispose()
    } catch {
        Write-Warning "Failed to read pipettes from worksheet '$WorksheetPath': $($_.Exception.Message)"
        return [PSCustomObject]@{ 'Pipettes Used:' = @(); 'Calibration Due Date:' = @() }
    }
    return [PSCustomObject]@{
        'Pipettes Used:'        = ($used  | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
        'Calibration Due Date:' = ($cal   | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    }
}

function Get-SealTestData {
    <#
        .SYNOPSIS
            Extracts header fields, tester information and weight loss violations from seal test workbooks.

        .DESCRIPTION
            This function opens both positive (POS) and negative (NEG) seal test workbooks and
            extracts a set of predefined header fields (ROBAL, part number, batch number, etc.),
            a list of testers (names in cell B43 of each sheet) and any weight loss violations
            (K column ≤ −2.4 or explicit FAIL in column L).  It also reports mismatches in
            certain header fields between POS and NEG (e.g. batch number).  If no files are
            provided, empty collections are returned.

        .PARAMETER PosPath
            Full path to the seal POS workbook.  Optional; can be null or empty.

        .PARAMETER NegPath
            Full path to the seal NEG workbook.  Optional; can be null or empty.

        .OUTPUTS
            A PSCustomObject with HeaderFields, Testers, Violations and Mismatches.
    #>
    param(
        [string]$PosPath,
        [string]$NegPath
    )
    # Return empty structure if both files are missing
    if (([string]::IsNullOrEmpty($PosPath) -or -not (Test-Path -LiteralPath $PosPath)) -and
        ([string]::IsNullOrEmpty($NegPath) -or -not (Test-Path -LiteralPath $NegPath))) {
        return [PSCustomObject]@{ HeaderFields = @{}; Testers = @{ POS=@(); NEG=@() }; Violations=@(); Mismatches=@() }
    }
    $pkgNeg = $null; $pkgPos = $null
    $headerFields = @{}; $testers = @{ POS=@(); NEG=@() }; $violations = @(); $mismatches = @()
    # Predefine fields to extract
    $fields = @(
        @{ Label='ROBAL';             Cell='F2'  },
        @{ Label='Part Number';       Cell='B2'  },
        @{ Label='Batch Number';      Cell='D2'  },
        @{ Label='Cartridge Number';  Cell='B6'  },
        @{ Label='PO Number';         Cell='B10' },
        @{ Label='Assay Family';      Cell='D10' },
        @{ Label='Weight Loss Spec';  Cell='F10' },
        @{ Label='Balance ID Number'; Cell='B14' },
        @{ Label='Balance Cal Due';   Cell='D14' },
        @{ Label='Vacuum Oven ID';    Cell='B20' },
        @{ Label='Vacuum Oven Cal';   Cell='D20' },
        @{ Label='Timer ID Number';   Cell='B25' },
        @{ Label='Timer Cal Due';     Cell='D25' }
    )
    try {
        if ($NegPath -and (Test-Path -LiteralPath $NegPath)) {
            $pkgNeg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($NegPath))
        }
        if ($PosPath -and (Test-Path -LiteralPath $PosPath)) {
            $pkgPos = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($PosPath))
        }
        foreach ($f in $fields) {
            $valNeg = ''; $valPos = ''
            if ($pkgNeg) {
                foreach ($ws in $pkgNeg.Workbook.Worksheets) {
                    if ($ws.Name -ne 'Worksheet Instructions') {
                        $v = $ws.Cells[$f.Cell].Text
                        if ($v) { $valNeg = "$v"; break }
                    }
                }
            }
            if ($pkgPos) {
                foreach ($ws in $pkgPos.Workbook.Worksheets) {
                    if ($ws.Name -ne 'Worksheet Instructions') {
                        $v = $ws.Cells[$f.Cell].Text
                        if ($v) { $valPos = "$v"; break }
                    }
                }
            }
            $headerFields[$f.Label] = @{ NEG=$valNeg; POS=$valPos }
            # detect mismatch for key fields
            if ($valNeg -ne $valPos -and $f.Label -in @('ROBAL','Part Number','Batch Number','Cartridge Number','PO Number','Assay Family','Weight Loss Spec')) {
                $mismatches += [PSCustomObject]@{ Field=$f.Label; NEG=$valNeg; POS=$valPos }
            }
        }
        # Gather testers from cell B43 on each sheet
        if ($pkgNeg) {
            foreach ($ws in $pkgNeg.Workbook.Worksheets) {
                if ($ws.Name -ne 'Worksheet Instructions') {
                    $txt = $ws.Cells['B43'].Text
                    if ($txt) { $testers.NEG += ($txt -split ',') }
                }
            }
        }
        if ($pkgPos) {
            foreach ($ws in $pkgPos.Workbook.Worksheets) {
                if ($ws.Name -ne 'Worksheet Instructions') {
                    $txt = $ws.Cells['B43'].Text
                    if ($txt) { $testers.POS += ($txt -split ',') }
                }
            }
        }
        $testers.NEG = $testers.NEG | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
        $testers.POS = $testers.POS | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
        # Weight loss violations
        $sheets = @()
        if ($pkgNeg) { $sheets += $pkgNeg.Workbook.Worksheets }
        if ($pkgPos) { $sheets += $pkgPos.Workbook.Worksheets }
        foreach ($ws in $sheets) {
            if ($ws.Name -eq 'Worksheet Instructions') { continue }
            $isNEG = $pkgNeg -and ($pkgNeg.Workbook.Worksheets -contains $ws)
            for ($r = 3; $r -le 45; $r++) {
                $valK = $ws.Cells["K$r"].Value
                $textL = $ws.Cells["L$r"].Text
                if ($valK -ne $null -and $valK -is [double]) {
                    if ($textL -eq 'FAIL' -or $valK -le -2.4) {
                        $violations += [PSCustomObject]@{
                            Type       = ($isNEG ? 'NEG' : 'POS')
                            Sheet      = $ws.Name
                            Cartridge  = $ws.Cells["H$r"].Text
                            InitialW   = $ws.Cells["I$r"].Value
                            FinalW     = $ws.Cells["J$r"].Value
                            WeightLoss = $valK
                            Status     = ($textL -eq 'FAIL' ? 'FAIL' : 'Minus')
                            Obs        = $ws.Cells["M$r"].Text
                        }
                    }
                }
            }
        }
    } catch {
        Write-Warning "Failed to parse seal files: $($_.Exception.Message)"
    } finally {
        if ($pkgNeg) { $pkgNeg.Dispose() }
        if ($pkgPos) { $pkgPos.Dispose() }
    }
    return [PSCustomObject]@{
        HeaderFields = $headerFields
        Testers      = $testers
        Violations   = $violations
        Mismatches   = $mismatches
    }
}

function Get-ControlInventory {
    <#
        .SYNOPSIS
            Reads a control inventory workbook and returns records matching an assay.

        .DESCRIPTION
            The control inventory contains product/lot/expiry details of control materials.
            This function scans columns 7–13 for text matching either a given assay name
            or a slang mapping (e.g. CTNG ➜ Xpert CT_NG).  It outputs rows where a match
            occurs, along with product details, quantities and storage location.

        .PARAMETER RawDataPath
            Full path to the control inventory Excel workbook (.xlsx).

        .PARAMETER AssayName
            Name of the assay to match.  Use '*' to match all.

        .PARAMETER SlangMap
            Optional mapping of slang names to assay names.

        .OUTPUTS
            An array of PSCustomObjects with control inventory details.
    #>
    param(
        [string]$RawDataPath,
        [string]$AssayName,
        [hashtable]$SlangMap
    )
    $records = @()
    if ([string]::IsNullOrEmpty($RawDataPath) -or -not (Test-Path -LiteralPath $RawDataPath)) {
        return $records
    }
    try {
        $fi = [IO.FileInfo]::new($RawDataPath)
        $package = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $fi
        $ws = $package.Workbook.Worksheets[1]
        $rowCount = $ws.Dimension.End.Row
        $assaySlang = $null
        if ($SlangMap -and $AssayName -and $AssayName -ne '*') {
            $assaySlang = $SlangMap.Keys | Where-Object { $SlangMap[$_] -eq $AssayName } | Select-Object -First 1
        }
        for ($r = 2; $r -le $rowCount; $r++) {
            $products = @(); $isMatch = $false
            for ($c = 7; $c -le 13; $c++) {
                $cell = $ws.Cells[$r,$c].Text
                if ($cell -and $cell -ne '' -and $cell.ToUpper() -ne 'N/A') {
                    if (($assaySlang -and $cell -like "*$assaySlang*") -or ($AssayName -ne '*' -and $cell -like "*$AssayName*")) {
                        $isMatch = $true
                    }
                    $products += $cell
                }
            }
            if ($isMatch) {
                $pn        = $ws.Cells[$r,1].Text
                $lotnr     = $ws.Cells[$r,2].Text
                $utg       = $ws.Cells[$r,3].Text
                $antal     = $ws.Cells[$r,4].Text
                $senast    = $ws.Cells[$r,5].Text
                $sign      = $ws.Cells[$r,6].Text
                $gmProd    = ($products -join ', ').Trim(',')
                $beskr     = $ws.Cells[$r,14].Text
                $plats     = $ws.Cells[$r,15].Text
                $labb      = $ws.Cells[$r,16].Text
                if (-not $antal -or "$antal".Trim().ToUpper() -eq 'N/A') { continue }
                $records += [PSCustomObject]@{
                    'P/N'                = $pn
                    'Lotnr.'             = $lotnr
                    'Utgångsdatum'       = $utg
                    'Antal i lager'      = $antal
                    'Senast inventering' = $senast
                    'SIGN'               = $sign
                    'Produkt (G-M)'      = $gmProd
                    'Beskrivning'        = $beskr
                    'Förvaringsplats'    = $plats
                    'Labb'               = $labb
                }
            }
        }
        $package.Dispose()
    } catch {
        Write-Warning "Failed to read control inventory from '$RawDataPath': $($_.Exception.Message)"
    }
    return $records
}

function Get-EquipmentReference {
    <#
        .SYNOPSIS
            Reads an equipment reference workbook and returns a map of instruments to calibrations and S/N.

        .DESCRIPTION
            This function parses an Excel workbook containing instrument names, serial numbers and calibration
            dates.  It uses internal dictionaries to map partial names and to group serial numbers by instrument.
            The return value is a hashtable keyed by instrument name with properties Calibration (MMM‑yy) and
            SNN (array of serial numbers).  If no file is provided or parsing fails, an empty hashtable is returned.

        .PARAMETER XlsPath
            Path to the equipment reference workbook (.xls).  The function assumes the first sheet contains
            the relevant data.

        .OUTPUTS
            A hashtable of equipment definitions.
    #>
    param([string]$XlsPath)
    $ref = @{}
    if ([string]::IsNullOrEmpty($XlsPath) -or -not (Test-Path -LiteralPath $XlsPath)) {
        return $ref
    }
    try {
        $fs = [System.IO.File]::Open($XlsPath, 'Open', 'Read')
        $wb = New-Object NPOI.HSSF.UserModel.HSSFWorkbook($fs)
        $sheet = $wb.GetSheetAt(0)
        # Instrument names to match in column 0 and mapping to key names
        $InstrumentMap = @{ 'GX1'='GeneXpert XVI - 16 Sites (Dell Computer), GX1';
                            'GX2'='GeneXpert XVI - 16 Sites (Dell Computer), GX2';
                            'GX3'='GeneXpert XVI - 16 Sites (Dell Computer), GX3';
                            'GX5'='GeneXpert XVI - 16 Sites (Dell Computor), GX5';
                            'GX6'='GeneXpert XVI - 16 Sites (Dell Computor), GX6';
                            'GX7'='GeneXpert XVI - 16 Sites (Dell Computor), GX-7';
                            'Infinity-I'='Infinity-I'; 'Infinity-III'='Infinity-III';
                            'Infinity-V'='Infinity-V'; 'Infinity-VI'='Infinity-VI';
                            'Infinity-VIII'='Infinity-VIII' }
        $SNMap = @{ 'GX1'=@('709863','709864','709865','709866');
                     'GX2'=@('709951','709952','709953','709954');
                     'GX3'=@('710084','710085','710086','710087');
                     'GX5'=@('750210','750211','750212','750213');
                     'GX6'=@('750246','750247','750248','750249');
                     'GX7'=@('750170','750171','750172','750173') }
        for ($i = 0; $i -le $sheet.LastRowNum; $i++) {
            if ($i -eq 8) { continue } # Skip header row
            $row = $sheet.GetRow($i)
            if (-not $row) { continue }
            $origName = $row.GetCell(0)?.ToString()
            $serial   = $row.GetCell(4)?.ToString()
            $calRaw   = $row.GetCell(8)?.ToString()
            foreach ($key in $InstrumentMap.Keys) {
                if ($origName -like "*${($InstrumentMap[$key])}*") {
                    $parsed = $null
                    $cal = if ([datetime]::TryParse($calRaw,[ref]$parsed)) {
                        $parsed.ToString('MMM-yy',[System.Globalization.CultureInfo]::InvariantCulture)
                    } else { $calRaw }
                    $snn = if ($SNMap.ContainsKey($key)) { $SNMap[$key] } else { @($serial) }
                    $ref[$key] = @{ Calibration=$cal; SNN=$snn }
                    break
                }
            }
        }
        $fs.Close()
    } catch {
        Write-Warning "Failed to read equipment reference from '$XlsPath': $($_.Exception.Message)"
    }
    return $ref
}

function Get-PipetteReference {
    <#
        .SYNOPSIS
            Reads a pipette reference workbook and returns a map of pipette ID to calibration date.

        .DESCRIPTION
            This function parses a sheet containing pipette information.  It returns a hashtable
            keyed by pipette ID with the calibration due date (MMM‑yy).  If the file is missing
            or the parsing fails, an empty hashtable is returned.

        .PARAMETER XlsPath
            Path to the pipette reference workbook (.xls).  Assumes the first sheet holds data.

        .OUTPUTS
            A hashtable mapping pipette IDs to calibration information.
    #>
    param([string]$XlsPath)
    $ref = @{}
    if ([string]::IsNullOrEmpty($XlsPath) -or -not (Test-Path -LiteralPath $XlsPath)) {
        return $ref
    }
    try {
        $fs = [System.IO.File]::Open($XlsPath, 'Open', 'Read')
        $wb = New-Object NPOI.HSSF.UserModel.HSSFWorkbook($fs)
        $sheet = $wb.GetSheetAt(0)
        for ($i = 30; $i -le 52; $i++) {
            $row = $sheet.GetRow($i)
            if (-not $row) { continue }
            $mappedName = $row.GetCell(2)?.ToString()
            $calRaw     = $row.GetCell(4)?.ToString()
            $parsed = $null
            $calDate = if ([datetime]::TryParse($calRaw,[ref]$parsed)) {
                $parsed.ToString('MMM-yy',[System.Globalization.CultureInfo]::InvariantCulture)
            } else { $calRaw }
            if (-not [string]::IsNullOrWhiteSpace($mappedName)) {
                $ref[$mappedName] = @{ Calibration=$calDate }
            }
        }
        $fs.Close()
    } catch {
        Write-Warning "Failed to read pipette reference from '$XlsPath': $($_.Exception.Message)"
    }
    return $ref
}
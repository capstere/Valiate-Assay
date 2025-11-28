function Convert-ToBoolSafe {
    param(
        [Parameter(ValueFromPipeline = $true)]
        $Value
    )
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }
    $s = $Value.ToString().Trim()
    if ($s -match '^(?i)(true|yes|ja|1)$') { return $true }
    if ($s -match '^(?i)(false|no|nej|0)$') { return $false }
    return $false
}

function Normalize-AssayKey {
    param([string]$Name)
    if (-not $Name) { return $null }
    $n = $Name.Trim().ToUpper()
    $n = $n -replace '[-\s]+', '_'  # replace spaces and dashes with _
    return $n
}

$script:ControlMaterialMap = $null

function Get-ColIndex {
    param(
        [hashtable]$Map,
        [string]$Name
    )
    if ($Map.ContainsKey($Name)) {
        return $Map[$Name]
    }
    else {
        return $null
    }
}

function Get-ControlMaterialMap {
    if ($script:ControlMaterialMap) {
        return $script:ControlMaterialMap
    }
    if (-not $Global:ControlMaterialMapPath -or -not (Test-Path -LiteralPath $Global:ControlMaterialMapPath)) {
        return $null
    }

    # Säkerställ att EPPlus är laddat innan vi skapar paketet
    if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'EPPlus' })) {
        $epplusLoaded = $false
        if (Get-Command Load-EPPlus -ErrorAction SilentlyContinue) {
            try { $epplusLoaded = [bool](Load-EPPlus) } catch { $epplusLoaded = $false }
        }
        if (-not $epplusLoaded -and (Get-Command Ensure-EPPlus -ErrorAction SilentlyContinue)) {
            $dllPath = $null
            try { $dllPath = Ensure-EPPlus -Version '4.5.3.3' } catch { $dllPath = $null }
            if ($dllPath -and (Test-Path -LiteralPath $dllPath)) {
                try {
                    $bytes = [System.IO.File]::ReadAllBytes($dllPath)
                    [System.Reflection.Assembly]::Load($bytes) | Out-Null
                    $epplusLoaded = $true
                }
                catch { $epplusLoaded = $false }
            }
        }
        if (-not $epplusLoaded) { throw "EPPlus kunde inte laddas." }
    }

    try {
        $file = New-Object IO.FileInfo ($Global:ControlMaterialMapPath)
        $pkg  = New-Object OfficeOpenXml.ExcelPackage $file
        $wsMaster = $pkg.Workbook.Worksheets['PartNoMaster']
        $wsUsage  = $pkg.Workbook.Worksheets['AssayUsage']
        if (-not $wsMaster -or -not $wsMaster.Dimension) { return $null }

        # Build PartNoIndex
        $partNoIndex = @{}
        $startRowM = $wsMaster.Dimension.Start.Row
        $endRowM   = $wsMaster.Dimension.End.Row
        $startColM = $wsMaster.Dimension.Start.Column
        $endColM   = $wsMaster.Dimension.End.Column

        # Map column names to indices
        $headerMapM = @{}
        for ($c = $startColM; $c -le $endColM; $c++) {
            $h = ($wsMaster.Cells[$startRowM, $c].Text + '').Trim()
            if ($h) { $headerMapM[$h] = $c }
        }

        $colPartNo       = Get-ColIndex -Map $headerMapM -Name 'PartNo'
        $colNameOfficial = Get-ColIndex -Map $headerMapM -Name 'NameOfficial'
        $colRole         = Get-ColIndex -Map $headerMapM -Name 'Role'
        $colPolarity     = Get-ColIndex -Map $headerMapM -Name 'ControlPolarity'
        $colCategory     = Get-ColIndex -Map $headerMapM -Name 'Category'
        $colMatrix       = Get-ColIndex -Map $headerMapM -Name 'Matrix'
        $colUseSpLot     = Get-ColIndex -Map $headerMapM -Name 'UseSharePointLot'
        $colActiveSE     = Get-ColIndex -Map $headerMapM -Name 'ActiveInSweden'
        $colSharedGroup  = Get-ColIndex -Map $headerMapM -Name 'SharedGroup'
        $colNotes        = Get-ColIndex -Map $headerMapM -Name 'Notes'

        for ($r = $startRowM + 1; $r -le $endRowM; $r++) {
            $pn = ($wsMaster.Cells[$r, $colPartNo].Text + '').Trim()
            if (-not $pn) { continue }
            $key = $pn.ToUpper().Trim()
            $rowObj = [pscustomobject]@{
                PartNo           = $pn
                NameOfficial     = if ($colNameOfficial) { ($wsMaster.Cells[$r, $colNameOfficial].Text + '').Trim() } else { $null }
                Role             = if ($colRole) { ($wsMaster.Cells[$r, $colRole].Text + '').Trim() } else { $null }
                ControlPolarity  = if ($colPolarity) { ($wsMaster.Cells[$r, $colPolarity].Text + '').Trim() } else { $null }
                Category         = if ($colCategory) { ($wsMaster.Cells[$r, $colCategory].Text + '').Trim() } else { $null }
                Matrix           = if ($colMatrix) { ($wsMaster.Cells[$r, $colMatrix].Text + '').Trim() } else { $null }
                UseSharePointLot = if ($colUseSpLot) { Convert-ToBoolSafe ($wsMaster.Cells[$r, $colUseSpLot].Value) } else { $false }
                ActiveInSweden   = if ($colActiveSE) { Convert-ToBoolSafe ($wsMaster.Cells[$r, $colActiveSE].Value) } else { $false }
                SharedGroup      = if ($colSharedGroup) { ($wsMaster.Cells[$r, $colSharedGroup].Text + '').Trim() } else { $null }
                Notes            = if ($colNotes) { ($wsMaster.Cells[$r, $colNotes].Text + '').Trim() } else { $null }
            }
            $partNoIndex[$key] = $rowObj
        }

        # Build AssayUsageIndex
        $assayUsageIndex = @{}
        if ($wsUsage -and $wsUsage.Dimension) {
            $startRowU = $wsUsage.Dimension.Start.Row
            $endRowU   = $wsUsage.Dimension.End.Row
            $startColU = $wsUsage.Dimension.Start.Column
            $endColU   = $wsUsage.Dimension.End.Column

            $headerMapU = @{}
            for ($c = $startColU; $c -le $endColU; $c++) {
                $h = ($wsUsage.Cells[$startRowU, $c].Text + '').Trim()
                if ($h) { $headerMapU[$h] = $c }
            }

            $colAssayKey    = Get-ColIndex -Map $headerMapU -Name 'AssayKey'
            $colAssayDisp   = Get-ColIndex -Map $headerMapU -Name 'AssayDisplayName'
            $colPartNoU     = Get-ColIndex -Map $headerMapU -Name 'PartNo'
            $colPolarityU   = Get-ColIndex -Map $headerMapU -Name 'ControlPolarity'
            $colActiveSEU   = Get-ColIndex -Map $headerMapU -Name 'ActiveInSweden'
            $colNotesU      = Get-ColIndex -Map $headerMapU -Name 'Notes'

            for ($r = $startRowU + 1; $r -le $endRowU; $r++) {
                $assayKeyRaw = ($wsUsage.Cells[$r, $colAssayKey].Text + '').Trim()
                $pnRaw       = ($wsUsage.Cells[$r, $colPartNoU].Text + '').Trim()
                if (-not $assayKeyRaw -or -not $pnRaw) { continue }

                $assayKeyNorm = Normalize-AssayKey $assayKeyRaw
                $pnNorm       = $pnRaw.ToUpper().Trim()

                $usageRow = [pscustomobject]@{
                    AssayKey         = $assayKeyNorm
                    AssayDisplayName = if ($colAssayDisp) { ($wsUsage.Cells[$r, $colAssayDisp].Text + '').Trim() } else { $assayKeyRaw }
                    PartNo           = $pnRaw
                    PartNoNorm       = $pnNorm
                    ControlPolarity  = if ($colPolarityU) { ($wsUsage.Cells[$r, $colPolarityU].Text + '').Trim() } else { $null }
                    ActiveInSweden   = if ($colActiveSEU) { Convert-ToBoolSafe ($wsUsage.Cells[$r, $colActiveSEU].Value) } else { $false }
                    Notes            = if ($colNotesU) { ($wsUsage.Cells[$r, $colNotesU].Text + '').Trim() } else { $null }
                }

                if (-not $assayUsageIndex.ContainsKey($assayKeyNorm)) {
                    $assayUsageIndex[$assayKeyNorm] = @()
                }
                $assayUsageIndex[$assayKeyNorm] += $usageRow
            }
        }

        $script:ControlMaterialMap = [pscustomobject]@{
            PartNoIndex     = $partNoIndex
            AssayUsageIndex = $assayUsageIndex
        }
        return $script:ControlMaterialMap
    }
    catch {
        $script:ControlMaterialMap = $null
        throw
    }
    finally {
        if ($pkg) { $pkg.Dispose() }
    }
}

function Get-TestSummaryControlsFromWorksheet {
    param(
        [Parameter(Mandatory = $true)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet
    )

    $results = @()
    if (-not $Worksheet -or -not $Worksheet.Dimension) { return $results }

    $maxR = $Worksheet.Dimension.End.Row
    $maxC = $Worksheet.Dimension.End.Column

    # Find the row containing Table 1 header
    $tableStart = $null
    for ($r = 1; $r -le $maxR; $r++) {
        $txt = ($Worksheet.Cells[$r, 1].Text + '').Trim()
        if ($txt -match '(?i)table\s*1.*test\s+and\s+control\s+material') {
            $tableStart = $r
            break
        }
    }
    if (-not $tableStart) { return $results }

    # Determine end of table by finding next "Table" header
    $tableEnd = $maxR
    for ($r = $tableStart + 1; $r -le [Math]::Min($maxR, $tableStart + 40); $r++) {
        $txt = ($Worksheet.Cells[$r, 1].Text + '').Trim()
        if ($txt -match '(?i)^table\s*\d') {
            $tableEnd = $r - 1
            break
        }
    }

    # Collect possible header rows (must contain at least Part and Lot keywords)
    $headerRows = @()
    for ($r = $tableStart + 1; $r -le $tableEnd; $r++) {
        $countKeywords = 0
        for ($c = 1; $c -le $maxC; $c++) {
            $cellText = ($Worksheet.Cells[$r, $c].Text + '').Trim()
            if ($cellText -match '(?i)part\s*no') { $countKeywords++ }
            if ($cellText -match '(?i)lot\s*no') { $countKeywords++ }
        }
        if ($countKeywords -ge 2) { $headerRows += $r }
    }

    foreach ($hr in $headerRows) {
        $nameRow = $hr - 1
        $dataRow = $hr + 1
        if ($nameRow -le $tableStart - 1 -or $dataRow -gt $tableEnd) { continue }

        # Identify columns containing Part No, Lot No and Exp info
        for ($c = 1; $c -le $maxC; $c++) {
            $hdrCell = ($Worksheet.Cells[$hr, $c].Text + '').Trim()
            if (-not $hdrCell -or $hdrCell -notmatch '(?i)part\s*no') { continue }

            $partCol = $c
            $lotCol  = $null
            $expCol  = $null

            # Determine adjacent columns for Lot and Exp
            if ($c + 1 -le $maxC) {
                $lotHdr = ($Worksheet.Cells[$hr, $c + 1].Text + '').Trim()
                if ($lotHdr -match '(?i)lot\s*no') { $lotCol = $c + 1 }
            }
            if ($c + 2 -le $maxC) {
                $expHdr = ($Worksheet.Cells[$hr, $c + 2].Text + '').Trim()
                if ($expHdr -match '(?i)exp') { $expCol = $c + 2 }
            }

            # Determine name (the row above the header).  If blank, look left/right.
            $name = ($Worksheet.Cells[$nameRow, $c].Text + '').Trim()
            if (-not $name) {
                for ($offset = -2; $offset -le 2; $offset++) {
                    $c2 = $c + $offset
                    if ($c2 -lt 1 -or $c2 -gt $maxC) { continue }
                    $name = ($Worksheet.Cells[$nameRow, $c2].Text + '').Trim()
                    if ($name) { break }
                }
            }
            if (-not $name) { continue }

            # Read data fields
            $partRaw = ($Worksheet.Cells[$dataRow, $partCol].Text + '').Trim()
            $lotRaw  = $null
            $expRaw  = $null
            $expVal  = $null

            if ($lotCol) { $lotRaw = ($Worksheet.Cells[$dataRow, $lotCol].Text + '').Trim() }
            if ($expCol) {
                $expRaw = ($Worksheet.Cells[$dataRow, $expCol].Text + '').Trim()
                $expVal = $Worksheet.Cells[$dataRow, $expCol].Value
            }

            # Normalise part numbers (split on newlines, commas and "or")
            $pnList = @()
            if ($partRaw) {
                $joined  = ($partRaw -split "(\r?\n)" | Where-Object { $_.Trim() }) -join ' '
                $matches = [regex]::Matches($joined, '(?<!\d)(\d{3}-\d{5})(?!\d)')
                foreach ($m in $matches) { $pnList += $m.Groups[1].Value }
                if ($pnList.Count -gt 0) { $pnList = @($pnList | Select-Object -Unique) }
            }

            # Determine canonical Lot number (first long alphanumeric token)
            $lotCanonical = $null
            $lotIsNA      = $false
            if ($lotRaw) {
                if ($lotRaw -match '^(?i)\s*n/?a\s*$') {
                    $lotIsNA = $true
                }
                else {
                    $tokens = $lotRaw -split '[,;\s]+'
                    foreach ($tok in $tokens) {
                        $tt = $tok.Trim()
                        if ($tt -match '^[0-9A-Z\-]{6,}$') { $lotCanonical = $tt; break }
                    }
                }
            }

            # Convert Expiration date
            $expDate = $null
            if ($expVal -is [datetime]) {
                $expDate = $expVal
            }
            elseif ($expVal -is [double] -or $expVal -is [int]) {
                try {
                    $base = Get-Date '1899-12-30'
                    $expDate = $base.AddDays([double]$expVal)
                }
                catch { $expDate = $null }
            }
            elseif ($expRaw) {
                try {
                    $expDate = Get-Date -Date $expRaw -ErrorAction Stop
                }
                catch { $expDate = $null }
            }

            $results += [pscustomobject]@{
                Name           = $name
                PartNoRaw      = $partRaw
                PartNos        = $pnList
                LotNoRaw       = $lotRaw
                LotNoCanonical = $lotCanonical
                ExpRaw         = $expRaw
                ExpDate        = $expDate
                Row            = $dataRow
                Column         = $partCol
            }
        }
    }
    return $results
}

function Get-TestSummaryControls {

    param(
        [Parameter(Mandatory = $true)]
        [OfficeOpenXml.ExcelPackage]$Pkg
    )

    $empty = [pscustomobject]@{ WorksheetName = $null; Controls = @() }
    if (-not $Pkg) { return $empty }

    $best = $null
    foreach ($ws in $Pkg.Workbook.Worksheets) {
        if (-not $ws.Dimension) { continue }
        if ($ws.Name -match '(?i)worksheet\s*instructions') { continue }

        $controls = Get-TestSummaryControlsFromWorksheet -Worksheet $ws
        if ($controls -and $controls.Count -gt 0) {
            if ($ws.Name -match '(?i)test\s*summary') {
                return [pscustomobject]@{ WorksheetName = $ws.Name; Controls = $controls }
            }
            $best = [pscustomobject]@{ WorksheetName = $ws.Name; Controls = $controls }
        }
    }

    if ($best) { return $best }
    return $empty
}

function Get-ControlMaterialSpecFromSheet {
    param(
        [Parameter(Mandatory = $true)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet
    )

    $items = @()
    if (-not $Worksheet -or -not $Worksheet.Dimension) { return $items }

    $maxR = $Worksheet.Dimension.End.Row

    # Find header row containing P/N and Lotnr.
    $hdrRow = $null
    for ($r = 1; $r -le $maxR; $r++) {
        $a = ($Worksheet.Cells[$r, 1].Text + '').Trim()
        $c = ($Worksheet.Cells[$r, 3].Text + '').Trim()
        if ($a -match '(?i)^p/?n' -and $c -match '(?i)lot') { $hdrRow = $r; break }
    }
    if (-not $hdrRow) { return $items }

    $row = $hdrRow + 1
    while ($row -le $maxR) {
        $pn    = ($Worksheet.Cells[$row, 1].Text + '').Trim()
        $art   = ($Worksheet.Cells[$row, 2].Text + '').Trim()
        $lot   = ($Worksheet.Cells[$row, 3].Text + '').Trim()
        $expVal = $Worksheet.Cells[$row, 4].Value

        if (-not $pn -and -not $art -and -not $lot -and -not $expVal) { break }

        if ($pn) {
            # Determine canonical lot
            $lotCanonical = $null
            if ($lot) {
                $tokens = $lot.Split(' ', ',')
                foreach ($tok in $tokens) {
                    $t = $tok.Trim().ToUpper()
                    if ($t -match '^[0-9A-Z\-]{6,}$') { $lotCanonical = $t; break }
                }
            }

            # Convert expiration
            $expDate = $null
            if ($expVal -is [datetime]) {
                $expDate = $expVal
            }
            elseif ($expVal -is [double] -or $expVal -is [int]) {
                try {
                    $base = Get-Date '1899-12-30'
                    $expDate = $base.AddDays([double]$expVal)
                }
                catch { $expDate = $null }
            }
            elseif ($expVal) {
                try { $expDate = Get-Date $expVal -ErrorAction Stop } catch { $expDate = $null }
            }

            $items += [pscustomobject]@{
                Name           = $art
                PartNos        = @($pn)
                LotNoCanonical = $lotCanonical
                ExpDate        = $expDate
            }
        }
        $row++
    }

    return $items
}

function Compare-TestSummaryControls {
    param(
        [Parameter(Mandatory = $true)] [array]$TsControls,
        [Parameter(Mandatory = $true)] [array]$SpecItems
    )

    $results = @()
    foreach ($ts in $TsControls) {
        $matched    = $false
        $candidates = @()

        foreach ($sp in $SpecItems) {
            # match if any part number overlaps
            $partMatch = $false
            foreach ($p in $ts.PartNos) {
                if ($sp.PartNos -contains $p) { $partMatch = $true; break }
            }
            if ($partMatch) {
                $candidates += $sp
                continue
            }

            # fallback: match by name (case-insensitive)
            if ($sp.Name -and $ts.Name -and ($sp.Name.Trim().ToLower() -eq $ts.Name.Trim().ToLower())) {
                $candidates += $sp
            }
        }

        if ($candidates.Count -eq 0) {
            $results += [pscustomobject]@{
                Name        = $ts.Name
                TsPartNos   = ($ts.PartNos -join ', ')
                TsLotNo     = $ts.LotNoRaw
                TsExpDate   = $ts.ExpDate
                SpecPartNos = ''
                SpecLotNo   = ''
                SpecExpDate = ''
                Status      = 'Missing in spec'
            }
            continue
        }

        # Evaluate first candidate (could refine later)
        $spBest  = $candidates[0]
        $specLot = $spBest.LotNoCanonical
        $specExp = $spBest.ExpDate

        $lotOk = $false
        if ($ts.LotNoCanonical -and $specLot) {
            if ($ts.LotNoCanonical.ToUpper() -eq $specLot.ToUpper()) { $lotOk = $true }
        }
        elseif (-not $ts.LotNoCanonical -and -not $specLot) {
            $lotOk = $true
        }

        $expOk = $true
        if ($ts.ExpDate -and $specExp -and ($ts.ExpDate -is [datetime]) -and ($specExp -is [datetime])) {
            if ($ts.ExpDate.Date -ne $specExp.Date) { $expOk = $false }
        }

        $status =
            if     ($lotOk -and $expOk) { 'OK' }
            elseif ($lotOk -and -not $expOk) { 'Expiration mismatch' }
            elseif (-not $lotOk -and $expOk) { 'Lot mismatch' }
            else { 'Lot & Exp mismatch' }

        $results += [pscustomobject]@{
            Name        = $ts.Name
            TsPartNos   = ($ts.PartNos -join ', ')
            TsLotNo     = $ts.LotNoCanonical
            TsExpDate   = $ts.ExpDate
            SpecPartNos = ($spBest.PartNos -join ', ')
            SpecLotNo   = $specLot
            SpecExpDate = $specExp
            Status      = $status
        }
    }

    return $results
}

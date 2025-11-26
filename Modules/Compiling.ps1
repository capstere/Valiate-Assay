if (Get-Command Convert-XpertSummaryCsvToData -ErrorAction SilentlyContinue) { Remove-Item Function:\Convert-XpertSummaryCsvToData -ErrorAction SilentlyContinue }
if (Get-Command Parse-CompilingSampleId -ErrorAction SilentlyContinue)     { Remove-Item Function:\Parse-CompilingSampleId     -ErrorAction SilentlyContinue }
if (Get-Command Get-CompilingAnalysis -ErrorAction SilentlyContinue)       { Remove-Item Function:\Get-CompilingAnalysis       -ErrorAction SilentlyContinue }

$script:CompilingAssayDesigns = @(
    @{
        Name         = 'Group1'
        DisplayName  = 'MTB RIF / FLUVID / FLUVID+ / SARS-COV-2 Plus / CTNG / SARS-COV-2 / MRSA SA / VAN AB / GBS'
        Assays       = @('MTB RIF','FLUVID+','FLUVID','SARS-COV-2 PLUS','SARS-COV-2','CTNG','MRSA SA','VAN AB','GBS')
        ControlTypes = @(
            @{ ControlType='0'; Label='NEG'; Bags=0..10; Replicates=1..10; ExpectedCount=110 },
            @{ ControlType='1'; Label='POS'; Bags=1..10; Replicates=11..20; ExpectedCount=100 }
        )
    },
    @{
        Name         = 'Group7' # Placera JP-variant först för mer specifik matchning
        DisplayName  = 'C.DIFF JP'
        Assays       = @('C.DIFF JP')
        ControlTypes = @(
            @{ ControlType='0'; Label='NEG'; Bags=0..10; Replicates=1..10; ExpectedCount=110 },
            @{ ControlType='1'; Label='POS'; Bags=1..10; Replicates=11..17; ExpectedCount=70 },
            @{ ControlType='2'; Label='POS'; Bags=1..10; Replicates=18..19; ExpectedCount=20 },
            @{ ControlType='3'; Label='POS'; Bags=1..10; Replicates=20..20; ExpectedCount=10 }
        )
    },
    @{
        Name         = 'Group6'
        DisplayName  = 'MTB JP'
        Assays       = @('MTB JP')
        ControlTypes = @(
            @{ ControlType='0'; Label='NEG'; Bags=0..10; Replicates=1..10; ExpectedCount=110 },
            @{ ControlType='1'; Label='POS'; Bags=1..10; Replicates=11..16; ExpectedCount=60 },
            @{ ControlType='2'; Label='POS'; Bags=1..10; Replicates=17..17; ExpectedCount=10 },
            @{ ControlType='3'; Label='POS'; Bags=1..10; Replicates=18..18; ExpectedCount=10 },
            @{ ControlType='4'; Label='POS'; Bags=1..10; Replicates=19..19; ExpectedCount=10 },
            @{ ControlType='5'; Label='POS'; Bags=1..10; Replicates=20..20; ExpectedCount=10 }
        )
    },
    @{
        Name         = 'Group2'
        DisplayName  = 'MTB ULTRA / MTB XDR / C.DIFF / EBOLA / FLU RSV / MRSA NxG / NORO / STREP A'
        Assays       = @('MTB ULTRA','MTB XDR','C.DIFF','EBOLA','FLU RSV','MRSA NXG','NORO','STREP A')
        ControlTypes = @(
            @{ ControlType='0'; Label='NEG'; Bags=0..10; Replicates=1..10; ExpectedCount=110 },
            @{ ControlType='1'; Label='POS'; Bags=1..10; Replicates=11..18; ExpectedCount=80 },
            @{ ControlType='2'; Label='POS'; Bags=1..10; Replicates=19..20; ExpectedCount=20 }
        )
    },
    @{
        Name         = 'Group3'
        DisplayName  = 'HBV VL / HCV VL / HCV VL FS / HIV VL / HIV VL XC / HIV QA / HIV QA XC'
        Assays       = @('HBV VL','HCV VL','HCV VL FS','HIV VL','HIV VL XC','HIV QA','HIV QA XC')
        ControlTypes = @(
            @{ ControlType='0'; Label='NEG'; Bags=1..10; Replicates=1..10; ExpectedCount=100 },
            @{ ControlType='1'; Label='POS'; Bags=0..10; Replicates=11..18; ExpectedCount=90 },
            @{ ControlType='2'; Label='POS'; Bags=1..10; Replicates=19..20; ExpectedCount=20 }
        )
    },
    @{
        Name         = 'Group4'
        DisplayName  = 'CARBA R'
        Assays       = @('CARBA R')
        ControlTypes = @(
            @{ ControlType='0'; Label='NEG'; Bags=0..10; Replicates=1..14; ExpectedCount=150 },
            @{ ControlType='1'; Label='POS'; Bags=1..10; Replicates=15..18; ExpectedCount=40 },
            @{ ControlType='2'; Label='POS'; Bags=1..10; Replicates=19..20; ExpectedCount=20 }
        )
    },
    @{
        Name         = 'Group5'
        DisplayName  = 'HPV'
        Assays       = @('HPV')
        ControlTypes = @(
            @{ ControlType='0'; Label='NEG'; Bags=1..10; Replicates=1..6; ExpectedCount=60 },
            @{ ControlType='1'; Label='POS'; Bags=0..10; Replicates=7..18; ExpectedCount=130 },
            @{ ControlType='2'; Label='POS'; Bags=1..10; Replicates=19..20; ExpectedCount=20 }
        )
    },
    @{
        Name         = 'Group8'
        DisplayName  = 'Respiratory Panel'
        Assays       = @('RESPIRATORY PANEL','RESPIRATORY PANEL JP','RESPIRATORY PANEL R')
        ControlTypes = @(
            @{ ControlType='0'; Label='NEG'; Bags=0..10; Replicates=1..10; ExpectedCount=110 },
            @{ ControlType='1'; Label='POS'; Bags=1..10; Replicates=11..14; ExpectedCount=40 },
            @{ ControlType='2'; Label='POS'; Bags=1..10; Replicates=15..16; ExpectedCount=20 },
            @{ ControlType='3'; Label='POS'; Bags=1..10; Replicates=17..18; ExpectedCount=20 },
            @{ ControlType='4'; Label='POS'; Bags=1..10; Replicates=19..20; ExpectedCount=20 }
        )
    }
)

function New-CompilingSummary {
    return [pscustomobject]@{
        AssayName               = ''
        AssayVersion            = ''
        LspSummary              = ''
        AssayGroup              = $null
        TestCount               = 0
        ControlMaterialCount    = [ordered]@{}
        ControlTypeStats        = @()
        MissingReplicates       = @()
        Replacements            = @()
        Delaminations           = @()
        DuplicateSampleIDs      = @()
        DuplicateCartridgeSN    = @()
        ErrorCodes              = @{}
        ErrorList               = @()
        InstrumentByType        = [ordered]@{}
        BagControlReplicateMap  = @{}
    }
}

function Convert-XpertSummaryCsvToData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) { return @() }

    $lines = $null
    try { $lines = Get-Content -LiteralPath $Path } catch { return @() }
    if (-not $lines -or $lines.Count -lt 8) { return @() }

    $delimiter = ';'
    try { $delimiter = Get-CsvDelimiter -Path $Path } catch {}

    $headerFields = ConvertTo-CsvFields $lines[7]
    $seen = @{}
    $header = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $headerFields.Count; $i++) {
        $raw = ($headerFields[$i] + '').Trim('"')
        switch -Regex ($raw) {
            '#\s*TBD\s*Test\s*Result\s*#' { $raw = 'Test Result' }
            '^(?i)Test\s*Result$'         { $raw = 'Test Result' }
            default { }
        }
        if (-not $raw) { $raw = "Column$($i+1)" }
        if ($seen.ContainsKey($raw)) {
            $seen[$raw]++
            $raw = "$raw$($seen[$raw])"
        } else {
            $seen[$raw] = 1
        }
        $null = $header.Add($raw)
    }

    $dataLines = @()
    if ($lines.Count -gt 9) {
        $dataLines = $lines[9..($lines.Count-1)] | Where-Object { $_ -and $_.Trim() }
    }
    if ($dataLines.Count -eq 0) { return @() }

    try {
        return ,@(ConvertFrom-Csv -InputObject ($dataLines -join "`n") -Delimiter $delimiter -Header $header)
    } catch {
        try { return ,@(ConvertFrom-Csv -InputObject ($dataLines -join "`n") -Header $header) } catch { return @() }
    }
}

function Parse-CompilingSampleId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)][string]$SampleId
    )

    $result = [pscustomobject]@{
        RawSampleId     = $SampleId
        ControlMaterial = $null
        Bag             = $null
        ControlType     = $null
        ReplicateNumber = $null
        ReplicateTags   = ''
        ReplacementType = 'None'
    }

    if ([string]::IsNullOrWhiteSpace($SampleId)) { return $result }

    $parts = ($SampleId + '') -split '_'
    if ($parts.Count -ge 1) { $result.ControlMaterial = ($parts[0] + '').Trim() }

    if ($parts.Count -ge 2) {
        $bagVal = $null
        if ([int]::TryParse($parts[1], [ref]$bagVal)) { $result.Bag = $bagVal }
    }

    if ($parts.Count -ge 3) {
        $mCt = [regex]::Match($parts[2], '(\d+)')
        if ($mCt.Success) { $result.ControlType = $mCt.Groups[1].Value }
    }

    if ($parts.Count -ge 4) {
        $repSegment = ($parts[3..($parts.Count-1)] -join '_')
        $mRep = [regex]::Match($repSegment, '^(\d{1,2})')
        if ($mRep.Success) { $result.ReplicateNumber = [int]$mRep.Groups[1].Value }
        $letters = [regex]::Matches($repSegment, '[A-Za-z]+') | ForEach-Object { $_.Value }
        if ($letters.Count -gt 0) { $result.ReplicateTags = (($letters -join '')).ToUpperInvariant() } else { $result.ReplicateTags = '' }
    }

    $tags = $result.ReplicateTags
    if ($tags -match 'D')            { $result.ReplacementType = 'Delam' }
    elseif ($tags -eq 'AAA')         { $result.ReplacementType = 'Replacement3' }
    elseif ($tags -eq 'AA')          { $result.ReplacementType = 'Replacement2' }
    elseif ($tags -match 'A')        { $result.ReplacementType = 'Replacement1' }
    else                             { $result.ReplacementType = 'None' }

    return $result
}

function Get-CompilingAssayDefinition {
    param([string]$AssayName)

    if (-not $AssayName) { return $null }
    $norm = Normalize-HeaderText $AssayName
    $upper = $norm.ToUpperInvariant()

    $best = $null
    $bestLen = 0
    foreach ($def in $script:CompilingAssayDesigns) {
        foreach ($pat in $def.Assays) {
            $pNorm = (Normalize-HeaderText $pat).ToUpperInvariant()
            if ($upper -like "*$pNorm*") {
                if ($pNorm.Length -gt $bestLen) {
                    $best = $def
                    $bestLen = $pNorm.Length
                }
            }
        }
    }
    return $best
}

function Format-CompilingRange {
    param(
        [int[]]$Numbers,
        [string]$Prefix = '',
        [string]$Format = '{0:00}'
    )
    if (-not $Numbers -or $Numbers.Count -eq 0) { return '' }
    $distinct = ($Numbers | Select-Object -Unique | Sort-Object)
    $first = $distinct[0]; $last = $distinct[-1]
    if ($distinct.Count -eq 1) {
        return "$Prefix$([string]::Format($Format, $first))"
    }
    return "$Prefix$([string]::Format($Format, $first))-$([string]::Format($Format, $last))"
}

function Get-CompilingAnalysis {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Data
    )

    $result = New-CompilingSummary
    $rows = @()
    if ($Data) { $rows = @($Data) } else { return $result }

    $result.TestCount = $rows.Count

    foreach ($row in $rows) {
        if (-not $result.AssayName -and $row.PSObject.Properties['Assay'] -and ($row.'Assay')) {
            $result.AssayName = ($row.'Assay' + '').Trim()
        }
        if (-not $result.AssayVersion -and $row.PSObject.Properties['Assay Version'] -and ($row.'Assay Version')) {
            $result.AssayVersion = ($row.'Assay Version' + '').Trim()
        }
    }

    $assayDef = Get-CompilingAssayDefinition -AssayName $result.AssayName
    $result.AssayGroup = $assayDef

    $parsedRows = New-Object System.Collections.Generic.List[object]
    $bagCtrlRepMap = @{}
    $repls = New-Object System.Collections.Generic.List[object]
    $delams = New-Object System.Collections.Generic.List[object]
    $errorsMap = @{}
    $errorList = New-Object System.Collections.Generic.List[string]

    $instLut = @{}
    try {
        if ($script:GXINF_Map) {
            foreach ($k in $script:GXINF_Map.Keys) {
                $codes = $script:GXINF_Map[$k].Split(',') | ForEach-Object { ($_ + '').Trim() } | Where-Object { $_ }
                foreach ($code in $codes) { $instLut[$code] = $k }
            }
        }
    } catch {}

    $reagentLots = New-Object System.Collections.Generic.List[string]
    $cartList    = New-Object System.Collections.Generic.List[string]
    $sampleList  = New-Object System.Collections.Generic.List[string]

    foreach ($row in $rows) {
        $sampleId = ''
        if ($row.PSObject.Properties['Sample ID']) { $sampleId = ($row.'Sample ID' + '').Trim() }
        elseif ($row.PSObject.Properties['Sample']) { $sampleId = ($row.'Sample' + '').Trim() }

        $parsed = Parse-CompilingSampleId -SampleId $sampleId
        $parsedRows.Add([pscustomobject]@{ Row=$row; Parsed=$parsed })

        if ($parsed.ControlMaterial) {
            if (-not $result.ControlMaterialCount.Contains($parsed.ControlMaterial)) { $result.ControlMaterialCount[$parsed.ControlMaterial] = 0 }
            $result.ControlMaterialCount[$parsed.ControlMaterial]++
        }

        if ($parsed.Bag -ne $null -and $parsed.ControlType -and $parsed.ReplicateNumber -ne $null) {
            $key = "{0}|{1}|{2}" -f $parsed.Bag, $parsed.ControlType, $parsed.ReplicateNumber
            if (-not $bagCtrlRepMap.ContainsKey($key)) { $bagCtrlRepMap[$key] = New-Object System.Collections.Generic.List[object] }
            $bagCtrlRepMap[$key].Add($row)
        }

        if ($parsed.ReplacementType -eq 'Delam') {
            $delams.Add($parsed.RawSampleId)
        } elseif ($parsed.ReplacementType -ne 'None') {
            $repls.Add([pscustomobject]@{ SampleId=$parsed.RawSampleId; ReplacementType=$parsed.ReplacementType })
        }

        if ($row.PSObject.Properties['Error']) {
            $errVal = ($row.'Error' + '').Trim()
            if ($errVal) {
                $mErr = [regex]::Match($errVal, '(?i)error\s*(\d{3,5})')
                if ($mErr.Success) {
                    $codeKey = "Error $($mErr.Groups[1].Value)"
                    if (-not $errorsMap.ContainsKey($codeKey)) { $errorsMap[$codeKey] = New-Object System.Collections.Generic.List[string] }
                    $errorsMap[$codeKey].Add($sampleId)
                    $errorList.Add("$codeKey $sampleId")
                }
            }
        }

        if ($row.PSObject.Properties['Reagent Lot ID'] -and $row.'Reagent Lot ID') {
            $lot = ($row.'Reagent Lot ID' + '').Trim()
            if ($lot) { $reagentLots.Add($lot) }
        }

        $cartVal = $null
        if ($row.PSObject.Properties['Cartridge S/N']) { $cartVal = ($row.'Cartridge S/N' + '').Trim() }
        elseif ($row.PSObject.Properties['Cartridge']) { $cartVal = ($row.'Cartridge' + '').Trim() }
        if ($cartVal) { $cartList.Add($cartVal) }
        if ($sampleId) { $sampleList.Add($sampleId) }

        $insVal = $null
        $propNames = @('Instrument S/N','Instrument','Instrument SN')
        foreach ($pn in $propNames) {
            if ($row.PSObject.Properties[$pn] -and $row.$pn) { $insVal = ($row.$pn + '').Trim(); break }
        }
        if (-not $insVal) {
            $cand = $row.PSObject.Properties.Name | Where-Object { $_ -match '(?i)instrument' } | Select-Object -First 1
            if ($cand) { $insVal = ($row.$cand + '').Trim() }
        }
        if ($insVal) {
            $type = if ($instLut.ContainsKey($insVal)) { $instLut[$insVal] } else { 'Unknown' }
            if (-not $result.InstrumentByType.Contains($type)) { $result.InstrumentByType[$type] = 0 }
            $result.InstrumentByType[$type]++
        }
    }

    $result.BagControlReplicateMap = $bagCtrlRepMap
    $result.Replacements  = @($repls.ToArray())
    $result.Delaminations = @($delams.ToArray())
    $result.ErrorCodes    = $errorsMap
    $result.ErrorList     = @($errorList.ToArray())

    $dupSamples = $sampleList | Where-Object { $_ } | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($dupSamples) {
        $result.DuplicateSampleIDs = $dupSamples | ForEach-Object { [pscustomobject]@{ SampleId=$_.Name; Count=$_.Count } }
    }
    $dupCart = $cartList | Where-Object { $_ } | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($dupCart) {
        $result.DuplicateCartridgeSN = $dupCart | ForEach-Object { [pscustomobject]@{ Cartridge=$_.Name; Count=$_.Count } }
    }

    if ($reagentLots.Count -gt 0) {
        $lotCounts = @{}
        foreach ($lot in $reagentLots) {
            $code = $lot
            $mLot = [regex]::Match($lot,'(?<!\d)(\d{5})(?!\d)')
            if ($mLot.Success) { $code = $mLot.Groups[1].Value }
            if (-not $lotCounts.ContainsKey($code)) { $lotCounts[$code] = 0 }
            $lotCounts[$code]++
        }
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($k in ($lotCounts.Keys | Sort-Object)) {
            $cnt = $lotCounts[$k]
            $parts.Add( $(if ($cnt -gt 1) { "$k x$cnt" } else { $k }) )
        }
        $totalLots = $lotCounts.Count
        if ($totalLots -eq 1) { $result.LspSummary = $parts[0] }
        else { $result.LspSummary = "$totalLots (" + ($parts -join ', ') + ")" }
    }

    $ctStats = New-Object System.Collections.Generic.List[object]
    if ($assayDef) {
        foreach ($ct in $assayDef.ControlTypes) {
            $bags = @($ct.Bags)
            $reps = @($ct.Replicates)
            $designExpected = $bags.Count * $reps.Count
            $expected = if ($ct.ContainsKey('ExpectedCount')) { [int]$ct.ExpectedCount } else { $designExpected }
            $actual = ($parsedRows | Where-Object { $_.Parsed.ControlType -eq $ct.ControlType }).Count
            $missing = New-Object System.Collections.Generic.List[object]
            foreach ($b in $bags) {
                foreach ($r in $reps) {
                    $key = "{0}|{1}|{2}" -f $b, $ct.ControlType, $r
                    if (-not $bagCtrlRepMap.ContainsKey($key)) {
                        $missing.Add([pscustomobject]@{ Bag=$b; ControlType=$ct.ControlType; ReplicateNumber=$r })
                    }
                }
            }
            $ctStats.Add([pscustomobject]@{
                ControlType        = $ct.ControlType
                Label              = $ct.Label
                ExpectedCount      = $expected
                DesignExpected     = $designExpected
                ActualCount        = $actual
                MissingCount       = $missing.Count
                MissingReplicates  = @($missing.ToArray())
                BagRangeText       = Format-CompilingRange -Numbers $bags -Prefix 'Bag '
                ReplicateRangeText = Format-CompilingRange -Numbers $reps -Prefix 'rep '
                Ok                 = (($missing.Count -eq 0) -and ($actual -eq $expected))
            })
            if ($missing.Count -gt 0) {
                $result.MissingReplicates += @($missing.ToArray())
            }
        }
    } else {
        $groups = $parsedRows | Where-Object { $_.Parsed.ControlType } | Group-Object { $_.Parsed.ControlType }
        foreach ($g in $groups) {
            $ctStats.Add([pscustomobject]@{
                ControlType        = $g.Name
                Label              = ''
                ExpectedCount      = $null
                DesignExpected     = $null
                ActualCount        = $g.Count
                MissingCount       = $null
                MissingReplicates  = @()
                BagRangeText       = ''
                ReplicateRangeText = ''
                Ok                 = $null
            })
        }
    }
    $result.ControlTypeStats = @($ctStats | Sort-Object { [int]($_.ControlType) })

    return $result
}

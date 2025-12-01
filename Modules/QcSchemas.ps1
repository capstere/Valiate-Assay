param(
    [string]$ScriptRoot = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

if (Get-Command Get-QcSchema -ErrorAction SilentlyContinue) { Remove-Item Function:\Get-QcSchema -ErrorAction SilentlyContinue }
if (Get-Command Find-QcAssayDefinition -ErrorAction SilentlyContinue) { Remove-Item Function:\Find-QcAssayDefinition -ErrorAction SilentlyContinue }
if (Get-Command Get-QcAssayControlDesign -ErrorAction SilentlyContinue) { Remove-Item Function:\Get-QcAssayControlDesign -ErrorAction SilentlyContinue }

$script:QcSchemaCache = $null

function Get-QcSchema {
    param(
        [switch]$ForceReload
    )

    if (-not $ForceReload -and $script:QcSchemaCache) {
        return $script:QcSchemaCache
    }

    $candidates = @()
    $candidates += Join-Path $ScriptRoot 'Config\QcSchemaData.ps1'
    if ($PSScriptRoot) { $candidates += Join-Path (Split-Path -Parent $PSScriptRoot) 'Config\QcSchemaData.ps1' }
    $candidates = $candidates | Where-Object { $_ } | Select-Object -Unique

    $dataPath = $null
    foreach ($cand in $candidates) {
        if (Test-Path -LiteralPath $cand) { $dataPath = $cand; break }
    }

    if (-not $dataPath) { return $null }

    . $dataPath
    if ($Global:QcSchema) {
        $script:QcSchemaCache = $Global:QcSchema
    } elseif ($QcSchemaData) {
        $script:QcSchemaCache = $QcSchemaData
    } else {
        $script:QcSchemaCache = $null
    }

    return $script:QcSchemaCache
}

function Convert-QcRangeToArray {
    param(
        $Value
    )

    if ($null -eq $Value) { return @() }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @()
        foreach ($v in $Value) { if ($v -ne $null -and $v -ne '') { $items += [int]$v } }
        return ,@($items)
    }

    $text = ($Value + '').Trim()
    if (-not $text) { return @() }

    $rxDash = [regex]::Match($text, '^\s*(\d+)\s*-\s*(\d+)\s*$')
    if ($rxDash.Success) {
        $start = [int]$rxDash.Groups[1].Value
        $end   = [int]$rxDash.Groups[2].Value
        return $start..$end
    }

    $rxDot = [regex]::Match($text, '^\s*(\d+)\s*\.\.\s*(\d+)\s*$')
    if ($rxDot.Success) {
        $start = [int]$rxDot.Groups[1].Value
        $end   = [int]$rxDot.Groups[2].Value
        return $start..$end
    }

    if ($text -match ',') {
        $parts = $text -split ','
        $arr = @()
        foreach ($p in $parts) {
            $pVal = ($p + '').Trim()
            if ($pVal -ne '') { $arr += [int]$pVal }
        }
        return ,@($arr)
    }

    if ($text -match '^\d+$') { return ,([int]$text) }
    return @()
}

function Find-QcAssayDefinition {
    param(
        [string]$AssayName,
        [switch]$ForceReload
    )

    if (-not $AssayName) { return $null }

    $schema = Get-QcSchema -ForceReload:$ForceReload
    if (-not $schema -or -not $schema.Assays) { return $null }

    $normName = if (Get-Command Normalize-HeaderText -ErrorAction SilentlyContinue) { (Normalize-HeaderText $AssayName) } else { $AssayName }
    $upper = ($normName + '').ToUpperInvariant()

    $best = $null
    $bestScore = 0
    foreach ($assay in $schema.Assays) {
        $score = 0
        $matched = $false

        if ($assay.AliasPatterns) {
            foreach ($pat in $assay.AliasPatterns) {
                if (-not $pat) { continue }
                try {
                    if ($AssayName -match $pat) {
                        $matched = $true
                        if ($pat.Length -gt $score) { $score = $pat.Length }
                    }
                } catch {}
            }
        }

        if (-not $matched) {
            $assayNameNorm = if (Get-Command Normalize-HeaderText -ErrorAction SilentlyContinue) { (Normalize-HeaderText $assay.AssayName) } else { $assay.AssayName }
            $assayKeyNorm  = if (Get-Command Normalize-HeaderText -ErrorAction SilentlyContinue) { (Normalize-HeaderText $assay.AssayKey) } else { $assay.AssayKey }

            $assayNameUpper = ($assayNameNorm + '').ToUpperInvariant()
            $assayKeyUpper  = ($assayKeyNorm + '').ToUpperInvariant()

            if ($upper -like "*$assayNameUpper*") { $matched = $true; if ($assayNameUpper.Length -gt $score) { $score = $assayNameUpper.Length } }
            if ($upper -like "*$assayKeyUpper*")  { $matched = $true; if ($assayKeyUpper.Length  -gt $score) { $score = $assayKeyUpper.Length  } }
        }

        if ($matched -and $score -ge $bestScore) {
            $best = $assay
            $bestScore = $score
        }
    }

    return $best
}

function Get-QcAssayControlDesign {
    param(
        [string]$AssayName,
        [switch]$ForceReload
    )

    $assay = Find-QcAssayDefinition -AssayName $AssayName -ForceReload:$ForceReload
    if (-not $assay) { return $null }

    $controlTypes = @()
    if ($assay.ControlScheme) {
        foreach ($ct in $assay.ControlScheme) {
            if (-not $ct) { continue }
            $bags = Convert-QcRangeToArray $ct.ExpectedBagRange
            $reps = Convert-QcRangeToArray $ct.ExpectedReplicateRange
            $ctIndex = $ct.ControlTypeIndex
            if ($null -eq $ctIndex -and $ct.ContainsKey('Idx')) { $ctIndex = $ct.Idx }

            $controlTypes += [pscustomobject]@{
                ControlType   = [string]$ctIndex
                Label         = $ct.ControlLabel
                Bags          = $bags
                Replicates    = $reps
                ExpectedCount = if ($ct.ContainsKey('ExpectedCount')) { [int]$ct.ExpectedCount } else { ($bags.Count * $reps.Count) }
                Raw           = $ct
            }
        }
    }

    return [pscustomobject]@{
        Name         = $assay.AssayKey
        DisplayName  = $assay.AssayName
        Assay        = $assay
        ControlTypes = $controlTypes
    }
}

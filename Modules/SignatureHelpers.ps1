function Get-DataSheets { param([OfficeOpenXml.ExcelPackage]$Pkg)
    $all = @($Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" })
    if ($all.Count -gt 1) { return $all | Select-Object -Skip 1 } else { return @() }
}

function Test-SignatureFormat {
    param([string]$Text)
    $raw = ($Text + '')
    $trim = $raw.Trim()
    $parts = $trim -split '\s*,\s*'
    $name = if ($parts.Count -ge 1) { $parts[0] } else { '' }
    $sign = if ($parts.Count -ge 2) { $parts[1] } else { '' }
    $date = if ($parts.Count -ge 3) { $parts[2] } else { '' }
    $dateOk = $false
    if ($date) { if ($date -match '^\d{4}-\d{2}-\d{2}$' -or $date -match '^\d{8}$') { $dateOk = $true } }
    [pscustomobject]@{ Raw=$raw; Name=$name; Sign=$sign; Date=$date; Parts=$parts.Count; DateOk=$dateOk; LooksOk=($name -ne '' -and $sign -ne '') }
}

function Confirm-SignatureInput { param([string]$Text)
    $info = Test-SignatureFormat $Text
    $hint = @()
    if (-not $info.Name) { $hint += '• Namn verkar saknas' }
    if (-not $info.Sign) { $hint += '• Signatur verkar saknas' }
    $msg = "Har du skrivit korrekt 'Print Full Name, Sign, and Date'?
Text: $($info.Raw)
Tolkning:
  • Namn   : $($info.Name)
  • Sign   : $($info.Sign)
  • Datum  : $($info.Date)
" + ($(if ($hint.Count){ "Obs:`n  " + ($hint -join "`n  ") } else { "Ser bra ut." }))
    $res = [System.Windows.Forms.MessageBox]::Show($msg, "Bekräfta signatur", 'YesNo', 'Question')
    return ($res -eq 'Yes')
}

function Normalize-Signature {
    param([string]$s)
    if (-not $s) { return '' }
    $x = $s.Trim().ToLowerInvariant()
    $x = [regex]::Replace($x, '\s+', ' ')
    $x = $x -replace '\s*,\s*', ','
    return $x
}
 
function Get-SignatureSetForDataSheets {
    param([OfficeOpenXml.ExcelPackage]$Pkg)
    $result = [pscustomobject]@{
        RawFirst  = $null
        NormSet   = New-Object 'System.Collections.Generic.HashSet[string]'
        Occ       = @{}
        RawByNorm = @{} 
    }
    if (-not $Pkg) { return $result }

    foreach ($ws in ($Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' })) {
        $h3 = ($ws.Cells['H3'].Text + '').Trim()
        if ($h3 -match '^[0-9]') {
            $raw = ($ws.Cells['B47'].Text + '').Trim()
            if ($raw) {
                $norm = Normalize-Signature $raw
                [void]$result.NormSet.Add($norm)
                if (-not $result.RawFirst) { $result.RawFirst = $raw }
                if (-not $result.Occ.ContainsKey($norm)) {
                    $result.Occ[$norm] = New-Object 'System.Collections.Generic.List[string]'
                }
                if (-not $result.RawByNorm.ContainsKey($norm)) {
                    $result.RawByNorm[$norm] = $raw
                }
                [void]$result.Occ[$norm].Add($ws.Name)
            }
        } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/?A|NA|Tomt( innehåll)?)$') {
            break
        }
    }
    return $result
}

function UrlEncode([string]$s){ try { [System.Uri]::EscapeDataString($s) } catch { $s } }

function Get-BatchNumberFromSealFile([string]$Path){
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    if (-not (Load-EPPlus)) { return $null }
    $pkg = $null
    try {
        $pkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($Path))
        foreach ($ws in (Get-DataSheets $pkg)) {
            $txt = ($ws.Cells['D2'].Text + '').Trim()   # "Batch Number"
            if ($txt) { return $txt }
        }
    } catch {
        Gui-Log "⚠️ Get-BatchNumberFromSealFile: $($_.Exception.Message)" 'Warn'
    } finally { if ($pkg) { try { $pkg.Dispose() } catch {} } }
    return $null
}

function Update-BatchLink {
    try {
        $selNeg = Get-CheckedFilePath $clbNeg
        $selPos = Get-CheckedFilePath $clbPos
        $bnNeg  = if ($selNeg) { Get-BatchNumberFromSealFile $selNeg } else { $null }
        $bnPos  = if ($selPos) { Get-BatchNumberFromSealFile $selPos } else { $null }
        $lsp    = $txtLSP.Text.Trim()
        $mismatch = ($bnNeg -and $bnPos -and ($bnNeg -ne $bnPos))
        if ($mismatch) {
            $slBatchLink.Text        = 'SharePoint: mismatch'
            $slBatchLink.Enabled     = $false
            $slBatchLink.Tag         = $null
            $slBatchLink.ToolTipText = "NEG: $bnNeg  |  POS: $bnPos"
            return
        }
        $batch = if ($bnPos) { $bnPos } elseif ($bnNeg) { $bnNeg } else { $null }
        if ($batch) {
            $url = $SharePointBatchLinkTemplate -replace '\{BatchNumber\}', (UrlEncode $batch) -replace '\{LSP\}', (UrlEncode $lsp)

            $slBatchLink.Text        = "SharePoint: $batch"
            $slBatchLink.Enabled     = $true
            $slBatchLink.Tag         = $url
            $slBatchLink.ToolTipText = $url
        } else {
            $slBatchLink.Text        = 'SharePoint: —'
            $slBatchLink.Enabled     = $false
            $slBatchLink.Tag         = $null
            $slBatchLink.ToolTipText = 'Direktlänk aktiveras när Batch# hittas i sökt LSP.'
        }
    } catch {
        Gui-Log "⚠️ Update-BatchLink: $($_.Exception.Message)" 'Warn'
    }
}
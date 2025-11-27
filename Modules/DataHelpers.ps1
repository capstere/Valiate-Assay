function Ensure-EPPlus {
    param(
        [string] $Version = "4.5.3.3",
        [string] $SourceDllPath = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Modules\EPPlus\EPPlus.4.5.3.3\.5.3.3\lib\net35\EPPlus.dll",
        [string] $LocalFolder = "$env:TEMP\EPPlus"
    )
    try {
        $candidatePaths = @()
        if ($SourceDllPath) { $candidatePaths += $SourceDllPath }
        $localScriptDll = Join-Path $PSScriptRoot 'EPPlus.dll'
        $candidatePaths += $localScriptDll
        $userModRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'
        if (Test-Path $userModRoot) {
            Get-ChildItem -Path (Join-Path $userModRoot 'EPPlus') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $candidatePaths += Join-Path $_.FullName 'lib\net45\EPPlus.dll'
                $candidatePaths += Join-Path $_.FullName 'lib\net40\EPPlus.dll'
            }
        }

        $progFiles = $env:ProgramFiles
        $systemModRoot = Join-Path $progFiles 'WindowsPowerShell\Modules'
        if (Test-Path $systemModRoot) {
            Get-ChildItem -Path (Join-Path $systemModRoot 'EPPlus') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $candidatePaths += Join-Path $_.FullName 'lib\net45\EPPlus.dll'
                $candidatePaths += Join-Path $_.FullName 'lib\net40\EPPlus.dll'
            }
        }
        foreach ($cand in $candidatePaths) {
            if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -LiteralPath $cand)) { return $cand }
        }

        $nugetUrl = "https://www.nuget.org/api/v2/package/EPPlus/$Version"
        try {
            $guid = [Guid]::NewGuid().ToString()
            $tempDir = Join-Path $env:TEMP "EPPlus_$guid"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            $zipPath  = Join-Path $tempDir 'EPPlus.zip'
            $reqParams = @{ Uri = $nugetUrl; OutFile = $zipPath; UseBasicParsing = $true; Headers = @{ 'User-Agent' = 'DocMerge/1.0' } }
            Invoke-WebRequest @reqParams -ErrorAction Stop | Out-Null 
            if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'System.IO.Compression.FileSystem' })) {
                Add-Type -AssemblyName 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
            }
            [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $tempDir)
            $extractedRoot = Join-Path $tempDir 'lib'
            if (Test-Path $extractedRoot) {
                $dllCandidates = Get-ChildItem -Path (Join-Path $extractedRoot 'net45'), (Join-Path $extractedRoot 'net40') -Filter 'EPPlus.dll' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($dllCandidates) {
                    try {
                        if (-not (Test-Path -LiteralPath $localScriptDll)) {
                            Copy-Item -Path $dllCandidates.FullName -Destination $localScriptDll -Force -ErrorAction SilentlyContinue
                        }
                    } catch { Gui-Log "⚠️ Kunde inte kopiera EPPlus.dll: $($_.Exception.Message)" 'Warn' }
                    return $localScriptDll
                }
            }
        } catch {
            Gui-Log "❌ EPPlus: Kunde inte hämta EPPlus ($Version): $($_.Exception.Message)" 'Error'
        }
    } catch {
        Gui-Log "❌ Ensure-EPPlus fel: $($_.Exception.Message)" 'Error'
    }
    Gui-Log "❌ EPPlus.dll hittades inte. Installera EPPlus $Version manuellt." 'Error'
    return $null
}

function Load-EPPlus {
    if ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'EPPlus' }) { return $true }
    $dllPath = Ensure-EPPlus -Version '4.5.3.3'
    if (-not $dllPath) { return $false }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($dllPath)
        [System.Reflection.Assembly]::Load($bytes) | Out-Null
        return $true
    } catch {
        Gui-Log "❌ EPPlus-fel: $($_.Exception.Message)" 'Error'
        return $false
    }
}

function Set-RowBorder {
    param ($ws, [int] $row, [int] $firstRow, [int] $lastRow)
    foreach ($col in 'B','C','D','E','F','G','H') {
        $ws.Cells["$col$row"].Style.Border.Left.Style   = "None"
        $ws.Cells["$col$row"].Style.Border.Right.Style  = "None"
        $ws.Cells["$col$row"].Style.Border.Top.Style    = "None"
        $ws.Cells["$col$row"].Style.Border.Bottom.Style = "None"
    }
    $ws.Cells["B$row"].Style.Border.Left.Style  = "Medium"
    $ws.Cells["H$row"].Style.Border.Right.Style = "Medium"
    foreach ($col in 'B','C','D','E','F','G') { $ws.Cells["$col$row"].Style.Border.Right.Style = "Thin" }
    $topStyle = if ($row -eq $firstRow) { "Medium" } else { "Thin" }
    $bottomStyle = if ($row -eq $lastRow)  { "Medium" } else { "Thin" }
    foreach ($col in 'B','C','D','E','F','G','H') {
        $ws.Cells["$col$row"].Style.Border.Top.Style    = $topStyle
        $ws.Cells["$col$row"].Style.Border.Bottom.Style = $bottomStyle
    }
}

function Style-Cell { param($cell,$bold,$bg,$border,$fontColor)
    if ($bold) { $cell.Style.Font.Bold = $true }
    if ($bg)   { $cell.Style.Fill.PatternType = "Solid"; $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$bg")) }
    if ($fontColor) { $cell.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$fontColor")) }
    if ($border) { $cell.Style.Border.Top.Style=$border; $cell.Style.Border.Bottom.Style=$border; $cell.Style.Border.Left.Style=$border; $cell.Style.Border.Right.Style=$border }
}

function Test-FileLocked { param([Parameter(Mandatory=$true)][string]$Path)
    try { $fs = [IO.File]::Open($Path,'Open','ReadWrite','None'); $fs.Close(); return $false } catch { return $true }
}

function Get-CsvDelimiter { param([string]$Path)
    try {
        $first = Get-Content -LiteralPath $Path -Encoding Default -TotalCount 30 | Where-Object { $_ -and $_.Trim() } | Select-Object -First 1
        if (-not $first) { return ';' }
        $sc = ($first -split ';').Count; $cc = ($first -split ',').Count
        if ($cc -gt $sc -and $cc -ge 2) { return ',' } else { return ';' }
    } catch {
        Gui-Log "⚠️ Get-CsvDelimiter misslyckades: $($_.Exception.Message)" 'Warn'
        return ';'
    }
}

function New-TextFieldParser { param([string]$Path,[string]$Delimiter)
    try {
        $tp = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, [System.Text.Encoding]::Default)
        $tp.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $tp.SetDelimiters($Delimiter)
        $tp.HasFieldsEnclosedInQuotes = $true
        $tp.TrimWhiteSpace = $true
        return $tp
    } catch {
        Gui-Log "⚠️ New-TextFieldParser: $($_.Exception.Message)" 'Warn'
        return $null
    }
}

function Get-AssayFromCsv { param([string]$Path,[int]$StartRow=10)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $tp = $null; $delim=Get-CsvDelimiter $Path; $row=0
    try {
        $tp = New-TextFieldParser -Path $Path -Delimiter $delim
        if (-not $tp) { return $null }
        while (-not $tp.EndOfData) {
            $row++; $f = $tp.ReadFields()
            if ($row -lt $StartRow) { continue }
            if (-not $f -or $f.Length -lt 1) { continue }
            $a=([string]$f[0]).Trim()
            if ($a -and $a -notmatch '^(?i)\s*assay\s*$') { return $a }
        }
    } catch {
        Gui-Log "⚠️ Get-AssayFromCsv: $($_.Exception.Message)" 'Warn'
    } finally { if ($tp){$tp.Close()} }
    return $null
}

function Import-CsvRows { param([string]$Path,[int]$StartRow=10)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $delim=Get-CsvDelimiter $Path; $tp=$null; $rows=@()
    try {
        $tp = New-TextFieldParser -Path $Path -Delimiter $delim
        if (-not $tp) { return @() }
        $r=0
        while (-not $tp.EndOfData) {
            $r++; $f=$tp.ReadFields()
            if ($r -lt $StartRow) { continue }
            if (-not $f -or ($f -join '').Trim().Length -eq 0) { continue }
            $rows += ,$f
        }
    } catch {
        Gui-Log "⚠️ Import-CsvRows: $($_.Exception.Message)" 'Warn'
    } finally { if ($tp){$tp.Close()} }
    return ,@($rows)
}

function ConvertTo-CsvFields {
    param([string]$Line)
    return [regex]::Split($Line, ',(?=(?:[^"]*"[^"]*")*[^"]*$)')
}

function Get-CsvStats {
    param(
        [string]$Path,
        [string[]]$Lines
    )
    $out = [ordered]@{
        TestCount       = 0
        DupCount        = 0
        Duplicates      = @()
        LspValues       = @()
        LspOK           = $null
        InstrumentByType= [ordered]@{}
    }
    $lines = $Lines
    if (-not $lines) {
        if (-not (Test-Path -LiteralPath $Path)) { return [pscustomobject]$out }
        try { $lines = Get-Content -LiteralPath $Path } catch { Gui-Log "⚠️ Get-CsvStats: $($_.Exception.Message)" 'Warn'; return [pscustomobject]$out }
    }
    if (-not $lines -or $lines.Count -lt 8) { return [pscustomobject]$out }
    $header = ConvertTo-CsvFields $lines[7]
    $dataLines = @()
    if ($lines.Count -gt 9) { $dataLines = $lines[9..($lines.Count-1)] }
    $csv = $null
    try {
        if ($dataLines.Count -gt 0) {
            $csv = ConvertFrom-Csv -InputObject ($dataLines -join "`n") -Header $header
        }
    } catch { $csv = $null }

    $cartSnList = New-Object System.Collections.Generic.List[string]
    $lspList    = New-Object System.Collections.Generic.List[string]
    $instrList  = New-Object System.Collections.Generic.List[string] 
    if ($csv) {
        foreach ($row in $csv) {
            $cart = $row.'Cartridge S/N'
            $lsp  = $row.'Reagent Lot ID'
            $ins  = $row.'Instrument S/N'
            if (-not $cart) { try { $cart = $row.Item(3) } catch {} }
            if (-not $ins)  { try { $ins  = $row.Item(6) } catch {} }
            if ($cart) { $cartSnList.Add( ($cart + '').Trim() ) }
            if ($lsp)  { $lspList.Add(  ($lsp  + '').Trim() ) }
            if ($ins)  { $instrList.Add(($ins  + '').Trim() ) }
        }
    } else {
        foreach ($ln in $dataLines) {
            if (-not $ln -or -not $ln.Trim()) { continue }
            $f = ConvertTo-CsvFields $ln
            if ($f.Count -lt 7) { continue }
            $cartSnList.Add( ($f[3] + '').Trim() )
            $lspList.Add(    ($f[4] + '').Trim() )
            $instrList.Add(  ($f[6] + '').Trim() )
        }
    }
    $cartSn = $cartSnList | Where-Object { $_ -and $_ -ne '' }
    $out.TestCount = @($cartSn).Count
    $dups = $cartSn | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($dups) {
        $out.DupCount   = $dups.Count
        $out.Duplicates = $dups | ForEach-Object { "$($_.Name) x$($_.Count)" }
    }
    $lspClean = $lspList | ForEach-Object {
        $m = [regex]::Match($_, '(?<!\d)(\d{5})(?!\d)')
        if ($m.Success) { $m.Groups[1].Value } else { $null }
    } | Where-Object { $_ } | Select-Object -Unique
    $out.LspValues = $lspClean
    if ($lspClean.Count -gt 0) {
        $out.LspOK = ($lspClean.Count -eq 1)
    }
    $lut = @{}
    foreach ($k in $script:GXINF_Map.Keys) {
        $codes = $script:GXINF_Map[$k].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        foreach ($code in $codes) { $lut[$code] = $k }
    }
    foreach ($ins in ($instrList | Where-Object { $_ })) {
        $t = $null
        if ($lut.ContainsKey($ins)) { $t = $lut[$ins] } else { $t = 'Unknown' }
        if (-not $out.InstrumentByType.Contains($t)) { $out.InstrumentByType[$t] = 0 }
        $out.InstrumentByType[$t]++
    }
    return [pscustomobject]$out
}

function Format-SpPresenceGrandTotalStrict {
    param([hashtable]$Counts) 
    if (-not $Counts -or $Counts.Count -eq 0) { return '—' }
    $present = @(
        $Counts.GetEnumerator() |
        Where-Object { $_.Value -gt 0 } |
        ForEach-Object { [int]$_.Key } |
        Sort-Object
    )
    if ($present.Count -eq 0) { return '—' }
    $parts = New-Object System.Collections.Generic.List[string]
    $i = 0
    while ($i -lt $present.Count) {
        $start = $present[$i]; $end = $start
        $j = $i + 1
        while ($j -lt $present.Count -and $present[$j] -eq $end + 1) {
            $end = $present[$j]; $j++
        }
        if ($start -eq $end) { $parts.Add( ('#{0:00}' -f $start) ) }
        else { $parts.Add( ('#{0:00}-{1:00}' -f $start, $end) ) }
        $i = $j
    }
    $total = (
        $Counts.GetEnumerator() |
        Where-Object { $_.Value -gt 0 } |
        Measure-Object -Property Value -Sum
    ).Sum
    return ( ($parts -join '+') + (' x{0}' -f $total) )
}

if (Get-Command Get-InfinitySpFromCsvStrict -ErrorAction SilentlyContinue) {
    Remove-Item Function:\Get-InfinitySpFromCsvStrict -ErrorAction SilentlyContinue
}

function Get-InfinitySpFromCsvStrict {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Alias('InfinitySerials')][string[]]$InstrumentSerials,
        [string[]]$Lines
    )
    if (-not $Lines -and -not (Test-Path -LiteralPath $Path)) { return '—' }
    $useConvertFn = $false
    try { $useConvertFn = [bool](Get-Command ConvertTo-CsvFields -ErrorAction Stop) } catch {}
    function Split-CsvSmart([string]$ln) {
        if ($ln -like '*;*' -and $ln -notlike '*,*') {
            return [regex]::Split($ln, ';(?=(?:[^"]*"[^"]*")*[^"]*$)')
        } else {
            return [regex]::Split($ln, ',(?=(?:[^"]*"[^"]*")*[^"]*$)')
        }
    }

    $serSet = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
    foreach ($s in ($InstrumentSerials | Where-Object { $_ })) { $null = $serSet.Add( ($s + '').Trim().Trim('"') ) }
    if ($serSet.Count -eq 0) { return '—' } 

    $lines = $Lines
    if (-not $lines) {
        try { $lines = Get-Content -LiteralPath $Path } catch { Gui-Log "⚠️ Get-InfinitySpFromCsvStrict: $($_.Exception.Message)" 'Warn'; return '—' }
    }
    if (-not $lines -or $lines.Count -lt 10) { return '—' }

    $headerIndex = 7
    $dataStart   = 9

    $hdr = if ($useConvertFn) { ConvertTo-CsvFields $lines[$headerIndex] } else { Split-CsvSmart $lines[$headerIndex] }
    $colCount = $hdr.Count

    $idxInstr = -1
    for ($i=0; $i -lt $colCount; $i++) {
        $h = (($hdr[$i] + '') -replace '[\uFEFF\u200B]','').Trim().ToLowerInvariant()
        if ($h -match 'instrument' -and ($h -match 's/?n' -or $h -match 'serial')) { $idxInstr = $i; break }
    }
    if ($idxInstr -lt 0) { $idxInstr = 6 }  # G
    $idxSample = 2
    $counts = @{}

    for ($r = $dataStart; $r -lt $lines.Count; $r++) {
        $ln = $lines[$r]; if (-not $ln -or -not $ln.Trim()) { continue }
        $f = if ($useConvertFn) { ConvertTo-CsvFields $ln } else { Split-CsvSmart $ln }
        if ($f.Count -le [Math]::Max($idxInstr,$idxSample)) { continue }

        $instr  = ($f[$idxInstr] + '').Trim().Trim('"')
        if (-not $serSet.Contains($instr)) { continue }  # inte Infinity → skippa

        $sample = ($f[$idxSample] + '').Trim().Trim('"')
        if (-not $sample) { continue }

        $m = [regex]::Match($sample, '_(\d{2})_')
         if ($m.Success) {
             $nRaw = 0
             if ([int]::TryParse($m.Groups[1].Value, [ref]$nRaw)) {

                $sp = $nRaw 
                if ($sp -ge 0 -and $sp -le 10) {
                     if (-not $counts.ContainsKey($sp)) { $counts[$sp] = 0 }
                     $counts[$sp]++
                 }
             }
         }
    }

    if ($counts.Count -eq 0) { return '—' }
    return (Format-SpPresenceGrandTotalStrict -Counts $counts)
 }

function Normalize-Assay { param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $x=$s.ToLowerInvariant(); $x=[regex]::Replace($x,'[^a-z0-9]+',' '); $x=$x.Trim() -replace '\s+',' '; return $x
}
$AssayMap = @(

    @{ Tab='MTB ULTRA';            Aliases=@('Xpert MTB-RIF Ultra') }
    @{ Tab='MTB RIF';              Aliases=@('Xpert MTB-RIF Assay G4') }
    @{ Tab='MTB JP';               Aliases=@('Xpert MTB-RIF JP IVD') }
    @{ Tab='MTB XDR';              Aliases=@('Xpert MTB-XDR') }
    @{ Tab='FLUVID | FLUVID+';     Aliases=@('Xpress SARS-CoV-2_Flu_RSV plus','Xpert Xpress_SARS-CoV-2_Flu_RSV') }
    @{ Tab='SARS-COV-2 Plus';      Aliases=@('Xpert Xpress CoV-2 plus') }
    @{ Tab='CTNG | CTNG JP';       Aliases=@('Xpert CT_NG','Xpert CT_CE') }
    @{ Tab='C.DIFF | C.DIFF JP';   Aliases=@('Xpert C.difficile G3','Xpert C.difficile BT') }
    @{ Tab='HPV';                  Aliases=@('Xpert HPV HR','Xpert HPV v2 HR') }
    @{ Tab='HBV VL';               Aliases=@('Xpert HBV Viral Load') }
    @{ Tab='HCV VL';               Aliases=@('Xpert HCV Viral Load','Xpert_HCV Viral Load') }
    @{ Tab='HCV VL FS';            Aliases=@('Xpert HCV VL Fingerstick') }
    @{ Tab='HIV VL';               Aliases=@('Xpert HIV-1 Viral Load','Xpert_HIV-1 Viral Load') }
    @{ Tab='HIV VL XC';            Aliases=@('Xpert HIV-1 Viral Load XC') }
    @{ Tab='HIV QA';               Aliases=@('Xpert HIV-1 Qual','Xpert_HIV-1 Qual') }
    @{ Tab='HIV QA XC';            Aliases=@('Xpert HIV-1 Qual XC PQC','Xpert HIV-1 Qual XC') }
    @{ Tab='SARS-COV-2';           Aliases=@('Xpert Xpress SARS-CoV-2 CE-IVD','Xpert Xpress SARS-CoV-2') }
    @{ Tab='FLU RSV';              Aliases=@('Xpert Xpress Flu-RSV','Xpress Flu IPT_EAT off') }
    @{ Tab='MRSA SA';              Aliases=@('Xpert SA Nasal Complete G3','Xpert MRSA-SA SSTI G3') }
    @{ Tab='MRSA NxG';             Aliases=@('Xpert MRSA NxG') }
    @{ Tab='NORO';                 Aliases=@('Xpert Norovirus') }
    @{ Tab='VAN AB';               Aliases=@('Xpert vanA vanB') }
    @{ Tab='GBS';                  Aliases=@('Xpert GBS LB XC','Xpert Xpress GBS','Xpert Xpress GBS US-IVD') }
    @{ Tab='STREP A';              Aliases=@('Xpert Xpress Strep A') }
    @{ Tab='CARBA R';              Aliases=@('Xpert Carba-R','Xpert_Carba-R') }
    @{ Tab='EBOLA';                Aliases=@('Xpert Ebola EUA','Xpert Ebola CE-IVD') }
    @{ Tab='Respiratory Panel';    Aliases=@('Respiratory Panel IUO') }

)

$AssayIndex = @{}
foreach($row in $AssayMap){ foreach($a in $row.Aliases){ $k=Normalize-Assay $a; if($k -and -not $AssayIndex.ContainsKey($k)){ $AssayIndex[$k]=$row.Tab } } }

function Get-ControlTabName {
    param([string]$AssayName)
    $k = Normalize-Assay $AssayName
    if ($k -and $AssayIndex.ContainsKey($k)) { return $AssayIndex[$k] }

    if (Test-Path $SlangAssayPath) {
        try {
            $mapPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($SlangAssayPath))
            $ws = $mapPkg.Workbook.Worksheets['Slang till Assay']; if (-not $ws) { $ws = $mapPkg.Workbook.Worksheets[1] }
            if ($ws -and $ws.Dimension) {
                for ($r=2; $r -le $ws.Dimension.End.Row; $r++){
                    $sheet=$ws.Cells[$r,1].Text.Trim()
                    $aliases=@($ws.Cells[$r,2].Text,$ws.Cells[$r,3].Text,$ws.Cells[$r,4].Text) | Where-Object { $_ -and $_.Trim() }
                    foreach($al in $aliases){ if (Normalize-Assay $AssayName -eq (Normalize-Assay $al)) { $mapPkg.Dispose(); return $sheet } }
                }
            }
            $mapPkg.Dispose()
        } catch {
            Gui-Log "⚠️ Get-ControlTabName: $($_.Exception.Message)" 'Warn'
        }
    }

    return $null

}

$MinitabMap = @(

    @{ Aliases=@('Xpert MTB-RIF Ultra');                           Macro='%D12547-MTBU-SWE' }
    @{ Aliases=@('Xpert MTB-RIF Assay G4');                        Macro='%D12547-MTB-SWE' }
    @{ Aliases=@('Xpress SARS-CoV-2_Flu_RSV plus','Xpert Xpress_SARS-CoV-2_Flu_RSV'); Macro='%D12547-XP3COV2FLURSV-SWE' }
    @{ Aliases=@('Xpert Xpress CoV-2 plus');                        Macro='%D12547-XP3SARSCOV2-SWE' }
    @{ Aliases=@('CT_NG','Xpert CT_CE');                            Macro='%D12547-CTNG-SWE' }
    @{ Aliases=@('Xpert C.difficile G3','Xpert C.difficile BT');    Macro='%D12547-CDCE-SWE' }
    @{ Aliases=@('Xpert HPV HR','Xpert HPV v2 HR');                 Macro='%D12547-HPV-SWE' }
    @{ Aliases=@('Xpert HBV Viral Load');                           Macro='%D12547-HBVVL-SWE' }
    @{ Aliases=@('Xpert HCV Viral Load','Xpert_HCV Viral Load');    Macro='%D12547-HCVVL-SWE' }
    @{ Aliases=@('Xpert HCV VL Fingerstick');                       Macro='%D12547-FSHCV-SWE' }
    @{ Aliases=@('Xpert HIV-1 Viral Load','Xpert_HIV-1 Viral Load'); Macro='%D12547-HIVVL-SWE' }
    @{ Aliases=@('Xpert HIV-1 Qual','Xpert_HIV-1 Qual');            Macro='%D12547-HIVQA-SWE' }
    @{ Aliases=@('Xpert Xpress SARS-CoV-2 CE-IVD','Xpert Xpress SARS-CoV-2'); Macro='%D12547-XPRSARSCOV2-SWE' }
    @{ Aliases=@('Xpert Xpress Flu-RSV');                           Macro='%D12547-XPFLURSV-SWE' }
    @{ Aliases=@('Xpress Flu IPT_EAT off');                         Macro='%D12547-FLUNG-SWE' } 
    @{ Aliases=@('Xpert Norovirus');                                Macro='%D12547-NORO-SWE' }
    @{ Aliases=@('Xpert vanA vanB');                                Macro='%D12547-VAB-SWE' }
    @{ Aliases=@('Xpert Xpress Strep A');                           Macro='%D12547-STREPA-SWE' }
    @{ Aliases=@('Xpert Carba-R','Xpert_Carba-R');                  Macro='%D12547-CARBAR-SWE' }
    @{ Aliases=@('Xpert Ebola EUA','Xpert Ebola CE-IVD');           Macro='%D12547-EBOLA-SWE' }
    @{ Aliases=@('Xpert SA Nasal Complete G3','Xpert MRSA-SA SSTI G3'); Macro='%D12547-SACOMP-SWE' }

    # N/A-gruppen:
    @{ Aliases=@('Xpert GBS LB XC','Xpert Xpress GBS','Xpert Xpress GBS US-IVD'); Macro=$null }
    @{ Aliases=@('Xpert HIV-1 Qual XC PQC','Xpert HIV-1 Qual XC');  Macro=$null }
    @{ Aliases=@('Xpert HIV-1 Viral Load XC');                      Macro=$null }
    @{ Aliases=@('Xpert MTB-RIF JP IVD');                           Macro=$null }
    @{ Aliases=@('Xpert MTB-XDR');                                  Macro=$null }
    @{ Aliases=@('Xpert MRSA NxG');                                 Macro=$null }
)

$MinitabIndex = @{}
foreach ($row in $MinitabMap) { foreach ($a in $row.Aliases) { $k = Normalize-Assay $a; if ($k -and -not $MinitabIndex.ContainsKey($k)) { $MinitabIndex[$k] = $row.Macro } } }

function Get-MinitabMacro { param([string]$AssayName)
    if ([string]::IsNullOrWhiteSpace($AssayName)) { return $null }
    $k = Normalize-Assay $AssayName
    if ($k -and $MinitabIndex.ContainsKey($k)) { return $MinitabIndex[$k] }
    return $null
}

function Find-ObservationCol { param($ws)

    $default = 13
    if (-not $ws -or -not $ws.Dimension) { return $default }
    $maxR = [Math]::Min(5, $ws.Dimension.End.Row)
    $maxC = $ws.Dimension.End.Column

    for ($r=1; $r -le $maxR; $r++) {
        for ($c=1; $c -le $maxC; $c++) {
            $t = ($ws.Cells[$r,$c].Text + '').Trim()
            if ($t -match '^(?i)\s*(obs|observation)') { return $c }
        }
    }
    return $default
}

if (-not (Get-Command Extract-WorksheetHeader -ErrorAction SilentlyContinue)) {
    function Extract-WorksheetHeader {
        param([OfficeOpenXml.ExcelPackage]$Pkg)
        $result = [pscustomobject]@{
            PartNo         = $null
            BatchNo        = $null
            CartridgeNo    = $null
            DocumentNumber = $null
            Attachment     = $null
            Rev            = $null
            Effective      = $null
        }
        if (-not $Pkg) { return $result }
        try {
            foreach ($ws in $Pkg.Workbook.Worksheets) {
                if (-not $ws.Dimension) { continue }
                if ($ws.Name -match '(?i)for reference only') { continue }
                $left  = (($ws.HeaderFooter.OddHeader.LeftAlignedText  + '') -replace '\r?\n',' ')
                if (-not $left)  { $left  = (($ws.HeaderFooter.EvenHeader.LeftAlignedText  + '') -replace '\r?\n',' ') }
                $right = (($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' ')
                if (-not $right) { $right = (($ws.HeaderFooter.EvenHeader.RightAlignedText + '') -replace '\r?\n',' ') }
                $raw = (($left + ' ' + $right) + '').Trim()
                if (-not $raw) { continue }
                $parsed = Parse-WorksheetHeaderRaw -Raw $raw
                if (-not $result.PartNo         -and $parsed.PartNo)         { $result.PartNo         = $parsed.PartNo }
                if (-not $result.BatchNo        -and $parsed.BatchNo)        { $result.BatchNo        = $parsed.BatchNo }
                if (-not $result.CartridgeNo    -and $parsed.CartridgeNo)    { $result.CartridgeNo    = $parsed.CartridgeNo }
                if (-not $result.DocumentNumber -and $parsed.DocumentNumber) { $result.DocumentNumber = $parsed.DocumentNumber }
                if (-not $result.Attachment     -and $parsed.Attachment)     { $result.Attachment     = $parsed.Attachment }
                if (-not $result.Rev            -and $parsed.Rev)            { $result.Rev            = $parsed.Rev }
                if (-not $result.Effective      -and $parsed.Effective)      { $result.Effective      = $parsed.Effective }
            }
        } catch { Gui-Log "⚠️ Extract-WorksheetHeader: $($_.Exception.Message)" 'Warn' }
        if (-not $result.CartridgeNo) {
            foreach ($ws in $Pkg.Workbook.Worksheets) {
                if (-not $ws.Dimension) { continue }
                if ($ws.Name -match '(?i)for reference only') { continue }
                $left  = (($ws.HeaderFooter.OddHeader.LeftAlignedText  + '') -replace '\r?\n',' ').Trim()
                $right = (($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' ').Trim()
                $m  = [regex]::Match($left,  '(?<!\d)(\d{5})(?!\d)')
                $m2 = if (-not $m.Success) { [regex]::Match($right,'(?<!\d)(\d{5})(?!\d)') } else { $null }
                if ($m.Success) { $result.CartridgeNo = $m.Groups[1].Value; break }
                if ($m2 -and $m2.Success) { $result.CartridgeNo = $m2.Groups[1].Value; break }
            }
        }
        return $result
    }
}

if (-not (Get-Command Get-WorksheetHeaderPerSheet -ErrorAction SilentlyContinue)) {
    function Get-WorksheetHeaderPerSheet {
        param([OfficeOpenXml.ExcelPackage]$Pkg)

        $rxPart      = '(?i)^\s*Part\s*(?:No|Number)\.?\s*(?:\(s\))?\s*$'
        $rxBatch     = '(?i)^\s*Batch\s*(?:No|Number)(?:\s*\(s\))?\.?\s*$'
        $rxCartridge = '(?i)^\s*Cartridge\s*(?:No|Number)?\s*(?:\(?LSP\)?)?\.?\s*$'
        $rxDoc       = '(?i)^\s*Document\s*(?:No|Number|#)\s*$'
        $rxRev       = '(?i)^\s*Rev(?:ision)?\.?\s*$'
        $rxEff       = '(?i)^\s*Effective(?:\s*Date)?\s*$'
        function Get-HeaderFooterText([Object]$ws) {
            $left  = Normalize-HeaderText ((($ws.HeaderFooter.OddHeader.LeftAlignedText  + '') -replace '\r?\n',' '))
            if (-not $left)  { $left  = Normalize-HeaderText ((($ws.HeaderFooter.EvenHeader.LeftAlignedText  + '') -replace '\r?\n',' ')) }
            $right = Normalize-HeaderText ((($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' '))
            if (-not $right) { $right = Normalize-HeaderText ((($ws.HeaderFooter.EvenHeader.RightAlignedText + '') -replace '\r?\n',' ')) }
            return @($left,$right)
        }

        function Try-FromHeaderFooter([string[]]$lr, [string]$kind) {
            $left,$right = $lr
            $raw = (($left + ' ' + $right) + '').Trim()
            if (-not $raw) {
                return @{ Has = $false; Val = $null }
            }
            $parsed = Parse-WorksheetHeaderRaw -Raw $raw
            $has = $false
            $val = $null

            switch ($kind) {
                'Part' {
                    $has = ($raw -match '(?i)\bPart\b')
                    $val = $parsed.PartNo
                }
                'Batch' {
                    $has = ($raw -match '(?i)\bBatch\b')
                    $val = $parsed.BatchNo
                }
                'Cartridge' {
                    $has = ($raw -match '(?i)\bCartridge\b')
                    $val = $parsed.CartridgeNo
                }
                'Doc' {
                    $has = ($raw -match '(?i)Document\s*Number')
                    $val = $parsed.DocumentNumber
                }
                'REV' {
                    $has = ($raw -match '(?i)\bRev(?:ision)?\b')
                    $val = $parsed.Rev
                }
                'EFF' {
                    $has = ($raw -match '(?i)\bEffective(\s*Date)?\b')
                    $val = $parsed.Effective
                }
            }
            return @{ Has = [bool]$has; Val = $val }
        }

        function Try-FromCells([Object]$ws, [string]$rxLabel, [int]$maxR=1000, [int]$maxC=200, [string]$canon='') {
            $has=$false; $val=$null
            for($r=1;$r -le $maxR;$r++){
                for($c=1;$c -le $maxC;$c++){
                    $t     = Normalize-HeaderText (""+$ws.Cells[$r,$c].Text)
                    if([string]::IsNullOrWhiteSpace($t)){ continue }
                    if($t -match $rxLabel){
                        $has=$true
$right = Normalize-HeaderText (""+$ws.Cells[$r,[Math]::Min($c+1,$maxC)].Text)
$down  = Normalize-HeaderText (""+$ws.Cells[[Math]::Min($r+1,$maxR),$c].Text)
                        $val = if($right){$right}elseif($down){$down}else{''}
                        if($canon -eq 'Cartridge' -and $val){
                        elseif ($canon -eq 'EFF' -and $val) {
    $dt = $null
    if (Try-Parse-HeaderDate $val ([ref]$dt)) { $val = $dt.ToString('yyyy-MM-dd') }
}
                            $m=[regex]::Match($val,'(?<!\d)(\d{5})(?!\d)'); if($m.Success){$val=$m.Groups[1].Value}
                        }
                        return @{Has=$has; Val=$val}
                    }
                }
            }
            return @{Has=$has; Val=$val}
        }
        $rows=@()
        if(-not $Pkg){ return $rows }

foreach ($ws in $Pkg.Workbook.Worksheets) {
    if ($ws.Name -eq 'Worksheet Instructions' -or
        $ws.Name -match '(?i)for reference only') {
        continue
    }

            $partLabel=$false; $batchLabel=$false; $cartLabel=$false; $docLabel=$false; $revLabel=$false; $effLabel=$false
            $part=$null; $batch=$null; $cart=$null; $doc=$null; $rev=$null; $eff=$null
            $lr = Get-HeaderFooterText $ws
            $r1 = Try-FromHeaderFooter $lr 'Part';      $partLabel=$partLabel -or $r1.Has; if($r1.Val){$part=$r1.Val}
            $r2 = Try-FromHeaderFooter $lr 'Batch';     $batchLabel=$batchLabel -or $r2.Has; if($r2.Val){$batch=$r2.Val}
            $r3 = Try-FromHeaderFooter $lr 'Cartridge'; $cartLabel=$cartLabel -or $r3.Has; if($r3.Val){$cart=$r3.Val}
            $r4 = Try-FromHeaderFooter $lr 'Doc';       $docLabel=$docLabel -or $r4.Has; if($r4.Val){$doc=$r4.Val}
            $r5 = Try-FromHeaderFooter $lr 'REV';       $revLabel=$revLabel -or $r5.Has; if($r5.Val){$rev=$r5.Val}
            $r6 = Try-FromHeaderFooter $lr 'EFF';       $effLabel=$effLabel -or $r6.Has; if($r6.Val){$eff=$r6.Val}
            if(-not $part){  $c=Try-FromCells $ws $rxPart      1000 200 'Part';      $partLabel=$partLabel -or $c.Has; if($c.Val){$part=$c.Val} }
            if(-not $batch){ $c=Try-FromCells $ws $rxBatch     1000 200 'Batch';     $batchLabel=$batchLabel -or $c.Has; if($c.Val){$batch=$c.Val} }
            if(-not $cart){  $c=Try-FromCells $ws $rxCartridge 1000 200 'Cartridge'; $cartLabel=$cartLabel -or $c.Has; if($c.Val){$cart=$c.Val} }
            if(-not $doc){   $c=Try-FromCells $ws $rxDoc       1000 200;             $docLabel=$docLabel -or $c.Has; if($c.Val){$doc=$c.Val} }
            if(-not $rev){   $c=Try-FromCells $ws $rxRev       1000 200;             $revLabel=$revLabel -or $c.Has; if($c.Val){$rev=$c.Val} }
            if(-not $eff){   $c=Try-FromCells $ws $rxEff       1000 200 'EFF';       $effLabel=$effLabel -or $c.Has; if($c.Val){$eff=$c.Val} }
            $rows += [pscustomobject]@{
                Sheet              = $ws.Name
                PartNo             = $part
                BatchNo            = $batch
                CartridgeNo        = $cart
                DocumentNumber     = $doc
                Rev                = $rev
                Effective          = $eff
                HasPartNoLabel     = [bool]$partLabel
                HasBatchNoLabel    = [bool]$batchLabel
                HasCartridgeLabel  = [bool]$cartLabel
                HasDocumentLabel   = [bool]$docLabel
                HasRevLabel        = [bool]$revLabel
                HasEffectiveLabel  = [bool]$effLabel
            }
        }
        return $rows
    }
}

function Get-TestSummaryEquipmentFromWorksheet {
    param(
        [OfficeOpenXml.ExcelWorksheet]$Worksheet
    )

    $result = [pscustomobject]@{
        WorksheetName = $null
        Pipettes      = @()
        Instruments   = @()
    }

    if (-not $Worksheet -or -not $Worksheet.Dimension) {
        return $result
    }

    $result.WorksheetName = $Worksheet.Name

    $maxR = $Worksheet.Dimension.End.Row
    $maxC = $Worksheet.Dimension.End.Column

    # Hjälp: hämta trim:ad text
    $getCellText = {
        param(
            $Ws,
            [int]$Row,
            [int]$Col
        )
        if ($Row -le 0 -or $Col -le 0) { return '' }
        if (-not $Ws.Dimension) { return '' }
        $t = $Ws.Cells[$Row,$Col].Text
        if ($t) { return $t.Trim() }
        return ''
    }

    $pipHeaderRows = @()
    for ($r = 1; $r -le $maxR; $r++) {
        $txt = & $getCellText $Worksheet $r 1
        if (-not $txt) { continue }

        # Ex: "Table 2; Pipette Information"
        if ($txt -match '(?i)table\s*2\b.*pipett(e)?\b') {
            $pipHeaderRows += $r
        }
    }

    $pipPairs = @()  # @{ IdRow = <row>; CalRow = <row> }

    foreach ($hdrRow in $pipHeaderRows) {
        $searchMax = [Math]::Min($maxR, $hdrRow + 12)

        for ($r = $hdrRow + 1; $r -le $searchMax; $r++) {
            $label = & $getCellText $Worksheet $r 1
            if ($label -match '(?i)cepheid\s*id') {

                # Hitta "Calibration Due Date" strax under
                $calRow = $null
                for ($cRow = $r + 1; $cRow -le [Math]::Min($maxR, $r + 6); $cRow++) {
                    $lab2 = & $getCellText $Worksheet $cRow 1
                    if ($lab2 -match '(?i)calibration\s+due\s+date') {
                        $calRow = $cRow
                        break
                    }
                }

                if ($calRow) {
                    $pipPairs += [pscustomobject]@{
                        IdRow  = $r
                        CalRow = $calRow
                    }
                }
            }
        }
    }

    $pipettes = @()

    # Regex för pipett-ID: "Nr. 1", "nr 7", "no 26", "nr.3" etc.
    $pipetteIdRegex = '(?i)^\s*(?:nr|no)\.?\s*\d+\s*$'

    foreach ($pair in $pipPairs) {
        for ($c = 2; $c -le $maxC; $c += 2) {
            $idTxt = & $getCellText $Worksheet $pair.IdRow $c
            if (-not $idTxt) { continue }

            if ($idTxt -match '^(?i)N/?A$') { continue }
            if ($idTxt -notmatch $pipetteIdRegex) { continue }

            $dueCell = $Worksheet.Cells[$pair.CalRow,$c]
            $dueVal  = $dueCell.Value
            $dueOut  = $null

            if ($dueVal -is [datetime]) {
                $dueOut = $dueVal
            }
            elseif ($dueVal -is [double] -or $dueVal -is [int]) {
                try {
                    $base   = Get-Date '1899-12-30'
                    $dueOut = $base.AddDays([double]$dueVal)
                } catch {
                    $dueOut = ($dueCell.Text + '').Trim()
                }
            }
            else {
                $txt = ($dueCell.Text + '').Trim()
                if ($txt) { $dueOut = $txt }
            }

            # Normalize pipette ID to "Nr. <number>"
            $idNorm = $idTxt
            try {
                $mId = [regex]::Match($idTxt, '(?i)^\s*(?:nr|no)\.?\s*(\d+)\s*$')
                if ($mId.Success) {
                    $idNorm = "Nr. {0}" -f $mId.Groups[1].Value
                } elseif ($idNorm) {
                    $idNorm = $idNorm.Trim()
                }
            } catch {
                $idNorm = $idTxt
            }
            $pipettes += [pscustomobject]@{
                Id                 = $idNorm
                CalibrationDueDate = $dueOut
            }
        }
    }

    if ($pipettes -and $pipettes.Count -gt 0) {
        $uniq = @()
        $seen = @{}
        foreach ($p in $pipettes) {
            if (-not $p.Id) { continue }
            if ($seen.ContainsKey($p.Id)) { continue }
            $seen[$p.Id] = $true
            $uniq += $p
        }
        $pipettes = $uniq
    }

    $result.Pipettes = $pipettes

    $instHeaderRows = @()
    for ($r = 1; $r -le $maxR; $r++) {
        $txt = & $getCellText $Worksheet $r 1
        if (-not $txt) { continue }

        # Bred match: "Table 3; GeneXperts/Infinitys Used", "Table 3; GeneXperts Used:" etc.
        if ($txt -match '(?i)table\s*3\b.*(gene\s*experts?|genexperts?|genex?perts?|infinitys?)') {
            $instHeaderRows += $r
        }
    }

    $instPairs = @()

    foreach ($hdrRow in $instHeaderRows) {
        $searchMax = [Math]::Min($maxR, $hdrRow + 20)

        for ($r = $hdrRow + 1; $r -le $searchMax; $r++) {
            $label = & $getCellText $Worksheet $r 1
            if ($label -match '(?i)cepheid\s*id') {

                $calRow = $null
                for ($cRow = $r + 1; $cRow -le [Math]::Min($maxR, $r + 6); $cRow++) {
                    $lab2 = & $getCellText $Worksheet $cRow 1
                    if ($lab2 -match '(?i)calibration\s+due\s+date') {
                        $calRow = $cRow
                        break
                    }
                }

                if ($calRow) {
                    $instPairs += [pscustomobject]@{
                        IdRow  = $r
                        CalRow = $calRow
                    }
                }
            }
        }
    }

    $insts = @()

    foreach ($pair in $instPairs) {
        for ($c = 2; $c -le $maxC; $c += 2) {
            $idTxt = & $getCellText $Worksheet $pair.IdRow $c
            if (-not $idTxt) { continue }
            if ($idTxt -match '^(?i)N/?A$') { continue }

            # Bara GX/Infinity (inklusive romerska siffror t.ex. Infinity-VI)
            if ($idTxt -notmatch '(?i)^\s*(?:infinity(?:[-\s]*[ivxlcdm]+)?|gx\s*\d+|gx\d+)') {
                continue
            }

            $dueCell = $Worksheet.Cells[$pair.CalRow,$c]
            $dueVal  = $dueCell.Value
            $dueOut  = $null

            if ($dueVal -is [datetime]) {
                $dueOut = $dueVal
            }
            elseif ($dueVal -is [double] -or $dueVal -is [int]) {
                try {
                    $base   = Get-Date '1899-12-30'
                    $dueOut = $base.AddDays([double]$dueVal)
                } catch {
                    $dueOut = ($dueCell.Text + '').Trim()
                }
            }
            else {
                $txt = ($dueCell.Text + '').Trim()
                if ($txt) { $dueOut = $txt }
            }

            $insts += [pscustomobject]@{
                Id                 = $idTxt
                CalibrationDueDate = $dueOut
            }
        }
    }

    if ($insts -and $insts.Count -gt 0) {
        $uniq = @()
        $seen = @{}
        foreach ($i in $insts) {
            if (-not $i.Id) { continue }
            if ($seen.ContainsKey($i.Id)) { continue }
            $seen[$i.Id] = $true
            $uniq += $i
        }
        $insts = $uniq
    }

    $result.Instruments = $insts

    return $result
}

function Get-TestSummaryEquipment {
    param(
        [OfficeOpenXml.ExcelPackage]$Pkg
    )

    $empty = [pscustomobject]@{
        WorksheetName = $null
        Pipettes      = @()
        Instruments   = @()
    }

    if (-not $Pkg) { return $empty }

    $best = $null

    try {
        foreach ($ws in $Pkg.Workbook.Worksheets) {
            if (-not $ws.Dimension) { continue }
            if ($ws.Name -match '(?i)worksheet\s*instructions') { continue }

            $info = Get-TestSummaryEquipmentFromWorksheet -Worksheet $ws
            if (-not $info) { continue }

            $pipCount  = if ($info.Pipettes)   { $info.Pipettes.Count }   else { 0 }
            $instCount = if ($info.Instruments){ $info.Instruments.Count } else { 0 }

            if (-not $best) { $best = $info }

            if ($pipCount -gt 0 -or $instCount -gt 0) {
                if ($ws.Name -match '(?i)test\s*summary') {
                    return $info
                }
                $best = $info
            }
        }
    } catch {
        return $empty
    }

    if ($best) { return $best }
    return $empty
}

if (-not (Get-Command Compare-WorksheetHeaderSet -ErrorAction SilentlyContinue)) {
    function Compare-WorksheetHeaderSet {
        param([Parameter(Mandatory)][object[]]$Rows)

        $devIssues  = 0
        $reqIssues  = 0
        $detailsLst = New-Object System.Collections.Generic.List[string]

function _canon([string]$raw, [string]$type) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    $txt = Normalize-HeaderText $raw
            switch ($type) {
                'Part'      { $m=[regex]::Match($txt,'(?i)\b(\d{3}-\d{4})\b'); return $(if($m.Success){$m.Groups[1].Value}else{$txt.ToUpper()}) }
                'Batch'     { $m=[regex]::Match($txt,'(?<!\d)(\d{10})(?!\d)');  return $(if($m.Success){$m.Groups[1].Value}else{$txt.ToUpper()}) }
                'Cartridge' { $m=[regex]::Match($txt,'(?<!\d)(\d{5})(?!\d)');   return $(if($m.Success){$m.Groups[1].Value}else{$txt.ToUpper()}) }
                'Doc'       { $m=[regex]::Match($txt,'(?i)\b(D\d+)\b');         return $(if($m.Success){$m.Groups[1].Value.ToUpper()}else{$txt.ToUpper()}) }
                'REV'       { return $txt.ToUpper() }
                'EFF' {
    $dt = $null
    if (Try-Parse-HeaderDate $txt ([ref]$dt)) { return $dt.ToString('yyyy-MM-dd') }
    return $txt
}
                default     { return $txt }
            }
        }
        $keys = @(

            @{K='PartNo';         T='Part';      Required=$true;  Label='HasPartNoLabel'    },
            @{K='BatchNo';        T='Batch';     Required=$true;  Label='HasBatchNoLabel'   },
            @{K='CartridgeNo';    T='Cartridge'; Required=$true;  Label='HasCartridgeLabel' },
            @{K='DocumentNumber'; T='Doc';       Required=$false; Label='HasDocumentLabel'  },
            @{K='Rev';            T='REV';       Required=$false; Label='HasRevLabel'       },
            @{K='Effective';      T='EFF';       Required=$false; Label='HasEffectiveLabel' }
        )

        foreach ($entry in $keys) {
            $k=$entry.K; $t=$entry.T; $req=$entry.Required; $labelProp=$entry.Label
            $vals = foreach($r in $Rows){
                [pscustomobject]@{
                    Sheet  = $r.Sheet
                    Canon  = _canon ($r.$k) $t
                    Raw    = if([string]::IsNullOrWhiteSpace($r.$k)){ $null } else { $r.$k.Trim() }
                    HasLbl = [bool]($r.$labelProp)
                }
            }
            $nonEmpty = $vals | Where-Object { $_.Canon }
            $useSet   = if($nonEmpty.Count -gt 0){$nonEmpty}else{$vals}
            $byVal    = $useSet | Group-Object Canon | Sort-Object Count -Descending
            $major    = if($byVal.Count -gt 0){ $byVal[0].Name } else { $null }
            $deviants = @()
            if($major){
                $deviants = $vals | Where-Object { $_.Canon -and ($_.Canon -ne $major) } | Select-Object -ExpandProperty Sheet -Unique
            }

            if($deviants.Count -gt 0){
                $devIssues++
                $majShow = if($major){"'$major'"}else{'(empty)'}
                $line = "- $k majoritet=$majShow | avvikande flikar: " + ($deviants -join ', ')
                [void]$detailsLst.Add($line)
            }

            if($req){
                $missingSheets = $vals | Where-Object { $_.HasLbl -and -not $_.Canon } | Select-Object -ExpandProperty Sheet -Unique
                if($missingSheets.Count -gt 0){
                    $reqIssues++
                    $line = "*Saknas* - $k " + ($missingSheets -join ', ')
                    [void]$detailsLst.Add($line)
                }
            }
        }

        $issuesTotal = $devIssues + $reqIssues
        $summary = if($issuesTotal -eq 0){ 'OK' } else { "Avvikelser: $issuesTotal (avvikare=$devIssues, saknas=$reqIssues)" }

        return [pscustomobject]@{
            Issues  = $issuesTotal
            Summary = $summary
            Details = ($detailsLst -join "`r`n")

        }
    }
}

if (-not (Get-Command Extract-SealTestHeader -ErrorAction SilentlyContinue)) {

    function Extract-SealTestHeader {
        param([OfficeOpenXml.ExcelPackage]$Pkg)
        $result = [pscustomobject]@{ DocumentNumber=$null; Rev=$null; Effective=$null }
        if (-not $Pkg) { return $result }

        foreach ($ws in $Pkg.Workbook.Worksheets) {

            if ($ws.Name -eq 'Worksheet Instructions') { continue }
            $right = (($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' ').Trim()

            if (-not $right) { $right = (($ws.HeaderFooter.EvenHeader.RightAlignedText + '') -replace '\r?\n',' ').Trim() }
            if (-not $result.DocumentNumber -and $right -match '(?i)\bDocument\s*(?:No|Number|#)\s*[:#]?\s*(D\d+)\b') { $result.DocumentNumber = $matches[1] }
            if (-not $result.Rev            -and $right -match '(?i)\bRev(?:ision)?\.?\s*[:#]?\s*([A-Z]{1,3}(?:\.\d+)?)\b') { $result.Rev = $matches[1] }
            if (-not $result.Effective      -and $right -match '(?i)\bEffective\s*[:#]?\s*([0-9]{1,2}[\/\-][0-9]{1,2}[\/\-][0-9]{4}|[0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2})') { $result.Effective = $matches[1] }
            if ($result.DocumentNumber -and $result.Rev -and $result.Effective) { break }
        }
        return $result

    }

}

function Normalize-Id {

    param([string]$Value, [ValidateSet('Batch','Part','Cartridge')] [string]$Type)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $v = $Value.Trim()
    switch ($Type) {
        'Batch'     { return ($v -replace '[^\d]', '') }                      # 10 siffror
        'Part'      { return (($v -replace '[^0-9A-Za-z\-]', '')).ToUpper() } # t.ex. 700-5702
        'Cartridge' { return ($v -replace '[^\d]', '') }                      # 5 siffror
        default     { return $v }

    }

}

function Normalize-HeaderText {

    param([string]$s)

    if ([string]::IsNullOrEmpty($s)) { return '' }
    # Ta bort BOM/zero-width

    $s = $s -replace "[\uFEFF\u200B]", ""
    # Ta bort inbäddade radbrytningsmarkörer från Excel-export

    # t.ex. 700-4370_x000D_ Batch No(s).: 1001502746_x000D_ ...
    $s = $s -replace "_x000D_", " "

    # NBSP/figur/narrow NBSP → vanligt mellanslag
    $s = $s -replace "[\u00A0\u2007\u202F]", " "

    # Normalisera olika streck till '-'
    $s = $s -replace "[\u2013\u2014\u2212]", "-"

    # Kollapsa whitespace och trimma
    $s = ($s -replace "\s+", " ").Trim()
    return $s

}

function Parse-WorksheetHeaderRaw {
    param(
        [string]$Raw
    )

    $obj = [pscustomobject]@{

        PartNo         = $null
        BatchNo        = $null
        CartridgeNo    = $null
        DocumentNumber = $null
        Attachment     = $null
        Rev            = $null
        Effective      = $null
    }

    if (-not $Raw) { return $obj }

    # Grundstädning
    $txt = $Raw + ''

    # Ta bort header-koder (sidnr m.m.)
    $txt = $txt -replace '&P', ''
    $txt = $txt -replace '&N', ''
    $txt = $txt -replace '&L', ' '

    # Normalisera (tar också bort _x000D_)
    $txt = Normalize-HeaderText $txt

    # Document Number + ev. Attachment
    $m = [regex]::Match(
        $txt,
        '(?i)\bDocument\s*Number\b\s*[:#]?\s*(D\d{5})(?:\s+Attachment[:\s]+([0-9A-Za-z\.]+))?'
    )

    if ($m.Success) {
        $obj.DocumentNumber = $m.Groups[1].Value.Trim()
        if ($m.Groups.Count -ge 3 -and $m.Groups[2].Value) {
            $obj.Attachment = $m.Groups[2].Value.Trim()
        }
    }

    # Fristående "Attachment" om det inte fångades ovan
    if (-not $obj.Attachment) {
        $m = [regex]::Match($txt, '(?i)\bAttachment\b\s*:?\s*([0-9A-Za-z\.]+)')
        if ($m.Success) {
            $obj.Attachment = $m.Groups[1].Value.Trim()
        }
    }

    # Rev (t.ex. F, AW, E.1)
    $m = [regex]::Match($txt, '(?i)\bRev(?:ision)?\b\s*:?\s*([A-Z0-9]+(?:\.[0-9]+)?)')
    if ($m.Success) {
        $obj.Rev = $m.Groups[1].Value.Trim()
    }

    # Effective (t.ex. 4/14/2025 eller 09/20/2023)
    $m = [regex]::Match($txt, '(\d{1,2}/\d{1,2}/\d{4})')
    if ($m.Success) {
        $obj.Effective = $m.Groups[1].Value.Trim()

    }

    # Part No (t.ex. 700-4370)
    $m = [regex]::Match($txt, '(?i)Part\s*No\.\s*:\s*([0-9A-Z\-]+)')
    if ($m.Success) {
        $obj.PartNo = $m.Groups[1].Value.Trim()
    }

    # Batch No (t.ex. 1001502746)
    $m = [regex]::Match($txt, '(?i)Batch\s*No\(s\)\.\s*:\s*(\d{6,12})')
    if ($m.Success) {
        $obj.BatchNo = $m.Groups[1].Value.Trim()

    }

    # Cartridge (LSP, t.ex. 78704)
    $m = [regex]::Match($txt, '(?i)Cartridge\s*No\.\s*\(LSP\)\s*:\s*(\d{5,7})')
    if ($m.Success) {
        $obj.CartridgeNo = $m.Groups[1].Value.Trim()

    }

    return $obj

}

function Try-Parse-HeaderDate {
    param([object]$cellValue, [ref]$out)
    $out.Value = $null
    if ($null -eq $cellValue) { return $false }

 

    if ($cellValue -is [double] -or $cellValue -is [int]) {
        try { $out.Value = [DateTime]::FromOADate([double]$cellValue); return $true } catch {}

    }

    if ($cellValue -is [DateTime]) { $out.Value = [DateTime]$cellValue; return $true }

    $s = Normalize-HeaderText ([string]$cellValue)
    if ([string]::IsNullOrWhiteSpace($s)) { return $false }

    $cultures = @('sv-SE','en-US','en-GB')
    foreach ($c in $cultures) {
        try {
            $dt = [DateTime]::Parse($s, [Globalization.CultureInfo]::GetCultureInfo($c), [Globalization.DateTimeStyles]::AssumeLocal)
            $out.Value = $dt; return $true
        } catch {}
    }

    $formats = @(
        'yyyy-MM-dd','yyyy/MM/dd','dd/MM/yyyy','d/M/yyyy','MM/dd/yyyy','M/d/yyyy',
        'dd-MM-yyyy','d-M-yyyy','dd.M.yyyy','d.M.yyyy','dd-MMM-yyyy','MMM-yy'
    )

    foreach ($fmt in $formats) {
        try {
            $dt = [DateTime]::ParseExact($s, $fmt, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal)
            $out.Value = $dt; return $true
        } catch {}
    }
    return $false
}

function Get-ConsensusValue {
    param(
        [string]$Ws,  [string]$Pos, [string]$Neg,
        [ValidateSet('Batch','Part','Cartridge')] [string]$Type
    )

    $nWs  = Normalize-Id $Ws  $Type
    $nPos = Normalize-Id $Pos $Type
    $nNeg = Normalize-Id $Neg $Type

    $present = @{}
    if ($nWs)  { $present['WS']  = $nWs  }
    if ($nPos) { $present['POS'] = $nPos }
    if ($nNeg) { $present['NEG'] = $nNeg }

    $chosenNorm = $null
    $sources    = @()

    $counts = @{}
    foreach ($k in $present.Keys) {

        $val = $present[$k]
        if (-not $counts.ContainsKey($val)) { $counts[$val] = 0 }
        $counts[$val]++

    }

    $maxCount = 0; $maxVal = $null
    foreach ($kv in $counts.GetEnumerator()) {
        if ($kv.Value -gt $maxCount) { $maxCount = $kv.Value; $maxVal = $kv.Key }

    }

    if ($maxCount -ge 2) {
        $chosenNorm = $maxVal
        foreach ($k in $present.Keys) { if ($present[$k] -eq $chosenNorm) { $sources += $k } }

    }

    if (-not $chosenNorm -and $nWs) {

        $posMatch = ($nPos -and $nPos -eq $nWs)
        $negMatch = ($nNeg -and $nNeg -eq $nWs)

        if (($posMatch -and -not $negMatch) -or ($negMatch -and -not $posMatch)) {
            $chosenNorm = $nWs
            $sources = @('WS')
            if ($posMatch) { $sources += 'POS' }
            if ($negMatch) { $sources += 'NEG' }
        }
    }

    if (-not $chosenNorm -and $nPos -and $nNeg -and $nPos -eq $nNeg) {
        $chosenNorm = $nPos
        $sources = @('POS','NEG')

    }

    if (-not $chosenNorm) {
        if ($nWs)      { $chosenNorm = $nWs;  $sources = @('WS')  }
        elseif ($nPos) { $chosenNorm = $nPos; $sources = @('POS') }
        elseif ($nNeg) { $chosenNorm = $nNeg; $sources = @('NEG') }
    }

    $orig = @{ WS=$Ws; POS=$Pos; NEG=$Neg }
    $pretty = $null
    foreach ($pref in @('WS','POS','NEG')) {

        if ($present.ContainsKey($pref) -and $present[$pref] -eq $chosenNorm) { $pretty = $orig[$pref]; break }
    }

 

    $note = $null

    if ($sources -and $sources.Count -gt 0) {
        $note = "Consensus: " + ($sources -join '+')
        $others = @()
        foreach ($k in @('WS','POS','NEG')) {

            if ($present.ContainsKey($k) -and ($sources -notcontains $k)) {

                $val = $orig[$k]

                if (-not [string]::IsNullOrEmpty($val)) { $others += ($k + '=' + $val) }
            }
        }
        if ($others.Count -gt 0) { $note += " | Others: " + ($others -join ', ') }
    }
    return @{ Value=$pretty; Source=($sources -join '+'); Note=$note }
}

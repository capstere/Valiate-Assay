param(
    [string]$ScriptRoot = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

# === Inställningar ===
$ScriptVersion = "v45.1.0"

$RootPaths = @(
    'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Tests',
    'N:\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING',
    'N:\QC\QC-1\IPT\4. IPT - KLART FÖR GRANSKNING'
)

$ikonSokvag = Join-Path $ScriptRoot "icon.png"
$UtrustningListPath = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project v2\Utrustninglista5.0.xlsx"
$RawDataPath        = "N:\QC\QC-1\IPT\KONTROLLPROVSFIL - Version 2.4.xlsm"
$SlangAssayPath     = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project\Click Less Project v2\slangassay.xlsx"

$OtherScriptPath = ''

$Script1Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Kontrollprovsfil 2025\Script Raw Data\Kontrollprovsfil_EPPlus_2025.ps1'
$Script2Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project v2\rename-GUI.bat'
$Script3Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project v2\rename-GUI.bat'

$env:PNPPOWERSHELL_UPDATECHECK = "Off"
$global:SP_ClientId   = "INSERT MYSELF"
$global:SP_Tenant     = "danaher.onmicrosoft.com"
$global:SP_CertBase64 = "INSERT MYSELF"
$global:SP_SiteUrl    = "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management"


# === Centraliserad konfiguration ===
$Config = [ordered]@{
    CsvPath       = ''           # Sökväg till CSV-fil
    SealNegPath   = ''           # Sökväg till Seal Test NEG
    SealPosPath   = ''           # Sökväg till Seal Test POS
    WorksheetPath = ''           # Sökväg till LSP worksheet (Worksheet.xlsx)
    SiteUrl      = $global:SP_SiteUrl
    Tenant       = $global:SP_Tenant
    ClientId     = $global:SP_ClientId
    Certificate  = $global:SP_CertBase64
    EpplusDllPath = (Join-Path $ScriptRoot 'EPPlus.dll')
    EpplusVersion = '4.5.3.3'
}

$script:GXINF_Map = @{
    'Infinity-VI'   = '847922'
    'Infinity-VIII' = '803094'
    'GX5'           = '750210,750211,750212,750213'
    'GX6'           = '750246,750247,750248,750249'
    'GX1'           = '709863,709864,709865,709866'
    'GX2'           = '709951,709952,709953,709954'
    'GX3'           = '710084,710085,710086,710087'
    'GX7'           = '750170,750171,750172,750213'
    'Infinity-I'    = '802069'
    'Infinity-III'  = '807363'
    'Infinity-V'    = '839032'
}

$SharePointBatchLinkTemplate = 'https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/ROBAL.aspx?viewid=6c9e53c9-a377-40c1-a154-13a13866b52b&view=7&q={BatchNumber}'

# --------------------------------------------------------------------------
# Kontrollmaterial-karta och rapportinställningar
# --------------------------------------------------------------------------
# Standardväg till kontrollmaterial-kartan (Excel-fil). Justera vid behov
$Global:ControlMaterialMapPath = Join-Path $ScriptRoot 'ControlMaterialMap_SE.xlsx'

# Rapportinställningar (toggles) som styr vilka sektioner som skrivs ut samt om
# detaljer för kontrollmaterial ska visas. Dessa kan ändras av användaren eller
# konfigureras här. Alla värden är booleans.
$Global:ReportOptions = [ordered]@{
    # Inkludera listning av saknade replikat (sektion D)
    IncludeMissingReplicates    = $true
    # Inkludera dubbletter (sektion F)
    IncludeDuplicates           = $true
    # Inkludera instrumentfel (sektion G)
    IncludeInstrumentErrors     = $true
    # Inkludera detaljerad information för kontrollmaterial (namn, kategori, källa)
    IncludeControlDetails       = $true
    # När sant, visas endast rader i kontrollmaterialssektionen som avviker
    HighlightMismatchesOnly     = $false
}

$DevLogDir = Join-Path $ScriptRoot 'Loggar'
if (-not (Test-Path $DevLogDir)) { New-Item -ItemType Directory -Path $DevLogDir -Force | Out-Null }
$global:LogPath = Join-Path $DevLogDir ("$($env:USERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt")

function Test-Config {
    $result = [pscustomobject]@{
        Ok       = $true
        Errors   = New-Object System.Collections.Generic.List[object]
        Warnings = New-Object System.Collections.Generic.List[object]
    }
    try {
        $templatePath = Join-Path $ScriptRoot 'output_template-v4.xlsx'
        if (-not (Test-Path -LiteralPath $templatePath)) {
            $null = $result.Errors.Add("Mallfil saknas: $templatePath")
        }
    } catch {
        $null = $result.Errors.Add("Test-Config (template): $($_.Exception.Message)")
    }
    try {
        if (-not (Test-Path -LiteralPath $UtrustningListPath)) {
            $null = $result.Warnings.Add("Utrustningslista saknas: $UtrustningListPath")
        }
    } catch {
        $null = $result.Warnings.Add("Test-Config (utrustning): $($_.Exception.Message)")
    }
    try {
        if (-not (Test-Path -LiteralPath $RawDataPath)) {
            $null = $result.Warnings.Add("Kontrollprovsfil saknas: $RawDataPath")
        }
    } catch {
        $null = $result.Warnings.Add("Test-Config (rawdata): $($_.Exception.Message)")
    }
    try {
        if (-not (Test-Path -LiteralPath $SlangAssayPath)) {
            $null = $result.Warnings.Add("Slang/assay-tabell saknas: $SlangAssayPath")
        }
    } catch {
        $null = $result.Warnings.Add("Test-Config (slang/assay): $($_.Exception.Message)")
    }
    try {
        if (-not (Test-Path -LiteralPath $Global:ControlMaterialMapPath)) {
            $null = $result.Warnings.Add("Kontrollmaterial-karta saknas: $Global:ControlMaterialMapPath")
        }
    } catch {
        $null = $result.Warnings.Add("Test-Config (control map): $($_.Exception.Message)")
    }
    try {
        if (-not (Test-Path -LiteralPath $DevLogDir)) {
            New-Item -ItemType Directory -Path $DevLogDir -Force | Out-Null
        }
        $probe = Join-Path $DevLogDir "write_probe.txt"
        Set-Content -Path $probe -Value 'probe' -Encoding UTF8 -Force
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    } catch {
        $null = $result.Warnings.Add("Kunde inte verifiera skrivning till loggmapp: $($_.Exception.Message)")
    }
    if ($result.Errors.Count -gt 0) { $result.Ok = $false }
    return $result
}

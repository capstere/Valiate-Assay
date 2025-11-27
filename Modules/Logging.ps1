function Get-OutputBox {
    try { return $script:outputBox } catch { return $null }
}

function Set-LogOutputControl {
	param([System.Windows.Forms.TextBox]$Control)
    $script:outputBox = $Control

}

function Gui-Log {
    param(
        [string] $Text,
        [ValidateSet('Info', 'Warn', 'Error')][string] $Severity = 'Info',
        [switch] $Immediate
    )
		
    $prefix = switch ($Severity) { 'Warn' { '⚠️' } 'Error' { '❌' } default { 'ℹ️' } }

    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $line = "[$timestamp] $prefix $Text"

    $append = {
        $outputBox.AppendText("$line`r`n")

        $outputBox.SelectionStart = $outputBox.TextLength
        $outputBox.ScrollToCaret()
        $outputBox.Refresh()
    }
    $outputBox = Get-OutputBox
    if ($outputBox) {
        try {
            if ($outputBox.InvokeRequired) {

                $null = $outputBox.BeginInvoke([System.Windows.Forms.MethodInvoker]$append)

            } else {
                & $append
            }
            if ($Immediate) {
							[System.Windows.Forms.Application]::DoEvents()
            }
        } catch {
            Write-Host $line
        }
    } else {
        Write-Host $line
    }

    if ($global:LogPath) {

        try { Add-Content -Path $global:LogPath -Value $line -Encoding UTF8 } catch { Write-Host "Loggning misslyckades: $($_.Exception.Message)" }
    }
}

function Write-BackendLog {
    param(
        [string]$Message,
        [ValidateSet('Info','Warn','Error')][string]$Severity = 'Info',
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    if (-not $global:LogPath) { return }

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Severity.ToUpper(), $Message

    try {
        Add-Content -Path $global:LogPath -Value $line -Encoding UTF8
        if ($ErrorRecord) {
            $errDetails = @(
                "    Exception : $($ErrorRecord.Exception.Message)",
                "    Category  : $($ErrorRecord.CategoryInfo)"
            )
            if ($ErrorRecord.ScriptStackTrace) {
                $errDetails += "    Stack     : $($ErrorRecord.ScriptStackTrace.Trim())"
            }
            Add-Content -Path $global:LogPath -Value $errDetails -Encoding UTF8
        }
    } catch {
        Write-Host "Backend-loggning misslyckades: $($_.Exception.Message)"
    }
}

function Write-BackendError {
    param(
        [string]$Context,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    if (-not $ErrorRecord) { return }
    $msg = if ($Context) { "$Context: $($ErrorRecord.Exception.Message)" } else { $ErrorRecord.Exception.Message }
    Write-BackendLog -Message $msg -Severity 'Error' -ErrorRecord $ErrorRecord
}

<#

.SYNOPSIS

Lägger till en rad i revisionslogg (CSV).

.DESCRIPTION

Skapar eller uppdaterar en audit-CSV per dag/användare/LSP. Misslyckanden loggas men stoppar inte körningen.

.PARAMETER Lsp

LSP-nummer om tillgängligt.

.PARAMETER Assay

Assay-namn om tillgängligt.

.PARAMETER BatchNumber

Batchnummer om tillgängligt.

.PARAMETER TestCount

Antal test i rapporten.

.PARAMETER Status

Kort status (OK/Warning/Error).

.PARAMETER ReportPath

Sökväg till genererad rapport.

.PARAMETER AuditDir

Valfri sökväg till audit-mapp (default: <scriptroot>/audit).

.EXAMPLE

Add-AuditEntry -Lsp 12345 -Assay 'MTB' -BatchNumber 'B001' -TestCount 10 -Status 'OK' -ReportPath 'C:\temp\out.xlsx'

.NOTES

Behåller textlogg parallellt; audit är frivillig och robust.

#>

function Add-AuditEntry {

    param(

        [string]$Lsp,
        [string]$Assay,
        [string]$BatchNumber,
        [int]$TestCount,
        [string]$Status,
        [string]$ReportPath,
        [string]$AuditDir
    )
    try {
        if (-not $AuditDir) {
            if ($PSScriptRoot) { $AuditDir = Join-Path $PSScriptRoot 'audit' }
            else { $AuditDir = Join-Path $env:TEMP 'audit' }
        }
        if (-not (Test-Path -LiteralPath $AuditDir)) { New-Item -ItemType Directory -Path $AuditDir -Force | Out-Null }
        $date = (Get-Date).ToString('yyyyMMdd')
        $safeLsp = if ($Lsp) { $Lsp } else { 'NA' }
        $file = Join-Path $AuditDir ("Audit_{0}_{1}_{2}.csv" -f $date, $env:USERNAME, $safeLsp)

        $row = [pscustomobject]@{

            Timestamp  = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Username   = $env:USERNAME
            LSP        = $Lsp
            Assay      = $Assay
            Batch      = $BatchNumber
            TestCount  = $TestCount
            Status     = if ($Status) { $Status } else { 'OK' }
            ReportPath = $ReportPath
        }
				
        $exists = Test-Path -LiteralPath $file
        $row | Export-Csv -Path $file -NoTypeInformation -Append:($exists) -Encoding UTF8
    } catch {

        Gui-Log "⚠️ Kunde inte skriva audit-fil: $($_.Exception.Message)" 'Warn'
    }
}

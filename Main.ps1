<# Main.ps1
    Main entry point for the IVDR validation system.

    This script performs the following high‑level tasks:
      1. Ensures all dependencies are installed (no admin rights required).
      2. Imports project modules for validation, equipment handling,
         SharePoint access and report generation.
      3. Defines logging utilities for traceability.
      4. Presents a simple WinForms GUI to accept user input via drag‑and‑drop
         or file dialogs.
      5. Orchestrates the validation process: extracting metadata, validating
         CSV content, gathering equipment/controls information and writing
         a colour‑coded Excel report.

    The intent is to provide a self‑contained, modular framework that can
    easily be extended.  Each module has clear responsibilities and is
    versioned separately.  Log messages are written to a timestamped file
    under the Logs directory.

    Usage:
      Simply run Run_Validation.bat or create a shortcut (.lnk) pointing to
      this script.  The GUI will guide the user through selecting files.
#>

param()

# Region: Constants
# Path to the XML rules file.  Modify if you maintain multiple versions.
$rulesPath = Join-Path -Path $PSScriptRoot -ChildPath 'AssayRules_final.xml'

# Destination directory for log files.  By default this uses a network or removable drive
# path (e.g. N:\min_destination) if configured, otherwise falls back to a 'Logs'
# subfolder in the script directory.  To override, set the environment variable
# IVDR_LOG_PATH before running the script.
$logRoot = if ($env:IVDR_LOG_PATH) { $env:IVDR_LOG_PATH } elseif (Test-Path 'N:\min_destination') { 'N:\min_destination' } else { Join-Path -Path $PSScriptRoot -ChildPath 'Logs' }
if (-not (Test-Path -LiteralPath $logRoot)) { New-Item -ItemType Directory -Path $logRoot -Force | Out-Null }
$logFile = Join-Path -Path $logRoot -ChildPath ("validation_" + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.log')

# Region: Logging
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $timestamp = (Get-Date).ToString('s')
    $line = "[$timestamp][$Level] $Message"
    Add-Content -Path $logFile -Value $line
    # Also echo to console for user feedback
    Write-Host $line
}


# Ensure dependencies
try {
    Write-Log "Installing dependencies if missing..."
    . (Join-Path -Path $PSScriptRoot -ChildPath 'Install-Dependencies.ps1')
    Install-Dependencies -ModulePath (Join-Path $PSScriptRoot 'Modules')
    # Load assemblies into current session
    Load-EPPlus -ModulePath (Join-Path $PSScriptRoot 'Modules') | Out-Null
    Load-NPOI   -ModulePath (Join-Path $PSScriptRoot 'Modules') | Out-Null
    Write-Log "Dependencies installed and assemblies loaded."
} catch {
    Write-Log "Dependency installation/loading failed: $_" 'ERROR'
    throw
}

# Import modules
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'AssayValidation.psm1') -Force
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'UsedEquipment.psm1')   -Force
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'SharePoint.psm1')      -Force
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'ReportGenerator.psm1') -Force

# Helper: extract 10‑digit batch number from CSV header by scanning first 3x3 fields
function Get-BatchNumberFromCsv {
    param([string]$FilePath)
    $batch = $null
    try {
        $lines = Get-Content -Path $FilePath -TotalCount 3
        foreach ($line in $lines) {
            $parts = $line -split '[;,]'
            foreach ($p in $parts) {
                if ($p -match '\b(\d{10})\b') {
                    $batch = $Matches[1]
                    break
                }
            }
            if ($batch) { break }
        }
    } catch {
        Write-Log "Failed to extract batch number from $FilePath: $_" 'ERROR'
    }
    return $batch
}

# GUI creation
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$form = New-Object System.Windows.Forms.Form
$form.Text = 'IVDR Validation'
$form.Size = New-Object System.Drawing.Size(600,400)
$form.StartPosition = 'CenterScreen'
$form.AllowDrop = $true

$label = New-Object System.Windows.Forms.Label
$label.Text = "Dra och släpp dina filer här (CSV obligatorisk, Worksheet/Seal valfria)"
$label.Dock = 'Top'
$label.AutoSize = $true
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Dock = 'Fill'
$form.Controls.Add($listBox)

$validateBtn = New-Object System.Windows.Forms.Button
$validateBtn.Text = 'Starta validering'
$validateBtn.Dock = 'Bottom'
$form.Controls.Add($validateBtn)

$selectedFiles = @()

function Add-File {
    param([string[]]$files)
    foreach ($f in $files) {
        if (-not $selectedFiles -contains $f) {
            $selectedFiles += $f
            $listBox.Items.Add([System.IO.Path]::GetFileName($f)) | Out-Null
        }
    }
}

# Drag and drop events
$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [Windows.Forms.DragDropEffects]::Copy
    }
})
$form.Add_DragDrop({
    $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    Add-File -files $files
})

# File selection via OpenFileDialog when double-clicking list
$form.Add_DoubleClick({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Multiselect = $true
    $ofd.Filter = 'CSV and Excel files (*.csv;*.xlsx)|*.csv;*.xlsx'
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-File -files $ofd.FileNames
    }
})

# Validate button logic
$validateBtn.Add_Click({
    try {
        Write-Log "Validation started."
        # Determine CSV file
        $csvFile = $selectedFiles | Where-Object { $_.ToLower().EndsWith('.csv') } | Select-Object -First 1
        if (-not $csvFile) {
            [System.Windows.Forms.MessageBox]::Show('Ingen CSV-fil har valts','Fel',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        # Optional worksheet and seal files
        $worksheetFile = $selectedFiles | Where-Object { $_.ToLower().EndsWith('.xlsx') -and $_ -notmatch 'seal.*pos|seal.*neg' } | Select-Object -First 1
        $sealPosFile   = $selectedFiles | Where-Object { $_.ToLower() -match 'seal.*pos' } | Select-Object -First 1
        $sealNegFile   = $selectedFiles | Where-Object { $_.ToLower() -match 'seal.*neg' } | Select-Object -First 1
        Write-Log "CSV: $csvFile"
        if ($worksheetFile) { Write-Log "Worksheet: $worksheetFile" }
        if ($sealPosFile)   { Write-Log "Seal POS: $sealPosFile" }
        if ($sealNegFile)   { Write-Log "Seal NEG: $sealNegFile" }
        # Extract batch number and sharepoint metadata
        $batchNum  = Get-BatchNumberFromCsv -FilePath $csvFile
        if ($batchNum) {
            Write-Log "Detected batch number $batchNum"
            $batchInfo = Get-SharePointBatchInfo -BatchNumber $batchNum
        } else {
            Write-Log "No batch number found in CSV header" 'WARN'
            $batchInfo = @{}
        }
        # Validate CSV
        $valResult = Validate-AssayFile -FilePath $csvFile -RulesPath $rulesPath
        $results   = $valResult.Results
        $metadata  = $valResult.Metadata
        # Extract used equipment

        # Determine the delimiter by inspecting the first line once
        $firstLine = Get-Content -Path $csvFile -TotalCount 1
        if (($firstLine -split ';').Length -gt ($firstLine -split ',').Length) {
            $delimiter = ';'
        } else {
            $delimiter = ','
        }
        $csvRows = Import-Csv -Path $csvFile -Delimiter $delimiter
        # Build column map for equipment extraction
        $csvColMap   = @{}
        foreach ($p in $csvRows[0].PSObject.Properties.Name) { $csvColMap[$p.ToLower().Trim()] = $p }
        $usedEquip   = Get-UsedEquipmentFromCsv -CsvRows $csvRows -ColumnMap $csvColMap
        # Used pipettes and seal test data
        $pipettes    = Get-UsedPipettesFromWorksheet -WorksheetPath $worksheetFile
        $sealData    = Get-SealTestData -PosPath $sealPosFile -NegPath $sealNegFile
        # Control inventory (optional)
        $control     = Get-ControlInventory -RawDataPath $null
        # Equipment and pipette reference lookups (placeholders)
        $equipRef    = Get-EquipmentReference -XlsPath $null
        $pipRef      = Get-PipetteReference   -XlsPath $null
        # Generate report
        $reportPath  = Join-Path -Path $PSScriptRoot -ChildPath ("ValidationReport_" + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.xlsx')
        Write-ValidationReport -OutputPath $reportPath -Results $results -Metadata $metadata -BatchInfo $batchInfo -EquipmentRef $equipRef -UsedEquipment $usedEquip -SealData $sealData -ControlInventory $control
        Write-Log "Validation finished. Report saved to $reportPath"
        [System.Windows.Forms.MessageBox]::Show("Valideringen är klar. Rapporten är sparad på:\n$reportPath","Färdig",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        Write-Log "Validation failed: $_" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show("Ett fel uppstod:\n$($_.Exception.Message)","Fel",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

[void]$form.ShowDialog()
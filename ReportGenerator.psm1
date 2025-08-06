<#
    Module: ReportGenerator.psm1

    This module is responsible for composing the final Excel report using
    EPPlus.  It accepts validation results, equipment information, seal test
    data and batch metadata, then arranges them into multiple worksheets.
    The design follows IVDR readability guidelines: structured, color coded
    by severity and includes metadata on the first sheet.

    NOTE: The EPPlus package must be available via the Install-Dependencies
    script.  The module assumes the assembly has been loaded (via Add-Type)
    or imported.
#>

function Write-ValidationReport {
    <#
        .SYNOPSIS
            Generates an Excel workbook summarising validation results and
            auxiliary information.

        .PARAMETER OutputPath
            Destination file path for the Excel report (.xlsx).

        .PARAMETER Results
            Array of result objects returned by Validate-AssayFile.  Each
            object must contain Row, SampleID, Assay, Validation and Message.

        .PARAMETER Metadata
            Hashtable containing metadata such as RulesVersion, ScriptVersion,
            FileValidated and Timestamp.

        .PARAMETER BatchInfo
            Hashtable of SharePoint metadata for the current batch (optional).

        .PARAMETER EquipmentRef
            Array of approved equipment reference entries (optional).

        .PARAMETER UsedEquipment
            Array of used equipment extracted from CSV (optional).

        .PARAMETER SealData
            Hashtable with Pos and Neg seal test arrays (optional).

        .PARAMETER ControlInventory
            Array of control material entries (optional).

        .EXAMPLE
            Write-ValidationReport -OutputPath 'report.xlsx' -Results $res.Results \
                                  -Metadata $res.Metadata -BatchInfo $batch \
                                  -EquipmentRef $equipRef -UsedEquipment $used \
                                  -SealData $seal -ControlInventory $ctrl

        Creates an Excel report at the specified path.
    #>
    [CmdletBinding()] param(
        [Parameter(Mandatory=$true)] [string]$OutputPath,
        [Parameter(Mandatory=$true)] [array]$Results,
        [Parameter(Mandatory=$true)] [hashtable]$Metadata,
        [hashtable]$BatchInfo,
        $EquipmentRef,
        [array]$UsedEquipment,
        $SealData,
        [array]$ControlInventory
    )
    try {
        # Ensure EPPlus namespace is available
        if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'EPPlus' })) {
            throw "EPPlus assembly is not loaded. Ensure Install-Dependencies has run."
        }
        $pkg = New-Object OfficeOpenXml.ExcelPackage
        # Colours for categories
        $colorMap = @{
            'OK'           = [System.Drawing.Color]::FromArgb(0, 176, 80)   # green
            'Minor Error'  = [System.Drawing.Color]::FromArgb(255, 192, 0) # yellow
            'Valid Error'  = [System.Drawing.Color]::FromArgb(0, 112, 192) # blue
            'Unknown Error'= [System.Drawing.Color]::FromArgb(197, 90, 17) # orange
            'Error'        = [System.Drawing.Color]::FromArgb(255, 0, 0)   # red
            'Major Error'  = [System.Drawing.Color]::FromArgb(192, 0, 0)   # dark red
            'Naming Error' = [System.Drawing.Color]::FromArgb(155, 0, 211) # purple
        }
        # Sheet 1: BatchInfo
        $sheet = $pkg.Workbook.Worksheets.Add('BatchInfo')
        $sheet.Cells[1,1].Value = 'Report generated'
        $sheet.Cells[1,2].Value = $Metadata.Timestamp
        $sheet.Cells[2,1].Value = 'Script version'
        $sheet.Cells[2,2].Value = $Metadata.ScriptVersion
        $sheet.Cells[3,1].Value = 'Rules version'
        $sheet.Cells[3,2].Value = $Metadata.RulesVersion
        $sheet.Cells[4,1].Value = 'Input file'
        $sheet.Cells[4,2].Value = $Metadata.FileValidated
        if ($BatchInfo) {
            $row = 6
            $sheet.Cells[$row,1].Value = 'Batch metadata'
            $row++
            foreach ($key in $BatchInfo.Keys) {
                $sheet.Cells[$row,1].Value = $key
                $sheet.Cells[$row,2].Value = $BatchInfo[$key]
                $row++
            }
        }
        $sheet.Cells.AutoFitColumns()
        # Sheet 2: Equipment reference
        $sheet2 = $pkg.Workbook.Worksheets.Add('EquipmentRef')
        # Determine the structure: if EquipmentRef is a hashtable (mapping instrument ➜ calibration/SNN),
        # output instrument name, calibration and serial numbers; otherwise assume array of objects
        $sheet2.Cells[1,1].Value = 'Instrument'
        $sheet2.Cells[1,2].Value = 'Calibration'
        $sheet2.Cells[1,3].Value = 'SerialNumbers'
        $row = 2
        if ($EquipmentRef) {
            if ($EquipmentRef -is [hashtable]) {
                foreach ($name in $EquipmentRef.Keys) {
                    $item = $EquipmentRef[$name]
                    $sheet2.Cells[$row,1].Value = $name
                    $sheet2.Cells[$row,2].Value = $item.Calibration
                    $sheet2.Cells[$row,3].Value = [string]::Join(',',$item.SNN)
                    $row++
                }
            } else {
                foreach ($eq in $EquipmentRef) {
                    $sheet2.Cells[$row,1].Value = $eq.InstrumentSN
                    $sheet2.Cells[$row,2].Value = ''
                    $sheet2.Cells[$row,3].Value = $eq.ModuleSN
                    $row++
                }
            }
        }
        $sheet2.Cells.AutoFitColumns()
        # Sheet 3: Used Equipment
        $sheet3 = $pkg.Workbook.Worksheets.Add('UsedEquipment')
        $sheet3.Cells[1,1].Value = 'InstrumentSN'
        $sheet3.Cells[1,2].Value = 'ModuleSN'
        $row = 2
        if ($UsedEquipment) {
            foreach ($eq in $UsedEquipment) {
                $sheet3.Cells[$row,1].Value = $eq.InstrumentSN
                $sheet3.Cells[$row,2].Value = $eq.ModuleSN
                $row++
            }
        }
        $sheet3.Cells.AutoFitColumns()
        # Sheet 4: SealTest
        $sheet4 = $pkg.Workbook.Worksheets.Add('SealTest')
        $sheet4.Cells[1,1].Value = 'Category'
        $sheet4.Cells[1,2].Value = 'Field'
        $sheet4.Cells[1,3].Value = 'NEG'
        $sheet4.Cells[1,4].Value = 'POS'
        $row = 2
        if ($SealData) {
            if ($SealData.HeaderFields) {
                foreach ($key in $SealData.HeaderFields.Keys) {
                    $sheet4.Cells[$row,1].Value = 'Header'
                    $sheet4.Cells[$row,2].Value = $key
                    $sheet4.Cells[$row,3].Value = $SealData.HeaderFields[$key].NEG
                    $sheet4.Cells[$row,4].Value = $SealData.HeaderFields[$key].POS
                    $row++
                }
            }
            # Testers lists
            $sheet4.Cells[$row,1].Value = 'Testers'; $sheet4.Cells[$row,2].Value = 'NEG'; $sheet4.Cells[$row,3].Value = [string]::Join(', ', $SealData.Testers.NEG)
            $sheet4.Cells[$row,4].Value = [string]::Join(', ', $SealData.Testers.POS)
            $row++
            # Violations
            if ($SealData.Violations.Count -gt 0) {
                $sheet4 = $pkg.Workbook.Worksheets.Add('SealViolations')
                $sheet4.Cells[1,1].Value = 'Type'
                $sheet4.Cells[1,2].Value = 'Sheet'
                $sheet4.Cells[1,3].Value = 'Cartridge'
                $sheet4.Cells[1,4].Value = 'InitialW'
                $sheet4.Cells[1,5].Value = 'FinalW'
                $sheet4.Cells[1,6].Value = 'WeightLoss'
                $sheet4.Cells[1,7].Value = 'Status'
                $sheet4.Cells[1,8].Value = 'Observation'
                $vr = 2
                foreach ($vi in $SealData.Violations) {
                    $sheet4.Cells[$vr,1].Value = $vi.Type
                    $sheet4.Cells[$vr,2].Value = $vi.Sheet
                    $sheet4.Cells[$vr,3].Value = $vi.Cartridge
                    $sheet4.Cells[$vr,4].Value = $vi.InitialW
                    $sheet4.Cells[$vr,5].Value = $vi.FinalW
                    $sheet4.Cells[$vr,6].Value = $vi.WeightLoss
                    $sheet4.Cells[$vr,7].Value = $vi.Status
                    $sheet4.Cells[$vr,8].Value = $vi.Obs
                    $vr++
                }
                $sheet4.Cells.AutoFitColumns()
            }
            # Mismatches
            if ($SealData.Mismatches.Count -gt 0) {
                $sheetMM = $pkg.Workbook.Worksheets.Add('SealMismatches')
                $sheetMM.Cells[1,1].Value = 'Field'
                $sheetMM.Cells[1,2].Value = 'NEG'
                $sheetMM.Cells[1,3].Value = 'POS'
                $mr = 2
                foreach ($mm in $SealData.Mismatches) {
                    $sheetMM.Cells[$mr,1].Value = $mm.Field
                    $sheetMM.Cells[$mr,2].Value = $mm.NEG
                    $sheetMM.Cells[$mr,3].Value = $mm.POS
                    $mr++
                }
                $sheetMM.Cells.AutoFitColumns()
            }
        }
        $sheet4.Cells.AutoFitColumns()
        # Sheet 5: Control Inventory
        $sheet5 = $pkg.Workbook.Worksheets.Add('ControlInventory')
        $sheet5.Cells[1,1].Value = 'ControlID'
        $row = 2
        if ($ControlInventory) {
            foreach ($c in $ControlInventory) {
                $sheet5.Cells[$row,1].Value = $c
                $row++
            }
        }
        $sheet5.Cells.AutoFitColumns()
        # Sheet 6: Deviations
        $sheet6 = $pkg.Workbook.Worksheets.Add('Deviations')
        $sheet6.Cells[1,1].Value = 'Row'
        $sheet6.Cells[1,2].Value = 'SampleID'
        $sheet6.Cells[1,3].Value = 'Assay'
        $sheet6.Cells[1,4].Value = 'Validation'
        $sheet6.Cells[1,5].Value = 'Message'
        $row = 2
        foreach ($res in $Results) {
            $sheet6.Cells[$row,1].Value = $res.Row
            $sheet6.Cells[$row,2].Value = $res.SampleID
            $sheet6.Cells[$row,3].Value = $res.Assay
            $sheet6.Cells[$row,4].Value = $res.Validation
            $sheet6.Cells[$row,5].Value = $res.Message
            # Apply colour based on validation
            if ($colorMap.ContainsKey($res.Validation)) {
                $color = $colorMap[$res.Validation]
                $sheet6.Cells[$row,1,$row,5].Style.Fill.PatternType = 'Solid'
                $sheet6.Cells[$row,1,$row,5].Style.Fill.BackgroundColor.SetColor($color)
            }
            $row++
        }
        $sheet6.Cells.AutoFitColumns()
        # Save workbook
        $bytes = $pkg.GetAsByteArray()
        [System.IO.File]::WriteAllBytes($OutputPath, $bytes)
        Write-Host "Report written to $OutputPath" -ForegroundColor Green
    } catch {
        throw "Failed to generate report: $_"
    }
}
<#

    Module: SharePoint

    This module encapsulates operations against SharePoint to retrieve batch
    metadata and other information needed for IVDR validation. The

    Module: SharePoint.psm1

    This module encapsulates operations against SharePoint to retrieve batch 

    implementation currently returns mock data; replace the stubbed logic with
    real PnP.PowerShell calls in production.
#>


function Get-SharePointBatchInfo {
    <#
        .SYNOPSIS
            Retrieves metadata for a given SAP batch number from SharePoint.


        .PARAMETER BatchNumber
            A 10-digit SAP batch number.

        .OUTPUTS
            PSCustomObject with batch information or $null on failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [string]$BatchNumber
    )

    if ($BatchNumber -notmatch '^\d{10}$') {
        throw 'Batch number must be exactly 10 digits.'
    }

    $env:PNPPOWERSHELL_UPDATECHECK = 'Off'
    $tenant = 'danaher.onmicrosoft.com'
    $clientId = 'INSERT LATER'
    $certificateprivatekey = 'INSERT-LATER'
    $siteUrl = 'https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management'

    try {
        Connect-PnPOnline -Url $siteUrl -Tenant $tenant -ClientId $clientId -CertificateBase64Encoded $certificateprivatekey -ErrorAction Stop
    } catch {
        Write-Warning "Could not connect to SharePoint: $($_.Exception.Message)"
        return $null
    }

    $fields = @(
        'Work_x0020_Center',
        'Title','Batch_x0023_',
        'SAP_x0020_Batch_x0023__x0020_2',
        'LSP','Material','BBD_x002f_SLED',
        'Actual_x0020_startdate_x002f__x0',
        'PAL_x0020__x002d__x0020_Sample_x',
        'Sample_x0020_Reagent_x0020_P_x00',
        'Order_x0020_quantity','Total_x0020_good',
        'ITP_x0020_Test_x0020_results',
        'IPT_x0020__x002d__x0020_Testing_0',
        'MES_x0020__x002d__x0020_Order_x0'
    )

    try {
        $items = Get-PnPListItem -List 'Cepheid | Production orders' -PageSize 5000 -Fields $fields -ErrorAction Stop
    } catch {
        Write-Warning "Failed to query SharePoint: $($_.Exception.Message)"
        return $null
    }

    $match = $null
    foreach ($item in $items) {
        if ($item.FieldValues['Batch_x0023_'] -eq $BatchNumber) {
            $match = $item
            break
        }
    }

    if (-not $match) {
        Write-Warning "Ingen post hittades för Batch: $BatchNumber"
        return $null
    }

    return [PSCustomObject]@{
        BatchNumber     = $BatchNumber
        LSP             = $match.FieldValues['LSP']
        ProductName     = $match.FieldValues['Material']
        ManufactureDate = $match.FieldValues['Actual_x0020_startdate_x002f__x0']
        ExpiryDate      = $match.FieldValues['BBD_x002f_SLED']
        Assay           = $match.FieldValues['Material']
    }
        .DESCRIPTION
            Stub implementation that returns mock data. Replace with
            Connect-PnPOnline and Get-PnPListItem calls for production use.

        .PARAMETER BatchNumber
            A 10-digit SAP batch number extracted from the worksheet xlsx header.

        .OUTPUTS
    #>

    param([string]$BatchNumber)
	Write-Verbose "Retrieving batch info for $BatchNumber from SharePoint"
    return @{
        BatchNumber     = $BatchNumber
        ProductName     = 'MockProduct'
        ManufactureDate = (Get-Date).AddDays(-7).ToString('yyyy-MM-dd')
        ExpiryDate      = (Get-Date).AddMonths(12).ToString('yyyy-MM-dd')
        Status          = 'Open'
    }

}


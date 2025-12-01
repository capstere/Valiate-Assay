param()
function New-Result {
    param(
        [bool]$Ok,
        [object]$Data = $null,
        [string[]]$Errors = @(),
        [string[]]$Warnings = @()
    )

    [pscustomobject]@{
        Ok       = [bool]$Ok
        Data     = $Data
        Errors   = $Errors
        Warnings = $Warnings
    }
}

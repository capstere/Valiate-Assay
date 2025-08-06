<#
    This module encapsulates assay validation logic used in the IVDR project.
    It exposes two primary functions:

        Get-SuffixLogic           – Determines whether the suffix parity (X/+)
                                      needs to be swapped based on batch data.
        Test-AssayFile            – Validates a single assay CSV file against
                                      XML rules and returns detailed results.

    The code is derived from prior scripts but refactored into a reusable
    module.  Additional helper functions can be added here as needed.

    Version history
    ---------------
    1.1.0 – Refactored into module, added metadata reporting, suffix swap logic.
    1.0.0 – Initial implementation.
#>

function Get-SuffixLogic {
    <#
        .SYNOPSIS
            Determines whether the suffix parity logic (X for odd, + for even)
            should be swapped for a given set of rows and assay rules.

        .DESCRIPTION
            The function iterates over an array of rows (imported from CSV) and
            counts how many times suffix 'X' appears with an even last digit in
            the cartridge serial number and vice versa.  If the reversed match
            count exceeds the standard match count, the function returns `$true`
            indicating that the suffix logic should be swapped.

        .PARAMETER Rows
            The array of data rows imported from a CSV or Excel sheet.

        .PARAMETER AssayRules
            A hashtable containing assay definitions keyed by assay name. Each
            definition must include a SampleIDRegex property.

        .OUTPUTS
            [bool] True if the suffix logic should be swapped, otherwise False.
    #>
    [CmdletBinding()] param(
        [Parameter(Mandatory=$true)] [array]$Rows,
        [Parameter(Mandatory=$true)] [hashtable]$AssayRules
    )
    $xOdd = 0; $xEven = 0; $plusOdd = 0; $plusEven = 0
    foreach ($row in $Rows) {
        $sampleID    = $row.'Sample ID'
        $assay       = $row.Assay
        $cartridgeSN = $row.'Cartridge S/N'
        if (-not $sampleID -or -not $cartridgeSN -or -not $assay) { continue }
        if (-not $AssayRules.ContainsKey($assay)) { continue }
        $regex = $AssayRules[$assay].SampleIDRegex
        if ($sampleID -match $regex) {
            $suffixFromID = $Matches[6]
            if ($cartridgeSN -match '([0-9])$') {
                $lastDigit = [int]$Matches[1]
                $isEven    = (($lastDigit % 2) -eq 0)
                switch ($suffixFromID) {
                    'X' { if ($isEven) { $xEven++ } else { $xOdd++ } }
                    '+' { if ($isEven) { $plusEven++ } else { $plusOdd++ } }
                }
            }
        }
    }
    $standardMatch = $xOdd + $plusEven
    $reversedMatch = $xEven + $plusOdd
    return ($reversedMatch -gt $standardMatch)
}

function Test-AssayFile {
    <#
        .SYNOPSIS
            Validates an assay data file against rules defined in an XML configuration.

        .DESCRIPTION
            This function reads an input CSV file (semicolon or comma delimited) and
            validates each row according to a set of rules stored in an XML file.
            It verifies sample ID format, status, test type codes, test results,
            numeric ranges, suffix parity and error codes.  Major errors
            (false positives/negatives) are detected when the test result does not
            match the expected list for its test type. Naming errors are
            distinguished when applicable.

        .PARAMETER FilePath
            The path to the CSV file to validate.

        .PARAMETER RulesPath
            The path to the XML file containing assay rules.

        .OUTPUTS
            A hashtable with two keys: Results (an array of validation objects) and
            Metadata (information about the run).
    #>
    [CmdletBinding()] param(
        [Parameter(Mandatory=$true)] [string]$FilePath,
        [Parameter(Mandatory=$true)] [string]$RulesPath
    )
    # Load rules
    [xml]$rulesXml = Get-Content -Path $RulesPath -ErrorAction Stop
    $rulesVersion = $rulesXml.AssayRules.Version
    # Parse assays
    $assayRules = @{}
    foreach ($assayNode in $rulesXml.AssayRules.Assay) {
        $name = $assayNode.Name
        $rule = [ordered]@{}
        $rule.SampleIDRegex  = $assayNode.SampleIDRegex
        $rule.SuffixRule     = $assayNode.SuffixRule
        $rule.StartTimeRegex = $assayNode.StartTimeRegex
        # Map test types to codes
        $rule.TestTypeToCode = @{}
        foreach ($map in $assayNode.TestTypeToCode.Map) { $rule.TestTypeToCode[$map.Type] = $map.Code }
        # Valid results per code
        $rule.ValidResults = @{}
        foreach ($result in $assayNode.ValidResults.Result) {
            $code = $result.Code
            $vals = @(); foreach ($v in $result.Value) { $vals += $v.'#text' }
            $rule.ValidResults[$code] = $vals
        }
        # Extra columns
        $rule.ExtraColumns = @{}
        foreach ($col in $assayNode.ExtraColumns.Column) { $rule.ExtraColumns[$col.Name] = $col.Key }
        # Allowed status
        $rule.AllowedStatus = @(); foreach ($st in $assayNode.AllowedStatus.Status) { $rule.AllowedStatus += $st.'#text' }
        # Number ranges per type
        $rule.NumberRange = @{}
        foreach ($type in $assayNode.NumberRange.Type) { $rule.NumberRange[$type.Name] = @{ Start = $type.Start; End = $type.End } }
        # Check mismatch flag
        $rule.CheckSampleTypeMismatch = $false
        if ($assayNode.CheckSampleTypeMismatch -and $assayNode.CheckSampleTypeMismatch.'#text' -match '^true$') {
            $rule.CheckSampleTypeMismatch = $true
        }
        # Placeholder for swap flag
        $rule.SwapSuffix = $false
        $assayRules[$name] = $rule
    }
    # Status rules
    $statusRules = @{}
    foreach ($sr in $rulesXml.AssayRules.StatusRules.Rule) {
        $statusRules[$sr.Name] = @{ RequireError = ($sr.RequireError -eq 'true'); AllowedTestResult = @() }
        foreach ($atr in $sr.AllowedTestResult) { $statusRules[$sr.Name].AllowedTestResult += $atr.'#text' }
    }
    # Error categories
    $minorErrors = @(); foreach ($code in $rulesXml.AssayRules.ErrorCategories.Minor.Code) { $minorErrors += $code.'#text' }
    $validErrors = @(); foreach ($code in $rulesXml.AssayRules.ErrorCategories.Valid.Code) { $validErrors += $code.'#text' }
    # Determine delimiter based on first line
    $firstLine = (Get-Content -Path $FilePath -TotalCount 1)
    $semiCount  = ($firstLine -split ';').Length - 1
    $commaCount = ($firstLine -split ',').Length - 1
    if ($semiCount -ge $commaCount) {
        $delimiter = ';'
    } else {
        $delimiter = ','
    }
    # Import CSV
    $rows = Import-Csv -Path $FilePath -Delimiter $delimiter
    if ($rows.Count -eq 0) { throw "Input file contains no rows." }
    # Normalize column names
    $colMap = @{}
    foreach ($prop in $rows[0].PSObject.Properties.Name) {
        $colMap[$prop.ToLower().Trim()] = $prop
    }
    # Required columns
    $required = @('assay','sample id','test type','cartridge s/n','test result','status','error','start time')
    foreach ($r in $required) {
        if (-not $colMap.ContainsKey($r)) { throw "Missing required column '$r' in input file." }
    }
    # Evaluate suffix swap per assay
    foreach ($assayName in $assayRules.Keys) {
        $batch = $rows | Where-Object { $_.$($colMap['assay']) -eq $assayName }
        if ($batch.Count -gt 0) {
            $assayRules[$assayName].SwapSuffix = Get-SuffixLogic -Rows $batch -AssayRules $assayRules
        }
    }
    # Results list
    $results = @()
    $rowIndex = 0
    foreach ($row in $rows) {
        $rowIndex++
        $sampleID    = $row.$($colMap['sample id'])
        $assayName   = $row.$($colMap['assay'])
        $testType    = $row.$($colMap['test type'])
        $cartridgeSN = $row.$($colMap['cartridge s/n'])
        $testResult  = $row.$($colMap['test result'])
        $status      = $row.$($colMap['status'])
        $errorText   = $row.$($colMap['error'])
        $startTime   = $row.$($colMap['start time'])
        $res = [PSCustomObject]@{
            Row      = $rowIndex
            SampleID = $sampleID
            Assay    = $assayName
            Validation = 'OK'
            Message    = ''
        }
        # Unknown assay
        if (-not $assayRules.ContainsKey($assayName)) {
            $res.Validation = 'Error'
            $res.Message    = "Unknown assay '$assayName'"
            $results += $res; continue
        }
        $rule = $assayRules[$assayName]
        # Status check
        if (-not $rule.AllowedStatus -contains $status) {
            $res.Validation = 'Error'
            $res.Message    = "Invalid status '$status'"
            $results += $res; continue
        }
        $srule = $statusRules[$status]
        # Missing error
        if ($srule.RequireError -and [string]::IsNullOrEmpty($errorText)) {
            $res.Validation = 'Error'
            $res.Message    = "Missing error text for status '$status'"
            $results += $res; continue
        }
        # Allowed test results for Aborted/Incomplete
        if ($srule.AllowedTestResult.Count -gt 0) {
            if (-not ($srule.AllowedTestResult -contains $testResult)) {
                $res.Validation = 'Error'
                $res.Message    = "Unexpected test result '$testResult' for status '$status'"
                $results += $res; continue
            }
            # classify error code
            if (-not [string]::IsNullOrEmpty($errorText) -and $errorText -match '\b([0-9]+)\b') {
                $errCode = $Matches[1]
                if ($minorErrors -contains $errCode) {
                    $res.Validation = 'Minor Error'
                    $res.Message    = "Minor error $errCode"
                } elseif ($validErrors -contains $errCode) {
                    $res.Validation = 'Valid Error'
                    $res.Message    = "Valid instrument error $errCode"
                } else {
                    $res.Validation = 'Unknown Error'
                    $res.Message    = "Unknown error code $errCode"
                }
            }
            $results += $res; continue
        }
        # For Done status, perform full validation
        if (-not $rule.TestTypeToCode.ContainsKey($testType)) {
            $res.Validation = 'Error'
            $res.Message    = "Unknown test type '$testType'"
            $results += $res; continue
        }
        $code = $rule.TestTypeToCode[$testType]
        # Sample ID pattern
        if ($sampleID -notmatch $rule.SampleIDRegex) {
            $res.Validation = 'Error'
            $res.Message    = "Sample ID does not match expected pattern"
            $results += $res; continue
        }
        $codeFromId   = $Matches[3]
        $sampleNum    = [int]$Matches[4]
        $alphaSuffix  = $Matches[5]
        $suffixChar   = $Matches[6]
        # Check mismatch
        if ($rule.CheckSampleTypeMismatch -and $codeFromId -ne $code) {
            $res.Validation = 'Naming Error'
            $res.Message    = "Sample ID code ($codeFromId) does not match Test Type code ($code)"
            $results += $res; continue
        }
        # Number range
        if ($rule.NumberRange.ContainsKey($testType)) {
            $range = $rule.NumberRange[$testType]
            if ($sampleNum -lt [int]$range.Start -or $sampleNum -gt [int]$range.End) {
                $res.Validation = 'Error'
                $res.Message    = "Sample number $sampleNum outside allowed range $($range.Start)-$($range.End) for $testType"
                $results += $res; continue
            }
        }
        # Suffix parity
        $lastDigit = $null
        if ($cartridgeSN -match '([0-9])$') { $lastDigit = [int]$Matches[1] }
        if ($lastDigit -ne $null) {
            $expectedSuffix = $null
            switch -regex ($rule.SuffixRule) {
                '^Odd=X,Even=\+$' {
                    if (($lastDigit % 2) -eq 1) { $expectedSuffix = 'X' } else { $expectedSuffix = '+' }
                }
                '^Always=\+$' { $expectedSuffix = '+' }
                '^Always=X'   { $expectedSuffix = 'X' }
            }
            if ($rule.SwapSuffix -and $expectedSuffix) {
                if ($expectedSuffix -eq 'X') { $expectedSuffix = '+' } else { $expectedSuffix = 'X' }
            }
            if ($expectedSuffix -and $suffixChar -ne $expectedSuffix) {
                $res.Validation = 'Error'
                $res.Message    = "Suffix '$suffixChar' does not match expected '$expectedSuffix' for cartridge $cartridgeSN"
                $results += $res; continue
            }
        }
        # Test result
        $expectedValues = $rule.ValidResults[$code]
        if (-not ($expectedValues -contains $testResult)) {
            $res.Validation = 'Major Error'
            $res.Message    = "Unexpected Test Result '$testResult' for code $code"
            $results += $res; continue
        }
        # Start time
        if ($startTime -notmatch $rule.StartTimeRegex) {
            $res.Validation = 'Error'
            $res.Message    = "Start Time does not match expected format"
            $results += $res; continue
        }
        # Error code if present on Done
        if (-not [string]::IsNullOrEmpty($errorText) -and $errorText -match '\b([0-9]+)\b') {
            $err = $Matches[1]
            if ($minorErrors -contains $err) {
                $res.Validation = 'Minor Error'; $res.Message = "Minor error $err"
            } elseif ($validErrors -contains $err) {
                $res.Validation = 'Valid Error'; $res.Message = "Valid instrument error $err"
            } else {
                $res.Validation = 'Unknown Error'; $res.Message = "Unknown error code $err"
            }
            $results += $res; continue
        }
        # All checks passed
        $results += $res
    }
    # Metadata
    $metadata = [PSCustomObject]@{
        RulesVersion  = $rulesVersion
        ScriptVersion = '1.1.0'
        FileValidated = (Get-Item -Path $FilePath).Name
        Timestamp     = (Get-Date).ToString('s')
    }
    return @{ Results = $results; Metadata = $metadata }
}
Set-Alias -Name Validate-AssayFile -Value Test-AssayFile

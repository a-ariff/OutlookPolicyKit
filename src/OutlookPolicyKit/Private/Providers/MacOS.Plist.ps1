<#
.SYNOPSIS
    macOS Plist provider for Outlook policies
    
.DESCRIPTION
    Provides functions to get and set Outlook policies using macOS defaults command
    and plist manipulation for Microsoft Outlook policies.
#>

function Get-OPKMacPolicy {
    <#
    .SYNOPSIS
        Gets an Outlook policy from macOS plist
    .PARAMETER Domain
        The domain to query (default: com.microsoft.Outlook)
    .PARAMETER Key
        The policy key to retrieve
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Domain = 'com.microsoft.Outlook',
        
        [Parameter(Mandatory = $true)]
        [string]$Key
    )
    
    try {
        # Try using defaults command first
        $result = & /usr/bin/defaults read $Domain $Key 2>$null
        
        if ($LASTEXITCODE -eq 0 -and $result) {
            # Determine the type based on the value
            $type = 'unknown'
            $value = $result
            
            if ($result -match '^\d+$') {
                $type = 'int'
                $value = [int]$result
            }
            elseif ($result -in @('0', '1', 'true', 'false', 'YES', 'NO')) {
                $type = 'bool'
                $value = $result -in @('1', 'true', 'YES')
            }
            else {
                $type = 'string'
                $value = $result
            }
            
            return [PSCustomObject]@{
                Domain = $Domain
                Key = $Key
                Value = $value
                Type = $type
                Source = 'defaults'
            }
        }
        
        # Fallback to PlistBuddy if defaults fails
        $plistPath = "~/Library/Preferences/$Domain.plist"
        if (Test-Path $plistPath) {
            $result = & /usr/libexec/PlistBuddy -c "Print $Key" $plistPath 2>$null
            if ($LASTEXITCODE -eq 0) {
                return [PSCustomObject]@{
                    Domain = $Domain
                    Key = $Key
                    Value = $result
                    Type = 'string'
                    Source = 'PlistBuddy'
                }
            }
        }
        
        return $null
    }
    catch {
        Write-Error "Failed to get policy '$Key' from domain '$Domain': $($_.Exception.Message)"
        return $null
    }
}

function Set-OPKMacPolicy {
    <#
    .SYNOPSIS
        Sets an Outlook policy in macOS plist
    .PARAMETER Domain
        The domain to set (default: com.microsoft.Outlook)
    .PARAMETER Key
        The policy key to set
    .PARAMETER Value
        The value to set
    .PARAMETER Type
        The type of value (bool, int, string)
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Domain = 'com.microsoft.Outlook',
        
        [Parameter(Mandatory = $true)]
        [string]$Key,
        
        [Parameter(Mandatory = $true)]
        $Value,
        
        [Parameter()]
        [ValidateSet('bool', 'int', 'string', 'float')]
        [string]$Type = 'string'
    )
    
    try {
        # Convert value based on type
        $convertedValue = switch ($Type) {
            'bool' { 
                if ($Value -is [bool]) { $Value.ToString().ToLower() }
                else { ([bool]$Value).ToString().ToLower() }
            }
            'int' { [int]$Value }
            'float' { [float]$Value }
            'string' { [string]$Value }
            default { [string]$Value }
        }
        
        # Try using defaults command first
        $cmd = @('/usr/bin/defaults', 'write', $Domain, $Key)
        
        switch ($Type) {
            'bool' { 
                $cmd += @('-bool', $convertedValue)
            }
            'int' { 
                $cmd += @('-int', $convertedValue)
            }
            'float' { 
                $cmd += @('-float', $convertedValue)
            }
            'string' { 
                $cmd += @('-string', $convertedValue)
            }
        }
        
        $result = & $cmd[0] $cmd[1..($cmd.Count-1)] 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Verbose "Set policy '$Key' = '$Value' ($Type) in domain '$Domain'"
            return $true
        }
        else {
            # Fallback to PlistBuddy
            Write-Warning "defaults command failed, trying PlistBuddy: $result"
            
            $plistPath = "~/Library/Preferences/$Domain.plist"
            
            # Create plist if it doesn't exist
            if (-not (Test-Path $plistPath)) {
                & /usr/bin/touch $plistPath
                & /usr/libexec/PlistBuddy -c "Add :$Key $Type $convertedValue" $plistPath
            }
            else {
                # Try to set the value, add if it doesn't exist
                $setBuddy = & /usr/libexec/PlistBuddy -c "Set :$Key $convertedValue" $plistPath 2>&1
                if ($LASTEXITCODE -ne 0) {
                    & /usr/libexec/PlistBuddy -c "Add :$Key $Type $convertedValue" $plistPath
                }
            }
            
            if ($LASTEXITCODE -eq 0) {
                Write-Verbose "Set policy '$Key' = '$Value' ($Type) in domain '$Domain' using PlistBuddy"
                return $true
            }
        }
        
        return $false
    }
    catch {
        Write-Error "Failed to set policy '$Key' in domain '$Domain': $($_.Exception.Message)"
        return $false
    }
}

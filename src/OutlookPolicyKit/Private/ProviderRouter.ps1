<#.SYNOPSIS
    Provider router for Outlook policy management
    
.DESCRIPTION
    Routes policy get/set operations to appropriate platform-specific providers
    and provides friendly name mapping for Outlook policies.
#>
# Policy name mappings between friendly names and implementation details
$script:PolicyMappings = @{
    # Windows Registry mappings
    'Windows' = @{
        'DisableExternalImages' = @{
            Name = 'ExternalContent'
            Type = 'DWORD'
            DefaultValue = 1
            Description = 'Block external content in HTML emails'
        }
        'BlockUnsafeHTML' = @{
            Name = 'BlockHTTPImages'
            Type = 'DWORD' 
            DefaultValue = 1
            Description = 'Block HTTP images in HTML emails'
        }
        'CachedMode' = @{
            Name = 'CachedMode'
            Type = 'DWORD'
            DefaultValue = 1
            Description = 'Enable Exchange cached mode'
        }
        'SyncWindowDays' = @{
            Name = 'SyncWindowSettingDays'
            Type = 'DWORD'
            DefaultValue = 365
            Description = 'Number of days to sync in cached mode'
        }
    }
    # macOS plist mappings
    'macOS' = @{
        'DisableExternalImages' = @{
            Key = 'DisableExternalContent'
            Type = 'bool'
            DefaultValue = $true
            Description = 'Block external content in HTML emails'
        }
        'BlockUnsafeHTML' = @{
            Key = 'BlockHTTPImages'
            Type = 'bool'
            DefaultValue = $true  
            Description = 'Block HTTP images in HTML emails'
        }
        'CachedMode' = @{
            Key = 'CachedExchangeMode'
            Type = 'bool'
            DefaultValue = $true
            Description = 'Enable Exchange cached mode'
        }
        'SyncWindowDays' = @{
            Key = 'SyncWindowSettingDays'
            Type = 'int'
            DefaultValue = 365
            Description = 'Number of days to sync in cached mode'
        }
    }
}
function Get-OPKPolicy {
    <#
    .SYNOPSIS
        Gets an Outlook policy using the appropriate provider
    .PARAMETER Platform
        The platform (Windows or macOS)
    .PARAMETER Name
        The friendly policy name
    .PARAMETER Scope
        For Windows: Machine or User (default: Machine)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Windows', 'macOS')]
        [string]$Platform,
        
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter()]
        [ValidateSet('Machine', 'User')]
        [string]$Scope = 'Machine'
    )
    
    try {
        if (-not $script:PolicyMappings[$Platform][$Name]) {
            Write-Error "Unknown policy '$Name' for platform '$Platform'"
            return $null
        }
        
        $policyInfo = $script:PolicyMappings[$Platform][$Name]
        
        switch ($Platform) {
            'Windows' {
                $result = Get-OPKWinPolicy -Name $policyInfo.Name -Scope $Scope
                if ($result) {
                    $result | Add-Member -NotePropertyName 'FriendlyName' -NotePropertyValue $Name -Force
                    $result | Add-Member -NotePropertyName 'Description' -NotePropertyValue $policyInfo.Description -Force
                }
                return $result
            }
            'macOS' {
                $result = Get-OPKMacPolicy -Key $policyInfo.Key
                if ($result) {
                    $result | Add-Member -NotePropertyName 'FriendlyName' -NotePropertyValue $Name -Force
                    $result | Add-Member -NotePropertyName 'Description' -NotePropertyValue $policyInfo.Description -Force
                }
                return $result
            }
        }
    }
    catch {
        Write-Error "Failed to get policy '$Name' on '$Platform': $($_.Exception.Message)"
        return $null
    }
}
function Set-OPKPolicy {
    <#
    .SYNOPSIS
        Sets an Outlook policy using the appropriate provider
    .PARAMETER Platform
        The platform (Windows or macOS)
    .PARAMETER Name
        The friendly policy name
    .PARAMETER Value
        The value to set
    .PARAMETER Scope
        For Windows: Machine or User (default: Machine)
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Windows', 'macOS')]
        [string]$Platform,
        
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $true)]
        $Value,
        
        [Parameter()]
        [ValidateSet('Machine', 'User')]
        [string]$Scope = 'Machine'
    )
    
    try {
        if (-not $script:PolicyMappings[$Platform][$Name]) {
            Write-Error "Unknown policy '$Name' for platform '$Platform'"
            return $false
        }
        
        $policyInfo = $script:PolicyMappings[$Platform][$Name]
        
        # Create descriptive target for ShouldProcess
        $target = "$Platform policy '$Name' (Scope: $Scope)"
        $action = "Set value to '$Value'"
        
        if ($PSCmdlet.ShouldProcess($target, $action)) {
            switch ($Platform) {
                'Windows' {
                    return Set-OPKWinPolicy -Name $policyInfo.Name -Value $Value -Scope $Scope -Type $policyInfo.Type
                }
                'macOS' {
                    return Set-OPKMacPolicy -Key $policyInfo.Key -Value $Value -Type $policyInfo.Type
                }
            }
        }
        else {
            Write-Verbose "Operation cancelled by user"
            return $false
        }
    }
    catch {
        Write-Error "Failed to set policy '$Name' on '$Platform': $($_.Exception.Message)"
        return $false
    }
}
function Get-OPKAvailablePolicies {
    <#
    .SYNOPSIS
        Gets the list of available policies for a platform
    .PARAMETER Platform
        The platform to get policies for
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('Windows', 'macOS')]
        [string]$Platform
    )
    
    if ($Platform) {
        return $script:PolicyMappings[$Platform].Keys | ForEach-Object {
            [PSCustomObject]@{
                Name = $_
                Platform = $Platform
                Description = $script:PolicyMappings[$Platform][$_].Description
                DefaultValue = $script:PolicyMappings[$Platform][$_].DefaultValue
            }
        }
    }
    else {
        $allPolicies = @()
        foreach ($plat in $script:PolicyMappings.Keys) {
            $policies = $script:PolicyMappings[$plat].Keys | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_
                    Platform = $plat
                    Description = $script:PolicyMappings[$plat][$_].Description
                    DefaultValue = $script:PolicyMappings[$plat][$_].DefaultValue
                }
            }
            $allPolicies += $policies
        }
        return $allPolicies
    }
}

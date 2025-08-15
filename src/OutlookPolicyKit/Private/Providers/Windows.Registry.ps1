<#
.SYNOPSIS
    Windows Registry provider for Outlook policies
    
.DESCRIPTION
    Provides functions to get and set Outlook policies in Windows registry
    for both HKLM (machine) and HKCU (user) scopes.
#>

function Get-OPKWinPolicy {
    <#
    .SYNOPSIS
        Gets an Outlook policy from Windows registry
    .PARAMETER Name
        The policy name to retrieve
    .PARAMETER Scope
        Machine or User scope (defaults to Machine)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter()]
        [ValidateSet('Machine', 'User')]
        [string]$Scope = 'Machine'
    )
    
    try {
        $registryPath = if ($Scope -eq 'Machine') {
            'HKLM:\Software\Policies\Microsoft\Office\16.0\outlook'
        } else {
            'HKCU:\Software\Policies\Microsoft\Office\16.0\outlook'
        }
        
        # Policy mappings
        $policyMappings = @{
            'ExternalContent' = 'DisableExternalContent'
            'BlockHTTPImages' = 'BlockHTTPImages' 
            'CachedMode' = 'CachedExchangeMode'
            'SyncWindowSettingDays' = 'SyncWindowSetting'
        }
        
        $registryName = $policyMappings[$Name]
        if (-not $registryName) {
            Write-Error "Unknown policy name: $Name"
            return $null
        }
        
        if (Test-Path $registryPath) {
            $value = Get-ItemProperty -Path $registryPath -Name $registryName -ErrorAction SilentlyContinue
            if ($value) {
                return [PSCustomObject]@{
                    Name = $Name
                    Value = $value.$registryName
                    Scope = $Scope
                    Path = "$registryPath\$registryName"
                    Type = 'Registry'
                }
            }
        }
        
        return $null
    }
    catch {
        Write-Error "Failed to get policy '$Name': $($_.Exception.Message)"
        return $null
    }
}

function Set-OPKWinPolicy {
    <#
    .SYNOPSIS
        Sets an Outlook policy in Windows registry
    .PARAMETER Name
        The policy name to set
    .PARAMETER Value
        The value to set
    .PARAMETER Scope
        Machine or User scope (defaults to Machine)
    .PARAMETER Type
        Registry value type (String, DWORD, etc.)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $true)]
        $Value,
        
        [Parameter()]
        [ValidateSet('Machine', 'User')]
        [string]$Scope = 'Machine',
        
        [Parameter()]
        [ValidateSet('String', 'DWORD', 'Binary', 'QWord')]
        [string]$Type = 'DWORD'
    )
    
    try {
        $registryPath = if ($Scope -eq 'Machine') {
            'HKLM:\Software\Policies\Microsoft\Office\16.0\outlook'
        } else {
            'HKCU:\Software\Policies\Microsoft\Office\16.0\outlook'
        }
        
        # Policy mappings
        $policyMappings = @{
            'ExternalContent' = 'DisableExternalContent'
            'BlockHTTPImages' = 'BlockHTTPImages'
            'CachedMode' = 'CachedExchangeMode' 
            'SyncWindowSettingDays' = 'SyncWindowSetting'
        }
        
        $registryName = $policyMappings[$Name]
        if (-not $registryName) {
            Write-Error "Unknown policy name: $Name"
            return $false
        }
        
        # Ensure the registry path exists
        if (-not (Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }
        
        # Set the registry value with appropriate type
        switch ($Type) {
            'String' { 
                Set-ItemProperty -Path $registryPath -Name $registryName -Value $Value -Type String
            }
            'DWORD' {
                Set-ItemProperty -Path $registryPath -Name $registryName -Value ([int]$Value) -Type DWord
            }
            'Binary' {
                Set-ItemProperty -Path $registryPath -Name $registryName -Value $Value -Type Binary
            }
            'QWord' {
                Set-ItemProperty -Path $registryPath -Name $registryName -Value ([long]$Value) -Type QWord
            }
        }
        
        Write-Verbose "Set policy '$Name' = '$Value' in $Scope scope"
        return $true
    }
    catch {
        Write-Error "Failed to set policy '$Name': $($_.Exception.Message)"
        return $false
    }
}

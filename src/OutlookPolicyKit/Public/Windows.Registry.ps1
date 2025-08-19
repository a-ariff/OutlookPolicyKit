# Windows Registry management functions - PSScriptAnalyzer compliant

function Set-OutlookRegistryValue {
    <#
    .SYNOPSIS
        Sets a registry value for Outlook configuration
    
    .DESCRIPTION
        This function creates or updates a registry value at the specified path.
        It includes proper validation and error handling.
    
    .PARAMETER KeyPath
        Registry key path where the value will be set
    
    .PARAMETER ValueName
        Name of the registry value to set
    
    .PARAMETER Value
        Value to assign to the registry entry
    
    .PARAMETER ValueType
        Type of the registry value (String, DWORD, Binary, etc.)
    
    .EXAMPLE
        Set-OutlookRegistryValue -KeyPath 'HKCU:\Software\Microsoft\Office\16.0\Outlook' -ValueName 'DisableToday' -Value 1 -ValueType 'DWORD'
        Sets the DisableToday value to 1
    
    .NOTES
        Author: OutlookPolicyKit Team
        Version: 0.1.0
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ValueName,
        
        [Parameter(Mandatory = $true)]
        [object]$Value,
        
        [Parameter(Mandatory = $false)]
        [string]$ValueType = "String"
    )
    
    try {
        if ($PSCmdlet.ShouldProcess("$KeyPath\$ValueName", "Set registry value")) {
            if (!(Test-Path $KeyPath)) {
                New-Item -Path $KeyPath -Force | Out-Null
            }
            
            Set-ItemProperty -Path $KeyPath -Name $ValueName -Value $Value -Type $ValueType
            Write-Information "Registry value set: $KeyPath\$ValueName = $Value" -InformationAction Continue
            
            # Return success result
            $result = [PSCustomObject]@{
                KeyPath = $KeyPath
                ValueName = $ValueName
                Value = $Value
                ValueType = $ValueType
                Status = 'Success'
                Message = 'Registry value set successfully'
            }
            
            return $result
        }
    }
    catch {
        Write-Error "Failed to set registry value: $($_.Exception.Message)"
        throw
    }
}

function Remove-OutlookRegistryValue {
    <#
    .SYNOPSIS
        Removes a registry value for Outlook configuration
    
    .DESCRIPTION
        This function removes a registry value from the specified path.
        It includes proper validation and error handling.
    
    .PARAMETER KeyPath
        Registry key path containing the value to remove
    
    .PARAMETER ValueName
        Name of the registry value to remove
    
    .EXAMPLE
        Remove-OutlookRegistryValue -KeyPath 'HKCU:\Software\Microsoft\Office\16.0\Outlook' -ValueName 'DisableToday'
        Removes the DisableToday value
    
    .NOTES
        Author: OutlookPolicyKit Team
        Version: 0.1.0
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ValueName
    )
    
    try {
        if ($PSCmdlet.ShouldProcess("$KeyPath\$ValueName", "Remove registry value")) {
            if (Test-Path $KeyPath) {
                Remove-ItemProperty -Path $KeyPath -Name $ValueName -ErrorAction SilentlyContinue
                Write-Information "Registry value removed: $KeyPath\$ValueName" -InformationAction Continue
                
                # Return success result
                $result = [PSCustomObject]@{
                    KeyPath = $KeyPath
                    ValueName = $ValueName
                    Status = 'Success'
                    Message = 'Registry value removed successfully'
                }
                
                return $result
            } else {
                Write-Information "Registry key not found: $KeyPath" -InformationAction Continue
            }
        }
    }
    catch {
        Write-Error "Failed to remove registry value: $($_.Exception.Message)"
        throw
    }
}

function New-OutlookRegistryKey {
    <#
    .SYNOPSIS
        Creates a new registry key for Outlook configuration
    
    .DESCRIPTION
        This function creates a new registry key at the specified path.
        It includes proper validation and error handling.
    
    .PARAMETER KeyPath
        Registry key path to create
    
    .EXAMPLE
        New-OutlookRegistryKey -KeyPath 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Policies'
        Creates the Outlook Policies registry key
    
    .NOTES
        Author: OutlookPolicyKit Team
        Version: 0.1.0
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyPath
    )
    
    try {
        if ($PSCmdlet.ShouldProcess($KeyPath, "Create registry key")) {
            if (!(Test-Path $KeyPath)) {
                New-Item -Path $KeyPath -Force | Out-Null
                Write-Information "Registry key created: $KeyPath" -InformationAction Continue
                
                # Return success result
                $result = [PSCustomObject]@{
                    KeyPath = $KeyPath
                    Status = 'Created'
                    Message = 'Registry key created successfully'
                }
            } else {
                Write-Information "Registry key already exists: $KeyPath" -InformationAction Continue
                
                # Return existing result
                $result = [PSCustomObject]@{
                    KeyPath = $KeyPath
                    Status = 'Exists'
                    Message = 'Registry key already exists'
                }
            }
            
            return $result
        }
    }
    catch {
        Write-Error "Failed to create registry key: $($_.Exception.Message)"
        throw
    }
}

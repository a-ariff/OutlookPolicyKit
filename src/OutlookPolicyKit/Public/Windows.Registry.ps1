# Windows Registry management functions missing SupportsShouldProcess

function Set-OutlookRegistryValue {
    # Missing [CmdletBinding(SupportsShouldProcess)] - PSScriptAnalyzer violation
    [CmdletBinding()]
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
        # This function modifies system state but lacks SupportsShouldProcess
        if (!(Test-Path $KeyPath)) {
            New-Item -Path $KeyPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $KeyPath -Name $ValueName -Value $Value -Type $ValueType
        Write-Host "Registry value set: $KeyPath\$ValueName = $Value"
    }
    catch {
        Write-Error "Failed to set registry value: $_"
        throw
    }
}

function Remove-OutlookRegistryValue {
    # Missing [CmdletBinding(SupportsShouldProcess)] - PSScriptAnalyzer violation
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ValueName
    )
    
    try {
        # This function removes registry values but lacks SupportsShouldProcess
        if (Test-Path $KeyPath) {
            Remove-ItemProperty -Path $KeyPath -Name $ValueName -ErrorAction SilentlyContinue
            Write-Host "Registry value removed: $KeyPath\$ValueName"
        }
    }
    catch {
        Write-Error "Failed to remove registry value: $_"
        throw
    }
}

function New-OutlookRegistryKey {
    # Missing [CmdletBinding(SupportsShouldProcess)] - PSScriptAnalyzer violation
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyPath
    )
    
    try {
        # This function creates registry keys but lacks SupportsShouldProcess
        if (!(Test-Path $KeyPath)) {
            New-Item -Path $KeyPath -Force | Out-Null
            Write-Host "Registry key created: $KeyPath"
        }
        else {
            Write-Host "Registry key already exists: $KeyPath"
        }
    }
    catch {
        Write-Error "Failed to create registry key: $_"
        throw
    }
}

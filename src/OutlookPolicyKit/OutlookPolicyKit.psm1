#Requires -Version 5.1

<#
.SYNOPSIS
    OutlookPolicyKit PowerShell Module

.DESCRIPTION
    Intune Outlook automation toolkit for policy management and remediation
    
.NOTES
    Author: OutlookPolicyKit Team
    Version: 0.1.0
    Copyright: (c) 2025 OutlookPolicyKit Team. All rights reserved.
#>

# Get public and private function definition files.
$Public  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )

# Dot source the files
ForEach($import in @($Public + $Private))
{
    Try
    {
        . $import.FullName
    }
    Catch
    {
        Write-Error -Message "Failed to import function $($import.FullName): $($_.Exception.Message)"
    }
}

# Export the public functions
Export-ModuleMember -Function $Public.Basename

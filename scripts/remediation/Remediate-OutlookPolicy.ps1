#Requires -Version 7.0
<#
.SYNOPSIS
    Outlook Policy Remediation Script for Microsoft Intune

.DESCRIPTION
    This script enforces Outlook policy baselines by applying configured settings
    across Windows and macOS platforms. It integrates with the OutlookPolicyKit
    module to remediate policy drift and ensure compliance.

.PARAMETER BaselinePath
    Path to the baseline JSON configuration file

.PARAMETER LogPath
    Path to write remediation logs (optional)

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File Remediate-OutlookPolicy.ps1

.NOTES
    Exit Codes:
    0 = Success - All policies applied successfully
    1 = Failure - One or more policies failed to apply
    
    Author: OutlookPolicyKit
    Version: 1.0.0
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$BaselinePath,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

# Initialize exit code
$ExitCode = 0

try {
    # Set up logging
    if ($LogPath) {
        $LogFile = Join-Path $LogPath "OutlookPolicy-Remediation-$(Get-Date -Format 'yyyy-MM-dd-HHmmss').log"
        if (!(Test-Path (Split-Path $LogFile -Parent))) {
            New-Item -ItemType Directory -Path (Split-Path $LogFile -Parent) -Force | Out-Null
        }
    }
    
    function Write-Log {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,
            [ValidateSet('INFO', 'WARNING', 'ERROR')]
            [string]$Level = 'INFO'
        )
        
        $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $LogMessage = "[$Timestamp] [$Level] $Message"
        
        Write-Output $LogMessage
        if ($LogFile) {
            Add-Content -Path $LogFile -Value $LogMessage -Encoding UTF8
        }
    }
    
    Write-Log "Starting Outlook Policy Remediation" -Level INFO
    Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)" -Level INFO
    Write-Log "Operating System: $($PSVersionTable.Platform ?? $PSVersionTable.OS)" -Level INFO
    
    # Determine baseline path
    if (-not $BaselinePath) {
        $ScriptDir = $PSScriptRoot
        $RepoRoot = Split-Path (Split-Path $ScriptDir -Parent) -Parent
        
        # Auto-detect platform and use appropriate baseline
        if ($IsWindows -or $PSVersionTable.PSVersion.Major -lt 6) {
            $BaselinePath = Join-Path $RepoRoot "examples/baselines/windows-secure.json"
            Write-Log "Auto-detected Windows platform" -Level INFO
        } elseif ($IsMacOS) {
            $BaselinePath = Join-Path $RepoRoot "examples/baselines/macos-secure.json"
            Write-Log "Auto-detected macOS platform" -Level INFO
        } else {
            Write-Log "Unable to detect platform, defaulting to Windows baseline" -Level WARNING
            $BaselinePath = Join-Path $RepoRoot "examples/baselines/windows-secure.json"
        }
    }
    
    Write-Log "Using baseline: $BaselinePath" -Level INFO
    
    # Check if baseline exists
    if (-not (Test-Path $BaselinePath)) {
        Write-Log "Baseline file not found: $BaselinePath" -Level ERROR
        $ExitCode = 1
        exit $ExitCode
    }
    
    # Load baseline configuration
    try {
        $Baseline = Get-Content -Path $BaselinePath -Raw | ConvertFrom-Json
        Write-Log "Successfully loaded baseline configuration" -Level INFO
        Write-Log "Baseline Name: $($Baseline.metadata.name)" -Level INFO
        Write-Log "Baseline Version: $($Baseline.metadata.version)" -Level INFO
        Write-Log "Policy Count: $($Baseline.policies.Count)" -Level INFO
    }
    catch {
        Write-Log "Failed to parse baseline JSON: $($_.Exception.Message)" -Level ERROR
        $ExitCode = 1
        exit $ExitCode
    }
    
    # Try to import OutlookPolicyKit module
    $ModulePath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "src/OutlookPolicyKit/OutlookPolicyKit.psd1"
    if (Test-Path $ModulePath) {
        try {
            Import-Module $ModulePath -Force
            Write-Log "Successfully imported OutlookPolicyKit module" -Level INFO
        }
        catch {
            Write-Log "Failed to import OutlookPolicyKit module: $($_.Exception.Message)" -Level ERROR
            $ExitCode = 1
            exit $ExitCode
        }
    } else {
        Write-Log "OutlookPolicyKit module not found at: $ModulePath" -Level ERROR
        $ExitCode = 1
        exit $ExitCode
    }
    
    # Apply remediation
    $Results = @()
    $SuccessCount = 0
    $FailureCount = 0
    
    Write-Log "Starting policy remediation..." -Level INFO
    
    try {
        # Use the Invoke-OPKRemediation cmdlet if available
        if (Get-Command "Invoke-OPKRemediation" -ErrorAction SilentlyContinue) {
            Write-Log "Using Invoke-OPKRemediation cmdlet" -Level INFO
            $RemediationResult = Invoke-OPKRemediation -BaselinePath $BaselinePath -Force -Verbose
            
            if ($RemediationResult.Success) {
                Write-Log "Remediation completed successfully" -Level INFO
                Write-Log "Applied Policies: $($RemediationResult.AppliedCount)" -Level INFO
                Write-Log "Failed Policies: $($RemediationResult.FailedCount)" -Level INFO
                
                if ($RemediationResult.FailedCount -gt 0) {
                    $ExitCode = 1
                    Write-Log "Some policies failed to apply" -Level ERROR
                }
            } else {
                Write-Log "Remediation failed: $($RemediationResult.Message)" -Level ERROR
                $ExitCode = 1
            }
        } else {
            # Fallback to basic policy application
            Write-Log "Invoke-OPKRemediation not available, using fallback method" -Level WARNING
            
            foreach ($Policy in $Baseline.policies) {
                try {
                    Write-Log "Applying policy: $($Policy.name)" -Level INFO
                    
                    # Build policy settings hashtable
                    $PolicySettings = @{}
                    $PolicySettings[$Policy.name] = $Policy.value
                    
                    # Apply the policy
                    $SetResult = Set-OPKOutlookPolicy -PolicySettings $PolicySettings -Verbose
                    
                    if ($SetResult) {
                        Write-Log "Successfully applied policy: $($Policy.name)" -Level INFO
                        $SuccessCount++
                    } else {
                        Write-Log "Failed to apply policy: $($Policy.name)" -Level ERROR
                        $FailureCount++
                    }
                }
                catch {
                    Write-Log "Error applying policy $($Policy.name): $($_.Exception.Message)" -Level ERROR
                    $FailureCount++
                }
            }
            
            if ($FailureCount -gt 0) {
                $ExitCode = 1
            }
        }
    }
    catch {
        Write-Log "Unexpected error during remediation: $($_.Exception.Message)" -Level ERROR
        $ExitCode = 1
    }
    
    # Generate summary
    Write-Log "=== REMEDIATION SUMMARY ===" -Level INFO
    Write-Log "Total Policies: $($Baseline.policies.Count)" -Level INFO
    Write-Log "Success Count: $SuccessCount" -Level INFO
    Write-Log "Failure Count: $FailureCount" -Level INFO
    Write-Log "Exit Code: $ExitCode" -Level INFO
    
    if ($LogFile) {
        Write-Log "Log file: $LogFile" -Level INFO
    }
}
catch {
    Write-Output "CRITICAL ERROR: $($_.Exception.Message)"
    if ($LogFile) {
        Add-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] CRITICAL ERROR: $($_.Exception.Message)" -Encoding UTF8
    }
    $ExitCode = 1
}
finally {
    # Ensure we exit with proper code
    Write-Output "Remediation completed with exit code: $ExitCode"
    exit $ExitCode
}

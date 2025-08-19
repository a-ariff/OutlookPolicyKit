function Invoke-OPKRemediation {
    <#
    .SYNOPSIS
        Performs policy remediation using baseline comparison and enforcement
    
    .DESCRIPTION
        This function compares current Outlook policy settings against a baseline,
        identifies non-compliant settings, and optionally enforces compliance.
        Creates detailed logs and returns comprehensive remediation reports.
    
    .PARAMETER BaselinePath
        Path to the baseline JSON file for comparison and remediation
    
    .PARAMETER Platform
        Target platform (Windows or macOS). Auto-detected if not specified.
    
    .PARAMETER Scope
        For Windows: Machine or User scope (default: Machine)
    
    .PARAMETER Enforce
        When specified, applies remediation to bring policies into compliance
    
    .PARAMETER LogPath
        Custom path for log files. Defaults to $env:ProgramData/OutlookPolicyKit/logs
    
    .PARAMETER PassThru
        Returns the remediation report object for further processing
    
    .EXAMPLE
        Invoke-OPKRemediation -BaselinePath 'C:\baselines\windows-secure.json'
        Performs compliance check without enforcement
    
    .EXAMPLE
        Invoke-OPKRemediation -BaselinePath 'C:\baselines\windows-secure.json' -Enforce
        Performs compliance check and enforces remediation for non-compliant policies
    
    .EXAMPLE
        $result = Invoke-OPKRemediation -BaselinePath './baseline.json' -Enforce -PassThru
        Performs remediation and captures the result for script processing
    
    .NOTES
        Author: OutlookPolicyKit Team
        Version: 0.2.0
        Requires: Appropriate privileges for the target scope when using -Enforce
        
        Exit Codes:
        0 - Success (all policies compliant or successfully remediated)
        1 - Non-compliance found (when not using -Enforce)
        2 - Remediation failures occurred
        3 - Critical error (baseline file issues, permissions, etc.)
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$BaselinePath,
        
        [Parameter()]
        [ValidateSet('Windows', 'macOS')]
        [string]$Platform,
        
        [Parameter()]
        [ValidateSet('Machine', 'User')]
        [string]$Scope = 'Machine',
        
        [Parameter()]
        [switch]$Enforce,
        
        [Parameter()]
        [string]$LogPath,
        
        [Parameter()]
        [switch]$PassThru
    )
    
    begin {
        Write-Verbose "Starting Invoke-OPKRemediation function"
        
        # Auto-detect platform if not specified
        if (-not $Platform) {
            $Platform = if ($IsWindows) { 'Windows' } elseif ($IsMacOS) { 'macOS' } else { 'Windows' }
            Write-Verbose "Auto-detected platform: $Platform"
        }
        
        # Set up logging path
        if (-not $LogPath) {
            $LogPath = if ($Platform -eq 'Windows') {
                "$env:ProgramData\OutlookPolicyKit\logs"
            } else {
                "/var/log/OutlookPolicyKit"
            }
        }
        
        # Ensure log directory exists
        try {
            if (-not (Test-Path $LogPath)) {
                New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
                Write-Verbose "Created log directory: $LogPath"
            }
        } catch {
            Write-Warning "Could not create log directory $LogPath : $($_.Exception.Message)"
            $LogPath = $env:TEMP  # Fallback to temp directory
        }
        
        # Initialize tracking variables
        $timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
        $logFileName = "OPK-Remediation-$timestamp.json"
        $logFilePath = Join-Path $LogPath $logFileName
        
        $complianceResults = @()
        $remediationResults = @()
        $exitCode = 0
    }
    
    process {
        try {
            Write-Information "Starting policy remediation against baseline: $BaselinePath" -InformationAction Continue
            
            # Load and validate baseline
            try {
                $baseline = Get-Content -Path $BaselinePath -Raw | ConvertFrom-Json
                if (-not $baseline.policies) {
                    throw "Invalid baseline file: missing 'policies' section"
                }
                Write-Verbose "Loaded baseline: $($baseline.metadata.name) v$($baseline.metadata.version)"
            } catch {
                Write-Error "Failed to load baseline: $($_.Exception.Message)"
                $exitCode = 3
                return
            }
            
            Write-Information "Baseline loaded: $($baseline.metadata.name) v$($baseline.metadata.version)" -InformationAction Continue
            
            # Phase 1: Compliance Assessment
            Write-Information "Phase 1: Assessing policy compliance..." -InformationAction Continue
            
            $totalPolicies = 0
            $compliantPolicies = 0
            $nonCompliantPolicies = 0
            $missingPolicies = 0
            
            foreach ($policy in $baseline.policies.PSObject.Properties) {
                $policyName = $policy.Name
                $expectedValue = $policy.Value.value
                $description = $policy.Value.description
                $severity = $policy.Value.severity
                $autoRemediate = $policy.Value.autoRemediate -eq $true
                
                $totalPolicies++
                Write-Verbose "Checking policy: $policyName"
                
                # Get current value
                $current = Get-OPKPolicy -Platform $Platform -Name $policyName -Scope $Scope
                $currentValue = if ($current) { $current.Value } else { $null }
                
                # Determine compliance
                $compliant = if ($current) {
                    $currentValue -eq $expectedValue
                } else {
                    $false
                }
                
                if ($compliant) {
                    $compliantPolicies++
                    $status = 'Compliant'
                } elseif ($current) {
                    $nonCompliantPolicies++
                    $status = 'Non-Compliant'
                } else {
                    $missingPolicies++
                    $status = 'Missing'
                }
                
                $complianceResults += [PSCustomObject]@{
                    PolicyName = $policyName
                    Description = $description
                    Severity = $severity
                    CurrentValue = $currentValue
                    ExpectedValue = $expectedValue
                    Status = $status
                    AutoRemediate = $autoRemediate
                    Timestamp = Get-Date -Format 'o'
                }
            }
            
            # Phase 2: Remediation (if -Enforce is specified)
            if ($Enforce) {
                if ($PSCmdlet.ShouldProcess("Outlook Policy Settings", "Apply baseline remediation")) {
                    Write-Information "Phase 2: Enforcing policy compliance..." -InformationAction Continue
                    
                    $remediationAttempts = 0
                    $remediationSuccesses = 0
                    $remediationFailures = 0
                    
                    foreach ($policy in $complianceResults | Where-Object { $_.Status -ne 'Compliant' }) {
                        $remediationAttempts++
                        
                        if ($policy.AutoRemediate) {
                            Write-Verbose "Remediating policy: $($policy.PolicyName) = $($policy.ExpectedValue)"
                            
                            try {
                                $success = Set-OPKPolicy -Platform $Platform -Name $policy.PolicyName -Value $policy.ExpectedValue -Scope $Scope
                                
                                if ($success) {
                                    $remediationSuccesses++
                                    Write-Information "✓ Remediated: $($policy.PolicyName)" -InformationAction Continue
                                    
                                    $remediationResults += [PSCustomObject]@{
                                        PolicyName = $policy.PolicyName
                                        Action = 'Remediated'
                                        OldValue = $policy.CurrentValue
                                        NewValue = $policy.ExpectedValue
                                        Success = $true
                                        Message = 'Policy successfully remediated'
                                        Timestamp = Get-Date -Format 'o'
                                    }
                                } else {
                                    $remediationFailures++
                                    Write-Warning "✗ Failed to remediate: $($policy.PolicyName)"
                                    
                                    $remediationResults += [PSCustomObject]@{
                                        PolicyName = $policy.PolicyName
                                        Action = 'Remediation Attempted'
                                        OldValue = $policy.CurrentValue
                                        NewValue = $policy.ExpectedValue
                                        Success = $false
                                        Message = 'Remediation failed'
                                        Timestamp = Get-Date -Format 'o'
                                    }
                                }
                            } catch {
                                $remediationFailures++
                                Write-Error "Error remediating $($policy.PolicyName): $($_.Exception.Message)"
                                
                                $remediationResults += [PSCustomObject]@{
                                    PolicyName = $policy.PolicyName
                                    Action = 'Remediation Error'
                                    OldValue = $policy.CurrentValue
                                    NewValue = $policy.ExpectedValue
                                    Success = $false
                                    Message = $_.Exception.Message
                                    Timestamp = Get-Date -Format 'o'
                                }
                            }
                        } else {
                            Write-Warning "⚠ Policy $($policy.PolicyName) is non-compliant but auto-remediation is disabled"
                            
                            $remediationResults += [PSCustomObject]@{
                                PolicyName = $policy.PolicyName
                                Action = 'Skipped'
                                OldValue = $policy.CurrentValue
                                NewValue = $policy.ExpectedValue
                                Success = $false
                                Message = 'Auto-remediation disabled for this policy'
                                Timestamp = Get-Date -Format 'o'
                            }
                        }
                    }
                    
                    # Set exit code based on remediation results
                    if ($remediationFailures -gt 0) {
                        $exitCode = 2  # Remediation failures occurred
                    } elseif ($nonCompliantPolicies + $missingPolicies -gt $remediationSuccesses) {
                        $exitCode = 1  # Some non-compliance remains
                    } else {
                        $exitCode = 0  # Success
                    }
                } else {
                    Write-Information "Remediation cancelled by user" -InformationAction Continue
                    $exitCode = 1
                }
            } else {
                # Not enforcing - set exit code based on compliance
                if ($nonCompliantPolicies -gt 0 -or $missingPolicies -gt 0) {
                    $exitCode = 1  # Non-compliance found
                } else {
                    $exitCode = 0  # All compliant
                }
            }
            
        } catch {
            Write-Error "Critical error during remediation: $($_.Exception.Message)"
            $exitCode = 3
        }
    }
    
    end {
        try {
            # Create comprehensive report
            $report = [PSCustomObject]@{
                Metadata = @{
                    Timestamp = Get-Date -Format 'o'
                    BaselinePath = $BaselinePath
                    BaselineName = $baseline.metadata.name
                    BaselineVersion = $baseline.metadata.version
                    Platform = $Platform
                    Scope = $Scope
                    EnforcementMode = $Enforce.IsPresent
                    ExitCode = $exitCode
                }
                Summary = @{
                    TotalPolicies = $totalPolicies
                    CompliantPolicies = $compliantPolicies
                    NonCompliantPolicies = $nonCompliantPolicies
                    MissingPolicies = $missingPolicies
                    RemediationAttempts = if ($Enforce) { $remediationAttempts } else { 0 }
                    RemediationSuccesses = if ($Enforce) { $remediationSuccesses } else { 0 }
                    RemediationFailures = if ($Enforce) { $remediationFailures } else { 0 }
                }
                ComplianceResults = $complianceResults
                RemediationResults = $remediationResults
            }
            
            # Write log file
            try {
                $report | ConvertTo-Json -Depth 10 | Out-File -FilePath $logFilePath -Encoding UTF8
                Write-Information "Remediation report saved: $logFilePath" -InformationAction Continue
            } catch {
                Write-Warning "Could not write log file: $($_.Exception.Message)"
            }
            
            # Output summary
            if ($Enforce) {
                Write-Information "Remediation completed: $remediationSuccesses/$remediationAttempts policies remediated" -InformationAction Continue
            } else {
                Write-Information "Compliance check completed: $compliantPolicies/$totalPolicies policies compliant" -InformationAction Continue
            }
            
            Write-Verbose "Completed Invoke-OPKRemediation function with exit code: $exitCode"
            
            # Set exit code for script scenarios
            if (-not $PassThru) {
                exit $exitCode
            }
            
            # Return report if PassThru is requested
            if ($PassThru) {
                return $report
            }
        } catch {
            Write-Error "Error in cleanup: $($_.Exception.Message)"
            exit 3
        }
    }
}

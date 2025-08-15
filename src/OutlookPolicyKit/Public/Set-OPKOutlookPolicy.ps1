function Set-OPKOutlookPolicy {
    <#
    .SYNOPSIS
        Sets Outlook policy settings using ProviderRouter
    
    .DESCRIPTION
        This function applies Outlook policy settings using the ProviderRouter.
        Supports individual policy settings via -Name/-Value or bulk application
        from baseline files. Includes -WhatIf support for testing changes.
    
    .PARAMETER Name
        The name of a specific policy to set
    
    .PARAMETER Value
        The value to set for the specified policy
    
    .PARAMETER BaselinePath
        Path to a baseline JSON file to apply all policies from
    
    .PARAMETER Platform
        Target platform (Windows or macOS). Auto-detected if not specified.
    
    .PARAMETER Scope
        For Windows: Machine or User scope (default: Machine)
    
    .PARAMETER WhatIf
        Shows what changes would be made without actually applying them
    
    .PARAMETER Confirm
        Prompts for confirmation before making changes
    
    .EXAMPLE
        Set-OPKOutlookPolicy -Name 'DisableExternalImages' -Value $true
        Sets a specific policy value
    
    .EXAMPLE
        Set-OPKOutlookPolicy -BaselinePath 'C:\baselines\windows-secure.json' -WhatIf
        Shows what policies would be applied from the baseline without making changes
    
    .EXAMPLE
        Set-OPKOutlookPolicy -Name 'CachedMode' -Value $true -Platform Windows -Scope User
        Sets cached mode for the current user on Windows
    
    .NOTES
        Author: OutlookPolicyKit Team
        Version: 0.2.0
        Requires: Appropriate privileges for the target scope
    #>
    
    [CmdletBinding(DefaultParameterSetName = 'SinglePolicy', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(ParameterSetName = 'SinglePolicy', Mandatory = $true)]
        [string]$Name,
        
        [Parameter(ParameterSetName = 'SinglePolicy', Mandatory = $true)]
        $Value,
        
        [Parameter(ParameterSetName = 'Baseline', Mandatory = $true)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$BaselinePath,
        
        [Parameter()]
        [ValidateSet('Windows', 'macOS')]
        [string]$Platform,
        
        [Parameter()]
        [ValidateSet('Machine', 'User')]
        [string]$Scope = 'Machine'
    )
    
    begin {
        Write-Verbose "Starting Set-OPKOutlookPolicy function"
        
        # Auto-detect platform if not specified
        if (-not $Platform) {
            $Platform = if ($IsWindows) { 'Windows' } elseif ($IsMacOS) { 'macOS' } else { 'Windows' }
            Write-Verbose "Auto-detected platform: $Platform"
        }
        
        # Initialize results tracking
        $results = @()
        $successCount = 0
        $failureCount = 0
    }
    
    process {
        try {
            switch ($PSCmdlet.ParameterSetName) {
                'SinglePolicy' {
                    Write-Verbose "Setting single policy: $Name = $Value"
                    
                    if ($PSCmdlet.ShouldProcess("$Platform policy '$Name'", "Set value to '$Value'")) {
                        $success = Set-OPKPolicy -Platform $Platform -Name $Name -Value $Value -Scope $Scope
                        
                        if ($success) {
                            $successCount++
                            Write-Information "Successfully set $Name = $Value" -InformationAction Continue
                            
                            $results += [PSCustomObject]@{
                                PolicyName = $Name
                                Value = $Value
                                Platform = $Platform
                                Scope = $Scope
                                Status = 'Success'
                                Message = 'Policy applied successfully'
                            }
                        } else {
                            $failureCount++
                            Write-Warning "Failed to set $Name = $Value"
                            
                            $results += [PSCustomObject]@{
                                PolicyName = $Name
                                Value = $Value
                                Platform = $Platform
                                Scope = $Scope
                                Status = 'Failed'
                                Message = 'Failed to apply policy'
                            }
                        }
                    } else {
                        Write-Information "Would set $Name = $Value (WhatIf)" -InformationAction Continue
                        $results += [PSCustomObject]@{
                            PolicyName = $Name
                            Value = $Value
                            Platform = $Platform
                            Scope = $Scope
                            Status = 'WhatIf'
                            Message = 'Would apply this policy'
                        }
                    }
                }
                
                'Baseline' {
                    Write-Verbose "Processing baseline from: $BaselinePath"
                    
                    # Load baseline
                    $baseline = Get-Content -Path $BaselinePath -Raw | ConvertFrom-Json
                    
                    # Validate baseline structure
                    if (-not $baseline.policies) {
                        throw "Invalid baseline file: missing 'policies' section"
                    }
                    
                    Write-Information "Applying baseline: $($baseline.metadata.name) v$($baseline.metadata.version)" -InformationAction Continue
                    
                    foreach ($policy in $baseline.policies.PSObject.Properties) {
                        $policyName = $policy.Name
                        $policyValue = $policy.Value.value
                        $severity = $policy.Value.severity
                        
                        Write-Verbose "Processing policy: $policyName = $policyValue (Severity: $severity)"
                        
                        $shouldProcessMessage = "Apply baseline policy '$policyName' = '$policyValue'"
                        if ($PSCmdlet.ShouldProcess("$Platform baseline policies", $shouldProcessMessage)) {
                            $success = Set-OPKPolicy -Platform $Platform -Name $policyName -Value $policyValue -Scope $Scope
                            
                            if ($success) {
                                $successCount++
                                Write-Verbose "Successfully applied $policyName = $policyValue"
                                
                                $results += [PSCustomObject]@{
                                    PolicyName = $policyName
                                    Value = $policyValue
                                    Platform = $Platform
                                    Scope = $Scope
                                    Severity = $severity
                                    Status = 'Success'
                                    Message = 'Policy applied successfully'
                                }
                            } else {
                                $failureCount++
                                Write-Warning "Failed to apply $policyName = $policyValue"
                                
                                $results += [PSCustomObject]@{
                                    PolicyName = $policyName
                                    Value = $policyValue
                                    Platform = $Platform
                                    Scope = $Scope
                                    Severity = $severity
                                    Status = 'Failed'
                                    Message = 'Failed to apply policy'
                                }
                            }
                        } else {
                            Write-Information "Would apply $policyName = $policyValue (WhatIf)" -InformationAction Continue
                            $results += [PSCustomObject]@{
                                PolicyName = $policyName
                                Value = $policyValue
                                Platform = $Platform
                                Scope = $Scope
                                Severity = $severity
                                Status = 'WhatIf'
                                Message = 'Would apply this policy'
                            }
                        }
                    }
                }
            }
        }
        catch {
            Write-Error "Failed to set Outlook policies: $($_.Exception.Message)"
            return $null
        }
    }
    
    end {
        # Output summary
        if (-not $WhatIfPreference) {
            Write-Information "Policy application completed: $successCount succeeded, $failureCount failed" -InformationAction Continue
        }
        
        Write-Verbose "Completed Set-OPKOutlookPolicy function"
        
        # Return results for further processing if needed
        return $results | Sort-Object PolicyName
    }
}

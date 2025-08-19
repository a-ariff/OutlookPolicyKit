function Get-OPKOutlookPolicy {
    <#
    .SYNOPSIS
        Retrieves Outlook policy settings and compares them against baselines
    
    .DESCRIPTION
        This function retrieves Outlook policy settings using the ProviderRouter
        and supports baseline comparison to show current vs expected values.
        Supports both Windows and macOS platforms.
    
    .PARAMETER Name
        The name of a specific policy to retrieve
    
    .PARAMETER BaselinePath
        Path to a baseline JSON file for comparison
    
    .PARAMETER Platform
        Target platform (Windows or macOS). Auto-detected if not specified.
    
    .PARAMETER Scope
        For Windows: Machine or User scope (default: Machine)
    
    .EXAMPLE
        Get-OPKOutlookPolicy -Name 'DisableExternalImages'
        Retrieves a specific policy setting
    
    .EXAMPLE
        Get-OPKOutlookPolicy -BaselinePath 'C:\baselines\windows-secure.json'
        Compares all policies in baseline against current settings
    
    .EXAMPLE
        Get-OPKOutlookPolicy -Platform macOS -Scope User
        Retrieves all available policies for macOS user scope
    
    .NOTES
        Author: OutlookPolicyKit Team
        Version: 0.2.0
    #>
    
    [CmdletBinding(DefaultParameterSetName = 'AllPolicies')]
    param(
        [Parameter(ParameterSetName = 'SinglePolicy', Mandatory = $true)]
        [string]$Name,
        
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
        Write-Verbose "Starting Get-OPKOutlookPolicy function"
        
        # Auto-detect platform if not specified
        if (-not $Platform) {
            $Platform = if ($IsWindows) { 'Windows' } elseif ($IsMacOS) { 'macOS' } else { 'Windows' }
            Write-Verbose "Auto-detected platform: $Platform"
        }
    }
    
    process {
        try {
            switch ($PSCmdlet.ParameterSetName) {
                'SinglePolicy' {
                    Write-Verbose "Retrieving single policy: $Name"
                    $result = Get-OPKPolicy -Platform $Platform -Name $Name -Scope $Scope
                    return $result
                }
                
                'Baseline' {
                    Write-Verbose "Processing baseline comparison from: $BaselinePath"
                    
                    # Load baseline
                    $baseline = Get-Content -Path $BaselinePath -Raw | ConvertFrom-Json
                    
                    # Validate baseline structure
                    if (-not $baseline.policies) {
                        throw "Invalid baseline file: missing 'policies' section"
                    }
                    
                    $results = @()
                    
                    foreach ($policy in $baseline.policies.PSObject.Properties) {
                        $policyName = $policy.Name
                        $expectedValue = $policy.Value.value
                        $description = $policy.Value.description
                        $severity = $policy.Value.severity
                        
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
                        
                        $results += [PSCustomObject]@{
                            PolicyName = $policyName
                            Description = $description
                            Severity = $severity
                            CurrentValue = $currentValue
                            ExpectedValue = $expectedValue
                            Compliant = $compliant
                            Status = if ($compliant) { 'OK' } elseif ($current) { 'Drift' } else { 'Missing' }
                            Platform = $Platform
                            Scope = $Scope
                        }
                    }
                    
                    # Return formatted results
                    Write-Verbose "Returning $($results.Count) policy comparison results"
                    return $results | Sort-Object PolicyName
                }
                
                'AllPolicies' {
                    Write-Verbose "Retrieving all available policies for platform: $Platform"
                    
                    # Get all available policies
                    $availablePolicies = Get-OPKAvailablePolicies -Platform $Platform
                    
                    $results = @()
                    foreach ($policy in $availablePolicies) {
                        $current = Get-OPKPolicy -Platform $Platform -Name $policy.Name -Scope $Scope
                        
                        $results += [PSCustomObject]@{
                            PolicyName = $policy.Name
                            Description = $policy.Description
                            CurrentValue = if ($current) { $current.Value } else { $null }
                            DefaultValue = $policy.DefaultValue
                            Platform = $Platform
                            Scope = $Scope
                            IsSet = $null -ne $current
                        }
                    }
                    
                    return $results | Sort-Object PolicyName
                }
            }
        }
        catch {
            Write-Error "Failed to retrieve Outlook policies: $($_.Exception.Message)"
            return $null
        }
    }
    
    end {
        Write-Verbose "Completed Get-OPKOutlookPolicy function"
    }
}

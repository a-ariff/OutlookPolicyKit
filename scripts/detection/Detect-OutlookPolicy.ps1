<#
.SYNOPSIS
    Detection script for Outlook policy compliance

.DESCRIPTION
    This script imports the OutlookPolicyKit module and runs detection against a baseline.
    Returns JSON output and sets appropriate exit codes for Intune compliance.

.PARAMETER BaselinePath
    Path to the baseline JSON file for detection

.EXAMPLE
    .\Detect-OutlookPolicy.ps1 -BaselinePath "C:\baselines\windows-secure.json"
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$BaselinePath
)

try {
    # Import the OutlookPolicyKit module
    $ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\..\src\OutlookPolicyKit\OutlookPolicyKit.psd1"
    if (Test-Path $ModulePath) {
        Import-Module $ModulePath -Force
    } else {
        # Try finding in PSModulePath
        Import-Module OutlookPolicyKit -Force
    }

    # Run policy detection
    $Results = Get-OPKOutlookPolicy -BaselinePath $BaselinePath
    
    # Count non-compliant policies
    $NonCompliant = @($Results | Where-Object { $_.Compliant -eq $false })
    $TotalPolicies = @($Results).Count
    
    # Prepare output object
    $Output = @{
        Timestamp = Get-Date -Format 'yyyy-MM-ddTHH:mm:ss.fffZ'
        BaselinePath = $BaselinePath
        TotalPolicies = $TotalPolicies
        NonCompliantCount = $NonCompliant.Count
        ComplianceRate = if ($TotalPolicies -gt 0) { [Math]::Round((($TotalPolicies - $NonCompliant.Count) / $TotalPolicies) * 100, 2) } else { 100 }
        OverallCompliant = $NonCompliant.Count -eq 0
        PolicyDetails = $Results | ForEach-Object {
            @{
                PolicyName = $_.PolicyName
                Description = $_.Description
                Severity = $_.Severity
                CurrentValue = $_.CurrentValue
                ExpectedValue = $_.ExpectedValue
                Compliant = $_.Compliant
                Status = $_.Status
                Platform = $_.Platform
                Scope = $_.Scope
            }
        }
    }
    
    # Output JSON
    $JsonOutput = $Output | ConvertTo-Json -Depth 10
    Write-Output $JsonOutput
    
    # Set exit code based on compliance
    if ($Output.OverallCompliant) {
        Write-Host "All policies are compliant" -ForegroundColor Green
        exit 0  # Success - compliant
    } else {
        Write-Warning "$($NonCompliant.Count) of $TotalPolicies policies are non-compliant"
        exit 1  # Non-compliant detected
    }
}
catch {
    # Error occurred during detection
    $ErrorOutput = @{
        Timestamp = Get-Date -Format 'yyyy-MM-ddTHH:mm:ss.fffZ'
        BaselinePath = $BaselinePath
        Error = $_.Exception.Message
        ErrorType = "DetectionError"
        OverallCompliant = $false
    }
    
    $ErrorJson = $ErrorOutput | ConvertTo-Json -Depth 5
    Write-Output $ErrorJson
    Write-Error "Detection failed: $($_.Exception.Message)"
    exit 2  # Error during detection
}

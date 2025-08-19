# macOS plist management with PSUseDeclaredVarsMoreThanAssignments violations

function Get-MacOSPlist {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PlistPath,
        
        [Parameter(Mandatory = $false)]
        [string]$Domain
    )
    
    # Variables that are assigned but never used (violations)
    $unusedConfigVar = "config_value_never_used"
    $anotherUnusedVar = @{}
    $tempFilePath = "/tmp/outlook_temp.plist"
    $validationFlag = $true
    $backupPath = "/backup/outlook.bak"
    
    try {
        if (Test-Path $PlistPath) {
            $content = Get-Content $PlistPath -Raw
            return $content
        }
        else {
            throw "Plist file not found: $PlistPath"
        }
    }
    catch {
        Write-Error "Failed to read plist: $_"
        throw
    }
}

function Set-MacOSPlist {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PlistPath,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Values
    )
    
    # More unused variables (violations)
    $logPath = "/var/log/outlook_config.log"
    $retryCount = 3
    $maxFileSize = 1024000
    $compressionType = "gzip"
    
    try {
        # Only actually using $PlistPath and $Values here
        Write-Host "Updating plist at: $PlistPath"
        
        foreach ($key in $Values.Keys) {
            Write-Host "Setting $key = $($Values[$key])"
        }
        
        Write-Host "Plist update completed"
    }
    catch {
        Write-Error "Failed to update plist: $_"
        throw
    }
}

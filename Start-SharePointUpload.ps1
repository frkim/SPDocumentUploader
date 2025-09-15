<#
.SYNOPSIS
    SharePoint Document Uploader - Main PowerShell Script

.DESCRIPTION
    Orchestrates the complete SharePoint Online document upload process.
    Supports file scanning, authentication, upload with progress tracking,
    error handling, and comprehensive logging for DevOps teams.

.PARAMETER ConfigPath
    Path to JSON configuration file. Defaults to 'config.json'

.PARAMETER SourcePath
    Local source directory to scan for files. Overrides config setting.

.PARAMETER TargetFolder
    SharePoint target folder path. Overrides config setting.

.PARAMETER DryRun
    Scan and validate files without uploading to SharePoint.

.PARAMETER Verbose
    Enable verbose logging output.

.PARAMETER Extensions
    Comma-separated list of file extensions to include (e.g., '.pdf,.docx')

.PARAMETER MaxSizeMB
    Maximum file size in MB to process.

.PARAMETER BatchSize
    Number of files to process per batch.

.PARAMETER Overwrite
    Overwrite existing files in SharePoint.

.PARAMETER ExportResults
    Export upload results to CSV file.

.PARAMETER Force
    Skip confirmation prompts in interactive mode.

.EXAMPLE
    .\Start-SharePointUpload.ps1
    # Use default configuration

.EXAMPLE
    .\Start-SharePointUpload.ps1 -SourcePath "\\server\share\docs" -DryRun -Verbose
    # Scan network drive with verbose output

.EXAMPLE
    .\Start-SharePointUpload.ps1 -TargetFolder "/sites/team/Projects" -Extensions ".pdf,.docx" -Force
    # Upload specific file types to custom folder

.NOTES
    Author: DevOps Team
    Version: 1.0
    Requires: PowerShell 5.1+, PnP.PowerShell module
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "config.json",
    
    [Parameter(Mandatory = $false)]
    [string]$SourcePath,
    
    [Parameter(Mandatory = $false)]
    [string]$TargetFolder,
    
    [Parameter(Mandatory = $false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory = $false)]
    [switch]$Verbose,
    
    [Parameter(Mandatory = $false)]
    [string]$Extensions,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxSizeMB,
    
    [Parameter(Mandatory = $false)]
    [int]$BatchSize,
    
    [Parameter(Mandatory = $false)]
    [switch]$Overwrite,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportResults,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# Set error action preference for better error handling
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

# Import required modules
$ModulePath = Join-Path $PSScriptRoot "Modules"
Import-Module (Join-Path $ModulePath "SPConfig.psm1") -Force
Import-Module (Join-Path $ModulePath "SPAuth.psm1") -Force
Import-Module (Join-Path $ModulePath "SPFileScanner.psm1") -Force
Import-Module (Join-Path $ModulePath "SPUploader.psm1") -Force
Import-Module (Join-Path $ModulePath "SPLogger.psm1") -Force

# Global variables
$Global:SPLogger = $null
$Global:SPAuth = $null
$Global:ExitCode = 0

function Main {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "SharePoint Document Uploader - PowerShell Edition" -ForegroundColor Cyan
        Write-Host "=" * 60 -ForegroundColor Cyan
        
        # Load and validate configuration
        $config = Initialize-Configuration
        
        # Initialize logging
        $Global:SPLogger = Initialize-SPLogger -Config $config
        Write-SPInformation "SharePoint Document Uploader started"
        
        # Display configuration summary
        Show-ConfigurationSummary -Config $config
        
        # Validate environment and connectivity
        if (-not (Test-Environment -Config $config)) {
            throw "Environment validation failed"
        }
        
        # Scan for files
        $scanResult = Get-FilesToUpload -Config $config -SourcePath $SourcePath
        $files = $scanResult.Files
        $scanStats = $scanResult.Statistics
        
        if (-not $files -or $files.Count -eq 0) {
            Write-SPWarning "No files found to upload"
            return $true
        }
        
        # Display scan results
        Show-ScanResults -Files $files -Statistics $scanStats
        
        # Confirm upload (unless Force or DryRun)
        if (-not $DryRun -and -not $Force -and [System.Environment]::UserInteractive) {
            if (-not (Confirm-Upload -Files $files)) {
                Write-SPInformation "Upload cancelled by user"
                return $true
            }
        }
        
        # Perform upload or dry run
        if ($DryRun) {
            Write-SPInformation "DRY RUN: Files that would be uploaded:"
            foreach ($file in $files) {
                Write-SPInformation "  $($file.RelativePath) ($([math]::Round($file.Size / 1MB, 2)) MB)"
            }
            Write-SPInformation "Dry run completed - no files were uploaded"
        }
        else {
            # Perform actual upload
            $uploadResult = Start-Upload -Files $files -Config $config
            Show-UploadResults -Results $uploadResult.Results -Statistics $uploadResult.Statistics
            
            # Export results if requested
            if ($ExportResults) {
                Export-UploadResults -Results $uploadResult.Results -OutputPath $ExportResults
            }
            
            # Set exit code based on results
            if ($uploadResult.Statistics.FailedFiles -gt 0) {
                $Global:ExitCode = 1
            }
        }
        
        Write-SPInformation "SharePoint Document Uploader completed successfully"
        return $true
    }
    catch {
        Write-SPError "Fatal error occurred" -Exception $_.Exception
        Write-Error $_.Exception.Message
        $Global:ExitCode = 1
        return $false
    }
    finally {
        # Cleanup resources
        Cleanup-Resources
    }
}

function Initialize-Configuration {
    [CmdletBinding()]
    param()
    
    try {
        Write-Verbose "Loading configuration from: $ConfigPath"
        
        # Check if config file exists
        if (-not (Test-Path $ConfigPath)) {
            throw "Configuration file not found: $ConfigPath. Run New-SharePointConfig to create one."
        }
        
        # Load configuration
        $config = Get-SharePointConfig -ConfigPath $ConfigPath -ApplyEnvOverrides
        
        # Apply command line overrides
        if ($SourcePath) { $config.LocalSourcePath = $SourcePath }
        if ($TargetFolder) { $config.TargetFolderPath = $TargetFolder }
        if ($Extensions) { $config.FileExtensions = $Extensions -split "," | ForEach-Object { $_.Trim() } }
        if ($MaxSizeMB) { $config.MaxFileSizeMB = $MaxSizeMB }
        if ($BatchSize) { $config.BatchSize = $BatchSize }
        if ($Overwrite) { $config.OverwriteExisting = $true }
        if ($Verbose) { $config.LogLevel = "Verbose" }
        
        # Re-validate after overrides
        $config.Validate()
        
        return $config
    }
    catch {
        Write-Error "Configuration initialization failed: $($_.Exception.Message)"
        throw
    }
}

function Show-ConfigurationSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [SharePointConfig]$Config
    )
    
    $summary = $Config.GetSummary()
    
    Write-Host "`n=== CONFIGURATION SUMMARY ===" -ForegroundColor Yellow
    foreach ($key in $summary.Keys | Sort-Object) {
        $value = $summary[$key]
        if ($value -is [array]) {
            $value = $value -join ", "
        }
        Write-Host "$key`: $value" -ForegroundColor White
    }
    Write-Host "=" * 30 -ForegroundColor Yellow
}

function Test-Environment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [SharePointConfig]$Config
    )
    
    Write-SPInformation "Validating environment and connectivity..."
    
    try {
        # Check PowerShell version
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            throw "PowerShell 5.1 or higher is required"
        }
        Write-SPInformation "✓ PowerShell version: $($PSVersionTable.PSVersion)"
        
        # Check PnP PowerShell module
        $pnpModule = Get-Module -Name "PnP.PowerShell" -ListAvailable | Select-Object -First 1
        if (-not $pnpModule) {
            throw "PnP.PowerShell module is required. Install with: Install-Module PnP.PowerShell -Scope CurrentUser"
        }
        Write-SPInformation "✓ PnP PowerShell version: $($pnpModule.Version)"
        
        # Validate local paths
        if (-not (Test-Path $Config.LocalSourcePath)) {
            throw "Source path does not exist: $($Config.LocalSourcePath)"
        }
        Write-SPInformation "✓ Source path accessible: $($Config.LocalSourcePath)"
        
        # Test SharePoint connection (unless dry run)
        if (-not $DryRun) {
            $Global:SPAuth = Connect-SharePointSite -Config $Config -TestConnection
            $testResult = Test-SharePointAuth -Config $Config
            
            if (-not $testResult.Connection) {
                throw "SharePoint connection failed: $($testResult.ErrorMessage)"
            }
            
            Write-SPInformation "✓ SharePoint connection successful"
            Write-SPInformation "✓ Site access verified"
            
            if ($testResult.LibraryAccess) {
                Write-SPInformation "✓ Document library accessible"
            }
            
            if ($testResult.UploadPermissions) {
                Write-SPInformation "✓ Upload permissions verified"
            } else {
                Write-SPWarning "⚠ Upload permissions could not be verified"
            }
        }
        
        return $true
    }
    catch {
        Write-SPError "Environment validation failed" -Exception $_.Exception
        return $false
    }
}

function Get-FilesToUpload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [SharePointConfig]$Config,
        
        [Parameter(Mandatory = $false)]
        [string]$SourcePath
    )
    
    Write-SPInformation "Scanning for files..."
    
    $startTime = Get-Date
    
    # Initialize file scanner
    $scanner = New-FileScanner -Config $Config
    
    # Scan directory
    $sourcePath = if ($SourcePath) { $SourcePath } else { $Config.LocalSourcePath }
    $files = Get-DirectoryFiles -Scanner $scanner -Path $sourcePath
    
    # Apply filters
    $validFiles = $files | Where-Object { $_.IsValid }
    
    # Apply additional command-line filters
    if ($MaxSizeMB) {
        $maxBytes = $MaxSizeMB * 1MB
        $validFiles = $validFiles | Where-Object { $_.Size -le $maxBytes }
    }
    
    if ($Extensions) {
        $extList = $Extensions -split "," | ForEach-Object { $_.Trim().ToLower() }
        $validFiles = $validFiles | Where-Object { $extList -contains $_.Extension.ToLower() }
    }
    
    $endTime = Get-Date
    $scanDuration = ($endTime - $startTime).TotalSeconds
    
    # Calculate statistics
    $stats = @{
        TotalFiles = $files.Count
        ValidFiles = $validFiles.Count
        ValidSize = ($validFiles | Measure-Object -Property Size -Sum).Sum
        ScanDuration = $scanDuration
        FileTypeBreakdown = @{}
        SizeDistribution = @{
            "Small (< 1MB)" = 0
            "Medium (1-10MB)" = 0
            "Large (10-50MB)" = 0
            "Extra Large (> 50MB)" = 0
        }
    }
    
    # Calculate file type breakdown
    $validFiles | Group-Object Extension | ForEach-Object {
        $stats.FileTypeBreakdown[$_.Name] = @{
            Count = $_.Count
            TotalSize = ($_.Group | Measure-Object -Property Size -Sum).Sum
        }
    }
    
    # Calculate size distribution
    foreach ($file in $validFiles) {
        $sizeMB = $file.Size / 1MB
        if ($sizeMB -lt 1) {
            $stats.SizeDistribution["Small (< 1MB)"]++
        }
        elseif ($sizeMB -le 10) {
            $stats.SizeDistribution["Medium (1-10MB)"]++
        }
        elseif ($sizeMB -le 50) {
            $stats.SizeDistribution["Large (10-50MB)"]++
        }
        else {
            $stats.SizeDistribution["Extra Large (> 50MB)"]++
        }
    }
    
    Write-SPInformation "File scan completed in $([math]::Round($scanDuration, 2)) seconds"
    
    return @{
        Files = $validFiles
        Statistics = $stats
    }
}

function Show-ScanResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [FileInfo[]]$Files,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Statistics
    )
    
    Write-Host "`n=== SCAN RESULTS ===" -ForegroundColor Green
    Write-Host "Total files found: $($Statistics.TotalFiles)" -ForegroundColor White
    Write-Host "Valid files to upload: $($Files.Count)" -ForegroundColor White
    Write-Host "Files skipped: $($Statistics.TotalFiles - $Files.Count)" -ForegroundColor White
    Write-Host "Total size: $([math]::Round($Statistics.ValidSize / 1MB, 2)) MB" -ForegroundColor White
    
    # Show file type breakdown
    if ($Statistics.FileTypeBreakdown) {
        Write-Host "`nFiles by type:" -ForegroundColor Yellow
        foreach ($ext in $Statistics.FileTypeBreakdown.Keys | Sort-Object) {
            $info = $Statistics.FileTypeBreakdown[$ext]
            $sizeMB = [math]::Round($info.TotalSize / 1MB, 2)
            Write-Host "  $ext`: $($info.Count) files ($sizeMB MB)" -ForegroundColor White
        }
    }
    
    # Show size distribution
    if ($Statistics.SizeDistribution) {
        Write-Host "`nSize distribution:" -ForegroundColor Yellow
        foreach ($category in $Statistics.SizeDistribution.Keys) {
            Write-Host "  $category`: $($Statistics.SizeDistribution[$category]) files" -ForegroundColor White
        }
    }
    
    Write-Host "=" * 20 -ForegroundColor Green
}

function Confirm-Upload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [FileInfo[]]$Files
    )
    
    $totalSizeMB = ($Files | Measure-Object -Property Size -Sum).Sum / 1MB
    
    Write-Host "`nReady to upload $($Files.Count) files ($([math]::Round($totalSizeMB, 2)) MB)" -ForegroundColor Cyan
    
    do {
        $response = Read-Host "Do you want to proceed? (Y/N)"
        $response = $response.Trim().ToUpper()
        
        if ($response -eq "Y" -or $response -eq "YES") {
            return $true
        }
        elseif ($response -eq "N" -or $response -eq "NO") {
            return $false
        }
        else {
            Write-Host "Please enter Y (Yes) or N (No)" -ForegroundColor Yellow
        }
    } while ($true)
}

function Start-Upload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [FileInfo[]]$Files,
        
        [Parameter(Mandatory = $true)]
        [SharePointConfig]$Config
    )
    
    Write-SPInformation "Starting upload of $($Files.Count) files..."
    
    return Measure-SPOperation -OperationName "FileUpload" -ScriptBlock {
        Start-SharePointUpload -Files $Files -Config $Config -Auth $Global:SPAuth -TargetFolder $TargetFolder
    } -AdditionalMetrics @{
        FileCount = $Files.Count
        TotalSizeMB = [math]::Round(($Files | Measure-Object -Property Size -Sum).Sum / 1MB, 2)
    }
}

function Show-UploadResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Results,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Statistics
    )
    
    Write-SPUploadSummary -Statistics $Statistics
    
    # Show failed uploads
    $failedResults = $Results | Where-Object { -not $_.Success -and -not $_.Skipped }
    if ($failedResults) {
        Write-Host "`n=== FAILED UPLOADS ===" -ForegroundColor Red
        foreach ($result in $failedResults) {
            Write-Host "File: $($result.FileInfo.Name)" -ForegroundColor White
            Write-Host "  Error: $($result.ErrorMessage)" -ForegroundColor Red
            Write-Host "  Retries: $($result.RetryCount)" -ForegroundColor Yellow
            Write-Host "  Path: $($result.FileInfo.FullPath)" -ForegroundColor Gray
            Write-Host "-" * 40 -ForegroundColor Gray
        }
    }
}

function Export-UploadResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Results,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )
    
    try {
        $exportData = $Results | Select-Object @{
            Name = "FileName"; Expression = { $_.FileInfo.Name }
        }, @{
            Name = "FilePath"; Expression = { $_.FileInfo.FullPath }
        }, @{
            Name = "RelativePath"; Expression = { $_.FileInfo.RelativePath }
        }, @{
            Name = "FileSize"; Expression = { $_.FileInfo.Size }
        }, @{
            Name = "Success"; Expression = { $_.Success }
        }, @{
            Name = "Skipped"; Expression = { $_.Skipped }
        }, @{
            Name = "ErrorMessage"; Expression = { $_.ErrorMessage }
        }, @{
            Name = "SharePointUrl"; Expression = { $_.SharePointUrl }
        }, @{
            Name = "RetryCount"; Expression = { $_.RetryCount }
        }, @{
            Name = "UploadTime"; Expression = { $_.UploadTime }
        }
        
        $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-SPInformation "Upload results exported to: $OutputPath"
    }
    catch {
        Write-SPError "Failed to export results" -Exception $_.Exception
    }
}

function Cleanup-Resources {
    [CmdletBinding()]
    param()
    
    try {
        if ($Global:SPAuth) {
            $Global:SPAuth.Disconnect()
            $Global:SPAuth = $null
        }
        
        if ($Global:SPLogger) {
            Stop-SPLogger
        }
        
        Write-Verbose "Resources cleaned up successfully"
    }
    catch {
        Write-Warning "Error during cleanup: $($_.Exception.Message)"
    }
}

# Script execution
try {
    # Set verbose preference if requested
    if ($Verbose) {
        $VerbosePreference = "Continue"
    }
    
    # Execute main function
    $success = Main
    
    # Exit with appropriate code
    if ($success -and $Global:ExitCode -eq 0) {
        Write-Host "`nScript completed successfully" -ForegroundColor Green
        exit 0
    }
    else {
        Write-Host "`nScript completed with errors" -ForegroundColor Red
        exit $Global:ExitCode
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Host "`nFor help, run: Get-Help .\Start-SharePointUpload.ps1 -Full" -ForegroundColor Yellow
    exit 1
}
finally {
    # Final cleanup
    Cleanup-Resources
}
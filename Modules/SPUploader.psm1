# SharePoint Document Uploader - Upload Module
# Handles file uploads to SharePoint Online using PnP PowerShell
# Author: DevOps Team
# Version: 1.0

<#
.SYNOPSIS
    SharePoint Online file upload module using PnP PowerShell

.DESCRIPTION
    Provides file upload functionality with progress tracking, error handling,
    retry mechanisms, and batch processing. Supports large files and folder structures.

.NOTES
    Requires PnP.PowerShell module and active SharePoint connection
#>

class UploadResult {
    [FileInfo]$FileInfo
    [bool]$Success = $false
    [bool]$Skipped = $false
    [string]$ErrorMessage
    [string]$SharePointUrl
    [int]$RetryCount = 0
    [datetime]$UploadTime
    [double]$UploadDurationSeconds
    [long]$BytesTransferred
    
    # Constructor
    UploadResult([FileInfo]$FileInfo) {
        $this.FileInfo = $FileInfo
        $this.UploadTime = Get-Date
    }
}

class SharePointUploader {
    [SharePointConfig]$Config
    [SharePointAuth]$Auth
    [hashtable]$Statistics
    [System.Collections.Generic.List[UploadResult]]$Results
    
    # Constructor
    SharePointUploader([SharePointConfig]$Config, [SharePointAuth]$Auth) {
        $this.Config = $Config
        $this.Auth = $Auth
        $this.ResetStatistics()
        $this.Results = [System.Collections.Generic.List[UploadResult]]::new()
    }
    
    # Reset upload statistics
    [void]ResetStatistics() {
        $this.Statistics = @{
            TotalFiles = 0
            UploadedFiles = 0
            SkippedFiles = 0
            FailedFiles = 0
            TotalSize = 0
            UploadedSize = 0
            StartTime = $null
            EndTime = $null
            Duration = 0
            AverageSpeed = 0
        }
    }
    
    # Upload multiple files
    [UploadResult[]]UploadFiles([FileInfo[]]$Files, [string]$TargetFolder = $null) {
        if (-not $TargetFolder) {
            $TargetFolder = $this.Config.TargetFolderPath
        }
        
        Write-Information "Starting upload of $($Files.Count) files to $TargetFolder" -InformationAction Continue
        
        $this.ResetStatistics()
        $this.Results.Clear()
        $this.Statistics.TotalFiles = $Files.Count
        $this.Statistics.StartTime = Get-Date
        
        # Ensure target folder exists
        if ($this.Config.CreateFolders) {
            try {
                $null = $this.Auth.GetOrCreateFolder($TargetFolder)
            }
            catch {
                Write-Error "Failed to create target folder: $($_.Exception.Message)"
                throw
            }
        }
        
        # Process files in batches
        $batchSize = $this.Config.BatchSize
        $totalBatches = [math]::Ceiling($Files.Count / $batchSize)
        
        for ($i = 0; $i -lt $Files.Count; $i += $batchSize) {
            $batch = $Files[$i..([math]::Min($i + $batchSize - 1, $Files.Count - 1))]
            $batchNumber = [math]::Floor($i / $batchSize) + 1
            
            Write-Verbose "Processing batch $batchNumber of $totalBatches ($($batch.Count) files)"
            
            # Process batch with throttling
            $batch | ForEach-Object -ThrottleLimit $this.Config.ThrottleLimit -Parallel {
                $uploader = $using:this
                $config = $using:Config
                $auth = $using:Auth
                $targetFolder = $using:TargetFolder
                
                $result = $uploader.UploadSingleFile($_, $targetFolder)
                $uploader.Results.Add($result)
                
                # Update statistics (thread-safe)
                if ($result.Success) {
                    [System.Threading.Interlocked]::Increment([ref]$uploader.Statistics.UploadedFiles)
                    [System.Threading.Interlocked]::Add([ref]$uploader.Statistics.UploadedSize, $result.BytesTransferred)
                }
                elseif ($result.Skipped) {
                    [System.Threading.Interlocked]::Increment([ref]$uploader.Statistics.SkippedFiles)
                }
                else {
                    [System.Threading.Interlocked]::Increment([ref]$uploader.Statistics.FailedFiles)
                }
            }
            
            # Progress update
            $completed = [math]::Min($i + $batchSize, $Files.Count)
            if ($this.Config.UseProgressBars) {
                $percentComplete = ($completed / $Files.Count) * 100
                Write-Progress -Activity "Uploading Files" -Status "$completed of $($Files.Count)" -PercentComplete $percentComplete
            }
            
            Write-Information "Batch $batchNumber complete: $completed of $($Files.Count) files processed" -InformationAction Continue
        }
        
        # Clear progress bar
        if ($this.Config.UseProgressBars) {
            Write-Progress -Activity "Uploading Files" -Completed
        }
        
        $this.Statistics.EndTime = Get-Date
        $this.Statistics.Duration = ($this.Statistics.EndTime - $this.Statistics.StartTime).TotalSeconds
        
        if ($this.Statistics.Duration -gt 0) {
            $this.Statistics.AverageSpeed = ($this.Statistics.UploadedSize / 1MB) / $this.Statistics.Duration
        }
        
        Write-Information "Upload complete: $($this.Statistics.UploadedFiles) uploaded, $($this.Statistics.FailedFiles) failed, $($this.Statistics.SkippedFiles) skipped" -InformationAction Continue
        
        return $this.Results.ToArray()
    }
    
    # Upload single file with retry logic
    [UploadResult]UploadSingleFile([FileInfo]$FileInfo, [string]$TargetFolder, [int]$RetryCount = 0) {
        $result = [UploadResult]::new($FileInfo)
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        
        try {
            Write-Verbose "Uploading: $($FileInfo.Name) ($([math]::Round($FileInfo.Size / 1MB, 2)) MB)"
            
            # Check if file exists and handle conflicts
            $conflictResult = $this.HandleFileConflict($FileInfo, $TargetFolder)
            if ($conflictResult) {
                return $conflictResult
            }
            
            # Ensure folder structure exists
            $this.EnsureFolderStructure($FileInfo, $TargetFolder)
            
            # Calculate target path
            $targetPath = $this.GetTargetPath($FileInfo, $TargetFolder)
            
            # Upload file based on size
            if ($FileInfo.Size -gt 100MB) {
                $uploadedFile = $this.UploadLargeFile($FileInfo, $targetPath)
            }
            else {
                $uploadedFile = $this.UploadStandardFile($FileInfo, $targetPath)
            }
            
            # Success
            $result.Success = $true
            $result.RetryCount = $RetryCount
            $result.BytesTransferred = $FileInfo.Size
            $result.SharePointUrl = $this.Auth.SiteUrl.TrimEnd('/') + $uploadedFile.ServerRelativeUrl
            
            Write-Verbose "Successfully uploaded: $($FileInfo.Name)"
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Warning "Upload failed for $($FileInfo.Name): $errorMessage"
            
            # Retry logic
            if ($RetryCount -lt $this.Config.MaxRetries) {
                Write-Information "Retrying upload ($($RetryCount + 1)/$($this.Config.MaxRetries)): $($FileInfo.Name)" -InformationAction Continue
                Start-Sleep -Seconds $this.Config.RetryDelaySeconds
                return $this.UploadSingleFile($FileInfo, $TargetFolder, $RetryCount + 1)
            }
            
            $result.Success = $false
            $result.ErrorMessage = $errorMessage
            $result.RetryCount = $RetryCount
        }
        finally {
            $stopwatch.Stop()
            $result.UploadDurationSeconds = $stopwatch.Elapsed.TotalSeconds
        }
        
        return $result
    }
    
    # Upload standard file (< 100MB)
    [object]UploadStandardFile([FileInfo]$FileInfo, [string]$TargetPath) {
        $this.Auth.EnsureConnection()
        
        $folderPath = Split-Path $TargetPath -Parent
        $fileName = Split-Path $TargetPath -Leaf
        
        # Upload file
        $uploadedFile = Add-PnPFile -Path $FileInfo.FullPath -Folder $folderPath -NewFileName $fileName -ErrorAction Stop
        
        return $uploadedFile
    }
    
    # Upload large file (>= 100MB) with chunked upload
    [object]UploadLargeFile([FileInfo]$FileInfo, [string]$TargetPath) {
        $this.Auth.EnsureConnection()
        
        Write-Verbose "Using chunked upload for large file: $($FileInfo.Name)"
        
        # Use PnP's built-in large file upload
        $uploadedFile = Add-PnPFile -Path $FileInfo.FullPath -Folder (Split-Path $TargetPath -Parent) -NewFileName (Split-Path $TargetPath -Leaf) -ChunkSize 10MB -ErrorAction Stop
        
        return $uploadedFile
    }
    
    # Handle file conflicts
    [UploadResult]HandleFileConflict([FileInfo]$FileInfo, [string]$TargetFolder) {
        if ($this.Config.OverwriteExisting) {
            return $null # Proceed with upload (will overwrite)
        }
        
        try {
            $targetPath = $this.GetTargetPath($FileInfo, $TargetFolder)
            $existingFile = Get-PnPFile -Url $targetPath -AsListItem -ErrorAction SilentlyContinue
            
            if ($existingFile) {
                Write-Verbose "File already exists, skipping: $($FileInfo.Name)"
                $result = [UploadResult]::new($FileInfo)
                $result.Skipped = $true
                $result.ErrorMessage = "File already exists and overwrite is disabled"
                return $result
            }
        }
        catch {
            # File doesn't exist, proceed with upload
        }
        
        return $null
    }
    
    # Ensure folder structure exists
    [void]EnsureFolderStructure([FileInfo]$FileInfo, [string]$TargetFolder) {
        if (-not $this.Config.CreateFolders) {
            return
        }
        
        $relativePath = Split-Path $FileInfo.RelativePath -Parent
        if ($relativePath -and $relativePath -ne ".") {
            $fullFolderPath = "$TargetFolder/$relativePath".Replace("\\", "/").Replace("//", "/")
            $null = $this.Auth.GetOrCreateFolder($fullFolderPath)
        }
    }
    
    # Get target path for file
    [string]GetTargetPath([FileInfo]$FileInfo, [string]$TargetFolder) {
        $targetPath = "$TargetFolder/$($FileInfo.RelativePath)".Replace("\\", "/").Replace("//", "/")
        return $targetPath.TrimStart("/")
    }
    
    # Get upload statistics
    [hashtable]GetStatistics() {
        $stats = $this.Statistics.Clone()
        
        # Calculate additional metrics
        if ($this.Results.Count -gt 0) {
            $successfulUploads = $this.Results | Where-Object { $_.Success }
            $failedUploads = $this.Results | Where-Object { -not $_.Success -and -not $_.Skipped }
            
            if ($successfulUploads) {
                $stats.AverageUploadTime = ($successfulUploads | Measure-Object -Property UploadDurationSeconds -Average).Average
                $stats.FastestUpload = ($successfulUploads | Measure-Object -Property UploadDurationSeconds -Minimum).Minimum
                $stats.SlowestUpload = ($successfulUploads | Measure-Object -Property UploadDurationSeconds -Maximum).Maximum
            }
            
            $stats.SuccessRate = ($this.Statistics.UploadedFiles / $this.Statistics.TotalFiles) * 100
        }
        
        return $stats
    }
    
    # Get failed uploads for retry
    [FileInfo[]]GetFailedFiles() {
        $failedResults = $this.Results | Where-Object { -not $_.Success -and -not $_.Skipped }
        return $failedResults.FileInfo
    }
    
    # Export results to CSV
    [void]ExportResults([string]$OutputPath) {
        $exportData = $this.Results | Select-Object @{
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
        }, @{
            Name = "DurationSeconds"; Expression = { $_.UploadDurationSeconds }
        }
        
        $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Information "Upload results exported to: $OutputPath" -InformationAction Continue
    }
}

# Function to upload files to SharePoint
function Start-SharePointUpload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [FileInfo[]]$Files,
        
        [Parameter(Mandatory = $true)]
        [SharePointConfig]$Config,
        
        [Parameter(Mandatory = $true)]
        [SharePointAuth]$Auth,
        
        [Parameter(Mandatory = $false)]
        [string]$TargetFolder,
        
        [Parameter(Mandatory = $false)]
        [string]$ExportResultsPath
    )
    
    try {
        $uploader = [SharePointUploader]::new($Config, $Auth)
        $results = $uploader.UploadFiles($Files, $TargetFolder)
        
        # Export results if requested
        if ($ExportResultsPath) {
            $uploader.ExportResults($ExportResultsPath)
        }
        
        return @{
            Results = $results
            Statistics = $uploader.GetStatistics()
            FailedFiles = $uploader.GetFailedFiles()
        }
    }
    catch {
        Write-Error "Upload operation failed: $($_.Exception.Message)"
        throw
    }
}

# Function to retry failed uploads
function Resume-SharePointUpload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [FileInfo[]]$FailedFiles,
        
        [Parameter(Mandatory = $true)]
        [SharePointConfig]$Config,
        
        [Parameter(Mandatory = $true)]
        [SharePointAuth]$Auth,
        
        [Parameter(Mandatory = $false)]
        [string]$TargetFolder
    )
    
    if ($FailedFiles.Count -eq 0) {
        Write-Information "No failed files to retry" -InformationAction Continue
        return $null
    }
    
    Write-Information "Retrying upload for $($FailedFiles.Count) failed files" -InformationAction Continue
    
    return Start-SharePointUpload -Files $FailedFiles -Config $Config -Auth $Auth -TargetFolder $TargetFolder
}

# Function to test upload performance
function Test-SharePointUploadPerformance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [SharePointAuth]$Auth,
        
        [Parameter(Mandatory = $false)]
        [string]$TargetFolder = "/",
        
        [Parameter(Mandatory = $false)]
        [int]$TestFileSizeMB = 10,
        
        [Parameter(Mandatory = $false)]
        [int]$NumberOfFiles = 5
    )
    
    Write-Information "Testing SharePoint upload performance..." -InformationAction Continue
    
    $testFiles = @()
    $tempDir = Join-Path $env:TEMP "SPUploadTest_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
    
    try {
        # Create test files
        for ($i = 1; $i -le $NumberOfFiles; $i++) {
            $testFileName = "test_file_$i.dat"
            $testFilePath = Join-Path $tempDir $testFileName
            
            # Create file with random data
            $bytes = [byte[]]::new($TestFileSizeMB * 1MB)
            (New-Object System.Random).NextBytes($bytes)
            [System.IO.File]::WriteAllBytes($testFilePath, $bytes)
            
            $fileObj = Get-Item $testFilePath
            $testFiles += [FileInfo]::new($fileObj, $tempDir)
        }
        
        # Test upload
        $config = [SharePointConfig]::new()
        $config.MaxFileSizeMB = $TestFileSizeMB * 2
        $config.BatchSize = $NumberOfFiles
        $config.UseProgressBars = $true
        
        $uploader = [SharePointUploader]::new($config, $Auth)
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        
        $results = $uploader.UploadFiles($testFiles, $TargetFolder)
        
        $stopwatch.Stop()
        
        # Calculate performance metrics
        $successfulUploads = $results | Where-Object { $_.Success }
        $totalSize = ($successfulUploads | Measure-Object -Property BytesTransferred -Sum).Sum
        $avgSpeed = if ($stopwatch.Elapsed.TotalSeconds -gt 0) { ($totalSize / 1MB) / $stopwatch.Elapsed.TotalSeconds } else { 0 }
        
        # Clean up test files from SharePoint
        foreach ($result in $successfulUploads) {
            try {
                Remove-PnPFile -ServerRelativeUrl $result.SharePointUrl.Replace($Auth.SiteUrl, "") -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Warning "Could not clean up test file: $($result.FileInfo.Name)"
            }
        }
        
        return @{
            TestDuration = $stopwatch.Elapsed.TotalSeconds
            FilesUploaded = $successfulUploads.Count
            TotalSizeMB = [math]::Round($totalSize / 1MB, 2)
            AverageSpeedMBps = [math]::Round($avgSpeed, 2)
            SuccessRate = ($successfulUploads.Count / $testFiles.Count) * 100
        }
    }
    finally {
        # Clean up temp directory
        if (Test-Path $tempDir) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# Export module members
Export-ModuleMember -Function Start-SharePointUpload, Resume-SharePointUpload, Test-SharePointUploadPerformance
Export-ModuleMember -Cmdlet *
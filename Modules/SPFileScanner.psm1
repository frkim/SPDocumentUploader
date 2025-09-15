# SharePoint Document Uploader - File Scanner Module
# Handles local file discovery and validation
# Author: DevOps Team
# Version: 1.0

<#
.SYNOPSIS
    File scanning module for SharePoint Document Uploader

.DESCRIPTION
    Provides file discovery, filtering, and validation capabilities for local drives.
    Supports recursive scanning, extension filtering, size validation, and metadata extraction.

.NOTES
    Requires PowerShell 5.1 or later
    Optimized for network drives and large file sets
#>

class FileInfo {
    [string]$FullPath
    [string]$RelativePath
    [string]$Name
    [long]$Size
    [datetime]$LastWriteTime
    [string]$Extension
    [bool]$IsValid
    [string]$ErrorMessage
    [string]$Hash
    
    # Constructor
    FileInfo([System.IO.FileInfo]$FileObj, [string]$BasePath) {
        $this.FullPath = $FileObj.FullName
        $this.Name = $FileObj.Name
        $this.Size = $FileObj.Length
        $this.LastWriteTime = $FileObj.LastWriteTime
        $this.Extension = $FileObj.Extension.ToLower()
        $this.IsValid = $true
        
        # Calculate relative path
        try {
            $this.RelativePath = $FileObj.FullName.Substring($BasePath.Length).TrimStart('\', '/')
            $this.RelativePath = $this.RelativePath.Replace('\', '/')
        }
        catch {
            $this.RelativePath = $this.Name
        }
    }
}

class FileScanner {
    [SharePointConfig]$Config
    [hashtable]$Statistics
    [string[]]$SkipPatterns = @("thumbs.db", "desktop.ini", ".ds_store", "~$*", "*.tmp", "*.temp")
    
    # Constructor
    FileScanner([SharePointConfig]$Config) {
        $this.Config = $Config
        $this.ResetStatistics()
    }
    
    # Reset statistics
    [void]ResetStatistics() {
        $this.Statistics = @{
            TotalFiles = 0
            ValidFiles = 0
            SkippedFiles = 0
            ErrorFiles = 0
            TotalSize = 0
            ValidSize = 0
            ScanDuration = 0
            DirectoriesScanned = 0
        }
    }
    
    # Scan directory for files
    [FileInfo[]]ScanDirectory([string]$SourcePath = $null) {
        if (-not $SourcePath) {
            $SourcePath = $this.Config.LocalSourcePath
        }
        
        if (-not (Test-Path -Path $SourcePath -PathType Container)) {
            throw "Source path does not exist or is not a directory: $SourcePath"
        }
        
        Write-Information "Scanning directory: $SourcePath" -InformationAction Continue
        Write-Verbose "Include subfolders: $($this.Config.IncludeSubfolders)"
        Write-Verbose "File extensions filter: $($this.Config.FileExtensions -join ', ')"
        
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $this.ResetStatistics()
        
        $files = @()
        
        try {
            # Get files based on subfolder setting
            if ($this.Config.IncludeSubfolders) {
                $fileObjects = Get-ChildItem -Path $SourcePath -File -Recurse -Force -ErrorAction SilentlyContinue
            }
            else {
                $fileObjects = Get-ChildItem -Path $SourcePath -File -Force -ErrorAction SilentlyContinue
            }
            
            $this.Statistics.TotalFiles = $fileObjects.Count
            Write-Verbose "Found $($fileObjects.Count) files"
            
            # Process files with progress
            $progress = 0
            foreach ($fileObj in $fileObjects) {
                $progress++
                
                if ($this.Config.UseProgressBars) {
                    $percentComplete = ($progress / $fileObjects.Count) * 100
                    Write-Progress -Activity "Scanning Files" -Status "$progress of $($fileObjects.Count)" -PercentComplete $percentComplete
                }
                
                try {
                    $fileInfo = [FileInfo]::new($fileObj, $SourcePath)
                    
                    # Validate file
                    $this.ValidateFile($fileInfo)
                    
                    if ($fileInfo.IsValid) {
                        $files += $fileInfo
                        $this.Statistics.ValidFiles++
                        $this.Statistics.ValidSize += $fileInfo.Size
                    }
                    else {
                        $this.Statistics.SkippedFiles++
                        Write-Verbose "Skipped: $($fileInfo.Name) - $($fileInfo.ErrorMessage)"
                    }
                }
                catch {
                    $this.Statistics.ErrorFiles++
                    Write-Warning "Error processing file $($fileObj.FullName): $($_.Exception.Message)"
                }
            }
            
            # Clear progress bar
            if ($this.Config.UseProgressBars) {
                Write-Progress -Activity "Scanning Files" -Completed
            }
        }
        catch {
            Write-Error "Error scanning directory: $($_.Exception.Message)"
            throw
        }
        finally {
            $stopwatch.Stop()
            $this.Statistics.ScanDuration = $stopwatch.Elapsed.TotalSeconds
        }
        
        Write-Information "Scan complete: $($this.Statistics.ValidFiles) valid files, $($this.Statistics.SkippedFiles) skipped" -InformationAction Continue
        Write-Information "Total size: $([math]::Round($this.Statistics.ValidSize / 1MB, 2)) MB" -InformationAction Continue
        
        return $files
    }
    
    # Validate individual file
    [void]ValidateFile([FileInfo]$FileInfo) {
        # Check file extension filter
        if ($this.Config.FileExtensions -and $this.Config.FileExtensions.Count -gt 0) {
            if ($FileInfo.Extension -notin $this.Config.FileExtensions) {
                $FileInfo.IsValid = $false
                $FileInfo.ErrorMessage = "Extension '$($FileInfo.Extension)' not in allowed list"
                return
            }
        }
        
        # Check file size
        $maxSizeBytes = $this.Config.MaxFileSizeMB * 1MB
        if ($FileInfo.Size -gt $maxSizeBytes) {
            $FileInfo.IsValid = $false
            $FileInfo.ErrorMessage = "File size ($([math]::Round($FileInfo.Size / 1MB, 2)) MB) exceeds maximum ($($this.Config.MaxFileSizeMB) MB)"
            return
        }
        
        # Check skip patterns
        foreach ($pattern in $this.SkipPatterns) {
            if ($FileInfo.Name -like $pattern) {
                $FileInfo.IsValid = $false
                $FileInfo.ErrorMessage = "File matches skip pattern: $pattern"
                return
            }
        }
        
        # Check file accessibility
        try {
            $stream = [System.IO.File]::OpenRead($FileInfo.FullPath)
            $stream.Close()
        }
        catch [System.UnauthorizedAccessException] {
            $FileInfo.IsValid = $false
            $FileInfo.ErrorMessage = "Access denied"
            return
        }
        catch [System.IO.IOException] {
            $FileInfo.IsValid = $false
            $FileInfo.ErrorMessage = "File is locked or in use"
            return
        }
        catch {
            $FileInfo.IsValid = $false
            $FileInfo.ErrorMessage = "File access error: $($_.Exception.Message)"
            return
        }
        
        # File is valid
        $FileInfo.IsValid = $true
    }
    
    # Get files grouped by extension
    [hashtable]GetFilesByExtension([FileInfo[]]$Files) {
        $grouped = @{}
        
        foreach ($file in $Files | Where-Object { $_.IsValid }) {
            $ext = if ($file.Extension) { $file.Extension } else { "no_extension" }
            
            if (-not $grouped.ContainsKey($ext)) {
                $grouped[$ext] = @()
            }
            
            $grouped[$ext] += $file
        }
        
        return $grouped
    }
    
    # Get large files above threshold
    [FileInfo[]]GetLargeFiles([FileInfo[]]$Files, [int]$ThresholdMB = 50) {
        $thresholdBytes = $ThresholdMB * 1MB
        return $Files | Where-Object { $_.IsValid -and $_.Size -gt $thresholdBytes }
    }
    
    # Filter files by various criteria
    [FileInfo[]]FilterFiles([FileInfo[]]$Files, [hashtable]$Filters) {
        $filtered = $Files | Where-Object { $_.IsValid }
        
        if ($Filters.ContainsKey("Extensions")) {
            $extensions = $Filters.Extensions
            $filtered = $filtered | Where-Object { $_.Extension -in $extensions }
        }
        
        if ($Filters.ContainsKey("MinSize")) {
            $minSize = $Filters.MinSize
            $filtered = $filtered | Where-Object { $_.Size -ge $minSize }
        }
        
        if ($Filters.ContainsKey("MaxSize")) {
            $maxSize = $Filters.MaxSize
            $filtered = $filtered | Where-Object { $_.Size -le $maxSize }
        }
        
        if ($Filters.ContainsKey("ModifiedAfter")) {
            $after = $Filters.ModifiedAfter
            $filtered = $filtered | Where-Object { $_.LastWriteTime -gt $after }
        }
        
        if ($Filters.ContainsKey("ModifiedBefore")) {
            $before = $Filters.ModifiedBefore
            $filtered = $filtered | Where-Object { $_.LastWriteTime -lt $before }
        }
        
        if ($Filters.ContainsKey("NamePattern")) {
            $pattern = $Filters.NamePattern
            $filtered = $filtered | Where-Object { $_.Name -like $pattern }
        }
        
        return $filtered
    }
    
    # Calculate file hashes for duplicate detection
    [void]CalculateFileHashes([FileInfo[]]$Files, [string]$Algorithm = "SHA256") {
        Write-Information "Calculating file hashes..." -InformationAction Continue
        
        $progress = 0
        foreach ($file in $Files | Where-Object { $_.IsValid }) {
            $progress++
            
            if ($this.Config.UseProgressBars) {
                $percentComplete = ($progress / $Files.Count) * 100
                Write-Progress -Activity "Calculating Hashes" -Status "$progress of $($Files.Count)" -PercentComplete $percentComplete
            }
            
            try {
                $hash = Get-FileHash -Path $file.FullPath -Algorithm $Algorithm -ErrorAction Stop
                $file.Hash = $hash.Hash
            }
            catch {
                Write-Warning "Failed to calculate hash for $($file.Name): $($_.Exception.Message)"
            }
        }
        
        if ($this.Config.UseProgressBars) {
            Write-Progress -Activity "Calculating Hashes" -Completed
        }
    }
    
    # Find duplicate files by hash
    [hashtable]FindDuplicateFiles([FileInfo[]]$Files) {
        $duplicates = @{}
        $hashGroups = $Files | Where-Object { $_.IsValid -and $_.Hash } | Group-Object -Property Hash
        
        foreach ($group in $hashGroups | Where-Object { $_.Count -gt 1 }) {
            $duplicates[$group.Name] = $group.Group
        }
        
        return $duplicates
    }
    
    # Export scan results to CSV
    [void]ExportToCsv([FileInfo[]]$Files, [string]$OutputPath) {
        $exportData = $Files | Select-Object FullPath, RelativePath, Name, Size, LastWriteTime, Extension, IsValid, ErrorMessage, Hash
        $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Information "Scan results exported to: $OutputPath" -InformationAction Continue
    }
    
    # Get detailed statistics
    [hashtable]GetDetailedStatistics([FileInfo[]]$Files) {
        $stats = $this.Statistics.Clone()
        
        # Add file type breakdown
        $fileTypes = $this.GetFilesByExtension($Files)
        $stats.FileTypeBreakdown = @{}
        
        foreach ($ext in $fileTypes.Keys) {
            $extFiles = $fileTypes[$ext]
            $stats.FileTypeBreakdown[$ext] = @{
                Count = $extFiles.Count
                TotalSize = ($extFiles | Measure-Object -Property Size -Sum).Sum
            }
        }
        
        # Add size distribution
        $stats.SizeDistribution = @{
            Small = ($Files | Where-Object { $_.IsValid -and $_.Size -lt 1MB }).Count
            Medium = ($Files | Where-Object { $_.IsValid -and $_.Size -ge 1MB -and $_.Size -lt 10MB }).Count
            Large = ($Files | Where-Object { $_.IsValid -and $_.Size -ge 10MB -and $_.Size -lt 100MB }).Count
            XLarge = ($Files | Where-Object { $_.IsValid -and $_.Size -ge 100MB }).Count
        }
        
        return $stats
    }
}

# Function to scan files with configuration
function Get-FilesToUpload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [SharePointConfig]$Config,
        
        [Parameter(Mandatory = $false)]
        [string]$SourcePath,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Filters,
        
        [Parameter(Mandatory = $false)]
        [switch]$CalculateHashes,
        
        [Parameter(Mandatory = $false)]
        [string]$ExportPath
    )
    
    try {
        $scanner = [FileScanner]::new($Config)
        $files = $scanner.ScanDirectory($SourcePath)
        
        # Apply additional filters if provided
        if ($Filters) {
            $files = $scanner.FilterFiles($files, $Filters)
        }
        
        # Calculate hashes if requested
        if ($CalculateHashes) {
            $scanner.CalculateFileHashes($files)
        }
        
        # Export if requested
        if ($ExportPath) {
            $scanner.ExportToCsv($files, $ExportPath)
        }
        
        # Return results with statistics
        return @{
            Files = $files
            Statistics = $scanner.GetDetailedStatistics($files)
        }
    }
    catch {
        Write-Error "File scan failed: $($_.Exception.Message)"
        throw
    }
}

# Function to analyze file distribution
function Get-FileAnalysis {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [FileInfo[]]$Files,
        
        [Parameter(Mandatory = $false)]
        [switch]$FindDuplicates
    )
    
    $analysis = @{
        TotalFiles = $Files.Count
        ValidFiles = ($Files | Where-Object { $_.IsValid }).Count
        TotalSize = ($Files | Where-Object { $_.IsValid } | Measure-Object -Property Size -Sum).Sum
        ExtensionBreakdown = @{}
        SizeDistribution = @{}
        DateDistribution = @{}
    }
    
    # Extension breakdown
    $extGroups = $Files | Where-Object { $_.IsValid } | Group-Object -Property Extension
    foreach ($group in $extGroups) {
        $analysis.ExtensionBreakdown[$group.Name] = @{
            Count = $group.Count
            TotalSize = ($group.Group | Measure-Object -Property Size -Sum).Sum
        }
    }
    
    # Size distribution
    $analysis.SizeDistribution = @{
        "Under 1MB" = ($Files | Where-Object { $_.IsValid -and $_.Size -lt 1MB }).Count
        "1-10MB" = ($Files | Where-Object { $_.IsValid -and $_.Size -ge 1MB -and $_.Size -lt 10MB }).Count
        "10-100MB" = ($Files | Where-Object { $_.IsValid -and $_.Size -ge 10MB -and $_.Size -lt 100MB }).Count
        "Over 100MB" = ($Files | Where-Object { $_.IsValid -and $_.Size -ge 100MB }).Count
    }
    
    # Date distribution (last 30 days)
    $cutoffDate = (Get-Date).AddDays(-30)
    $analysis.DateDistribution = @{
        "Last 7 days" = ($Files | Where-Object { $_.IsValid -and $_.LastWriteTime -gt (Get-Date).AddDays(-7) }).Count
        "Last 30 days" = ($Files | Where-Object { $_.IsValid -and $_.LastWriteTime -gt $cutoffDate }).Count
        "Older than 30 days" = ($Files | Where-Object { $_.IsValid -and $_.LastWriteTime -le $cutoffDate }).Count
    }
    
    # Find duplicates if requested
    if ($FindDuplicates -and ($Files | Where-Object { $_.Hash })) {
        $scanner = [FileScanner]::new($null)
        $analysis.Duplicates = $scanner.FindDuplicateFiles($Files)
    }
    
    return $analysis
}

# Export module members
Export-ModuleMember -Function Get-FilesToUpload, Get-FileAnalysis
Export-ModuleMember -Cmdlet *
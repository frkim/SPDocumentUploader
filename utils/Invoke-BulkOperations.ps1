<#
.SYNOPSIS
    Bulk Operations Utility for SharePoint Document Uploader

.DESCRIPTION
    Provides bulk operations for SharePoint document management including
    batch uploads from multiple sources, cleanup operations, and maintenance tasks.

.PARAMETER Operation
    Bulk operation to perform: MultiSource, Cleanup, Sync, Report, Archive

.PARAMETER ConfigPath
    Path to JSON configuration file

.PARAMETER SourceListFile
    Text file containing list of source directories (one per line)

.PARAMETER Days
    Number of days for cleanup/archive operations

.PARAMETER DryRun
    Preview operations without making changes

.PARAMETER Force
    Skip confirmation prompts

.EXAMPLE
    .\Invoke-BulkOperations.ps1 -Operation MultiSource -SourceListFile "sources.txt"
    # Upload from multiple source directories

.EXAMPLE
    .\Invoke-BulkOperations.ps1 -Operation Cleanup -Days 30 -DryRun
    # Preview cleanup of files older than 30 days

.EXAMPLE
    .\Invoke-BulkOperations.ps1 -Operation Report -ConfigPath "prod-config.json"
    # Generate upload activity report

.NOTES
    Author: DevOps Team
    Requires: PnP.PowerShell module, SharePoint modules
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("MultiSource", "Cleanup", "Sync", "Report", "Archive", "Verify")]
    [string]$Operation,
    
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "config.json",
    
    [Parameter(Mandatory = $false)]
    [string]$SourceListFile,
    
    [Parameter(Mandatory = $false)]
    [int]$Days = 30,
    
    [Parameter(Mandatory = $false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

$ErrorActionPreference = "Stop"

# Import required modules
$ModulePath = Join-Path (Split-Path $PSScriptRoot -Parent) "Modules"
Import-Module (Join-Path $ModulePath "SPConfig.psm1") -Force
Import-Module (Join-Path $ModulePath "SPAuth.psm1") -Force
Import-Module (Join-Path $ModulePath "SPFileScanner.psm1") -Force
Import-Module (Join-Path $ModulePath "SPUploader.psm1") -Force
Import-Module (Join-Path $ModulePath "SPLogger.psm1") -Force

# Global variables
$Global:SPAuth = $null
$Global:SPLogger = $null
$Global:Config = $null

function Invoke-BulkOperation {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "SharePoint Bulk Operations Utility" -ForegroundColor Cyan
        Write-Host "Operation: $Operation" -ForegroundColor Yellow
        Write-Host "=" * 40 -ForegroundColor Cyan
        
        # Load configuration
        $Global:Config = Get-SharePointConfig -ConfigPath $ConfigPath -ApplyEnvOverrides
        $Global:SPLogger = Initialize-SPLogger -Config $Global:Config
        
        Write-SPInformation "Starting bulk operation: $Operation"
        
        # Execute operation
        switch ($Operation) {
            "MultiSource" { Invoke-MultiSourceUpload }
            "Cleanup" { Invoke-CleanupOperation }
            "Sync" { Invoke-SyncOperation }
            "Report" { Invoke-ReportGeneration }
            "Archive" { Invoke-ArchiveOperation }
            "Verify" { Invoke-VerifyOperation }
        }
        
        Write-SPInformation "Bulk operation completed successfully"
        return $true
    }
    catch {
        Write-SPError "Bulk operation failed" -Exception $_.Exception
        return $false
    }
    finally {
        Cleanup-BulkResources
    }
}

function Invoke-MultiSourceUpload {
    Write-SPInformation "Starting multi-source upload operation"
    
    if (-not $SourceListFile -or -not (Test-Path $SourceListFile)) {
        throw "Source list file required for MultiSource operation: $SourceListFile"
    }
    
    # Read source directories
    $sourceDirs = Get-Content $SourceListFile | Where-Object { $_ -and $_.Trim() -and -not $_.StartsWith("#") }
    
    if (-not $sourceDirs) {
        throw "No valid source directories found in: $SourceListFile"
    }
    
    Write-SPInformation "Found $($sourceDirs.Count) source directories to process"
    
    # Connect to SharePoint
    $Global:SPAuth = Connect-SharePointSite -Config $Global:Config -TestConnection
    
    $totalResults = @()\n    $totalStats = @{\n        ProcessedDirs = 0\n        TotalFiles = 0\n        SuccessfulUploads = 0\n        FailedUploads = 0\n        SkippedFiles = 0\n        TotalSizeMB = 0\n        StartTime = Get-Date\n    }\n    \n    foreach ($sourceDir in $sourceDirs) {\n        $sourceDir = $sourceDir.Trim()\n        \n        try {\n            Write-SPInformation "Processing source directory: $sourceDir"\n            \n            # Validate source directory\n            if (-not (Test-Path $sourceDir)) {\n                Write-SPWarning "Source directory not found, skipping: $sourceDir"\n                continue\n            }\n            \n            # Scan files in source directory\n            $scanner = New-FileScanner -Config $Global:Config\n            $files = Get-DirectoryFiles -Scanner $scanner -Path $sourceDir\n            $validFiles = $files | Where-Object { $_.IsValid }\n            \n            if (-not $validFiles) {\n                Write-SPWarning "No valid files found in: $sourceDir"\n                continue\n            }\n            \n            Write-SPInformation "Found $($validFiles.Count) files to upload from: $sourceDir"\n            \n            # Create target folder based on source directory name\n            $sourceDirName = Split-Path $sourceDir -Leaf\n            $targetFolder = Join-Path $Global:Config.TargetFolderPath $sourceDirName -Replace \"\\\\\", \"/\"\n            \n            if ($DryRun) {\n                Write-SPInformation \"DRY RUN: Would upload $($validFiles.Count) files to $targetFolder\"\n                $totalStats.TotalFiles += $validFiles.Count\n                $totalStats.TotalSizeMB += ($validFiles | Measure-Object -Property Size -Sum).Sum / 1MB\n            }\n            else {\n                # Perform upload\n                $uploadResult = Start-SharePointUpload -Files $validFiles -Config $Global:Config -Auth $Global:SPAuth -TargetFolder $targetFolder\n                \n                $totalResults += $uploadResult.Results\n                $totalStats.SuccessfulUploads += $uploadResult.Statistics.SuccessfulFiles\n                $totalStats.FailedUploads += $uploadResult.Statistics.FailedFiles\n                $totalStats.SkippedFiles += $uploadResult.Statistics.SkippedFiles\n                $totalStats.TotalSizeMB += $uploadResult.Statistics.TotalSizeMB\n            }\n            \n            $totalStats.ProcessedDirs++\n        }\n        catch {\n            Write-SPError \"Failed to process source directory: $sourceDir\" -Exception $_.Exception\n            continue\n        }\n    }\n    \n    # Show summary\n    $totalStats.EndTime = Get-Date\n    $duration = ($totalStats.EndTime - $totalStats.StartTime).TotalMinutes\n    \n    Write-Host \"`n=== MULTI-SOURCE UPLOAD SUMMARY ===\" -ForegroundColor Green\n    Write-Host \"Processed directories: $($totalStats.ProcessedDirs)\" -ForegroundColor White\n    Write-Host \"Total files: $($totalStats.TotalFiles)\" -ForegroundColor White\n    \n    if (-not $DryRun) {\n        Write-Host \"Successful uploads: $($totalStats.SuccessfulUploads)\" -ForegroundColor Green\n        Write-Host \"Failed uploads: $($totalStats.FailedUploads)\" -ForegroundColor Red\n        Write-Host \"Skipped files: $($totalStats.SkippedFiles)\" -ForegroundColor Yellow\n    }\n    \n    Write-Host \"Total size: $([math]::Round($totalStats.TotalSizeMB, 2)) MB\" -ForegroundColor White\n    Write-Host \"Duration: $([math]::Round($duration, 2)) minutes\" -ForegroundColor White\n    \n    # Export detailed results if not dry run\n    if (-not $DryRun -and $totalResults) {\n        $resultsFile = \"MultiSourceUpload_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv\"\n        $totalResults | Export-Csv -Path $resultsFile -NoTypeInformation\n        Write-SPInformation \"Detailed results exported to: $resultsFile\"\n    }\n}\n\nfunction Invoke-CleanupOperation {\n    Write-SPInformation \"Starting cleanup operation (files older than $Days days)\"\n    \n    if (-not $Force -and -not $DryRun) {\n        $confirm = Read-Host \"This will delete files from SharePoint. Continue? (y/N)\"\n        if ($confirm -notmatch \"^[Yy]\") {\n            Write-SPInformation \"Cleanup operation cancelled\"\n            return\n        }\n    }\n    \n    # Connect to SharePoint\n    $Global:SPAuth = Connect-SharePointSite -Config $Global:Config -TestConnection\n    \n    # Calculate cutoff date\n    $cutoffDate = (Get-Date).AddDays(-$Days)\n    Write-SPInformation \"Cleaning up files modified before: $($cutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))\"\n    \n    try {\n        # Get all files in target folder\n        $targetFolder = $Global:Config.TargetFolderPath\n        if (-not $targetFolder) {\n            $targetFolder = \"/\"\n        }\n        \n        $listItems = Get-PnPListItem -List $Global:Config.DocumentLibraryName -FolderServerRelativeUrl $targetFolder -PageSize 1000\n        \n        $oldFiles = $listItems | Where-Object { \n            $_.FileSystemObjectType -eq \"File\" -and \n            $_.FieldValues.Modified -lt $cutoffDate \n        }\n        \n        Write-SPInformation \"Found $($oldFiles.Count) files eligible for cleanup\"\n        \n        if ($oldFiles.Count -eq 0) {\n            Write-SPInformation \"No files found for cleanup\"\n            return\n        }\n        \n        $deletedCount = 0\n        $errorCount = 0\n        $totalSize = 0\n        \n        foreach ($file in $oldFiles) {\n            try {\n                $fileName = $file.FieldValues.FileLeafRef\n                $fileSize = $file.FieldValues.File_x0020_Size\n                $modifiedDate = $file.FieldValues.Modified\n                \n                Write-SPInformation \"Processing: $fileName (Modified: $modifiedDate, Size: $([math]::Round($fileSize / 1KB, 2)) KB)\"\n                \n                if ($DryRun) {\n                    Write-SPInformation \"DRY RUN: Would delete $fileName\"\n                }\n                else {\n                    Remove-PnPFile -Identity $file.FieldValues.FileRef -Force -ErrorAction Stop\n                    Write-SPInformation \"Deleted: $fileName\"\n                }\n                \n                $deletedCount++\n                $totalSize += $fileSize\n            }\n            catch {\n                Write-SPError \"Failed to delete file: $fileName\" -Exception $_.Exception\n                $errorCount++\n            }\n        }\n        \n        # Show summary\n        Write-Host \"`n=== CLEANUP SUMMARY ===\" -ForegroundColor Green\n        \n        if ($DryRun) {\n            Write-Host \"Files that would be deleted: $deletedCount\" -ForegroundColor Yellow\n        }\n        else {\n            Write-Host \"Files deleted: $deletedCount\" -ForegroundColor Green\n        }\n        \n        Write-Host \"Errors encountered: $errorCount\" -ForegroundColor $(if ($errorCount -gt 0) { \"Red\" } else { \"Green\" })\n        Write-Host \"Total size freed: $([math]::Round($totalSize / 1MB, 2)) MB\" -ForegroundColor White\n    }\n    catch {\n        Write-SPError \"Cleanup operation failed\" -Exception $_.Exception\n        throw\n    }\n}\n\nfunction Invoke-SyncOperation {\n    Write-SPInformation \"Starting synchronization operation\"\n    \n    # Connect to SharePoint\n    $Global:SPAuth = Connect-SharePointSite -Config $Global:Config -TestConnection\n    \n    # Get local files\n    $scanner = New-FileScanner -Config $Global:Config\n    $localFiles = Get-DirectoryFiles -Scanner $scanner -Path $Global:Config.LocalSourcePath\n    $validLocalFiles = $localFiles | Where-Object { $_.IsValid }\n    \n    # Get SharePoint files\n    $spFiles = @()\n    try {\n        $listItems = Get-PnPListItem -List $Global:Config.DocumentLibraryName -FolderServerRelativeUrl $Global:Config.TargetFolderPath -PageSize 1000\n        $spFiles = $listItems | Where-Object { $_.FileSystemObjectType -eq \"File\" }\n    }\n    catch {\n        Write-SPWarning \"Could not retrieve SharePoint files: $($_.Exception.Message)\"\n    }\n    \n    # Compare files\n    $localFileMap = @{}\n    foreach ($file in $validLocalFiles) {\n        $localFileMap[$file.Name.ToLower()] = $file\n    }\n    \n    $spFileMap = @{}\n    foreach ($file in $spFiles) {\n        $spFileMap[$file.FieldValues.FileLeafRef.ToLower()] = $file\n    }\n    \n    # Find differences\n    $newFiles = @()\n    $updatedFiles = @()\n    $deletedFiles = @()\n    \n    # Files to upload (new or updated)\n    foreach ($localFile in $validLocalFiles) {\n        $fileName = $localFile.Name.ToLower()\n        \n        if (-not $spFileMap.ContainsKey($fileName)) {\n            $newFiles += $localFile\n        }\n        else {\n            $spFile = $spFileMap[$fileName]\n            $spModified = $spFile.FieldValues.Modified\n            $localModified = $localFile.LastWriteTime\n            \n            if ($localModified -gt $spModified) {\n                $updatedFiles += $localFile\n            }\n        }\n    }\n    \n    # Files to delete (exist in SharePoint but not locally)\n    foreach ($spFileName in $spFileMap.Keys) {\n        if (-not $localFileMap.ContainsKey($spFileName)) {\n            $deletedFiles += $spFileMap[$spFileName]\n        }\n    }\n    \n    # Show sync analysis\n    Write-Host \"`n=== SYNCHRONIZATION ANALYSIS ===\" -ForegroundColor Cyan\n    Write-Host \"Local files: $($validLocalFiles.Count)\" -ForegroundColor White\n    Write-Host \"SharePoint files: $($spFiles.Count)\" -ForegroundColor White\n    Write-Host \"New files to upload: $($newFiles.Count)\" -ForegroundColor Green\n    Write-Host \"Updated files to upload: $($updatedFiles.Count)\" -ForegroundColor Yellow\n    Write-Host \"Files to delete: $($deletedFiles.Count)\" -ForegroundColor Red\n    \n    if ($newFiles.Count -eq 0 -and $updatedFiles.Count -eq 0 -and $deletedFiles.Count -eq 0) {\n        Write-SPInformation \"SharePoint is already in sync with local directory\"\n        return\n    }\n    \n    if (-not $Force -and -not $DryRun) {\n        $confirm = Read-Host \"Proceed with synchronization? (y/N)\"\n        if ($confirm -notmatch \"^[Yy]\") {\n            Write-SPInformation \"Synchronization cancelled\"\n            return\n        }\n    }\n    \n    $syncStats = @{\n        UploadedFiles = 0\n        DeletedFiles = 0\n        Errors = 0\n    }\n    \n    # Upload new and updated files\n    $filesToUpload = $newFiles + $updatedFiles\n    if ($filesToUpload) {\n        Write-SPInformation \"Uploading $($filesToUpload.Count) files...\"\n        \n        if ($DryRun) {\n            Write-SPInformation \"DRY RUN: Would upload $($filesToUpload.Count) files\"\n            $syncStats.UploadedFiles = $filesToUpload.Count\n        }\n        else {\n            $uploadResult = Start-SharePointUpload -Files $filesToUpload -Config $Global:Config -Auth $Global:SPAuth\n            $syncStats.UploadedFiles = $uploadResult.Statistics.SuccessfulFiles\n            $syncStats.Errors += $uploadResult.Statistics.FailedFiles\n        }\n    }\n    \n    # Delete removed files\n    if ($deletedFiles) {\n        Write-SPInformation \"Removing $($deletedFiles.Count) deleted files...\"\n        \n        foreach ($file in $deletedFiles) {\n            try {\n                $fileName = $file.FieldValues.FileLeafRef\n                \n                if ($DryRun) {\n                    Write-SPInformation \"DRY RUN: Would delete $fileName\"\n                }\n                else {\n                    Remove-PnPFile -Identity $file.FieldValues.FileRef -Force\n                    Write-SPInformation \"Deleted: $fileName\"\n                }\n                \n                $syncStats.DeletedFiles++\n            }\n            catch {\n                Write-SPError \"Failed to delete: $($file.FieldValues.FileLeafRef)\" -Exception $_.Exception\n                $syncStats.Errors++\n            }\n        }\n    }\n    \n    # Show sync summary\n    Write-Host \"`n=== SYNC SUMMARY ===\" -ForegroundColor Green\n    \n    if ($DryRun) {\n        Write-Host \"Files that would be uploaded: $($syncStats.UploadedFiles)\" -ForegroundColor Yellow\n        Write-Host \"Files that would be deleted: $($syncStats.DeletedFiles)\" -ForegroundColor Yellow\n    }\n    else {\n        Write-Host \"Files uploaded: $($syncStats.UploadedFiles)\" -ForegroundColor Green\n        Write-Host \"Files deleted: $($syncStats.DeletedFiles)\" -ForegroundColor Green\n    }\n    \n    Write-Host \"Errors: $($syncStats.Errors)\" -ForegroundColor $(if ($syncStats.Errors -gt 0) { \"Red\" } else { \"Green\" })\n}\n\nfunction Invoke-ReportGeneration {\n    Write-SPInformation \"Generating SharePoint upload activity report\"\n    \n    # Connect to SharePoint\n    $Global:SPAuth = Connect-SharePointSite -Config $Global:Config -TestConnection\n    \n    try {\n        # Get all files from document library\n        $listItems = Get-PnPListItem -List $Global:Config.DocumentLibraryName -FolderServerRelativeUrl $Global:Config.TargetFolderPath -PageSize 2000\n        $files = $listItems | Where-Object { $_.FileSystemObjectType -eq \"File\" }\n        \n        if (-not $files) {\n            Write-SPInformation \"No files found for reporting\"\n            return\n        }\n        \n        Write-SPInformation \"Analyzing $($files.Count) files...\"\n        \n        # Generate report data\n        $reportData = $files | Select-Object @{\n            Name = \"FileName\"; Expression = { $_.FieldValues.FileLeafRef }\n        }, @{\n            Name = \"FilePath\"; Expression = { $_.FieldValues.FileRef }\n        }, @{\n            Name = \"FileSize\"; Expression = { $_.FieldValues.File_x0020_Size }\n        }, @{\n            Name = \"FileSizeMB\"; Expression = { [math]::Round($_.FieldValues.File_x0020_Size / 1MB, 2) }\n        }, @{\n            Name = \"FileExtension\"; Expression = { [System.IO.Path]::GetExtension($_.FieldValues.FileLeafRef) }\n        }, @{\n            Name = \"Created\"; Expression = { $_.FieldValues.Created }\n        }, @{\n            Name = \"Modified\"; Expression = { $_.FieldValues.Modified }\n        }, @{\n            Name = \"CreatedBy\"; Expression = { $_.FieldValues.Author.LookupValue }\n        }, @{\n            Name = \"ModifiedBy\"; Expression = { $_.FieldValues.Editor.LookupValue }\n        }\n        \n        # Generate statistics\n        $stats = @{\n            TotalFiles = $files.Count\n            TotalSizeMB = [math]::Round(($reportData | Measure-Object -Property FileSize -Sum).Sum / 1MB, 2)\n            FileTypes = $reportData | Group-Object FileExtension | Sort-Object Count -Descending\n            LargestFiles = $reportData | Sort-Object FileSizeMB -Descending | Select-Object -First 10\n            RecentFiles = $reportData | Sort-Object Modified -Descending | Select-Object -First 10\n            TopUploaders = $reportData | Group-Object CreatedBy | Sort-Object Count -Descending | Select-Object -First 5\n        }\n        \n        # Export detailed report\n        $reportFile = \"SharePointReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv\"\n        $reportData | Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8\n        \n        # Generate summary report\n        $summaryFile = \"SharePointSummary_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt\"\n        $summary = @\"\nSharePoint Upload Activity Report\nGenerated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')\nSite: $($Global:Config.SharePointSiteUrl)\nLibrary: $($Global:Config.DocumentLibraryName)\nTarget Folder: $($Global:Config.TargetFolderPath)\n\n=== SUMMARY STATISTICS ===\nTotal Files: $($stats.TotalFiles)\nTotal Size: $($stats.TotalSizeMB) MB\n\n=== FILE TYPES ===\n\"@\n        \n        foreach ($type in $stats.FileTypes) {\n            $sizeMB = [math]::Round(($reportData | Where-Object { $_.FileExtension -eq $type.Name } | Measure-Object -Property FileSizeMB -Sum).Sum, 2)\n            $summary += \"$($type.Name): $($type.Count) files ($sizeMB MB)`n\"\n        }\n        \n        $summary += \"`n=== LARGEST FILES ===`n\"\n        foreach ($file in $stats.LargestFiles) {\n            $summary += \"$($file.FileName): $($file.FileSizeMB) MB`n\"\n        }\n        \n        $summary += \"`n=== RECENT UPLOADS ===`n\"\n        foreach ($file in $stats.RecentFiles) {\n            $summary += \"$($file.FileName): $($file.Modified.ToString('yyyy-MM-dd HH:mm:ss')) by $($file.CreatedBy)`n\"\n        }\n        \n        $summary += \"`n=== TOP UPLOADERS ===`n\"\n        foreach ($uploader in $stats.TopUploaders) {\n            $summary += \"$($uploader.Name): $($uploader.Count) files`n\"\n        }\n        \n        $summary | Out-File -FilePath $summaryFile -Encoding UTF8\n        \n        Write-Host \"`n=== REPORT GENERATED ===\" -ForegroundColor Green\n        Write-Host \"Detailed report: $reportFile\" -ForegroundColor White\n        Write-Host \"Summary report: $summaryFile\" -ForegroundColor White\n        Write-Host \"Total files analyzed: $($stats.TotalFiles)\" -ForegroundColor White\n        Write-Host \"Total size: $($stats.TotalSizeMB) MB\" -ForegroundColor White\n        \n        Write-SPInformation \"Report generation completed successfully\"\n    }\n    catch {\n        Write-SPError \"Report generation failed\" -Exception $_.Exception\n        throw\n    }\n}\n\nfunction Invoke-ArchiveOperation {\n    Write-SPInformation \"Starting archive operation (files older than $Days days)\"\n    \n    # Similar to cleanup but moves to archive folder instead of deleting\n    # Implementation would move files to an \"Archive\" folder in SharePoint\n    \n    Write-SPWarning \"Archive operation not yet implemented\"\n    Write-SPInformation \"Use Cleanup operation to delete old files, or implement custom archiving logic\"\n}\n\nfunction Invoke-VerifyOperation {\n    Write-SPInformation \"Starting file verification operation\"\n    \n    # Connect to SharePoint\n    $Global:SPAuth = Connect-SharePointSite -Config $Global:Config -TestConnection\n    \n    # Get local files\n    $scanner = New-FileScanner -Config $Global:Config\n    $localFiles = Get-DirectoryFiles -Scanner $scanner -Path $Global:Config.LocalSourcePath\n    $validLocalFiles = $localFiles | Where-Object { $_.IsValid }\n    \n    Write-SPInformation \"Verifying $($validLocalFiles.Count) local files against SharePoint...\"\n    \n    $verificationResults = @()\n    $stats = @{\n        Verified = 0\n        Missing = 0\n        SizeMismatch = 0\n        Errors = 0\n    }\n    \n    foreach ($localFile in $validLocalFiles) {\n        try {\n            $fileName = $localFile.Name\n            $targetPath = Join-Path $Global:Config.TargetFolderPath $localFile.RelativePath -Replace \"\\\\\", \"/\"\n            \n            # Try to get file from SharePoint\n            $spFile = $null\n            try {\n                $spFile = Get-PnPFile -Url $targetPath -AsListItem -ErrorAction SilentlyContinue\n            }\n            catch {\n                # File doesn't exist in SharePoint\n            }\n            \n            $result = [PSCustomObject]@{\n                LocalFile = $localFile.FullPath\n                SharePointPath = $targetPath\n                LocalSize = $localFile.Size\n                SharePointSize = if ($spFile) { $spFile.FieldValues.File_x0020_Size } else { $null }\n                Status = \"\"\n                LastModified = $localFile.LastWriteTime\n                SharePointModified = if ($spFile) { $spFile.FieldValues.Modified } else { $null }\n            }\n            \n            if (-not $spFile) {\n                $result.Status = \"Missing\"\n                $stats.Missing++\n                Write-SPWarning \"Missing in SharePoint: $fileName\"\n            }\n            elseif ($result.LocalSize -ne $result.SharePointSize) {\n                $result.Status = \"Size Mismatch\"\n                $stats.SizeMismatch++\n                Write-SPWarning \"Size mismatch for $fileName`: Local=$($result.LocalSize), SharePoint=$($result.SharePointSize)\"\n            }\n            else {\n                $result.Status = \"Verified\"\n                $stats.Verified++\n            }\n            \n            $verificationResults += $result\n        }\n        catch {\n            Write-SPError \"Verification error for $($localFile.Name)\" -Exception $_.Exception\n            $stats.Errors++\n        }\n    }\n    \n    # Export verification results\n    $verificationFile = \"FileVerification_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv\"\n    $verificationResults | Export-Csv -Path $verificationFile -NoTypeInformation -Encoding UTF8\n    \n    # Show summary\n    Write-Host \"`n=== VERIFICATION SUMMARY ===\" -ForegroundColor Green\n    Write-Host \"Files verified: $($stats.Verified)\" -ForegroundColor Green\n    Write-Host \"Files missing: $($stats.Missing)\" -ForegroundColor Red\n    Write-Host \"Size mismatches: $($stats.SizeMismatch)\" -ForegroundColor Yellow\n    Write-Host \"Errors: $($stats.Errors)\" -ForegroundColor Red\n    Write-Host \"Verification report: $verificationFile\" -ForegroundColor White\n    \n    if ($stats.Missing -gt 0) {\n        Write-SPInformation \"Run upload operation to sync missing files\"\n    }\n    \n    Write-SPInformation \"File verification completed\"\n}\n\nfunction Cleanup-BulkResources {\n    try {\n        if ($Global:SPAuth) {\n            $Global:SPAuth.Disconnect()\n        }\n        \n        if ($Global:SPLogger) {\n            Stop-SPLogger\n        }\n    }\n    catch {\n        Write-Warning \"Error during cleanup: $($_.Exception.Message)\"\n    }\n}\n\n# Execute the bulk operation\nInvoke-BulkOperation
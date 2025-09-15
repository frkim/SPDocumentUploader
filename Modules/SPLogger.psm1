# SharePoint Document Uploader - Logging Module
# Handles comprehensive logging with PowerShell native capabilities
# Author: DevOps Team
# Version: 1.0

<#
.SYNOPSIS
    Comprehensive logging module for SharePoint Document Uploader

.DESCRIPTION
    Provides structured logging with multiple outputs, transcript logging,
    event log integration, and performance tracking. Designed for DevOps
    monitoring and troubleshooting.

.NOTES
    Requires PowerShell 5.1 or later
    Supports Windows Event Log integration for monitoring systems
#>

# Global variables for logging
$Global:SPLoggerConfig = $null
$Global:SPLoggerTranscriptPath = $null
$Global:SPLoggerEventLog = $null

class SPLogger {
    [SharePointConfig]$Config
    [string]$LogDirectory
    [string]$LogFileName
    [string]$ErrorLogFileName
    [bool]$TranscriptStarted = $false
    [System.IO.StreamWriter]$LogFileWriter
    [System.IO.StreamWriter]$ErrorLogWriter
    
    # Constructor
    SPLogger([SharePointConfig]$Config) {
        $this.Config = $Config
        $this.InitializeLogging()
    }
    
    # Initialize logging infrastructure
    [void]InitializeLogging() {
        try {
            # Create log directory
            $this.LogDirectory = $this.Config.LogDirectory
            if (-not (Test-Path $this.LogDirectory)) {
                New-Item -Path $this.LogDirectory -ItemType Directory -Force | Out-Null
            }
            
            # Set log file names
            $timestamp = Get-Date -Format "yyyyMMdd"
            $this.LogFileName = Join-Path $this.LogDirectory "SPUpload_$timestamp.log"
            $this.ErrorLogFileName = Join-Path $this.LogDirectory "SPErrors_$timestamp.log"
            
            # Initialize file writers if logging to file is enabled
            if ($this.Config.LogToFile) {
                $this.LogFileWriter = [System.IO.StreamWriter]::new($this.LogFileName, $true, [System.Text.Encoding]::UTF8)
                $this.LogFileWriter.AutoFlush = $true
                
                $this.ErrorLogWriter = [System.IO.StreamWriter]::new($this.ErrorLogFileName, $true, [System.Text.Encoding]::UTF8)
                $this.ErrorLogWriter.AutoFlush = $true
            }
            
            # Start transcript if enabled
            if ($this.Config.EnableTranscript -and -not $this.TranscriptStarted) {
                $transcriptPath = Join-Path $this.LogDirectory "SPTranscript_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
                Start-Transcript -Path $transcriptPath -Append -ErrorAction SilentlyContinue
                $this.TranscriptStarted = $true
                $Global:SPLoggerTranscriptPath = $transcriptPath
            }
            
            # Initialize Windows Event Log
            $this.InitializeEventLog()
            
            # Set global logger config
            $Global:SPLoggerConfig = $this.Config
            
            $this.WriteLog("Information", "Logging system initialized")
            $this.WriteLog("Information", "Log directory: $($this.LogDirectory)")
            $this.WriteLog("Information", "Log level: $($this.Config.LogLevel)")
        }
        catch {
            Write-Error "Failed to initialize logging: $($_.Exception.Message)"
            throw
        }
    }
    
    # Initialize Windows Event Log
    [void]InitializeEventLog() {
        try {
            $logName = "Application"
            $sourceName = "SharePointUploader"
            
            # Check if event source exists
            if (-not [System.Diagnostics.EventLog]::SourceExists($sourceName)) {
                # Create event source (requires admin rights)
                try {
                    New-EventLog -LogName $logName -Source $sourceName -ErrorAction Stop
                    $this.WriteLog("Information", "Created Windows Event Log source: $sourceName")
                }
                catch {
                    Write-Warning "Could not create Event Log source (requires admin rights): $($_.Exception.Message)"
                    return
                }
            }
            
            $Global:SPLoggerEventLog = $sourceName
        }
        catch {
            Write-Warning "Event Log initialization failed: $($_.Exception.Message)"
        }
    }
    
    # Write log entry
    [void]WriteLog([string]$Level, [string]$Message, [hashtable]$Properties = $null) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # Add properties if provided
        if ($Properties) {
            $propString = ($Properties.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
            $logEntry += " | $propString"
        }
        
        # Check log level
        if (-not $this.ShouldLog($Level)) {
            return
        }
        
        # Write to console with color coding
        $this.WriteToConsole($Level, $Message, $timestamp)
        
        # Write to log file
        if ($this.Config.LogToFile -and $this.LogFileWriter) {
            $this.LogFileWriter.WriteLine($logEntry)
        }
        
        # Write errors to separate error log
        if ($Level -in @("Error", "Critical") -and $this.Config.LogToFile -and $this.ErrorLogWriter) {
            $errorEntry = "[$timestamp] [$Level] $Message"
            if ($Properties) {
                $errorEntry += " | Context: $($Properties | ConvertTo-Json -Compress)"
            }
            $this.ErrorLogWriter.WriteLine($errorEntry)
        }
        
        # Write to Windows Event Log
        $this.WriteToEventLog($Level, $Message)
    }
    
    # Check if message should be logged based on level
    [bool]ShouldLog([string]$Level) {
        $levelHierarchy = @{
            "Verbose" = 1
            "Debug" = 2
            "Information" = 3
            "Warning" = 4
            "Error" = 5
            "Critical" = 6
        }
        
        $currentLevel = $levelHierarchy[$this.Config.LogLevel]
        $messageLevel = $levelHierarchy[$Level]
        
        return $messageLevel -ge $currentLevel
    }
    
    # Write to console with color coding
    [void]WriteToConsole([string]$Level, [string]$Message, [string]$Timestamp) {
        $colors = @{
            "Verbose" = "Gray"
            "Debug" = "Cyan"
            "Information" = "Green"
            "Warning" = "Yellow"
            "Error" = "Red"
            "Critical" = "Magenta"
        }
        
        $color = $colors[$Level]
        $consoleMessage = "[$($Timestamp.Substring(11, 8))] [$Level] $Message"
        
        Write-Host $consoleMessage -ForegroundColor $color
    }
    
    # Write to Windows Event Log
    [void]WriteToEventLog([string]$Level, [string]$Message) {
        if (-not $Global:SPLoggerEventLog) {
            return
        }
        
        try {
            $eventType = switch ($Level) {
                "Error" { "Error" }
                "Critical" { "Error" }
                "Warning" { "Warning" }
                default { "Information" }
            }
            
            $eventId = switch ($Level) {
                "Critical" { 1001 }
                "Error" { 1002 }
                "Warning" { 1003 }
                default { 1000 }
            }
            
            Write-EventLog -LogName "Application" -Source $Global:SPLoggerEventLog -EntryType $eventType -EventId $eventId -Message $Message -ErrorAction SilentlyContinue
        }
        catch {
            # Silently fail for event log errors
        }
    }
    
    # Log performance metrics
    [void]LogPerformance([string]$Operation, [double]$DurationSeconds, [hashtable]$Metrics = $null) {
        $message = "Performance: $Operation completed in $([math]::Round($DurationSeconds, 3)) seconds"
        
        $properties = @{
            Operation = $Operation
            Duration = $DurationSeconds
        }
        
        if ($Metrics) {
            $properties += $Metrics
        }
        
        $this.WriteLog("Information", $message, $properties)
    }
    
    # Log upload progress
    [void]LogUploadProgress([int]$Current, [int]$Total, [string]$CurrentFile, [long]$BytesTransferred = 0) {
        $percentComplete = if ($Total -gt 0) { ($Current / $Total) * 100 } else { 0 }
        $message = "Upload Progress: $Current/$Total ($([math]::Round($percentComplete, 1))%) - Current: $CurrentFile"
        
        $properties = @{
            Current = $Current
            Total = $Total
            PercentComplete = $percentComplete
            CurrentFile = $CurrentFile
            BytesTransferred = $BytesTransferred
        }
        
        $this.WriteLog("Information", $message, $properties)
    }
    
    # Log upload summary
    [void]LogUploadSummary([hashtable]$Statistics) {
        $this.WriteLog("Information", "=== UPLOAD SUMMARY ===")
        
        foreach ($key in $Statistics.Keys) {
            $value = $Statistics[$key]
            if ($value -is [double] -or $value -is [float]) {
                $value = [math]::Round($value, 2)
            }
            $this.WriteLog("Information", "$key`: $value")
        }
        
        $this.WriteLog("Information", "=== END SUMMARY ===")
    }
    
    # Clean up logging resources
    [void]Cleanup() {
        try {
            if ($this.LogFileWriter) {
                $this.LogFileWriter.Close()
                $this.LogFileWriter.Dispose()
            }
            
            if ($this.ErrorLogWriter) {
                $this.ErrorLogWriter.Close()
                $this.ErrorLogWriter.Dispose()
            }
            
            if ($this.TranscriptStarted) {
                Stop-Transcript -ErrorAction SilentlyContinue
                $this.TranscriptStarted = $false
            }
            
            $this.WriteLog("Information", "Logging cleanup completed")
        }
        catch {
            Write-Warning "Error during logging cleanup: $($_.Exception.Message)"
        }
    }
    
    # Clean up old log files
    [void]CleanupOldLogs([int]$DaysToKeep = 30) {
        try {
            $cutoffDate = (Get-Date).AddDays(-$DaysToKeep)
            $logFiles = Get-ChildItem -Path $this.LogDirectory -Filter "*.log" -File
            $transcriptFiles = Get-ChildItem -Path $this.LogDirectory -Filter "*.txt" -File
            
            $oldFiles = ($logFiles + $transcriptFiles) | Where-Object { $_.LastWriteTime -lt $cutoffDate }
            
            if ($oldFiles) {
                $deletedCount = 0
                foreach ($file in $oldFiles) {
                    try {
                        Remove-Item $file.FullName -Force
                        $deletedCount++
                    }
                    catch {
                        $this.WriteLog("Warning", "Could not delete old log file: $($file.Name)")
                    }
                }
                
                $this.WriteLog("Information", "Cleaned up $deletedCount old log files (older than $DaysToKeep days)")
            }
        }
        catch {
            $this.WriteLog("Error", "Error during log cleanup: $($_.Exception.Message)")
        }
    }
}

# Initialize global logger
function Initialize-SPLogger {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [SharePointConfig]$Config
    )
    
    return [SPLogger]::new($Config)
}

# Logging functions for easy use
function Write-SPLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Verbose", "Debug", "Information", "Warning", "Error", "Critical")]
        [string]$Level,
        
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Properties
    )
    
    if ($Global:SPLogger) {
        $Global:SPLogger.WriteLog($Level, $Message, $Properties)
    }
    else {
        Write-Host "[$Level] $Message" -ForegroundColor $(
            switch ($Level) {
                "Verbose" { "Gray" }
                "Debug" { "Cyan" }
                "Information" { "Green" }
                "Warning" { "Yellow" }
                "Error" { "Red" }
                "Critical" { "Magenta" }
            }
        )
    }
}

function Write-SPVerbose {
    [CmdletBinding()]
    param([string]$Message, [hashtable]$Properties)
    Write-SPLog -Level "Verbose" -Message $Message -Properties $Properties
}

function Write-SPDebug {
    [CmdletBinding()]
    param([string]$Message, [hashtable]$Properties)
    Write-SPLog -Level "Debug" -Message $Message -Properties $Properties
}

function Write-SPInformation {
    [CmdletBinding()]
    param([string]$Message, [hashtable]$Properties)
    Write-SPLog -Level "Information" -Message $Message -Properties $Properties
}

function Write-SPWarning {
    [CmdletBinding()]
    param([string]$Message, [hashtable]$Properties)
    Write-SPLog -Level "Warning" -Message $Message -Properties $Properties
}

function Write-SPError {
    [CmdletBinding()]
    param([string]$Message, [hashtable]$Properties, [Exception]$Exception)
    
    if ($Exception) {
        $Message += " | Exception: $($Exception.Message)"
        if (-not $Properties) { $Properties = @{} }
        $Properties.ExceptionType = $Exception.GetType().Name
        $Properties.StackTrace = $Exception.StackTrace
    }
    
    Write-SPLog -Level "Error" -Message $Message -Properties $Properties
}

function Write-SPCritical {
    [CmdletBinding()]
    param([string]$Message, [hashtable]$Properties)
    Write-SPLog -Level "Critical" -Message $Message -Properties $Properties
}

# Performance logging
function Measure-SPOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OperationName,
        
        [Parameter(Mandatory = $true)]
        [ScriptBlock]$ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$AdditionalMetrics
    )
    
    Write-SPInformation "Starting operation: $OperationName"
    
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    try {
        $result = & $ScriptBlock
        $stopwatch.Stop()
        
        if ($Global:SPLogger) {
            $Global:SPLogger.LogPerformance($OperationName, $stopwatch.Elapsed.TotalSeconds, $AdditionalMetrics)
        }
        
        return $result
    }
    catch {
        $stopwatch.Stop()
        Write-SPError "Operation failed: $OperationName" -Exception $_.Exception
        throw
    }
}

# Upload progress logging
function Write-SPUploadProgress {
    [CmdletBinding()]
    param(
        [int]$Current,
        [int]$Total,
        [string]$CurrentFile,
        [long]$BytesTransferred = 0
    )
    
    if ($Global:SPLogger) {
        $Global:SPLogger.LogUploadProgress($Current, $Total, $CurrentFile, $BytesTransferred)
    }
}

# Upload summary logging
function Write-SPUploadSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Statistics
    )
    
    if ($Global:SPLogger) {
        $Global:SPLogger.LogUploadSummary($Statistics)
    }
}

# Cleanup function
function Stop-SPLogger {
    [CmdletBinding()]
    param()
    
    if ($Global:SPLogger) {
        $Global:SPLogger.Cleanup()
        $Global:SPLogger = $null
    }
    
    $Global:SPLoggerConfig = $null
    $Global:SPLoggerTranscriptPath = $null
    $Global:SPLoggerEventLog = $null
}

# Export module members
Export-ModuleMember -Function Initialize-SPLogger, Write-SPLog, Write-SPVerbose, Write-SPDebug, 
                              Write-SPInformation, Write-SPWarning, Write-SPError, Write-SPCritical,
                              Measure-SPOperation, Write-SPUploadProgress, Write-SPUploadSummary, Stop-SPLogger
Export-ModuleMember -Variable Global:SPLogger
Export-ModuleMember -Cmdlet *
# SharePoint Document Uploader - PowerShell Configuration Module
# Handles configuration management for SharePoint operations
# Author: DevOps Team
# Version: 1.0

<#
.SYNOPSIS
    Configuration management module for SharePoint Document Uploader

.DESCRIPTION
    Provides configuration loading, validation, and management functions
    for SharePoint Online document upload operations. Supports JSON configuration
    files and environment variable overrides.

.NOTES
    Requires PowerShell 5.1 or later
    Configuration file should be in JSON format
#>

class SharePointConfig {
    # SharePoint Settings
    [string]$SharePointSiteUrl
    [string]$DocumentLibrary = "Documents"
    [string]$TargetFolderPath = "/Shared Documents/Upload"
    
    # Authentication Settings
    [string]$AuthMethod = "Interactive" # Interactive, AppRegistration, Certificate
    [string]$ClientId
    [string]$ClientSecret
    [string]$TenantId
    [string]$CertificateThumbprint
    [string]$CertificatePath
    [string]$Username
    
    # Local Drive Settings
    [string]$LocalSourcePath
    [bool]$IncludeSubfolders = $true
    [string[]]$FileExtensions = @(".pdf", ".docx", ".xlsx", ".pptx", ".txt", ".jpg", ".png")
    
    # Upload Settings
    [int]$BatchSize = 10
    [int]$MaxFileSizeMB = 100
    [bool]$OverwriteExisting = $false
    [bool]$CreateFolders = $true
    [int]$MaxRetries = 3
    [int]$RetryDelaySeconds = 5
    
    # Logging Settings
    [string]$LogLevel = "Information" # Verbose, Information, Warning, Error
    [bool]$LogToFile = $true
    [string]$LogDirectory = "logs"
    [bool]$EnableTranscript = $true
    
    # Performance Settings
    [int]$ThrottleLimit = 5
    [bool]$UseProgressBars = $true
    
    # Constructor
    SharePointConfig() {
        # Default constructor
    }
    
    SharePointConfig([string]$ConfigPath) {
        $this.LoadFromFile($ConfigPath)
    }
    
    # Load configuration from JSON file
    [void]LoadFromFile([string]$ConfigPath) {
        if (-not (Test-Path $ConfigPath)) {
            throw "Configuration file not found: $ConfigPath"
        }
        
        try {
            $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
            
            # Map JSON properties to class properties
            $this.SharePointSiteUrl = $config.SharePointSiteUrl
            $this.DocumentLibrary = $config.DocumentLibrary ?? $this.DocumentLibrary
            $this.TargetFolderPath = $config.TargetFolderPath ?? $this.TargetFolderPath
            
            # Authentication
            $this.AuthMethod = $config.AuthMethod ?? $this.AuthMethod
            $this.ClientId = $config.ClientId
            $this.ClientSecret = $config.ClientSecret
            $this.TenantId = $config.TenantId
            $this.CertificateThumbprint = $config.CertificateThumbprint
            $this.CertificatePath = $config.CertificatePath
            $this.Username = $config.Username
            
            # Local Drive
            $this.LocalSourcePath = $config.LocalSourcePath
            $this.IncludeSubfolders = $config.IncludeSubfolders ?? $this.IncludeSubfolders
            if ($config.FileExtensions) {
                $this.FileExtensions = $config.FileExtensions
            }
            
            # Upload Settings
            $this.BatchSize = $config.BatchSize ?? $this.BatchSize
            $this.MaxFileSizeMB = $config.MaxFileSizeMB ?? $this.MaxFileSizeMB
            $this.OverwriteExisting = $config.OverwriteExisting ?? $this.OverwriteExisting
            $this.CreateFolders = $config.CreateFolders ?? $this.CreateFolders
            $this.MaxRetries = $config.MaxRetries ?? $this.MaxRetries
            $this.RetryDelaySeconds = $config.RetryDelaySeconds ?? $this.RetryDelaySeconds
            
            # Logging
            $this.LogLevel = $config.LogLevel ?? $this.LogLevel
            $this.LogToFile = $config.LogToFile ?? $this.LogToFile
            $this.LogDirectory = $config.LogDirectory ?? $this.LogDirectory
            $this.EnableTranscript = $config.EnableTranscript ?? $this.EnableTranscript
            
            # Performance
            $this.ThrottleLimit = $config.ThrottleLimit ?? $this.ThrottleLimit
            $this.UseProgressBars = $config.UseProgressBars ?? $this.UseProgressBars
            
        }
        catch {
            throw "Failed to parse configuration file: $($_.Exception.Message)"
        }
    }
    
    # Apply environment variable overrides
    [void]ApplyEnvironmentOverrides() {
        if ($env:SP_SITE_URL) { $this.SharePointSiteUrl = $env:SP_SITE_URL }
        if ($env:SP_DOC_LIBRARY) { $this.DocumentLibrary = $env:SP_DOC_LIBRARY }
        if ($env:SP_TARGET_FOLDER) { $this.TargetFolderPath = $env:SP_TARGET_FOLDER }
        if ($env:SP_CLIENT_ID) { $this.ClientId = $env:SP_CLIENT_ID }
        if ($env:SP_CLIENT_SECRET) { $this.ClientSecret = $env:SP_CLIENT_SECRET }
        if ($env:SP_TENANT_ID) { $this.TenantId = $env:SP_TENANT_ID }
        if ($env:LOCAL_SOURCE_PATH) { $this.LocalSourcePath = $env:LOCAL_SOURCE_PATH }
        if ($env:MAX_FILE_SIZE_MB) { $this.MaxFileSizeMB = [int]$env:MAX_FILE_SIZE_MB }
        if ($env:BATCH_SIZE) { $this.BatchSize = [int]$env:BATCH_SIZE }
        if ($env:LOG_LEVEL) { $this.LogLevel = $env:LOG_LEVEL }
    }
    
    # Validate configuration
    [void]Validate() {
        $errors = @()
        
        if (-not $this.SharePointSiteUrl) {
            $errors += "SharePoint Site URL is required"
        }
        
        if (-not $this.LocalSourcePath) {
            $errors += "Local source path is required"
        }
        elseif (-not (Test-Path $this.LocalSourcePath)) {
            $errors += "Local source path does not exist: $($this.LocalSourcePath)"
        }
        
        # Validate authentication method
        switch ($this.AuthMethod) {
            "AppRegistration" {
                if (-not $this.ClientId -or -not $this.ClientSecret -or -not $this.TenantId) {
                    $errors += "App Registration requires ClientId, ClientSecret, and TenantId"
                }
            }
            "Certificate" {
                if (-not $this.ClientId -or -not $this.TenantId) {
                    $errors += "Certificate authentication requires ClientId and TenantId"
                }
                if (-not $this.CertificateThumbprint -and -not $this.CertificatePath) {
                    $errors += "Certificate authentication requires either CertificateThumbprint or CertificatePath"
                }
            }
        }
        
        # Validate numeric ranges
        if ($this.BatchSize -lt 1 -or $this.BatchSize -gt 100) {
            $errors += "BatchSize must be between 1 and 100"
        }
        
        if ($this.MaxFileSizeMB -lt 1 -or $this.MaxFileSizeMB -gt 1000) {
            $errors += "MaxFileSizeMB must be between 1 and 1000"
        }
        
        if ($this.MaxRetries -lt 0 -or $this.MaxRetries -gt 10) {
            $errors += "MaxRetries must be between 0 and 10"
        }
        
        if ($errors.Count -gt 0) {
            throw "Configuration validation failed:`n$($errors -join "`n")"
        }
    }
    
    # Get configuration summary (excluding sensitive data)
    [hashtable]GetSummary() {
        return @{
            SharePointSiteUrl = $this.SharePointSiteUrl
            DocumentLibrary = $this.DocumentLibrary
            TargetFolderPath = $this.TargetFolderPath
            AuthMethod = $this.AuthMethod
            LocalSourcePath = $this.LocalSourcePath
            IncludeSubfolders = $this.IncludeSubfolders
            FileExtensions = $this.FileExtensions -join ", "
            BatchSize = $this.BatchSize
            MaxFileSizeMB = $this.MaxFileSizeMB
            OverwriteExisting = $this.OverwriteExisting
            CreateFolders = $this.CreateFolders
            MaxRetries = $this.MaxRetries
            LogLevel = $this.LogLevel
            LogToFile = $this.LogToFile
        }
    }
    
    # Save configuration to JSON file
    [void]SaveToFile([string]$ConfigPath) {
        $configObj = @{
            SharePointSiteUrl = $this.SharePointSiteUrl
            DocumentLibrary = $this.DocumentLibrary
            TargetFolderPath = $this.TargetFolderPath
            AuthMethod = $this.AuthMethod
            ClientId = $this.ClientId
            TenantId = $this.TenantId
            CertificateThumbprint = $this.CertificateThumbprint
            CertificatePath = $this.CertificatePath
            Username = $this.Username
            LocalSourcePath = $this.LocalSourcePath
            IncludeSubfolders = $this.IncludeSubfolders
            FileExtensions = $this.FileExtensions
            BatchSize = $this.BatchSize
            MaxFileSizeMB = $this.MaxFileSizeMB
            OverwriteExisting = $this.OverwriteExisting
            CreateFolders = $this.CreateFolders
            MaxRetries = $this.MaxRetries
            RetryDelaySeconds = $this.RetryDelaySeconds
            LogLevel = $this.LogLevel
            LogToFile = $this.LogToFile
            LogDirectory = $this.LogDirectory
            EnableTranscript = $this.EnableTranscript
            ThrottleLimit = $this.ThrottleLimit
            UseProgressBars = $this.UseProgressBars
        }
        
        # Don't save sensitive data to file
        $configObj.Remove("ClientSecret")
        
        $configObj | ConvertTo-Json -Depth 3 | Out-File -FilePath $ConfigPath -Encoding UTF8
    }
}

# Function to create a default configuration file
function New-SharePointConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )
    
    if ((Test-Path $ConfigPath) -and -not $Force) {
        throw "Configuration file already exists. Use -Force to overwrite."
    }
    
    $defaultConfig = @{
        SharePointSiteUrl = "https://yourtenant.sharepoint.com/sites/yoursite"
        DocumentLibrary = "Documents"
        TargetFolderPath = "/Shared Documents/Upload"
        AuthMethod = "Interactive"
        ClientId = ""
        TenantId = ""
        LocalSourcePath = "\\\\server\\share\\documents"
        IncludeSubfolders = $true
        FileExtensions = @(".pdf", ".docx", ".xlsx", ".pptx", ".txt", ".jpg", ".png")
        BatchSize = 10
        MaxFileSizeMB = 100
        OverwriteExisting = $false
        CreateFolders = $true
        MaxRetries = 3
        RetryDelaySeconds = 5
        LogLevel = "Information"
        LogToFile = $true
        LogDirectory = "logs"
        EnableTranscript = $true
        ThrottleLimit = 5
        UseProgressBars = $true
    }
    
    $defaultConfig | ConvertTo-Json -Depth 3 | Out-File -FilePath $ConfigPath -Encoding UTF8
    Write-Information "Default configuration created at: $ConfigPath" -InformationAction Continue
}

# Function to load and validate configuration
function Get-SharePointConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ConfigPath = "config.json",
        
        [Parameter(Mandatory = $false)]
        [switch]$ApplyEnvOverrides
    )
    
    try {
        $config = [SharePointConfig]::new($ConfigPath)
        
        if ($ApplyEnvOverrides) {
            $config.ApplyEnvironmentOverrides()
        }
        
        $config.Validate()
        return $config
    }
    catch {
        Write-Error "Failed to load configuration: $($_.Exception.Message)"
        throw
    }
}

# Export module members
Export-ModuleMember -Function New-SharePointConfig, Get-SharePointConfig
Export-ModuleMember -Cmdlet *
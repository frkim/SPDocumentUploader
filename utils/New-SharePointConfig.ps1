<#
.SYNOPSIS
    SharePoint Document Uploader Configuration Setup Utility

.DESCRIPTION
    Interactive wizard to create and configure JSON configuration files
    for the SharePoint Document Uploader. Supports multiple authentication
    methods and validates settings.

.PARAMETER ConfigPath
    Output path for configuration file. Defaults to 'config.json'

.PARAMETER Template
    Load a pre-defined template: Basic, Advanced, DevOps, Certificate

.PARAMETER NonInteractive
    Run in non-interactive mode using environment variables

.PARAMETER Validate
    Only validate existing configuration without creating new

.EXAMPLE
    .\New-SharePointConfig.ps1
    # Interactive configuration wizard

.EXAMPLE
    .\New-SharePointConfig.ps1 -ConfigPath "prod-config.json" -Template DevOps
    # Create DevOps template configuration

.EXAMPLE
    .\New-SharePointConfig.ps1 -Validate -ConfigPath "existing-config.json"
    # Validate existing configuration

.NOTES
    Author: DevOps Team
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "config.json",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Basic", "Advanced", "DevOps", "Certificate")]
    [string]$Template,
    
    [Parameter(Mandatory = $false)]
    [switch]$NonInteractive,
    
    [Parameter(Mandatory = $false)]
    [switch]$Validate
)

$ErrorActionPreference = "Stop"

# Import required modules
$ModulePath = Join-Path (Split-Path $PSScriptRoot -Parent) "Modules"
Import-Module (Join-Path $ModulePath "SPConfig.psm1") -Force

function New-SharePointConfiguration {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "SharePoint Document Uploader - Configuration Setup" -ForegroundColor Cyan
        Write-Host "=" * 50 -ForegroundColor Cyan
        
        if ($Validate) {
            return Validate-ExistingConfiguration
        }
        
        if ($Template) {
            Write-Host "`nUsing template: $Template" -ForegroundColor Yellow
            $config = Get-ConfigurationTemplate -Template $Template
        }
        elseif ($NonInteractive) {
            Write-Host "`nCreating configuration from environment variables..." -ForegroundColor Yellow
            $config = Get-ConfigurationFromEnvironment
        }
        else {
            Write-Host "`nStarting interactive configuration wizard..." -ForegroundColor Yellow
            $config = Start-InteractiveWizard
        }
        
        # Validate configuration
        Write-Host "`nValidating configuration..." -ForegroundColor Yellow
        $config.Validate()
        Write-Host "✓ Configuration is valid" -ForegroundColor Green
        
        # Save configuration
        Write-Host "`nSaving configuration to: $ConfigPath" -ForegroundColor Yellow
        Save-SharePointConfig -Config $config -ConfigPath $ConfigPath
        Write-Host "✓ Configuration saved successfully" -ForegroundColor Green
        
        # Show summary
        Show-ConfigurationSummary -Config $config
        
        # Offer to test connection
        if (-not $NonInteractive -and [System.Environment]::UserInteractive) {
            $testConnection = Read-Host "`nWould you like to test the connection now? (Y/N)"
            if ($testConnection -match "^[Yy]") {
                & (Join-Path $PSScriptRoot "Test-SharePointConnection.ps1") -ConfigPath $ConfigPath
            }
        }
        
        Write-Host "`n=== SETUP COMPLETE ===" -ForegroundColor Green
        Write-Host "Configuration created successfully!" -ForegroundColor Green
        Write-Host "You can now run the uploader with:" -ForegroundColor White
        Write-Host "  .\Start-SharePointUpload.ps1 -ConfigPath '$ConfigPath'" -ForegroundColor Cyan
        
        return $true
    }
    catch {
        Write-Host "`n=== SETUP FAILED ===" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        
        if ($_.Exception.InnerException) {
            Write-Host "Details: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        }
        
        return $false
    }
}

function Validate-ExistingConfiguration {
    Write-Host "`nValidating existing configuration: $ConfigPath" -ForegroundColor Yellow
    
    if (-not (Test-Path $ConfigPath)) {
        throw "Configuration file not found: $ConfigPath"
    }
    
    $config = Get-SharePointConfig -ConfigPath $ConfigPath
    $config.Validate()
    
    Write-Host "✓ Configuration is valid" -ForegroundColor Green
    Show-ConfigurationSummary -Config $config
    
    return $true
}

function Get-ConfigurationTemplate {
    param([string]$Template)
    
    $templates = @{
        "Basic" = @{
            SharePointSiteUrl = "https://yourtenant.sharepoint.com/sites/yoursite"
            DocumentLibraryName = "Shared Documents"
            TargetFolderPath = ""
            LocalSourcePath = "C:\SharedDocuments"
            AuthenticationMethod = "Interactive"
            FileExtensions = @()
            MaxFileSizeMB = 100
            BatchSize = 10
            RetryAttempts = 3
            RetryDelaySeconds = 5
            OverwriteExisting = $false
            EnableProgressBar = $true
            LogLevel = "Information"
            LogToFile = $true
            LogToEventLog = $false
            LogFilePath = "logs\sharepoint-upload.log"
        }
        
        "Advanced" = @{
            SharePointSiteUrl = "https://yourtenant.sharepoint.com/sites/yoursite"
            DocumentLibraryName = "Documents"
            TargetFolderPath = "/Projects/Current"
            LocalSourcePath = "\\server\share\documents"
            AuthenticationMethod = "AppRegistration"
            ClientId = ""
            ClientSecret = ""
            TenantId = ""
            FileExtensions = @(".pdf", ".docx", ".xlsx", ".pptx")
            MaxFileSizeMB = 250
            BatchSize = 5
            RetryAttempts = 5
            RetryDelaySeconds = 10
            OverwriteExisting = $true
            EnableProgressBar = $true
            PreserveTimestamps = $true
            CreateFolderStructure = $true
            LogLevel = "Verbose"
            LogToFile = $true
            LogToEventLog = $true
            LogFilePath = "logs\sharepoint-upload-advanced.log"
            EventLogSource = "SharePointUploader"
        }
        
        "DevOps" = @{
            SharePointSiteUrl = "https://company.sharepoint.com/sites/devops"
            DocumentLibraryName = "Shared Documents"
            TargetFolderPath = "/Automated/$(Get-Date -Format 'yyyy-MM')"
            LocalSourcePath = "C:\BuildArtifacts"
            AuthenticationMethod = "Certificate"
            ClientId = ""
            CertificateThumbprint = ""
            TenantId = ""
            FileExtensions = @(".zip", ".msi", ".exe", ".pdf", ".log")
            MaxFileSizeMB = 500
            BatchSize = 3
            RetryAttempts = 10
            RetryDelaySeconds = 30
            OverwriteExisting = $true
            EnableProgressBar = $false
            PreserveTimestamps = $true
            CreateFolderStructure = $true
            ValidateChecksums = $true
            LogLevel = "Information"
            LogToFile = $true
            LogToEventLog = $true
            LogFilePath = "C:\Logs\SharePoint\upload.log"
            EventLogSource = "DevOpsUploader"
            EnableMetrics = $true
            MetricsEndpoint = ""
        }
        
        "Certificate" = @{
            SharePointSiteUrl = ""
            DocumentLibraryName = "Shared Documents"
            TargetFolderPath = ""
            LocalSourcePath = ""
            AuthenticationMethod = "Certificate"
            ClientId = ""
            CertificateThumbprint = ""
            TenantId = ""
            FileExtensions = @()
            MaxFileSizeMB = 100
            BatchSize = 10
            RetryAttempts = 3
            RetryDelaySeconds = 5
            OverwriteExisting = $false
            EnableProgressBar = $true
            LogLevel = "Information"
            LogToFile = $true
            LogToEventLog = $true
            LogFilePath = "logs\sharepoint-upload.log"
            EventLogSource = "SharePointUploader"
        }
    }
    
    if (-not $templates.ContainsKey($Template)) {
        throw "Unknown template: $Template"
    }
    
    $config = New-SharePointConfig
    $templateData = $templates[$Template]
    
    foreach ($key in $templateData.Keys) {
        $config.$key = $templateData[$key]
    }
    
    return $config
}

function Get-ConfigurationFromEnvironment {
    $config = New-SharePointConfig
    
    # Required settings
    $requiredVars = @{
        "SP_SITE_URL" = "SharePointSiteUrl"
        "SP_LOCAL_SOURCE" = "LocalSourcePath"
    }
    
    foreach ($envVar in $requiredVars.Keys) {
        $value = [Environment]::GetEnvironmentVariable($envVar)
        if (-not $value) {
            throw "Required environment variable not set: $envVar"
        }
        $config.($requiredVars[$envVar]) = $value
    }
    
    # Optional settings with defaults
    $optionalVars = @{
        "SP_LIBRARY_NAME" = @{ Property = "DocumentLibraryName"; Default = "Shared Documents" }
        "SP_TARGET_FOLDER" = @{ Property = "TargetFolderPath"; Default = "" }
        "SP_AUTH_METHOD" = @{ Property = "AuthenticationMethod"; Default = "Interactive" }
        "SP_CLIENT_ID" = @{ Property = "ClientId"; Default = "" }
        "SP_CLIENT_SECRET" = @{ Property = "ClientSecret"; Default = "" }
        "SP_TENANT_ID" = @{ Property = "TenantId"; Default = "" }
        "SP_CERT_THUMBPRINT" = @{ Property = "CertificateThumbprint"; Default = "" }
        "SP_FILE_EXTENSIONS" = @{ Property = "FileExtensions"; Default = @(); Transform = { $_.Split(",") | ForEach-Object { $_.Trim() } } }
        "SP_MAX_FILE_SIZE_MB" = @{ Property = "MaxFileSizeMB"; Default = 100; Transform = { [int]$_ } }
        "SP_BATCH_SIZE" = @{ Property = "BatchSize"; Default = 10; Transform = { [int]$_ } }
        "SP_LOG_LEVEL" = @{ Property = "LogLevel"; Default = "Information" }
    }
    
    foreach ($envVar in $optionalVars.Keys) {
        $value = [Environment]::GetEnvironmentVariable($envVar)
        $setting = $optionalVars[$envVar]
        
        if ($value) {
            if ($setting.Transform) {
                $value = & $setting.Transform $value
            }
            $config.($setting.Property) = $value
        } else {
            $config.($setting.Property) = $setting.Default
        }
    }
    
    return $config
}

function Start-InteractiveWizard {
    $config = New-SharePointConfig
    
    Write-Host "`n=== Basic Settings ===" -ForegroundColor Cyan
    
    # SharePoint Site URL
    do {
        $siteUrl = Read-Host "SharePoint site URL (e.g., https://company.sharepoint.com/sites/team)"
        if ($siteUrl -match "^https://.*\.sharepoint\.com/") {
            $config.SharePointSiteUrl = $siteUrl
            break
        }
        Write-Host "Please enter a valid SharePoint Online URL" -ForegroundColor Red
    } while ($true)
    
    # Document Library
    $library = Read-Host "Document library name [Shared Documents]"
    $config.DocumentLibraryName = if ($library) { $library } else { "Shared Documents" }
    
    # Target Folder
    $targetFolder = Read-Host "Target folder path (optional, e.g., /Projects/Current)"
    $config.TargetFolderPath = $targetFolder
    
    # Local Source Path
    do {
        $sourcePath = Read-Host "Local source directory path"
        if ($sourcePath -and (Test-Path $sourcePath -PathType Container)) {
            $config.LocalSourcePath = $sourcePath
            break
        }
        Write-Host "Please enter a valid directory path" -ForegroundColor Red
    } while ($true)
    
    Write-Host "`n=== Authentication Settings ===" -ForegroundColor Cyan
    
    # Authentication Method
    Write-Host "Available authentication methods:"
    Write-Host "1. Interactive (browser login)"
    Write-Host "2. App Registration (client ID + secret)"
    Write-Host "3. Certificate (client ID + certificate)"
    Write-Host "4. Managed Identity (Azure only)"
    
    do {
        $authChoice = Read-Host "Select authentication method [1-4]"
        switch ($authChoice) {
            "1" { 
                $config.AuthenticationMethod = "Interactive"
                break
            }
            "2" {
                $config.AuthenticationMethod = "AppRegistration"
                $config.ClientId = Read-Host "Client ID"
                $config.ClientSecret = Read-Host "Client Secret" -AsSecureString | ConvertFrom-SecureString
                $config.TenantId = Read-Host "Tenant ID (optional)"
                break
            }
            "3" {
                $config.AuthenticationMethod = "Certificate"
                $config.ClientId = Read-Host "Client ID"
                $config.CertificateThumbprint = Read-Host "Certificate thumbprint"
                $config.TenantId = Read-Host "Tenant ID"
                break
            }
            "4" {
                $config.AuthenticationMethod = "ManagedIdentity"
                break
            }
            default {
                Write-Host "Please select 1, 2, 3, or 4" -ForegroundColor Red
                continue
            }
        }
        break
    } while ($true)
    
    Write-Host "`n=== File Processing Settings ===" -ForegroundColor Cyan
    
    # File Extensions
    $extensions = Read-Host "File extensions to include (comma-separated, leave empty for all)"
    if ($extensions) {
        $config.FileExtensions = $extensions -split "," | ForEach-Object { $_.Trim() }
    }
    
    # Max File Size
    $maxSize = Read-Host "Maximum file size in MB [100]"
    $config.MaxFileSizeMB = if ($maxSize) { [int]$maxSize } else { 100 }
    
    # Batch Size
    $batchSize = Read-Host "Batch size (files processed together) [10]"
    $config.BatchSize = if ($batchSize) { [int]$batchSize } else { 10 }
    
    # Overwrite Setting
    $overwrite = Read-Host "Overwrite existing files? (y/N)"
    $config.OverwriteExisting = $overwrite -match "^[Yy]"
    
    Write-Host "`n=== Logging Settings ===" -ForegroundColor Cyan
    
    # Log Level
    Write-Host "Log levels: Verbose, Information, Warning, Error"
    $logLevel = Read-Host "Log level [Information]"
    $config.LogLevel = if ($logLevel) { $logLevel } else { "Information" }
    
    # File Logging
    $logToFile = Read-Host "Enable file logging? (Y/n)"
    $config.LogToFile = -not ($logToFile -match "^[Nn]")
    
    if ($config.LogToFile) {
        $logPath = Read-Host "Log file path [logs\sharepoint-upload.log]"
        $config.LogFilePath = if ($logPath) { $logPath } else { "logs\sharepoint-upload.log" }
    }
    
    # Event Log (Windows only)
    $logToEvent = Read-Host "Enable Windows Event Log? (y/N)"
    $config.LogToEventLog = $logToEvent -match "^[Yy]"
    
    return $config
}

function Show-ConfigurationSummary {
    param([SharePointConfig]$Config)
    
    $summary = $Config.GetSummary()
    
    Write-Host "`n=== CONFIGURATION SUMMARY ===" -ForegroundColor Yellow
    foreach ($key in $summary.Keys | Sort-Object) {
        $value = $summary[$key]
        if ($value -is [array]) {
            $value = $value -join ", "
        }
        # Mask sensitive information
        if ($key -match "(Secret|Password|Thumbprint)") {
            $value = "***HIDDEN***"
        }
        Write-Host "$key`: $value" -ForegroundColor White
    }
    Write-Host "=" * 30 -ForegroundColor Yellow
}

function Save-SharePointConfig {
    param(
        [SharePointConfig]$Config,
        [string]$ConfigPath
    )
    
    # Create directory if it doesn't exist
    $configDir = Split-Path $ConfigPath -Parent
    if ($configDir -and -not (Test-Path $configDir)) {
        New-Item -Path $configDir -ItemType Directory -Force | Out-Null
    }
    
    # Convert to JSON and save
    $configJson = $Config.ToJson()
    $configJson | Out-File -FilePath $ConfigPath -Encoding UTF8
    
    # Set restrictive permissions on config file
    if ($IsWindows -or $PSVersionTable.PSVersion.Major -le 5) {
        try {
            $acl = Get-Acl $ConfigPath
            $acl.SetAccessRuleProtection($true, $false)  # Remove inherited permissions
            
            # Add current user with full control
            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,
                "FullControl",
                "Allow"
            )
            $acl.SetAccessRule($accessRule)
            
            Set-Acl -Path $ConfigPath -AclObject $acl
        }
        catch {
            Write-Warning "Could not set file permissions: $($_.Exception.Message)"
        }
    }
}

# Execute the configuration setup
New-SharePointConfiguration
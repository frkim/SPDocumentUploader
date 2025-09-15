<#
.SYNOPSIS
    Test SharePoint Online connection and permissions

.DESCRIPTION
    Validates SharePoint connectivity, authentication, and upload permissions
    for the document uploader. Provides detailed diagnostic information
    for troubleshooting connection issues.

.PARAMETER ConfigPath
    Path to JSON configuration file. Defaults to '../config.json'

.PARAMETER SiteUrl
    SharePoint site URL to test. Overrides config setting.

.PARAMETER AuthMethod
    Authentication method to use: Interactive, AppRegistration, Certificate, ManagedIdentity

.PARAMETER Detailed
    Show detailed connection information and diagnostics

.EXAMPLE
    .\Test-SharePointConnection.ps1
    # Test using default configuration

.EXAMPLE
    .\Test-SharePointConnection.ps1 -SiteUrl "https://tenant.sharepoint.com/sites/team" -AuthMethod Interactive
    # Test specific site with interactive authentication

.EXAMPLE
    .\Test-SharePointConnection.ps1 -Detailed
    # Run detailed diagnostics

.NOTES
    Author: DevOps Team
    Requires: PnP.PowerShell module
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "..\config.json",
    
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Interactive", "AppRegistration", "Certificate", "ManagedIdentity")]
    [string]$AuthMethod,
    
    [Parameter(Mandatory = $false)]
    [switch]$Detailed
)

$ErrorActionPreference = "Stop"

# Import required modules
$ModulePath = Join-Path (Split-Path $PSScriptRoot -Parent) "Modules"
Import-Module (Join-Path $ModulePath "SPConfig.psm1") -Force
Import-Module (Join-Path $ModulePath "SPAuth.psm1") -Force
Import-Module (Join-Path $ModulePath "SPLogger.psm1") -Force

function Test-SharePointConnection {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "SharePoint Connection Test Utility" -ForegroundColor Cyan
        Write-Host "=" * 40 -ForegroundColor Cyan
        
        # Load configuration
        $config = Get-ConfigurationForTest
        
        # Initialize logger
        $logger = Initialize-SPLogger -Config $config -LogLevel "Information"
        
        Write-Host "`n1. Testing Environment Prerequisites..." -ForegroundColor Yellow
        Test-Prerequisites
        
        Write-Host "`n2. Validating Configuration..." -ForegroundColor Yellow
        Test-Configuration -Config $config
        
        Write-Host "`n3. Testing SharePoint Authentication..." -ForegroundColor Yellow
        $authResult = Test-Authentication -Config $config
        
        if ($authResult.Success) {
            Write-Host "`n4. Testing SharePoint Permissions..." -ForegroundColor Yellow
            Test-Permissions -Config $config -Auth $authResult.Auth
            
            if ($Detailed) {
                Write-Host "`n5. Detailed Site Information..." -ForegroundColor Yellow
                Show-DetailedSiteInfo -Config $config -Auth $authResult.Auth
            }
        }
        
        Write-Host "`n=== TEST SUMMARY ===" -ForegroundColor Green
        Write-Host "✓ Connection test completed successfully" -ForegroundColor Green
        Write-Host "Your SharePoint configuration is ready for uploads!" -ForegroundColor Green
        
        return $true
    }
    catch {
        Write-Host "`n=== TEST FAILED ===" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        
        if ($_.Exception.InnerException) {
            Write-Host "Details: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        }
        
        Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
        Write-Host "- Check your internet connection" -ForegroundColor White
        Write-Host "- Verify SharePoint site URL is correct" -ForegroundColor White
        Write-Host "- Ensure you have appropriate permissions" -ForegroundColor White
        Write-Host "- Check authentication credentials" -ForegroundColor White
        
        return $false
    }
    finally {
        # Cleanup
        if (Get-PnPConnection -ErrorAction SilentlyContinue) {
            Disconnect-PnPOnline
        }
    }
}

function Get-ConfigurationForTest {
    try {
        # Check if config file exists
        $configFullPath = Resolve-Path $ConfigPath -ErrorAction SilentlyContinue
        if (-not $configFullPath) {
            throw "Configuration file not found: $ConfigPath"
        }
        
        # Load configuration
        $config = Get-SharePointConfig -ConfigPath $configFullPath -ApplyEnvOverrides
        
        # Apply command line overrides
        if ($SiteUrl) {
            $config.SharePointSiteUrl = $SiteUrl
        }
        
        if ($AuthMethod) {
            $config.AuthenticationMethod = $AuthMethod
        }
        
        return $config
    }
    catch {
        throw "Configuration loading failed: $($_.Exception.Message)"
    }
}

function Test-Prerequisites {
    Write-Host "  Checking PowerShell version..." -NoNewline
    if ($PSVersionTable.PSVersion.Major -ge 5) {
        Write-Host " ✓" -ForegroundColor Green
        Write-Host "    Version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
    } else {
        Write-Host " ✗" -ForegroundColor Red
        throw "PowerShell 5.1 or higher is required"
    }
    
    Write-Host "  Checking PnP PowerShell module..." -NoNewline
    $pnpModule = Get-Module -Name "PnP.PowerShell" -ListAvailable | Select-Object -First 1
    if ($pnpModule) {
        Write-Host " ✓" -ForegroundColor Green
        Write-Host "    Version: $($pnpModule.Version)" -ForegroundColor Gray
    } else {
        Write-Host " ✗" -ForegroundColor Red
        throw "PnP.PowerShell module not found. Install with: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
    
    Write-Host "  Checking network connectivity..." -NoNewline
    if (Test-Connection "sharepoint.com" -Count 1 -Quiet) {
        Write-Host " ✓" -ForegroundColor Green
    } else {
        Write-Host " ✗" -ForegroundColor Red
        throw "Cannot reach SharePoint Online services"
    }
}

function Test-Configuration {
    param([SharePointConfig]$Config)
    
    Write-Host "  Validating SharePoint site URL..." -NoNewline
    if ($Config.SharePointSiteUrl -match "^https://.*\.sharepoint\.com/") {
        Write-Host " ✓" -ForegroundColor Green
        Write-Host "    URL: $($Config.SharePointSiteUrl)" -ForegroundColor Gray
    } else {
        Write-Host " ✗" -ForegroundColor Red
        throw "Invalid SharePoint site URL format"
    }
    
    Write-Host "  Checking authentication method..." -NoNewline
    if ($Config.AuthenticationMethod -in @("Interactive", "AppRegistration", "Certificate", "ManagedIdentity")) {
        Write-Host " ✓" -ForegroundColor Green
        Write-Host "    Method: $($Config.AuthenticationMethod)" -ForegroundColor Gray
    } else {
        Write-Host " ✗" -ForegroundColor Red
        throw "Invalid authentication method: $($Config.AuthenticationMethod)"
    }
    
    # Validate authentication-specific settings
    switch ($Config.AuthenticationMethod) {
        "AppRegistration" {
            Write-Host "  Checking App Registration settings..." -NoNewline
            if ($Config.ClientId -and $Config.ClientSecret) {
                Write-Host " ✓" -ForegroundColor Green
            } else {
                Write-Host " ✗" -ForegroundColor Red
                throw "ClientId and ClientSecret required for App Registration authentication"
            }
        }
        "Certificate" {
            Write-Host "  Checking certificate settings..." -NoNewline
            if ($Config.ClientId -and $Config.CertificateThumbprint) {
                Write-Host " ✓" -ForegroundColor Green
            } else {
                Write-Host " ✗" -ForegroundColor Red
                throw "ClientId and CertificateThumbprint required for Certificate authentication"
            }
        }
    }
}

function Test-Authentication {
    param([SharePointConfig]$Config)
    
    Write-Host "  Attempting SharePoint connection..." -NoNewline
    
    try {
        $auth = Connect-SharePointSite -Config $Config -TestConnection
        
        if (Get-PnPConnection -ErrorAction SilentlyContinue) {
            Write-Host " ✓" -ForegroundColor Green
            
            # Get current context
            $context = Get-PnPContext
            Write-Host "    Connected to: $($context.Url)" -ForegroundColor Gray
            Write-Host "    User: $($context.CurrentUser.LoginName)" -ForegroundColor Gray
            
            return @{
                Success = $true
                Auth = $auth
            }
        } else {
            Write-Host " ✗" -ForegroundColor Red
            throw "Connection established but no active context found"
        }
    }
    catch {
        Write-Host " ✗" -ForegroundColor Red
        throw "Authentication failed: $($_.Exception.Message)"
    }
}

function Test-Permissions {
    param(
        [SharePointConfig]$Config,
        [SharePointAuth]$Auth
    )
    
    try {
        Write-Host "  Testing site access..." -NoNewline
        $site = Get-PnPSite -ErrorAction Stop
        Write-Host " ✓" -ForegroundColor Green
        Write-Host "    Site Title: $($site.Title)" -ForegroundColor Gray
        
        Write-Host "  Testing document library access..." -NoNewline
        $library = Get-PnPList -Identity $Config.DocumentLibraryName -ErrorAction Stop
        Write-Host " ✓" -ForegroundColor Green
        Write-Host "    Library: $($library.Title)" -ForegroundColor Gray
        
        Write-Host "  Testing folder access..." -NoNewline
        if ($Config.TargetFolderPath) {
            $folder = Get-PnPFolder -Url $Config.TargetFolderPath -ErrorAction SilentlyContinue
            if ($folder) {
                Write-Host " ✓" -ForegroundColor Green
                Write-Host "    Folder exists: $($Config.TargetFolderPath)" -ForegroundColor Gray
            } else {
                Write-Host " ⚠" -ForegroundColor Yellow
                Write-Host "    Folder will be created: $($Config.TargetFolderPath)" -ForegroundColor Gray
            }
        } else {
            Write-Host " ✓" -ForegroundColor Green
            Write-Host "    Using library root" -ForegroundColor Gray
        }
        
        Write-Host "  Testing upload permissions..." -NoNewline
        # Test by creating a temporary test file
        $testFileName = "test_upload_$(Get-Date -Format 'yyyyMMddHHmmss').txt"
        $testContent = "This is a test file created by SharePoint Connection Test utility"
        
        try {
            $testFile = Add-PnPFile -Text $testContent -NewFileName $testFileName -Folder $Config.TargetFolderPath -ErrorAction Stop
            Write-Host " ✓" -ForegroundColor Green
            
            # Clean up test file
            Remove-PnPFile -Identity $testFile.Name -Folder $Config.TargetFolderPath -Force -ErrorAction SilentlyContinue
            Write-Host "    Upload/delete permissions verified" -ForegroundColor Gray
        }
        catch {
            Write-Host " ✗" -ForegroundColor Red
            throw "Upload test failed: $($_.Exception.Message)"
        }
    }
    catch {
        Write-Host " ✗" -ForegroundColor Red
        throw "Permission test failed: $($_.Exception.Message)"
    }
}

function Show-DetailedSiteInfo {
    param(
        [SharePointConfig]$Config,
        [SharePointAuth]$Auth
    )
    
    try {
        Write-Host "`n--- Site Information ---" -ForegroundColor Cyan
        $site = Get-PnPSite
        Write-Host "Title: $($site.Title)"
        Write-Host "URL: $($site.Url)"
        Write-Host "Owner: $($site.Owner.Title)"
        Write-Host "Storage Used: $([math]::Round($site.StorageUsage / 1024, 2)) GB"
        Write-Host "Storage Quota: $([math]::Round($site.StorageQuota / 1024, 2)) GB"
        
        Write-Host "`n--- Web Information ---" -ForegroundColor Cyan
        $web = Get-PnPWeb
        Write-Host "Title: $($web.Title)"
        Write-Host "Description: $($web.Description)"
        Write-Host "Language: $($web.Language)"
        Write-Host "Created: $($web.Created)"
        Write-Host "Last Modified: $($web.LastItemModifiedDate)"
        
        Write-Host "`n--- Document Library ---" -ForegroundColor Cyan
        $library = Get-PnPList -Identity $Config.DocumentLibraryName
        Write-Host "Name: $($library.Title)"
        Write-Host "Items: $($library.ItemCount)"
        Write-Host "Size: $([math]::Round($library.RootFolder.StorageUsage / 1024, 2)) KB"
        Write-Host "Version: $($library.MajorVersionLimit) major, $($library.MinorVersionLimit) minor"
        
        if ($Config.TargetFolderPath) {
            Write-Host "`n--- Target Folder ---" -ForegroundColor Cyan
            $folder = Get-PnPFolder -Url $Config.TargetFolderPath -ErrorAction SilentlyContinue
            if ($folder) {
                Write-Host "Name: $($folder.Name)"
                Write-Host "URL: $($folder.ServerRelativeUrl)"
                Write-Host "Items: $($folder.ItemCount)"
            } else {
                Write-Host "Folder does not exist and will be created during upload"
            }
        }
        
        Write-Host "`n--- Current User Permissions ---" -ForegroundColor Cyan
        $permissions = Get-PnPUserEffectivePermissions -User (Get-PnPContext).CurrentUser.LoginName -List $Config.DocumentLibraryName
        $relevantPerms = @("AddListItems", "EditListItems", "DeleteListItems", "ViewListItems", "OpenItems")
        foreach ($perm in $relevantPerms) {
            $hasPermission = $permissions.Contains($perm)
            $status = if ($hasPermission) { "✓" } else { "✗" }
            $color = if ($hasPermission) { "Green" } else { "Red" }
            Write-Host "$status $perm" -ForegroundColor $color
        }
    }
    catch {
        Write-Warning "Could not retrieve detailed information: $($_.Exception.Message)"
    }
}

# Execute the test
Test-SharePointConnection
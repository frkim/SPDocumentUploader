<#
.SYNOPSIS
    Setup and Installation Script for SharePoint Document Uploader

.DESCRIPTION
    Automates the installation and configuration of the SharePoint Document Uploader
    PowerShell solution. Installs dependencies, sets up modules, creates configuration,
    and validates the environment.

.PARAMETER InstallScope
    Scope for PowerShell module installation: CurrentUser or AllUsers

.PARAMETER ConfigPath
    Path where configuration file will be created

.PARAMETER SkipConfig
    Skip the configuration wizard

.PARAMETER Force
    Force installation and overwrite existing components

.PARAMETER Validate
    Only validate existing installation

.EXAMPLE
    .\Setup-SharePointUploader.ps1
    # Interactive setup with default options

.EXAMPLE
    .\Setup-SharePointUploader.ps1 -InstallScope AllUsers -Force
    # Install for all users and overwrite existing

.EXAMPLE
    .\Setup-SharePointUploader.ps1 -SkipConfig -Validate
    # Skip configuration and validate installation

.NOTES
    Author: DevOps Team
    Version: 1.0
    Requires: PowerShell 5.1+, Administrator rights (for AllUsers scope)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("CurrentUser", "AllUsers")]
    [string]$InstallScope = "CurrentUser",
    
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "config.json",
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipConfig,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force,
    
    [Parameter(Mandatory = $false)]
    [switch]$Validate
)

$ErrorActionPreference = "Stop"
$Global:SetupResults = @{
    Prerequisites = $false
    ModuleInstall = $false
    ModuleImport = $false
    Configuration = $false
    Validation = $false
}

function Start-SharePointUploaderSetup {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "SharePoint Document Uploader - Setup & Installation" -ForegroundColor Cyan
        Write-Host "=" * 60 -ForegroundColor Cyan
        Write-Host ""
        
        # Check if running as administrator (for AllUsers scope)
        if ($InstallScope -eq "AllUsers") {
            Test-AdminRights
        }
        
        # Step 1: Check prerequisites
        Write-Host "Step 1: Checking Prerequisites..." -ForegroundColor Yellow
        Test-Prerequisites
        Write-Host "✓ Prerequisites validated" -ForegroundColor Green
        $Global:SetupResults.Prerequisites = $true
        
        if ($Validate) {
            # Validation mode - check existing installation
            Write-Host "`nStep 2: Validating Existing Installation..." -ForegroundColor Yellow
            Test-ExistingInstallation
            Show-SetupSummary
            return $true
        }
        
        # Step 2: Install PnP PowerShell module
        Write-Host "`nStep 2: Installing PnP PowerShell Module..." -ForegroundColor Yellow
        Install-PnPPowerShell
        Write-Host "✓ PnP PowerShell module installed" -ForegroundColor Green
        $Global:SetupResults.ModuleInstall = $true
        
        # Step 3: Import and validate modules
        Write-Host "`nStep 3: Importing SharePoint Uploader Modules..." -ForegroundColor Yellow
        Import-SharePointModules
        Write-Host "✓ SharePoint modules imported successfully" -ForegroundColor Green
        $Global:SetupResults.ModuleImport = $true
        
        # Step 4: Create configuration (unless skipped)
        if (-not $SkipConfig) {
            Write-Host "`nStep 4: Creating Configuration..." -ForegroundColor Yellow
            New-Configuration
            Write-Host "✓ Configuration created successfully" -ForegroundColor Green
            $Global:SetupResults.Configuration = $true
        } else {
            Write-Host "`nStep 4: Configuration creation skipped" -ForegroundColor Yellow
            $Global:SetupResults.Configuration = "Skipped"
        }
        
        # Step 5: Validate installation
        Write-Host "`nStep 5: Validating Installation..." -ForegroundColor Yellow
        Test-Installation
        Write-Host "✓ Installation validated successfully" -ForegroundColor Green
        $Global:SetupResults.Validation = $true
        
        # Show summary and next steps
        Show-SetupSummary
        Show-NextSteps
        
        return $true
    }
    catch {
        Write-Host "`n=== SETUP FAILED ===" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        
        if ($_.Exception.InnerException) {
            Write-Host "Details: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        }
        
        Write-Host "`nSetup was not completed successfully." -ForegroundColor Red
        Write-Host "Please review the error above and try again." -ForegroundColor Yellow
        
        return $false
    }
}

function Test-AdminRights {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        throw "Administrator rights required for AllUsers installation scope. Please run as Administrator or use -InstallScope CurrentUser"
    }
    
    Write-Host "  ✓ Running with Administrator rights" -ForegroundColor Green
}

function Test-Prerequisites {
    # Check PowerShell version
    Write-Host "  Checking PowerShell version..." -NoNewline
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Host " ✗" -ForegroundColor Red
        throw "PowerShell 5.1 or higher is required. Current version: $($PSVersionTable.PSVersion)"
    }
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "    Version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
    
    # Check .NET Framework version (required for PnP PowerShell)
    Write-Host "  Checking .NET Framework..." -NoNewline
    try {
        $netVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\" -Name Release -ErrorAction SilentlyContinue
        if ($netVersion -and $netVersion.Release -ge 461808) {  # .NET 4.7.2
            Write-Host " ✓" -ForegroundColor Green
            Write-Host "    .NET Framework 4.7.2+ detected" -ForegroundColor Gray
        } else {
            Write-Host " ⚠" -ForegroundColor Yellow
            Write-Host "    .NET Framework 4.7.2+ recommended for optimal compatibility" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host " ⚠" -ForegroundColor Yellow
        Write-Host "    Could not verify .NET Framework version" -ForegroundColor Yellow
    }
    
    # Check internet connectivity
    Write-Host "  Checking internet connectivity..." -NoNewline
    if (Test-Connection "www.microsoft.com" -Count 1 -Quiet) {
        Write-Host " ✓" -ForegroundColor Green
    } else {
        Write-Host " ⚠" -ForegroundColor Yellow
        Write-Host "    Internet connectivity required for module installation" -ForegroundColor Yellow
    }
    
    # Check execution policy
    Write-Host "  Checking execution policy..." -NoNewline
    $executionPolicy = Get-ExecutionPolicy
    if ($executionPolicy -in @("Unrestricted", "RemoteSigned", "Bypass")) {
        Write-Host " ✓" -ForegroundColor Green
        Write-Host "    Current policy: $executionPolicy" -ForegroundColor Gray
    } else {
        Write-Host " ⚠" -ForegroundColor Yellow
        Write-Host "    Current policy: $executionPolicy" -ForegroundColor Yellow
        Write-Host "    You may need to set execution policy: Set-ExecutionPolicy RemoteSigned" -ForegroundColor Yellow
    }
}

function Install-PnPPowerShell {
    # Check if PnP PowerShell is already installed
    $existingModule = Get-Module -Name "PnP.PowerShell" -ListAvailable | Select-Object -First 1
    
    if ($existingModule -and -not $Force) {
        Write-Host "  PnP.PowerShell module already installed (Version: $($existingModule.Version))" -ForegroundColor Green
        
        # Check if it's a recent version
        $minVersion = [Version]"1.12.0"
        if ($existingModule.Version -lt $minVersion) {
            Write-Host "  ⚠ Older version detected. Consider updating with: Update-Module PnP.PowerShell" -ForegroundColor Yellow
        }
        return
    }
    
    Write-Host "  Installing PnP.PowerShell module..." -ForegroundColor White
    
    try {
        # Set TLS version for compatibility with PowerShell Gallery
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        # Install module with appropriate scope
        $installParams = @{
            Name = "PnP.PowerShell"
            Scope = $InstallScope
            Force = $Force
            AllowClobber = $true
        }
        
        Install-Module @installParams
        
        # Verify installation
        $installedModule = Get-Module -Name "PnP.PowerShell" -ListAvailable | Select-Object -First 1
        if ($installedModule) {
            Write-Host "    ✓ PnP.PowerShell version $($installedModule.Version) installed successfully" -ForegroundColor Green
        } else {
            throw "Module installation completed but module not found"
        }
    }
    catch {
        Write-Host "    ✗ Failed to install PnP.PowerShell module" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
        
        # Provide troubleshooting guidance
        Write-Host "`n    Troubleshooting steps:" -ForegroundColor Yellow
        Write-Host "    1. Check internet connectivity" -ForegroundColor White
        Write-Host "    2. Run as Administrator (for AllUsers scope)" -ForegroundColor White
        Write-Host "    3. Update PowerShell: Install-Module PowerShellGet -Force" -ForegroundColor White
        Write-Host "    4. Manual install: Install-Module PnP.PowerShell -Scope CurrentUser -Force" -ForegroundColor White
        
        throw
    }
}

function Import-SharePointModules {
    $modulePath = Join-Path $PSScriptRoot "Modules"
    
    if (-not (Test-Path $modulePath)) {
        throw "Modules directory not found: $modulePath"
    }
    
    $modules = @(
        "SPConfig.psm1",
        "SPAuth.psm1", 
        "SPFileScanner.psm1",
        "SPUploader.psm1",
        "SPLogger.psm1"
    )
    
    foreach ($module in $modules) {
        $modulePath = Join-Path $PSScriptRoot "Modules" $module
        
        Write-Host "    Importing $module..." -ForegroundColor White
        
        if (-not (Test-Path $modulePath)) {
            throw "Module not found: $modulePath"
        }
        
        try {
            Import-Module $modulePath -Force -ErrorAction Stop
            Write-Host "      ✓ $module imported successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "      ✗ Failed to import $module" -ForegroundColor Red
            throw "Module import failed for $module`: $($_.Exception.Message)"
        }
    }
    
    # Verify module functions are available
    $testFunctions = @(
        "New-SharePointConfig",
        "Connect-SharePointSite", 
        "New-FileScanner",
        "Start-SharePointUpload",
        "Initialize-SPLogger"
    )
    
    foreach ($function in $testFunctions) {
        if (-not (Get-Command $function -ErrorAction SilentlyContinue)) {
            throw "Required function not available after import: $function"
        }
    }
    
    Write-Host "    ✓ All module functions validated" -ForegroundColor Green
}

function New-Configuration {
    if (Test-Path $ConfigPath -and -not $Force) {
        $overwrite = Read-Host "Configuration file exists. Overwrite? (y/N)"
        if ($overwrite -notmatch "^[Yy]") {
            Write-Host "    Configuration creation skipped" -ForegroundColor Yellow
            $Global:SetupResults.Configuration = "Skipped"
            return
        }
    }
    
    # Run configuration wizard
    try {
        $configScriptPath = Join-Path $PSScriptRoot "Utils" "New-SharePointConfig.ps1"
        
        if (-not (Test-Path $configScriptPath)) {
            throw "Configuration script not found: $configScriptPath"
        }
        
        Write-Host "    Starting configuration wizard..." -ForegroundColor White
        & $configScriptPath -ConfigPath $ConfigPath
        
        if (Test-Path $ConfigPath) {
            Write-Host "    ✓ Configuration file created: $ConfigPath" -ForegroundColor Green
        } else {
            throw "Configuration file was not created"
        }
    }
    catch {
        Write-Host "    ✗ Configuration creation failed" -ForegroundColor Red
        throw "Configuration setup failed: $($_.Exception.Message)"
    }
}

function Test-Installation {
    Write-Host "    Testing module imports..." -ForegroundColor White
    
    # Test module availability
    $requiredModules = @("PnP.PowerShell")
    foreach ($module in $requiredModules) {
        $moduleTest = Get-Module -Name $module -ListAvailable
        if (-not $moduleTest) {
            throw "Required module not available: $module"
        }
        Write-Host "      ✓ $module available" -ForegroundColor Green
    }
    
    # Test SharePoint module functions
    Write-Host "    Testing SharePoint module functions..." -ForegroundColor White
    try {
        $testConfig = New-SharePointConfig
        if ($testConfig) {
            Write-Host "      ✓ Configuration module functional" -ForegroundColor Green
        }
    }
    catch {
        throw "SharePoint module test failed: $($_.Exception.Message)"
    }
    
    # Test configuration file (if exists)
    if (Test-Path $ConfigPath) {
        Write-Host "    Testing configuration file..." -ForegroundColor White
        try {
            $config = Get-SharePointConfig -ConfigPath $ConfigPath
            $config.Validate()
            Write-Host "      ✓ Configuration file is valid" -ForegroundColor Green
        }
        catch {
            Write-Host "      ⚠ Configuration validation failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "        Run configuration wizard to fix: .\Utils\New-SharePointConfig.ps1" -ForegroundColor Yellow
        }
    }
    
    Write-Host "    ✓ Installation validation completed" -ForegroundColor Green
}

function Test-ExistingInstallation {
    $results = @{
        PnPModule = $false
        SharePointModules = $false
        Configuration = $false
        UtilityScripts = $false
    }
    
    # Check PnP PowerShell module
    Write-Host "    Checking PnP.PowerShell module..." -ForegroundColor White
    $pnpModule = Get-Module -Name "PnP.PowerShell" -ListAvailable | Select-Object -First 1
    if ($pnpModule) {
        Write-Host "      ✓ PnP.PowerShell version $($pnpModule.Version) found" -ForegroundColor Green
        $results.PnPModule = $true
    } else {
        Write-Host "      ✗ PnP.PowerShell module not found" -ForegroundColor Red
    }
    
    # Check SharePoint modules
    Write-Host "    Checking SharePoint modules..." -ForegroundColor White
    $moduleFiles = @("SPConfig.psm1", "SPAuth.psm1", "SPFileScanner.psm1", "SPUploader.psm1", "SPLogger.psm1")
    $modulesPath = Join-Path $PSScriptRoot "Modules"
    $allModulesExist = $true
    
    foreach ($moduleFile in $moduleFiles) {
        $modulePath = Join-Path $modulesPath $moduleFile
        if (Test-Path $modulePath) {
            Write-Host "      ✓ $moduleFile found" -ForegroundColor Green
        } else {
            Write-Host "      ✗ $moduleFile missing" -ForegroundColor Red
            $allModulesExist = $false
        }
    }
    
    $results.SharePointModules = $allModulesExist
    
    # Check configuration
    Write-Host "    Checking configuration..." -ForegroundColor White
    if (Test-Path $ConfigPath) {
        try {
            $config = Get-SharePointConfig -ConfigPath $ConfigPath
            $config.Validate()
            Write-Host "      ✓ Configuration file is valid" -ForegroundColor Green
            $results.Configuration = $true
        }
        catch {
            Write-Host "      ⚠ Configuration file has issues: $($_.Exception.Message)" -ForegroundColor Yellow
            $results.Configuration = "Invalid"
        }
    } else {
        Write-Host "      ✗ Configuration file not found: $ConfigPath" -ForegroundColor Red
        $results.Configuration = $false
    }
    
    # Check utility scripts
    Write-Host "    Checking utility scripts..." -ForegroundColor White
    $utilityScripts = @("New-SharePointConfig.ps1", "Test-SharePointConnection.ps1", "Invoke-BulkOperations.ps1")
    $utilsPath = Join-Path $PSScriptRoot "Utils"
    $allUtilsExist = $true
    
    foreach ($script in $utilityScripts) {
        $scriptPath = Join-Path $utilsPath $script
        if (Test-Path $scriptPath) {
            Write-Host "      ✓ $script found" -ForegroundColor Green
        } else {
            Write-Host "      ✗ $script missing" -ForegroundColor Red
            $allUtilsExist = $false
        }
    }
    
    $results.UtilityScripts = $allUtilsExist
    
    # Summary
    $allValid = $results.PnPModule -and $results.SharePointModules -and ($results.Configuration -eq $true) -and $results.UtilityScripts
    
    if ($allValid) {
        Write-Host "`n    ✓ Installation is complete and valid" -ForegroundColor Green
        $Global:SetupResults.Validation = $true
    } else {
        Write-Host "`n    ⚠ Installation has issues that need to be addressed" -ForegroundColor Yellow
        $Global:SetupResults.Validation = "Issues Found"
    }
}

function Show-SetupSummary {
    Write-Host "`n" + "=" * 60 -ForegroundColor Cyan
    Write-Host "SETUP SUMMARY" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    $summaryItems = @(
        @{ Name = "Prerequisites Check"; Status = $Global:SetupResults.Prerequisites },
        @{ Name = "PnP Module Installation"; Status = $Global:SetupResults.ModuleInstall },
        @{ Name = "SharePoint Modules"; Status = $Global:SetupResults.ModuleImport },
        @{ Name = "Configuration Setup"; Status = $Global:SetupResults.Configuration },
        @{ Name = "Installation Validation"; Status = $Global:SetupResults.Validation }
    )
    
    foreach ($item in $summaryItems) {
        $status = $item.Status
        $color = switch ($status) {
            $true { "Green"; $statusText = "✓ Complete" }
            $false { "Red"; $statusText = "✗ Failed" }
            "Skipped" { "Yellow"; $statusText = "⊝ Skipped" }
            "Issues Found" { "Yellow"; $statusText = "⚠ Issues Found" }
            "Invalid" { "Yellow"; $statusText = "⚠ Invalid" }
            default { "Gray"; $statusText = "? Unknown" }
        }
        
        Write-Host "$($item.Name.PadRight(25)) : $statusText" -ForegroundColor $color
    }
    
    Write-Host "=" * 60 -ForegroundColor Cyan
}

function Show-NextSteps {
    Write-Host "`nNEXT STEPS:" -ForegroundColor Yellow
    Write-Host "1. Test your configuration:" -ForegroundColor White
    Write-Host "   .\Utils\Test-SharePointConnection.ps1 -ConfigPath '$ConfigPath'" -ForegroundColor Cyan
    
    Write-Host "`n2. Start uploading files:" -ForegroundColor White
    Write-Host "   .\Start-SharePointUpload.ps1 -ConfigPath '$ConfigPath'" -ForegroundColor Cyan
    
    Write-Host "`n3. For help and documentation:" -ForegroundColor White
    Write-Host "   Get-Help .\Start-SharePointUpload.ps1 -Full" -ForegroundColor Cyan
    Write-Host "   .\README.md" -ForegroundColor Cyan
    
    if ($Global:SetupResults.Configuration -eq "Skipped") {
        Write-Host "`n⚠ Configuration was skipped. Run the following to create configuration:" -ForegroundColor Yellow
        Write-Host "   .\Utils\New-SharePointConfig.ps1 -ConfigPath '$ConfigPath'" -ForegroundColor Cyan
    }
    
    Write-Host "`n🎉 SharePoint Document Uploader setup completed!" -ForegroundColor Green
}

# Execute setup
try {
    $success = Start-SharePointUploaderSetup
    exit $(if ($success) { 0 } else { 1 })
}
catch {
    Write-Host "`nSetup failed with error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
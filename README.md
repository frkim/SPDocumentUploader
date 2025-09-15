# SharePoint Document Uploader - PowerShell Edition

A comprehensive PowerShell solution for uploading files from local drives to SharePoint Online, designed specifically for DevOps teams and automated workflows.

## 🚀 Features

- **Multiple Authentication Methods**: Interactive, App Registration, Certificate-based, and Managed Identity
- **Batch Processing**: Efficient upload with configurable batch sizes and retry logic
- **Comprehensive Logging**: Console, file, and Windows Event Log integration
- **Progress Tracking**: Real-time upload progress with detailed statistics
- **Error Handling**: Robust retry mechanisms and detailed error reporting
- **File Filtering**: Support for file extensions, size limits, and custom filters
- **DevOps Ready**: Non-interactive mode, exit codes, and CI/CD integration
- **Bulk Operations**: Multi-source uploads, cleanup, sync, and verification utilities
- **Configuration Management**: JSON-based configuration with environment variable overrides

## 📋 Prerequisites

- **PowerShell 5.1** or higher (PowerShell 7+ recommended)
- **PnP.PowerShell Module** (automatically installed during setup)
- **SharePoint Online** access with appropriate permissions
- **.NET Framework 4.7.2** or higher (for PnP PowerShell)

## 🛠️ Installation

### Option 1: Quick Setup (Recommended)

1. **Clone or download** this repository to your local machine
2. **Run the setup script** as Administrator:
   ```powershell
   .\Setup-SharePointUploader.ps1
   ```
3. **Follow the interactive prompts** to configure your environment

### Option 2: Manual Installation

1. **Install PnP PowerShell module**:
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser -Force
   ```

2. **Import the modules**:
   ```powershell
   Import-Module .\Modules\SPConfig.psm1
   Import-Module .\Modules\SPAuth.psm1
   Import-Module .\Modules\SPFileScanner.psm1
   Import-Module .\Modules\SPUploader.psm1
   Import-Module .\Modules\SPLogger.psm1
   ```

3. **Create configuration**:
   ```powershell
   .\Utils\New-SharePointConfig.ps1
   ```

## ⚙️ Configuration

### Interactive Configuration

Run the configuration wizard to create your settings interactively:

```powershell
.\Utils\New-SharePointConfig.ps1
```

### Template-based Configuration

Use pre-defined templates for quick setup:

```powershell
# Basic configuration
.\Utils\New-SharePointConfig.ps1 -Template Basic

# DevOps configuration with certificate authentication
.\Utils\New-SharePointConfig.ps1 -Template DevOps

# Advanced configuration with app registration
.\Utils\New-SharePointConfig.ps1 -Template Advanced
```

### Environment Variables

Set configuration via environment variables for CI/CD scenarios:

```powershell
$env:SP_SITE_URL = "https://company.sharepoint.com/sites/team"
$env:SP_LOCAL_SOURCE = "C:\BuildArtifacts"
$env:SP_AUTH_METHOD = "Certificate"
$env:SP_CLIENT_ID = "your-client-id"
$env:SP_CERT_THUMBPRINT = "certificate-thumbprint"
$env:SP_TENANT_ID = "your-tenant-id"

.\Utils\New-SharePointConfig.ps1 -NonInteractive
```

### Configuration File Format

The configuration is stored in JSON format:

```json
{
  "SharePointSiteUrl": "https://company.sharepoint.com/sites/team",
  "DocumentLibraryName": "Shared Documents",
  "TargetFolderPath": "/Projects/Current",
  "LocalSourcePath": "C:\\SharedDocuments",
  "AuthenticationMethod": "Interactive",
  "FileExtensions": [".pdf", ".docx", ".xlsx"],
  "MaxFileSizeMB": 100,
  "BatchSize": 10,
  "RetryAttempts": 3,
  "LogLevel": "Information",
  "LogToFile": true,
  "LogToEventLog": true
}
```

## 🚀 Usage

### Basic Upload

Upload files using default configuration:

```powershell
.\Start-SharePointUpload.ps1
```

### Advanced Usage

```powershell
# Upload from specific directory
.\Start-SharePointUpload.ps1 -SourcePath "\\server\share\docs"

# Dry run to preview uploads
.\Start-SharePointUpload.ps1 -DryRun -Verbose

# Upload specific file types
.\Start-SharePointUpload.ps1 -Extensions ".pdf,.docx" -MaxSizeMB 50

# Non-interactive mode for automation
.\Start-SharePointUpload.ps1 -Force -ConfigPath "prod-config.json"

# Export results to CSV
.\Start-SharePointUpload.ps1 -ExportResults "upload-results.csv"
```

### Utility Operations

**Test Connection:**
```powershell
.\Utils\Test-SharePointConnection.ps1 -Detailed
```

**Bulk Operations:**
```powershell
# Multi-source upload
.\Utils\Invoke-BulkOperations.ps1 -Operation MultiSource -SourceListFile "sources.txt"

# Cleanup old files
.\Utils\Invoke-BulkOperations.ps1 -Operation Cleanup -Days 30 -DryRun

# Generate reports
.\Utils\Invoke-BulkOperations.ps1 -Operation Report

# Verify file integrity
.\Utils\Invoke-BulkOperations.ps1 -Operation Verify
```

## 🔐 Authentication Methods

### 1. Interactive Authentication
- **Use Case**: Development, manual uploads
- **Setup**: No additional configuration required
- **Process**: Browser-based login prompt

### 2. App Registration (Client ID + Secret)
- **Use Case**: Automated scripts, scheduled tasks
- **Setup**: Register app in Azure AD, generate client secret
- **Configuration**:
  ```json
  {
    "AuthenticationMethod": "AppRegistration",
    "ClientId": "your-client-id",
    "ClientSecret": "your-client-secret",
    "TenantId": "your-tenant-id"
  }
  ```

### 3. Certificate Authentication
- **Use Case**: High-security environments, DevOps pipelines
- **Setup**: Register app in Azure AD, upload certificate
- **Configuration**:
  ```json
  {
    "AuthenticationMethod": "Certificate",
    "ClientId": "your-client-id",
    "CertificateThumbprint": "cert-thumbprint",
    "TenantId": "your-tenant-id"
  }
  ```

### 4. Managed Identity
- **Use Case**: Azure VMs, Azure DevOps agents
- **Setup**: Enable managed identity on Azure resource
- **Configuration**:
  ```json
  {
    "AuthenticationMethod": "ManagedIdentity"
  }
  ```

## 📊 Logging and Monitoring

### Log Levels
- **Verbose**: Detailed debug information
- **Information**: Standard operational messages
- **Warning**: Non-critical issues
- **Error**: Critical failures only

### Log Destinations
- **Console**: Colored output with progress indicators
- **File**: Structured logs with rotation support
- **Windows Event Log**: Integration with system monitoring

### Performance Metrics
- Upload speed and throughput
- File processing statistics
- Error rates and retry patterns
- Resource utilization tracking

## 🔧 Troubleshooting

### Common Issues

**Connection Failures:**
```powershell
# Test connectivity
.\Utils\Test-SharePointConnection.ps1 -Detailed

# Check module versions
Get-Module PnP.PowerShell -ListAvailable
```

**Permission Issues:**
```powershell
# Verify permissions
Get-PnPUserEffectivePermissions -User "user@domain.com" -List "Documents"

# Test upload permissions
.\Utils\Test-SharePointConnection.ps1 -Detailed
```

**File Upload Failures:**
- Check file size limits (SharePoint Online: 250GB max)
- Verify file name restrictions (no special characters)
- Ensure sufficient storage quota
- Review retry settings in configuration

### Debug Mode

Enable detailed debugging:

```powershell
$VerbosePreference = "Continue"
$DebugPreference = "Continue"

.\Start-SharePointUpload.ps1 -Verbose -ConfigPath "config.json"
```

## 🔄 CI/CD Integration

### Azure DevOps Pipeline

```yaml
steps:
- task: PowerShell@2
  displayName: 'Upload to SharePoint'
  inputs:
    targetType: 'filePath'
    filePath: 'Scripts/Start-SharePointUpload.ps1'
    arguments: '-Force -ConfigPath "$(Pipeline.Workspace)/config.json"'
    workingDirectory: '$(Pipeline.Workspace)/SharePointUploader'
  env:
    SP_CLIENT_ID: $(SharePointClientId)
    SP_CLIENT_SECRET: $(SharePointClientSecret)
    SP_TENANT_ID: $(SharePointTenantId)
```

### GitHub Actions

```yaml
- name: Upload to SharePoint
  run: |
    .\Start-SharePointUpload.ps1 -Force -ConfigPath "config.json"
  shell: pwsh
  env:
    SP_CLIENT_ID: ${{ secrets.SHAREPOINT_CLIENT_ID }}
    SP_CLIENT_SECRET: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
    SP_TENANT_ID: ${{ secrets.SHAREPOINT_TENANT_ID }}
```

## 📁 Project Structure

```
SharePointUploader/
├── Modules/                    # PowerShell modules
│   ├── SPConfig.psm1          # Configuration management
│   ├── SPAuth.psm1            # Authentication handling
│   ├── SPFileScanner.psm1     # File discovery and validation
│   ├── SPUploader.psm1        # Upload functionality
│   └── SPLogger.psm1          # Logging and monitoring
├── Utils/                      # Utility scripts
│   ├── New-SharePointConfig.ps1      # Configuration wizard
│   ├── Test-SharePointConnection.ps1 # Connection testing
│   └── Invoke-BulkOperations.ps1     # Bulk operations
├── Start-SharePointUpload.ps1  # Main execution script
├── Setup-SharePointUploader.ps1 # Installation script
├── config.json                 # Configuration file
├── README.md                   # This documentation
└── CHANGELOG.md               # Version history
```

## 🛡️ Security Best Practices

### Configuration Security
- Store sensitive credentials in environment variables or secure vaults
- Use certificate authentication for production environments
- Restrict configuration file permissions (recommended: 600)
- Regularly rotate client secrets and certificates

### SharePoint Permissions
- Follow principle of least privilege
- Use dedicated service accounts for automation
- Implement proper folder-level permissions
- Monitor upload activities through audit logs

### Network Security
- Use HTTPS for all SharePoint connections
- Implement proper firewall rules
- Consider using private endpoints for Azure-hosted solutions
- Enable conditional access policies where appropriate

## 📄 License

This project is licensed under the MIT License. See the LICENSE file for details.

## 🏷️ Version History

### v1.0.0 (Current)
- Complete PowerShell implementation
- Multi-authentication support
- Comprehensive logging and monitoring
- Bulk operations utilities
- DevOps integration features
- Performance optimizations

---

**Note**: This PowerShell edition replaces the original Python implementation to better support DevOps teams working in Windows environments. All functionality has been preserved and enhanced with Windows-native features.
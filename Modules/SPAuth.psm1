# SharePoint Document Uploader - Authentication Module
# Handles SharePoint Online authentication using PnP PowerShell
# Author: DevOps Team
# Version: 1.0

<#
.SYNOPSIS
    SharePoint Online authentication module using PnP PowerShell

.DESCRIPTION
    Provides authentication functions for SharePoint Online using various methods:
    - Interactive authentication
    - App registration with client secret
    - Certificate-based authentication
    - Managed identity (for Azure environments)

.NOTES
    Requires PnP.PowerShell module
    Install with: Install-Module PnP.PowerShell -Scope CurrentUser
#>

# Import required modules
if (-not (Get-Module -Name PnP.PowerShell -ListAvailable)) {
    throw "PnP.PowerShell module is required. Install with: Install-Module PnP.PowerShell -Scope CurrentUser"
}

class SharePointAuth {
    [string]$SiteUrl
    [string]$AuthMethod
    [hashtable]$ConnectionParams
    [bool]$IsConnected = $false
    [datetime]$LastConnectionTime
    
    # Constructor
    SharePointAuth([SharePointConfig]$Config) {
        $this.SiteUrl = $Config.SharePointSiteUrl
        $this.AuthMethod = $Config.AuthMethod
        $this.PrepareConnectionParams($Config)
    }
    
    # Prepare connection parameters based on auth method
    [void]PrepareConnectionParams([SharePointConfig]$Config) {
        $this.ConnectionParams = @{
            Url = $this.SiteUrl
        }
        
        switch ($Config.AuthMethod) {
            "Interactive" {
                $this.ConnectionParams.Interactive = $true
                if ($Config.Username) {
                    $this.ConnectionParams.Credentials = $Config.Username
                }
            }
            
            "AppRegistration" {
                if (-not $Config.ClientId -or -not $Config.ClientSecret -or -not $Config.TenantId) {
                    throw "App Registration requires ClientId, ClientSecret, and TenantId"
                }
                $this.ConnectionParams.ClientId = $Config.ClientId
                $this.ConnectionParams.ClientSecret = $Config.ClientSecret
                $this.ConnectionParams.Tenant = $Config.TenantId
            }
            
            "Certificate" {
                if (-not $Config.ClientId -or -not $Config.TenantId) {
                    throw "Certificate authentication requires ClientId and TenantId"
                }
                $this.ConnectionParams.ClientId = $Config.ClientId
                $this.ConnectionParams.Tenant = $Config.TenantId
                
                if ($Config.CertificateThumbprint) {
                    $this.ConnectionParams.Thumbprint = $Config.CertificateThumbprint
                }
                elseif ($Config.CertificatePath) {
                    $this.ConnectionParams.CertificatePath = $Config.CertificatePath
                }
                else {
                    throw "Certificate authentication requires either CertificateThumbprint or CertificatePath"
                }
            }
            
            "ManagedIdentity" {
                $this.ConnectionParams.ManagedIdentity = $true
            }
            
            default {
                throw "Unsupported authentication method: $($Config.AuthMethod)"
            }
        }
    }
    
    # Connect to SharePoint
    [void]Connect() {
        try {
            Write-Verbose "Connecting to SharePoint using $($this.AuthMethod) authentication"
            Write-Verbose "Site URL: $($this.SiteUrl)"
            
            # Add error action to connection params
            $this.ConnectionParams.ErrorAction = "Stop"
            
            # Connect using PnP PowerShell
            Connect-PnPOnline @($this.ConnectionParams)
            
            $this.IsConnected = $true
            $this.LastConnectionTime = Get-Date
            
            Write-Information "Successfully connected to SharePoint" -InformationAction Continue
            
            # Test connection by getting web properties
            $this.TestConnection()
        }
        catch {
            $this.IsConnected = $false
            Write-Error "Failed to connect to SharePoint: $($_.Exception.Message)"
            throw
        }
    }
    
    # Test the connection
    [void]TestConnection() {
        try {
            $web = Get-PnPWeb -ErrorAction Stop
            Write-Verbose "Connected to: $($web.Title) ($($web.Url))"
            
            # Test list access
            $lists = Get-PnPList -ErrorAction Stop
            Write-Verbose "Found $($lists.Count) lists in the site"
        }
        catch {
            Write-Error "Connection test failed: $($_.Exception.Message)"
            throw
        }
    }
    
    # Disconnect from SharePoint
    [void]Disconnect() {
        try {
            if ($this.IsConnected) {
                Disconnect-PnPOnline -ErrorAction SilentlyContinue
                $this.IsConnected = $false
                Write-Verbose "Disconnected from SharePoint"
            }
        }
        catch {
            Write-Warning "Error during disconnect: $($_.Exception.Message)"
        }
    }
    
    # Ensure connection is active
    [void]EnsureConnection() {
        if (-not $this.IsConnected) {
            $this.Connect()
        }
        else {
            # Check if connection is still valid (reconnect after 30 minutes)
            $timeSinceConnection = (Get-Date) - $this.LastConnectionTime
            if ($timeSinceConnection.TotalMinutes -gt 30) {
                Write-Verbose "Connection expired, reconnecting..."
                $this.Disconnect()
                $this.Connect()
            }
        }
    }
    
    # Get document library
    [object]GetDocumentLibrary([string]$LibraryName) {
        $this.EnsureConnection()
        
        try {
            $library = Get-PnPList -Identity $LibraryName -ErrorAction Stop
            Write-Verbose "Found document library: $($library.Title)"
            return $library
        }
        catch {
            Write-Error "Failed to get document library '$LibraryName': $($_.Exception.Message)"
            throw
        }
    }
    
    # Get or create folder
    [object]GetOrCreateFolder([string]$FolderPath) {
        $this.EnsureConnection()
        
        try {
            # Clean folder path
            $cleanPath = $FolderPath.Trim("/").Replace("\", "/")
            
            # Check if folder exists
            try {
                $folder = Get-PnPFolder -Url $cleanPath -ErrorAction Stop
                Write-Verbose "Folder exists: $cleanPath"
                return $folder
            }
            catch {
                # Folder doesn't exist, create it
                Write-Information "Creating folder: $cleanPath" -InformationAction Continue
                
                # Create folder structure recursively
                $pathParts = $cleanPath -split "/"
                $currentPath = ""
                
                foreach ($part in $pathParts) {
                    if ($part) {
                        if ($currentPath) {
                            $currentPath += "/$part"
                        } else {
                            $currentPath = $part
                        }
                        
                        try {
                            $null = Get-PnPFolder -Url $currentPath -ErrorAction Stop
                        }
                        catch {
                            # Create this folder level
                            $parentPath = $currentPath.Substring(0, $currentPath.LastIndexOf("/"))
                            if (-not $parentPath) { $parentPath = "/" }
                            
                            $null = Add-PnPFolder -Name $part -Folder $parentPath -ErrorAction Stop
                            Write-Verbose "Created folder: $currentPath"
                        }
                    }
                }
                
                # Return the final folder
                return Get-PnPFolder -Url $cleanPath -ErrorAction Stop
            }
        }
        catch {
            Write-Error "Failed to get or create folder '$FolderPath': $($_.Exception.Message)"
            throw
        }
    }
    
    # Test upload permissions
    [bool]TestUploadPermissions([string]$TargetFolder = "/") {
        $this.EnsureConnection()
        
        try {
            Write-Information "Testing upload permissions..." -InformationAction Continue
            
            # Create a test file
            $testFileName = "test_upload_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $testContent = "SharePoint upload test - $(Get-Date)"
            $tempFile = Join-Path $env:TEMP $testFileName
            
            $testContent | Out-File -FilePath $tempFile -Encoding UTF8
            
            try {
                # Try to upload the test file
                $uploadResult = Add-PnPFile -Path $tempFile -Folder $TargetFolder -ErrorAction Stop
                Write-Verbose "Test upload successful: $($uploadResult.Name)"
                
                # Clean up test file from SharePoint
                try {
                    Remove-PnPFile -ServerRelativeUrl $uploadResult.ServerRelativeUrl -Force -ErrorAction SilentlyContinue
                    Write-Verbose "Test file cleaned up from SharePoint"
                }
                catch {
                    Write-Warning "Could not clean up test file from SharePoint: $testFileName"
                }
                
                return $true
            }
            catch {
                Write-Error "Upload permissions test failed: $($_.Exception.Message)"
                return $false
            }
            finally {
                # Clean up local temp file
                if (Test-Path $tempFile) {
                    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
                }
            }
        }
        catch {
            Write-Error "Error during permission test: $($_.Exception.Message)"
            return $false
        }
    }
    
    # Get site information
    [hashtable]GetSiteInfo() {
        $this.EnsureConnection()
        
        try {
            $web = Get-PnPWeb -Includes Title, Description, Url, Created, LastItemModifiedDate
            $site = Get-PnPSite -Includes Owner, Usage
            
            return @{
                Title = $web.Title
                Description = $web.Description
                Url = $web.Url
                Created = $web.Created
                LastModified = $web.LastItemModifiedDate
                Owner = $site.Owner.LoginName
                StorageUsed = $site.Usage.Storage
                StorageQuota = $site.Usage.StorageQuota
            }
        }
        catch {
            Write-Error "Failed to get site information: $($_.Exception.Message)"
            throw
        }
    }
}

# Function to create and test SharePoint connection
function Connect-SharePointSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [SharePointConfig]$Config,
        
        [Parameter(Mandatory = $false)]
        [switch]$TestConnection
    )
    
    try {
        $auth = [SharePointAuth]::new($Config)
        $auth.Connect()
        
        if ($TestConnection) {
            $siteInfo = $auth.GetSiteInfo()
            Write-Information "Connected to: $($siteInfo.Title)" -InformationAction Continue
            Write-Information "Site URL: $($siteInfo.Url)" -InformationAction Continue
            Write-Information "Owner: $($siteInfo.Owner)" -InformationAction Continue
        }
        
        return $auth
    }
    catch {
        Write-Error "Failed to connect to SharePoint: $($_.Exception.Message)"
        throw
    }
}

# Function to test SharePoint authentication
function Test-SharePointAuth {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [SharePointConfig]$Config
    )
    
    $testResults = @{
        Connection = $false
        SiteAccess = $false
        LibraryAccess = $false
        UploadPermissions = $false
        ErrorMessage = $null
    }
    
    try {
        Write-Information "Testing SharePoint authentication..." -InformationAction Continue
        
        # Test connection
        $auth = [SharePointAuth]::new($Config)
        $auth.Connect()
        $testResults.Connection = $true
        Write-Information "✓ Connection successful" -InformationAction Continue
        
        # Test site access
        $siteInfo = $auth.GetSiteInfo()
        $testResults.SiteAccess = $true
        Write-Information "✓ Site access: $($siteInfo.Title)" -InformationAction Continue
        
        # Test library access
        $library = $auth.GetDocumentLibrary($Config.DocumentLibrary)
        $testResults.LibraryAccess = $true
        Write-Information "✓ Library access: $($library.Title)" -InformationAction Continue
        
        # Test upload permissions
        $canUpload = $auth.TestUploadPermissions($Config.TargetFolderPath)
        $testResults.UploadPermissions = $canUpload
        if ($canUpload) {
            Write-Information "✓ Upload permissions verified" -InformationAction Continue
        } else {
            Write-Warning "✗ Upload permissions test failed"
        }
        
        # Clean up
        $auth.Disconnect()
        
        return $testResults
    }
    catch {
        $testResults.ErrorMessage = $_.Exception.Message
        Write-Error "Authentication test failed: $($_.Exception.Message)"
        return $testResults
    }
}

# Export module members
Export-ModuleMember -Function Connect-SharePointSite, Test-SharePointAuth
Export-ModuleMember -Cmdlet *
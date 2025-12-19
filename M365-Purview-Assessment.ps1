#Requires -Version 5.1

<#
.SYNOPSIS
    Microsoft Purview Implementation Due Diligence Assessment
.DESCRIPTION
    Comprehensive assessment of Microsoft 365 environment for Purview implementation planning.
    Collects workload information, licensing, security controls, integration points, and generates a professional Word document report.
.PARAMETER OutputPath
    Path where the JSON output file will be saved. Defaults to Desktop with timestamp.
.PARAMETER SkipModuleCheck
    Skip the PowerShell module installation check. Use if modules are already installed.
.PARAMETER SharePointTenantName
    Override the SharePoint tenant name if auto-detection fails. 
    Example: If your SharePoint URL is https://contoso-admin.sharepoint.com, use 'contoso'
.EXAMPLE
    .\M365-Purview-Assessment.ps1
    Runs the assessment with default settings.
.EXAMPLE
    .\M365-Purview-Assessment.ps1 -SharePointTenantName 'indigorx'
    Runs the assessment with a specific SharePoint tenant name.
.EXAMPLE
    .\M365-Purview-Assessment.ps1 -OutputPath "C:\Reports\Assessment.json" -SkipModuleCheck
    Runs the assessment with custom output path and skips module checks.
.NOTES
    Author: Vivek
    Version: 1.3.3
    Supports MFA authentication
    
    IMPORTANT: PowerShell 7+ is recommended for best compatibility.
    PowerShell 5.1 has a function limit that can cause issues with Microsoft.Graph module.
    This script works around the issue by importing only specific Graph sub-modules.
    
    To install PowerShell 7: https://aka.ms/powershell
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\M365_Purview_Assessment_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipModuleCheck,
    
    [Parameter(Mandatory=$false)]
    [string]$SharePointTenantName = $null
)

# Global variables
$Script:AssessmentData = @{}
$Script:ErrorLog = @()
$Script:TenantDomain = $null
$Script:ConnectedServices = @{
    Exchange = $false
    Graph = $false
    ComplianceCenter = $false
    SharePoint = $false
    Teams = $false
}

#region Helper Functions

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info','Warning','Error','Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $colors = @{
        'Info' = 'Cyan'
        'Warning' = 'Yellow'
        'Error' = 'Red'
        'Success' = 'Green'
    }
    
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $colors[$Level]
    
    if ($Level -eq 'Error') {
        $Script:ErrorLog += @{
            Timestamp = $timestamp
            Message = $Message
        }
    }
}

function Test-PowerShellVersion {
    $psVersion = $PSVersionTable.PSVersion
    Write-Log "PowerShell Version: $($psVersion.Major).$($psVersion.Minor)" -Level Info
    
    if ($psVersion.Major -eq 5) {
        Write-Log "You are using PowerShell 5.x which has a function limit that can cause issues." -Level Warning
        Write-Log "This script will work around the issue, but PowerShell 7+ is recommended." -Level Warning
        Write-Log "Download PowerShell 7: https://aka.ms/powershell" -Level Info
        Write-Host ""
    }
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-ElevatedSession {
    if (-not (Test-Administrator)) {
        Write-Log "Script requires elevation. Restarting as Administrator..." -Level Warning
        
        # Build argument list with all original parameters
        $argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"")
        
        # Add SharePointTenantName if provided
        if ($PSBoundParameters.ContainsKey('SharePointTenantName')) {
            $argList += "-SharePointTenantName"
            $argList += "'$SharePointTenantName'"
        }
        
        # Add OutputPath if provided
        if ($PSBoundParameters.ContainsKey('OutputPath')) {
            $argList += "-OutputPath"
            $argList += "`"$OutputPath`""
        }
        
        # Add SkipModuleCheck if provided
        if ($PSBoundParameters.ContainsKey('SkipModuleCheck')) {
            $argList += "-SkipModuleCheck"
        }
        
        Write-Log "Restarting with parameters: $($argList -join ' ')" -Level Info
        Start-Process powershell.exe -Verb RunAs -ArgumentList $argList
        exit
    }
}

function Install-RequiredModule {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory=$false)]
        [string]$MinimumVersion,
        
        [Parameter(Mandatory=$false)]
        [switch]$SkipImport
    )
    
    Write-Log "Checking module: $ModuleName" -Level Info
    
    $installedModule = Get-Module -ListAvailable -Name $ModuleName | 
        Sort-Object Version -Descending | 
        Select-Object -First 1
    
    if ($installedModule) {
        if ($MinimumVersion -and $installedModule.Version -lt [version]$MinimumVersion) {
            Write-Log "Module $ModuleName version $($installedModule.Version) is installed but minimum required is $MinimumVersion. Updating..." -Level Warning
            try {
                Update-Module -Name $ModuleName -Force -ErrorAction Stop
                Write-Log "Module $ModuleName updated successfully" -Level Success
            } catch {
                Write-Log "Failed to update module $ModuleName : $_" -Level Error
                throw
            }
        } else {
            Write-Log "Module $ModuleName is already installed (Version: $($installedModule.Version))" -Level Success
        }
    } else {
        Write-Log "Installing module: $ModuleName" -Level Info
        try {
            # Ensure NuGet provider is installed first (without prompting)
            $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
            if (-not $nuget) {
                Write-Log "Installing NuGet provider..." -Level Info
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
            }
            
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            Write-Log "Module $ModuleName installed successfully" -Level Success
        } catch {
            Write-Log "Failed to install module $ModuleName : $_" -Level Error
            throw
        }
    }
    
    # Import the module (unless SkipImport is specified)
    if (-not $SkipImport) {
        try {
            Import-Module $ModuleName -ErrorAction Stop -WarningAction SilentlyContinue
            Write-Log "Module $ModuleName imported successfully" -Level Success
        } catch {
            Write-Log "Failed to import module $ModuleName : $_" -Level Error
            throw
        }
    }
}

function Install-GraphSubModules {
    Write-Log "Installing Microsoft.Graph sub-modules (to avoid PowerShell 5.1 function limit)..." -Level Info
    
    # Ensure NuGet provider is installed
    $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $nuget) {
        Write-Log "Installing NuGet provider..." -Level Info
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
    }
    
    # List of specific Graph sub-modules we need
    $graphModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Identity.SignIns',
        'Microsoft.Graph.Applications',
        'Microsoft.Graph.Reports'
    )
    
    foreach ($module in $graphModules) {
        try {
            Write-Log "Checking sub-module: $module" -Level Info
            
            $installedModule = Get-Module -ListAvailable -Name $module | 
                Sort-Object Version -Descending | 
                Select-Object -First 1
            
            if (-not $installedModule) {
                Write-Log "Installing sub-module: $module" -Level Info
                Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                Write-Log "Sub-module $module installed successfully" -Level Success
            } else {
                Write-Log "Sub-module $module is already installed (Version: $($installedModule.Version))" -Level Success
            }
            
            # Import the sub-module
            Import-Module $module -ErrorAction Stop -WarningAction SilentlyContinue
            Write-Log "Sub-module $module imported successfully" -Level Success
            
        } catch {
            Write-Log "Warning: Could not install/import $module : $_" -Level Warning
        }
    }
}

function Connect-M365Services {
    Write-Log "Connecting to Microsoft 365 services..." -Level Info
    Write-Log "You will be prompted for credentials. MFA is supported." -Level Info
    
    # Track successful connections
    $Script:ConnectedServices = @{
        Exchange = $false
        Graph = $false
        ComplianceCenter = $false
        SharePoint = $false
        Teams = $false
    }
    
    try {
        # Connect to Exchange Online
        Write-Log "Connecting to Exchange Online..." -Level Info
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Connected to Exchange Online" -Level Success
        $Script:ConnectedServices.Exchange = $true
        
        # Get tenant domain from Exchange
        $acceptedDomain = Get-AcceptedDomain | Where-Object {$_.Default -eq $true}
        $Script:TenantDomain = $acceptedDomain.DomainName
        Write-Log "Tenant Domain: $Script:TenantDomain" -Level Info
        
    } catch {
        Write-Log "Failed to connect to Exchange Online: $_" -Level Error
        throw
    }
    
    # Connect to Microsoft Graph with fallback options
    $graphConnected = $false
    $graphScopes = @(
        "User.Read.All",
        "Group.Read.All",
        "Directory.Read.All",
        "Policy.Read.All",
        "Application.Read.All",
        "Organization.Read.All",
        "RoleManagement.Read.All"
    )
    
    # Try interactive browser authentication first
    try {
        Write-Log "Connecting to Microsoft Graph (browser authentication)..." -Level Info
        Connect-MgGraph -Scopes $graphScopes -NoWelcome -ErrorAction Stop
        Write-Log "Connected to Microsoft Graph" -Level Success
        $Script:ConnectedServices.Graph = $true
        $graphConnected = $true
    } catch {
        Write-Log "Browser authentication failed: $_" -Level Warning
    }
    
    # If browser auth failed, try device code flow
    if (-not $graphConnected) {
        try {
            Write-Log "Trying device code authentication..." -Level Info
            Write-Log "A code will be displayed. Copy it and visit https://microsoft.com/devicelogin" -Level Info
            Connect-MgGraph -Scopes $graphScopes -UseDeviceCode -NoWelcome -ErrorAction Stop
            Write-Log "Connected to Microsoft Graph via device code" -Level Success
            $Script:ConnectedServices.Graph = $true
            $graphConnected = $true
        } catch {
            Write-Log "Device code authentication failed: $_" -Level Warning
        }
    }
    
    # If both methods failed, try with minimal scopes
    if (-not $graphConnected) {
        try {
            Write-Log "Trying with minimal permissions..." -Level Info
            $minimalScopes = @("User.Read.All", "Directory.Read.All", "Organization.Read.All")
            Connect-MgGraph -Scopes $minimalScopes -UseDeviceCode -NoWelcome -ErrorAction Stop
            Write-Log "Connected to Microsoft Graph with minimal permissions" -Level Success
            Write-Log "Note: Some data collection may be limited due to reduced permissions" -Level Warning
            $Script:ConnectedServices.Graph = $true
            $graphConnected = $true
        } catch {
            Write-Log "Failed to connect to Microsoft Graph with all methods: $_" -Level Error
            Write-Log "Some data collection will be skipped. The assessment will continue with available data." -Level Warning
        }
    }
    
    try {
        # Connect to Security & Compliance Center
        Write-Log "Connecting to Security & Compliance Center..." -Level Info
        Connect-IPPSSession -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Log "Connected to Security & Compliance Center" -Level Success
        $Script:ConnectedServices.ComplianceCenter = $true
    } catch {
        Write-Log "Warning: Could not connect to Security & Compliance Center: $_" -Level Warning
        Write-Log "DLP and compliance data collection will be limited" -Level Warning
    }
    
    try {
        # Connect to SharePoint Online Admin
        if ($Script:TenantDomain) {
            # Use provided SharePoint tenant name if specified
            if ($SharePointTenantName) {
                $tenantName = $SharePointTenantName
                Write-Log "Using provided SharePoint tenant name: $tenantName" -Level Info
            } else {
                # Try to auto-detect from domain
                $tenantName = ($Script:TenantDomain -split '\.')[0]
                Write-Log "Auto-detected SharePoint tenant name from domain: $tenantName" -Level Info
            }
            
            $adminUrl = "https://$tenantName-admin.sharepoint.com"
            Write-Log "Connecting to SharePoint Online Admin: $adminUrl" -Level Info
            
            $maxAttempts = 3
            $attempt = 1
            $connected = $false
            
            while (-not $connected -and $attempt -le $maxAttempts) {
                try {
                    Connect-SPOService -Url $adminUrl -ErrorAction Stop
                    Write-Log "Connected to SharePoint Online" -Level Success
                    $Script:ConnectedServices.SharePoint = $true
                    $connected = $true
                } catch {
                    Write-Log "Failed to connect to SharePoint Online at $adminUrl" -Level Warning
                    
                    if ($attempt -lt $maxAttempts) {
                        Write-Host ""
                        Write-Host "Unable to connect to SharePoint admin site: $adminUrl" -ForegroundColor Yellow
                        Write-Host "This might be because:" -ForegroundColor Yellow
                        Write-Host "  1. The SharePoint tenant name is different from your Exchange domain" -ForegroundColor Yellow
                        Write-Host "  2. You don't have SharePoint Administrator permissions" -ForegroundColor Yellow
                        Write-Host "  3. The URL is incorrect" -ForegroundColor Yellow
                        Write-Host ""
                        
                        # Prompt for correct tenant name
                        $userInput = Read-Host "Enter your SharePoint tenant name (e.g., 'indigorx' for https://indigorx-admin.sharepoint.com), or press Enter to skip SharePoint"
                        
                        if ([string]::IsNullOrWhiteSpace($userInput)) {
                            Write-Log "Skipping SharePoint connection as requested by user" -Level Warning
                            break
                        } else {
                            $tenantName = $userInput.Trim()
                            $adminUrl = "https://$tenantName-admin.sharepoint.com"
                            Write-Log "Retrying with user-provided tenant name: $tenantName" -Level Info
                            $attempt++
                        }
                    } else {
                        Write-Log "Maximum connection attempts reached for SharePoint" -Level Warning
                        Write-Log "SharePoint data collection will be skipped" -Level Warning
                    }
                }
            }
        }
    } catch {
        Write-Log "Warning: Error setting up SharePoint connection: $_" -Level Warning
        Write-Log "SharePoint data collection will be skipped" -Level Warning
    }
    
    try {
        # Connect to Microsoft Teams
        Write-Log "Connecting to Microsoft Teams..." -Level Info
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        Write-Log "Connected to Microsoft Teams" -Level Success
        $Script:ConnectedServices.Teams = $true
    } catch {
        Write-Log "Warning: Could not connect to Microsoft Teams: $_" -Level Warning
        Write-Log "Teams data collection will be skipped" -Level Warning
    }
    
    # Summary of connections
    Write-Host ""
    Write-Log "Connection Summary:" -Level Info
    Write-Log "  Exchange Online: $($Script:ConnectedServices.Exchange)" -Level Info
    Write-Log "  Microsoft Graph: $($Script:ConnectedServices.Graph)" -Level Info
    Write-Log "  Compliance Center: $($Script:ConnectedServices.ComplianceCenter)" -Level Info
    Write-Log "  SharePoint Online: $($Script:ConnectedServices.SharePoint)" -Level Info
    Write-Log "  Microsoft Teams: $($Script:ConnectedServices.Teams)" -Level Info
    Write-Host ""
}

function Get-UserLicensing {
    Write-Log "Collecting user licensing information..." -Level Info
    
    if (-not $Script:ConnectedServices.Graph) {
        Write-Log "Skipping user licensing collection - Microsoft Graph not connected" -Level Warning
        $Script:AssessmentData['UserLicensing'] = @{ 
            Error = "Microsoft Graph connection not available"
            Note = "Connect to Microsoft Graph to collect licensing data"
        }
        return
    }
    
    try {
        # Get all users
        $users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses,UserType
        
        $totalUsers = $users.Count
        $enabledUsers = ($users | Where-Object {$_.AccountEnabled -eq $true}).Count
        $guestUsers = ($users | Where-Object {$_.UserType -eq 'Guest'}).Count
        
        # Get all subscribed SKUs
        $skus = Get-MgSubscribedSku
        
        $licenseSummary = @()
        foreach ($sku in $skus) {
            $licenseSummary += @{
                ProductName = $sku.SkuPartNumber
                TotalLicenses = $sku.PrepaidUnits.Enabled
                AssignedLicenses = $sku.ConsumedUnits
                AvailableLicenses = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
            }
        }
        
        # Check for Purview licensing
        $purviewLicenses = $licenseSummary | Where-Object {
            $_.ProductName -like "*COMPLIANCE*" -or 
            $_.ProductName -like "*E5*" -or 
            $_.ProductName -like "*INFORMATION_PROTECTION*"
        }
        
        $Script:AssessmentData['UserLicensing'] = @{
            TotalUsers = $totalUsers
            EnabledUsers = $enabledUsers
            GuestUsers = $guestUsers
            LicenseSummary = $licenseSummary
            PurviewLicensing = $purviewLicenses
            AssessmentDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        
        Write-Log "User licensing data collected: $totalUsers total users, $enabledUsers enabled" -Level Success
        
    } catch {
        Write-Log "Error collecting user licensing: $_" -Level Error
        $Script:AssessmentData['UserLicensing'] = @{ Error = $_.Exception.Message }
    }
}

function Get-M365WorkloadStatus {
    Write-Log "Collecting M365 workload information..." -Level Info
    
    $workloads = @{
        Exchange = @{ Enabled = $false; UsageStats = @{} }
        SharePoint = @{ Enabled = $false; UsageStats = @{} }
        OneDrive = @{ Enabled = $false; UsageStats = @{} }
        Teams = @{ Enabled = $false; UsageStats = @{} }
        Yammer = @{ Enabled = $false; UsageStats = @{} }
    }
    
    if ($Script:ConnectedServices.Exchange) {
        try {
            # Exchange Online
            $mailboxes = Get-EXOMailbox -ResultSize Unlimited
            $workloads.Exchange.Enabled = $true
            $workloads.Exchange.UsageStats = @{
                TotalMailboxes = $mailboxes.Count
                UserMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'UserMailbox'}).Count
                SharedMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'SharedMailbox'}).Count
                RoomMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'RoomMailbox'}).Count
            }
            Write-Log "Exchange Online data collected" -Level Success
        } catch {
            Write-Log "Error collecting Exchange data: $_" -Level Warning
            $workloads.Exchange.Error = $_.Exception.Message
        }
    } else {
        Write-Log "Skipping Exchange data collection - not connected" -Level Warning
        $workloads.Exchange.Error = "Not connected"
    }
    
    if ($Script:ConnectedServices.SharePoint) {
        try {
            # SharePoint Online
            $sites = Get-SPOSite -Limit All
            $workloads.SharePoint.Enabled = $true
            $workloads.SharePoint.UsageStats = @{
                TotalSites = $sites.Count
                StorageUsedGB = [math]::Round(($sites | Measure-Object StorageUsageCurrent -Sum).Sum / 1024, 2)
                StorageQuotaGB = [math]::Round(($sites | Measure-Object StorageQuota -Sum).Sum / 1024, 2)
            }
            Write-Log "SharePoint Online data collected" -Level Success
        } catch {
            Write-Log "Error collecting SharePoint data: $_" -Level Warning
            $workloads.SharePoint.Error = $_.Exception.Message
        }
    } else {
        Write-Log "Skipping SharePoint data collection - not connected" -Level Warning
        $workloads.SharePoint.Error = "Not connected"
    }
    
    if ($Script:ConnectedServices.SharePoint) {
        try {
            # OneDrive
            $oneDriveSites = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'"
            $workloads.OneDrive.Enabled = $true
            $workloads.OneDrive.UsageStats = @{
                TotalOneDriveSites = $oneDriveSites.Count
                StorageUsedGB = [math]::Round(($oneDriveSites | Measure-Object StorageUsageCurrent -Sum).Sum / 1024, 2)
            }
            Write-Log "OneDrive data collected" -Level Success
        } catch {
            Write-Log "Error collecting OneDrive data: $_" -Level Warning
            $workloads.OneDrive.Error = $_.Exception.Message
        }
    } else {
        Write-Log "Skipping OneDrive data collection - SharePoint not connected" -Level Warning
        $workloads.OneDrive.Error = "SharePoint not connected"
    }
    
    if ($Script:ConnectedServices.Teams) {
        try {
            # Microsoft Teams
            $teams = Get-Team
            $workloads.Teams.Enabled = $true
            $workloads.Teams.UsageStats = @{
                TotalTeams = $teams.Count
            }
            Write-Log "Teams data collected" -Level Success
        } catch {
            Write-Log "Error collecting Teams data: $_" -Level Warning
            $workloads.Teams.Error = $_.Exception.Message
        }
    } else {
        Write-Log "Skipping Teams data collection - not connected" -Level Warning
        $workloads.Teams.Error = "Not connected"
    }
    
    $Script:AssessmentData['M365Workloads'] = $workloads
}

function Get-SharingSettings {
    Write-Log "Collecting sharing settings..." -Level Info
    
    $sharingSettings = @{}
    
    if ($Script:ConnectedServices.SharePoint) {
        try {
            # SharePoint sharing settings
            $tenantConfig = Get-SPOTenant
            $sharingSettings['SharePoint'] = @{
                SharingCapability = $tenantConfig.SharingCapability
                RequireAnonymousLinksExpireInDays = $tenantConfig.RequireAnonymousLinksExpireInDays
                DefaultSharingLinkType = $tenantConfig.DefaultSharingLinkType
                DefaultLinkPermission = $tenantConfig.DefaultLinkPermission
                PreventExternalUsersFromResharing = $tenantConfig.PreventExternalUsersFromResharing
                ShowPeoplePickerSuggestionsForGuestUsers = $tenantConfig.ShowPeoplePickerSuggestionsForGuestUsers
                FileAnonymousLinkType = $tenantConfig.FileAnonymousLinkType
                FolderAnonymousLinkType = $tenantConfig.FolderAnonymousLinkType
            }
            Write-Log "SharePoint sharing settings collected" -Level Success
            
            # OneDrive sharing settings
            $sharingSettings['OneDrive'] = @{
                OneDriveSharingCapability = $tenantConfig.OneDriveSharingCapability
                OneDriveStorageQuota = $tenantConfig.OneDriveStorageQuota
            }
            Write-Log "OneDrive sharing settings collected" -Level Success
        } catch {
            Write-Log "Error collecting SharePoint sharing settings: $_" -Level Warning
            $sharingSettings['SharePoint'] = @{ Error = $_.Exception.Message }
        }
    } else {
        Write-Log "Skipping SharePoint/OneDrive sharing settings - not connected" -Level Warning
        $sharingSettings['SharePoint'] = @{ Error = "SharePoint not connected" }
        $sharingSettings['OneDrive'] = @{ Error = "SharePoint not connected" }
    }
    
    if ($Script:ConnectedServices.Teams) {
        try {
            # Teams external access settings
            $teamsConfig = Get-CsTenantFederationConfiguration
            $sharingSettings['Teams'] = @{
                AllowFederatedUsers = $teamsConfig.AllowFederatedUsers
                AllowPublicUsers = $teamsConfig.AllowPublicUsers
                AllowTeamsConsumer = $teamsConfig.AllowTeamsConsumer
            }
            Write-Log "Teams sharing settings collected" -Level Success
        } catch {
            Write-Log "Error collecting Teams sharing settings: $_" -Level Warning
            $sharingSettings['Teams'] = @{ Error = $_.Exception.Message }
        }
    } else {
        Write-Log "Skipping Teams sharing settings - not connected" -Level Warning
        $sharingSettings['Teams'] = @{ Error = "Teams not connected" }
    }
    
    $Script:AssessmentData['SharingSettings'] = $sharingSettings
}

function Get-DLPPolicies {
    Write-Log "Collecting DLP policies..." -Level Info
    
    if (-not $Script:ConnectedServices.ComplianceCenter) {
        Write-Log "Skipping DLP policies collection - Compliance Center not connected" -Level Warning
        $Script:AssessmentData['DLPPolicies'] = @{ 
            Error = "Compliance Center connection not available"
        }
        return
    }
    
    try {
        $dlpPolicies = Get-DlpCompliancePolicy -ErrorAction SilentlyContinue
        $dlpRules = Get-DlpComplianceRule -ErrorAction SilentlyContinue
        
        $dlpSummary = @()
        foreach ($policy in $dlpPolicies) {
            $rules = $dlpRules | Where-Object {$_.ParentPolicyName -eq $policy.Name}
            $dlpSummary += @{
                PolicyName = $policy.Name
                Enabled = $policy.Enabled
                Mode = $policy.Mode
                Workload = $policy.Workload -join ', '
                RuleCount = $rules.Count
                Rules = $rules | ForEach-Object {
                    @{
                        RuleName = $_.Name
                        Disabled = $_.Disabled
                        ContentContainsSensitiveInformation = $_.ContentContainsSensitiveInformation.Count
                    }
                }
            }
        }
        
        $Script:AssessmentData['DLPPolicies'] = @{
            TotalPolicies = $dlpPolicies.Count
            EnabledPolicies = ($dlpPolicies | Where-Object {$_.Enabled -eq $true}).Count
            Policies = $dlpSummary
        }
        
        Write-Log "DLP policies collected: $($dlpPolicies.Count) policies found" -Level Success
    } catch {
        Write-Log "Error collecting DLP policies: $_" -Level Warning
        $Script:AssessmentData['DLPPolicies'] = @{ Error = $_.Exception.Message }
    }
}

function Get-SensitivityLabels {
    Write-Log "Collecting sensitivity labels..." -Level Info
    
    if (-not $Script:ConnectedServices.ComplianceCenter) {
        Write-Log "Skipping sensitivity labels collection - Compliance Center not connected" -Level Warning
        $Script:AssessmentData['SensitivityLabels'] = @{ 
            Error = "Compliance Center connection not available"
        }
        return
    }
    
    try {
        $labels = Get-Label -ErrorAction SilentlyContinue
        $labelPolicies = Get-LabelPolicy -ErrorAction SilentlyContinue
        
        $labelSummary = @()
        foreach ($label in $labels) {
            $labelSummary += @{
                DisplayName = $label.DisplayName
                Name = $label.Name
                Tooltip = $label.Tooltip
                Disabled = $label.Disabled
                Priority = $label.Priority
                ParentLabelDisplayName = $label.ParentLabelDisplayName
                EncryptionEnabled = $label.EncryptionEnabled
                ContentType = $label.ContentType -join ', '
            }
        }
        
        $policySummary = @()
        foreach ($policy in $labelPolicies) {
            $policySummary += @{
                Name = $policy.Name
                Enabled = $policy.Enabled
                Workload = $policy.Workload -join ', '
                Labels = $policy.Labels -join ', '
            }
        }
        
        $Script:AssessmentData['SensitivityLabels'] = @{
            TotalLabels = $labels.Count
            EnabledLabels = ($labels | Where-Object {$_.Disabled -eq $false}).Count
            Labels = $labelSummary
            TotalPolicies = $labelPolicies.Count
            Policies = $policySummary
        }
        
        Write-Log "Sensitivity labels collected: $($labels.Count) labels, $($labelPolicies.Count) policies" -Level Success
    } catch {
        Write-Log "Error collecting sensitivity labels: $_" -Level Warning
        $Script:AssessmentData['SensitivityLabels'] = @{ Error = $_.Exception.Message }
    }
}

function Get-InformationProtectionPolicies {
    Write-Log "Collecting Information Protection policies..." -Level Info
    
    if (-not $Script:ConnectedServices.ComplianceCenter) {
        Write-Log "Skipping Information Protection policies collection - Compliance Center not connected" -Level Warning
        $Script:AssessmentData['InformationProtection'] = @{ 
            Error = "Compliance Center connection not available"
        }
        return
    }
    
    try {
        $retentionPolicies = Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue
        $retentionRules = Get-RetentionComplianceRule -ErrorAction SilentlyContinue
        
        $retentionSummary = @()
        foreach ($policy in $retentionPolicies) {
            $rules = $retentionRules | Where-Object {$_.ParentPolicyName -eq $policy.Name}
            $retentionSummary += @{
                Name = $policy.Name
                Enabled = $policy.Enabled
                Workload = $policy.Workload -join ', '
                RuleCount = $rules.Count
            }
        }
        
        $Script:AssessmentData['InformationProtection'] = @{
            RetentionPolicies = @{
                TotalPolicies = $retentionPolicies.Count
                EnabledPolicies = ($retentionPolicies | Where-Object {$_.Enabled -eq $true}).Count
                Policies = $retentionSummary
            }
        }
        
        Write-Log "Information Protection policies collected" -Level Success
    } catch {
        Write-Log "Error collecting Information Protection policies: $_" -Level Warning
        $Script:AssessmentData['InformationProtection'] = @{ Error = $_.Exception.Message }
    }
}

function Get-ConditionalAccessPolicies {
    Write-Log "Collecting Conditional Access policies..." -Level Info
    
    if (-not $Script:ConnectedServices.Graph) {
        Write-Log "Skipping Conditional Access collection - Microsoft Graph not connected" -Level Warning
        $Script:AssessmentData['ConditionalAccess'] = @{ 
            Error = "Microsoft Graph connection not available"
        }
        return
    }
    
    try {
        $caPolicies = Get-MgIdentityConditionalAccessPolicy -All
        
        $caSummary = @()
        foreach ($policy in $caPolicies) {
            $caSummary += @{
                DisplayName = $policy.DisplayName
                State = $policy.State
                CreatedDateTime = $policy.CreatedDateTime
                ModifiedDateTime = $policy.ModifiedDateTime
                Conditions = @{
                    UserRiskLevels = $policy.Conditions.UserRiskLevels -join ', '
                    SignInRiskLevels = $policy.Conditions.SignInRiskLevels -join ', '
                    Platforms = $policy.Conditions.Platforms.IncludePlatforms -join ', '
                    Locations = if($policy.Conditions.Locations.IncludeLocations) { $policy.Conditions.Locations.IncludeLocations -join ', ' } else { 'Any' }
                }
                GrantControls = if($policy.GrantControls) { $policy.GrantControls.BuiltInControls -join ', ' } else { 'None' }
            }
        }
        
        $Script:AssessmentData['ConditionalAccess'] = @{
            TotalPolicies = $caPolicies.Count
            EnabledPolicies = ($caPolicies | Where-Object {$_.State -eq 'enabled'}).Count
            ReportOnlyPolicies = ($caPolicies | Where-Object {$_.State -eq 'enabledForReportingButNotEnforced'}).Count
            Policies = $caSummary
        }
        
        Write-Log "Conditional Access policies collected: $($caPolicies.Count) policies" -Level Success
    } catch {
        Write-Log "Error collecting Conditional Access policies: $_" -Level Warning
        $Script:AssessmentData['ConditionalAccess'] = @{ Error = $_.Exception.Message }
    }
}

function Get-ThirdPartyApps {
    Write-Log "Collecting third-party app integrations..." -Level Info
    
    if (-not $Script:ConnectedServices.Graph) {
        Write-Log "Skipping third-party apps collection - Microsoft Graph not connected" -Level Warning
        $Script:AssessmentData['ThirdPartyApps'] = @{ 
            Error = "Microsoft Graph connection not available"
        }
        return
    }
    
    try {
        $servicePrincipals = Get-MgServicePrincipal -All -Filter "servicePrincipalType eq 'Application'"
        
        $appSummary = @()
        foreach ($sp in $servicePrincipals) {
            if ($sp.AppOwnerOrganizationId -ne (Get-MgOrganization).Id) {
                $appSummary += @{
                    DisplayName = $sp.DisplayName
                    AppId = $sp.AppId
                    Homepage = $sp.Homepage
                    PublisherName = $sp.PublisherName
                    SignInAudience = $sp.SignInAudience
                }
            }
        }
        
        $Script:AssessmentData['ThirdPartyApps'] = @{
            TotalApps = $appSummary.Count
            Apps = $appSummary | Select-Object -First 50  # Limit to first 50 for report
        }
        
        Write-Log "Third-party apps collected: $($appSummary.Count) apps" -Level Success
    } catch {
        Write-Log "Error collecting third-party apps: $_" -Level Warning
        $Script:AssessmentData['ThirdPartyApps'] = @{ Error = $_.Exception.Message }
    }
}

function Get-GuestUserPolicies {
    Write-Log "Collecting guest user and external collaboration settings..." -Level Info
    
    if (-not $Script:ConnectedServices.Graph) {
        Write-Log "Skipping guest user policies collection - Microsoft Graph not connected" -Level Warning
        $Script:AssessmentData['GuestUserPolicies'] = @{ 
            Error = "Microsoft Graph connection not available"
        }
        return
    }
    
    try {
        $authPolicy = Get-MgPolicyAuthorizationPolicy
        
        $guestSettings = @{
            GuestUserRoleId = $authPolicy.GuestUserRoleId
            AllowInvitesFrom = $authPolicy.AllowInvitesFrom
            AllowedToSignUpEmailBasedSubscriptions = $authPolicy.AllowedToSignUpEmailBasedSubscriptions
            AllowedToUseSSPR = $authPolicy.AllowedToUseSSPR
            AllowEmailVerifiedUsersToJoinOrganization = $authPolicy.AllowEmailVerifiedUsersToJoinOrganization
            BlockMsolPowerShell = $authPolicy.BlockMsolPowerShell
        }
        
        # Get guest users count
        $guestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All
        $guestSettings['TotalGuestUsers'] = $guestUsers.Count
        
        $Script:AssessmentData['GuestUserPolicies'] = $guestSettings
        
        Write-Log "Guest user policies collected: $($guestUsers.Count) guest users" -Level Success
    } catch {
        Write-Log "Error collecting guest user policies: $_" -Level Warning
        $Script:AssessmentData['GuestUserPolicies'] = @{ Error = $_.Exception.Message }
    }
}

function Get-DataRepositories {
    Write-Log "Mapping data repositories..." -Level Info
    
    $repositories = @{
        SharePointSites = @()
        OneDriveSites = @()
        Teams = @()
        ExchangeMailboxes = @()
    }
    
    if ($Script:ConnectedServices.SharePoint) {
        try {
            # SharePoint sites
            $sites = Get-SPOSite -Limit All
            foreach ($site in $sites) {
                $repositories.SharePointSites += @{
                    Url = $site.Url
                    Title = $site.Title
                    StorageUsedGB = [math]::Round($site.StorageUsageCurrent / 1024, 2)
                    StorageQuotaGB = [math]::Round($site.StorageQuota / 1024, 2)
                    SharingCapability = $site.SharingCapability
                    Template = $site.Template
                }
            }
            Write-Log "Mapped $($sites.Count) SharePoint sites" -Level Success
        } catch {
            Write-Log "Error mapping SharePoint sites: $_" -Level Warning
        }
        
        try {
            # OneDrive sites
            $oneDriveSites = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'"
            foreach ($site in $oneDriveSites) {
                $repositories.OneDriveSites += @{
                    Url = $site.Url
                    Owner = $site.Owner
                    StorageUsedGB = [math]::Round($site.StorageUsageCurrent / 1024, 2)
                }
            }
            Write-Log "Mapped $($oneDriveSites.Count) OneDrive sites" -Level Success
        } catch {
            Write-Log "Error mapping OneDrive sites: $_" -Level Warning
        }
    } else {
        Write-Log "Skipping SharePoint/OneDrive mapping - not connected" -Level Warning
    }
    
    if ($Script:ConnectedServices.Teams) {
        try {
            # Teams
            $teams = Get-Team
            foreach ($team in $teams) {
                $repositories.Teams += @{
                    DisplayName = $team.DisplayName
                    GroupId = $team.GroupId
                    Visibility = $team.Visibility
                    Archived = $team.Archived
                }
            }
            Write-Log "Mapped $($teams.Count) Teams" -Level Success
        } catch {
            Write-Log "Error mapping Teams: $_" -Level Warning
        }
    } else {
        Write-Log "Skipping Teams mapping - not connected" -Level Warning
    }
    
    if ($Script:ConnectedServices.Exchange) {
        try {
            # Exchange mailboxes (sample for performance)
            $mailboxes = Get-EXOMailbox -ResultSize 100
            foreach ($mailbox in $mailboxes) {
                $repositories.ExchangeMailboxes += @{
                    DisplayName = $mailbox.DisplayName
                    PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                    RecipientTypeDetails = $mailbox.RecipientTypeDetails
                }
            }
            Write-Log "Sampled $($mailboxes.Count) Exchange mailboxes" -Level Success
        } catch {
            Write-Log "Error mapping Exchange mailboxes: $_" -Level Warning
        }
    } else {
        Write-Log "Skipping Exchange mailbox mapping - not connected" -Level Warning
    }
    
    $Script:AssessmentData['DataRepositories'] = $repositories
}

function Export-AssessmentData {
    Write-Log "Exporting assessment data to JSON..." -Level Info
    
    try {
        $jsonContent = $Script:AssessmentData | ConvertTo-Json -Depth 10
        
        # Write UTF-8 without BOM (PowerShell's Out-File adds BOM which breaks Node.js parsing)
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($OutputPath, $jsonContent, $utf8NoBom)
        
        Write-Log "Assessment data exported to: $OutputPath" -Level Success
        return $OutputPath
    } catch {
        Write-Log "Error exporting assessment data: $_" -Level Error
        throw
    }
}

function New-AssessmentReport {
    param(
        [Parameter(Mandatory=$true)]
        [string]$JsonPath
    )
    
    Write-Log "Generating Word document report..." -Level Info
    
    # Create Node.js script to generate Word document
    $nodeScript = @'
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
        AlignmentType, WidthType, BorderStyle, HeadingLevel, ShadingType, LevelFormat } = require('docx');

// Read JSON data and remove BOM if present
let jsonContent = fs.readFileSync(process.argv[2], 'utf8');
// Remove BOM if present (handles both UTF-8 BOM and other BOMs)
if (jsonContent.charCodeAt(0) === 0xFEFF) {
    jsonContent = jsonContent.slice(1);
}
const jsonData = JSON.parse(jsonContent);

// Table border configuration
const tableBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const cellBorders = { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder };

// Helper function to create a two-column table row
function createTableRow(label, value, isHeader = false) {
    return new TableRow({
        tableHeader: isHeader,
        children: [
            new TableCell({
                borders: cellBorders,
                width: { size: 4680, type: WidthType.DXA },
                shading: isHeader ? { fill: "D5E8F0", type: ShadingType.CLEAR } : undefined,
                children: [new Paragraph({ 
                    children: [new TextRun({ text: String(label || ''), bold: isHeader, size: 22 })]
                })]
            }),
            new TableCell({
                borders: cellBorders,
                width: { size: 4680, type: WidthType.DXA },
                shading: isHeader ? { fill: "D5E8F0", type: ShadingType.CLEAR } : undefined,
                children: [new Paragraph({ 
                    children: [new TextRun({ text: String(value || 'N/A'), bold: isHeader, size: 22 })]
                })]
            })
        ]
    });
}

// Helper function to create multi-column table
function createMultiColumnTable(headers, rows) {
    const colWidth = Math.floor(9360 / headers.length);
    
    const tableRows = [
        new TableRow({
            tableHeader: true,
            children: headers.map(header => new TableCell({
                borders: cellBorders,
                width: { size: colWidth, type: WidthType.DXA },
                shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                children: [new Paragraph({ 
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: String(header), bold: true, size: 22 })]
                })]
            }))
        })
    ];
    
    rows.forEach(row => {
        tableRows.push(new TableRow({
            children: row.map(cell => new TableCell({
                borders: cellBorders,
                width: { size: colWidth, type: WidthType.DXA },
                children: [new Paragraph({ 
                    children: [new TextRun({ text: String(cell || 'N/A'), size: 22 })]
                })]
            }))
        }));
    });
    
    return new Table({
        columnWidths: Array(headers.length).fill(colWidth),
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        rows: tableRows
    });
}

// Build document sections
const sections = [];

// Title page
sections.push(
    new Paragraph({ 
        heading: HeadingLevel.TITLE, 
        children: [new TextRun("Microsoft Purview Implementation")],
        spacing: { after: 200 }
    }),
    new Paragraph({ 
        heading: HeadingLevel.TITLE, 
        children: [new TextRun("Due Diligence Assessment")],
        spacing: { after: 400 }
    }),
    new Paragraph({ 
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ 
            text: `Assessment Date: ${new Date().toLocaleDateString()}`,
            size: 24,
            bold: true
        })],
        spacing: { after: 200 }
    }),
    new Paragraph({ text: "" })
);

// Executive Summary
sections.push(
    new Paragraph({ 
        heading: HeadingLevel.HEADING_1, 
        children: [new TextRun("Executive Summary")],
        pageBreakBefore: true
    }),
    new Paragraph({ 
        children: [new TextRun("This assessment provides a comprehensive analysis of the current Microsoft 365 environment to support Microsoft Purview implementation planning.")]
    }),
    new Paragraph({ text: "" })
);

// User Licensing Section
if (jsonData.UserLicensing && !jsonData.UserLicensing.Error) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("User Licensing Overview")]
        }),
        new Paragraph({ text: "" })
    );
    
    const userTable = new Table({
        columnWidths: [4680, 4680],
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        rows: [
            createTableRow("Metric", "Value", true),
            createTableRow("Total Users", jsonData.UserLicensing.TotalUsers),
            createTableRow("Enabled Users", jsonData.UserLicensing.EnabledUsers),
            createTableRow("Guest Users", jsonData.UserLicensing.GuestUsers)
        ]
    });
    sections.push(userTable, new Paragraph({ text: "" }));
    
    // License summary with all products
    if (jsonData.UserLicensing.LicenseSummary && jsonData.UserLicensing.LicenseSummary.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("License Summary - All Products")]
            }),
            new Paragraph({ text: "" })
        );
        
        const licenseRows = jsonData.UserLicensing.LicenseSummary.map(lic => [
            lic.ProductName || 'N/A',
            lic.TotalLicenses || 0,
            lic.AssignedLicenses || 0,
            lic.AvailableLicenses || 0,
            lic.AssignedLicenses > 0 ? Math.round((lic.AssignedLicenses / lic.TotalLicenses) * 100) + '%' : '0%'
        ]);
        
        sections.push(
            createMultiColumnTable(
                ["Product Name", "Total", "Assigned", "Available", "Usage %"],
                licenseRows
            ),
            new Paragraph({ text: "" })
        );
    }
    
    // Purview-specific licensing
    if (jsonData.UserLicensing.PurviewLicensing && jsonData.UserLicensing.PurviewLicensing.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("Purview & Compliance Licensing")]
            }),
            new Paragraph({ 
                children: [new TextRun("The following licenses include Purview/Compliance features:")]
            }),
            new Paragraph({ text: "" })
        );
        
        const purviewRows = jsonData.UserLicensing.PurviewLicensing.map(lic => [
            lic.ProductName || 'N/A',
            lic.TotalLicenses || 0,
            lic.AssignedLicenses || 0,
            lic.AvailableLicenses || 0
        ]);
        
        sections.push(
            createMultiColumnTable(
                ["Product Name", "Total", "Assigned", "Available"],
                purviewRows
            ),
            new Paragraph({ text: "" })
        );
        
        const totalPurviewAssigned = jsonData.UserLicensing.PurviewLicensing.reduce((sum, lic) => sum + (lic.AssignedLicenses || 0), 0);
        const usersNeedingLicenses = jsonData.UserLicensing.EnabledUsers - totalPurviewAssigned;
        
        sections.push(
            new Paragraph({ 
                children: [new TextRun({ 
                    text: `Summary: ${totalPurviewAssigned} users have Purview licensing. ${usersNeedingLicenses > 0 ? usersNeedingLicenses + ' users may need Purview licenses.' : 'All users covered.'}`,
                    bold: true,
                    color: usersNeedingLicenses > 0 ? "FF6600" : "008000"
                })]
            }),
            new Paragraph({ text: "" })
        );
    } else {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("Purview & Compliance Licensing")]
            }),
            new Paragraph({ 
                children: [new TextRun({ 
                    text: "WARNING: No Purview or E5 Compliance licenses detected. Additional licensing may be required for full Purview functionality.",
                    bold: true,
                    color: "FF0000"
                })]
            }),
            new Paragraph({ text: "" })
        );
    }
} else if (jsonData.UserLicensing && jsonData.UserLicensing.Error) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("User Licensing Overview")]
        }),
        new Paragraph({ 
            children: [new TextRun({ text: "Data not available: " + jsonData.UserLicensing.Error, italics: true })]
        }),
        new Paragraph({ text: "" })
    );
}

// M365 Workloads Section
if (jsonData.M365Workloads) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("M365 Workloads Status")],
            pageBreakBefore: true
        }),
        new Paragraph({ text: "" })
    );
    
    Object.keys(jsonData.M365Workloads).forEach(workload => {
        const wl = jsonData.M365Workloads[workload];
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun(workload)]
            })
        );
        
        const rows = [
            createTableRow("Property", "Value", true),
            createTableRow("Enabled", wl.Enabled ? "Yes" : "No")
        ];
        
        if (wl.Error) {
            rows.push(createTableRow("Status", wl.Error));
        } else if (wl.UsageStats && Object.keys(wl.UsageStats).length > 0) {
            Object.keys(wl.UsageStats).forEach(key => {
                rows.push(createTableRow(key, wl.UsageStats[key]));
            });
        }
        
        sections.push(
            new Table({
                columnWidths: [4680, 4680],
                margins: { top: 100, bottom: 100, left: 180, right: 180 },
                rows: rows
            }),
            new Paragraph({ text: "" })
        );
    });
}

// Sharing Settings Section
if (jsonData.SharingSettings) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("Sharing Settings")],
            pageBreakBefore: true
        }),
        new Paragraph({ text: "" })
    );
    
    Object.keys(jsonData.SharingSettings).forEach(service => {
        const settings = jsonData.SharingSettings[service];
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun(service)]
            })
        );
        
        if (settings.Error) {
            sections.push(
                new Paragraph({ 
                    children: [new TextRun({ text: "Data not available: " + settings.Error, italics: true })]
                }),
                new Paragraph({ text: "" })
            );
        } else {
            const rows = [createTableRow("Setting", "Value", true)];
            Object.keys(settings).forEach(key => {
                rows.push(createTableRow(key, settings[key]));
            });
            
            sections.push(
                new Table({
                    columnWidths: [4680, 4680],
                    margins: { top: 100, bottom: 100, left: 180, right: 180 },
                    rows: rows
                }),
                new Paragraph({ text: "" })
            );
        }
    });
}

// DLP Policies Section
if (jsonData.DLPPolicies && !jsonData.DLPPolicies.Error) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("Data Loss Prevention (DLP) Policies")],
            pageBreakBefore: true
        }),
        new Paragraph({ text: "" })
    );
    
    const dlpSummary = new Table({
        columnWidths: [4680, 4680],
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        rows: [
            createTableRow("Metric", "Value", true),
            createTableRow("Total Policies", jsonData.DLPPolicies.TotalPolicies || 0),
            createTableRow("Enabled Policies", jsonData.DLPPolicies.EnabledPolicies || 0)
        ]
    });
    sections.push(dlpSummary, new Paragraph({ text: "" }));
    
    if (jsonData.DLPPolicies.Policies && jsonData.DLPPolicies.Policies.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("DLP Policy Details")]
            }),
            new Paragraph({ text: "" })
        );
        
        const policyRows = jsonData.DLPPolicies.Policies.map(pol => [
            pol.PolicyName || 'N/A',
            pol.Enabled ? "Yes" : "No",
            pol.Mode || 'N/A',
            pol.Workload || 'N/A',
            pol.RuleCount || 0
        ]);
        
        sections.push(
            createMultiColumnTable(
                ["Policy Name", "Enabled", "Mode", "Workload", "Rules"],
                policyRows
            ),
            new Paragraph({ text: "" })
        );
        
        // Add rule details for each policy
        jsonData.DLPPolicies.Policies.forEach(policy => {
            if (policy.Rules && policy.Rules.length > 0) {
                sections.push(
                    new Paragraph({ 
                        children: [new TextRun({ text: `Rules for: ${policy.PolicyName}`, bold: true })]
                    }),
                    new Paragraph({ text: "" })
                );
                
                const ruleRows = policy.Rules.map(rule => [
                    rule.RuleName || 'N/A',
                    rule.Disabled ? 'Disabled' : 'Enabled',
                    rule.ContentContainsSensitiveInformation || 0
                ]);
                
                sections.push(
                    createMultiColumnTable(
                        ["Rule Name", "Status", "Sensitive Info Types"],
                        ruleRows
                    ),
                    new Paragraph({ text: "" })
                );
            }
        });
    } else {
        sections.push(
            new Paragraph({ 
                children: [new TextRun({ text: "No DLP policies configured. This is a HIGH PRIORITY compliance gap.", bold: true, color: "FF0000" })]
            }),
            new Paragraph({ text: "" })
        );
    }
}

// Sensitivity Labels Section
if (jsonData.SensitivityLabels && !jsonData.SensitivityLabels.Error) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("Sensitivity Labels")],
            pageBreakBefore: true
        }),
        new Paragraph({ text: "" })
    );
    
    const labelSummary = new Table({
        columnWidths: [4680, 4680],
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        rows: [
            createTableRow("Metric", "Value", true),
            createTableRow("Total Labels", jsonData.SensitivityLabels.TotalLabels || 0),
            createTableRow("Enabled Labels", jsonData.SensitivityLabels.EnabledLabels || 0),
            createTableRow("Total Policies", jsonData.SensitivityLabels.TotalPolicies || 0)
        ]
    });
    sections.push(labelSummary, new Paragraph({ text: "" }));
    
    // Detailed label information
    if (jsonData.SensitivityLabels.Labels && jsonData.SensitivityLabels.Labels.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("Label Details")]
            }),
            new Paragraph({ text: "" })
        );
        
        const labelRows = jsonData.SensitivityLabels.Labels.map(label => [
            label.DisplayName || 'N/A',
            label.Tooltip || 'No description',
            label.EncryptionEnabled ? 'Yes' : 'No',
            label.Disabled ? 'Disabled' : 'Enabled',
            label.Priority || 'N/A'
        ]);
        
        sections.push(
            createMultiColumnTable(
                ["Label Name", "Description", "Encryption", "Status", "Priority"],
                labelRows
            ),
            new Paragraph({ text: "" })
        );
    }
    
    // Label policies
    if (jsonData.SensitivityLabels.Policies && jsonData.SensitivityLabels.Policies.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("Label Policies")]
            }),
            new Paragraph({ text: "" })
        );
        
        const policyRows = jsonData.SensitivityLabels.Policies.map(policy => [
            policy.Name || 'N/A',
            policy.Enabled ? 'Enabled' : 'Disabled',
            policy.Workload || 'N/A',
            policy.Labels || 'N/A'
        ]);
        
        sections.push(
            createMultiColumnTable(
                ["Policy Name", "Status", "Workload", "Labels Applied"],
                policyRows
            ),
            new Paragraph({ text: "" })
        );
    }
}

// Conditional Access Section
if (jsonData.ConditionalAccess && !jsonData.ConditionalAccess.Error) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("Conditional Access Policies")],
            pageBreakBefore: true
        }),
        new Paragraph({ text: "" })
    );
    
    const caSummary = new Table({
        columnWidths: [4680, 4680],
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        rows: [
            createTableRow("Metric", "Value", true),
            createTableRow("Total Policies", jsonData.ConditionalAccess.TotalPolicies || 0),
            createTableRow("Enabled Policies", jsonData.ConditionalAccess.EnabledPolicies || 0),
            createTableRow("Report-Only Policies", jsonData.ConditionalAccess.ReportOnlyPolicies || 0)
        ]
    });
    sections.push(caSummary, new Paragraph({ text: "" }));
    
    // Detailed policy information
    if (jsonData.ConditionalAccess.Policies && jsonData.ConditionalAccess.Policies.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("Policy Details")]
            }),
            new Paragraph({ text: "" })
        );
        
        const policyRows = jsonData.ConditionalAccess.Policies.map(policy => [
            policy.DisplayName || 'N/A',
            policy.State || 'N/A',
            policy.GrantControls || 'None',
            policy.Conditions?.Locations || 'Any',
            policy.Conditions?.Platforms || 'Any'
        ]);
        
        sections.push(
            createMultiColumnTable(
                ["Policy Name", "State", "Grant Controls", "Locations", "Platforms"],
                policyRows
            ),
            new Paragraph({ text: "" })
        );
        
        sections.push(
            new Paragraph({ 
                children: [new TextRun({ 
                    text: "Note: Conditional Access policies may affect how users access Purview-protected content.",
                    italics: true
                })]
            }),
            new Paragraph({ text: "" })
        );
    }
} else if (jsonData.ConditionalAccess && jsonData.ConditionalAccess.Error) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("Conditional Access Policies")]
        }),
        new Paragraph({ 
            children: [new TextRun({ text: "Data not available: " + jsonData.ConditionalAccess.Error, italics: true })]
        }),
        new Paragraph({ text: "" })
    );
}

// Third-Party Apps Section
if (jsonData.ThirdPartyApps && !jsonData.ThirdPartyApps.Error) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("Third-Party Applications")],
            pageBreakBefore: true
        }),
        new Paragraph({ text: "" })
    );
    
    const appSummary = new Table({
        columnWidths: [4680, 4680],
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        rows: [
            createTableRow("Metric", "Value", true),
            createTableRow("Total Third-Party Apps", jsonData.ThirdPartyApps.TotalApps || 0)
        ]
    });
    sections.push(appSummary, new Paragraph({ text: "" }));
    
    // Detailed app list
    if (jsonData.ThirdPartyApps.Apps && jsonData.ThirdPartyApps.Apps.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("Connected Applications")]
            }),
            new Paragraph({ 
                children: [new TextRun(`Showing ${jsonData.ThirdPartyApps.Apps.length} of ${jsonData.ThirdPartyApps.TotalApps} total apps`)]
            }),
            new Paragraph({ text: "" })
        );
        
        const appRows = jsonData.ThirdPartyApps.Apps.map(app => [
            app.DisplayName || 'N/A',
            app.PublisherName || 'Unknown',
            app.SignInAudience || 'N/A',
            app.Homepage || 'N/A'
        ]);
        
        sections.push(
            createMultiColumnTable(
                ["Application Name", "Publisher", "Audience", "Homepage"],
                appRows
            ),
            new Paragraph({ text: "" })
        );
        
        sections.push(
            new Paragraph({ 
                children: [new TextRun({ 
                    text: "⚠️ Review: Each third-party app should be evaluated for data access permissions and compliance requirements.",
                    italics: true
                })]
            }),
            new Paragraph({ text: "" })
        );
    }
} else if (jsonData.ThirdPartyApps && jsonData.ThirdPartyApps.Error) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("Third-Party Applications")]
        }),
        new Paragraph({ 
            children: [new TextRun({ text: "Data not available: " + jsonData.ThirdPartyApps.Error, italics: true })]
        }),
        new Paragraph({ text: "" })
    );
}

// Data Repositories Section
if (jsonData.DataRepositories) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("Data Repository Mapping")],
            pageBreakBefore: true
        }),
        new Paragraph({ text: "" })
    );
    
    const repoSummary = new Table({
        columnWidths: [4680, 4680],
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        rows: [
            createTableRow("Repository Type", "Count", true),
            createTableRow("SharePoint Sites", (jsonData.DataRepositories.SharePointSites || []).length),
            createTableRow("OneDrive Sites", (jsonData.DataRepositories.OneDriveSites || []).length),
            createTableRow("Teams", (jsonData.DataRepositories.Teams || []).length),
            createTableRow("Exchange Mailboxes (Sample)", (jsonData.DataRepositories.ExchangeMailboxes || []).length)
        ]
    });
    sections.push(repoSummary, new Paragraph({ text: "" }));
    
    // SharePoint Sites Details
    if (jsonData.DataRepositories.SharePointSites && jsonData.DataRepositories.SharePointSites.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("SharePoint Sites")]
            }),
            new Paragraph({ text: "" })
        );
        
        const totalStorage = jsonData.DataRepositories.SharePointSites.reduce((sum, site) => sum + (site.StorageUsedGB || 0), 0);
        sections.push(
            new Paragraph({ 
                children: [new TextRun({ text: `Total Storage Used: ${totalStorage.toFixed(2)} GB`, bold: true })]
            }),
            new Paragraph({ text: "" })
        );
        
        const siteRows = jsonData.DataRepositories.SharePointSites.map(site => [
            site.Title || 'N/A',
            site.Url || 'N/A',
            (site.StorageUsedGB || 0).toFixed(2) + ' GB',
            (site.StorageQuotaGB || 0).toFixed(2) + ' GB',
            site.SharingCapability || 'N/A'
        ]);
        
        // Limit to top 50 sites by storage
        const topSites = siteRows.sort((a, b) => parseFloat(b[2]) - parseFloat(a[2])).slice(0, 50);
        
        sections.push(
            createMultiColumnTable(
                ["Site Title", "URL", "Used", "Quota", "Sharing"],
                topSites
            ),
            new Paragraph({ text: "" })
        );
        
        if (jsonData.DataRepositories.SharePointSites.length > 50) {
            sections.push(
                new Paragraph({ 
                    children: [new TextRun({ text: `Showing top 50 of ${jsonData.DataRepositories.SharePointSites.length} total sites by storage usage.`, italics: true })]
                }),
                new Paragraph({ text: "" })
            );
        }
    }
    
    // OneDrive Sites Details
    if (jsonData.DataRepositories.OneDriveSites && jsonData.DataRepositories.OneDriveSites.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("OneDrive Sites")]
            }),
            new Paragraph({ text: "" })
        );
        
        const totalODStorage = jsonData.DataRepositories.OneDriveSites.reduce((sum, site) => sum + (site.StorageUsedGB || 0), 0);
        sections.push(
            new Paragraph({ 
                children: [new TextRun({ text: `Total Storage Used: ${totalODStorage.toFixed(2)} GB`, bold: true })]
            }),
            new Paragraph({ text: "" })
        );
        
        const odRows = jsonData.DataRepositories.OneDriveSites.map(site => [
            site.Owner || 'N/A',
            site.Url || 'N/A',
            (site.StorageUsedGB || 0).toFixed(2) + ' GB'
        ]);
        
        // Limit to top 50 OneDrive sites by storage
        const topODSites = odRows.sort((a, b) => parseFloat(b[2]) - parseFloat(a[2])).slice(0, 50);
        
        sections.push(
            createMultiColumnTable(
                ["Owner", "URL", "Storage Used"],
                topODSites
            ),
            new Paragraph({ text: "" })
        );
        
        if (jsonData.DataRepositories.OneDriveSites.length > 50) {
            sections.push(
                new Paragraph({ 
                    children: [new TextRun({ text: `Showing top 50 of ${jsonData.DataRepositories.OneDriveSites.length} total OneDrive sites by storage usage.`, italics: true })]
                }),
                new Paragraph({ text: "" })
            );
        }
    }
    
    // Teams Details
    if (jsonData.DataRepositories.Teams && jsonData.DataRepositories.Teams.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("Microsoft Teams")]
            }),
            new Paragraph({ text: "" })
        );
        
        const teamRows = jsonData.DataRepositories.Teams.map(team => [
            team.DisplayName || 'N/A',
            team.Visibility || 'N/A',
            team.Archived ? 'Yes' : 'No',
            team.GroupId || 'N/A'
        ]);
        
        sections.push(
            createMultiColumnTable(
                ["Team Name", "Visibility", "Archived", "Group ID"],
                teamRows
            ),
            new Paragraph({ text: "" })
        );
    }
    
    // Exchange Mailboxes Sample
    if (jsonData.DataRepositories.ExchangeMailboxes && jsonData.DataRepositories.ExchangeMailboxes.length > 0) {
        sections.push(
            new Paragraph({ 
                heading: HeadingLevel.HEADING_2, 
                children: [new TextRun("Exchange Mailboxes (Sample)")]
            }),
            new Paragraph({ text: "" })
        );
        
        const mailboxRows = jsonData.DataRepositories.ExchangeMailboxes.slice(0, 25).map(mbx => [
            mbx.DisplayName || 'N/A',
            mbx.PrimarySmtpAddress || 'N/A',
            mbx.RecipientTypeDetails || 'N/A'
        ]);
        
        sections.push(
            createMultiColumnTable(
                ["Display Name", "Email Address", "Type"],
                mailboxRows
            ),
            new Paragraph({ text: "" })
        );
        
        sections.push(
            new Paragraph({ 
                children: [new TextRun({ text: `Showing sample of 25 mailboxes. Full inventory available in JSON data.`, italics: true })]
            }),
            new Paragraph({ text: "" })
        );
    }
}

// Information Protection & Retention Policies Section
if (jsonData.InformationProtection && !jsonData.InformationProtection.Error) {
    sections.push(
        new Paragraph({ 
            heading: HeadingLevel.HEADING_1, 
            children: [new TextRun("Information Protection & Retention Policies")],
            pageBreakBefore: true
        }),
        new Paragraph({ text: "" })
    );
    
    if (jsonData.InformationProtection.RetentionPolicies) {
        const retentionSummary = new Table({
            columnWidths: [4680, 4680],
            margins: { top: 100, bottom: 100, left: 180, right: 180 },
            rows: [
                createTableRow("Metric", "Value", true),
                createTableRow("Total Retention Policies", jsonData.InformationProtection.RetentionPolicies.TotalPolicies || 0),
                createTableRow("Enabled Retention Policies", jsonData.InformationProtection.RetentionPolicies.EnabledPolicies || 0)
            ]
        });
        sections.push(retentionSummary, new Paragraph({ text: "" }));
        
        // Detailed retention policy information
        if (jsonData.InformationProtection.RetentionPolicies.Policies && jsonData.InformationProtection.RetentionPolicies.Policies.length > 0) {
            sections.push(
                new Paragraph({ 
                    heading: HeadingLevel.HEADING_2, 
                    children: [new TextRun("Retention Policy Details")]
                }),
                new Paragraph({ text: "" })
            );
            
            const policyRows = jsonData.InformationProtection.RetentionPolicies.Policies.map(policy => [
                policy.Name || 'N/A',
                policy.Enabled ? 'Enabled' : 'Disabled',
                policy.Workload || 'N/A',
                policy.RuleCount || 0
            ]);
            
            sections.push(
                createMultiColumnTable(
                    ["Policy Name", "Status", "Workload", "Rules"],
                    policyRows
                ),
                new Paragraph({ text: "" })
            );
        }
    }
}

// Compliance Gaps Section
sections.push(
    new Paragraph({ 
        heading: HeadingLevel.HEADING_1, 
        children: [new TextRun("Compliance Gap Analysis")],
        pageBreakBefore: true
    }),
    new Paragraph({ text: "" }),
    new Paragraph({ 
        children: [new TextRun("Based on the assessment, the following areas require attention for Purview implementation:")]
    }),
    new Paragraph({ text: "" })
);

const gapRows = [];
if (!jsonData.DLPPolicies || jsonData.DLPPolicies.TotalPolicies === 0) {
    gapRows.push(["Data Loss Prevention", "No DLP policies configured", "High"]);
}
if (!jsonData.SensitivityLabels || jsonData.SensitivityLabels.TotalLabels === 0) {
    gapRows.push(["Sensitivity Labels", "No sensitivity labels deployed", "High"]);
}
if (!jsonData.InformationProtection || !jsonData.InformationProtection.RetentionPolicies || jsonData.InformationProtection.RetentionPolicies.TotalPolicies === 0) {
    gapRows.push(["Retention Policies", "No retention policies configured", "Medium"]);
}

if (gapRows.length > 0) {
    sections.push(
        createMultiColumnTable(
            ["Gap Area", "Description", "Priority"],
            gapRows
        )
    );
} else {
    sections.push(
        new Paragraph({ 
            children: [new TextRun("No critical compliance gaps identified.")]
        })
    );
}

// Create document
const doc = new Document({
    styles: {
        default: { 
            document: { run: { font: "Arial", size: 24 } } 
        },
        paragraphStyles: [
            { 
                id: "Heading1", 
                name: "Heading 1", 
                basedOn: "Normal", 
                next: "Normal", 
                quickFormat: true,
                run: { size: 32, bold: true, color: "2E5090", font: "Arial" },
                paragraph: { spacing: { before: 240, after: 180 }, outlineLevel: 0 }
            },
            { 
                id: "Heading2", 
                name: "Heading 2", 
                basedOn: "Normal", 
                next: "Normal", 
                quickFormat: true,
                run: { size: 28, bold: true, color: "2E5090", font: "Arial" },
                paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 1 }
            }
        ]
    },
    sections: [{
        properties: { 
            page: { 
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
            }
        },
        children: sections
    }]
});

// Generate Word file
const outputPath = process.argv[3];
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(outputPath, buffer);
    console.log(`Report generated: ${outputPath}`);
}).catch(err => {
    console.error('Error generating document:', err);
    process.exit(1);
});
'@

    $reportPath = $JsonPath -replace '\.json$', '.docx'
    
    try {
        # Check if Node.js is installed
        $nodeVersion = node --version 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Node.js is not installed. Please install Node.js from https://nodejs.org/" -Level Error
            throw "Node.js is required to generate the Word document"
        }
        
        Write-Log "Node.js version: $nodeVersion" -Level Info
        
        # Create a temporary directory for the Node.js project
        $tempProjectDir = Join-Path $env:TEMP "purview-report-$(Get-Date -Format 'yyyyMMddHHmmss')"
        New-Item -ItemType Directory -Path $tempProjectDir -Force | Out-Null
        
        # Create package.json
        $packageJson = @{
            name = "purview-report-generator"
            version = "1.0.0"
            dependencies = @{
                docx = "^8.5.0"
            }
        } | ConvertTo-Json
        
        $packageJson | Out-File -FilePath (Join-Path $tempProjectDir "package.json") -Encoding UTF8
        
        # Save the node script
        $scriptPath = Join-Path $tempProjectDir "generate-report.js"
        $nodeScript | Out-File -FilePath $scriptPath -Encoding UTF8
        
        # Install docx package locally
        Write-Log "Installing docx package locally..." -Level Info
        Push-Location $tempProjectDir
        try {
            npm install 2>&1 | Out-Null
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to install docx package"
            }
            Write-Log "Docx package installed successfully" -Level Success
        } finally {
            Pop-Location
        }
        
        # Run the node script
        Write-Log "Generating Word document..." -Level Info
        Push-Location $tempProjectDir
        try {
            node generate-report.js $JsonPath $reportPath 2>&1 | ForEach-Object {
                if ($_ -match "Report generated:") {
                    Write-Log $_ -Level Success
                } elseif ($_ -match "Error") {
                    Write-Log $_ -Level Error
                } else {
                    Write-Verbose $_
                }
            }
            
            if ($LASTEXITCODE -ne 0) {
                throw "Node.js script failed with exit code $LASTEXITCODE"
            }
        } finally {
            Pop-Location
        }
        
        # Clean up temp directory
        Remove-Item $tempProjectDir -Recurse -Force -ErrorAction SilentlyContinue
        
        if (Test-Path $reportPath) {
            Write-Log "Word document report generated successfully: $reportPath" -Level Success
            return $reportPath
        } else {
            throw "Word document was not created"
        }
        
    } catch {
        Write-Log "Error generating Word document: $_" -Level Error
        Write-Log "JSON data is still available at: $JsonPath" -Level Info
        return $JsonPath
    }
}
function Disconnect-M365Services {
    Write-Log "Disconnecting from Microsoft 365 services..." -Level Info
    
    try {
        if ($Script:ConnectedServices.Exchange) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
        if ($Script:ConnectedServices.Graph) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
        if ($Script:ConnectedServices.SharePoint) {
            Disconnect-SPOService -ErrorAction SilentlyContinue
        }
        if ($Script:ConnectedServices.Teams) {
            Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
        }
        Write-Log "Disconnected from all services" -Level Success
    } catch {
        Write-Log "Error during disconnection: $_" -Level Warning
    }
}

#endregion

#region Main Execution

try {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Microsoft Purview Due Diligence Assessment" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Check PowerShell version and provide recommendations
    Test-PowerShellVersion
    
    # Check for administrator privileges
    if (-not (Test-Administrator)) {
        Write-Log "This script requires administrator privileges for module installation." -Level Warning
        Write-Log "Attempting to elevate..." -Level Info
        Start-ElevatedSession
    }
    
    # Install required modules
    if (-not $SkipModuleCheck) {
        Write-Log "Checking and installing required PowerShell modules..." -Level Info
        
        # Install non-Graph modules normally
        Install-RequiredModule -ModuleName 'ExchangeOnlineManagement' -MinimumVersion '3.0.0'
        Install-RequiredModule -ModuleName 'Microsoft.Online.SharePoint.PowerShell' -MinimumVersion '16.0.0'
        Install-RequiredModule -ModuleName 'MicrosoftTeams' -MinimumVersion '5.0.0'
        
        # Install Microsoft.Graph module but don't import it (to avoid function limit)
        # Instead, we'll import only the sub-modules we need
        Write-Log "Installing Microsoft.Graph module (will import specific sub-modules only)..." -Level Info
        Install-RequiredModule -ModuleName 'Microsoft.Graph' -MinimumVersion '2.0.0' -SkipImport
        
        # Now install and import only the specific Graph sub-modules we need
        Install-GraphSubModules
    }
    
    # Connect to M365 services
    Connect-M365Services
    
    Write-Host ""
    Write-Log "Starting data collection..." -Level Info
    Write-Host ""
    
    # Collect all assessment data
    Get-UserLicensing
    Get-M365WorkloadStatus
    Get-SharingSettings
    Get-DLPPolicies
    Get-SensitivityLabels
    Get-InformationProtectionPolicies
    Get-ConditionalAccessPolicies
    Get-ThirdPartyApps
    Get-GuestUserPolicies
    Get-DataRepositories
    
    # Export data to JSON
    $jsonPath = Export-AssessmentData
    
    # Generate Word document report
    Write-Host ""
    $reportPath = New-AssessmentReport -JsonPath $jsonPath
    
    # Summary
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Assessment Complete!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Reports generated:" -ForegroundColor Cyan
    Write-Host "  - JSON Data: $jsonPath" -ForegroundColor White
    if ($reportPath -like "*.docx") {
        Write-Host "  - Word Report: $reportPath" -ForegroundColor White
    }
    Write-Host ""
    
    if ($Script:ErrorLog.Count -gt 0) {
        Write-Host "Errors encountered during assessment: $($Script:ErrorLog.Count)" -ForegroundColor Yellow
        Write-Host "Review the console output above for details." -ForegroundColor Yellow
    }
    
} catch {
    Write-Log "Fatal error during assessment: $_" -Level Error
    Write-Log $_.ScriptStackTrace -Level Error
    throw
} finally {
    # Disconnect from services
    Disconnect-M365Services
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

#endregion

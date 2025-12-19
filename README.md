# Microsoft 365 Purview Implementation Assessment Tool

## Overview

Comprehensive PowerShell script for conducting due diligence assessments of Microsoft 365 environments to support Microsoft Purview implementation planning. Automatically collects configuration data across all M365 workloads and generates a professional Word document report with actionable insights.

## What It Does

**Automated Data Collection:**
- User licensing and SKU assignments (with Purview license gap analysis)
- M365 workload status and usage (Exchange, SharePoint, OneDrive, Teams)
- Data Loss Prevention (DLP) policies and rules
- Sensitivity labels and label policies
- Information Protection and retention policies
- Conditional Access policies
- Third-party application integrations
- Guest user and external collaboration settings
- Complete data repository mapping with storage metrics
- External sharing configurations across all workloads

**Professional Reporting:**
- Generates detailed Word document (.docx) with tables, metrics, and analysis
- Exports structured JSON data for further analysis
- Includes compliance gap analysis with prioritized recommendations
- Color-coded warnings for critical findings
- Storage calculations and usage percentages
- Detailed inventories (not just counts)

## Key Features

### Robust Authentication
- **MFA Support**: Works seamlessly with Multi-Factor Authentication
- **Multiple Auth Methods**: Browser authentication with automatic fallback to device code flow
- **Graceful Degradation**: Continues assessment even if some services fail to connect
- **Interactive SharePoint Prompting**: Asks for correct tenant name if auto-detection fails

### Intelligent Module Management
- **Auto-Installation**: Installs required PowerShell modules automatically
- **PowerShell 5.1 Compatibility**: Works around function limit issues with selective Graph sub-module imports
- **Local Package Installation**: Uses temporary npm project for Word generation (no global permissions needed)

### User-Friendly Experience
- **Parameter Preservation**: Maintains parameters when elevating to administrator
- **Real-Time Progress**: Shows detailed logging throughout execution
- **Connection Summary**: Displays which services connected successfully
- **Error Recovery**: Up to 3 attempts for SharePoint connection with user prompting
- **Flexible Execution**: Can skip SharePoint or provide custom tenant names

### Production-Ready Design
- **UTF-8 BOM Handling**: Properly formats JSON for Node.js parsing
- **No Data Loss**: Exports JSON even if Word generation fails
- **Session Cleanup**: Properly disconnects from all services
- **Comprehensive Error Logging**: Tracks all errors with timestamps

## What Problems It Solves

1. **Manual Discovery Pain**: Eliminates hours of manual portal navigation and data collection
2. **Incomplete Assessments**: Ensures comprehensive data gathering across all M365 services
3. **License Planning Gaps**: Calculates exactly how many users need Purview licensing
4. **Data Location Mystery**: Maps all SharePoint sites, OneDrive, Teams, and mailboxes with storage metrics
5. **Compliance Blind Spots**: Identifies missing DLP policies, unused sensitivity labels, and configuration gaps
6. **Disconnected Data**: Consolidates scattered information into a single actionable report
7. **SharePoint Tenant Confusion**: Intelligently handles mismatched Exchange/SharePoint domain names
8. **Authentication Complexity**: Manages MFA and multiple authentication methods automatically

## Technical Details

**Prerequisites:**
- Windows PowerShell 5.1+ (PowerShell 7+ recommended)
- Node.js 14+ (for Word document generation)
- Global Reader or Global Administrator role
- Internet connectivity

**PowerShell Modules** (auto-installed):
- ExchangeOnlineManagement
- Microsoft.Graph (Authentication, Users, Groups, Identity, Applications, Reports)
- Microsoft.Online.SharePoint.PowerShell
- MicrosoftTeams

**Output Files:**
- `M365_Purview_Assessment_[timestamp].json` - Complete raw data
- `M365_Purview_Assessment_[timestamp].docx` - Professional formatted report

## Usage

**Basic:**
```powershell
.\M365-Purview-Assessment.ps1
```

**With SharePoint tenant override:**
```powershell
.\M365-Purview-Assessment.ps1 -SharePointTenantName 'your-tenant'
```

**With all parameters:**
```powershell
.\M365-Purview-Assessment.ps1 -SharePointTenantName 'your-tenant' -OutputPath "C:\Reports\Assessment.json" -SkipModuleCheck
```

## Report Contents

**Executive Summary** | **User Licensing** (with Purview gap analysis) | **Workload Status** (Exchange, SharePoint, OneDrive, Teams) | **Sharing Settings** | **DLP Policies** (with rule details) | **Sensitivity Labels** (with descriptions & encryption status) | **Retention Policies** | **Conditional Access** | **Third-Party Apps** (with publishers) | **Data Repository Mapping** (sites, storage, sharing) | **Compliance Gap Analysis**

## Evolution Through Development

Built through iterative refinement addressing real-world challenges:
- Fixed PowerShell 5.1 function limit with Microsoft.Graph module
- Implemented graceful service degradation for failed connections
- Added device code authentication fallback for restrictive environments
- Removed invalid Graph API scopes for tenant compatibility
- Fixed UTF-8 BOM encoding issues for Node.js JSON parsing
- Implemented local npm package installation for reliability
- Added SharePoint tenant name parameter and auto-prompting
- Preserved parameters through administrative elevation
- Enhanced reporting from summary counts to complete detailed inventories

## Version

Current: **1.4.2** (December 2024)

---

**Author:** Vivek | **Purpose:** Microsoft Purview Implementation Due Diligence | **License:** [Your Choice]

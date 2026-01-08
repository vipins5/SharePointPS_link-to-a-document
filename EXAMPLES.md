# Example Usage Guide

This file provides practical examples for using the Export-SPLinkToDocument.ps1 script.

## Step 1: Install Prerequisites

First, ensure you have the PnP.PowerShell module installed:

```powershell
# Check if module is installed
Get-Module -ListAvailable -Name PnP.PowerShell

# If not installed, install it
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
```

## Step 2: Prepare Your Environment

1. Note your SharePoint site URL (e.g., `https://contoso.sharepoint.com/sites/mysite`)
2. Note the document library name (e.g., `Documents`, `Shared Documents`)
3. Ensure you have at least Read permissions on the library

## Step 3: Run the Script

### Scenario 1: Quick Export with Default Settings
Export all "Link to a Document" items to a file named `LinkToDocument_Export.csv` in the current directory:

```powershell
.\Export-SPLinkToDocument.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -LibraryName "Documents"
```

### Scenario 2: Export to a Specific Location
Export to a specific file path:

```powershell
.\Export-SPLinkToDocument.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/projectsite" `
    -LibraryName "Shared Documents" `
    -OutputPath "C:\Reports\SharePoint\links_$(Get-Date -Format 'yyyyMMdd').csv"
```

### Scenario 3: Export from Multiple Libraries
If you need to export from multiple libraries, run the script multiple times:

```powershell
# Export from Documents library
.\Export-SPLinkToDocument.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/mysite" `
    -LibraryName "Documents" `
    -OutputPath ".\exports\documents_links.csv"

# Export from another library
.\Export-SPLinkToDocument.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/mysite" `
    -LibraryName "Project Files" `
    -OutputPath ".\exports\project_links.csv"
```

## Step 4: Review the Output

Open the generated CSV file to review the exported data. The file will contain:
- ID: Unique item identifier
- Title: Link title
- URL: The target URL
- URLDescription: Description of the link
- ContentType: Content type name
- Created/Modified dates
- Created/Modified by user names
- FilePath: SharePoint path

## Step 5: Analyze the Data

You can import the CSV into Excel or PowerShell for further analysis:

```powershell
# Import and analyze in PowerShell
$links = Import-Csv -Path "LinkToDocument_Export.csv"

# Count total links
$links.Count

# Group by created user
$links | Group-Object -Property CreatedBy | Select-Object Name, Count

# Find recently modified links (last 30 days)
$thirtyDaysAgo = (Get-Date).AddDays(-30)
$links | Where-Object { [DateTime]$_.Modified -gt $thirtyDaysAgo } | Format-Table Title, Modified, ModifiedBy

# Export broken links to a separate file (if URL is empty)
$links | Where-Object { -not $_.URL } | Export-Csv -Path "broken_links.csv" -NoTypeInformation
```

## Troubleshooting Examples

### If you get "Access Denied"
Verify your permissions:
```powershell
# Connect manually to test access
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/mysite" -Interactive

# Try to list the library
Get-PnPList -Identity "Documents"

# Disconnect
Disconnect-PnPOnline
```

### If no items are found
The script will show available content types. You can manually check:
```powershell
# Connect to the site
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/mysite" -Interactive

# Get all items and check content types
$items = Get-PnPListItem -List "Documents"
$items | Select-Object Id, @{Name="ContentType";Expression={$_.ContentType.Name}} | Group-Object ContentType

# Disconnect
Disconnect-PnPOnline
```

## Advanced: Scheduled Export

You can create a scheduled task to run this export regularly:

```powershell
# Create a scheduled task (requires admin privileges)
$action = New-ScheduledTaskAction -Execute "pwsh.exe" `
    -Argument "-File C:\Scripts\Export-SPLinkToDocument.ps1 -SiteUrl 'https://contoso.sharepoint.com/sites/mysite' -LibraryName 'Documents' -OutputPath 'C:\Reports\links.csv'"

$trigger = New-ScheduledTaskTrigger -Daily -At 9am

Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "SharePoint Link Export" -Description "Daily export of SharePoint link items"
```

## Notes

- The script uses interactive authentication, which means it will open a browser window for sign-in
- For automated/scheduled scenarios, you may need to use certificate-based authentication or app-only authentication
- Large libraries may take several minutes to process
- The script respects SharePoint throttling limits with pagination

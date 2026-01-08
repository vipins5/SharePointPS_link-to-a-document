<#
.SYNOPSIS
    Exports "Link to a Document" items from a SharePoint Online document library to a CSV file.

.DESCRIPTION
    This script connects to a SharePoint Online site and exports all items with the 
    "Link to a Document" content type from a specified document library to a CSV file.

.PARAMETER SiteUrl
    The URL of the SharePoint Online site (e.g., https://contoso.sharepoint.com/sites/sitename)

.PARAMETER LibraryName
    The name of the document library to export from

.PARAMETER OutputPath
    The path where the CSV file will be saved (default: LinkToDocument_Export.csv in current directory)

.EXAMPLE
    .\Export-SPLinkToDocument.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -LibraryName "Documents" -OutputPath "C:\Export\links.csv"
    
    Exports all "Link to a Document" items from the Documents library to the specified CSV file.

.EXAMPLE
    .\Export-SPLinkToDocument.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -LibraryName "Shared Documents"
    
    Exports all "Link to a Document" items using the default output filename.

.NOTES
    Requires PnP.PowerShell module to be installed.
    Install with: Install-Module -Name PnP.PowerShell -Scope CurrentUser
    
    Requires appropriate permissions to access the SharePoint site and library.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="SharePoint Online site URL")]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true, HelpMessage="Document library name")]
    [ValidateNotNullOrEmpty()]
    [string]$LibraryName,
    
    [Parameter(Mandatory=$false, HelpMessage="Output CSV file path")]
    [string]$OutputPath = "LinkToDocument_Export.csv"
)

# Function to check if PnP.PowerShell module is installed
function Test-PnPModule {
    $module = Get-Module -ListAvailable -Name PnP.PowerShell
    if (-not $module) {
        Write-Error "PnP.PowerShell module is not installed. Please install it using: Install-Module -Name PnP.PowerShell -Scope CurrentUser"
        return $false
    }
    return $true
}

# Function to export Link to Document items
function Export-LinkToDocumentItems {
    param(
        [string]$SiteUrl,
        [string]$LibraryName,
        [string]$OutputPath
    )
    
    try {
        Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Cyan
        
        # Connect to SharePoint Online site using interactive login
        Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
        
        Write-Host "Successfully connected to SharePoint site" -ForegroundColor Green
        Write-Host "Retrieving items from library: $LibraryName" -ForegroundColor Cyan
        
        # Get all items from the library
        # Filter by ContentType name "Link to a Document" or by content type ID (0x0101...)
        $allItems = Get-PnPListItem -List $LibraryName -PageSize 2000
        
        if ($allItems.Count -eq 0) {
            Write-Warning "No items found in the library '$LibraryName'"
            return
        }
        
        Write-Host "Total items retrieved: $($allItems.Count)" -ForegroundColor Yellow
        
        # Filter for "Link to a Document" content type
        # The content type name is typically "Link to a Document"
        # Content Type ID starts with 0x0105 for Link to a Document
        $linkItems = $allItems | Where-Object { 
            $_.ContentType.Name -eq "Link to a Document" -or
            $_.ContentType.Id -like "0x0105*"
        }
        
        if ($linkItems.Count -eq 0) {
            Write-Warning "No 'Link to a Document' items found in the library"
            Write-Host "Available content types found:" -ForegroundColor Yellow
            $allItems | Select-Object -ExpandProperty ContentType -Unique | ForEach-Object {
                Write-Host "  - $($_.Name) (ID: $($_.Id))" -ForegroundColor Gray
            }
            return
        }
        
        Write-Host "Found $($linkItems.Count) 'Link to a Document' items" -ForegroundColor Green
        
        # Create export data using ArrayList for better performance
        $exportData = [System.Collections.ArrayList]::new()
        
        foreach ($item in $linkItems) {
            # Safely get lookup values with null checks
            $createdBy = if ($item["Author"] -and $item["Author"].LookupValue) { 
                $item["Author"].LookupValue 
            } else { 
                "" 
            }
            
            $modifiedBy = if ($item["Editor"] -and $item["Editor"].LookupValue) { 
                $item["Editor"].LookupValue 
            } else { 
                "" 
            }
            
            $exportItem = [PSCustomObject]@{
                ID = $item.Id
                Title = $item["Title"]
                Name = $item["FileLeafRef"]
                URL = $item["URL"]
                URLDescription = $item["URLDescription"]
                ContentType = $item.ContentType.Name
                Created = $item["Created"]
                CreatedBy = $createdBy
                Modified = $item["Modified"]
                ModifiedBy = $modifiedBy
                FilePath = $item["FileRef"]
            }
            
            $null = $exportData.Add($exportItem)
        }
        
        # Export to CSV
        $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        
        Write-Host "`nExport completed successfully!" -ForegroundColor Green
        Write-Host "Exported $($exportData.Count) items to: $OutputPath" -ForegroundColor Green
        
        # Disconnect from SharePoint
        Disconnect-PnPOnline
        
    }
    catch {
        Write-Error "An error occurred: $_"
        Write-Error $_.Exception.Message
        
        # Try to disconnect if connected
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore disconnect errors
        }
        
        throw
    }
}

# Main execution
try {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Export SharePoint Link to Document Items" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Check if PnP module is available
    if (-not (Test-PnPModule)) {
        exit 1
    }
    
    # Validate output path directory exists
    $outputDir = Split-Path -Path $OutputPath -Parent
    if ($outputDir -and -not (Test-Path $outputDir)) {
        Write-Error "Output directory does not exist: $outputDir"
        exit 1
    }
    
    # Execute export
    Export-LinkToDocumentItems -SiteUrl $SiteUrl -LibraryName $LibraryName -OutputPath $OutputPath
    
    Write-Host "`nScript execution completed." -ForegroundColor Cyan
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}

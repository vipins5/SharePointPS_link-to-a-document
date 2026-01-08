# SharePointPS_link-to-a-document

Exports "Link to a Document" items from a SharePoint Online document library to a CSV file.

## Description

This PowerShell script connects to a SharePoint Online site and exports all items with the "Link to a Document" content type from a specified document library to a CSV file. This is useful for:
- Auditing document links in SharePoint libraries
- Migrating or documenting link references
- Creating reports of linked documents
- Backup and inventory of document references

## Prerequisites

- **PowerShell 5.1 or later** (PowerShell 7+ recommended)
- **PnP.PowerShell module** - Install using:
  ```powershell
  Install-Module -Name PnP.PowerShell -Scope CurrentUser
  ```
- **SharePoint Online access** with appropriate permissions to read the document library

## Usage

### Basic Syntax

```powershell
.\Export-SPLinkToDocument.ps1 -SiteUrl <SharePoint Site URL> -LibraryName <Library Name> [-OutputPath <CSV File Path>]
```

### Parameters

- **SiteUrl** (Required): The full URL of the SharePoint Online site
  - Example: `https://contoso.sharepoint.com/sites/mysite`
  
- **LibraryName** (Required): The name of the document library to export from
  - Example: `Documents` or `Shared Documents`
  
- **OutputPath** (Optional): The path where the CSV file will be saved
  - Default: `LinkToDocument_Export.csv` in the current directory
  - Example: `C:\Export\links.csv`

### Examples

#### Example 1: Export with default output file
```powershell
.\Export-SPLinkToDocument.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -LibraryName "Documents"
```

#### Example 2: Export with custom output path
```powershell
.\Export-SPLinkToDocument.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -LibraryName "Documents" -OutputPath "C:\Export\links.csv"
```

#### Example 3: Export from a specific library
```powershell
.\Export-SPLinkToDocument.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/projectsite" -LibraryName "Shared Documents" -OutputPath ".\project_links.csv"
```

## Output

The script generates a CSV file with the following columns:

- **ID**: The unique identifier of the item
- **Title**: The title of the link item
- **Name**: The file/item name
- **URL**: The URL the link points to
- **URLDescription**: Description of the URL
- **ContentType**: The content type name (e.g., "Link to a Document")
- **Created**: Date and time when the item was created
- **CreatedBy**: User who created the item
- **Modified**: Date and time when the item was last modified
- **ModifiedBy**: User who last modified the item
- **FilePath**: The SharePoint path to the item

## Authentication

The script uses interactive authentication with modern authentication support. When you run the script, a browser window will open asking you to sign in to your Microsoft 365 account.

## Error Handling

The script includes comprehensive error handling and will:
- Verify that the PnP.PowerShell module is installed
- Validate the output directory exists
- Check if items exist in the library
- Display available content types if no "Link to a Document" items are found
- Provide detailed error messages if something goes wrong

## Notes

- The script uses pagination (2000 items per page) to handle large libraries efficiently
- The script filters items by content type name containing "Link" or matching "Link to a Document"
- All dates are exported in the format they are stored in SharePoint
- The CSV file is encoded in UTF-8 format

## Troubleshooting

### Module Not Found
If you get an error about PnP.PowerShell not being installed:
```powershell
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
```

### Access Denied
Ensure you have at least Read permissions on the SharePoint site and document library.

### No Items Found
The script will display all available content types if no "Link to a Document" items are found. This helps you verify the correct content type name in your environment.

## License

This project is provided as-is for use with SharePoint Online environments.

## Contributing

Feel free to submit issues or pull requests to improve the script.

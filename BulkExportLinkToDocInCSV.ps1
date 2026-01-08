<#
.SYNOPSIS
Exports “Link to a Document” items from a SharePoint Online document library to a CSV.

.DESCRIPTION
Connects to a SharePoint Online site using PnP.PowerShell interactive authentication, retrieves up to 5,000 items from
a specified library, filters to only “.aspx” link items (commonly used by the “Link to a Document” content type),
extracts the target URL from the library’s URL field, and exports the results (Name, Title, URL) to a UTF-8 CSV file.

NOTES
- Requires PnP.PowerShell.
- The script assumes the link target is stored in the field with internal name "URL".
- The "URL" field can be returned either as a string ("https://... , description") or as a FieldUrlValue object;
  the script handles both and extracts the raw URL.
- If your library contains more than 5,000 items, consider paging with CAML queries or using -PageSize with additional logic.
#>

# -----------------------
# CONFIG (PII removed)
# -----------------------
$siteUrl  = "https://<TENANT>.sharepoint.com/sites/<SITE>"
$listName = "<DOCUMENT_LIBRARY_NAME>"
$clientId = "<ENTRA_APP_CLIENT_ID_GUID>"
$outCsv   = "C:\Path\To\Export_Links.csv"

# Connect to SharePoint Online (interactive sign-in)
Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId

# Retrieve items and pre-load the fields we need for export
$items = Get-PnPListItem -List $listName -PageSize 5000 -Fields "FileLeafRef","Title","URL","ContentType"

# Transform list items into export rows
$rows = foreach ($it in $items) {
    $name = $it["FileLeafRef"]
    if (-not $name) { continue }

    # Export only .aspx link items (adjust filter if your link items use a different extension)
    if ($name -notlike "*.aspx") { continue }

    # Extract the URL from the "URL" field (can be string or FieldUrlValue)
    $urlVal = $it["URL"]
    $url = $null

    if ($urlVal -is [string]) {
        # Example string format: "https://example.com, Display Text"
        $url = ($urlVal -split ",")[0].Trim()
    }
    elseif ($urlVal -and $urlVal.Url) {
        # FieldUrlValue object
        $url = $urlVal.Url
    }

    # Output object for CSV export
    [pscustomobject]@{
        Name  = $name
        Title = $it["Title"]
        URL   = $url
    }
}

# Export to CSV (UTF-8)
$rows | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8

Write-Host "Exported: $($rows.Count) items to $outCsv"

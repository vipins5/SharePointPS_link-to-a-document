<#
.SYNOPSIS
Bulk-creates SharePoint Online “Link to a Document” items in a document library from a CSV input.

.DESCRIPTION
This script creates SharePoint Online link items (commonly represented as .aspx files for the “Link to a Document”
content type) by performing a server-side copy of a known-good template link item. A server-side copy is used to
preserve SharePoint’s internal “safe to execute” classification (e.g., NoExecute=0), which avoids the behavior where
programmatically uploaded .aspx files are forced to download instead of opening as links.

The script:
- Connects to SharePoint Online using PnP.PowerShell interactive authentication.
- Validates that a template link file exists (created manually once via “New -> Link”).
- Ensures a template folder exists (e.g., "_Templates") and moves the template file into it (server-side copy + delete)
  as a “hide” mechanism (best-effort).
- Reads a CSV containing Title + URL (and optional Description) and creates new link items by copying the template.
- Updates metadata on each new item (Title, FileLeafRef, URL field, and optional supporting fields).
- Writes execution transcript and error logs to timestamped log files.

CSV FORMAT
Required headers:
  Title,URL
Optional header:
  Description  (used as URL display text; defaults to Title)

PREREQUISITES
- Install-Module PnP.PowerShell -Scope CurrentUser
- The template link file must exist in the target library at least once (create manually via “New -> Link”).
#>

# -----------------------
# CONFIG (PII removed)
# -----------------------
$siteUrl   = "https://<TENANT>.sharepoint.com/sites/<SITE>"
$library   = "<DOCUMENT_LIBRARY_NAME>"
$csvPath   = "C:\Path\To\links.csv"
$ctName    = "Link to a Document"
$clientId  = "<ENTRA_APP_CLIENT_ID_GUID>"

# Template link created manually via SharePoint UI: "New -> Link"
# NOTE: The template must exist at least once; this script copies it server-side.
$templateFileName = "test.aspx"

# Folder used to keep the template out of typical end-user navigation
$templateFolderName = "_Templates"

# Log output directory
$logRoot = "C:\Path\To\Logs"

# -----------------------
# LOGGING HELPERS
# -----------------------
$ts = Get-Date -Format "yyyyMMdd-HHmmss"
New-Item -ItemType Directory -Force -Path $logRoot | Out-Null

$errLog = Join-Path $logRoot "BulkLinkToDoc-Errors-$ts.log"
$runLog = Join-Path $logRoot "BulkLinkToDoc-Run-$ts.log"

# Capture full console output for troubleshooting/auditing
Start-Transcript -Path $runLog -Append | Out-Null

function Write-ErrLog {
    param(
        [string]$Message,
        [object]$Exception = $null
    )

    $line = "[{0}] {1}" -f (Get-Date -Format "s"), $Message
    Add-Content -Path $errLog -Value $line

    if ($Exception) {
        Add-Content -Path $errLog -Value ("    Exception: " + $Exception.ToString())
    }
}

function New-SafeBaseName([string]$name) {
    # Produces a filesystem-safe base name for an .aspx file
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) { $name = $name.Replace($c, "_") }
    $name = $name.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "Link" }
    return $name
}

function Get-UniqueAspxName([string]$serverRelFolder, [string]$baseName) {
    # Generates a unique .aspx filename by checking for collisions in the library folder
    $n = 0
    while ($true) {
        $candidate = if ($n -eq 0) { "$baseName.aspx" } else { "$baseName-$n.aspx" }
        $candidateUrl = "$serverRelFolder/$candidate"
        $exists = Get-PnPFile -Url $candidateUrl -ErrorAction SilentlyContinue
        if (-not $exists) { return $candidate }
        $n++
    }
}

try {
    # -----------------------
    # CONNECT
    # -----------------------
    Write-Host "Connecting to $siteUrl ..."
    Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId

    # Resolve target library and its root folder (server-relative)
    $list = Get-PnPList -Identity $library
    $rootFolder = $list.RootFolder.ServerRelativeUrl.TrimEnd("/")

    # Validate the expected URL field exists (internal name "URL" in this pattern)
    $urlField = Get-PnPField -List $list | Where-Object { $_.InternalName -eq "URL" } | Select-Object -First 1
    if (-not $urlField) {
        throw "URL field (InternalName 'URL') not found in library '$library'."
    }

    # -----------------------
    # ENSURE TEMPLATE FOLDER EXISTS
    # -----------------------
    $templateFolderUrl = "$rootFolder/$templateFolderName"

    if (-not (Get-PnPFolder -Url $templateFolderUrl -ErrorAction SilentlyContinue)) {
        Write-Host "Creating template folder: $templateFolderUrl"
        Add-PnPFolder -Name $templateFolderName -Folder $rootFolder | Out-Null
    }

    # -----------------------
    # LOCATE TEMPLATE FILE
    # -----------------------
    $templateAtRootUrl   = "$rootFolder/$templateFileName"
    $templateAtHiddenUrl = "$templateFolderUrl/$templateFileName"

    $templateSourceUrl = $null

    # Prefer template in the designated template folder
    if (Get-PnPFile -Url $templateAtHiddenUrl -ErrorAction SilentlyContinue) {
        $templateSourceUrl = $templateAtHiddenUrl
        Write-Host "Template found in template folder: $templateSourceUrl"
    }
    elseif (Get-PnPFile -Url $templateAtRootUrl -ErrorAction SilentlyContinue) {
        $templateSourceUrl = $templateAtRootUrl
        Write-Host "Template found at root: $templateSourceUrl"
    }
    else {
        throw "Template '$templateFileName' not found. Create it manually (New -> Link) in '$library' first."
    }

    # -----------------------
    # MOVE TEMPLATE INTO TEMPLATE FOLDER (server-side copy + delete)
    # -----------------------
    # Rationale: Avoid end-user interaction with the template; preserve its “safe to execute” state by copying server-side.
    if ($templateSourceUrl -eq $templateAtRootUrl) {
        Write-Host "Moving template into template folder..."
        Copy-PnPFile -SourceUrl $templateAtRootUrl -TargetUrl $templateAtHiddenUrl -OverwriteIfAlreadyExists:$true -Force
        Remove-PnPFile -ServerRelativeUrl $templateAtRootUrl -Force -Recycle
        $templateSourceUrl = $templateAtHiddenUrl
    }

    # Best-effort: mark template item hidden (note: some tenants/views ignore this for document libraries)
    try {
        $tmplItem = Get-PnPFile -Url $templateSourceUrl -AsListItem -ErrorAction Stop
        Set-PnPListItem -List $list -Identity $tmplItem.Id -Values @{ "Hidden" = $true } -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        Write-ErrLog "Could not set Hidden=true on template item (non-fatal)." $_.Exception
    }

    # -----------------------
    # IMPORT CSV
    # -----------------------
    if (-not (Test-Path $csvPath)) {
        throw "CSV not found: $csvPath"
    }

    $rows = Import-Csv $csvPath
    if (-not $rows -or $rows.Count -eq 0) {
        throw "CSV is empty or unreadable: $csvPath"
    }

    Write-Host "Rows in CSV: $($rows.Count)"
    Write-Host "Template used: $templateSourceUrl"

    # -----------------------
    # CREATE LINKS
    # -----------------------
    foreach ($r in $rows) {
        $title = $null
        $url = $null

        try {
            $title = ($r.Title | ForEach-Object { "$_".Trim() })
            $url   = ($r.URL   | ForEach-Object { "$_".Trim() })

            if ([string]::IsNullOrWhiteSpace($title) -or [string]::IsNullOrWhiteSpace($url)) {
                Write-ErrLog "Skipping row (missing Title or URL). Row=$($r | ConvertTo-Json -Compress)"
                continue
            }

            # Optional display text for the URL field
            $desc = if ($r.PSObject.Properties.Name -contains "Description" -and $r.Description) {
                "$($r.Description)".Trim()
            } else {
                $title
            }

            # Generate a safe, unique .aspx filename under the library root folder
            $baseName  = New-SafeBaseName $title
            $newName   = Get-UniqueAspxName -serverRelFolder $rootFolder -baseName $baseName
            $targetUrl = "$rootFolder/$newName"

            # 1) Server-side copy of template (preserves internal execution classification)
            Copy-PnPFile -SourceUrl $templateSourceUrl -TargetUrl $targetUrl -OverwriteIfAlreadyExists:$false -Force

            # 2) Update metadata on the new link item
            $item = Get-PnPFile -Url $targetUrl -AsListItem -ErrorAction Stop

            # URL field expects the common "url, description" format
            $urlValue = "$url, $desc"

            Set-PnPListItem -List $list -Identity $item.Id -ContentType $ctName -Values @{
                "Title"       = $title
                "FileLeafRef" = $newName
                $urlField.InternalName = $urlValue
                "_SourceUrl"  = $url
                # Intentionally do not set _ShortcutUrl unless your template item requires it.
            } | Out-Null

            Write-Host "Created link: $newName -> $url"
        }
        catch {
            Write-ErrLog ("Failed item. Title='{0}' URL='{1}'" -f $title, $url) $_.Exception
            Write-Warning "Failed '$title': $($_.Exception.Message)"
        }
    }

    Write-Host "Completed."
    Write-Host "Error log: $errLog"
    Write-Host "Run transcript: $runLog"
}
catch {
    Write-ErrLog "Fatal error. Script terminated." $_.Exception
    throw
}
finally {
    Stop-Transcript | Out-Null
}

# SharePoint – Fix “Link to a Document” Download Issue

This repository provides PowerShell scripts to fix an issue in SharePoint Online where “Link to a Document” items download instead of opening when clicked.
The fix requires exporting existing links and re-creating them correctly using a server-side copy approach.

**Problem Overview**

In modern SharePoint Online:
“Link to a Document” items are implemented as .aspx files
When .aspx files are uploaded or created programmatically, SharePoint marks them as non-executable
As a result, clicking the item downloads the file instead of opening the link
System fields controlling this behavior (e.g., NoExecute) cannot be updated

➡️ Existing broken links cannot be repaired in place

**Solution Approach**

The only reliable and supported fix is:
Export existing link items to CSV
Re-create the links using a server-side copy of a valid template
This mirrors how SharePoint creates links via the UI (New → Link) and preserves correct behavior.

Scripts Included
1. Export Existing Link Items

Exports existing .aspx “Link to a Document” items to CSV.

Exports:
1. File name
2. Title
3. Target URL

**   .\BulkExportLinkToDocInCSV.ps1**

**2. Import / Re-Create Link Items (Fix Download Issue)**

Re-creates links so they open correctly instead of downloading.

**Key behaviors:**

* Uses a manually created template link
* Copies the template server-side (Copy-PnPFile)
* Preserves SharePoint’s internal execution classification
* Hides the template in a _Templates folder
* Generates execution and error logs

**.\BulkCreateLinkToDocFromCSV_V1.ps1**

**Prerequisites**

* PowerShell 7+
* SharePoint Online access
* VSCode
* PnP PowerShell

Install-Module PnP.PowerShell -Scope CurrentUser

* Entra ID App Registration (Client ID)
* One manually created “Link to a Document” item in the target library
(used as a reusable template) **"Test.Aspx"**

**Important Notes**

* The template link must exist and cannot be deleted
* Existing broken links must be deleted and re-created
* The script does not force links to open in a new tab
(SharePoint does not support this via item properties)

**Recommended Workflow**

* Run BulkExportLinkToDocInCSV.ps1
* Review the exported CSV
* Delete broken .aspx link items from the library
* Create CSV with output from previous step
* Run BulkCreateLinkToDocFromCSV_V1.ps1
* Validate links now open instead of downloading

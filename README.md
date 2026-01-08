# SharePoint Online – Fix “Link to a Document” Download Issue

![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue?logo=powershell)
![SharePoint](https://img.shields.io/badge/SharePoint-Online-0078D4?logo=microsoft-sharepoint)
![PnP PowerShell](https://img.shields.io/badge/PnP-PowerShell-orange)

This repository contains PowerShell automation to remediate a known SharePoint Online issue where **“Link to a Document”** items download a `.aspx` file instead of redirecting users to the intended URL.

The remediation exports existing link items and **re-creates them using a server-side copy mechanism**, matching the behavior of links created through the SharePoint UI.

---

## TL;DR

- “Link to a Document” items are stored as `.aspx` files.
- When created programmatically or via migration tools, they may be flagged as **non-executable**.
- This causes the link to **download instead of redirect**.
- Execution-related system fields are **read-only**.
- **There is no supported in-place fix.**
- The only supported remediation is to **delete and re-create** the links using a UI-created template and a server-side copy.

---

## Problem Overview

In modern SharePoint Online, **“Link to a Document”** items are implemented as `.aspx` files.

- **Issue**  
  When these files are created outside the SharePoint UI (for example, through migration tooling or scripted uploads), SharePoint classifies them as **non-executable**.

- **Result**  
  Users experience file download behavior rather than URL redirection.

- **Constraint**  
  Internal execution controls (for example, `NoExecute`) are system-managed, read-only fields and cannot be modified via PowerShell, REST, or CSOM.

> **Conclusion:**  
> Existing broken link items cannot be repaired in place and must be re-created.

---

## Solution Overview

The only supported and reliable remediation is to replicate how SharePoint creates link items through the UI (**New → Link**).

### Remediation Strategy

1. Export existing “Link to a Document” items to CSV.
2. Re-create each link by performing a **server-side copy** of a valid, UI-created template file.

Using `Copy-PnPFile` preserves SharePoint’s internal execution classification and prevents download behavior.

---

## Prerequisites

The following requirements must be met before running the scripts:

- **PowerShell:** Version 7 or later
- **Permissions:** SharePoint Online Site Collection Administrator
- **Authentication:** Entra ID App Registration (Client ID) for PnP PowerShell
- **Module:** PnP PowerShell

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

---

## Required Template Link

A manually created **“Link to a Document”** item must exist in the target document library.

- **File name:** `Test.aspx` (configurable in script)
- **Creation method:**  
  Document Library → **New** → **Link** → Enter any URL

This file acts as the execution-safe template used for server-side copies.

---

## Scripts Included

### Export Existing Link Items

**Script:**  
`BulkExportLinkToDocInCSV.ps1`

**Function:**  
Scans the library and exports existing `.aspx` link items.

**Captured fields:**
- File Name
- Title
- Target URL

---

### Re-Create Link Items

**Script:**  
`BulkCreateLinkToDocFromCSV_V1.ps1`

**Function:**  
Re-creates link items using a server-side copy of the template file.

**Key behaviors:**
- Uses a UI-created template (`Test.aspx`)
- Performs server-side copy (`Copy-PnPFile`)
- Stores template in a hidden `_Templates` folder
- Generates execution and error logs

---

## Recommended Remediation Workflow

1. Run `BulkExportLinkToDocInCSV.ps1`
2. Review the exported CSV for accuracy
3. Delete or rename existing broken `.aspx` link items
4. Validate the CSV contains only desired links
5. Run `BulkCreateLinkToDocFromCSV_V1.ps1`
6. Validate link behavior in the SharePoint UI

---

## Known Limitations and FAQ

### Can existing links be fixed without deletion?
No. SharePoint does not support modifying execution-related system fields on existing `.aspx` files.

### Can this be fixed using CSOM, REST, or Graph?
No. All supported APIs respect the same system-level restrictions.

### Why does copying the file work?
A server-side copy of a UI-created link preserves SharePoint’s internal execution classification.

### Can links be forced to open in a new tab?
No. SharePoint does not expose a supported property to control new-tab behavior for `.aspx` link items.

### Is this supported by Microsoft?
Yes. Re-creating the link using native UI behavior is the only supported remediation and aligns with Microsoft escalation guidance.

---

## Important Notes

- The template file must exist for the duration of script execution.
- Existing broken links must be removed or renamed to avoid naming conflicts.
- This solution does not modify system-managed fields directly.

---

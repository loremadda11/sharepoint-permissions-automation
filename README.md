# SharePoint Permissions Automation

PowerShell scripts to bulk-assign granular folder permissions across hundreds of SharePoint Online project folders — using PnP PowerShell and Microsoft Graph.

---

## The problem

A procurement site on SharePoint Online contained **300 GB of data** organised in ~150 project folders (`01_PROJECTS/`). Each project had a `CONTRACTS` subfolder requiring specific, non-inherited permissions (different from the parent site):

| User / Group       | Permission level |
|--------------------|------------------|
| Project manager    | Full Control     |
| Project collaborator | Contribute     |
| Read-only user     | Read             |
| Site Members group | Read             |

Setting these manually through the SharePoint UI would have taken **approximately one week**. These scripts reduced that to a single automated run.

An additional complication: some older projects had the folder named `CONTRACTS` (no numeric prefix), while newer ones used `02_CONTRACTS`. Two separate scripts handle each case.

---

## Scripts

### `Set-ContractFolderPermissions.ps1`
Targets folders named `02_CONTRACTS` inside each top-level project folder.

### `Set-LegacyContractFolderPermissions.ps1`
Targets folders named `CONTRACTS` (no prefix), skipping any project that already has `02_CONTRACTS` (to avoid double-processing).

Both scripts share the same logic:

1. Connect to SharePoint Online via PnP PowerShell (interactive login)
2. Enumerate all top-level project folders
3. Detect which ones contain the target subfolder
4. For each target folder:
   - Reset any existing unique permissions
   - Break inheritance (without copying parent permissions)
   - Remove all leftover role assignments
   - Assign the correct permissions to the defined users/group
5. Export a CSV log with timestamp, result (OK / ERROR), and error detail

---

## Key design decisions

**Dry-run mode** (`$DryRun = $true`): shows exactly what would change without touching anything. Always run this first.

**Throttling-aware retry** (`Invoke-WithRetry`): SharePoint Graph API returns HTTP 429 (Too Many Requests) and 503 under load. The retry function handles transient errors with exponential backoff (up to 60 seconds), so the script can run unattended on large sites without failing.

**Batch limit** (`$MaxFoldersToProcess`): set to a small number (e.g. 3) for initial real-world testing before processing all folders. Set to 0 to process everything.

**Clean permission reset**: instead of adding permissions on top of existing ones (which can leave ghost assignments), the scripts always: reset inheritance → break inheritance fresh → delete all existing assignments → assign only the intended permissions. This guarantees a predictable, auditable final state.

**CSV audit log**: every run produces two CSV files — a discovery log (what was found/skipped) and a results log (what was changed and whether it succeeded). Useful for auditing and troubleshooting.

---

## Requirements

- [PnP PowerShell](https://pnp.github.io/powershell/) (`Install-Module PnP.PowerShell`)
- An Entra ID App Registration with `Sites.FullControl.All` permission (or run interactively with a SharePoint admin account)
- PowerShell 7+ recommended

---

## Setup

1. Clone this repository
2. Open the script and update the `CONFIGURAZIONE` section at the top:

```powershell
$ClientId            = "<YOUR-APP-CLIENT-ID>"
$SiteUrl             = "https://your-tenant.sharepoint.com/sites/YourSite"
$LibraryTitle        = "Shared Documents"       # or your document library name
$RootProjectsFolder  = "/sites/YourSite/Shared Documents/01_PROJECTS"
$ManagerLogin        = "manager@your-tenant.com"
$CollabLogin         = "collaborator@your-tenant.com"
$ReadOnlyLogin       = "readonly@your-tenant.com"
```

3. Run in dry-run mode first:

```powershell
$DryRun = $true
.\Set-ContractFolderPermissions.ps1
```

4. Review the discovery log CSV, then run for real on a small batch:

```powershell
$DryRun = $false
$MaxFoldersToProcess = 3
.\Set-ContractFolderPermissions.ps1
```

5. If results look correct, process everything:

```powershell
$DryRun = $false
$MaxFoldersToProcess = 0
.\Set-ContractFolderPermissions.ps1
```

---

## What I learned

- SharePoint's permission inheritance model: when and why to use `BreakRoleInheritance($false, $true)` vs `$false, $false`
- How PnP PowerShell wraps the CSOM API, and where you still need to call `ExecuteQuery()` manually
- Handling Microsoft Graph throttling gracefully in long-running scripts
- The importance of dry-run + batched testing before touching production data at scale

---

## Technologies

`PowerShell` · `PnP PowerShell` · `SharePoint Online` · `Microsoft 365` · `Entra ID`

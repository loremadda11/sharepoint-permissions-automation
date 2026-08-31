# SharePoint Permissions Automation

PowerShell automation to bulk-assign granular folder permissions across hundreds of SharePoint Online project folders using PnP PowerShell and Microsoft 365 tooling.

> **Portfolio context** — This is a sanitized example of the kind of operational automation I build as an IT Specialist: start from a repetitive real-world problem, define a safer workflow, automate it, test it in small batches, and make the result auditable.

## My role

I treated this as an operational/process problem rather than a pure scripting task:

- analysed the existing manual permission workflow and the different legacy folder structures;
- defined the target permission model and the required safety checks;
- used AI-assisted development tools alongside PowerShell to accelerate implementation while reviewing and testing the resulting logic;
- added dry-run, batching, retry logic and CSV reporting so the tool could be used safely on a large live environment;
- validated the workflow progressively before wider execution.

## The problem

A procurement site on SharePoint Online contained **300 GB of data** organised in ~150 project folders (`01_PROJECTS/`). Each project had a `CONTRACTS` subfolder requiring specific, non-inherited permissions (different from the parent site):

| User / Group | Permission level |
|---|---|
| Project manager | Full Control |
| Project collaborator | Contribute |
| Read-only user | Read |
| Site Members group | Read |

Setting these manually through the SharePoint UI would have taken **approximately one week**. The automation reduced that to a controlled batch process.

An additional complication: some older projects had the folder named `CONTRACTS`, while newer ones used `02_CONTRACTS`. Two scripts handle the two cases without double-processing.

## Workflow

```text
SharePoint project folders
        ↓
Discover target CONTRACTS folders
        ↓
Dry-run / small batch validation
        ↓
Reset and rebuild permissions
        ↓
Retry transient failures
        ↓
CSV discovery + execution reports
```

## Scripts

### `Set-ContractFolderPermissions.ps1`
Targets folders named `02_CONTRACTS` inside each top-level project folder.

### `Set-LegacyContractFolderPermissions.ps1`
Targets folders named `CONTRACTS` (no prefix), skipping any project that already has `02_CONTRACTS`.

Both scripts share the same logic:

1. Connect to SharePoint Online via PnP PowerShell
2. Enumerate all top-level project folders
3. Detect which ones contain the target subfolder
4. For each target folder:
   - reset any existing unique permissions;
   - break inheritance without copying parent permissions;
   - remove leftover role assignments;
   - assign only the intended users/groups and permission levels;
5. Export CSV logs with timestamp, result and error detail.

## Key design decisions

**Dry-run mode** — shows exactly what would change without touching anything.

**Throttling-aware retry** — transient HTTP 429 / 503 responses are handled with exponential backoff, allowing long runs to continue safely.

**Batch limit** — initial real-world testing can be restricted to a small number of folders before processing the full set.

**Clean permission reset** — permissions are rebuilt from a known state instead of layering new assignments on top of unknown legacy permissions.

**CSV audit log** — every run produces discovery and results reports for auditing and troubleshooting.

## Practical impact

The main value was removing a large amount of repetitive manual permission work while reducing inconsistency and making the process repeatable, reviewable and auditable.

## Requirements

- [PnP PowerShell](https://pnp.github.io/powershell/) (`Install-Module PnP.PowerShell`)
- An Entra ID App Registration with `Sites.FullControl.All` permission (or an appropriately privileged interactive account)
- PowerShell 7+ recommended

## Setup

1. Clone this repository
2. Update the configuration section with your own test tenant/site values
3. Start with dry-run enabled
4. Review the discovery CSV
5. Run a small real batch
6. Process the full set only after validation

Example:

```powershell
$DryRun = $true
.\Set-ContractFolderPermissions.ps1
```

Then a small batch:

```powershell
$DryRun = $false
$MaxFoldersToProcess = 3
.\Set-ContractFolderPermissions.ps1
```

## Security notes

- Use placeholder/sample tenant data in public repositories
- Do not commit real user addresses, internal site URLs, secrets or production identifiers
- Prefer dry-run and small-batch validation before changing permissions at scale
- Review the scripts before using them in your own environment

## What I learned

- SharePoint permission inheritance and role assignment behaviour
- How PnP PowerShell interacts with SharePoint CSOM
- Handling throttling and transient failures in long-running automation
- Why dry-run, batching and audit output matter more than raw script speed in production operations
- Using AI-assisted coding as an implementation accelerator while keeping problem definition, validation and operational safety human-owned

## Technologies

`PowerShell` · `PnP PowerShell` · `SharePoint Online` · `Microsoft 365` · `Entra ID`

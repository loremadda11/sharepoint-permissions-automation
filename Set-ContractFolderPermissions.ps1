$ErrorActionPreference = "Stop"

# ============================================================
# CONFIGURAZIONE — adatta questi valori al tuo ambiente
# ============================================================

$ClientId = "<YOUR-APP-CLIENT-ID>"
$SiteUrl  = "https://contoso.sharepoint.com/sites/Procurement"

$LibraryTitle = "Shared Documents"

# Cartella radice che contiene tutte le commesse/progetti
$RootProjectsFolder = "/sites/Procurement/Shared Documents/01_PROJECTS"

# Nome della sottocartella contratti da cercare in ogni progetto
$TargetFolderName = "02_CONTRACTS"

# Utenti e ruoli
$ManagerLogin    = "manager@contoso.com"
$CollabLogin     = "collaborator@contoso.com"
$ReadOnlyLogin   = "readonly.user@contoso.com"

$RoleManager     = "Full Control"
$RoleCollab      = "Contribute"
$RoleMembers     = "Read"

# TRUE  = mostra solo cosa farebbe, senza modificare nulla (consigliato al primo avvio)
# FALSE = applica davvero le modifiche
$DryRun = $true

# 0 = processa tutte le cartelle trovate
# N = processa solo le prime N (utile per test iniziali)
$MaxFoldersToProcess = 0

# TRUE = rimuove anche i permessi unici dentro 02_CONTRACTS
#        così le sottocartelle erediteranno i nuovi permessi
$ClearUniquePermissionsInside = $true

$BaseLogDir = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    (Get-Location).Path
} else {
    $PSScriptRoot
}

$LogPath = Join-Path $BaseLogDir ("log_permissions_contracts_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

# ============================================================
# FUNZIONI
# ============================================================

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory)][string]$Operation,
        [int]$MaxAttempts = 3
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            $msg = $_.Exception.Message
            $isTransient =
                $msg -match "429"      -or
                $msg -match "503"      -or
                $msg -match "throttl"  -or
                $msg -match "timeout"  -or
                $msg -match "temporar"

            if ($attempt -lt $MaxAttempts -and $isTransient) {
                $delay = [Math]::Min(60, 5 * $attempt)
                Write-Warning "Transient error during '$Operation'. Retrying in $delay seconds. Detail: $msg"
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

function Set-CleanFolderPermissions {
    param(
        [Parameter(Mandatory)][string]$FolderUrl
    )

    $Folder = Get-PnPFolder -Url $FolderUrl -Includes ListItemAllFields -ErrorAction Stop
    $Item   = $Folder.ListItemAllFields
    $Ctx    = $Item.Context

    Invoke-WithRetry -Operation "Load item $FolderUrl" -ScriptBlock {
        $Ctx.Load($Item)
        $Ctx.ExecuteQuery()
    }

    # 1. Se la cartella aveva già permessi unici, ripristina prima l'ereditarietà
    if ($Item.HasUniqueRoleAssignments) {
        Invoke-WithRetry -Operation "Reset inheritance $FolderUrl" -ScriptBlock {
            $Item.ResetRoleInheritance()
            $Item.Update()
            $Ctx.ExecuteQuery()
        }
        Write-Host "  Inheritance temporarily restored." -ForegroundColor DarkCyan
    } else {
        Write-Host "  Folder was already inheriting permissions." -ForegroundColor DarkCyan
    }

    # 2. Rompe l'ereditarietà senza copiare i permessi del padre
    Invoke-WithRetry -Operation "BreakRoleInheritance $FolderUrl" -ScriptBlock {
        $Item.BreakRoleInheritance($false, $ClearUniquePermissionsInside)
        $Item.Update()
        $Ctx.ExecuteQuery()
    }

    # 3. Rimuove tutti i permessi rimasti
    Invoke-WithRetry -Operation "Load existing permissions $FolderUrl" -ScriptBlock {
        $Ctx.Load($Item.RoleAssignments)
        $Ctx.ExecuteQuery()
    }

    @($Item.RoleAssignments) | ForEach-Object { $_.DeleteObject() }

    Invoke-WithRetry -Operation "Delete existing permissions $FolderUrl" -ScriptBlock {
        $Ctx.ExecuteQuery()
    }
    Write-Host "  Existing permissions removed." -ForegroundColor DarkCyan

    # 4. Assegna i permessi corretti
    Invoke-WithRetry -Operation "Assign Manager - Full Control" -ScriptBlock {
        Set-PnPListItemPermission -List $LibraryTitle -Identity $Item.Id `
            -User $ManagerLogin -AddRole $RoleManager -ErrorAction Stop
    }

    Invoke-WithRetry -Operation "Assign Collaborator - Contribute" -ScriptBlock {
        Set-PnPListItemPermission -List $LibraryTitle -Identity $Item.Id `
            -User $CollabLogin -AddRole $RoleCollab -ErrorAction Stop
    }

    Invoke-WithRetry -Operation "Assign ReadOnly - Read" -ScriptBlock {
        Set-PnPListItemPermission -List $LibraryTitle -Identity $Item.Id `
            -User $ReadOnlyLogin -AddRole $RoleCollab -ErrorAction Stop
    }

    Invoke-WithRetry -Operation "Assign Members group - Read" -ScriptBlock {
        Set-PnPListItemPermission -List $LibraryTitle -Identity $Item.Id `
            -Group $MembersGroup.Title -AddRole $RoleMembers -ErrorAction Stop
    }
}

# ============================================================
# CONNESSIONE
# ============================================================

Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$List = Get-PnPList -Identity $LibraryTitle -ErrorAction Stop

$MembersGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction Stop
if (-not $MembersGroup -or -not $MembersGroup.Title) {
    throw "Unable to retrieve the site Members group."
}
Write-Host "Members group found: $($MembersGroup.Title)" -ForegroundColor Green

New-PnPUser -LoginName $ManagerLogin  -ErrorAction Stop | Out-Null
New-PnPUser -LoginName $CollabLogin   -ErrorAction Stop | Out-Null
New-PnPUser -LoginName $ReadOnlyLogin -ErrorAction Stop | Out-Null

# ============================================================
# DISCOVERY: trova le cartelle da processare
# ============================================================

Write-Host ""
Write-Host "Reading projects under:" -ForegroundColor Cyan
Write-Host $RootProjectsFolder -ForegroundColor Cyan

$RootFolder = Get-PnPFolder -Url $RootProjectsFolder -Includes Folders, ServerRelativeUrl -ErrorAction Stop
$Ctx = Get-PnPContext

Invoke-WithRetry -Operation "Load project folders" -ScriptBlock {
    $Ctx.Load($RootFolder.Folders)
    $Ctx.ExecuteQuery()
}

[array]$Projects = @(
    $RootFolder.Folders |
    Where-Object { $_.Name -ne "Forms" -and $_.Name -notlike "_*" } |
    Sort-Object Name
)

Write-Host "Top-level project folders found: $($Projects.Count)" -ForegroundColor Green

$DiscoveryResults = New-Object System.Collections.Generic.List[object]
$TargetFolders    = New-Object System.Collections.Generic.List[object]

foreach ($Project in $Projects) {
    $CandidateUrl = $Project.ServerRelativeUrl.TrimEnd("/") + "/" + $TargetFolderName

    try {
        Get-PnPFolder -Url $CandidateUrl -Includes ListItemAllFields -ErrorAction Stop | Out-Null

        $TargetFolders.Add([PSCustomObject]@{ Project = $Project.Name; FolderUrl = $CandidateUrl })
        $DiscoveryResults.Add([PSCustomObject]@{ Result = "FOUND";   Project = $Project.Name; Folder = $CandidateUrl; Error = "" })
    }
    catch {
        $DiscoveryResults.Add([PSCustomObject]@{ Result = "SKIPPED"; Project = $Project.Name; Folder = $CandidateUrl; Error = "$TargetFolderName not found" })
    }
}

if ($MaxFoldersToProcess -gt 0) {
    [array]$TargetFolders = @($TargetFolders | Select-Object -First $MaxFoldersToProcess)
}

Write-Host ""
Write-Host "$TargetFolderName folders to process: $($TargetFolders.Count)" -ForegroundColor Green

$DiscoveryLogPath = $LogPath.Replace(".csv", "_discovery.csv")
$DiscoveryResults | Export-Csv -Path $DiscoveryLogPath -NoTypeInformation -Encoding UTF8
Write-Host "Discovery log saved to: $DiscoveryLogPath" -ForegroundColor Green

if ($TargetFolders.Count -eq 0) {
    Write-Host "No $TargetFolderName folders found. Nothing to do." -ForegroundColor Yellow
    return
}

if ($DryRun) {
    Write-Host ""
    Write-Host "DRY-RUN MODE: no changes will be made." -ForegroundColor Yellow
    Write-Host "Folders that would be updated:" -ForegroundColor Yellow
    $TargetFolders | ForEach-Object { Write-Host "  $($_.FolderUrl)" }
    Write-Host ""
    Write-Host "To run a limited real test, set:" -ForegroundColor Yellow
    Write-Host '  $DryRun = $false  |  $MaxFoldersToProcess = 3'
    Write-Host "To process all folders:" -ForegroundColor Yellow
    Write-Host '  $DryRun = $false  |  $MaxFoldersToProcess = 0'
    return
}

# ============================================================
# APPLICAZIONE PERMESSI
# ============================================================

$Results = New-Object System.Collections.Generic.List[object]
$Index   = 0

Write-Host ""
Write-Warning "LIVE MODE: applying permission changes."
Write-Host ""

foreach ($Target in $TargetFolders) {
    $Index++
    $FolderUrl = $Target.FolderUrl

    Write-Progress -Activity "Updating permissions" `
        -Status "$Index / $($TargetFolders.Count) — $FolderUrl" `
        -PercentComplete (($Index / [Math]::Max($TargetFolders.Count,1)) * 100)

    try {
        Write-Host "[$Index/$($TargetFolders.Count)] Processing: $FolderUrl" -ForegroundColor Cyan
        Set-CleanFolderPermissions -FolderUrl $FolderUrl
        Write-Host "[OK]" -ForegroundColor Green

        $Results.Add([PSCustomObject]@{
            Timestamp = Get-Date; Result = "OK"
            Project = $Target.Project; Folder = $FolderUrl; Error = ""
        })
    }
    catch {
        $Err = $_.Exception.Message
        Write-Host "[ERROR] $Err" -ForegroundColor Red

        $Results.Add([PSCustomObject]@{
            Timestamp = Get-Date; Result = "ERROR"
            Project = $Target.Project; Folder = $FolderUrl; Error = $Err
        })
    }
}

Write-Progress -Activity "Updating permissions" -Completed
$Results | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8

$OkCount  = ($Results | Where-Object { $_.Result -eq "OK"    }).Count
$ErrCount = ($Results | Where-Object { $_.Result -eq "ERROR" }).Count

Write-Host ""
Write-Host "Done. OK: $OkCount  |  Errors: $ErrCount" -ForegroundColor Green
Write-Host "Log saved to: $LogPath" -ForegroundColor Green

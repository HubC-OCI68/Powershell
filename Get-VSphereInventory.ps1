#Requires -Version 5.1
<#
.SYNOPSIS
    Collecte l'état complet d'une infrastructure vSphere 8 via vCenter :
    inventaire, capacité, provisionnement, performances, santé et points d'attention.

.DESCRIPTION
    Se connecte à un vCenter avec VMware PowerCLI et génère dans un dossier horodaté :
      - un classeur Excel multi-onglets (si le module ImportExcel est installé),
      - un fichier CSV par section (séparateur ';', UTF-8 - s'ouvre directement dans Excel FR),
      - un rapport HTML de synthèse avec les points d'attention colorés,
      - le journal d'exécution (transcript).

    Données collectées :
      - vCenter : version, build, licences (clés masquées), alarmes déclenchées
      - Clusters : HA / contrôle d'admission / DRS / EVC / vSAN, capacité CPU-RAM, utilisation,
                   ratios de consolidation vCPU:cœur et vRAM:RAM, tolérance N+1, homogénéité des builds
      - Hôtes ESXi : matériel, n° de série, version/build, uptime, charge CPU/RAM, VMs hébergées,
                     NTP, SSH, Shell, lockdown, DNS, syslog, politique d'énergie,
                     cartes physiques, VMkernel, HBA, vSwitch standard
      - VMs : état, OS, IP, vCPU/RAM, charge instantanée, mémoire active/consommée/ballooning/swap,
              disques (thin/thick), stockage provisionné/consommé, espace libre invité,
              VMware Tools, version matérielle, réservations/limites, snapshots
      - Performances historiques (VMs et hôtes) : CPU, RAM, CPU Ready (moyenne, P95, max)
                     + indications de sur/sous-dimensionnement
      - Datastores : type, capacité, libre, provisionné, surprovisionnement, version VMFS / cible NFS
      - Réseau : portgroups standard et distribués (VLAN), vDS

    Le script est en LECTURE SEULE : il ne modifie rien.
    Un compte vCenter avec le rôle "Lecture seule" propagé depuis la racine de vCenter suffit.

.PARAMETER vCenter
    FQDN ou IP du vCenter.
.PARAMETER Credential
    Identifiants vCenter. Demandés interactivement si absents.
.PARAMETER OutputPath
    Dossier racine des rapports (défaut : Documents\vSphere-Inventaire).
.PARAMETER StatsDays
    Profondeur d'historique des performances en jours (défaut 7).
    1 = pas de 5 min, <=7 = 30 min, <=30 = 2 h, au-delà = 1 jour.
.PARAMETER SkipPerfStats
    Ne collecte pas les performances historiques (exécution beaucoup plus rapide).
.PARAMETER SkipHostDetails
    Ne collecte pas la configuration détaillée des hôtes (services, NIC, VMkernel, HBA, vSwitch).
.PARAMETER SnapshotAgeWarningDays
    Âge (jours) au-delà duquel un snapshot est signalé (défaut 3).
.PARAMETER DatastoreWarningPct
    Taux de remplissage (%) au-delà duquel un datastore est signalé (défaut 85).
.PARAMETER IgnoreInvalidCertificate
    N'affiche pas d'avertissement si le certificat du vCenter n'est pas approuvé par le poste.
.PARAMETER NoOpen
    N'ouvre pas le rapport HTML à la fin.

.EXAMPLE
    .\Get-VSphereInventory.ps1 -vCenter vcsa.domaine.local

.EXAMPLE
    .\Get-VSphereInventory.ps1 -vCenter vcsa.domaine.local -StatsDays 30 -OutputPath D:\Rapports

.EXAMPLE
    $cred = Get-Credential
    .\Get-VSphereInventory.ps1 -vCenter 10.0.0.10 -Credential $cred -SkipPerfStats -IgnoreInvalidCertificate

.NOTES
    Prérequis (poste Windows, PowerShell 5.1 ou 7.x) :
      Install-Module VCF.PowerCLI -Scope CurrentUser     # PowerCLI 9.x (ou VMware.PowerCLI 13.x)
      Install-Module ImportExcel  -Scope CurrentUser     # optionnel : génération du .xlsx
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$vCenter,

    [System.Management.Automation.PSCredential]$Credential,

    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'vSphere-Inventaire'),

    [ValidateRange(1, 365)]
    [int]$StatsDays = 7,

    [switch]$SkipPerfStats,

    [switch]$SkipHostDetails,

    [ValidateRange(1, 3650)]
    [int]$SnapshotAgeWarningDays = 3,

    [ValidateRange(50, 99)]
    [int]$DatastoreWarningPct = 85,

    [switch]$IgnoreInvalidCertificate,

    [switch]$NoOpen
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

#region ======================= Fonctions utilitaires =======================

function New-List { , (New-Object 'System.Collections.Generic.List[object]') }

function Rnd {
    param([object]$Value, [int]$Digits = 1)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    [math]::Round([double]$Value, $Digits)
}

function Num {
    # Convertit n'importe quelle valeur numérique de l'API (Int16, Int64, Nullable...) en [double] ; $null -> 0
    param([object]$Value)
    if ($null -eq $Value -or "$Value" -eq '') { return [double]0 }
    [double]$Value
}

function Write-Step {
    param([string]$Message)
    Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message) -ForegroundColor Cyan
}

function Invoke-Section {
    # Exécute une section de collecte ; une erreur n'interrompt pas le reste du script.
    param([string]$Name, [scriptblock]$Body)
    Write-Step $Name
    try { & $Body | Out-Null }
    catch {
        $msg = "$Name : $($_.Exception.Message) (ligne $($_.InvocationInfo.ScriptLineNumber) : $($_.InvocationInfo.Line.Trim()))"
        Write-Warning $msg
        $script:CollectErrors.Add($msg)
    }
}

function Add-Finding {
    param(
        [ValidateSet('CRITIQUE', 'ALERTE', 'INFO')][string]$Severite,
        [string]$Categorie, [string]$Objet, [string]$Detail
    )
    $script:findings.Add([pscustomobject]@{
            Severite  = $Severite
            Categorie = $Categorie
            Objet     = $Objet
            Detail    = $Detail
        })
}

function Get-ClusterNameOfHost {
    param($HostView)
    if (-not $HostView) { return '' }
    $p = "$($HostView.Parent)"
    if ($script:clusterById.ContainsKey($p)) { return $script:clusterById[$p].Name }
    return '(hôte autonome)'
}

function Get-SnapshotTree {
    param($List, [string]$VmName, [string]$VmId)
    foreach ($s in $List) {
        $created = $s.CreateTime
        [pscustomobject]@{
            _Id         = "$($s.Snapshot)"
            _VmId       = $VmId
            VM          = $VmName
            Snapshot    = $s.Name
            Description = $s.Description
            Cree_le     = $created.ToLocalTime()
            Age_Jours   = Rnd (((Get-Date).ToUniversalTime() - $created.ToUniversalTime()).TotalDays)
            Avec_RAM    = ($s.State -eq 'poweredOn')
            Quiesce     = $s.Quiesced
            Taille_Go   = $null
        }
        if ($s.ChildSnapshotList) {
            Get-SnapshotTree -List $s.ChildSnapshotList -VmName $VmName -VmId $VmId
        }
    }
}

function Get-StatSummary {
    param($Values)
    $a = @($Values | Where-Object { $null -ne $_ })
    if ($a.Count -eq 0) { return @{ Avg = $null; P95 = $null; Max = $null; Count = 0 } }
    $sorted = @($a | Sort-Object)
    $idx = [math]::Max(0, [math]::Ceiling(0.95 * $sorted.Count) - 1)
    @{
        Avg   = Rnd (($a | Measure-Object -Average).Average)
        P95   = Rnd $sorted[$idx]
        Max   = Rnd $sorted[$sorted.Count - 1]
        Count = $a.Count
    }
}

#endregion

#region ======================= Initialisation =======================

$CollectErrors = New-List
$findings      = New-List
$summaryRows   = New-List
$licenseRows   = New-List
$clusterRows   = New-List
$hostRows      = New-List
$vmRows        = New-List
$snapRows      = New-List
$perfRows      = New-List
$dsRows        = New-List
$alarmRows     = New-List
$hostCfgRows   = New-List
$nicRows       = New-List
$vmkRows       = New-List
$hbaRows       = New-List
$vssRows       = New-List
$pgRows        = New-List
$vdsRows       = New-List

$hostById    = @{}
$clusterById = @{}
$dsNameById  = @{}
$netNameById = @{}
$dcByEntity  = @{}

$startTime = Get-Date
$stamp     = Get-Date -Format 'yyyyMMdd-HHmm'
$safeName  = $vCenter -replace '[^\w\.-]', '_'
$runFolder = Join-Path $OutputPath "$($safeName)_$stamp"
New-Item -ItemType Directory -Path $runFolder -Force | Out-Null
Start-Transcript -Path (Join-Path $runFolder 'execution.log') | Out-Null

#endregion

try {

    #region ======================= PowerCLI & connexion =======================

    Write-Step 'Chargement de PowerCLI'
    if (-not (Get-Module -ListAvailable -Name VMware.VimAutomation.Core)) {
        throw "PowerCLI n'est pas installé. Exécutez : Install-Module VCF.PowerCLI -Scope CurrentUser"
    }
    Import-Module VMware.VimAutomation.Core -ErrorAction Stop -WarningAction SilentlyContinue
    Import-Module VMware.VimAutomation.Vds  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

    $certAction = if ($IgnoreInvalidCertificate) { 'Ignore' } else { 'Warn' }
    Set-PowerCLIConfiguration -Scope Session -InvalidCertificateAction $certAction -DefaultVIServerMode Single -Confirm:$false | Out-Null
    try { Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false -ErrorAction Stop | Out-Null } catch { }

    if (-not $Credential) {
        $Credential = Get-Credential -Message "Compte vCenter pour $vCenter (rôle Lecture seule suffisant)"
    }

    Write-Step "Connexion à $vCenter"
    $vi    = Connect-VIServer -Server $vCenter -Credential $Credential -ErrorAction Stop
    $si    = Get-View ServiceInstance -Server $vi
    $about = $si.Content.About

    #endregion

    #region ======================= Objets de base (Get-View, rapide) =======================

    Write-Step 'Collecte des objets de base'
    $vmProps = @(
        'Name', 'OverallStatus', 'Config.Template', 'Config.Hardware.NumCPU', 'Config.Hardware.NumCoresPerSocket',
        'Config.Hardware.MemoryMB', 'Config.Hardware.Device', 'Config.GuestFullName', 'Config.Version',
        'Config.Annotation', 'Config.CpuHotAddEnabled', 'Config.MemoryHotAddEnabled', 'Config.CpuAllocation',
        'Config.MemoryAllocation', 'Config.Firmware', 'Runtime.PowerState', 'Runtime.Host',
        'Runtime.ConnectionState', 'Summary.Storage', 'Summary.QuickStats', 'Guest.ToolsRunningStatus',
        'Guest.ToolsVersionStatus2', 'Guest.IpAddress', 'Guest.HostName', 'Guest.GuestFullName',
        'Guest.Net', 'Guest.Disk', 'Snapshot', 'Datastore', 'Network'
    )

    $dcViews      = @(Get-View -ViewType Datacenter -Property Name)
    $clusterViews = @(Get-View -ViewType ClusterComputeResource -Property Name, Parent, Host, Summary, ConfigurationEx, OverallStatus)
    $hostViews    = @(Get-View -ViewType HostSystem -Property Name, Parent, Summary, Runtime, Hardware.SystemInfo, Config.LockdownMode, OverallStatus)
    $dsViews      = @(Get-View -ViewType Datastore -Property Name, Summary, Info, Host, Vm, OverallStatus)
    $netViews     = @(Get-View -ViewType Network -Property Name)
    $vmViews      = @(Get-View -ViewType VirtualMachine -Property $vmProps)

    foreach ($c in $clusterViews) { $clusterById["$($c.MoRef)"] = $c }
    foreach ($h in $hostViews)    { $hostById["$($h.MoRef)"] = $h }
    foreach ($d in $dsViews)      { $dsNameById["$($d.MoRef)"] = $d.Name }
    foreach ($n in $netViews)     { $netNameById["$($n.MoRef)"] = $n.Name }
    foreach ($dc in $dcViews) {
        foreach ($t in 'ClusterComputeResource', 'HostSystem', 'Datastore') {
            foreach ($o in @(Get-View -ViewType $t -SearchRoot $dc.MoRef -Property Name)) {
                $dcByEntity["$($o.MoRef)"] = $dc.Name
            }
        }
    }

    #endregion

    #region ======================= Machines virtuelles =======================

    Invoke-Section 'Machines virtuelles' {
        $templates = 0
        foreach ($v in $vmViews) {
            $cfg = $v.Config
            if ($null -eq $cfg) {
                Add-Finding 'CRITIQUE' 'VM' $v.Name "VM inaccessible / orpheline (état : $($v.Runtime.ConnectionState))"
                continue
            }
            if ($cfg.Template) { $templates++; continue }

            $vmId    = "$($v.MoRef)"
            $hostId  = "$($v.Runtime.Host)"
            $h       = $hostById[$hostId]
            $qs      = $v.Summary.QuickStats
            $st      = $v.Summary.Storage
            $nCpu    = [int]$cfg.Hardware.NumCPU
            $memMB   = [int]$cfg.Hardware.MemoryMB
            $on      = ($v.Runtime.PowerState -eq 'poweredOn')
            $hostMhz = if ($h) { Num $h.Summary.Hardware.CpuMhz } else { [double]0 }

            $cpuPct = $null
            if ($on -and $nCpu -gt 0 -and $hostMhz -gt 0) { $cpuPct = Rnd (100 * (Num $qs.OverallCpuUsage) / ($nCpu * $hostMhz)) }

            # Disques virtuels
            $disks  = @($cfg.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualDisk] })
            $diskGB = 0; $thin = 0
            foreach ($d in $disks) {
                $diskGB += (Num $d.CapacityInKB) / 1MB
                if ($d.Backing.ThinProvisioned) { $thin++ }
            }

            # Espace libre minimal côté OS invité (nécessite les VMware Tools)
            $guestMinFree = $null; $guestVol = ''
            $gd = @($v.Guest.Disk | Where-Object { $_.Capacity -gt 0 })
            if ($gd.Count) {
                $worst = $gd | Sort-Object { $_.FreeSpace / $_.Capacity } | Select-Object -First 1
                $guestMinFree = Rnd (100 * (Num $worst.FreeSpace) / (Num $worst.Capacity))
                $guestVol = $worst.DiskPath
            }

            $ips = @($v.Guest.Net | ForEach-Object { $_.IpAddress } | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' }) -join ', '
            if (-not $ips) { $ips = $v.Guest.IpAddress }

            $snaps = @()
            if ($v.Snapshot) {
                $snaps = @(Get-SnapshotTree -List $v.Snapshot.RootSnapshotList -VmName $v.Name -VmId $vmId)
                foreach ($s in $snaps) { $snapRows.Add($s) }
            }

            $cpuLim = $cfg.CpuAllocation.Limit
            $memLim = $cfg.MemoryAllocation.Limit
            $cpuLimTxt = if ($null -eq $cpuLim -or $cpuLim -lt 0) { 'Illimitée' } else { $cpuLim }
            $memLimTxt = if ($null -eq $memLim -or $memLim -lt 0) { 'Illimitée' } else { $memLim }

            $row = [pscustomobject]@{
                _Id                      = $vmId
                _HostId                  = $hostId
                _ClusterId               = if ($h) { "$($h.Parent)" } else { '' }
                Datacenter               = $dcByEntity[$hostId]
                Cluster                  = Get-ClusterNameOfHost $h
                Hote                     = if ($h) { $h.Name } else { '' }
                VM                       = $v.Name
                Etat                     = "$($v.Runtime.PowerState)"
                Statut                   = "$($v.OverallStatus)"
                OS_Configure             = $cfg.GuestFullName
                OS_Detecte               = $v.Guest.GuestFullName
                Nom_DNS                  = $v.Guest.HostName
                IP                       = $ips
                vCPU                     = $nCpu
                Coeurs_par_Socket        = $cfg.Hardware.NumCoresPerSocket
                RAM_Go                   = Rnd ($memMB / 1024)
                CPU_Utilise_MHz          = if ($on) { $qs.OverallCpuUsage } else { $null }
                CPU_Utilise_Pct          = $cpuPct
                RAM_Active_Go            = if ($on) { Rnd ((Num $qs.GuestMemoryUsage) / 1024) } else { $null }
                RAM_Consommee_Go         = if ($on) { Rnd ((Num $qs.HostMemoryUsage) / 1024) } else { $null }
                RAM_Ballooning_Mo        = $qs.BalloonedMemory
                RAM_Swap_Mo              = $qs.SwappedMemory
                Nb_Disques               = $disks.Count
                Disques_Thin             = $thin
                Disques_Taille_Go        = Rnd $diskGB
                Stockage_Consomme_Go     = Rnd ($st.Committed / 1GB)
                Stockage_Provisionne_Go  = Rnd (($st.Committed + $st.Uncommitted) / 1GB)
                Invite_LibreMin_Pct      = $guestMinFree
                Invite_VolumeLePlusPlein = $guestVol
                Datastores               = @($v.Datastore | ForEach-Object { $dsNameById["$_"] }) -join ', '
                Reseaux                  = @($v.Network | ForEach-Object { $netNameById["$_"] }) -join ', '
                Tools_Etat               = "$($v.Guest.ToolsRunningStatus)"
                Tools_Version            = "$($v.Guest.ToolsVersionStatus2)"
                Version_Materielle       = $cfg.Version
                Firmware                 = $cfg.Firmware
                HotAdd_CPU               = $cfg.CpuHotAddEnabled
                HotAdd_RAM               = $cfg.MemoryHotAddEnabled
                Reservation_CPU_MHz      = $cfg.CpuAllocation.Reservation
                Limite_CPU_MHz           = $cpuLimTxt
                Reservation_RAM_Mo       = $cfg.MemoryAllocation.Reservation
                Limite_RAM_Mo            = $memLimTxt
                Nb_Snapshots             = $snaps.Count
                Uptime_Jours             = if ($on -and (Num $qs.UptimeSeconds) -gt 0) { Rnd ((Num $qs.UptimeSeconds) / 86400) } else { $null }
                Notes                    = ($cfg.Annotation -replace '[\r\n]+', ' ')
            }
            $vmRows.Add($row)

            # Points d'attention VM
            if ($on -and $v.Guest.ToolsRunningStatus -ne 'guestToolsRunning') {
                Add-Finding 'ALERTE' 'VM' $v.Name 'VMware Tools non démarrés sur une VM allumée'
            }
            if ("$($v.Guest.ToolsVersionStatus2)" -in 'guestToolsNeedUpgrade', 'guestToolsTooOld', 'guestToolsSupportedOld') {
                Add-Finding 'INFO' 'VM' $v.Name "VMware Tools à mettre à jour ($($v.Guest.ToolsVersionStatus2))"
            }
            if ("$($v.Guest.ToolsVersionStatus2)" -eq 'guestToolsNotInstalled') {
                Add-Finding 'ALERTE' 'VM' $v.Name 'VMware Tools non installés'
            }
            if ($null -ne $guestMinFree -and $guestMinFree -lt 10) {
                $sev = if ($guestMinFree -lt 5) { 'CRITIQUE' } else { 'ALERTE' }
                Add-Finding $sev 'VM' $v.Name "Volume invité $guestVol : $guestMinFree % libre"
            }
            if ((Num $qs.BalloonedMemory) -gt 0 -or (Num $qs.SwappedMemory) -gt 0) {
                Add-Finding 'ALERTE' 'VM' $v.Name "Contention mémoire : ballooning $($qs.BalloonedMemory) Mo / swap $($qs.SwappedMemory) Mo"
            }
            if ($cpuLimTxt -ne 'Illimitée' -or $memLimTxt -ne 'Illimitée') {
                Add-Finding 'ALERTE' 'VM' $v.Name "Limite définie (CPU : $cpuLimTxt MHz / RAM : $memLimTxt Mo)"
            }
            if ("$($v.OverallStatus)" -eq 'red') { Add-Finding 'ALERTE' 'VM' $v.Name 'Statut global rouge dans vCenter' }
        }
        $script:templateCount = $templates
    }

    # Taille des snapshots + alertes d'âge
    Invoke-Section 'Snapshots' {
        if ($snapRows.Count) {
            $ids = @($snapRows | ForEach-Object { $_._VmId } | Select-Object -Unique)
            $sizes = @{}
            Get-VM -Id $ids | Get-Snapshot | ForEach-Object { $sizes[$_.Id] = Rnd $_.SizeGB 2 }
            foreach ($s in $snapRows) {
                if ($sizes.ContainsKey($s._Id)) { $s.Taille_Go = $sizes[$s._Id] }
                if ($s.Age_Jours -ge $SnapshotAgeWarningDays) {
                    Add-Finding 'ALERTE' 'Snapshot' $s.VM "Snapshot '$($s.Snapshot)' âgé de $($s.Age_Jours) j ($($s.Taille_Go) Go)"
                }
            }
        }
    }

    #endregion

    #region ======================= Hôtes ESXi =======================

    Invoke-Section 'Hôtes ESXi' {
        $vmByHost = @{}
        foreach ($r in $vmRows) {
            if (-not $r._HostId) { continue }
            if (-not $vmByHost.ContainsKey($r._HostId)) { $vmByHost[$r._HostId] = New-List }
            $vmByHost[$r._HostId].Add($r)
        }

        foreach ($h in ($hostViews | Sort-Object Name)) {
            $id   = "$($h.MoRef)"
            $hw   = $h.Summary.Hardware
            $qs   = $h.Summary.QuickStats
            $prod = $h.Summary.Config.Product
            # Pas de "$x = if (...) { @(...) } else { @() }" : bug PowerShell "Argument types do not match"
            $vms = @()
            if ($vmByHost.ContainsKey($id)) { $vms = $vmByHost[$id].ToArray() }
            $on   = @($vms | Where-Object { $_.Etat -eq 'poweredOn' })
            [double]$vcpuOn = 0; [double]$vramOn = 0
            foreach ($x in $on) { $vcpuOn += (Num $x.vCPU); $vramOn += (Num $x.RAM_Go) }

            # Conversion explicite des valeurs de l'API (Int16 / Nullable<Int32> / Int64) en [double]
            # (évite "Les types des arguments ne correspondent pas" sous Windows PowerShell 5.1)
            $nCores    = Num $hw.NumCpuCores
            $cpuMhz    = Num $hw.CpuMhz
            $cpuUsed   = Num $qs.OverallCpuUsage
            $memUsedMB = Num $qs.OverallMemoryUsage
            $uptimeD   = (Num $qs.Uptime) / 86400

            $cpuCapMhz = $nCores * $cpuMhz
            $memGB     = (Num $hw.MemorySize) / 1GB
            $memUsedGB = $memUsedMB / 1024
            $cpuPct    = if ($cpuCapMhz -gt 0) { Rnd (100 * $cpuUsed / $cpuCapMhz) } else { $null }
            $memPct    = if ($memGB -gt 0) { Rnd (100 * $memUsedGB / $memGB) } else { $null }

            $hostRows.Add([pscustomobject]@{
                    _Id                   = $id
                    _ClusterId            = "$($h.Parent)"
                    _CpuCapMhz            = $cpuCapMhz
                    _CpuUsedMhz           = $cpuUsed
                    _MemUsedGB            = $memUsedGB
                    Datacenter            = $dcByEntity[$id]
                    Cluster               = Get-ClusterNameOfHost $h
                    Hote                  = $h.Name
                    Connexion             = "$($h.Runtime.ConnectionState)"
                    Maintenance           = [bool]$h.Runtime.InMaintenanceMode
                    Statut                = "$($h.OverallStatus)"
                    Constructeur          = $hw.Vendor
                    Modele                = $hw.Model
                    Num_Serie             = $h.Hardware.SystemInfo.SerialNumber
                    CPU_Modele            = $hw.CpuModel
                    Sockets               = [int](Num $hw.NumCpuPkgs)
                    Coeurs                = [int]$nCores
                    Threads               = [int](Num $hw.NumCpuThreads)
                    CPU_MHz               = [int]$cpuMhz
                    RAM_Go                = Rnd $memGB 0
                    ESXi_Version          = $prod.Version
                    ESXi_Build            = $prod.Build
                    Uptime_Jours          = if ($uptimeD -gt 0) { Rnd $uptimeD } else { $null }
                    CPU_Capacite_GHz      = Rnd ($cpuCapMhz / 1000)
                    CPU_Utilise_GHz       = Rnd ($cpuUsed / 1000)
                    CPU_Pct               = $cpuPct
                    RAM_Utilisee_Go       = Rnd $memUsedGB
                    RAM_Pct               = $memPct
                    VMs_Total             = $vms.Count
                    VMs_Allumees          = $on.Count
                    vCPU_VMs_Allumees     = [int]$vcpuOn
                    Ratio_vCPU_par_Coeur  = if ($nCores -gt 0) { Rnd ($vcpuOn / $nCores) 2 } else { $null }
                    vRAM_VMs_Allumees_Go  = Rnd $vramOn 0
                    Ratio_vRAM_RAM        = if ($memGB -gt 0) { Rnd ($vramOn / $memGB) 2 } else { $null }
                })

            if ("$($h.Runtime.ConnectionState)" -ne 'connected') {
                Add-Finding 'CRITIQUE' 'Hôte' $h.Name "Hôte $($h.Runtime.ConnectionState)"
            }
            if ($h.Runtime.InMaintenanceMode) { Add-Finding 'INFO' 'Hôte' $h.Name 'Hôte en mode maintenance' }
            switch ("$($h.OverallStatus)") {
                'red'    { Add-Finding 'CRITIQUE' 'Hôte' $h.Name 'Statut global rouge dans vCenter' }
                'yellow' { Add-Finding 'ALERTE'   'Hôte' $h.Name 'Statut global jaune dans vCenter' }
            }
            if ($cpuPct -ge 85) { Add-Finding 'ALERTE' 'Hôte' $h.Name "CPU instantané à $cpuPct %" }
            if ($memPct -ge 90) { Add-Finding 'ALERTE' 'Hôte' $h.Name "RAM utilisée à $memPct %" }
            if ($uptimeD -gt 365) {
                Add-Finding 'INFO' 'Hôte' $h.Name "Uptime > 1 an ($([int]$uptimeD) j) : vérifier l'application des correctifs"
            }
        }
    }

    #endregion

    #region ======================= Clusters =======================

    Invoke-Section 'Clusters' {
        foreach ($c in ($clusterViews | Sort-Object Name)) {
            $cid      = "$($c.MoRef)"
            $hs       = @($hostRows | Where-Object { $_._ClusterId -eq $cid })
            $hsActive = @($hs | Where-Object { $_.Connexion -eq 'connected' -and -not $_.Maintenance })

            $cores = 0; $cpuCap = 0; $cpuUsed = 0; $memCap = 0; $memUsed = 0; $maxHostMem = 0; $maxHostCpu = 0
            foreach ($x in $hsActive) {
                $cores += $x.Coeurs; $cpuCap += $x._CpuCapMhz; $cpuUsed += $x._CpuUsedMhz
                $memCap += $x.RAM_Go; $memUsed += $x._MemUsedGB
                if ($x.RAM_Go -gt $maxHostMem) { $maxHostMem = $x.RAM_Go }
                if ($x._CpuCapMhz -gt $maxHostCpu) { $maxHostCpu = $x._CpuCapMhz }
            }

            $vms   = @($vmRows | Where-Object { $_._ClusterId -eq $cid })
            $vmsOn = @($vms | Where-Object { $_.Etat -eq 'poweredOn' })
            $vcpuAll = 0; $vcpuOn = 0; $vramAll = 0; $vramOn = 0; $stoProv = 0; $stoUsed = 0
            foreach ($x in $vms) { $vcpuAll += $x.vCPU; $vramAll += $x.RAM_Go; $stoProv += $x.Stockage_Provisionne_Go; $stoUsed += $x.Stockage_Consomme_Go }
            foreach ($x in $vmsOn) { $vcpuOn += $x.vCPU; $vramOn += $x.RAM_Go }

            $das = $c.ConfigurationEx.DasConfig
            $drs = $c.ConfigurationEx.DrsConfig
            $n1Mem = ($hsActive.Count -gt 1) -and ($memUsed -le ($memCap - $maxHostMem))
            $n1Cpu = ($hsActive.Count -gt 1) -and ($cpuUsed -le ($cpuCap - $maxHostCpu))
            $builds = @($hs | ForEach-Object { $_.ESXi_Build } | Select-Object -Unique)
            $ratioCpu = if ($cores) { Rnd ($vcpuOn / $cores) 2 } else { $null }
            $ratioMem = if ($memCap) { Rnd ($vramOn / $memCap) 2 } else { $null }

            $clusterRows.Add([pscustomobject]@{
                    Datacenter                = $dcByEntity[$cid]
                    Cluster                   = $c.Name
                    Statut                    = "$($c.OverallStatus)"
                    Hotes                     = $hs.Count
                    Hotes_Actifs              = $hsActive.Count
                    HA                        = $das.Enabled
                    HA_Controle_Admission     = $das.AdmissionControlEnabled
                    HA_Surveillance_Hotes     = $das.HostMonitoring
                    DRS                       = $drs.Enabled
                    DRS_Mode                  = "$($drs.DefaultVmBehavior)"
                    EVC                       = $c.Summary.CurrentEVCModeKey
                    vSAN                      = $c.ConfigurationEx.VsanConfigInfo.Enabled
                    Builds_ESXi               = $builds -join ', '
                    Coeurs_Physiques          = $cores
                    CPU_Capacite_GHz          = Rnd ($cpuCap / 1000)
                    CPU_Utilise_GHz           = Rnd ($cpuUsed / 1000)
                    CPU_Pct                   = if ($cpuCap) { Rnd (100 * $cpuUsed / $cpuCap) } else { $null }
                    RAM_Totale_Go             = Rnd $memCap 0
                    RAM_Utilisee_Go           = Rnd $memUsed 0
                    RAM_Pct                   = if ($memCap) { Rnd (100 * $memUsed / $memCap) } else { $null }
                    VMs_Total                 = $vms.Count
                    VMs_Allumees              = $vmsOn.Count
                    vCPU_Total                = $vcpuAll
                    vCPU_VMs_Allumees         = $vcpuOn
                    Ratio_vCPU_par_Coeur      = $ratioCpu
                    vRAM_Totale_Go            = Rnd $vramAll 0
                    vRAM_VMs_Allumees_Go      = Rnd $vramOn 0
                    Ratio_vRAM_RAM            = $ratioMem
                    Stockage_VMs_Provisionne_Go = Rnd $stoProv 0
                    Stockage_VMs_Consomme_Go  = Rnd $stoUsed 0
                    N1_RAM_OK                 = $n1Mem
                    N1_CPU_OK                 = $n1Cpu
                    RAM_Marge_N1_Go           = Rnd ($memCap - $maxHostMem - $memUsed) 0
                })

            if (-not $das.Enabled) { Add-Finding 'ALERTE' 'Cluster' $c.Name 'vSphere HA désactivé' }
            elseif (-not $das.AdmissionControlEnabled) { Add-Finding 'INFO' 'Cluster' $c.Name "HA : contrôle d'admission désactivé" }
            if (-not $drs.Enabled) { Add-Finding 'INFO' 'Cluster' $c.Name 'DRS désactivé' }
            if ($hsActive.Count -gt 1 -and -not $n1Mem) {
                Add-Finding 'ALERTE' 'Cluster' $c.Name "RAM consommée ($([int]$memUsed) Go) > capacité N+1 ($([int]($memCap - $maxHostMem)) Go) : perte d'un hôte non absorbable"
            }
            if ($hsActive.Count -le 1) { Add-Finding 'INFO' 'Cluster' $c.Name 'Moins de 2 hôtes actifs : pas de redondance' }
            if ($builds.Count -gt 1) { Add-Finding 'INFO' 'Cluster' $c.Name "Builds ESXi hétérogènes : $($builds -join ', ')" }
            if ($ratioCpu -gt 5) { Add-Finding 'ALERTE' 'Cluster' $c.Name "Ratio vCPU:cœur élevé (${ratioCpu}:1)" }
            if ($ratioMem -gt 1.2) { Add-Finding 'ALERTE' 'Cluster' $c.Name "RAM surallouée (vRAM/RAM = $ratioMem)" }
        }
    }

    #endregion

    #region ======================= Datastores =======================

    Invoke-Section 'Datastores' {
        foreach ($d in ($dsViews | Sort-Object Name)) {
            $s = $d.Summary
            $cap = [double]$s.Capacity; $free = [double]$s.FreeSpace; $unc = [double]$s.Uncommitted
            $used = $cap - $free; $prov = $used + $unc
            $usedPct = if ($cap) { Rnd (100 * $used / $cap) } else { $null }
            $provPct = if ($cap) { Rnd (100 * $prov / $cap) } else { $null }
            $id = "$($d.MoRef)"

            $dsRows.Add([pscustomobject]@{
                    Datacenter      = $dcByEntity[$id]
                    Datastore       = $d.Name
                    Type            = $s.Type
                    Accessible      = $s.Accessible
                    Statut          = "$($d.OverallStatus)"
                    Maintenance     = $s.MaintenanceMode
                    Capacite_Go     = Rnd ($cap / 1GB) 0
                    Utilise_Go      = Rnd ($used / 1GB) 0
                    Libre_Go        = Rnd ($free / 1GB) 0
                    Utilise_Pct     = $usedPct
                    Provisionne_Go  = Rnd ($prov / 1GB) 0
                    Provisionne_Pct = $provPct
                    Nb_Hotes        = @($d.Host).Count
                    Nb_VMs          = @($d.Vm).Count
                    VMFS_Version    = $d.Info.Vmfs.MajorVersion
                    LUN             = @($d.Info.Vmfs.Extent | ForEach-Object { $_.DiskName }) -join ', '
                    NFS_Cible       = if ($d.Info.Nas) { "$($d.Info.Nas.RemoteHost):$($d.Info.Nas.RemotePath)" } else { '' }
                    URL             = $s.Url
                })

            if (-not $s.Accessible) { Add-Finding 'CRITIQUE' 'Datastore' $d.Name 'Datastore inaccessible' }
            elseif ($usedPct -ge 95) { Add-Finding 'CRITIQUE' 'Datastore' $d.Name "Rempli à $usedPct % ($(Rnd ($free / 1GB) 0) Go libres)" }
            elseif ($usedPct -ge $DatastoreWarningPct) { Add-Finding 'ALERTE' 'Datastore' $d.Name "Rempli à $usedPct % ($(Rnd ($free / 1GB) 0) Go libres)" }
            if ($provPct -gt 150) { Add-Finding 'INFO' 'Datastore' $d.Name "Surprovisionnement thin : $provPct % de la capacité" }
        }
    }

    #endregion

    #region ======================= Performances historiques =======================

    if (-not $SkipPerfStats) {
        Invoke-Section "Performances historiques ($StatsDays j)" {
            $interval = if ($StatsDays -le 1) { 5 } elseif ($StatsDays -le 7) { 30 } elseif ($StatsDays -le 30) { 120 } else { 1440 }
            $intervalMs = $interval * 60 * 1000
            $start = (Get-Date).AddDays(-$StatsDays)
            $vmRowById = @{}; foreach ($r in $vmRows) { $vmRowById[$r._Id] = $r }

            # --- VMs ---
            $vmOn = @(Get-VM | Where-Object { $_.PowerState -eq 'PoweredOn' })
            if ($vmOn.Count) {
                $byKey = @{}
                $stats = Get-Stat -Entity $vmOn -Stat 'cpu.usage.average', 'mem.usage.average', 'cpu.ready.summation' `
                    -Start $start -IntervalMins $interval -ErrorAction SilentlyContinue
                foreach ($s in $stats) {
                    if ($s.Instance -ne '') { continue }
                    $k = "$($s.Entity.Id)|$($s.MetricId)"
                    if (-not $byKey.ContainsKey($k)) { $byKey[$k] = New-List }
                    $byKey[$k].Add([double]$s.Value)
                }
                foreach ($vm in $vmOn) {
                    $row = $vmRowById[$vm.Id]
                    $n = [math]::Max(1, $vm.NumCpu)
                    $cpu = Get-StatSummary $byKey["$($vm.Id)|cpu.usage.average"]
                    $mem = Get-StatSummary $byKey["$($vm.Id)|mem.usage.average"]
                    $rdy = Get-StatSummary (@($byKey["$($vm.Id)|cpu.ready.summation"]) | Where-Object { $null -ne $_ } | ForEach-Object { 100 * $_ / $intervalMs / $n })

                    $hints = @()
                    if ($rdy.Avg -ge 5) { $hints += 'Contention CPU (Ready élevé)'; Add-Finding 'ALERTE' 'Performance' $vm.Name "CPU Ready moyen $($rdy.Avg) % par vCPU" }
                    if ($cpu.P95 -ge 85) { $hints += 'CPU saturé (P95)' }
                    elseif ($vm.NumCpu -ge 4 -and $null -ne $cpu.P95 -and $cpu.P95 -lt 25) { $hints += 'vCPU probablement surdimensionnés' }
                    if ($mem.P95 -ge 90) { $hints += 'Mémoire active élevée' }
                    elseif ($vm.MemoryGB -ge 8 -and $null -ne $mem.P95 -and $mem.P95 -lt 20) { $hints += 'RAM possiblement surdimensionnée (à confirmer côté OS)' }

                    $perfRows.Add([pscustomobject]@{
                            Type              = 'VM'
                            Nom               = $vm.Name
                            Cluster           = if ($row) { $row.Cluster } else { '' }
                            vCPU_ou_Coeurs    = $vm.NumCpu
                            RAM_Go            = Rnd $vm.MemoryGB
                            CPU_Moy_Pct       = $cpu.Avg
                            CPU_P95_Pct       = $cpu.P95
                            CPU_Max_Pct       = $cpu.Max
                            RAM_Moy_Pct       = $mem.Avg
                            RAM_P95_Pct       = $mem.P95
                            RAM_Max_Pct       = $mem.Max
                            CPU_Ready_Moy_Pct = $rdy.Avg
                            CPU_Ready_Max_Pct = $rdy.Max
                            Echantillons      = $cpu.Count
                            Indications       = $hints -join ' ; '
                        })
                }
            }

            # --- Hôtes ---
            $hostsOk = @(Get-VMHost | Where-Object { $_.ConnectionState -in 'Connected', 'Maintenance' })
            if ($hostsOk.Count) {
                $byKey = @{}
                $stats = Get-Stat -Entity $hostsOk -Stat 'cpu.usage.average', 'mem.usage.average' `
                    -Start $start -IntervalMins $interval -ErrorAction SilentlyContinue
                foreach ($s in $stats) {
                    if ($s.Instance -ne '') { continue }
                    $k = "$($s.Entity.Id)|$($s.MetricId)"
                    if (-not $byKey.ContainsKey($k)) { $byKey[$k] = New-List }
                    $byKey[$k].Add([double]$s.Value)
                }
                foreach ($vh in $hostsOk) {
                    $cpu = Get-StatSummary $byKey["$($vh.Id)|cpu.usage.average"]
                    $mem = Get-StatSummary $byKey["$($vh.Id)|mem.usage.average"]
                    $hints = @()
                    if ($cpu.P95 -ge 80) { $hints += 'CPU hôte chargé (P95)'; Add-Finding 'ALERTE' 'Performance' $vh.Name "CPU P95 à $($cpu.P95) % sur $StatsDays j" }
                    if ($mem.P95 -ge 90) { $hints += 'RAM hôte chargée (P95)'; Add-Finding 'ALERTE' 'Performance' $vh.Name "RAM P95 à $($mem.P95) % sur $StatsDays j" }
                    $perfRows.Add([pscustomobject]@{
                            Type              = 'Hôte'
                            Nom               = $vh.Name
                            Cluster           = "$($vh.Parent)"
                            vCPU_ou_Coeurs    = $vh.NumCpu
                            RAM_Go            = Rnd $vh.MemoryTotalGB 0
                            CPU_Moy_Pct       = $cpu.Avg
                            CPU_P95_Pct       = $cpu.P95
                            CPU_Max_Pct       = $cpu.Max
                            RAM_Moy_Pct       = $mem.Avg
                            RAM_P95_Pct       = $mem.P95
                            RAM_Max_Pct       = $mem.Max
                            CPU_Ready_Moy_Pct = $null
                            CPU_Ready_Max_Pct = $null
                            Echantillons      = $cpu.Count
                            Indications       = $hints -join ' ; '
                        })
                }
            }
        }
    }

    #endregion

    #region ======================= Configuration détaillée des hôtes =======================

    if (-not $SkipHostDetails) {
        Invoke-Section 'Configuration des hôtes (services, réseau, stockage)' {
            $vmhosts = @(Get-VMHost | Where-Object { $_.ConnectionState -in 'Connected', 'Maintenance' } | Sort-Object Name)
            $i = 0
            foreach ($vh in $vmhosts) {
                $i++
                Write-Progress -Activity 'Configuration des hôtes' -Status $vh.Name -PercentComplete (100 * $i / $vmhosts.Count)
                try {
                    $svc     = Get-VMHostService -VMHost $vh
                    $ntpSvc  = $svc | Where-Object { $_.Key -eq 'ntpd' }
                    $sshSvc  = $svc | Where-Object { $_.Key -eq 'TSM-SSH' }
                    $shSvc   = $svc | Where-Object { $_.Key -eq 'TSM' }
                    $ntpSrv  = @(Get-VMHostNtpServer -VMHost $vh) -join ', '
                    $net     = Get-VMHostNetwork -VMHost $vh
                    $syslog  = @(Get-VMHostSysLogServer -VMHost $vh -ErrorAction SilentlyContinue | ForEach-Object { "$($_.Host):$($_.Port)" }) -join ', '
                    $lock    = $hostById[$vh.Id].Config.LockdownMode
                    $power   = $vh.ExtensionData.Config.PowerSystemInfo.CurrentPolicy.ShortName

                    $hostCfgRows.Add([pscustomobject]@{
                            Hote              = $vh.Name
                            Cluster           = "$($vh.Parent)"
                            NTP_Serveurs      = $ntpSrv
                            NTP_Demarre       = $ntpSvc.Running
                            NTP_Politique     = $ntpSvc.Policy
                            SSH_Demarre       = $sshSvc.Running
                            SSH_Politique     = $sshSvc.Policy
                            Shell_Demarre     = $shSvc.Running
                            Lockdown          = "$lock"
                            DNS               = @($net.DnsAddress) -join ', '
                            Domaine           = $net.DomainName
                            Syslog            = $syslog
                            Politique_Energie = $power
                            HyperThreading    = $vh.HyperthreadingActive
                        })

                    if (-not $ntpSrv) { Add-Finding 'ALERTE' 'Hôte' $vh.Name 'Aucun serveur NTP configuré' }
                    elseif (-not $ntpSvc.Running) { Add-Finding 'ALERTE' 'Hôte' $vh.Name 'Service NTP arrêté' }
                    if ($sshSvc.Running) { Add-Finding 'INFO' 'Sécurité' $vh.Name 'SSH démarré' }
                    if ($shSvc.Running) { Add-Finding 'INFO' 'Sécurité' $vh.Name 'ESXi Shell démarré' }
                    if (-not $syslog) { Add-Finding 'INFO' 'Hôte' $vh.Name 'Aucun serveur syslog distant' }

                    foreach ($p in @(Get-VMHostNetworkAdapter -VMHost $vh -Physical)) {
                        $speed = [int]$p.BitRatePerSec
                        $nicRows.Add([pscustomobject]@{
                                Hote         = $vh.Name
                                Carte        = $p.Name
                                MAC          = $p.Mac
                                Vitesse_Mbps = $speed
                                Lien         = if ($speed -gt 0) { 'Up' } else { 'Down' }
                                FullDuplex   = $p.FullDuplex
                            })
                    }

                    foreach ($k in @(Get-VMHostNetworkAdapter -VMHost $vh -VMKernel)) {
                        $vmkRows.Add([pscustomobject]@{
                                Hote       = $vh.Name
                                Interface  = $k.Name
                                IP         = $k.IP
                                Masque     = $k.SubnetMask
                                DHCP       = $k.DhcpEnabled
                                PortGroup  = $k.PortGroupName
                                MTU        = $k.Mtu
                                Management = $k.ManagementTrafficEnabled
                                vMotion    = $k.VMotionEnabled
                                vSAN       = $k.VsanTrafficEnabled
                                FT         = $k.FaultToleranceLoggingEnabled
                            })
                    }

                    foreach ($b in @(Get-VMHostHba -VMHost $vh)) {
                        $wwn = ''
                        if ($b.Type -eq 'FibreChannel' -and $b.PortWorldWideName) {
                            $wwn = ('{0:x16}' -f $b.PortWorldWideName) -replace '(..)(?!$)', '$1:'
                        }
                        $hbaRows.Add([pscustomobject]@{
                                Hote    = $vh.Name
                                Device  = $b.Device
                                Type    = "$($b.Type)"
                                Modele  = $b.Model
                                Pilote  = $b.Driver
                                Statut  = $b.Status
                                WWPN    = $wwn
                                IQN     = $b.IScsiName
                            })
                    }

                    foreach ($sw in @(Get-VirtualSwitch -VMHost $vh -Standard)) {
                        $vssRows.Add([pscustomobject]@{
                                Hote     = $vh.Name
                                vSwitch  = $sw.Name
                                MTU      = $sw.Mtu
                                Uplinks  = @($sw.Nic) -join ', '
                            })
                    }
                    foreach ($pg in @(Get-VirtualPortGroup -VMHost $vh -Standard)) {
                        $pgRows.Add([pscustomobject]@{
                                Type      = 'Standard'
                                Hote      = $vh.Name
                                Switch    = $pg.VirtualSwitchName
                                PortGroup = $pg.Name
                                VLAN      = $pg.VLanId
                                Nb_Ports  = $null
                                Binding   = $null
                            })
                    }
                }
                catch {
                    $msg = "Hôte $($vh.Name) : $($_.Exception.Message)"
                    Write-Warning $msg
                    $CollectErrors.Add($msg)
                }
            }
            Write-Progress -Activity 'Configuration des hôtes' -Completed

            # Commutateurs distribués
            if (Get-Command Get-VDSwitch -ErrorAction SilentlyContinue) {
                foreach ($vds in @(Get-VDSwitch)) {
                    $vdsRows.Add([pscustomobject]@{
                            vDS          = $vds.Name
                            Version      = $vds.Version
                            MTU          = $vds.Mtu
                            Nb_Uplinks   = $vds.NumUplinkPorts
                            Nb_Hotes     = @($vds.ExtensionData.Summary.HostMember).Count
                            Decouverte   = "$($vds.LinkDiscoveryProtocol) / $($vds.LinkDiscoveryProtocolOperation)"
                        })
                }
                foreach ($pg in @(Get-VDPortgroup | Where-Object { -not $_.IsUplink })) {
                    $pgRows.Add([pscustomobject]@{
                            Type      = 'Distribué'
                            Hote      = ''
                            Switch    = $pg.VDSwitch.Name
                            PortGroup = $pg.Name
                            VLAN      = if ($pg.VlanConfiguration) { "$($pg.VlanConfiguration)" } else { 'Aucun' }
                            Nb_Ports  = $pg.NumPorts
                            Binding   = "$($pg.PortBinding)"
                        })
                }
            }
        }
    }

    #endregion

    #region ======================= Alarmes déclenchées =======================

    Invoke-Section 'Alarmes déclenchées' {
        $alarmName = @{}; $entityName = @{}; $seen = @{}
        $roots = @(Get-View $si.Content.RootFolder -Property TriggeredAlarmState) + @(Get-View -ViewType Datacenter -Property TriggeredAlarmState)
        foreach ($r in $roots) {
            foreach ($a in @($r.TriggeredAlarmState)) {
                if (-not $a -or $seen.ContainsKey($a.Key)) { continue }
                $seen[$a.Key] = $true
                $ak = "$($a.Alarm)"; $ek = "$($a.Entity)"
                if (-not $alarmName.ContainsKey($ak)) { $alarmName[$ak] = (Get-View $a.Alarm -Property Info).Info.Name }
                if (-not $entityName.ContainsKey($ek)) { $entityName[$ek] = (Get-View $a.Entity -Property Name).Name }
                $alarmRows.Add([pscustomobject]@{
                        Severite   = "$($a.OverallStatus)"
                        Alarme     = $alarmName[$ak]
                        Objet      = $entityName[$ek]
                        Type_Objet = $a.Entity.Type
                        Date       = $a.Time.ToLocalTime()
                        Acquittee  = $a.Acknowledged
                    })
                $sev = if ("$($a.OverallStatus)" -eq 'red') { 'CRITIQUE' } else { 'ALERTE' }
                Add-Finding $sev 'Alarme vCenter' $entityName[$ek] $alarmName[$ak]
            }
        }
    }

    #endregion

    #region ======================= Licences =======================

    Invoke-Section 'Licences' {
        $lm = Get-View $si.Content.LicenseManager
        foreach ($l in @($lm.Licenses)) {
            $exp = ($l.Properties | Where-Object { $_.Key -eq 'expirationDate' }).Value
            $key = if ($l.LicenseKey -and $l.LicenseKey.Length -ge 5) { 'XXXXX-XXXXX-XXXXX-XXXXX-' + $l.LicenseKey.Substring($l.LicenseKey.Length - 5) } else { $l.LicenseKey }
            $licenseRows.Add([pscustomobject]@{
                    Licence    = $l.Name
                    Edition    = $l.EditionKey
                    Cle        = $key
                    Total      = $l.Total
                    Utilise    = $l.Used
                    Unite      = $l.CostUnit
                    Expiration = if ($exp) { $exp } else { 'Permanente' }
                })
            if ($exp -and ([datetime]$exp - (Get-Date)).TotalDays -lt 60) {
                Add-Finding 'ALERTE' 'Licence' $l.Name "Expire le $(([datetime]$exp).ToString('dd/MM/yyyy'))"
            }
            if ($l.Total -gt 0 -and $l.Used -gt $l.Total) {
                Add-Finding 'CRITIQUE' 'Licence' $l.Name "Utilisation $($l.Used) > total $($l.Total)"
            }
        }
    }

    #endregion

    #region ======================= Synthèse =======================

    Write-Step 'Synthèse'
    $vmsOnAll = @($vmRows | Where-Object { $_.Etat -eq 'poweredOn' })
    $tCores = 0; $tMem = 0; $tCpuCap = 0; $tCpuUsed = 0; $tMemUsed = 0
    foreach ($x in $hostRows) { $tCores += $x.Coeurs; $tMem += $x.RAM_Go; $tCpuCap += $x._CpuCapMhz; $tCpuUsed += $x._CpuUsedMhz; $tMemUsed += $x._MemUsedGB }
    $tVcpu = 0; $tVram = 0; $tVcpuOn = 0; $tVramOn = 0
    foreach ($x in $vmRows) { $tVcpu += $x.vCPU; $tVram += $x.RAM_Go }
    foreach ($x in $vmsOnAll) { $tVcpuOn += $x.vCPU; $tVramOn += $x.RAM_Go }
    $dsCap = 0; $dsFree = 0; $dsProv = 0
    foreach ($x in $dsRows) { $dsCap += $x.Capacite_Go; $dsFree += $x.Libre_Go; $dsProv += $x.Provisionne_Go }

    $summary = [ordered]@{
        'Date de collecte'                 = $startTime.ToString('dd/MM/yyyy HH:mm')
        'vCenter'                          = $vCenter
        'Produit'                          = $about.FullName
        'Version / Build'                  = "$($about.Version) / $($about.Build)"
        'Compte utilisé'                   = $Credential.UserName
        'Datacenters'                      = $dcViews.Count
        'Clusters'                         = $clusterViews.Count
        'Hôtes ESXi'                       = $hostRows.Count
        'VMs (hors templates)'             = $vmRows.Count
        'VMs allumées'                     = $vmsOnAll.Count
        'Templates'                        = $script:templateCount
        'Snapshots'                        = $snapRows.Count
        'Cœurs physiques'                  = $tCores
        'RAM physique (Go)'                = Rnd $tMem 0
        'CPU utilisé (%)'                  = if ($tCpuCap) { Rnd (100 * $tCpuUsed / $tCpuCap) } else { $null }
        'RAM utilisée (%)'                 = if ($tMem) { Rnd (100 * $tMemUsed / $tMem) } else { $null }
        'vCPU provisionnés (tous / allumés)' = "$tVcpu / $tVcpuOn"
        'vRAM provisionnée Go (toutes / allumées)' = "$(Rnd $tVram 0) / $(Rnd $tVramOn 0)"
        'Ratio vCPU:cœur (VMs allumées)'   = if ($tCores) { Rnd ($tVcpuOn / $tCores) 2 } else { $null }
        'Ratio vRAM:RAM (VMs allumées)'    = if ($tMem) { Rnd ($tVramOn / $tMem) 2 } else { $null }
        'Datastores'                       = $dsRows.Count
        'Stockage capacité (To)'           = Rnd ($dsCap / 1024) 2
        'Stockage libre (To)'              = Rnd ($dsFree / 1024) 2
        'Stockage provisionné (To)'        = Rnd ($dsProv / 1024) 2
        'Alarmes déclenchées'              = $alarmRows.Count
        'Points CRITIQUES'                 = @($findings | Where-Object { $_.Severite -eq 'CRITIQUE' }).Count
        'Points ALERTE'                    = @($findings | Where-Object { $_.Severite -eq 'ALERTE' }).Count
        'Points INFO'                      = @($findings | Where-Object { $_.Severite -eq 'INFO' }).Count
        'Historique de performances'       = if ($SkipPerfStats) { 'Non collecté' } else { "$StatsDays jours" }
        'Erreurs de collecte'              = $CollectErrors.Count
    }
    foreach ($k in $summary.Keys) { $summaryRows.Add([pscustomobject]@{ Indicateur = $k; Valeur = $summary[$k] }) }

    $sevOrder = @{ 'CRITIQUE' = 0; 'ALERTE' = 1; 'INFO' = 2 }
    $findingsSorted = @($findings | Sort-Object { $sevOrder[$_.Severite] }, Categorie, Objet)

    #endregion

    #region ======================= Exports =======================

    Write-Step 'Export des résultats'
    $sheets = [ordered]@{
        'Synthese'         = $summaryRows
        'Points_attention' = $findingsSorted
        'Clusters'         = $clusterRows
        'Hotes'            = $hostRows
        'VMs'              = $vmRows
        'Performances'     = $perfRows
        'Datastores'       = $dsRows
        'Snapshots'        = $snapRows
        'Alarmes'          = $alarmRows
        'Licences'         = $licenseRows
        'Hotes_Config'     = $hostCfgRows
        'Hotes_NIC'        = $nicRows
        'Hotes_VMkernel'   = $vmkRows
        'Hotes_HBA'        = $hbaRows
        'vSwitch_Standard' = $vssRows
        'PortGroups'       = $pgRows
        'vDS'              = $vdsRows
        'Erreurs_collecte' = @($CollectErrors | ForEach-Object { [pscustomobject]@{ Erreur = $_ } })
    }

    $hasExcel = [bool](Get-Module -ListAvailable -Name ImportExcel)
    $xlsx = Join-Path $runFolder "vSphere_$($safeName)_$stamp.xlsx"
    if ($hasExcel) { Import-Module ImportExcel -ErrorAction SilentlyContinue }

    $csvFolder = Join-Path $runFolder 'CSV'
    New-Item -ItemType Directory -Path $csvFolder -Force | Out-Null

    foreach ($name in $sheets.Keys) {
        $data = @($sheets[$name] | Select-Object * -ExcludeProperty _*)
        if ($data.Count -eq 0) { continue }
        $data | Export-Csv -Path (Join-Path $csvFolder "$name.csv") -NoTypeInformation -Delimiter ';' -Encoding UTF8
        if ($hasExcel) {
            try {
                $data | Export-Excel -Path $xlsx -WorksheetName $name -TableName "T_$name" -TableStyle Medium2 `
                    -AutoSize -FreezeTopRow -BoldTopRow
            }
            catch { Write-Warning "Excel ($name) : $($_.Exception.Message)" }
        }
    }
    if (-not $hasExcel) {
        Write-Warning "Module ImportExcel absent : pas de .xlsx (CSV générés). Pour l'activer : Install-Module ImportExcel -Scope CurrentUser"
    }

    # ---------- Rapport HTML ----------
    $css = @'
<style>
body{font-family:"Segoe UI",Arial,sans-serif;font-size:12px;margin:20px;color:#222;background:#fff}
h1{color:#1d4f7a;margin-bottom:4px} .sub{color:#666;margin-top:0}
h2{color:#1d4f7a;border-bottom:2px solid #1d4f7a;padding-bottom:3px;margin-top:32px}
.toc a{margin-right:14px;color:#1d4f7a;text-decoration:none;font-weight:600}
.wrap{overflow-x:auto;max-height:600px;overflow-y:auto;border:1px solid #ddd}
table{border-collapse:collapse;width:100%}
th{background:#1d4f7a;color:#fff;padding:5px 7px;text-align:left;position:sticky;top:0;white-space:nowrap}
td{padding:4px 7px;border-bottom:1px solid #e3e3e3;white-space:nowrap}
tr:nth-child(even){background:#f5f7f9}
tr.crit td{background:#f8d0d0} tr.warn td{background:#fde9c4} tr.info td{background:#e1effb}
</style>
'@
    $htmlSections = [ordered]@{
        'Synthese'         = 'Synthèse'
        'Points_attention' = "Points d'attention"
        'Clusters'         = 'Clusters'
        'Hotes'            = 'Hôtes ESXi'
        'Datastores'       = 'Datastores'
        'Performances'     = "Performances ($StatsDays j)"
        'Snapshots'        = 'Snapshots'
        'Alarmes'          = 'Alarmes déclenchées'
        'Licences'         = 'Licences'
        'Hotes_Config'     = 'Configuration des hôtes'
        'Hotes_VMkernel'   = 'Interfaces VMkernel'
        'Hotes_NIC'        = 'Cartes physiques'
        'PortGroups'       = 'Port groups'
        'VMs'              = 'Machines virtuelles'
        'Erreurs_collecte' = 'Erreurs de collecte'
    }
    $toc = ($htmlSections.Keys | ForEach-Object { "<a href='#$_'>$($htmlSections[$_])</a>" }) -join ''
    $body = New-Object System.Text.StringBuilder
    [void]$body.Append("<h1>État de l'infrastructure vSphere</h1><p class='sub'>$vCenter - $($about.FullName) - collecte du $($startTime.ToString('dd/MM/yyyy HH:mm'))</p>")
    [void]$body.Append("<div class='toc'>$toc</div>")
    foreach ($k in $htmlSections.Keys) {
        [void]$body.Append("<h2 id='$k'>$($htmlSections[$k])</h2>")
        $data = @($sheets[$k] | Select-Object * -ExcludeProperty _*)
        if ($data.Count -eq 0) { [void]$body.Append('<p><i>Aucune donnée.</i></p>'); continue }
        $frag = ($data | ConvertTo-Html -Fragment | Out-String)
        $frag = $frag -replace '<tr><td>CRITIQUE</td>', '<tr class="crit"><td>CRITIQUE</td>' `
            -replace '<tr><td>ALERTE</td>', '<tr class="warn"><td>ALERTE</td>' `
            -replace '<tr><td>INFO</td>', '<tr class="info"><td>INFO</td>' `
            -replace '<tr><td>red</td>', '<tr class="crit"><td>red</td>' `
            -replace '<tr><td>yellow</td>', '<tr class="warn"><td>yellow</td>'
        [void]$body.Append("<div class='wrap'>$frag</div>")
    }
    $htmlPath = Join-Path $runFolder "vSphere_$($safeName)_$stamp.html"
    $html = "<!DOCTYPE html><html><head><meta charset='utf-8'><title>vSphere - $vCenter</title>$css</head><body>$($body.ToString())</body></html>"
    [System.IO.File]::WriteAllText($htmlPath, $html, (New-Object System.Text.UTF8Encoding($true)))

    #endregion

    $duration = (Get-Date) - $startTime
    Write-Host ''
    Write-Host '==================== Terminé ====================' -ForegroundColor Green
    $summaryRows | Format-Table -AutoSize | Out-String | Write-Host
    Write-Host "Durée            : $([int]$duration.TotalMinutes) min $($duration.Seconds) s"
    Write-Host "Dossier          : $runFolder"
    Write-Host "Rapport HTML     : $htmlPath"
    if ($hasExcel) { Write-Host "Classeur Excel   : $xlsx" }
    if ($CollectErrors.Count) { Write-Warning "$($CollectErrors.Count) erreur(s) de collecte, voir l'onglet Erreurs_collecte." }

    if (-not $NoOpen) { Invoke-Item $htmlPath }
}
catch {
    Write-Error "Erreur bloquante : $($_.Exception.Message)"
}
finally {
    if ($global:DefaultVIServers) { Disconnect-VIServer -Server * -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    Stop-Transcript | Out-Null
}

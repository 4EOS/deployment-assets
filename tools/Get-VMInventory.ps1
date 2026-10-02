<#
.SYNOPSIS
  Enumerates every VM on a hypervisor host and emits one row per VM.

.DESCRIPTION
  Closes the "which VMs exist on this host?" gap for sites that no API can see.
  Three modes, all read-only:

    -Mode VCenter  (default)  vSphere / vCenter or standalone ESXi.
                              Uses VMware PowerCLI when available, otherwise the
                              vSphere Automation REST API (/api then /rest).
    -Mode HyperV              Any Windows host with the Hyper-V role. Uses CIM,
                              no module install required.

  Output is CSV and/or JSON on stdout or to -OutFile. Progress/log lines go to
  stderr so RMM parsers never choke on them.

.PARAMETER Server
  vCenter or ESXi FQDN/IP, e.g. 10.1.0.120 (PFS vCenter) or 10.5.105.10 (Plymouth).

.PARAMETER Site
  Free-text site label stamped on every row, e.g. "Plymouth".

.EXAMPLE
  # main campus vCenter, all hosts, CSV for the inventory workbook
  .\Get-VMInventory.ps1 -Server 10.1.0.120 -Site "Main campus" -User 'administrator@vsphere.local' `
      -Password $env:VC_PW -OutFile .\pfs-main-vms.csv

.EXAMPLE
  # a remote standalone host, JSON back to the RMM
  .\Get-VMInventory.ps1 -Server 10.5.105.10 -Site Plymouth -User root -Password $env:ESXI_PW -Json

.EXAMPLE
  # Hyper-V host, no VMware anything
  .\Get-VMInventory.ps1 -Mode HyperV -Site "Fort Wayne" -OutFile .\ftwayne-vms.csv

.NOTES
  Exit codes: 0 ok · 2 no credentials · 3 connect/query failure · 4 nothing found.
  Secrets come from parameters or environment variables, never the script body.
#>
[CmdletBinding(DefaultParameterSetName = 'VCenter')]
param(
    [Parameter(ParameterSetName = 'VCenter', Position = 0)]
    [string]$Server,

    [Parameter(ParameterSetName = 'HyperV')]
    [switch]$Mode_HyperV_unused,   # placeholder to keep param sets tidy

    [ValidateSet('VCenter', 'HyperV')]
    [string]$Mode = 'VCenter',

    [string]$Site = 'unknown',
    [string]$User,
    [string]$Password = $env:VC_PW,
    [switch]$IgnoreCertificate,
    [switch]$Json,
    [string]$OutFile
)

$ErrorActionPreference = 'Stop'
function Log($m) { [Console]::Error.WriteLine("[Get-VMInventory] $m") }

if ($Mode -eq 'HyperV') {
    Log "mode HyperV on $env:COMPUTERNAME"
    try {
        $vms = Get-CimInstance -Namespace root/virtualization/v2 -ClassName Msvm_ComputerSystem -ErrorAction Stop |
            Where-Object { $_.Caption -eq 'Virtual Machine' }
    } catch {
        Log "CIM query failed: $($_.Exception.Message)"
        exit 3
    }
    $rows = foreach ($v in $vms) {
        $mem = Get-CimInstance -Namespace root/virtualization/v2 -ClassName Msvm_MemorySettingData -Filter "InstanceID LIKE '$($v.Name)%'" -ErrorAction SilentlyContinue
        [pscustomobject]@{
            site = $Site; host = $env:COMPUTERNAME; role = 'Hyper-V'
            vm = $v.ElementName
            power = switch ($v.EnabledState) { 2 { 'on' } 3 { 'off' } 32768 { 'paused' } 32769 { 'saved' } default { "state$($v.EnabledState)" } }
            vcpu = ($v.NumberOfProcessors)
            memoryGB = if ($mem) { [math]::Round($mem[0].VirtualQuantity / 1GB, 1) } else { $null }
            guestOS = $null; ip = $null; hostName = $null
        }
    }
} else {
    if (-not $Server) { Log "no -Server given"; exit 2 }
    Log "mode VCenter against $Server (site: $Site)"
    if (-not $User -or -not $Password) { Log "credentials required (-User and -Password/dollar-VC_PW)"; exit 2 }
    if ($IgnoreCertificate) { Log 'certificate validation skipped (lab/self-signed)'
        if ($PSVersionTable.PSVersion.Major -ge 6) { $PSDefaultParameterValues['Invoke-RestMethod:SkipCertificateCheck'] = $true }
        else { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } }
    }

    $powerCli = Get-Module -ListAvailable -Name VMware.PowerCLI
    if ($powerCli) {
        Log "PowerCLI $($powerCli[0].Version) found - using it"
        Import-Module VMware.VimAutomation.Core -ErrorAction Stop
        $cred = New-Object System.Management.Automation.PSCredential($User, (ConvertTo-SecureString $Password -AsPlainText -Force))
        try { $null = Connect-VIServer -Server $Server -Credential $cred -ErrorAction Stop }
        catch { Log "Connect-VIServer failed: $($_.Exception.Message)"; exit 3 }
        $rows = Get-VM | ForEach-Object {
            [pscustomobject]@{
                site = $Site
                host = if ($_.VMHost) { $_.VMHost.Name } else { '(unknown)' }
                role = if ($_.VMHost -and $_.VMHost.Parent) { $_.VMHost.Parent.Name } else { '' }
                vm = $_.Name
                power = "$($_.PowerState)".ToLower()
                vcpu = $_.NumCpu
                memoryGB = [math]::Round($_.MemoryGB, 1)
                guestOS = $_.Guest.OSFullName
                ip = ($_.Guest.IPAddress -join ', ')
                hostName = $_.Guest.HostName
            }
        }
        Disconnect-VIServer -Server $Server -Confirm:$false -ErrorAction SilentlyContinue
    } else {
        Log 'no PowerCLI - falling back to the vSphere Automation REST API'
        $base = $null
        foreach ($p in '/api', '/rest') {           # 7.0+ uses /api, 6.5/6.7 use /rest
            try {
                $sess = Invoke-RestMethod -Method Post -Uri "https://$Server$p/session" `
                        -Credential (New-Object System.Management.Automation.PSCredential($User, (ConvertTo-SecureString $Password -AsPlainText -Force))) `
                        -SkipCertificateCheck -ErrorAction Stop
                $base = $p; $sessionId = $sess; break
            } catch { Log "  $p not available: $($_.Exception.Message.Split([Environment]::NewLine)[0])" }
        }
        if (-not $base) { Log 'REST API unreachable on this host (ESXi 6.5 has no /rest) - use PowerCLI or the host client'; exit 3 }
        $hdr = @{ 'vmware-api-session-id' = $sessionId }
        $vmList = (Invoke-RestMethod -Uri "https://$Server$base/vcenter/vm" -Headers $hdr -SkipCertificateCheck).value
        $hosts = @{}
        try {
            foreach ($h in (Invoke-RestMethod -Uri "https://$Server$base/vcenter/host" -Headers $hdr -SkipCertificateCheck).value) {
                $hosts[$h.host] = $h.name
            }
        } catch { Log 'host list not available (standalone ESXi?)' }
        $rows = foreach ($vm in $vmList) {
            $ip = $null
            try {
                $nics = (Invoke-RestMethod -Uri "https://$Server$base/vcenter/vm/$($vm.vm)/guest/networking/interfaces" -Headers $hdr -SkipCertificateCheck).value
                $ip = (($nics | ForEach-Object { $_.ip.ip_addresses } | Where-Object { $_ -notmatch ':' } ) -join ', ')
            } catch { }
            [pscustomobject]@{
                site = $Site
                host = if ($vm.host) { $hosts[$vm.host] } else { '(unknown)' }
                role = ''
                vm = $vm.name
                power = "$($vm.power_state)".ToLower()
                vcpu = $vm.cpu_count
                memoryGB = [math]::Round($vm.memory_size_MiB / 1024, 1)
                guestOS = $vm.guest_OS    # only populated by some builds
                ip = $ip
                hostName = $null
            }
        }
        Invoke-RestMethod -Method Delete -Uri "https://$Server$base/session" -Headers $hdr -SkipCertificateCheck -ErrorAction SilentlyContinue | Out-Null
    }
}

$rows = @($rows)
if ($rows.Count -eq 0) { Log "no VMs returned from $Server"; exit 4 }
Log "$($rows.Count) VMs"

if ($Json) { $rows | ConvertTo-Json -Depth 4 }
if ($OutFile) { $rows | Export-Csv -NoTypeInformation -Path $OutFile; Log "wrote $OutFile" }
if (-not $Json -and -not $OutFile) { $rows | Format-Table -AutoSize | Out-String | Write-Output }
exit 0

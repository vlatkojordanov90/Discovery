<#
.SYNOPSIS
  Windows Server 2019 -> CSV-only export:
   - Hardware summary
   - Hyper-V VM summary + VM list
   - Running services list
   - IIS details (if installed): sites/bindings + app pools

.OUTPUT
  A single CSV file containing multiple sections. Each row has:
    Section, ItemType, Name, Key, Value, Extra1, Extra2

.USAGE
  .\Export-Server2019-ReportCsv.ps1 -OutCsv "C:\Temp\server2019_report.csv"
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [string]$OutCsv
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$rows = New-Object System.Collections.Generic.List[object]

function Add-Row {
  param(
    [string]$Section,
    [string]$ItemType,
    [string]$Name,
    [string]$Key,
    [string]$Value,
    [string]$Extra1 = "",
    [string]$Extra2 = ""
  )
  $rows.Add([pscustomobject]@{
    Section  = $Section
    ItemType = $ItemType
    Name     = $Name
    Key      = $Key
    Value    = $Value
    Extra1   = $Extra1
    Extra2   = $Extra2
  }) | Out-Null
}

# ---------------- Hardware ----------------
$cs   = Get-CimInstance Win32_ComputerSystem
$os   = Get-CimInstance Win32_OperatingSystem
$bios = Get-CimInstance Win32_BIOS
$cpu  = Get-CimInstance Win32_Processor
$uptime = (Get-Date) - $os.LastBootUpTime

Add-Row "Hardware" "Summary" $env:COMPUTERNAME "Manufacturer" $cs.Manufacturer
Add-Row "Hardware" "Summary" $env:COMPUTERNAME "Model" $cs.Model
Add-Row "Hardware" "Summary" $env:COMPUTERNAME "SerialNumber" $bios.SerialNumber
Add-Row "Hardware" "Summary" $env:COMPUTERNAME "OS" $os.Caption
Add-Row "Hardware" "Summary" $env:COMPUTERNAME "OSVersion" $os.Version
Add-Row "Hardware" "Summary" $env:COMPUTERNAME "InstallDate" ($os.InstallDate.ToString("s"))
Add-Row "Hardware" "Summary" $env:COMPUTERNAME "LastBootUpTime" ($os.LastBootUpTime.ToString("s"))
Add-Row "Hardware" "Summary" $env:COMPUTERNAME "Uptime" ("{0}d {1}h {2}m" -f $uptime.Days, $uptime.Hours, $uptime.Minutes)

Add-Row "Hardware" "CPU" $env:COMPUTERNAME "CPU_Name" (($cpu | Select-Object -First 1).Name)
Add-Row "Hardware" "CPU" $env:COMPUTERNAME "CPU_Sockets" (@($cpu).Count)
Add-Row "Hardware" "CPU" $env:COMPUTERNAME "CPU_CoresTotal" (($cpu | Measure-Object NumberOfCores -Sum).Sum)
Add-Row "Hardware" "CPU" $env:COMPUTERNAME "CPU_LogicalTotal" (($cpu | Measure-Object NumberOfLogicalProcessors -Sum).Sum)

Add-Row "Hardware" "Memory" $env:COMPUTERNAME "RAM_GB" ([math]::Round(($cs.TotalPhysicalMemory / 1GB), 2))

# Physical disks
Get-CimInstance Win32_DiskDrive | ForEach-Object {
  Add-Row "Hardware" "Disk" $_.Model "SizeGB" ([math]::Round(($_.Size / 1GB), 2)) $_.InterfaceType ($_.SerialNumber)
}

# Volumes
try {
  Get-Volume | Where-Object DriveLetter | ForEach-Object {
    Add-Row "Hardware" "Volume" ("$($_.DriveLetter):") "SizeGB" ([math]::Round(($_.Size / 1GB), 2)) `
      ("FreeGB=" + [math]::Round(($_.SizeRemaining / 1GB), 2)) ("FS=" + $_.FileSystem)
  }
} catch { }

# NICs
Get-CimInstance Win32_NetworkAdapter |
  Where-Object { $_.PhysicalAdapter -eq $true -and $_.NetEnabled -eq $true } |
  ForEach-Object {
    $speed = if ($_.Speed) { [math]::Round(($_.Speed / 1MB), 0) } else { "" }
    Add-Row "Hardware" "NIC" $_.Name "MAC" $_.MACAddress ("SpeedMbps=" + $speed) ""
  }

# ---------------- Hyper-V ----------------
if (Get-Command Get-VM -ErrorAction SilentlyContinue) {
  $vms = Get-VM
  Add-Row "HyperV" "Summary" $env:COMPUTERNAME "TotalVMs" (@($vms).Count)
  Add-Row "HyperV" "Summary" $env:COMPUTERNAME "RunningVMs" (@($vms | Where-Object State -eq 'Running').Count)
  Add-Row "HyperV" "Summary" $env:COMPUTERNAME "OffVMs" (@($vms | Where-Object State -eq 'Off').Count)

  foreach ($vm in ($vms | Sort-Object Name)) {
    $ips = @()
    try { $ips = (Get-VMNetworkAdapter -VMName $vm.Name).IPAddresses } catch { }
    $ipv4 = ($ips | Where-Object { $_ -and $_ -notmatch ':' }) -join ';'

    Add-Row "HyperV" "VM" $vm.Name "State" $vm.State `
      ("CPU=" + $vm.ProcessorCount) `
      ("StartupMB=" + [math]::Round(($vm.MemoryStartup / 1MB), 0) + "; DynMem=" + $vm.DynamicMemoryEnabled + "; IPs=" + $ipv4)
  }
} else {
  Add-Row "HyperV" "Summary" $env:COMPUTERNAME "Available" "False" "Hyper-V module/cmdlets not found" ""
}

# ---------------- Services ----------------
Get-Service | Where-Object Status -eq 'Running' | Sort-Object DisplayName | ForEach-Object {
  Add-Row "Services" "Service" $_.DisplayName "Name" $_.Name ("StartType=" + $_.StartType) ("Status=" + $_.Status)
}

# ---------------- IIS ----------------
# Check if IIS role installed
$iisInstalled = $false
try {
  $iisFeature = Get-WindowsFeature -Name Web-Server
  if ($iisFeature -and $iisFeature.Installed) { $iisInstalled = $true }
} catch { }

if (-not $iisInstalled) {
  Add-Row "IIS" "Summary" $env:COMPUTERNAME "Installed" "False" "" ""
} else {
  Add-Row "IIS" "Summary" $env:COMPUTERNAME "Installed" "True" "" ""

  # IIS version from registry
  try {
    $iisVer = Get-ItemProperty "HKLM:\Software\Microsoft\InetStp" -ErrorAction Stop
    Add-Row "IIS" "Version" $env:COMPUTERNAME "Major" $iisVer.MajorVersion
    Add-Row "IIS" "Version" $env:COMPUTERNAME "Minor" $iisVer.MinorVersion
    Add-Row "IIS" "Version" $env:COMPUTERNAME "Build" $iisVer.BuildNumber
  } catch { }

  Import-Module WebAdministration -ErrorAction SilentlyContinue

  if (Get-Command Get-Website -ErrorAction SilentlyContinue) {
    # Sites
    foreach ($site in (Get-Website | Sort-Object Name)) {
      Add-Row "IIS" "Site" $site.Name "State" $site.State ("Path=" + $site.PhysicalPath) ("AppPool=" + $site.ApplicationPool)

      # Bindings as separate rows
      try {
        Get-WebBinding -Name $site.Name | ForEach-Object {
          Add-Row "IIS" "Binding" $site.Name "Protocol" $_.protocol ("Binding=" + $_.bindingInformation) ("SslFlags=" + $_.sslFlags)
        }
      } catch { }
    }

    # App Pools
    foreach ($ap in (Get-ChildItem IIS:\AppPools | Sort-Object Name)) {
      $state = ""
      try { $state = (Get-WebAppPoolState -Name $ap.Name).Value } catch { }
      Add-Row "IIS" "AppPool" $ap.Name "State" $state `
        ("Runtime=" + $ap.managedRuntimeVersion + "; Pipeline=" + $ap.managedPipelineMode) `
        ("Identity=" + $ap.processModel.identityType + "; StartMode=" + $ap.startMode + "; AutoStart=" + $ap.autoStart)
    }
  } else {
    Add-Row "IIS" "Summary" $env:COMPUTERNAME "WebAdministrationModule" "NotAvailable" "" ""
  }
}

# ---------------- Export ----------------
$dir = Split-Path -Parent $OutCsv
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

$rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $OutCsv
Write-Host "Saved CSV: $OutCsv" -ForegroundColor Green

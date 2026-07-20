<#
.SYNOPSIS
    Collects effective policy state from a Windows endpoint for offline audit.

.DESCRIPTION
    Read-only. Dumps GPO, MDM/CSP, security policy and component state to a
    folder of CSV/TXT/HTML files. Designed for hybrid-joined machines where
    both Group Policy and Intune are live.

    Optionally runs a refresh race test to detect settings that flip between
    the two channels.

.PARAMETER OutputPath
    Destination folder. Created if missing.

.PARAMETER RaceTest
    Runs gpupdate /force then an MDM sync, snapshotting HKLM\SOFTWARE\Policies
    before and after each. Takes ~3 minutes. Produces race-*.txt.

.EXAMPLE
    .\Collect-PolicyState.ps1 -OutputPath C:\Audit
    .\Collect-PolicyState.ps1 -OutputPath C:\Audit -RaceTest

.NOTES
    Run elevated. Non-elevated runs will silently skip several sections.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = "$env:SystemDrive\PolicyAudit\$(Get-Date -f yyyyMMdd-HHmmss)",
    [switch]$RaceTest
)

$ErrorActionPreference = 'Continue'
$null = New-Item -ItemType Directory -Force -Path $OutputPath

$elevated = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

function Step {
    param([string]$Name, [scriptblock]$Action)
    Write-Host "[*] $Name" -ForegroundColor Cyan
    try { & $Action }
    catch { Write-Host "    ! $($_.Exception.Message)" -ForegroundColor Yellow }
}

function Out-File2 { param($Name) Join-Path $OutputPath $Name }

# --------------------------------------------------------------------------
# 0. Host and join context
# --------------------------------------------------------------------------
Step 'Join state and host info' {
    dsregcmd /status                | Out-File (Out-File2 'dsregcmd.txt')
    systeminfo                      | Out-File (Out-File2 'systeminfo.txt')

    [PSCustomObject]@{
        Collected     = Get-Date -Format o
        ComputerName  = $env:COMPUTERNAME
        Elevated      = $elevated
        OSCaption     = (Get-CimInstance Win32_OperatingSystem).Caption
        OSBuild       = (Get-CimInstance Win32_OperatingSystem).BuildNumber
        UBR           = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').UBR
        LastBoot      = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    } | Export-Csv (Out-File2 'host.csv') -NoTypeInformation
}

# --------------------------------------------------------------------------
# 1. MDM enrollments and co-management
# --------------------------------------------------------------------------
Step 'MDM enrollments' {
    Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Enrollments' -EA SilentlyContinue |
        ForEach-Object { Get-ItemProperty $_.PSPath } |
        Where-Object ProviderID |
        Select-Object PSChildName, ProviderID, UPN, EnrollmentType,
                      DiscoveryServiceFullURL, EnrollmentState |
        Export-Csv (Out-File2 'enrollments.csv') -NoTypeInformation
}

Step 'Co-management (ConfigMgr)' {
    $cm = Get-CimInstance -Namespace root\ccm\policy\machine\actualconfig `
             -ClassName CCM_ComanagementSettings -EA SilentlyContinue
    if ($cm) {
        $cm | Select-Object * | Export-Csv (Out-File2 'comanagement.csv') -NoTypeInformation
    } else {
        'No CCM_ComanagementSettings found. ConfigMgr client likely absent.' |
            Out-File (Out-File2 'comanagement.csv')
    }
    Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\CCM\CoManagementFlags' -EA SilentlyContinue |
        Out-File (Out-File2 'comanagement-flags.txt')
}

# --------------------------------------------------------------------------
# 2. MDM / CSP state
# --------------------------------------------------------------------------
function Get-PolicyManagerSet {
    param([string]$Root, [string]$Scope)
    Get-ChildItem $Root -Recurse -EA SilentlyContinue | ForEach-Object {
        $k = $_
        $area = ($k.Name -split '\\')[-1]
        $props = Get-ItemProperty $k.PSPath -EA SilentlyContinue
        if (-not $props) { return }
        $props.PSObject.Properties |
            Where-Object { $_.Name -notmatch '^PS' -and $_.Name -notmatch '_WinningProvider$' } |
            ForEach-Object {
                [PSCustomObject]@{
                    Scope   = $Scope
                    Area    = $area
                    Setting = $_.Name
                    Value   = ($_.Value -join ',')
                    Winner  = $props."$($_.Name)_WinningProvider"
                    Id      = "$area\$($_.Name)"
                }
            }
    }
}

Step 'CSP effective state (current)' {
    $eff  = Get-PolicyManagerSet 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device' 'Device'
    $eff += Get-PolicyManagerSet 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\user'   'User'
    $eff | Sort-Object Scope, Area, Setting |
        Export-Csv (Out-File2 'csp-effective.csv') -NoTypeInformation
    Write-Host "    $($eff.Count) settings" -ForegroundColor DarkGray
}

Step 'CSP pushed state (per provider)' {
    $pushed = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\PolicyManager\providers' `
                -Recurse -EA SilentlyContinue | ForEach-Object {
        $k = $_
        $prov = if ($k.PSPath -match '\\providers\\([^\\]+)') { $Matches[1] } else { '?' }
        $k.GetValueNames() | Where-Object { $_ } | ForEach-Object {
            [PSCustomObject]@{
                Provider = $prov
                Area     = ($k.Name -split '\\')[-1]
                Setting  = $_
                Value    = ($k.GetValue($_) -join ',')
                Id       = "$(($k.Name -split '\\')[-1])\$_"
            }
        }
    }
    $pushed | Sort-Object Provider, Area, Setting |
        Export-Csv (Out-File2 'csp-pushed.csv') -NoTypeInformation

    # settings configured by more than one enrollment
    $pushed | Group-Object Id |
        Where-Object { ($_.Group.Provider | Select-Object -Unique).Count -gt 1 } |
        Select-Object Name, Count |
        Export-Csv (Out-File2 'FINDING-multi-provider.csv') -NoTypeInformation
}

Step 'ControlPolicyConflict / MDMWinsOverGP' {
    $cpc = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\ControlPolicyConflict' -EA SilentlyContinue
    $val = if ($cpc) { $cpc.MDMWinsOverGP } else { $null }
    [PSCustomObject]@{
        MDMWinsOverGP = if ($null -eq $val) { 'NOT CONFIGURED' } else { $val }
        Meaning       = switch ($val) {
                            1       { 'MDM precedence, conflicting GP blocked' }
                            0       { 'Default - race condition possible' }
                            default { 'Default - race condition possible' }
                        }
    } | Export-Csv (Out-File2 'mdmwinsovergp.csv') -NoTypeInformation
}

Step 'MDM diagnostics report' {
    $d = Join-Path $OutputPath 'MDMDiag'
    $null = New-Item -ItemType Directory -Force -Path 'C:\Users\Public\Documents\MDMDiagnostics'
    $null = New-Item -ItemType Directory -Force -Path $d
    & mdmdiagnosticstool.exe -out $d 2>&1 | Out-File (Out-File2 'mdmdiag-stdout.txt')
}

# --------------------------------------------------------------------------
# 3. Group Policy state
# --------------------------------------------------------------------------
Step 'Group Policy results' {
    & gpresult /h (Out-File2 'gpresult.html') /f 2>&1 | Out-Null
    & gpresult /x (Out-File2 'gpresult.xml')  /f 2>&1 | Out-Null
    & gpresult /r /scope:computer | Out-File (Out-File2 'gpresult-computer.txt')
}

Step 'Local GPO artifacts' {
    $paths = @(
        "$env:SystemRoot\System32\GroupPolicy"
        "$env:SystemRoot\System32\GroupPolicyUsers"
        "$env:ProgramData\Microsoft\Group Policy\History"
    )
    $paths | ForEach-Object {
        Get-ChildItem $_ -Recurse -File -EA SilentlyContinue |
            Select-Object FullName, Length, LastWriteTime
    } | Export-Csv (Out-File2 'gpo-files.csv') -NoTypeInformation
}

Step 'Registry policy tree' {
    Get-ChildItem 'HKLM:\SOFTWARE\Policies' -Recurse -EA SilentlyContinue | ForEach-Object {
        $k = $_
        $k.GetValueNames() | ForEach-Object {
            [PSCustomObject]@{
                Key   = $k.Name
                Value = $_
                Data  = ($k.GetValue($_) -join ',')
            }
        }
    } | Sort-Object Key, Value | Export-Csv (Out-File2 'registry-policies.csv') -NoTypeInformation
}

# --------------------------------------------------------------------------
# 4. Security policy, audit, rights
# --------------------------------------------------------------------------
Step 'Security policy and audit' {
    if ($elevated) {
        & secedit /export /cfg (Out-File2 'secpol.inf') /quiet 2>&1 | Out-Null
        & auditpol /get /category:* | Out-File (Out-File2 'auditpol.txt')
    } else {
        'Skipped - requires elevation' | Out-File (Out-File2 'secpol.inf')
    }
    net accounts        | Out-File (Out-File2 'net-accounts-local.txt')
    net accounts /domain 2>&1 | Out-File (Out-File2 'net-accounts-domain.txt')
}

# --------------------------------------------------------------------------
# 5. Component actual state (truth, not intent)
# --------------------------------------------------------------------------
Step 'Defender / firewall / BitLocker / LAPS' {
    Get-MpPreference     -EA SilentlyContinue | Format-List * | Out-File (Out-File2 'defender-prefs.txt')
    Get-MpComputerStatus -EA SilentlyContinue | Format-List * | Out-File (Out-File2 'defender-status.txt')

    Get-NetFirewallProfile -EA SilentlyContinue |
        Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, LogFileName |
        Export-Csv (Out-File2 'firewall-profiles.csv') -NoTypeInformation

    Get-NetFirewallRule -PolicyStore ActiveStore -EA SilentlyContinue |
        Select-Object DisplayName, Enabled, Direction, Action, Profile |
        Export-Csv (Out-File2 'firewall-rules.csv') -NoTypeInformation

    Get-BitLockerVolume -EA SilentlyContinue |
        Select-Object MountPoint, VolumeStatus, ProtectionStatus, EncryptionMethod, EncryptionPercentage |
        Export-Csv (Out-File2 'bitlocker.csv') -NoTypeInformation

    Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Policies\LAPS' -EA SilentlyContinue |
        Out-File (Out-File2 'laps.txt')

    Get-ChildItem 'C:\Windows\System32\CodeIntegrity\CiPolicies\Active' -EA SilentlyContinue |
        Select-Object Name, Length, LastWriteTime |
        Export-Csv (Out-File2 'wdac-policies.csv') -NoTypeInformation
}

# --------------------------------------------------------------------------
# 6. Logs
# --------------------------------------------------------------------------
Step 'Event logs' {
    Get-WinEvent -LogName 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin' `
        -MaxEvents 3000 -EA SilentlyContinue |
        Select-Object TimeCreated, Id, LevelDisplayName, Message |
        Export-Csv (Out-File2 'log-mdm.csv') -NoTypeInformation

    Get-WinEvent -LogName 'Microsoft-Windows-GroupPolicy/Operational' `
        -MaxEvents 2000 -EA SilentlyContinue |
        Select-Object TimeCreated, Id, LevelDisplayName, Message |
        Export-Csv (Out-File2 'log-gpo.csv') -NoTypeInformation
}

Step 'Intune Management Extension logs' {
    $ime = 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs'
    if (Test-Path $ime) {
        Copy-Item $ime -Destination (Join-Path $OutputPath 'IME-Logs') -Recurse -Force -EA SilentlyContinue
    }
}

# --------------------------------------------------------------------------
# 7. Optional: race test
# --------------------------------------------------------------------------
function Get-PolicySnapshot {
    Get-ChildItem 'HKLM:\SOFTWARE\Policies' -Recurse -EA SilentlyContinue | ForEach-Object {
        $k = $_
        $k.GetValueNames() | ForEach-Object { "$($k.Name)\$_ = $($k.GetValue($_) -join ',')" }
    } | Sort-Object
}

if ($RaceTest) {
    Step 'Race test: baseline snapshot' {
        $script:snapBefore = Get-PolicySnapshot
        $snapBefore | Out-File (Out-File2 'race-1-before.txt')
    }
    Step 'Race test: gpupdate /force' {
        & gpupdate /force /target:computer 2>&1 | Out-File (Out-File2 'race-gpupdate.txt')
        Start-Sleep -Seconds 30
        $script:snapGP = Get-PolicySnapshot
        $snapGP | Out-File (Out-File2 'race-2-after-gp.txt')
    }
    Step 'Race test: MDM sync' {
        Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' `
            -TaskName PushLaunch -EA SilentlyContinue | Start-ScheduledTask
        Start-Sleep -Seconds 90
        $script:snapMDM = Get-PolicySnapshot
        $snapMDM | Out-File (Out-File2 'race-3-after-mdm.txt')
    }
    Step 'Race test: diffs' {
        Compare-Object $snapBefore $snapGP  |
            Out-File (Out-File2 'race-diff-gp.txt')
        Compare-Object $snapGP    $snapMDM |
            Out-File (Out-File2 'FINDING-race-diff-mdm.txt')
    }
}

# --------------------------------------------------------------------------
Write-Host ''
Write-Host "Done. Output: $OutputPath" -ForegroundColor Green
if (-not $elevated) {
    Write-Host 'WARNING: not elevated - secedit, auditpol and some logs were skipped.' -ForegroundColor Yellow
}
Get-ChildItem $OutputPath | Select-Object Name, Length | Format-Table -AutoSize
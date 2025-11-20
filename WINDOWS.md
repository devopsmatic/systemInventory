```powershell
<#
What are we capturing:
Software installed and certain key configurations or software settings by the developer

How to RUN:
1. Any developer with admin rights can peform this action.
2. Save the script as Collect-Software-Fast.ps1 in Downloads folder.
3. Run PowerShell as Administrator.

Send us this:
4. This script will create a folder structure:
C:\MYPC
└───InstalledSoftware
    └───2025-NOV-20-15-39-46-MYPC
        └───Configs
Where MYPC = the current Computer Name

Zip this folder MYPC and send us. 
Example: MYPC.zip (MYPC is my Computer Name)

What if you run this script multiple times:
NOTHING TO WORRY

Each run creates an unique folder that has a different DATE TIME like: 
2025-NOV-20-15-39-46-MYPC
2025-NOV-20-18-00-23-MYPC 
2025-NOV-22-11-55-11-MYPC

Just ZIP the TOP LEVEL FOLDER that is the Computer Name and send us
#>


param(
    [switch]$FullMode  # if specified, do the slower provider-by-provider Get-Package enumeration
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

#region Helpers
function Get-HumanRunLabel {
    $dt = Get-Date
    return "{0}-{1}-{2}-{3}-{4}-{5}" -f $dt.Year, $dt.ToString('MMM').ToUpper(), $dt.ToString('dd'), $dt.ToString('HH'), $dt.ToString('mm'), $dt.ToString('ss')
}
function Normalize-InstallDate { param($d) if (-not $d) { return $null } $s=[string]$d; if ($s -match '^\d{8}$') { try { return ([datetime]::ParseExact($s,'yyyyMMdd',$null)).ToString('yyyy-MM-dd') } catch { return $s } } try { return ([datetime]$s).ToString('yyyy-MM-dd') } catch { return $s } }
function Try-InstallImportExcel {
    try {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        }
        return $true
    } catch {
        Write-Host ("ImportExcel install failed: " + $_.Exception.Message) -ForegroundColor Yellow
        return $false
    }
}
function Show-Step { param($i,$t,$title) $pct = [int](($i/$t)*100); Write-Host ""; Write-Host ('#' * 72) -ForegroundColor DarkGray; Write-Host ("STEP {0}/{1} - {2} - {3}% complete" -f $i,$t,$title,$pct) -ForegroundColor White; Write-Progress -Activity "Collect-DevProfile" -Status $title -PercentComplete $pct -Id 1 }
function Show-Warn { param($m) Write-Host ("WARNING: " + $m) -ForegroundColor Yellow }
function Show-Err { param($m) Write-Host ("ERROR: " + $m) -ForegroundColor Red }
#endregion

#region Prepare output folders
# Updated steps count (one extra for hwinfo)
$totalSteps = 17
$currentStep = 0

$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Preparing output folder"
$computerName = $env:COMPUTERNAME
$baseFolder = Join-Path "C:\" "$computerName\InstalledSoftware"
try {
    if (-not (Test-Path $baseFolder)) { New-Item -Path $baseFolder -ItemType Directory -Force | Out-Null; Write-Host ("Created base folder: " + $baseFolder) -ForegroundColor Green }
    else { Write-Host ("Base folder exists: " + $baseFolder) -ForegroundColor Green }
} catch { Show-Err ("Cannot create base folder: " + $_.Exception.Message); throw }

$runLabel = Get-HumanRunLabel
$runFolderName = "$runLabel-$computerName"
$runFolder = Join-Path $baseFolder $runFolderName
New-Item -Path $runFolder -ItemType Directory -Force | Out-Null
$configFolder = Join-Path $runFolder "Configs"
New-Item -Path $configFolder -ItemType Directory -Force | Out-Null

$csvPath = Join-Path $runFolder "InstalledSoftware.csv"
$excelPath = Join-Path $runFolder "InstalledSoftware.xlsx"

Write-Host ("Outputs under: " + $runFolder) -ForegroundColor Cyan
#endregion

#region Collect: Registry uninstall entries (fast, guarded)
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Collecting registry Uninstall entries"
$softwareList = [System.Collections.Generic.List[PSObject]]::new()
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($path in $regPaths) {
    try {
        $items = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            try {
                if ($item -and $item.PSObject -and $item.PSObject.Properties['DisplayName'] -and $item.DisplayName -and $item.DisplayName.Trim()) {
                    $regPathFull = $item.PSPath -replace '.*?::',''
                    $keyName = ($regPathFull -split '\\')[-1]
                    $hasWindowsInstaller = $item.PSObject.Properties['WindowsInstaller'] -ne $null
                    $msiCode = if ($hasWindowsInstaller -and $item.WindowsInstaller -eq 1 -and ($keyName -match '^\{[0-9A-Fa-f-]{36}\}$')) { $keyName } else { $null }
                    $softwareList.Add([PSCustomObject]@{
                        Source = "Registry"
                        Name = ($item.DisplayName -as [string]).Trim()
                        Version = ($item.DisplayVersion -as [string])
                        Publisher = $item.Publisher
                        InstallDateRaw = $item.InstallDate
                        InstallDate = Normalize-InstallDate $item.InstallDate
                        InstallLocation = $item.InstallLocation
                        UninstallString = $item.UninstallString
                        RegistryKey = $regPathFull
                        Provider = if ($hasWindowsInstaller -and $item.WindowsInstaller -eq 1) { 'MSI' } else { $null }
                        MSIProductCode = $msiCode
                        ExtraSource = $null
                    })
                }
            } catch {
                $msg = $_.Exception.Message
                if ($msg -match "The property '([^']+)' cannot be found") {
                    Show-Warn ("Skipped non-existent property " + $Matches[1])
                } else {
                    Show-Warn ("Skipped one registry entry due to error: " + $msg)
                }
            }
        }
    } catch {
        Show-Warn ("Failed to read registry path " + $path + ": " + $_.Exception.Message)
    }
}
Write-Host ("Registry entries collected: " + $softwareList.Count) -ForegroundColor Green
#endregion

#region Collect: MSI products via COM (fast and reliable)
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Enumerating MSI products (COM)"
try {
    $msi = New-Object -ComObject WindowsInstaller.Installer
    $prods = @($msi.Products())
    $i = 0; $total = $prods.Count
    foreach ($p in $prods) {
        $i++; $pct = [int](($i/$total)*100)
        Write-Progress -Id 2 -Activity "MSI enumeration" -Status ("Processing MSI product $i / $total") -PercentComplete $pct
        try {
            $name = $msi.ProductInfo($p, "InstalledProductName") -as [string]
        } catch { $name = $null }
        try { $ver = $msi.ProductInfo($p, "VersionString") -as [string] } catch { $ver = $null }
        $softwareList.Add([PSCustomObject]@{
            Source = "MSI"
            Name = $name
            Version = $ver
            Publisher = $null
            InstallDateRaw = $null
            InstallDate = $null
            InstallLocation = $null
            UninstallString = $null
            RegistryKey = $null
            Provider = "MSI"
            MSIProductCode = $p
            ExtraSource = $null
        })
    }
    Write-Progress -Id 2 -Activity "MSI enumeration" -Completed
    Write-Host ("MSI products added: " + $prods.Count) -ForegroundColor Green
} catch {
    Show-Warn ("MSI COM enumeration failed: " + $_.Exception.Message)
}
#endregion

#region Collect: Appx (Store) local
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Collecting Appx (Store) packages"
try {
    $appx = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    if ($appx) {
        foreach ($a in $appx) {
            $softwareList.Add([PSCustomObject]@{
                Source = "AppxPackage"
                Name = ($a.Name -as [string])
                Version = ($a.Version -as [string])
                Publisher = $a.Publisher
                InstallDateRaw = $null
                InstallDate = $null
                InstallLocation = $a.InstallLocation
                UninstallString = "Remove-AppxPackage -Package $($a.PackageFullName)"
                RegistryKey = $null
                Provider = "Appx"
                MSIProductCode = $null
                ExtraSource = $a.PackageFamilyName
            })
        }
        Write-Host ("Appx packages: " + $appx.Count) -ForegroundColor Green
    } else { Write-Host "No Appx packages or Get-AppxPackage failed." -ForegroundColor Yellow }
} catch { Show-Warn ("Get-AppxPackage failed: " + $_.Exception.Message) }
#endregion

#region Collect: winget local DB (fast local)
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Collecting winget local DB (if present)"
try {
    $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
    if ($wingetCmd) {
        $json = $null
        try { $json = & winget list --source winget --accept-package-agreements --accept-source-agreements --output json 2>$null } catch { $json = $null }
        if ($json) {
            try {
                $parsed = $json | ConvertFrom-Json
                foreach ($w in $parsed) {
                    $softwareList.Add([PSCustomObject]@{
                        Source = "winget"
                        Name = $w.Name
                        Version = $w.Version
                        Publisher = ($w.Publisher -replace '\s*$','')
                        InstallDateRaw = $null
                        InstallDate = $null
                        InstallLocation = $null
                        UninstallString = $null
                        RegistryKey = $null
                        Provider = "winget"
                        MSIProductCode = $null
                        ExtraSource = $w.Id
                    })
                }
                Write-Host ("winget entries: " + ($parsed.Count)) -ForegroundColor Green
            } catch { Show-Warn ("winget parse failed: " + $_.Exception.Message) }
        } else { Write-Host "winget returned no JSON; skipping." -ForegroundColor Yellow }
    } else { Write-Host "winget not found; skipping." -ForegroundColor Yellow }
} catch { Show-Warn ("winget invocation failed: " + $_.Exception.Message) }
#endregion

#region Collect: Program Files quick heuristic (portable apps)
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Scanning Program Files (heuristic)"
try {
    $pfRoots = @("$env:ProgramFiles","$env:ProgramFiles(x86)","$env:LocalAppData\Programs")
    foreach ($root in $pfRoots) {
        if (Test-Path $root) {
            foreach ($d in Get-ChildItem -Path $root -Directory -ErrorAction SilentlyContinue) {
                $exe = Get-ChildItem -Path $d.FullName -Filter *.exe -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($exe) {
                    $softwareList.Add([PSCustomObject]@{
                        Source = "ProgramFilesScan"
                        Name = $d.Name
                        Version = $null
                        Publisher = $null
                        InstallDateRaw = $null
                        InstallDate = $null
                        InstallLocation = $d.FullName
                        UninstallString = $null
                        RegistryKey = $null
                        Provider = $null
                        MSIProductCode = $null
                        ExtraSource = $null
                    })
                }
            }
        }
    }
    Write-Host "ProgramFiles scan complete." -ForegroundColor Green
} catch { Show-Warn ("ProgramFiles scan failed: " + $_.Exception.Message) }
#endregion

#region Optional: FullMode (provider-by-provider) - only if user asked explicitly
if ($FullMode) {
    $currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Running Get-Package (provider-by-provider, FULL MODE)"
    try {
        $allPkgs = Get-Package -ErrorAction SilentlyContinue
        if ($allPkgs) {
            $providers = $allPkgs | Group-Object -Property ProviderName
            $pi = 0; $tp = $providers.Count
            foreach ($g in $providers) {
                $pi++; $provName = $g.Name
                $pkgs = $g.Group; $totalPkgs = $pkgs.Count; $i=0
                $sw = [Diagnostics.Stopwatch]::StartNew()
                foreach ($p in $pkgs) {
                    $i++;
                    $pct = [int](($i/$totalPkgs)*100)
                    $elapsed = $sw.Elapsed.TotalSeconds
                    $avg = if ($i -gt 1) { $elapsed / $i } else { 0.5 }
                    $eta = [int](($totalPkgs - $i) * $avg)
                    Write-Progress -Id 2 -Activity ("Get-Package ($provName)") -Status ("$i/$totalPkgs ETA: " + (New-TimeSpan -Seconds $eta).ToString()) -PercentComplete $pct
                    $softwareList.Add([PSCustomObject]@{
                        Source = "Get-Package"
                        Name = ($p.Name -as [string]).Trim()
                        Version = ($p.Version -as [string])
                        Publisher = $p.ProviderName
                        InstallDateRaw = $null
                        InstallDate = $null
                        InstallLocation = $null
                        UninstallString = $null
                        RegistryKey = $null
                        Provider = $p.ProviderName
                        MSIProductCode = $null
                        ExtraSource = $p.Source
                    })
                }
                Write-Progress -Id 2 -Activity "Get-Package" -Completed
            }
        } else { Write-Host "Get-Package returned nothing." -ForegroundColor Yellow }
    } catch { Show-Warn ("Get-Package full mode failed: " + $_.Exception.Message) }
} else {
    Write-Host "Running in LOCAL-ONLY fast mode. Use -FullMode to enable provider enumeration." -ForegroundColor Cyan
}
#endregion

#region Deduplicate & shape output
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Deduplicating and shaping results"
try {
    $grouped = $softwareList | Group-Object @{Expression = { ($_.Name -as [string]).ToLower() + '|' + ($_.Version -as [string]) } }
    $deduped = foreach ($g in $grouped) {
        $items = $g.Group
        if ($items.Count -eq 1) { $items[0] } else {
            $reg = $items | Where-Object { $_.Source -eq 'Registry' -and ($_.UninstallString -or $_.InstallLocation) } | Select-Object -First 1
            if ($reg) { $reg } else {
                $pkg = $items | Where-Object { $_.Source -eq 'Get-Package' } | Select-Object -First 1
                if ($pkg) { $pkg } else {
                    $appx = $items | Where-Object { $_.Source -eq 'AppxPackage' } | Select-Object -First 1
                    if ($appx) { $appx } else { $items[0] }
                }
            }
        }
    }
    $result = $deduped | Sort-Object @{Expression = { ($_.Name -as [string]).ToLower() } } |
        Select-Object Source, Name, Version, Publisher, @{Name='InstallDate';Expression={$_.InstallDate}}, InstallLocation, UninstallString, Provider, MSIProductCode, ExtraSource
    Write-Host ("Unique rows: " + $result.Count) -ForegroundColor Green
} catch { Show-Err ("Dedup failed: " + $_.Exception.Message) }
#endregion

#region Export CSV & Excel attempt
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Exporting CSV and XLSX (if available)"
try {
    $result | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host ("CSV created: " + $csvPath) -ForegroundColor Green
} catch { Show-Err ("CSV export failed: " + $_.Exception.Message) }

try {
    if (Try-InstallImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        $result | Export-Excel -Path $excelPath -WorksheetName "SoftwareInventory" -AutoSize -FreezeTopRow -ErrorAction Stop
        Write-Host ("XLSX created: " + $excelPath) -ForegroundColor Green
    } else {
        Show-Warn "Excel export skipped (ImportExcel not available or install failed)."
    }
} catch { Show-Warn ("Excel export failed: " + $_.Exception.Message) }
#endregion

#region Export Configs: envs, git, python (parallel pip), node, vscode, docker, summary
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Exporting environment variables (Machine/User/All)"
try {
    $envMachinePath = Join-Path $configFolder "Env-System.csv"
    $m = [System.Environment]::GetEnvironmentVariables('Machine')
    $m.GetEnumerator() | ForEach-Object { [PSCustomObject]@{ Scope='Machine'; Name=$_.Key; Value=[string]$_.Value } } | Sort-Object Name | Export-Csv -Path $envMachinePath -NoTypeInformation -Encoding UTF8 -Force
    $envUserPath = Join-Path $configFolder "Env-User.csv"
    $u = [System.Environment]::GetEnvironmentVariables('User')
    $u.GetEnumerator() | ForEach-Object { [PSCustomObject]@{ Scope='User'; Name=$_.Key; Value=[string]$_.Value } } | Sort-Object Name | Export-Csv -Path $envUserPath -NoTypeInformation -Encoding UTF8 -Force
    $envAllPath = Join-Path $configFolder "All-EnvironmentVariables.csv"
    $proc = [System.Environment]::GetEnvironmentVariables('Process')
    $combined = @()
    foreach ($k in $m.Keys) { $combined += [PSCustomObject]@{ Scope='Machine'; Name=$k; Value=[string]$m[$k] } }
    foreach ($k in $u.Keys) { $combined += [PSCustomObject]@{ Scope='User'; Name=$k; Value=[string]$u[$k] } }
    foreach ($k in $proc.Keys) { $combined += [PSCustomObject]@{ Scope='Process'; Name=$k; Value=[string]$proc[$k] } }
    $combined | Sort-Object Scope, Name | Export-Csv -Path $envAllPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "Environment exports done." -ForegroundColor Green
} catch { Show-Warn ("Env export failed: " + $_.Exception.Message) }

# Git
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Exporting Git & SSH config"
try {
    $gitConfigPath = Join-Path $configFolder "Git-Config.txt"
    $gitSshConfigPath = Join-Path $configFolder "Git-SSH-Config.txt"
    $gout = @()
    $gout += "git --version:`n" + ((& git --version 2>&1) -join "`n")
    $gout += "`n-- git global config --`n" + ((& git config --global --list 2>&1) -join "`n")
    $gitcfg = Join-Path $env:USERPROFILE ".gitconfig"
    if (Test-Path $gitcfg) { $gout += "`n-- .gitconfig --`n" + (Get-Content $gitcfg -Raw) }
    Set-Content -Path $gitConfigPath -Value $gout -Encoding UTF8 -Force
    $sshCfg = Join-Path $env:USERPROFILE ".ssh\config"
    if (Test-Path $sshCfg) { Copy-Item -Path $sshCfg -Destination $gitSshConfigPath -Force } else { Set-Content -Path $gitSshConfigPath -Value "No SSH config found." -Encoding UTF8 -Force }
    Write-Host "Git export done." -ForegroundColor Green
} catch { Show-Warn ("Git export failed: " + $_.Exception.Message) }

# Python - detect interpreters and pip freeze in parallel with timeout
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Detecting Python interpreters & pip packages (parallel)"
try {
    $pythonInterpreters = @()
    try { $pyOut = & py -0p 2>$null } catch { $pyOut = $null }
    if ($pyOut) { $lines = $pyOut -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }; foreach ($ln in $lines) { if ($ln -match '^\s*(?:-p)?\s*(.+)$') { $path = $Matches[1].Trim(); if (Test-Path $path) { $pythonInterpreters += (Resolve-Path $path).Path } } } }
    try { $where = (& where.exe python 2>$null) -split "`n" } catch { $where = @() }
    foreach ($w in $where) { if ($w -and (Test-Path $w.Trim())) { $pythonInterpreters += (Resolve-Path $w.Trim()).Path } }
    $pythonInterpreters = $pythonInterpreters | Sort-Object -Unique

    $pipJobs = @()
    $pipTimeoutSec = 20
    foreach ($py in $pythonInterpreters) {
        $pipJobs += Start-Job -ArgumentList $py -ScriptBlock {
            param($pythonPath)
            try {
                $ver = & $pythonPath --version 2>&1
                $pkgs = & $pythonPath -m pip freeze 2>&1
                return @{Interpreter=$pythonPath; Version=$ver; Packages=$pkgs}
            } catch { return @{Interpreter=$pythonPath; Version=$null; Packages=@("ERROR: "+$_.Exception.Message)} }
        }
    }

    $deadline = (Get-Date).AddSeconds( ($pythonInterpreters.Count + 1) * $pipTimeoutSec )
    while ((Get-Job -State Running).Count -gt 0 -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }

    $pySummary = Join-Path $configFolder "Python-Packages-Summary.txt"
    $out = @()
    foreach ($j in Get-Job) {
        $res = Receive-Job -Job $j -ErrorAction SilentlyContinue
        if ($res) {
            $fn = Join-Path $configFolder ("Python-" + ((Split-Path $res.Interpreter -Leaf) -replace '[^0-9A-Za-z\._-]','') + "-Packages.txt")
            if ($res.Packages) { Set-Content -Path $fn -Value ($res.Packages -join "`n") -Encoding UTF8 -Force } else { Set-Content -Path $fn -Value "No packages returned." -Encoding UTF8 -Force }
            $out += [PSCustomObject]@{ Interpreter=$res.Interpreter; Version=$res.Version; PackagesFile=$fn }
        }
        Remove-Job -Job $j -Force -ErrorAction SilentlyContinue
    }
    $out | ConvertTo-Csv -NoTypeInformation | Set-Content -Path $pySummary -Encoding UTF8 -Force
    Write-Host "Python detection complete." -ForegroundColor Green
} catch { Show-Warn ("Python detection failed: " + $_.Exception.Message) }

# NodeJS / npm summary
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Collecting NodeJS / npm info"
try {
    $nodeF = Join-Path $configFolder "NodeJS.txt"
    $nOut = @()
    $nOut += "node --version:`n" + (($(& node --version 2>&1) -join "`n"))
    $nOut += "`n npm --version:`n" + (($(& npm --version 2>&1) -join "`n"))
    $nOut += "`n npm -g list --depth=0:`n" + (($(& npm list -g --depth=0 2>&1) -join "`n"))
    Set-Content -Path $nodeF -Value $nOut -Encoding UTF8 -Force
    Write-Host "Node info exported." -ForegroundColor Green
} catch { Show-Warn ("Node step failed: " + $_.Exception.Message) }

# VSCode
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Exporting VSCode settings & extensions"
try {
    $vSettings = Join-Path $configFolder "VSCode-Settings.json"
    $vExts = Join-Path $configFolder "VSCode-Extensions.txt"
    $appdata = $env:APPDATA
    $vsSettingsSrc = Join-Path $appdata "Code\User\settings.json"
    if (Test-Path $vsSettingsSrc) { Copy-Item -Path $vsSettingsSrc -Destination $vSettings -Force } else { Set-Content -Path $vSettings -Value "VSCode settings not found." -Encoding UTF8 -Force }
    $codeCmd = Get-Command code -ErrorAction SilentlyContinue
    if ($codeCmd) { $exts = & code --list-extensions 2>$null; if ($exts) { $exts | Set-Content -Path $vExts -Encoding UTF8 -Force } else { Set-Content -Path $vExts -Value "No extensions returned." -Encoding UTF8 -Force } } else { Set-Content -Path $vExts -Value "VSCode CLI not found." -Encoding UTF8 -Force }
    Write-Host "VSCode exports done." -ForegroundColor Green
} catch { Show-Warn ("VSCode export failed: " + $_.Exception.Message) }

# Docker info
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Collecting Docker info (if present)"
try {
    $dockerF = Join-Path $configFolder "Docker-Info.txt"
    $dv = (& docker --version 2>&1) -join "`n"
    if ($dv -and -not ($dv -match 'not found|is not recognized')) {
        Set-Content -Path $dockerF -Value ((& docker --version 2>&1) -join "`n" + "`n" + ((& docker info 2>&1) -join "`n")) -Encoding UTF8 -Force
    } else { Set-Content -Path $dockerF -Value "Docker not installed or not in PATH." -Encoding UTF8 -Force }
    Write-Host "Docker info exported." -ForegroundColor Green
} catch { Show-Warn ("Docker step failed: " + $_.Exception.Message) }
#endregion

#region Collect: Hardware & Network Info -> hwinfo.csv + hwinfo.xlsx
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Collecting Hardware & Network info (hwinfo)"

try {
    $hwSummary = [ordered]@{}

    # OS info
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($os) {
        $hwSummary.OS_Caption = $os.Caption
        $hwSummary.OS_Name = $os.Caption -replace "Microsoft\s+",""
        $hwSummary.OS_Version = $os.Version
        $hwSummary.OS_Build = $os.BuildNumber
        $hwSummary.OS_Architecture = $os.OSArchitecture
        $hwSummary.LastBootUpTime = if ($os.LastBootUpTime) { ([Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)).ToString('yyyy-MM-dd HH:mm:ss') } else { $null }
    } else {
        Show-Warn "Unable to read Win32_OperatingSystem"
    }

    # Windows edition & activation state (best-effort)
    try {
        $slp = Get-CimInstance -Namespace root\cimv2 -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue |
               Where-Object { $_.PartialProductKey -and ($_.Name -match 'Windows') } |
               Select-Object -First 1
        if ($slp) {
            $hwSummary.Windows_Edition = $slp.Name
            $hwSummary.Windows_LicenseStatus = $slp.LicenseStatus
            $lsMap = @{
                0 = 'Unlicensed/Unknown'; 1='Licensed'; 2='Out-Of-Box Grace'; 3='Out-Of-Tolerance Grace'; 4='Non-Genuine Grace';
                5='Notification'; 6='Extended Grace'
            }
            $hwSummary.Windows_LicenseStatusText = $lsMap[$slp.LicenseStatus]
        } else {
            $sls = Get-CimInstance -Namespace root\cimv2 -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue
            if ($sls) { $hwSummary.Windows_LicenseStatusText = "OOBMgmtChannelEnabled:$($sls.OOBManagementChannelEnabled)" }
        }
    } catch {
        Show-Warn ("Windows license query failed: " + $_.Exception.Message)
    }

    # Manufacturer / Model
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
    if ($cs) {
        $hwSummary.Manufacturer = $cs.Manufacturer
        $hwSummary.Model = $cs.Model
        $hwSummary.SystemType = $cs.SystemType
        $hwSummary.TotalPhysicalMemoryBytes = [int64]$cs.TotalPhysicalMemory
        $hwSummary.TotalPhysicalMemoryGB = if ($cs.TotalPhysicalMemory) { [math]::Round($cs.TotalPhysicalMemory / 1GB,2) } else { $null }
    }

    # CPU(s)
    $cpus = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue
    $cpuRows = @()
    if ($cpus) {
        foreach ($c in $cpus) {
            $cpuRows += [PSCustomObject]@{
                Name = $c.Name
                Manufacturer = $c.Manufacturer
                NumberOfCores = $c.NumberOfCores
                NumberOfLogicalProcessors = $c.NumberOfLogicalProcessors
                MaxClockSpeedMHz = $c.MaxClockSpeed
                Architecture = $c.AddressWidth
            }
        }
        $hwSummary.CPU = ($cpus | Select-Object -First 1).Name
    }

    # RAM modules
    $mem = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue
    $memRows = @()
    if ($mem) {
        foreach ($m in $mem) {
            $memRows += [PSCustomObject]@{
                DeviceLocator = $m.DeviceLocator
                CapacityBytes = [int64]$m.Capacity
                CapacityGB = if ($m.Capacity) { [math]::Round($m.Capacity / 1GB, 2) } else { $null }
                SpeedMHz = $m.Speed
                Manufacturer = $m.Manufacturer
                PartNumber = $m.PartNumber
            }
        }
    }

    # Video / GPU
    $gpus = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue
    $gpuRows = @()
    if ($gpus) {
        foreach ($g in $gpus) {
            $gpuRows += [PSCustomObject]@{
                Name = $g.Name
                DriverVersion = $g.DriverVersion
                VideoProcessor = $g.VideoProcessor
                AdapterRAM_Bytes = if ($g.AdapterRAM) { [int64]$g.AdapterRAM } else { $null }
                AdapterRAM_GB = if ($g.AdapterRAM) { [math]::Round($g.AdapterRAM/1GB,2) } else { $null }
            }
        }
        $hwSummary.PrimaryGPU = ($gpus | Select-Object -First 1).Name
    }

    # Disks (physical and logical)
    $diskDrive = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue
    $diskRows = @()
    if ($diskDrive) {
        foreach ($d in $diskDrive) {
            $diskRows += [PSCustomObject]@{
                Model = $d.Model
                InterfaceType = $d.InterfaceType
                MediaType = $d.MediaType
                SizeBytes = if ($d.Size) { [int64]$d.Size } else { $null }
                SizeGB = if ($d.Size) { [math]::Round($d.Size/1GB,2) } else { $null }
                SerialNumber = $d.SerialNumber
            }
        }
        $primaryDisk = $diskRows | Select-Object -First 1
        if ($primaryDisk) { $hwSummary.PrimaryDiskModel = $primaryDisk.Model; $hwSummary.PrimaryDiskSizeGB = $primaryDisk.SizeGB }
    }

    # Logical drives (to show drive letters + sizes)
    $logical = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
    $logicalRows = @()
    if ($logical) {
        foreach ($ld in $logical) {
            $logicalRows += [PSCustomObject]@{
                DeviceID = $ld.DeviceID
                VolumeName = $ld.VolumeName
                FileSystem = $ld.FileSystem
                SizeBytes = if ($ld.Size) { [int64]$ld.Size } else { $null }
                SizeGB = if ($ld.Size) { [math]::Round($ld.Size/1GB,2) } else { $null }
                FreeBytes = if ($ld.FreeSpace) { [int64]$ld.FreeSpace } else { $null }
                FreeGB = if ($ld.FreeSpace) { [math]::Round($ld.FreeSpace/1GB,2) } else { $null }
            }
        }
    }

    # Network adapters & current IPs
    $netAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True" -ErrorAction SilentlyContinue
    $netRows = @()
    $primaryIP = $null
    if ($netAdapters) {
        foreach ($na in $netAdapters) {
            $ips = $na.IPAddress -join ';'
            $netRows += [PSCustomObject]@{
                Description = $na.Description
                MACAddress = $na.MACAddress
                IPAddress = $ips
                DefaultIPGateway = ($na.DefaultIPGateway -join ';')
                DNSServerSearchOrder = ($na.DNSServerSearchOrder -join ';')
                DHCPEnabled = $na.DHCPEnabled
                ServiceName = $na.ServiceName
            }
            if (-not $primaryIP -and $na.IPAddress) {
                $firstipv4 = ($na.IPAddress | Where-Object { $_ -and ($_ -notmatch '^169\.254\.') -and ($_ -match '^\d{1,3}(\.\d{1,3}){3}$') } | Select-Object -First 1)
                if ($firstipv4) { $primaryIP = $firstipv4 }
            }
        }
    } else {
        try {
            $ips = Get-NetIPAddress -AddressFamily IPv4 -PrefixOrigin 'Dhcp','Manual' -ErrorAction SilentlyContinue |
                   Where-Object { $_.IPAddress -notlike '169.254.*' -and $_.IPAddress -ne '127.0.0.1' } |
                   Select-Object -First 1
            if ($ips) { $primaryIP = $ips.IPAddress }
        } catch {}
    }
    $hwSummary.PrimaryIP = $primaryIP

    # USB controllers and serial ports (physical ports)
    $usb = Get-CimInstance -ClassName Win32_USBController -ErrorAction SilentlyContinue
    $usbRows = @()
    if ($usb) { foreach ($u in $usb) { $usbRows += [PSCustomObject]@{ Name = $u.Name; DeviceID = $u.DeviceID; PNPDeviceID = $u.PNPDeviceID } } }

    $serial = Get-CimInstance -ClassName Win32_SerialPort -ErrorAction SilentlyContinue
    $serialRows = @()
    if ($serial) { foreach ($s in $serial) { $serialRows += [PSCustomObject]@{ DeviceID = $s.DeviceID; Name = $s.Name; ProviderType = $s.ProviderType } } }

    # BIOS / Firmware info
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
    if ($bios) {
        $hwSummary.BIOS_Manufacturer = $bios.Manufacturer
        $hwSummary.BIOS_SerialNumber = $bios.SerialNumber
        $hwSummary.BIOS_Version = ($bios.SMBIOSBIOSVersion -as [string])
    }

    # Build summary table row (flatten some fields)
    $summaryRow = [PSCustomObject]@{
        ComputerName = $computerName
        Manufacturer = ($hwSummary.Manufacturer -as [string])
        Model = ($hwSummary.Model -as [string])
        SystemType = ($hwSummary.SystemType -as [string])
        OS = ($hwSummary.OS_Caption -as [string])
        OS_Version = ($hwSummary.OS_Version -as [string])
        OS_Build = ($hwSummary.OS_Build -as [string])
        Windows_Edition = ($hwSummary.Windows_Edition -as [string])
        Windows_LicenseStatus = ($hwSummary.Windows_LicenseStatusText -as [string])
        CPU = ($hwSummary.CPU -as [string])
        RAM_GB = ($hwSummary.TotalPhysicalMemoryGB -as [string])
        PrimaryGPU = ($hwSummary.PrimaryGPU -as [string])
        PrimaryDiskModel = ($hwSummary.PrimaryDiskModel -as [string])
        PrimaryDiskSize_GB = ($hwSummary.PrimaryDiskSizeGB -as [string])
        PrimaryIP = ($hwSummary.PrimaryIP -as [string])
        BIOS_Serial = ($hwSummary.BIOS_SerialNumber -as [string])
    }

    # Export summary CSV
    $hwCsv = Join-Path $configFolder "hwinfo.csv"
    $summaryRow | Export-Csv -Path $hwCsv -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "hwinfo summary CSV created: $hwCsv" -ForegroundColor Green

    # Also export detailed component CSVs
    $cpuCsv = Join-Path $configFolder "hw-cpus.csv"; $cpuRows | Export-Csv -Path $cpuCsv -NoTypeInformation -Encoding UTF8 -Force
    $memCsv = Join-Path $configFolder "hw-memory.csv"; $memRows | Export-Csv -Path $memCsv -NoTypeInformation -Encoding UTF8 -Force
    $gpuCsv = Join-Path $configFolder "hw-gpus.csv"; $gpuRows | Export-Csv -Path $gpuCsv -NoTypeInformation -Encoding UTF8 -Force
    $diskCsv = Join-Path $configFolder "hw-disks.csv"; $diskRows | Export-Csv -Path $diskCsv -NoTypeInformation -Encoding UTF8 -Force
    $logicalCsv = Join-Path $configFolder "hw-logicaldisks.csv"; $logicalRows | Export-Csv -Path $logicalCsv -NoTypeInformation -Encoding UTF8 -Force
    $netCsv = Join-Path $configFolder "hw-netadapters.csv"; $netRows | Export-Csv -Path $netCsv -NoTypeInformation -Encoding UTF8 -Force
    $usbCsv = Join-Path $configFolder "hw-usbcontrollers.csv"; $usbRows | Export-Csv -Path $usbCsv -NoTypeInformation -Encoding UTF8 -Force
    $serialCsv = Join-Path $configFolder "hw-serialports.csv"; $serialRows | Export-Csv -Path $serialCsv -NoTypeInformation -Encoding UTF8 -Force

    # Try to produce hwinfo.xlsx with multiple sheets (ImportExcel)
    $excelOk = $false
    try {
        if (Try-InstallImportExcel) {
            Import-Module ImportExcel -ErrorAction Stop
            $wb = Join-Path $configFolder "hwinfo.xlsx"
            $summaryRow | Export-Excel -Path $wb -WorksheetName 'Summary' -AutoSize -TableName 'Summary' -FreezeTopRow -ClearSheet
            if ($cpuRows.Count -gt 0) { $cpuRows | Export-Excel -Path $wb -WorksheetName 'Processors' -AutoSize -TableName 'Processors' -Append }
            if ($memRows.Count -gt 0) { $memRows | Export-Excel -Path $wb -WorksheetName 'Memory' -AutoSize -TableName 'Memory' -Append }
            if ($gpuRows.Count -gt 0) { $gpuRows | Export-Excel -Path $wb -WorksheetName 'VideoControllers' -AutoSize -TableName 'Video' -Append }
            if ($diskRows.Count -gt 0) { $diskRows | Export-Excel -Path $wb -WorksheetName 'Disks' -AutoSize -TableName 'Disks' -Append }
            if ($logicalRows.Count -gt 0) { $logicalRows | Export-Excel -Path $wb -WorksheetName 'LogicalDisks' -AutoSize -TableName 'LogicalDisks' -Append }
            if ($netRows.Count -gt 0) { $netRows | Export-Excel -Path $wb -WorksheetName 'NetworkAdapters' -AutoSize -TableName 'Network' -Append }
            if ($usbRows.Count -gt 0) { $usbRows | Export-Excel -Path $wb -WorksheetName 'USBControllers' -AutoSize -TableName 'USB' -Append }
            if ($serialRows.Count -gt 0) { $serialRows | Export-Excel -Path $wb -WorksheetName 'SerialPorts' -AutoSize -TableName 'Serial' -Append }
            Write-Host "hwinfo Excel workbook created: $wb" -ForegroundColor Green
            $excelOk = $true
        } else {
            Show-Warn "ImportExcel module unavailable; only CSVs created."
        }
    } catch {
        Show-Warn ("Writing hwinfo.xlsx failed: " + $_.Exception.Message)
    }
} catch {
    Show-Warn ("hwinfo collection failed: " + $_.Exception.Message)
}
#endregion

#region Summary & Consolidated CSV
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Writing summary and consolidated CSV"
try {
    $summaryFile = Join-Path $configFolder "Summary.txt"
    $summary = @()
    $summary += ("Machine: " + $computerName)
    $summary += ("User: " + $env:USERNAME)
    $summary += ("RunLabel: " + $runFolderName)
    $summary += ("TimestampLocal: " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $summary += ("Rows: " + $result.Count)
    $summary | Set-Content -Path $summaryFile -Encoding UTF8 -Force

    $consolidated = Join-Path $runFolder "Consolidated-Software-And-Env.csv"
    $tableCsv = $result | Select-Object Name,Version,Provider,InstallLocation,MSIProductCode | ConvertTo-Csv -NoTypeInformation
    $tableCsv | Set-Content -Path $consolidated -Encoding UTF8 -Force
    Write-Host "Summary & consolidated CSV written." -ForegroundColor Green
} catch { Show-Warn ("Summary write failed: " + $_.Exception.Message) }
#endregion

#region Finish
$currentStep = $totalSteps
Show-Step -i $currentStep -t $totalSteps -title "Completed"
Write-Progress -Activity "Collect-DevProfile" -Completed -Id 1
Write-Host ""
Write-Host "Completed. Outputs are in:" -ForegroundColor Cyan
Write-Host $runFolder -ForegroundColor Green
if (Test-Path $excelPath) { Write-Host ("CSV -> " + $csvPath + "  ;  Excel -> " + $excelPath) -ForegroundColor Cyan } else { Write-Host ("CSV -> " + $csvPath + "  ;  Excel not produced.") -ForegroundColor Yellow }

# Return object for interactive runs
return $result
#endregion
```

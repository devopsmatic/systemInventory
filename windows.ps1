```powershell
<#
Full_Inventory.ps1
- Software -> C:\<ComputerName>\Software\<RUNLABEL>\
- Hardware -> C:\<ComputerName>\Hardware\<RUNLABEL>\ (single hwinfo.xlsx workbook)
- Usage: .\Full_Inventory.ps1

You ZIP and send us the fodler: C:\<ComputerName>
Example, My Computer name is: MYPC
The folder I have to ZIP is: C:\MYPC and send MYPC.zip
#>

param(
    [switch]$FullMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# Relaunch under Bypass once if process policy not Bypass
if (-not $env:COLLECT_SOFT_BYPASS) {
    try { $procPolicy = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue } catch { $procPolicy = $null }
    if ($procPolicy -ne 'Bypass') {
        Write-Host "Relaunching script with ExecutionPolicy Bypass..." -ForegroundColor Cyan
        $pwshCmd = Get-Command -Name pwsh -ErrorAction SilentlyContinue
        $psCmd = Get-Command -Name powershell -ErrorAction SilentlyContinue
        $exe = if ($pwshCmd) { $pwshCmd.Source } elseif ($psCmd) { $psCmd.Source } else { $null }
        if (-not $exe) {
            Write-Host "Cannot find powershell executable to relaunch. Run manually with:" -ForegroundColor Yellow
            Write-Host "  powershell -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Definition)`"" -ForegroundColor Yellow
            exit 1
        }
        $scriptPath = $MyInvocation.MyCommand.Definition
        $fmArg = if ($FullMode) { '-FullMode' } else { '' }
        # Build command to run in the child process
        $childCmd = '$env:COLLECT_SOFT_BYPASS="1"; &' + " `"$scriptPath`" $fmArg"
        $args = @('-NoProfile','-ExecutionPolicy','Bypass','-Command',$childCmd)
        $p = Start-Process -FilePath $exe -ArgumentList $args -Wait -PassThru
        if ($p) { exit $p.ExitCode } else { exit 0 }
    }
}

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
$totalSteps = 16
$currentStep = 0

$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Preparing output folders"
$computerName = $env:COMPUTERNAME
$runLabel = Get-HumanRunLabel

# Software folder
$baseSoftware = Join-Path "C:\" "$computerName\Software"
if (-not (Test-Path $baseSoftware)) { New-Item -Path $baseSoftware -ItemType Directory -Force | Out-Null }
$softwareRunFolder = Join-Path $baseSoftware $runLabel
New-Item -Path $softwareRunFolder -ItemType Directory -Force | Out-Null

# Hardware folder
$baseHardware = Join-Path "C:\" "$computerName\Hardware"
if (-not (Test-Path $baseHardware)) { New-Item -Path $baseHardware -ItemType Directory -Force | Out-Null }
$hardwareRunFolder = Join-Path $baseHardware $runLabel
New-Item -Path $hardwareRunFolder -ItemType Directory -Force | Out-Null

# Common paths
$softwareCsv = Join-Path $softwareRunFolder "InstalledSoftware.csv"
$softwareXlsx = Join-Path $softwareRunFolder "InstalledSoftware.xlsx"
$hwCsvSummary = Join-Path $hardwareRunFolder "hwinfo.csv"
$hwExcelWorkbook = Join-Path $hardwareRunFolder "hwinfo.xlsx"
$hwErrorFile = Join-Path $hardwareRunFolder "hwinfo-errors.txt"

Write-Host "Software outputs -> $softwareRunFolder" -ForegroundColor Cyan
Write-Host "Hardware outputs -> $hardwareRunFolder" -ForegroundColor Cyan
#endregion

#region Collect: Registry uninstall entries
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

#region Collect: MSI via COM
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Enumerating MSI products (COM)"
try {
    $msi = New-Object -ComObject WindowsInstaller.Installer
    $prods = @($msi.Products())
    $i = 0; $total = $prods.Count
    foreach ($p in $prods) {
        $i++; $pct = [int](($i/$total)*100)
        Write-Progress -Id 2 -Activity "MSI enumeration" -Status ("Processing MSI product $i / $total") -PercentComplete $pct
        try { $name = $msi.ProductInfo($p, "InstalledProductName") -as [string] } catch { $name = $null }
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

#region Collect: Appx packages
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

#region Collect: winget local DB
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

#region Collect: ProgramFiles heuristic (portable apps)
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Scanning Program Files (heuristic)"
try {
    $pfRoots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, (Join-Path $env:LocalAppData 'Programs'))
    foreach ($root in $pfRoots) {
        if ($root -and (Test-Path $root)) {
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

#region Optional: FullMode Get-Package provider enumeration
if ($FullMode) {
    $currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Running Get-Package (provider-by-provider, FULL MODE)"
    try {
        $allPkgs = Get-Package -ErrorAction SilentlyContinue
        if ($allPkgs) {
            $providers = $allPkgs | Group-Object -Property ProviderName
            foreach ($g in $providers) {
                $provName = $g.Name
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
    $grouped = $softwareList | Where-Object { $_.Name } | Group-Object @{Expression = { ($_.Name -as [string]).ToLower() + '|' + ($_.Version -as [string]) } }
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

#region Export Software outputs (all placed directly into software run folder)
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Exporting Software CSV and other files"
try {
    # CSV
    $result | Export-Csv -Path $softwareCsv -NoTypeInformation -Encoding UTF8 -Force
    Write-Host ("InstalledSoftware CSV created: " + $softwareCsv) -ForegroundColor Green
} catch { Show-Err ("CSV export failed: " + $_.Exception.Message) }

# Try Excel for software
try {
    if (Try-InstallImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        $result | Export-Excel -Path $softwareXlsx -WorksheetName "SoftwareInventory" -AutoSize -FreezeTopRow -ErrorAction Stop
        Write-Host ("InstalledSoftware XLSX created: " + $softwareXlsx) -ForegroundColor Green
    } else {
        Show-Warn "Software Excel export skipped (ImportExcel not available or install failed)."
    }
} catch { Show-Warn ("Software Excel export failed: " + $_.Exception.Message) }

# Environment exports (Machine/User/All) -> placed in software folder
try {
    $envMachinePath = Join-Path $softwareRunFolder "Env-System.csv"
    $m = [System.Environment]::GetEnvironmentVariables('Machine')
    $m.GetEnumerator() | ForEach-Object { [PSCustomObject]@{ Scope='Machine'; Name=$_.Key; Value=[string]$_.Value } } | Sort-Object Name | Export-Csv -Path $envMachinePath -NoTypeInformation -Encoding UTF8 -Force

    $envUserPath = Join-Path $softwareRunFolder "Env-User.csv"
    $u = [System.Environment]::GetEnvironmentVariables('User')
    $u.GetEnumerator() | ForEach-Object { [PSCustomObject]@{ Scope='User'; Name=$_.Key; Value=[string]$_.Value } } | Sort-Object Name | Export-Csv -Path $envUserPath -NoTypeInformation -Encoding UTF8 -Force

    $envAllPath = Join-Path $softwareRunFolder "All-EnvironmentVariables.csv"
    $proc = [System.Environment]::GetEnvironmentVariables('Process')
    $combined = @()
    foreach ($k in $m.Keys) { $combined += [PSCustomObject]@{ Scope='Machine'; Name=$k; Value=[string]$m[$k] } }
    foreach ($k in $u.Keys) { $combined += [PSCustomObject]@{ Scope='User'; Name=$k; Value=[string]$u[$k] } }
    foreach ($k in $proc.Keys) { $combined += [PSCustomObject]@{ Scope='Process'; Name=$k; Value=[string]$proc[$k] } }
    $combined | Sort-Object Scope, Name | Export-Csv -Path $envAllPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "Environment exports done." -ForegroundColor Green
} catch { Show-Warn ("Env export failed: " + $_.Exception.Message) }

# Git
try {
    $gitConfigPath = Join-Path $softwareRunFolder "Git-Config.txt"
    $gitSshConfigPath = Join-Path $softwareRunFolder "Git-SSH-Config.txt"
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

# Python: detect interpreters and pip freeze (parallel jobs, per-interpreter file placed in software folder)
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

    $pySummary = Join-Path $softwareRunFolder "Python-Packages-Summary.csv"
    $out = @()
    foreach ($j in Get-Job) {
        $res = Receive-Job -Job $j -ErrorAction SilentlyContinue
        if ($res) {
            $fnSafe = ((Split-Path $res.Interpreter -Leaf) -replace '[^0-9A-Za-z\._-]','')
            $fn = Join-Path $softwareRunFolder ("Python-" + $fnSafe + "-Packages.txt")
            if ($res.Packages) { Set-Content -Path $fn -Value ($res.Packages -join "`n") -Encoding UTF8 -Force } else { Set-Content -Path $fn -Value "No packages returned." -Encoding UTF8 -Force }
            $out += [PSCustomObject]@{ Interpreter=$res.Interpreter; Version=$res.Version; PackagesFile=$fn }
        }
        Remove-Job -Job $j -Force -ErrorAction SilentlyContinue
    }
    if ($out.Count -gt 0) { $out | Export-Csv -Path $pySummary -NoTypeInformation -Encoding UTF8 -Force }
    Write-Host "Python detection complete." -ForegroundColor Green
} catch { Show-Warn ("Python detection failed: " + $_.Exception.Message) }

# NodeJS / npm
try {
    $nodeF = Join-Path $softwareRunFolder "NodeJS.txt"
    $nOut = @()
    $nOut += "node --version:`n" + (($(& node --version 2>&1) -join "`n"))
    $nOut += "`n npm --version:`n" + (($(& npm --version 2>&1) -join "`n"))
    $nOut += "`n npm -g list --depth=0:`n" + (($(& npm list -g --depth=0 2>&1) -join "`n"))
    Set-Content -Path $nodeF -Value $nOut -Encoding UTF8 -Force
    Write-Host "Node info exported." -ForegroundColor Green
} catch { Show-Warn ("Node step failed: " + $_.Exception.Message) }

# VSCode settings & extensions
try {
    $vSettings = Join-Path $softwareRunFolder "VSCode-Settings.json"
    $vExts = Join-Path $softwareRunFolder "VSCode-Extensions.txt"
    $appdata = $env:APPDATA
    $vsSettingsSrc = Join-Path $appdata "Code\User\settings.json"
    if (Test-Path $vsSettingsSrc) { Copy-Item -Path $vsSettingsSrc -Destination $vSettings -Force } else { Set-Content -Path $vSettings -Value "VSCode settings not found." -Encoding UTF8 -Force }
    $codeCmd = Get-Command code -ErrorAction SilentlyContinue
    if ($codeCmd) { $exts = & code --list-extensions 2>$null; if ($exts) { $exts | Set-Content -Path $vExts -Encoding UTF8 -Force } else { Set-Content -Path $vExts -Value "No extensions returned." -Encoding UTF8 -Force } } else { Set-Content -Path $vExts -Value "VSCode CLI not found." -Encoding UTF8 -Force }
    Write-Host "VSCode exports done." -ForegroundColor Green
} catch { Show-Warn ("VSCode export failed: " + $_.Exception.Message) }

# Docker info
try {
    $dockerF = Join-Path $softwareRunFolder "Docker-Info.txt"
    $dv = (& docker --version 2>&1) -join "`n"
    if ($dv -and -not ($dv -match 'not found|is not recognized')) {
        Set-Content -Path $dockerF -Value ((& docker --version 2>&1) -join "`n" + "`n" + ((& docker info 2>&1) -join "`n")) -Encoding UTF8 -Force
    } else { Set-Content -Path $dockerF -Value "Docker not installed or not in PATH." -Encoding UTF8 -Force }
    Write-Host "Docker info exported." -ForegroundColor Green
} catch { Show-Warn ("Docker step failed: " + $_.Exception.Message) }
#endregion

#region Summary & Consolidated CSV (software)
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Writing summary and consolidated CSV (software)"
try {
    $summaryFile = Join-Path $softwareRunFolder "Summary.txt"
    $summary = @()
    $summary += ("Computer: " + $computerName)
    $summary += ("User: " + $env:USERNAME)
    $summary += ("RunLabel: " + $runLabel)
    $summary += ("TimestampLocal: " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $summary += ("SoftwareRows: " + ($result.Count -as [string]))
    $summary | Set-Content -Path $summaryFile -Encoding UTF8 -Force

    $consolidated = Join-Path $softwareRunFolder "Consolidated-Software.csv"
    if ($result -and $result.Count -gt 0) {
        $result | Select-Object Name,Version,Provider,InstallLocation,MSIProductCode | Export-Csv -Path $consolidated -NoTypeInformation -Encoding UTF8 -Force
    } else {
        "Name,Version,Provider,InstallLocation,MSIProductCode" | Out-File -FilePath $consolidated -Encoding UTF8
    }
    Write-Host "Summary & consolidated CSV written." -ForegroundColor Green
} catch { Show-Warn ("Summary write failed: " + $_.Exception.Message) }
#endregion

#region Hardware collection -> single hwinfo.xlsx (or fallback hwinfo.csv) + hwinfo-errors.txt
$currentStep++; Show-Step -i $currentStep -t $totalSteps -title "Collecting Hardware -> single workbook"
try {
    # Prepare hw error file
    "`nHWINFO RUN: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" | Out-File -FilePath $hwErrorFile -Encoding UTF8 -Force
    function Log-HwError { param($m, $ex) $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $m"; if ($ex) { $entry += "`nException: $($ex.Message)`nStack:`n$($ex.StackTrace)`n" } ; Add-Content -Path $hwErrorFile -Value $entry }

    # OS & system
    try { $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop } catch { Log-HwError "Win32_OperatingSystem query failed" $_ ; $os = $null }
    try { $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop } catch { Log-HwError "Win32_ComputerSystem query failed" $_ ; $cs = $null }

    # CPU
    try { $cpus = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop } catch { Log-HwError "Win32_Processor query failed" $_ ; $cpus = @() }
    $cpuRows = foreach ($c in $cpus) {
        [PSCustomObject]@{
            Name = $c.Name
            Manufacturer = $c.Manufacturer
            Cores = $c.NumberOfCores
            LogicalProcessors = $c.NumberOfLogicalProcessors
            MaxClock_MHz = $c.MaxClockSpeed
            AddressWidth = $c.AddressWidth
        }
    }

    # Memory
    try { $mem = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction Stop } catch { Log-HwError "Win32_PhysicalMemory query failed" $_ ; $mem = @() }
    $memRows = foreach ($m in $mem) {
        [PSCustomObject]@{
            DeviceLocator = $m.DeviceLocator
            CapacityBytes = if ($m.Capacity) { [int64]$m.Capacity } else { $null }
            CapacityGB = if ($m.Capacity) { [math]::Round($m.Capacity/1GB,2) } else { $null }
            SpeedMHz = $m.Speed
            Manufacturer = $m.Manufacturer
            PartNumber = $m.PartNumber
        }
    }

    # GPUs
    try { $gpus = Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop } catch { Log-HwError "Win32_VideoController query failed" $_ ; $gpus = @() }
    $gpuRows = foreach ($g in $gpus) {
        [PSCustomObject]@{
            Name = $g.Name
            DriverVersion = $g.DriverVersion
            VideoProcessor = $g.VideoProcessor
            AdapterRAM_Bytes = if ($g.AdapterRAM) { [int64]$g.AdapterRAM } else { $null }
            AdapterRAM_GB = if ($g.AdapterRAM) { [math]::Round($g.AdapterRAM/1GB,2) } else { $null }
        }
    }

    # Disks
    try { $disks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction Stop } catch { Log-HwError "Win32_DiskDrive query failed" $_ ; $disks = @() }
    $diskRows = foreach ($d in $disks) {
        [PSCustomObject]@{
            Model = $d.Model
            InterfaceType = $d.InterfaceType
            MediaType = $d.MediaType
            SizeBytes = if ($d.Size) { [int64]$d.Size } else { $null }
            SizeGB = if ($d.Size) { [math]::Round($d.Size/1GB,2) } else { $null }
            SerialNumber = $d.SerialNumber
        }
    }

    # Logical disks
    try { $logical = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop } catch { Log-HwError "Win32_LogicalDisk query failed" $_ ; $logical = @() }
    $logicalRows = foreach ($ld in $logical) {
        [PSCustomObject]@{
            DeviceID = $ld.DeviceID
            VolumeName = $ld.VolumeName
            FileSystem = $ld.FileSystem
            SizeBytes = if ($ld.Size) { [int64]$ld.Size } else { $null }
            SizeGB = if ($ld.Size) { [math]::Round($ld.Size/1GB,2) } else { $null }
            FreeBytes = if ($ld.FreeSpace) { [int64]$ld.FreeSpace } else { $null }
            FreeGB = if ($ld.FreeSpace) { [math]::Round($ld.FreeSpace/1GB,2) } else { $null }
        }
    }

    # Network & primary IP
    try { $netAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True" -ErrorAction Stop } catch { Log-HwError "Win32_NetworkAdapterConfiguration query failed" $_ ; $netAdapters = @() }
    $netRows = @(); $primaryIP = $null
    foreach ($na in $netAdapters) {
        $ips = if ($na.IPAddress) { $na.IPAddress -join ';' } else { '' }
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

    # USB controllers & serial
    try { $usb = Get-CimInstance -ClassName Win32_USBController -ErrorAction Stop } catch { Log-HwError "Win32_USBController query failed" $_ ; $usb = @() }
    $usbRows = foreach ($u in $usb) { [PSCustomObject]@{ Name=$u.Name; DeviceID=$u.DeviceID; PNPDeviceID=$u.PNPDeviceID } }

    try { $serial = Get-CimInstance -ClassName Win32_SerialPort -ErrorAction Stop } catch { Log-HwError "Win32_SerialPort query failed" $_ ; $serial = @() }
    $serialRows = foreach ($s in $serial) { [PSCustomObject]@{ DeviceID=$s.DeviceID; Name=$s.Name; ProviderType=$s.ProviderType } }

    # BIOS
    try { $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop } catch { Log-HwError "Win32_BIOS query failed" $_ ; $bios = @() }
    $biosInfo = $bios | Select-Object -First 1

    # Summary row flatten
    $summaryRow = [PSCustomObject]@{
        ComputerName = $computerName
        Manufacturer = if ($cs) { ($cs | Select-Object -First 1).Manufacturer } else { $null }
        Model = if ($cs) { ($cs | Select-Object -First 1).Model } else { $null }
        SystemType = if ($cs) { ($cs | Select-Object -First 1).SystemType } else { $null }
        OS = if ($os) { $os.Caption } else { $null }
        OS_Version = if ($os) { $os.Version } else { $null }
        OS_Build = if ($os) { $os.BuildNumber } else { $null }
        CPU = if ($cpuRows.Count -gt 0) { $cpuRows[0].Name } else { $null }
        RAM_GB = if ($cs -and $cs.TotalPhysicalMemory) { [math]::Round($cs.TotalPhysicalMemory/1GB,2) } else { $null }
        PrimaryGPU = if ($gpuRows.Count -gt 0) { $gpuRows[0].Name } else { $null }
        PrimaryDiskModel = if ($diskRows.Count -gt 0) { $diskRows[0].Model } else { $null }
        PrimaryDiskSize_GB = if ($diskRows.Count -gt 0) { $diskRows[0].SizeGB } else { $null }
        PrimaryIP = $primaryIP
        BIOS_Serial = if ($biosInfo) { $biosInfo.SerialNumber } else { $null }
    }

    # Export single hw CSV summary (always)
    try { $summaryRow | Export-Csv -Path $hwCsvSummary -NoTypeInformation -Encoding UTF8 -Force; Write-Host "hwinfo.csv written" -ForegroundColor Green } catch { Log-HwError "Failed writing hwinfo.csv" $_ }

    # Try to write single Excel workbook - ImportExcel required
    $excelOk = $false
    try {
        if (Try-InstallImportExcel) {
            Import-Module ImportExcel -ErrorAction Stop
            $summaryRow | Export-Excel -Path $hwExcelWorkbook -WorksheetName 'Summary' -AutoSize -ClearSheet
            if ($cpuRows.Count -gt 0) { $cpuRows | Export-Excel -Path $hwExcelWorkbook -WorksheetName 'Processors' -AutoSize -Append }
            if ($memRows.Count -gt 0) { $memRows | Export-Excel -Path $hwExcelWorkbook -WorksheetName 'Memory' -AutoSize -Append }
            if ($gpuRows.Count -gt 0) { $gpuRows | Export-Excel -Path $hwExcelWorkbook -WorksheetName 'Video' -AutoSize -Append }
            if ($diskRows.Count -gt 0) { $diskRows | Export-Excel -Path $hwExcelWorkbook -WorksheetName 'Disks' -AutoSize -Append }
            if ($logicalRows.Count -gt 0) { $logicalRows | Export-Excel -Path $hwExcelWorkbook -WorksheetName 'LogicalDisks' -AutoSize -Append }
            if ($netRows.Count -gt 0) { $netRows | Export-Excel -Path $hwExcelWorkbook -WorksheetName 'NetworkAdapters' -AutoSize -Append }
            if ($usbRows.Count -gt 0) { $usbRows | Export-Excel -Path $hwExcelWorkbook -WorksheetName 'USBControllers' -AutoSize -Append }
            if ($serialRows.Count -gt 0) { $serialRows | Export-Excel -Path $hwExcelWorkbook -WorksheetName 'SerialPorts' -AutoSize -Append }
            Write-Host "hwinfo.xlsx written: $hwExcelWorkbook" -ForegroundColor Green
            $excelOk = $true
        } else {
            Log-HwError "ImportExcel unavailable or failed to install; hwinfo.xlsx not created."
        }
    } catch {
        Log-HwError "Failed to write hwinfo.xlsx" $_
    }

    if (-not $excelOk) {
        Write-Host "hwinfo.xlsx not produced. hwinfo.csv and hwinfo-errors.txt are available in: $hardwareRunFolder" -ForegroundColor Yellow
    }

} catch {
    Add-Content -Path $hwErrorFile -Value ("Unhandled hwinfo failure: " + $_.Exception.Message + "`n" + $_.Exception.StackTrace)
    Show-Err "Unhandled hwinfo collection failure. See $hwErrorFile"
}
#endregion

#region Finish
$currentStep = $totalSteps
Show-Step -i $currentStep -t $totalSteps -title "Completed"
Write-Progress -Activity "Collect-DevProfile" -Completed -Id 1
Write-Host ""
Write-Host "Completed." -ForegroundColor Cyan
Write-Host "Software outputs -> $softwareRunFolder" -ForegroundColor Green
Write-Host "Hardware outputs -> $hardwareRunFolder" -ForegroundColor Green
if (Test-Path $softwareXlsx) { Write-Host ("Software Excel -> " + $softwareXlsx) -ForegroundColor Cyan }
if (Test-Path $hwExcelWorkbook) { Write-Host ("Hardware Excel -> " + $hwExcelWorkbook) -ForegroundColor Cyan }
Write-Host "If ImportExcel failed, check hwinfo-errors.txt in the hardware folder." -ForegroundColor Yellow

# Return object for interactive runs
return $result
#endregion
```

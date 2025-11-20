### You have to perform 2 actions:
1. Execute the script
2. ZIP the folder and send us

#### Folder structure for your understanding:
If the PC name is MYPC, the structure would be
```

C:\MYPC
├───Hardware
│   └───2025-NOV-20-20-45-23
│       ├─ hwinfo.xlsx                 # consolidated workbook with sheets: Summary, Processors, Memory, Video, Disks, LogicalDisks, Network, USB, Serial
│       ├─ hwinfo.csv                  # flattened summary row (fallback / quick view)
│       └─ hwinfo-errors.txt           # any hwinfo exceptions / diagnostics
└───Software
    └───2025-NOV-20-20-45-23
        ├─ InstalledSoftware.csv       # deduplicated software inventory (master CSV)
        ├─ InstalledSoftware.xlsx      # Excel export of software inventory
        ├─ Consolidated-Software-And-Env.csv
        ├─ Env-System.csv              # machine environment variables
        ├─ Env-User.csv                # user environment variables
        ├─ All-EnvironmentVariables.csv
        ├─ Git-Config.txt
        ├─ Git-SSH-Config.txt
        ├─ Python-<interpreter>-Packages.txt  # pip freeze outputs per interpreter
        ├─ Python-Packages-Summary.txt
        ├─ NodeJS.txt
        ├─ VSCode-Settings.json
        ├─ VSCode-Extensions.txt
        ├─ Docker-Info.txt
        └─ Summary.txt                 # run summary (rows, timestamp, user, run label)
```

ZIP the parent **C:\MYPC** folder

 ### Executing the script
##### Option 1 — Run without downloading
```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

iwr https://raw.githubusercontent.com/devopsmatic/systemInventory/refs/heads/main/windows_inventory.ps1 -UseBasicParsing | iex

```
or
```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Invoke-WebRequest https://raw.githubusercontent.com/devopsmatic/systemInventory/refs/heads/main/windows_inventory.ps1 -UseBasicParsing | Invoke-Expression
```

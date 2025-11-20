### You have to perform 2 actions:
1. Execute the script
2. ZIP the folder and send us

#### Folder structure for your understanding:
If the PC name is MYPC, the structure would be

C:\MYPC
├───Hardware
│ └───2025-NOV-20-20-45-23
└───Software
└───2025-NOV-20-20-45-23

ZIP the parent **C:\MYPC** folder

 ### Executing the script
##### Option 1 — Run without downloading
```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

iwr https://raw.githubusercontent.com/devopsmatic/systemInventory/refs/heads/main/windows.ps1 -UseBasicParsing | iex

```
or
```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Invoke-WebRequest https://raw.githubusercontent.com/devopsmatic/systemInventory/refs/heads/main/windows.ps1 -UseBasicParsing | Invoke-Expression
```

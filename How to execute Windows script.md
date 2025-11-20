### How to execute Windows script "windows.ps1"

```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

iwr https://raw.githubusercontent.com/devopsmatic/systemInventory/refs/heads/main/windows_inventory.ps1 -UseBasicParsing | iex

```
or
```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Invoke-WebRequest https://raw.githubusercontent.com/devopsmatic/systemInventory/refs/heads/main/windows_inventory.ps1 -UseBasicParsing | Invoke-Expression
```

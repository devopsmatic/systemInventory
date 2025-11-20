### How to execute Windows script "windows.ps1"
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
##### Option 2 — Download first, then run
```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script = "$env:TEMP\windows.ps1"
Invoke-WebRequest "https://raw.githubusercontent.com/devopsmatic/systemInventory/refs/heads/main/windows.ps1" -OutFile $script
powershell -ExecutionPolicy Bypass -File $script

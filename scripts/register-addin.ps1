# Register TEAMPPT add-in (run elevated). ASCII only.
$ErrorActionPreference = 'Continue'
$log = 'c:\Projects\teamppt-addin\register.log'
function W($m) { Add-Content -Path $log -Value ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) }
Set-Content -Path $log -Value '=== register-addin start ==='

$dll = 'c:\Projects\teamppt-addin\src\TeampptAddin\bin\Debug\TeampptAddin.dll'
$regasm = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if (-not (Test-Path $dll))    { W "ERROR: DLL not found: $dll"; exit 1 }
if (-not (Test-Path $regasm)) { W "ERROR: RegAsm not found: $regasm"; exit 1 }

# 1) RegAsm with codebase + tlb
$out = & $regasm $dll /codebase /tlb
W ("RegAsm output: " + ($out -join ' | '))

# 2) Control category (must be added manually; RegAsm drops it)
$catPath = 'Registry::HKEY_CLASSES_ROOT\CLSID\{2D4E6F8A-1B3C-5D7E-9F0A-4C6E8D2B1A3F}\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}'
New-Item -Path $catPath -Force | Out-Null
W ("Control category exists: " + (Test-Path $catPath))

# 3) Add-in registry (HKCU of the elevated-but-same user)
$regPath = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\TeampptAddin.Connect'
New-Item -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name 'FriendlyName' -Value 'TEAMPPT'
Set-ItemProperty -Path $regPath -Name 'Description'  -Value 'TEAMPPT Header Assets Add-in'
Set-ItemProperty -Path $regPath -Name 'LoadBehavior' -Type DWord -Value 3
$lb = (Get-ItemProperty -Path $regPath -Name 'LoadBehavior').LoadBehavior
W ("Addin LoadBehavior=" + $lb)

W '=== register-addin done ==='

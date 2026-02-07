#requires -version 5.1
[CmdletBinding()]param(
  [ValidateSet('Studio','GameReady')][string]$DriverBranch='Studio',
  [int]$TargetVoltageMv=870,
  [int]$TargetClockMhz=1700,
  [int]$TempLimitC=80
)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$ROOT = Split-Path -Parent $MyInvocation.MyCommand.Path
$LOG  = Join-Path $ROOT 'optimize.log'
$BIN  = Join-Path $ROOT 'bin'
New-Item -ItemType Directory -Force -Path $BIN | Out-Null
function Write-Log{param([string]$m,[string]$lvl='INFO'); $ts=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); "[$ts][$lvl] $m" | Tee-Object -FilePath $LOG -Append }
function Assert-Admin{ $p=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent(); if(-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){ Write-Log 'Elevando a Administrador (UAC)...'; $args=@('-NoProfile','-ExecutionPolicy','Bypass','-File',"$PSCommandPath",'-DriverBranch',"$DriverBranch",'-TargetVoltageMv',"$TargetVoltageMv",'-TargetClockMhz',"$TargetClockMhz",'-TempLimitC',"$TempLimitC"); Start-Process -Verb RunAs -FilePath powershell.exe -ArgumentList $args; exit 0 } }
function Get-NvInfo{ $i=[ordered]@{GPUName=$null;DriverVersion=$null;TempC=$null;SMClockMHz=$null}; try{ $smi=Join-Path ${env:ProgramFiles} 'NVIDIA Corporation\NVSMI\nvidia-smi.exe'; if(Test-Path $smi){ $csv=& $smi --query-gpu=name,driver_version,temperature.gpu,clocks.sm --format=csv,noheader,nounits 2>$null; if($csv){ $p=$csv -split ',' | % { $_.Trim() }; $i.GPUName=$p[0]; $i.DriverVersion=$p[1]; if($p.Count -gt 2){$i.TempC=[int]$p[2]}; if($p.Count -gt 3){$i.SMClockMHz=[int]$p[3]} } } }catch{}; if(-not $i.DriverVersion){ try{ $gpu=Get-CimInstance Win32_PnPSignedDriver | ? { $_.DeviceClass -eq 'DISPLAY' -and $_.DriverProviderName -match 'NVIDIA' } | select -First 1; if($gpu){ $i.GPUName=$gpu.DeviceName; $i.DriverVersion=$gpu.DriverVersion } }catch{} }; return $i }
function Set-WinPower{ $ultimate='e9a42b02-d5df-448d-aa00-03f14749eb61'; $high='8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'; $plans=powercfg -L; if($plans -match $ultimate){ powercfg -setactive $ultimate } else { powercfg -setactive $high }; Write-Log 'Plan de energía aplicado: Alto/Ultimate' }
function Disable-NvTelemetry{ foreach($s in @('NvTelemetryContainer','NvContainerNetworkService')){ $svc=Get-Service -Name $s -ErrorAction SilentlyContinue; if($svc){ try{ if($svc.Status -ne 'Stopped'){ Stop-Service $svc -Force -ErrorAction SilentlyContinue } }catch{}; try{ Set-Service $s -StartupType Disabled }catch{}; Write-Log "Servicio deshabilitado: $s" } }; try{ $tasks = schtasks /Query /TN "\NVIDIA\*" /FO LIST 2>$null | Select-String 'TaskName:' | % { ($_ -replace 'TaskName:\s*','').Trim() }; foreach($t in $tasks){ try{ schtasks /Change /TN "$t" /Disable | Out-Null; Write-Log "Tarea deshabilitada: $t" }catch{} } }catch{ Write-Log 'No hay tareas NVIDIA para deshabilitar' } }
function Ensure-NVCPL{ $app=Get-AppxPackage -AllUsers -Name 'NVIDIACorp.NVIDIAControlPanel' -ErrorAction SilentlyContinue; if(-not $app){ Write-Log 'Instalando NVIDIA Control Panel (Store)...'; try{ winget install -e --id 9NF8H0H7WMLT --accept-source-agreements --accept-package-agreements --silent | Out-Null }catch{ Write-Log 'Fallo instalando NVCPL' 'WARN' } } else { Write-Log 'NVIDIA Control Panel presente' } }
function Apply-Global3D{ try{ $npi=Join-Path $BIN 'nvidiaProfileInspector.exe'; if(-not (Test-Path $npi)){ Write-Log 'Descargando NVIDIA Profile Inspector...'; curl.exe -L -o "$npi" https://github.com/Orbmu2k/nvidiaProfileInspector/releases/latest/download/nvidiaProfileInspector.exe | Out-Null } ; $nip=@"
<?xml version="1.0" encoding="utf-16"?>
<ArrayOfProfile xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <Profile>
    <ProfileName>Global profile</ProfileName>
    <Executeables />
    <Settings>
      <Setting><SettingID>0x10F9DC81</SettingID><Value>0x00000001</Value></Setting>
      <Setting><SettingID>0x000FB6C1</SettingID><Value>0x00000000</Value></Setting>
      <Setting><SettingID>0x00FF4D5B</SettingID><Value>0x00000003</Value></Setting>
    </Settings>
  </Profile>
</ArrayOfProfile>
"@; $nipPath=Join-Path $BIN 'global.nip'; $nip | Out-File -FilePath $nipPath -Encoding Unicode -Force; & $npi -import $nipPath | Out-Null; Write-Log 'Ajustes 3D globales aplicados' }catch{ Write-Log "No se pudieron aplicar ajustes 3D: $($_.Exception.Message)" 'WARN' } }
function Install-Afterburner{ $found=(winget list --id Guru3D.Afterburner -e 2>$null | Select-String 'Guru3D.Afterburner'); if(-not $found){ Write-Log 'Instalando MSI Afterburner...'; winget install -e --id Guru3D.Afterburner --accept-source-agreements --accept-package-agreements --silent | Out-Null }; Write-Log 'MSI Afterburner instalado' }
function Create-AB-Profiles{ $script = Join-Path $ROOT 'create_afterburner_profiles.ps1'; $args = "-TargetVoltageMv $TargetVoltageMv -TargetClockMhz $TargetClockMhz -TempLimitC $TempLimitC"; Write-Log 'Creando perfiles de Afterburner (requiere UAC)...'; Start-Process -Verb RunAs -FilePath powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$script`" $args" -Wait }
function Install-Driver-TNUC{ Write-Log 'Instalando TinyNvidiaUpdateChecker...'; winget source update | Out-Null; winget install -e --id Hawaii_Beach.TinyNvidiaUpdateChecker --accept-source-agreements --accept-package-agreements --silent | Out-Null; $cfgDir = Join-Path $env:LOCALAPPDATA 'Hawaii_Beach\TinyNvidiaUpdateChecker'; New-Item -ItemType Directory -Force -Path $cfgDir | Out-Null; $cfg = @"
Check for Updates=true
Minimal install=true
Use Experimental Metadata=true
Driver type=$(if($DriverBranch -eq 'Studio'){'sd'} else {'grd'})
"@; $cfg | Out-File -FilePath (Join-Path $cfgDir 'app.config') -Encoding ASCII -Force; $exe = "$env:USERPROFILE\AppData\Local\Microsoft\WindowsApps\TinyNvidiaUpdateChecker.exe"; if(-not (Test-Path $exe)){ $exe = (Get-Command TinyNvidiaUpdateChecker.exe -ErrorAction SilentlyContinue).Source }; if(-not $exe){ $exe = (Get-ChildItem "$env:LOCALAPPDATA\Hawaii_Beach" -Recurse -Filter TinyNvidiaUpdateChecker.exe -ErrorAction SilentlyContinue | Select-Object -First 1).FullName }; if(-not $exe){ throw 'No se encontró TinyNvidiaUpdateChecker.exe' }; Write-Log 'Descargando e instalando driver NVIDIA (minimal, notebook, WHQL DCH)...'; Start-Process -FilePath $exe -ArgumentList "--quiet --confirm-dl --override-notebook --noprompt" -Wait }

# MAIN
Assert-Admin
Write-Log 'Inicio driver-only pipeline'
$before = Get-NvInfo
Write-Log ("Estado inicial: GPU={0} Driver={1} Temp={2}" -f $before.GPUName,$before.DriverVersion,$before.TempC)

Disable-NvTelemetry
Install-Driver-TNUC
Ensure-NVCPL
Disable-NvTelemetry
Set-WinPower
Apply-Global3D
Install-Afterburner
Create-AB-Profiles

$after = Get-NvInfo
$services = @('NvTelemetryContainer','NvContainerNetworkService') | % { $svc=Get-Service -Name $_ -ErrorAction SilentlyContinue; [ordered]@{Name=$_; Exists=[bool]$svc; Status= if($svc){$svc.Status.ToString()} else {'N/A'} } }
$powerPlan = (powercfg -getactivescheme) -replace '.*:\s*',''
$report = [ordered]@{
  InstalledDriverVersion=$after.DriverVersion
  GPUNDetected=$after.GPUName
  CurrentBoostClockMHz=$after.SMClockMHz
  UndervoltTargetmV=$TargetVoltageMv
  TempBeforeC=$before.TempC
  TempAfterC=$after.TempC
  ActivePowerMode=$powerPlan
  RemainingNVServices=$services
  ProfilesCreated=@('AI (Profile1)','Gaming (Profile2)','Dev (Profile3)')
}
$report | ConvertTo-Json -Depth 5 | Out-File -Encoding utf8 -FilePath (Join-Path $ROOT 'report.json')
Write-Log 'Driver-only pipeline completado'
#requires -version 5.1
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [ValidateSet('Studio','GameReady')]
    [string]$DriverBranch = 'Studio',
    [int]$TargetVoltageMv = 870,
    [int]$TargetClockMhz = 1700,
    [int]$TempLimitC = 80
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Globals
$ROOT = Split-Path -Parent $MyInvocation.MyCommand.Path
$LOG  = Join-Path $ROOT 'optimize.log'
$PKG  = Join-Path $ROOT 'pkg'
$BIN  = Join-Path $ROOT 'bin'
$TMP  = Join-Path $ROOT 'tmp'
$REPORT_JSON = Join-Path $ROOT 'report.json'

New-Item -Force -ItemType Directory -Path $PKG,$BIN,$TMP | Out-Null

function Write-Log {
    param([string]$Message,[string]$Level='INFO')
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts][$Level] $Message"
    $line | Tee-Object -FilePath $LOG -Append
}

function Assert-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Log 'Re-ejecutando con privilegios de administrador (UAC)...'
        $a = @('-NoProfile','-ExecutionPolicy','Bypass','-File',"$PSCommandPath")
        if ($PSBoundParameters.Count) { $PSBoundParameters.GetEnumerator() | ForEach-Object { $a += @("-"+$_.Key, "$($_.Value)") } }
        Start-Process -FilePath "powershell.exe" -ArgumentList $a -Verb RunAs | Out-Null
        exit 0
    }
}

function Get-OSId {
    # NVIDIA API OS IDs: Win10 x64=57, Win11=135
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $ver = [version]$os.Version
    if ($ver.Major -ge 10 -and $ver.Build -ge 22000) { return 135 } else { return 57 }
}

function Get-NvidiaInfo {
    $info = [ordered]@{ GPUName=$null; DriverVersion=$null; TempC=$null; SMClockMHz=$null; PowerW=$null }
    try {
        $smi = Join-Path ${env:ProgramFiles} 'NVIDIA Corporation\NVSMI\nvidia-smi.exe'
        if (Test-Path $smi) {
            $csv = & $smi --query-gpu=name,driver_version,temperature.gpu,clocks.sm,power.draw --format=csv,noheader,nounits 2>$null
            if ($csv) {
                $parts = $csv -split ',' | ForEach-Object { $_.Trim() }
                $info.GPUName = $parts[0]
                $info.DriverVersion = if($parts.Count -gt 1){$parts[1]} else {$null}
                $info.TempC = if($parts.Count -gt 2){[int]$parts[2]} else {$null}
                $info.SMClockMHz = if($parts.Count -gt 3){[int]$parts[3]} else {$null}
                $info.PowerW = if($parts.Count -gt 4){[decimal]$parts[4]} else {$null}
                return $info
            }
        }
    } catch { Write-Log "nvidia-smi no disponible: $($_.Exception.Message)" 'WARN' }
    try {
        $gpu = Get-PnpDevice -PresentOnly | Where-Object { $_.Class -eq 'Display' -and $_.InstanceId -match 'VEN_10DE' } | Select-Object -First 1
        if ($gpu) {
            $drv = Get-CimInstance Win32_PnPSignedDriver | Where-Object { $_.DeviceID -eq $gpu.InstanceId }
            $info.GPUName = $gpu.FriendlyName
            $info.DriverVersion = $drv.DriverVersion
        }
    } catch {}
    return $info
}

function Uninstall-NvidiaApps {
    Write-Log 'Desinstalando NVIDIA App / GeForce Experience (método seguro)...'
    $toRemove = @(
        @{ Name='NVIDIA App'; Winget='NVIDIACorporation.NVIDIAApp' },
        @{ Name='GeForce Experience'; Winget='NVIDIA.GeForceExperience' }
    )
    foreach ($pkg in $toRemove) {
        try {
            $id = $pkg.Winget
            $found = $null
            try { $found = (winget list --id $id -e 2>$null | Select-String $id) } catch {}
            if ($found) {
                Write-Log "Desinstalando via winget: $($pkg.Name)"
                winget uninstall --id $id -e --silent --accept-source-agreements --accept-package-agreements | Out-Null
            } else {
                Write-Log "No instalado (omitido): $($pkg.Name)"
            }
        } catch { Write-Log "Error desinstalando $($pkg.Name): $($_.Exception.Message)" 'WARN' }
    }

    # Telemetría conocida. No tocar NVDisplay.ContainerLocalSystem
    foreach ($s in @('NvTelemetryContainer','NvContainerNetworkService')) {
        $svc = Get-Service -Name $s -ErrorAction SilentlyContinue
        if ($svc) {
            try { if ($svc.Status -ne 'Stopped') { Stop-Service $svc -Force -ErrorAction SilentlyContinue } } catch {}
            try { Set-Service $s -StartupType Disabled } catch {}
            Write-Log "Servicio deshabilitado: $s"
        }
    }

    # Tareas en carpeta \NVIDIA
    try {
        $nvtasks = schtasks /Query /TN "\NVIDIA\*" /FO LIST 2>$null | Select-String 'TaskName:' | ForEach-Object { ($_ -replace 'TaskName:\s*','').Trim() }
        foreach ($t in $nvtasks) { try { schtasks /Change /TN "$t" /Disable | Out-Null; Write-Log "Tarea deshabilitada: $t" } catch {} }
    } catch { Write-Log 'No hay tareas NVIDIA para deshabilitar' 'INFO' }
}

function Get-DriverDownloadUrl {
    param([int]$OsId,[string]$Branch)
    Write-Log 'Resolviendo URL del último controlador WHQL (notebook, DCH) desde NVIDIA...'

    [xml]$seriesXml = Invoke-WebRequest -Uri 'https://www.nvidia.com/Download/API/lookupValueSearch.aspx?TypeID=2' -UseBasicParsing | Select-Object -ExpandProperty Content | Out-String
    $seriesXml = [xml]$seriesXml
    $series = $seriesXml.LookupValueSearch.LookupValues.LookupValue | Where-Object { $_.Name -eq 'GeForce RTX 20 Series (Notebooks)' } | Select-Object -First 1
    if (-not $series) { throw 'No se encontró la serie RTX 20 (Notebooks) en API.' }
    $psid = [int]$series.Value

    [xml]$prodXml = Invoke-WebRequest -Uri 'https://www.nvidia.com/Download/API/lookupValueSearch.aspx?TypeID=3' -UseBasicParsing | Select-Object -ExpandProperty Content | Out-String
    $prodXml = [xml]$prodXml
    $product = $prodXml.LookupValueSearch.LookupValues.LookupValue | Where-Object { $_.Name -eq 'GeForce RTX 2080 SUPER with Max-Q Design' } | Select-Object -First 1
    if (-not $product) { throw 'No se encontró el producto RTX 2080 SUPER Max-Q en API.' }
    $pfid = [int]$product.Value

    $procUrl = "https://www.nvidia.com/Download/processFind.aspx?psid=$psid&pfid=$pfid&osid=$OsId&lid=1&whql=1&lang=en-us&ctk=0&dtcid=1"
    $res = Invoke-WebRequest -Uri $procUrl -UseBasicParsing
    $links = ($res.Links | Where-Object { $_.href -like '*driverResults.aspx*' })
    if (-not $links) { throw 'No se encontraron resultados de drivers en processFind.' }

    $preferred = $links | Where-Object { $_.outerText -match $Branch } | Select-Object -First 1
    if (-not $preferred) { $preferred = $links | Select-Object -First 1 }

    $resultUrl = $preferred.href
    if ($resultUrl -notmatch '^https?://') { $resultUrl = 'https://www.nvidia.com' + $resultUrl }

    $drvPage = Invoke-WebRequest -Uri $resultUrl -UseBasicParsing
    $m = [regex]::Match($drvPage.Content, 'downloadURL\s*:\s*"(?<url>https?://[^"]+?\.exe)"')
    if (-not $m.Success) { $m = [regex]::Match($drvPage.Content, 'https?://[\w\./-]+-notebook-win(?:10|11)-?win11?-64bit-[\w-]*dch[\w-]*\.exe') }
    if (-not $m.Success) { throw 'No se pudo extraer el enlace directo del instalador.' }
    $dl = $m.Groups['url'].Value
    if (-not $dl) { $dl = $m.Value }

    Write-Log "URL de descarga detectada: $dl"
    return $dl
}

function Save-File {
    param([string]$Url,[string]$OutPath)
    Write-Log "Descargando: $Url -> $OutPath"
    Invoke-WebRequest -Uri $Url -OutFile $OutPath -UseBasicParsing
}

function Extract-Driver {
    param([string]$Exe,[string]$OutDir)
    Write-Log "Extrayendo controlador a $OutDir"
    New-Item -Force -ItemType Directory -Path $OutDir | Out-Null
    try {
        $nargs = @('-s','-noreboot','-clean','-extract',"$OutDir")
        Start-Process -FilePath $Exe -ArgumentList $nargs -PassThru -Wait -WindowStyle Hidden | Out-Null
    } catch {
        Write-Log 'Fallo extracción directa, intentando auto-extracción por defecto' 'WARN'
        Start-Process -FilePath $Exe -ArgumentList @('-s','-noreboot','-clean') -PassThru -Wait -WindowStyle Hidden | Out-Null
        if (Test-Path 'C:\NVIDIA\DisplayDriver') {
            $latest = Get-ChildItem 'C:\NVIDIA\DisplayDriver' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($latest) { Copy-Item -Recurse -Force $latest.FullName $OutDir }
        }
    }
    if (-not (Test-Path (Join-Path $OutDir 'Display.Driver'))) { throw 'No se encontró la carpeta Display.Driver tras extracción.' }
}

function Install-Driver-INF {
    param([string]$DriverDir)
    Write-Log "Instalando solo el controlador mediante pnputil desde $DriverDir"
    $infGlob = Join-Path (Join-Path $DriverDir 'Display.Driver') '*.inf'
    $infs = Get-ChildItem $infGlob -ErrorAction SilentlyContinue
    if (-not $infs) { throw 'No se encontraron INF en Display.Driver' }
    foreach ($inf in $infs) {
        try { pnputil /add-driver "$($inf.FullName)" /install | Out-Null } catch { Write-Log "Error instalando $($inf.Name): $($_.Exception.Message)" 'WARN' }
    }
}

function Ensure-NVCPL {
    Write-Log 'Verificando NVIDIA Control Panel (Microsoft Store)...'
    $app = Get-AppxPackage -AllUsers -Name 'NVIDIACorp.NVIDIAControlPanel' -ErrorAction SilentlyContinue
    if (-not $app) {
        Write-Log 'Instalando NVIDIA Control Panel desde Microsoft Store via winget...'
        try { winget install -e --id NVIDIACorporation.NVIDIAControlPanel --accept-source-agreements --accept-package-agreements --silent } catch { Write-Log 'Fallo de winget para NVCPL' 'WARN' }
    } else { Write-Log 'NVIDIA Control Panel presente.' }
}

function Set-WindowsPowerPlan {
    Write-Log 'Configurando modo de energía de Windows: Alto rendimiento/Ultimate'
    $ultimate = 'e9a42b02-d5df-448d-aa00-03f14749eb61'
    $high     = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
    $plans = powercfg -L
    if ($plans -match $ultimate) { powercfg -setactive $ultimate } else { powercfg -setactive $high }
}

function Download-Afterburner {
    Write-Log 'Descargando e instalando MSI Afterburner (silencioso)...'
    $abZip = Join-Path $PKG 'MSIAfterburner.zip'
    $abUrl = 'https://download.msi.com/uti_exe/vga/MSIAfterburnerSetup.zip'
    Save-File -Url $abUrl -OutPath $abZip
    Expand-Archive -Path $abZip -DestinationPath $TMP -Force
    $setup = Get-ChildItem $TMP -Recurse -Filter 'MSIAfterburnerSetup*.exe' | Select-Object -First 1
    if (-not $setup) { throw 'No se encontró el instalador de Afterburner en el ZIP.' }
    Start-Process -FilePath $setup.FullName -ArgumentList '/S' -Wait
}

function Build-Afterburner-Profiles {
    Write-Log 'Creando perfiles de MSI Afterburner (AI / Gaming / Dev) con undervolt y curva de ventilador...'
    $abDir = '${env:ProgramFiles(x86)}\MSI Afterburner'
    if (-not (Test-Path $abDir)) { throw 'MSI Afterburner no está instalado.' }
    $profilesDir = Join-Path $abDir 'Profiles'
    New-Item -Force -ItemType Directory -Path $profilesDir | Out-Null

    $devId = (Get-PnpDevice -PresentOnly | Where-Object { $_.Class -eq 'Display' -and $_.InstanceId -match 'VEN_10DE' } | Select-Object -First 1).InstanceId
    $san = ($devId -replace '[^A-Z0-9&_]','_')

    function New-ProfileCfg([string]$path,[int]$vMv,[int]$clk,[int]$temp,[int]$pl,[string]$fanCurve) {
@"
[Startup]
Format=2
PowerLimit=$pl
ThermalLimit=$temp
ThermalPrioritize=0
CoreClockBoost=0
MemoryClockBoost=0
ApplyFan=1
FanSpeedMode=1
FanCurvePoints=$fanCurve

[Settings]
VFCurveEditor=1
VFPointCount=1
VFPoint0_Voltage=$vMv
VFPoint0_Frequency=$clk
LockVoltageFrequency=1
KBoost=0
"@ | Out-File -FilePath $path -Encoding ASCII -Force
    }

    $fanQuiet    = '30;20,40;25,55;35,65;45,70;55,75;65,80;75,85;85'
    $fanBalanced = '30;25,40;30,55;45,65;55,70;65,75;72,80;78,85;85'
    $fanAggro    = '30;30,40;35,55;50,65;60,70;70,75;78,80;84,85;90'

    New-ProfileCfg (Join-Path $profilesDir 'Profile1.cfg') $TargetVoltageMv $TargetClockMhz $TempLimitC 100 $fanAggro
    New-ProfileCfg (Join-Path $profilesDir 'Profile2.cfg') ($TargetVoltageMv-8) ($TargetClockMhz-25) $TempLimitC 100 $fanBalanced
    New-ProfileCfg (Join-Path $profilesDir 'Profile3.cfg') ($TargetVoltageMv-20) ($TargetClockMhz-75) ($TempLimitC-5) 90  $fanQuiet

    $cfgMain = Join-Path $abDir 'MSIAfterburner.cfg'
    if (-not (Test-Path $cfgMain)) { New-Item -ItemType File -Path $cfgMain -Force | Out-Null }
    $content = Get-Content $cfgMain -ErrorAction SilentlyContinue
    if ($content -notmatch '^ApplyOCAtStartup=') { Add-Content $cfgMain 'ApplyOCAtStartup=1' } else { (Get-Content $cfgMain) -replace '^ApplyOCAtStartup=.*','ApplyOCAtStartup=1' | Set-Content $cfgMain }

    $shell = New-Object -ComObject WScript.Shell
    $desk = [Environment]::GetFolderPath('Desktop')
    $abExe = Join-Path $abDir 'MSIAfterburner.exe'
    @(
        @{Name='GPU - AI Profile'; Args='-Profile1'},
        @{Name='GPU - Gaming Profile'; Args='-Profile2'},
        @{Name='GPU - Dev Profile'; Args='-Profile3'}
    ) | ForEach-Object {
        $sc = $shell.CreateShortcut((Join-Path $desk ("$($_.Name).lnk")))
        $sc.TargetPath = $abExe
        $sc.Arguments = $_.Args
        $sc.Save()
    }
}

function Apply-Nvidia-Global-Settings {
    Write-Log 'Ajustando NVIDIA Control Panel (energía máximo rendimiento, baja latencia desactivado, filtrado texturas alto rendimiento)'
    try {
        $npiUrl = 'https://github.com/Orbmu2k/nvidiaProfileInspector/releases/latest/download/nvidiaProfileInspector.exe'
        $npiExe = Join-Path $BIN 'nvidiaProfileInspector.exe'
        Save-File -Url $npiUrl -OutPath $npiExe
        $nip = @"
<?xml version="1.0" encoding="utf-16"?>
<ArrayOfProfile xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <Profile>
    <ProfileName>Global profile</ProfileName>
    <Executeables />
    <Settings>
      <Setting>
        <SettingID>0x10F9DC81</SettingID>
        <Value>0x00000001</Value>
      </Setting>
      <Setting>
        <SettingID>0x000FB6C1</SettingID>
        <Value>0x00000000</Value>
      </Setting>
      <Setting>
        <SettingID>0x00FF4D5B</SettingID>
        <Value>0x00000003</Value>
      </Setting>
    </Settings>
  </Profile>
</ArrayOfProfile>
"@
        $nipPath = Join-Path $BIN 'global.nip'
        $nip | Out-File -FilePath $nipPath -Encoding Unicode -Force
        & $npiExe -import $nipPath | Out-Null
    } catch { Write-Log "Fallo aplicando ajustes NPI: $($_.Exception.Message)" 'WARN' }
}

function Stress-TestGPU {
    param([int]$Seconds=120)
    Write-Log "Descargando GPUTest (FurMark CLI) y ejecutando prueba corta de $Seconds s..."
    try {
        $gtZip = Join-Path $PKG 'GPUTest.zip'
        $gtUrl = 'https://www.geeks3d.com/dl/get/731'
        Save-File -Url $gtUrl -OutPath $gtZip
        Expand-Archive -Path $gtZip -DestinationPath $BIN -Force
        $gputest = Get-ChildItem $BIN -Recurse -Filter 'gputest.exe' | Select-Object -First 1
        if ($gputest) {
            $smi = Join-Path ${env:ProgramFiles} 'NVIDIA Corporation\NVSMI\nvidia-smi.exe'
            $monLog = Join-Path $ROOT 'stress_mon.csv'
            $job = Start-Job -ScriptBlock {
                param($smi,$log)
                for ($i=0;$i -lt 9999;$i++) {
                    try { & $smi --query-gpu=timestamp,clocks.sm,temperature.gpu,power.draw --format=csv,noheader,nounits | Add-Content $log } catch {}
                    Start-Sleep -Milliseconds 500
                }
            } -ArgumentList $smi,$monLog
            Start-Process -FilePath $gputest.FullName -ArgumentList 'furmark /nogui /noscore /max_fps=0 /msaa=0 /width=1280 /height=720 /duration_ms=' + ($Seconds*1000) -Wait
            Stop-Job $job -Force | Out-Null
            Receive-Job $job -Keep | Out-Null
        }
    } catch { Write-Log "Fallo en stress test: $($_.Exception.Message)" 'WARN' }
}

function Build-Report {
    Write-Log 'Construyendo informe final...'
    $before = $script:BeforeInfo
    $after  = Get-NvidiaInfo
    $services = @('NvTelemetryContainer','NvContainerNetworkService') | ForEach-Object {
        $svc = Get-Service -Name $_ -ErrorAction SilentlyContinue
        [ordered]@{ Name=$_; Exists=[bool]$svc; Status= if($svc){$svc.Status.ToString()} else {'N/A'} }
    }
    $powerPlan = (powercfg -getactivescheme) -replace '.*:\s*',''

    $report = [ordered]@{
        InstalledDriverVersion = $after.DriverVersion
        GPUNDetected            = $after.GPUName
        CurrentBoostClockMHz    = $after.SMClockMHz
        UndervoltTargetmV       = $TargetVoltageMv
        TempBeforeC             = $before.TempC
        TempAfterC              = $after.TempC
        ActivePowerMode         = $powerPlan
        RemainingNVServices     = $services
        ProfilesCreated         = @('AI Profile (Profile1)','Gaming Profile (Profile2)','Dev Profile (Profile3)')
    }
    $report | ConvertTo-Json -Depth 5 | Out-File -FilePath $REPORT_JSON -Encoding UTF8 -Force
    Write-Log "Informe guardado en $REPORT_JSON"
}

# ---- Main ----
Assert-Admin
Write-Log "Iniciando optimización NVIDIA para RTX 2080 Super Max-Q"
$script:BeforeInfo = Get-NvidiaInfo
Write-Log ("Estado inicial: GPU={0} Driver={1} Temp={2}C SMClock={3}MHz" -f $BeforeInfo.GPUName,$BeforeInfo.DriverVersion,$BeforeInfo.TempC,$BeforeInfo.SMClockMHz)

Uninstall-NvidiaApps

$osid = Get-OSId
$dlUrl = Get-DriverDownloadUrl -OsId $osid -Branch $DriverBranch
$drvExe = Join-Path $PKG (Split-Path $dlUrl -Leaf)
Save-File -Url $dlUrl -OutPath $drvExe

$drvOut = Join-Path $TMP 'nv_driver'
Extract-Driver -Exe $drvExe -OutDir $drvOut
Install-Driver-INF -DriverDir $drvOut

Ensure-NVCPL

Set-WindowsPowerPlan
Apply-Nvidia-Global-Settings

Download-Afterburner
Build-Afterburner-Profiles

Stress-TestGPU -Seconds 90

Build-Report

Write-Log 'Proceso completado.'
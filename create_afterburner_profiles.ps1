param(
    [int]$TargetVoltageMv=870,
    [int]$TargetClockMhz=1700,
    [int]$TempLimitC=80
)
$ErrorActionPreference='Stop'
$abDir = "${env:ProgramFiles(x86)}\MSI Afterburner"
if(-not (Test-Path $abDir)){ throw "MSI Afterburner no está instalado en $abDir" }
$profilesDir = Join-Path $abDir 'Profiles'
New-Item -Force -ItemType Directory -Path $profilesDir | Out-Null

function New-ProfileCfg([string]$path,[int]$vMv,[int]$clk,[int]$temp,[int]$pl,[string]$fanCurve){
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
$shortcuts = @(
    @{Name='GPU - AI Profile'; Args='-Profile1'},
    @{Name='GPU - Gaming Profile'; Args='-Profile2'},
    @{Name='GPU - Dev Profile'; Args='-Profile3'}
)
foreach($scDef in $shortcuts){
  $sc = $shell.CreateShortcut((Join-Path $desk ("$($scDef.Name).lnk")))
  $sc.TargetPath = $abExe
  $sc.Arguments = $scDef.Args
  $sc.Save()
}
"Perfiles Afterburner creados"
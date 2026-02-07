$ErrorActionPreference='Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$bin = Join-Path $root 'bin'
New-Item -ItemType Directory -Force -Path $bin | Out-Null
$npiExe = Join-Path $bin 'nvidiaProfileInspector.exe'
Invoke-WebRequest -Uri 'https://github.com/Orbmu2k/nvidiaProfileInspector/releases/latest/download/nvidiaProfileInspector.exe' -OutFile $npiExe -UseBasicParsing

function New-Nip([string]$path,[int]$lowLatency){
$xml = @"
<?xml version="1.0" encoding="utf-16"?>
<ArrayOfProfile xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <Profile>
    <ProfileName>Global profile</ProfileName>
    <Executeables />
    <Settings>
      <!-- Power management mode = Prefer maximum performance -->
      <Setting>
        <SettingID>0x10F9DC81</SettingID>
        <Value>0x00000001</Value>
      </Setting>
      <!-- Low Latency Mode -->
      <Setting>
        <SettingID>0x000FB6C1</SettingID>
        <Value>0x0000000$lowLatency</Value>
      </Setting>
      <!-- Texture filtering - Quality = High performance -->
      <Setting>
        <SettingID>0x00FF4D5B</SettingID>
        <Value>0x00000003</Value>
      </Setting>
    </Settings>
  </Profile>
</ArrayOfProfile>
"@
$xml | Out-File -FilePath $path -Encoding Unicode -Force
}

$aiNip = Join-Path $bin 'ai_profile.nip'
$gmNip = Join-Path $bin 'gaming_profile.nip'
$dvNip = Join-Path $bin 'dev_profile.nip'
New-Nip $aiNip 0
New-Nip $gmNip 1
New-Nip $dvNip 0

$shell = New-Object -ComObject WScript.Shell
$desk = [Environment]::GetFolderPath('Desktop')
@(
  @{Name='GPU - Apply AI Driver Profile'; Path=$aiNip},
  @{Name='GPU - Apply Gaming Driver Profile'; Path=$gmNip},
  @{Name='GPU - Apply Dev Driver Profile'; Path=$dvNip}
) | ForEach-Object {
  $sc = $shell.CreateShortcut((Join-Path $desk ("$($_.Name).lnk")))
  $sc.TargetPath = $npiExe
  $sc.Arguments = "-import `"$($_.Path)`""
  $sc.Save()
}
"Perfiles de driver (NPI) listos"
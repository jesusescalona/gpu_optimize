$ErrorActionPreference='Stop'
$bin = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'bin'
New-Item -ItemType Directory -Force -Path $bin | Out-Null
$npi = Join-Path $bin 'nvidiaProfileInspector.exe'
if(-not (Test-Path $npi)){
  Write-Host 'Descargando NPI...'
  curl.exe -L -o $npi https://github.com/Orbmu2k/nvidiaProfileInspector/releases/latest/download/nvidiaProfileInspector.exe | Out-Null
}
$nip = @"
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
"@
$nipPath = Join-Path $bin 'global.nip'
$nip | Out-File -FilePath $nipPath -Encoding Unicode -Force
& $npi -import $nipPath
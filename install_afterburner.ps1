$ErrorActionPreference='Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$pkg = Join-Path $root 'pkg'
$tmp = Join-Path $root 'tmp'
New-Item -ItemType Directory -Force -Path $pkg,$tmp | Out-Null
$abZip = Join-Path $pkg 'MSIAfterburner.zip'
$abUrl = 'https://download.msi.com/uti_exe/vga/MSIAfterburnerSetup.zip'
Invoke-WebRequest -Uri $abUrl -OutFile $abZip -UseBasicParsing
Expand-Archive -Path $abZip -DestinationPath $tmp -Force
$setup = Get-ChildItem $tmp -Recurse -Filter 'MSIAfterburnerSetup*.exe' | Select-Object -First 1
if(-not $setup){ throw 'MSI Afterburner installer not found.' }
Start-Process -FilePath $setup.FullName -ArgumentList '/S' -Wait
"Installed MSI Afterburner"
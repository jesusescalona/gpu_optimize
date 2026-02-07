$ErrorActionPreference='Stop'
[xml]$a=(Invoke-WebRequest 'https://www.nvidia.com/Download/API/lookupValueSearch.aspx?TypeID=2' -UseBasicParsing).Content
$series=$a.LookupValueSearch.LookupValues.LookupValue | Where-Object { $_.Name -eq 'GeForce RTX 20 Series (Notebooks)' } | Select-Object -First 1
[xml]$b=(Invoke-WebRequest 'https://www.nvidia.com/Download/API/lookupValueSearch.aspx?TypeID=3' -UseBasicParsing).Content
$prod=$b.LookupValueSearch.LookupValues.LookupValue | Where-Object { $_.Name -match '(?i)GeForce RTX 2080\s*SUPER.*Max-Q' } | Select-Object -First 1
"psid=$($series.Value) pfid=$($prod.Value) name=$($prod.Name)"
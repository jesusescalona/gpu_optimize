$procs = Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -match 'optimize_nvidia.ps1' }
foreach($p in $procs){ try{ Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
"Stopped: $($procs.ProcessId -join ', ')"
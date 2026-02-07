$ultimate='e9a42b02-d5df-448d-aa00-03f14749eb61'
$high='8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
$plans = powercfg -L
if($plans -match $ultimate){ powercfg -setactive $ultimate } else { powercfg -setactive $high }
"Power plan aplicado"
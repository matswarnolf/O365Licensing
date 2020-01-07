Set-Location HKLM:

Set-Location "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\"  
$root = get-childitem 

foreach ($branch in $Root) {$path =Get-ItemProperty -path $branch  -ErrorAction SilentlyContinue

$path.DisableADALatopWAMOverride
}
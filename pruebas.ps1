#"Registry::HKEY_USERS\*\Software\microsoft\cryptography\calais\smartcards\FNMT-RCM_SLE_FN-20" | remove-item -Force
#"Registry::HKEY_USERS\*\Software\microsoft\cryptography\calais\smartcards\Software\microsoft\cryptography\calais\smartcards\FNMT-RCM_ST" | Remove-Item -Force
#"Registry::HKEY_USERS\*\Software\microsoft\cryptography\calais\smartcards\Software\microsoft\cryptography\calais\smartcards\FNTM-RCM_TC" | remove-item -force



Clear-Host
Echo "Keep-alive with Scroll Lock..."

$WShell = New-Object -com "Wscript.Shell"

while ($true)
{
  $WShell.sendkeys("{SCROLLLOCK}")
  Start-Sleep -Milliseconds 30
  $WShell.sendkeys("{SCROLLLOCK}")
  Start-Sleep -Seconds 30
}

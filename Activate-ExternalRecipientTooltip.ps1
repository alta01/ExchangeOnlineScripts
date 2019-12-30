##Defines cred variable. Office365 credentials from user - username is email
if ($UserCredential -eq $null)
{
  $UserCredential = Get-Credential
}

#Exchange-Online connection
Set-ExecutionPolicy Unrestricted -Scope Process -Force
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking -AllowClobber

$Mailtip = (Get-OrganizationConfig).MailTipsExternalRecipientsTipsEnabled

if ($Mailtip -eq $False)
  {
    Write-Host "External Recipient Configuration is currently set to False"
    $prompt = Read-Host -Prompt "Would you like to enable the External Mailtip? (Y/N)"
    if ($prompt -eq 'Y')
    {
      Write-Host "Confirmed. Enabling External Mailtip Notifications"
      Set-OrganizationConfig -MailTipsExternalRecipientsTipsEnabled:$true
      Write-Host "Done"
    }

    elseif ($prompt -ne 'Y')
    {
      Remove-Pssession $Session
      Exit
    }
  }

elseif ($Mailtip -eq $null)
{
  Write-Host "Could not retrieve value"
  Remove-PSSession $Session
  Exit
}

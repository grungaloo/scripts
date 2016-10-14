#TO-DO: Error checking, general


# O365 Archiving Script
# This script will take an existing user account and archive it according to Matt's ever popular KB article
# The following steps are taken to archive the account:
# 1. User login in blocked

#Connect to Exchange online and MSOL
$credential = Get-Credential
Import-Module MsOnline
Connect-MsolService -Credential $credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

#Grab a few pieces on info to be used during the script
$UserEmail = Read-Host -Prompt 'Email Address to be archived'
$ArchivedEmail = Read-Host -Prompt 'Email address after archiving'
$ArchivedDisplayName = Read-Host -Prompt 'Display Name after archiving'

#Disable the users login
Set-MsolUser -UserPrincipalName $UserEmail -BlockCredential $true
Echo 'User login blocked'

#TO-DO see why the display name isn't getting changed...
#Change the display name of the user to "Archives - <User Name>"
Set-User $UserEmail -DisplayName $ArchivedDisplayName
Echo 'User display name changed'

#Change the primary email address to the archived address
Set-Mailbox $UserEmail -EmailAddresses $ArchivedEmail
Echo 'User email address changed'

#Convert to a shared mailbox
Set-Mailbox $ArchivedEmail -Type Shared
Echo 'User mailbox converted to shared'

#TO-DO: check if a user has that license assigned before trying to remove it
#Grab the Sku of each license and remove it from the user account
#This will probably throw a bunch of errors because it doesn't check if the user actually has the license assigned - it just tries to remove it
Get-MsolAccountSku | ForEach {
    Set-MsolUserLicense -UserPrincipalName $ArchivedEmail -RemoveLicenses $_.AccountSkuID
}
Echo 'Licenses removed'

#TO-DO: create a function to remove the user from all distro groups
#Might not be worth it, O(n^2) - very slow

#Check if forwarding will be setup
$Forwarding = Read-Host -Prompt 'Do you want to setup forwarding (y/n)?'

If ($Forwarding.CompareTo('y')) {
    $ForwardTo = Read-Host -Prompt 'Enter the email where email will be forwarded'
    $ForwardToName = Read-Host -Prompt 'Enter the name of the forwarding group'
    New-DistributionGroup -Name $ForwardToName -Members $ForwardTo -PrimarySMTPAddress $UserEmail
}



#THIS SHOULD GO AT THE END
#Disconnect and kill exchage session. Msol terminates itself
Remove-PSSession $exchangeSession 

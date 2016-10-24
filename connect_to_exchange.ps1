param($Work)

#restart PS with -noexit, the same script, and 1
if (!$Work){
	powershell -noexit -file $MyInvocation.MyCommand.Path 1
	return
}

#Connect to Exchange online and MSOL
$credential = Get-Credential
Import-Module MsOnline
Connect-MsolService -Credential $credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

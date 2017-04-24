#Script:    Stale_AD_Objects.vbs 
#Purpose:  To check AD for stale computer objects based on date logon criteria and disable / delete 
#Author:   Paperclips     
#Email:    pwd9000@hotmail.co.uk 
#Date:     Oct 2013 
#Comments: Can be scheduled to run e.g. weekly to eleviate manual checks 
#Notes:   
 
$disablerange = (Get-Date).AddDays(-183) 
$deleterange = (Get-Date).AddDays(-365) 

# Disable computer objects and move to disabled OU: 
Get-ADComputer -Property Name,lastLogonDate -Filter {lastLogonDate -lt $disablerange} | Set-ADComputer -Enabled $false 
Get-ADComputer -Property Name,Enabled -Filter {Enabled -eq $False} | Move-ADObject -TargetPath "OU=Disabled Computers,DC=tucson,DC=local" 
 
# Delete Older Disabled computer objects: 
Get-ADComputer -Property Name,lastLogonDate -Filter {lastLogonDate -lt $deleterange} | Remove-ADObject -Confirm:$false -recursive
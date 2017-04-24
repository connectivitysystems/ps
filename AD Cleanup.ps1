################################################################################ 
     # PowerShell routine to move Computers into OU structure based on IP # 
################################################################################ 
 
 
##################### 
# Environment Setup # 
##################### 
 
#Add the Active Directory PowerShell module 
Import-Module ActiveDirectory 
 
#Set the threshold for an "old" computer which will be moved to the Disabled OU 
$old = (Get-Date).AddDays(-183)
 
#Set the threshold for an "very old" computer which will be deleted 
$veryold = (Get-Date).AddDays(-365)
 
 
############################## 
# Set the Location IP ranges # 
############################## 
 
$Site0IPRange = "\b(?:(?:192)\.)" + "\b(?:(?:168)\.)" + "\b(?:(?:0)\.)" + "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))" # 192.168.0.0/24 
$Site1IPRange = "\b(?:(?:192)\.)" + "\b(?:(?:168)\.)" + "\b(?:(?:1)\.)" + "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))" # 192.168.1.0/24 
$Site2IPRange = "\b(?:(?:192)\.)" + "\b(?:(?:168)\.)" + "\b(?:(?:2)\.)" + "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))" # 192.168.2.0/24 
$Site3IPRange = "\b(?:(?:192)\.)" + "\b(?:(?:168)\.)" + "\b(?:(?:3)\.)" + "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))" # 192.168.3.0/24 
$Site4IPRange = "\b(?:(?:192)\.)" + "\b(?:(?:168)\.)" + "\b(?:(?:4)\.)" + "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))" # 192.168.4.0/24 
$Site5IPRange = "\b(?:(?:192)\.)" + "\b(?:(?:168)\.)" + "\b(?:(?:5)\.)" + "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))" # 192.168.5.0/24 
$Site6IPRange = "\b(?:(?:192)\.)" + "\b(?:(?:168)\.)" + "\b(?:(?:6)\.)" + "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))" # 192.168.6.0/24 
 
######################## 
# Set the Location OUs # 
######################## 
 
# Disabled OU 
$DisabledDN = "OU=Disabled Computers,DC=tucson,DC=local" 
 
# OU Locations 
$Site0DN = "OU=PHX,OU=Computers,OU=MyBusiness,DC=tucson,DC=local" 
$Site1DN = "OU=TUC,OU=Computers,OU=MyBusiness,DC=tucson,DC=local" 
$Site2DN = "OU=SA,OU=Computers,OU=MyBusiness,DC=tucson,DC=local" 
$Site3DN = "OU=DEN,OU=Computers,OU=MyBusiness,DC=tucson,DC=local" 
$Site4DN = "OU=PERU,OU=Computers,OU=MyBusiness,DC=tucson,DC=local" 
$Site5DN = "OU=SAC,OU=Computers,OU=MyBusiness,DC=tucson,DC=local" 
$Site6DN = "OU=SLC,OU=Computers,OU=MyBusiness,DC=tucson,DC=local" 
 
############### 
# The process # 
############### 
 
# Query Active Directory for Computers and move the objects to the correct OU based on IP
##  -LDAPFilter "(!(OperatingSystem=*Server*))" - this excludes servers, remove to move them too
Get-ADComputer -LDAPFilter "(!(OperatingSystem=*Server*))" -Properties PasswordLastSet | ForEach-Object { 
 
    # Ignore Error Messages and continue on 
    trap [System.Net.Sockets.SocketException] { continue; } 
 
    # Set variables for Name and current OU 
    $ComputerName = $_.Name 
    $ComputerDN = $_.distinguishedName 
    $ComputerPasswordLastSet = $_.PasswordLastSet 
    $ComputerContainer = $ComputerDN.Replace( "CN=$ComputerName," , "") 
 
    # If the computer is more than 90 days off the network, remove the computer object 
    if ($ComputerPasswordLastSet -le $veryold) {  
        Remove-ADObject -Identity $ComputerDN -Recursive
    } 
 
    # Check to see if it is an "old" computer account and move it to the Disabled\Computers OU 
    if ($ComputerPasswordLastSet -le $old) {  
        $DestinationDN = $DisabledDN 
        Move-ADObject -Identity $ComputerDN -TargetPath $DestinationDN 
    } 
 
    # Query DNS for IP  
    # First we clear the previous IP. If the lookup fails it will retain the previous IP and incorrectly identify the subnet 
    $IP = $NULL 
    $IP = [System.Net.Dns]::GetHostAddresses("$ComputerName") 
 
    # Use the $IPLocation to determine the computer's destination network location 
    # 
    # 
    if ($IP -match $Site0IPRange) { 
        $DestinationDN = $Site0DN 
    } 
    Elseif ($IP -match $Site1IPRange) { 
        $DestinationDN = $Site1DN 
    } 
    ElseIf ($IP -match $Site2IPRange) { 
        $DestinationDN = $Site2DN 
    } 
    ElseIf ($IP -match $Site3IPRange) { 
        $DestinationDN = $Site3DN 
    } 
    ElseIf ($IP -match $Site4IPRange) { 
        $DestinationDN = $Site4DN 
    } 
    ElseIf ($IP -match $Site5IPRange) { 
        $DestinationDN = $Site5DN 
    } 
    ElseIf ($IP -match $Site6IPRange) { 
        $DestinationDN = $Site6DN 
    } 
    Else { 
        # If the subnet does not match we should not move the computer so we do Nothing 
        $DestinationDN = $ComputerContainer     
    } 
 
    # Move the Computer object to the appropriate OU 
    # If the IP is NULL we will trust it is an "old" or "very old" computer so we won't move it again 
    if ($IP -ne $NULL) { 
        Move-ADObject -Identity $ComputerDN -TargetPath $DestinationDN 
    } 
}
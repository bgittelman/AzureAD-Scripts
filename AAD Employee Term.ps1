#Connect to Exchange Online using your Office 365 administrative credentials
$cred = Get-Credential 
$Session = New-Pssession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
Import-Pssession $session
#Connect to SharePoint Online
$SharepointURL1 = Read-Host "Enter your SharePoint company name (found in your SharePoint URL: companyname.sharepoint.com/sites/main you would enter companyname)"
$SharepointURL2 = 'https://' + $SharepointURL1 + '-admin.sharepoint.com'
Connect-SPOService -Url $SharepointURL2 -credential $cred
Connect-PNPOnline -Url $SharepointURL2 -credential $cred
#Connect to O365
Connect-MsolService -Cred $cred
#Connect to AzureAD
Connect-AzureAD -Credential $cred
Write-Host "Connected to O365"
#Set the user variable with the user that is to be offboarded
$Username = Read-Host "Enter the username of the terminated employee (ie employeefirst.last@companyname.com)"
$Managername = Read-Host "Enter the username of the manager to grant OneDrive and Email access set forwarding to (ie managerfirst.last@company.com)"
#set AAD users for certain cmdlets
$User = Get-AzureADUser -ObjectId $Username
$Manager = Get-AZureADUser -ObjectID $Managername
Write-Host "Variables Set"
#set OoO language
$OutOfOfficeBody = @"
Thank you for your message.    
I am currently unavailable, however your message has been forwarded to $($Manager.DisplayName) at $($Manager.UserPrincipalName).  
Thanks!
"@
#Set Sign in Blocked
Set-AzureADUser -ObjectId $user.ObjectId -AccountEnabled $false
Write-Host "Azure Sign-in Blocked"
#Disconnect Existing Sessions
Revoke-SPOUserSession -User $Username -confirm:$False
Revoke-AzureADUserAllRefreshToken -ObjectId $user.ObjectId
Write-Host "Azure & Sharepoint Sessions Disconnected"
#Convert to Shared Mailbox, Grant access to Manager
Set-Mailbox $Username -Type Shared
Add-MailboxPermission -Identity $Username -User $Managername -AccessRights FullAccess -InheritanceType All
Write-Host "Mailbox Converted to Shared"
#Set Out Of Office
Set-MailboxAutoReplyConfiguration -Identity $username -ExternalMessage $OutOfOfficeBody -InternalMessage $OutOfOfficeBody -AutoReplyState Enabled
Write-Host "Out of Office Set"
#Forward e-mails to manager
Set-Mailbox $Username -ForwardingAddress $Manager.UserPrincipalName -DeliverToMailboxAndForward $True -HiddenFromAddressListsEnabled $true
Write-Host "Emails Forwarded"
#Remove From Distribution Groups
$DistributionGroups= Get-DistributionGroup | where { (Get-DistributionGroupMember $_.Name | foreach {$_.PrimarySmtpAddress}) -contains "$Username"}
$DistributionGroups | Select-Object DisplayName,ExchangeObjectID | Out-File $username-DGs.txt
foreach ($dg in $DistributionGroups)
    {
    Remove-DistributionGroupMember $dg.name -Member $Username -Confirm:$false
    }
Write-Host "Removed from Exchange DGs"
#Create OneDrive link for Manager
$OneDriveUrl = Get-PnPUserProfileProperty -Account $username | select PersonalUrl
Set-SPOUser $Manager.UserPrincipalName -Site $OneDriveUrl.PersonalUrl -IsSiteCollectionAdmin:$false
$OneDriveUrl | Out-File $username-onedrive.txt
Write-Host "OneDrive Link Created"
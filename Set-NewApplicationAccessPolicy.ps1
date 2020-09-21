#DLC - Penn - 09/17/2020 - PowerShell Cmdlets used to set Application Access Policy for Scoping MSGraphAPI Application Level Permissions
#ref: https://docs.microsoft.com/en-us/graph/auth-limit-mailbox-access
#
#  1. Ensure Settings.json is filled out with App and Tenant Information 
#  2. Fill in $global:SecGroup (Global) variable below with mail-enabled security group

#region Settings
$settings = get-content .\settings.json | ConvertFrom-Json
#endregion

#region Global Variables
$Global:TenantID = $settings.mailboxSettingsPOCApp.TenantID
$global:AppID = $settings.mailboxSettingsPOCApp.ClientID
$global:AppRedirectURI = $settings.mailboxSettingsPOCApp.RedirectURL
$global:AppSecret = $settings.mailboxSettingsPOCApp.ClientSecretPlainText
$global:SecGroup = 'mail-enabledsecuritygroup@mydomain.com'
#endregion

#set Policy
New-ApplicationAccessPolicy -AppId $global:AppID -PolicyScopeGroupId $global:SecGroup -AccessRight RestrictAccess -Description "Restrict this app to members of distribution group."

#test/verify Policy 
Test-ApplicationAccessPolicy -Identity $global:SecGroup -AppId $global:AppID
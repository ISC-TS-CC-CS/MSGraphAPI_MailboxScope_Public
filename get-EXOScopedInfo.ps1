#DLC - Penn - 09/20/2020 - Microsoft GraphAPI - Scoping application permissions to specific Exchange Online mailboxes
#get-EXOScopedInfo.ps1
# 
# Reference: https://docs.microsoft.com/en-us/graph/auth-limit-mailbox-access
# Note: Using /beta/ GraphAPI URI - can be replaced with /v1.0/
# 
# Requirements:  
#       Azure registered app with Application level permissions assigned - with Admin Consent
#       A mail-enabled security group, with target accounts as members
#       An Application Access Policy (see included Set-NewApplicationAccessPolicy)
# 
# 
#Application level Permissions that work with Scoping:
# Mail.Read
# Mail.ReadBasic
# Mail.ReadBasic.All
# Mail.ReadWrite
# Mail.Send
# MailboxSettings.Read
# MailboxSettings.ReadWrite
# Calendars.Read
# Calendars.ReadWrite
# Contacts.Read
# Contacts.ReadWrite
#
#

#region Settings
$settings = get-content .\settings.json | ConvertFrom-Json
#endregion

#region Global Variables
$Global:TenantID = $settings.mailboxScopeApp.TenantID
$global:AppID = $settings.mailboxScopeApp.ClientID
$global:AppRedirectURI = $settings.mailboxScopeApp.RedirectURL
$global:AppSecret = $settings.mailboxScopeApp.ClientSecretPlainText
#endregion


#region Functions
function get-accesstoken {
    [CmdletBinding()]
    param($ClientID, $redirectURL, $clientSecret, $tenantID)

#very basic - minimal error handling

    $getaccesstokenerr = $null

    try {
        $result = Invoke-RestMethod https://login.microsoftonline.com/$($tenantID)/oauth2/token `
            -Method Post -ContentType "application/x-www-form-urlencoded" `
            -Body @{client_id = $clientId; 
            client_secret     = $clientSecret; 
            redirect_uri      = $redirectURL; 
            grant_type        = "client_credentials";
            resource          = "https://graph.microsoft.com";
            state             = "32"
        } -ErrorVariable getaccesstokenerr
    

        if ($null -ne $result) { return $result }
    }
    catch {
        write-host -f Red "Could not retrieve Auth Token"
        write-host -f Red $getaccesstokenerr
        BREAK
    }
    
}
function get-authheader {
    #stores AuthHeader in global variable '$authHeader'
    #very basic - minimal error handling

    $accesstoken = Get-AccessToken -ClientID $global:AppID -redirectURL $global:AppRedirectURI -clientSecret $global:AppSecret -tenantID $Global:TenantID

    $token = $accesstoken.Access_Token
    $tokenexp = $accesstoken.expires_on

    write-host -f black ""
    write-host -f Magenta "AuthToken Retrieved:"
    write-host -f Magenta "$token"
    write-host -f Magenta "Token Expiration Date:"
    write-host -f Magenta "$tokenexp"

    $global:authHeader = @{
        'Content-Type'  = 'application/json'
        'Authorization' = "Bearer " + $token
        'ExpiresOn'     = $tokenexp
    }

    write-host -f black ""
    write-host -f Magenta "AuthHeader Formed:"
    write-host -f Magenta "$global:authHeader"
    write-host -f black ""

}
function get-mailReadViaGAPI{
    param(
        # Account GUID or UPN
        [Parameter(Mandatory=$true)]
        [string]
        $AccountGUIDorUPN
    )

    #reads email messages - Basic - first 'pull' only, implement @odata.nextlink in production or for more messages
    #basic - no error handling or pre-check for authHeader time

    [URI]$URI = "https://graph.microsoft.com/beta/users/$($AccountGUIDorUPN)/messages"
    Invoke-RestMethod -Method Get -Uri $URI.AbsoluteUri -Headers $authHeader -ContentType "application/json"

}
function get-mailboxSettingsViaGAPI{
    param(
        # UserGUID
        [Parameter(Mandatory=$true)]
        [string]
        $UserGUID
    )

    #retrieves Mailbox Settings
    #basic - no error handling or pre-check for authHeader time

    [uri]$URI = "https://graph.microsoft.com/beta/users/$($UserGuid)/mailboxSettings"
    Invoke-RestMethod -Method Get -Uri $URI.AbsoluteUri -Headers $authHeader -ContentType "application/json"

}
function new-sendEmailViaGAPI{
    param(
        # Account GUID or UPN to Send AS
        [Parameter(Mandatory=$true)]
        [string]
        $AccountGUIDorUPN,
        # Email Subject
        [Parameter(Mandatory=$true)]
        [string]
        $emailSubject,
        # Email Body - as String
        [Parameter(Mandatory=$true)]
        [string]
        $emailBodyAsString,
        # email Content Type - "text" or "HTML" - not Mandatory - default HTML
        [Parameter(Mandatory=$false)]
        [string]
        $emailContentType = "HTML",
        # To: Recipients
        [Parameter(Mandatory=$true)]
        [array]
        $toRecipients,
        # CC Recipients - not Mandatory
        [Parameter(Mandatory=$false)]
        [array]
        $ccRecipients,
        # save To Sent Items Bool flag - not Mandatory - default True
        [Parameter(Mandatory=$false)]
        [bool]
        $saveToSentItems = $true
    )

#DLC 09/21/2020 - unfinished - Hash to PSObject to JSON, to POST

# $fullEmailHash = @{}
# $emailBodyHash = @{}
# $messageBodyHash = @{}
# $toRecipientsHash = @{}
# $toRecipientsAddressHash = @{}
# $ccRecipeintsHash = @{}
# $ccRecipientsAddressHash = @{}
# $saveToSentItemsHash = @{}

# $emailBodyHash = @{"contentType" = $emailContentType;"content" = $emailBodyAsString}
# $messageBodyHash = @{"subject" = $emailSubject;"body" = $emailBodyHash}
# $toAddressPSObject = New-Object psobject 
# $toRecipients | %{$toRecipientsHash["emailAddress"] = @{"address" = $_}}


# if($ccRecipients){$ccRecipients | % {$ccRecipients["address"] = $_}}

# $thisEmail = new-object psobject
# $thisEmail | Add-Member -MemberType NoteProperty -Name "message" -Value $messageBodyHash
# $thisEmail.message | add-member -MemberType NoteProperty -Name "subject" -Value $emailSubject



}
#endregion

#region Process
get-authheader


$positiveTestResult = get-mailboxSettingsViaGAPI -UserGUID '~Fill in User GUID or UPN of member of mail-enabled Security Group~'
#---------------------------------------------
#Should Return Correct Results


$negativeTestResult = get-mailboxSettingsViaGAPI -UserGUID '~Fill in User GUID or UPN of account NOT in mail-enabled Security Group~'
# --------------------------------------------
#Returns 
# "code": "ErrorAccessDenied",
# "message": "Access to OData is disabled."

$falseUnsupportedErrorResult = get-mailboxSettingsViaGAPI -UserGUID '~Fill in User GUID or UPN of account without mailbox license, but has AAD Account~'
# --------------------------------------------
#Returns
# "code": "MailboxNotEnabledForRESTAPI",
# "message": "REST API is not yet supported for this mailbox."
#

#endregion
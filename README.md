# DLC - Penn - 09/17/2020 - MSGraphAPI_MailboxScope

 MS GraphAPI - Application permisson Mailbox Scope

Set-NewApplicationAccessPolicy.ps1 - Used After Azure App is setup and mail-enabled group is created - Creates/Verifies Application Access Policy - Users in Group will be allowed to be 'targets' of GraphAPI Get/Post/Patch/etc

get-EXOScopedInfo.ps1 - Examples of retrieving Mailbox Information using GraphAPI

settings.json - App/Tenant Specific Information used in get-EXOScopedINfo.ps1

Steps:
    1. Azure Admin - Create Azure App with expected security permissions (see get-EXOScopedInfo.ps1 header)
    2. Azure Admin - Create mail-enabled Security Group
    3. Azure Admin - Run/utilize Set-NewApplicationAccessPoliy.ps1 to link App and Security Group
    4. Add Users/Targets as Members to mail-enabled Security Group
    5. Utilize get-EXOScopedInfo.ps1 for verification of Scoping/Access

<!-- #ref: https://docs.microsoft.com/en-us/graph/auth-limit-mailbox-access
 -->

<#
.SYNOPSIS
Create an Azure Access Package including Catalog and Policies for a Team - CreateAccessPackagesForSpecificTeams.ps1

.NOTES
Author:         Jakob Schaefer, thinformatics AG
Version:        1.0
Creation Date:  2021/06/16

.DESCRIPTION
This script can be used to deploy access Packages for existing Teams
For every entry in the mandatory CSV an Access Package + Catalog is generated.
The Access Package Catalog will contain the Office 365 Group if the Team and will be related to the new Access Package.
Two Policies will be created fpr the access Package, one for the internal usage and one for external Guests. 
The external Policy is only enabled if the Team is enabled for external Guests (steered by the Column GuestsEnabled in the CSV), otherwise the external policy will be created but stay disabled.
The Parameters for the Policies and some general naming settings were defined by variables in the first part of the script

This script uses the Microsoft Graph API with delegated Permissions to create the Azure Ressources. 

Prerequisites:
- You need an Account with the Admin Role "User Administrator"
- The Teams for which we create the Access Packages have to exist already
- An AAD P2 License has to be enabled in the tenant and every user who is targeted by the access Packages needs an AAD P2 License also

Prepare Script:
We make use of an app based authentication to configure the entitlement Management. To get this done we have to create an App Registration in the tenant. 
To do so proceed as described here:
1. Login to the Source Tenant as Admin with the Role "Application Administrator permissions" (e.g.) to https://aad.portal.azure.com and navigate to "Azure Active Directory" > App Registrations
2. Create a new registration with "New registration"
3. Determine a name for the new registration. E.g.: "CreateAccessPackagesForTeams"
4. The specified account type should be "Accounts in this organizational directory only"
5. Determine a name for the Redirect URI (Web): https://CreateAccessPackagesForTeams.sourcetenant.onmicrosoft.com
6. In the new created App registration navigate Overview and note the ApplicationID
7. Navigate to "Certificates & secrets".
8. Create a new client secret by click on "+ New client secret"
9. Chose an expiry date range and give it a description (e.g.: Creation Time+Requested User Name: 052020 Jakob)
10. After click on "Add" you will get the client secret. Please note it immediately, you can´t make it visible again later. 
11. Navigate to API Permissions and click on "Add a permission"
12. Chose Microsoft Graph API > Delegated Permissions > and chose the following permissions: Group.Read.All, EntitlementManagement.ReadWriteAll
13. Grant admin consent to the new application by click on "Grant admin consent for ..."
14. Insert the noted ClientID, ClientSecret and Redirect URI in this script at line 66 ff (TBD)

.PARAMETER CSVPath
The path to the CSV which is used as a source to specify the Teams for which we should create Access Packages. We need the colums GroupID, DisplayName and GuestsEnabled in the csv
Example-Content:

GroupID;Displayname;GuestsEnabled
bfb87479-b9be-4261-94eb-cb3fc6087e27;"FT Any Teams";True

.PARAMETER ReportPath
Optional: If a reportPath is submitted, a csv will be generated which will contain the AccessPackage Informations for the processed Teams

.EXAMPLE
Create Access Packages for all Teams in the CSV and deposit a Report in c:\temp\
.\CreateAccessPackagesForSpecificTeams.ps1 -CSVPath c:\temp\TeamList.csv -ReportPath c:\temp\AccessPackageReport.csv
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)][string]$CSVPath,
    [Parameter(Mandatory=$false)][string]$ReportPath
)


# The resource URI
$resource = "https://graph.microsoft.com"
$GraphVersion = "/beta"
$GraphURL=$resource+$GraphVersion

#Source App Data: Fill in Client ID, Secret and RedirectUri here:
$ClientIDSource="e6cb6af1-169c-4b2f-8775-322705ac2b21"
$ClientSecretSource="1dAGkz7Pk1R47._9EZR-V_5Ul-_LcWmF_u"
$redirectUriSource="https://CreateAccessPAckagesForSpecificTEams.uccacademy.onmicrosoft.com"

#Var to allow skipping of a specific owner
$SkippedOwnerID="78e4fc82-8e7e-4d13-af11-3146047060d1"

#Naming Vars
$AccessPackagePrefix="AP-Teams-"
$AccessPackageCatalogPrefix="AP-Catalog-Teams-"
$AccessPackagePolicyPrefix="AP-Policy-Teams-"

#Acess Policy Vars
$InternalPolicyAssignmentDuration="365"
$InternalPolicyAccessReviewDuration="25"
$InternalPolicyAccessReviewReoccurence="annual"
$ExternalPolicyAccessReviewDuration="25"
$ExternalPolicyAssignmentDuration="180"
$ExternalPolicyAccessReviewReoccurence="annual"

#import csv
try {
    $CSVData=import-csv $CSVPath -Delimiter ';' -ErrorAction Stop
    Write-Host -ForegroundColor Gray -Object "INFO: Successfully imported the CSV. It contains $(($CSVData).count) entries"
}
catch {
    Write-Host -ForegroundColor Red -Object "ERROR: An Error occured while importing the csv. Please check the path and content of the csv, then restart this script"
    Read-Host -Prompt "Press any key to exit"
    exit
}

#Region Auth
# Function to popup Auth Dialog Windows Form
Function Get-AuthCode {
    Add-Type -AssemblyName System.Windows.Forms

    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url -f ($Scope -join "%20")) }

    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri        
        if ($Global:uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null

    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
    }

    $output
}

#####################################################################
#################SourceConnection###############################
#####################################################################

Write-Host -ForegroundColor Gray -Object "Action: Please Logon to the AzureAD of the Tenant." # TBD Permissions Admin or any user?
#Connect-AzureAD
# UrlEncode the ClientID and ClientSecret and URL's for special characters 
Add-Type -AssemblyName System.Web
$clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($ClientIDSource)
$clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($ClientSecretSource)
$redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($redirectUriSource)
$resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resource)
$scopeEncoded = [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/user.readwrite.all")

# Get AuthCode
$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientIDSource&resource=$resourceEncoded&prompt=admin_consent&scope=$scopeEncoded"
Get-AuthCode

# Extract Access token from the returned URI
$regex = '(?<=code=)(.*)(?=&)'
$authCode  = ($uri | Select-string -pattern $regex).Matches[0].Value

#Write-output "Received an authCode, $authCode"

#get Access Token
$body = "grant_type=authorization_code&redirect_uri=$redirectUriSource&client_id=$clientIdSource&client_secret=$clientSecretEncoded&code=$authCode&resource=$resource"
$tokenResponse = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
    -Method Post -ContentType "application/x-www-form-urlencoded" `
    -Body $body `
    -ErrorAction STOP

#EndRegion Auth
$tokenSource=$tokenResponse.access_token

#now do stuff  

#fetch general tenant info
$getTenantDefaultDomainURI="https://graph.microsoft.com/v1.0/organization"
$getTenantDefaultDomainQuery=Invoke-RestMethod -Method Get -Uri $getTenantDefaultDomainURI  -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer "+$tokenSource }
$verifieddomains=$getTenantDefaultDomainQuery.value.verifieddomains
$initialdomain=($verifieddomains | where {$_.isInitial -eq $true}).name

#initialize output var
$output = @()

foreach ($entry in $CSVData){
    Write-Host -ForegroundColor Gray -Object "INFO: Start processing Access Package for Team $($entry.Displayname)"
    #Access Package Catalog Creation
    $CreateAccessPackageCatalogURI = $GraphURL+"/IdentityGovernance/entitlementManagement/accessPackageCatalogs"
    $CatalogDisplayName=$AccessPackageCatalogPrefix+$entry.DisplayName
    $AccessPackageCatalogDescription="Access Package Catalog created by a script"
    $CreateAccessPackageCatalogBody =@"
    {
        "displayName": "$CatalogDisplayName",
        "description": "$AccessPackageCatalogDescription",
        "isExternallyVisible": false
        }
"@
    $queryCreateAccessPackageCatalog = Invoke-RestMethod -Method POST -Uri $CreateAccessPackageCatalogURI -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer "+$tokenSource } -Body $CreateAccessPackageCatalogBody
    $AccessPackageCatalogID=$queryCreateAccessPackageCatalog.id

    #Access Package Ressource linking
    $CreateAccessPackageRessourceURI = $GraphURL+"/IdentityGovernance/entitlementManagement/accessPackageResourceRequests"
    $CreateAccessPackageRessourceBody =@"
    {
        "catalogId": "$AccessPackageCatalogID",
        "requestType": "AdminAdd",
        "justification": "Automated addition by a script",
        "accessPackageResource": {
            "originId": "$($entry.GroupID)",
            "originSystem": "AadGroup",
            "resourceType": "O365 Group"
        }
        }
"@
    $queryCreateAccessPackageRessource = Invoke-RestMethod -Method POST -Uri $CreateAccessPackageRessourceURI -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer "+$tokenSource } -Body $CreateAccessPackageRessourceBody
    $AccessPackageRessourceID=$queryCreateAccessPackageRessource.id

    #Get ID of the new connected Access Package Ressource
    $GetAccessPackageRessourcesURI = $GraphURL+"/identityGovernance/entitlementManagement/accessPackageCatalogs/"+$AccessPackageCatalogID+"/accessPackageResources"
    $queryGetAccessPackageRessources = Invoke-RestMethod -Method GET -Uri $GetAccessPackageRessourcesURI -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer "+$tokenSource }
    $RessourceID=$queryGetAccessPackageRessources.value.id

    #Access Package creation
    $CreateAccessPackageURI = $GraphURL+"/IdentityGovernance/entitlementManagement/accessPackages"
    $AccessPackageDisplayName=$AccessPackagePrefix+$entry.Displayname
    $AccessPackageDescription="Access Package created by a script"

    $CreateAccessPackageBody =@"
    {
        "catalogId": "$AccessPackageCatalogID",
        "displayName": "$AccessPackageDisplayName",
        "description": "$AccessPackageDescription",
        "isHidden": true
        }
"@

    $queryCreateAccessPackage = Invoke-RestMethod -Method POST -Uri $CreateAccessPackageURI -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer "+$tokenSource } -Body $CreateAccessPackageBody
    $AccessPackageID=$queryCreateAccessPackage.id


    #Region Access Package Policy
    #not needed right now
    #internal users policy
    #get team owners to make them to the backup approver
    #allow users from the own tenant to request Access to the group without an approval, but with business justification

    $GetTeamOwnerURI = $GraphURL+"/groups/"+$entry.GroupID+"/owners"
    $GetTeamOwnersQuery = Invoke-RestMethod -Method Get -Uri $GetTeamOwnerURI -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer "+$tokenSource }
    $TeamOwners= $GetTeamOwnersQuery.Value
    #Remove Owner from array we need to skip
    $TeamOwners= $TeamOwners | where {$_.id -ne $SkippedOwnerID}

    $CreateAccessPackagePolicyURI = $GraphURL+"/IdentityGovernance/entitlementManagement/accessPackageAssignmentPolicies"
    $AccessPackagePolicyDisplayName=$AccessPackagePolicyPrefix+$entry.Displayname+"-Internal"
    $AccessPackagePolicyDescription="Access Package Policy created by a script"
    $AssignmentDuration=$InternalPolicyAssignmentDuration
    $ReviewReoccurence=$InternalPolicyAccessReviewReoccurence
    $ReviewDuration=$InternalPolicyAccessReviewDuration

    $CreateAccessPackagePolicyBody =@"
        {
            "accessPackageId": "$AccessPackageID",
            "displayName": "$AccessPackagePolicyDisplayName",
            "description": "$AccessPackagePolicyDescription",
            "canExtend": "true",
            "durationInDays": "$AssignmentDuration",
            "requestorSettings": {
                "scopeType": "AllExistingDirectoryMemberUsers",
                "acceptRequests": true
            },
            "requestApprovalSettings": {
                "isApprovalRequired": true,
                "isApprovalRequiredForExtension": false,
                "isRequestorJustificationRequired": false,
                "approvalMode": "SingleStage",
                "approvalStages": [
                    {
                        "approvalStageTimeOutInDays": 14,
                        "isApproverJustificationRequired":false,
                        "isEscalationEnabled": false,
                        "escalationTimeInMinutes": 0,
                        "primaryApprovers": [
"@               
        foreach ($owner in $TeamOwners){
            $CreateAccessPackagePolicyBody+=@"
            {
                "@odata.type": "#microsoft.graph.singleUser",
                "isBackup": false,
                "id": "$($owner.id)"
            },
"@
        }
        $CreateAccessPackagePolicyBody+= @"
                        ]
                    }
                ]
            },
            "accessReviewSettings": {
                "isEnabled": true,
                "recurrenceType": "$ReviewReoccurence",
                "reviewerType": "Reviewers",
                "startDateTime": "$(get-date -format o)",
                "durationInDays": $ReviewDuration,
                "reviewers": [
"@               
                            foreach ($owner in $TeamOwners){
                                $CreateAccessPackagePolicyBody+=@"
                                {
                                    "@odata.type": "#microsoft.graph.singleUser",
                                    "isBackup": false,
                                    "id": "$($owner.id)"
                                },
"@
                            }
                            $CreateAccessPackagePolicyBody+= @"
                    
                ]
            }
        }
"@

    $queryCreateAccessPackagePolicy= Invoke-RestMethod -Method POST -Uri $CreateAccessPackagePolicyURI -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer "+$tokenSource } -Body $CreateAccessPackagePolicyBody
    #EndRegion Access Package Policy


    #external users policy
    $CreateAccessPackagePolicyURI = $GraphURL+"/IdentityGovernance/entitlementManagement/accessPackageAssignmentPolicies"
    $AccessPackagePolicyDisplayName=$AccessPackagePolicyPrefix+$entry.DisplayName+"-Guests"
    $AccessPackagePolicyDescription="Access Package Policy created by a script"
    $AssignmentDuration=$ExternalPolicyAssignmentDuration
    $ReviewDuration=$ExternalPolicyAccessReviewDuration
    $ReviewReoccurence=$ExternalPolicyAccessReviewReoccurence
    #$reviewers="" #Insert the Owners of the Team here. #the reviewerType is Reviewers, this collection specifies the users who will be reviewers, either by ID or as members of a group, using a collection of singleUser and groupMembers.

    $CreateAccessPackagePolicyBody =@"
    {
        "accessPackageId": "$AccessPackageID",
        "displayName": "$AccessPackagePolicyDisplayName",
        "description": "$AccessPackagePolicyDescription",
        "canExtend": "true",
        "durationInDays": "$AssignmentDuration",
        "requestorSettings": {
            "scopeType": "AllExternalSubjects",
            "acceptRequests": $([string]$entry.GuestsEnabled.ToLower())
        },
        "requestApprovalSettings": {
            "isApprovalRequired": true,
            "isApprovalRequiredForExtension": true,
            "isRequestorJustificationRequired": true,
            "approvalMode": "SingleStage",
            "approvalStages": [
                {
                    "approvalStageTimeOutInDays": 14,
                    "isApproverJustificationRequired": false,
                    "isEscalationEnabled": false,
                    "escalationTimeInMinutes": 0,
                    "primaryApprovers": [
"@               
    foreach ($owner in $TeamOwners){
        $CreateAccessPackagePolicyBody+=@"
        {
            "@odata.type": "#microsoft.graph.singleUser",
            "isBackup": false,
            "id": "$($owner.id)"
        },
"@
    }
    $CreateAccessPackagePolicyBody+= @"

                    ]
                }
            ]
        },
        "accessReviewSettings": {
            "isEnabled": true,
            "recurrenceType": "$ReviewReoccurence",
            "reviewerType": "Reviewers",
            "startDateTime": "$(get-date -format o)",
            "durationInDays": $ReviewDuration,
            "reviewers": [
"@               
                        foreach ($owner in $TeamOwners){
                            $CreateAccessPackagePolicyBody+=@"
                            {
                                "@odata.type": "#microsoft.graph.singleUser",
                                "isBackup": false,
                                "id": "$($owner.id)"
                            },
"@
                        }
                        $CreateAccessPackagePolicyBody+= @"
                
            ]
        }
    }
"@

    $queryCreateAccessPackagePolicy= Invoke-RestMethod -Method POST -Uri $CreateAccessPackagePolicyURI -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer "+$tokenSource } -Body $CreateAccessPackagePolicyBody


    #Ressource Role Scope (here: which Teams)
    $CreateAccessPackageRessourceRoleScopeURI = $GraphURL+"/IdentityGovernance/entitlementManagement/accessPackages/"+$AccessPackageID+"/accessPackageResourceRoleScopes"

    $originID="Member_"+$($entry.GroupID)

    $CreateAccessPackageResourceRoleScopeBody=@"
    {
        "accessPackageResourceRole": {
            "displayName": "Member",
            "description": "anythint",
            "originSystem": "AadGroup",
            "originId": "$OriginID",
            "accessPackageResource": {
                "id": "$RessourceID",
                "originSystem": "AadGroup"
            }
        },
        "accessPackageResourceScope": {
            "originId": "$($entry.GroupID)",
            "originSystem": "AadGroup"
        }
    }
"@

    $queryCreateAccessPackagePolicy= Invoke-RestMethod -Method POST -Uri $CreateAccessPackageRessourceRoleScopeURI  -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer "+$tokenSource } -Body $CreateAccessPackageResourceRoleScopeBody

    #generate output
    $line =  New-Object psobject
    $line | add-member -Membertype NoteProperty -Name TeamDisplayName -value $entry.Displayname
    $line | add-member -Membertype NoteProperty -Name GroupID -value $entry.GroupID
    $line | add-member -Membertype NoteProperty -Name AccessPackageID -value $AccessPackageID
    $line | add-member -Membertype NoteProperty -Name AccessPackageURI -value ("https://myaccess.microsoft.com/@"+$initialdomain+"#/access-packages/"+$AccessPackageID)
    $output+=$line

    Write-Host -ForegroundColor Gray -Object "INFO: Finished processing Access Package for Team $($entry.Displayname)"
    Write-Host -ForegroundColor Gray -Object "-----------------------------------------------------------------------"
}

write-output $output

if ($ReportPath) {
    $output |export-csv -Path $ReportPath -Delimiter ';' -Encoding Unicode -NoClobber -NoTypeInformation    
}
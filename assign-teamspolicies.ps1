<#
    .SYNOPSIS
    Teams Policy assignment - assign-teamspolicies.ps1
   
    Jakob Schaefer
	
	02.04.2020
	
    .DESCRIPTION
    This script should help to assign Teams Policies to Teams users based on theire Group Membership or AD Attributes
    
   	PARAMETER PolicyType
    Name of the PolicyTypes that should be assigned. Multiple Values are possible
    PARAMETER PolicyName
    Name of the Policy that should be assigned. If different PolicyTypes should be assigned at once, the names of the different Policies have to be equal
    PARAMETER scope
    Decide if the policy should be assigned to members of a group or to accounts with specific userattributes
    PARAMETER Group
    If Scope = Group: Name of the Group, to which members the policies should be assigned
    PARAMETER UserFilterAttribute
    If Scope = User: Name of the attribute which should be used to filter the affected users
    PARAMETER UserFilterValue
    If Scope = User: Value of the attribute that should be used to filter the affected users
             
	EXAMPLES
    .\assign-teamspolicies.ps1 -PolicyType MeetingPolicy,MessagingPolicy -PolicyName "Lehrer" -Scope Group -GroupName "Alle Lehrer"
    .\assign-teamspolicies.ps1 -PolicyType AppPermissionPolicy -PolicyName "Erweiterter App Katalog" -Scope UserFilter -UserFilterAttribute "jobtitle" -UserFilterValue "Teams Pro"
#>

[CmdletBinding()]
param (
    [ValidateSet("AppPermissionPolicy", "AppSetupPolicy", "CallingPolicy", "CallParkPolicy", "ChannelsPolicy", "ComplianceRecordingPolicy", "EmergencyCallingPolicy", "EmergendyCallRoutingPolicy", "FeedbackPolicy", "IPPhonePolicy", "MeetingBroadcastPolicy", "MeetingPolicy", "MessagingPolicy", "UpgradePolicy", "VerticalPackagePolicy", "VideoInteropServicePolicy")]
    [Parameter(Mandatory=$true)][string[]]$PolicyType,
    [Parameter(Mandatory=$true)][string]$PolicyName,
    [ValidateSet("Group","UserFilter")]
    [Parameter(Mandatory=$true)][string]$Scope,
    [Parameter(Mandatory=$false)][string]$GroupName,
    [Parameter(Mandatory=$false)][string]$UserFilterAttribute,
    [Parameter(Mandatory=$false)][string]$UserFilterValue
)

try {
    $installedAzureADModule=get-module -ListAvailable -Name AzureAD -ErrorAction Stop
}
catch {
    Write-Host -ForegroundColor Red -Object "There is no Azure AD Powershell Module installed, Please install it."
    Read-Host -Prompt "Press any key to exit"
    exit
}

try {
    $installedSkypeOnlineModule=get-module -ListAvailable -Name SkypeOnlineConnector -ErrorAction Stop
    if(!($installedSkypeOnlineModule.Version.Major -ge 7)){
        Write-Host -ForegroundColor Red -Object "The installed Skype Online Module doesn´t meet the requirements. Please install it (min V 1.0.20). Install it from https://www.microsoft.com/de-de/download/details.aspx?id=39366"
        Read-Host -Prompt "Press any key to exit"
        exit
    }
}
catch {
    Write-Host -ForegroundColor Red -Object "There is no Skype Online Module installed, Please install it (min V 1.0.20). Install it from https://www.microsoft.com/de-de/download/details.aspx?id=39366"
    Read-Host -Prompt "Press any key to exit"
    exit
}

Write-Host -ForegroundColor Gray -Object "ACTION: Please enter Admin Credentials to login to AzureAD"
$SessionAAD=Connect-AzureAD

Write-Host -ForegroundColor Gray -Object "ACTION: Please enter Admin Credentials to login to Skype Online"
$sfbSession = New-CsOnlineSession
Import-PSSession $sfbSession -AllowClobber

if($GroupName -AND ($Scope -eq "Group")){
    $group = Get-AzureADGroup -SearchString $GroupName
    if($group){
        Write-Host -ForegroundColor "Gray" -Object "INFO: The Members of the Group $GroupName will be processed"
        $members = Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true | Where-Object {$_.ObjectType -eq "User"} 
    }else {
        write-host -ForegroundColor Red -Object "ERROR: No Group Found with the Name $GroupName"
    }
}

if($UserFilterAttribute -AND $UserFilterValue -AND ($Scope -eq "UserFilter")){
    $members = Get-AzureADUser -Filter "$UserFilterAttribute eq '$UserFilterValue'"
}

if($members){
    foreach ($PolicyTypeValue in $PolicyType){
        $command="Get-CSTeams"+$PolicyTypeValue+ " "+$PolicyName
        $policy = invoke-expression -Command $command 
        if ($policy){
            foreach ($member in $members){
                try{
                        $command= "Grant-CsTeams"+ $PolicyTypeValue+ " '" + $policy.Identity +"' -Identity " + $member.UserPrincipalName + " -ErrorAction Stop"
                        invoke-expression -Command $command 
                        Write-Host -ForegroundColor Gray -Object "INFO: The Policy $PolicyName (Type: $PolicyTypeValue) was assigned to the user $($member.UserPrincipalName) successfully"
                }catch{
                    Write-Host -ForegroundColor RED -Object "ERROR: An Error occured assigning the App Permission Policy $PolicyName to the user $($member.UserPrincipalName)"
                }
            }
        }else{
            write-host -ForegroundColor Red -Object "ERROR: No Policy found with the Name $PolicyName"    
        }
    }
}

Disconnect-AzureAD 
get-pssession | Remove-Pssession
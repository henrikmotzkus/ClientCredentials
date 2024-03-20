<#
.SYNOPSIS
  This script takes a Azure DevOPS User Story Export file (See template xlsx) and imports it to M365 Planner
.DESCRIPTION
  First you need to export all your userstories from Azure DevOPS into an Excel (See template). 
  Use the Visual Studio Office Integration https://learn.microsoft.com/de-de/azure/devops/boards/backlogs/office/track-work?view=azure-devops&tabs=open-excel
  You need a User with appropriate permissions: "User.Read", "Group.ReadWrite.All", "Group.Read.All", "Tasks.Read", "Tasks.ReadWrite"
  Column Description and Acceptance Criteria will be merged.
  You get the planID from Browser Url when you open the planner.
.PARAMETER <Parameter_Name>
  -PlanId   -> Every PLan in M365 Planner has an ID. Use this ID here.
  -FilePath -> The Path to the Excel
.NOTES
  Version:        1.0
  Author:         Henrik Motzkus
  Creation Date:  19.3.2023
  Purpose/Change: Initial script development
  
  .EXAMPLE
  Upload-TasktoPlanner -PlanID <YOPUR PLAN ID> -FilePath <PATH TO EXCEL FILE>
  
#>
  
  param (
      [String]$PlanId,
      [String]$FilePath
      )
      
$ExcelData = Import-Excel -Path $FilePath
Connect-MgGraph -Scopes "User.Read", "Group.ReadWrite.All", "Group.Read.All", "Tasks.Read", "Tasks.ReadWrite"

function ConvertTo-HashtableFromPsCustomObject { 
    param ( 
        [Parameter(  
            Position = 0,   
            Mandatory = $true,   
            ValueFromPipeline = $true,  
            ValueFromPipelineByPropertyName = $true  
        )] [object] $psCustomObject 
    );
    Write-Verbose "[Start]:: ConvertTo-HashtableFromPsCustomObject"

    $output = @{}; 
    $psCustomObject | Get-Member -MemberType *Property | % {
        $output.($_.name) = $psCustomObject.($_.name); 
    } 
    
    Write-Verbose "[Exit]:: ConvertTo-HashtableFromPsCustomObject"

    return  $output;
}

function CreatePlannerTask {
    param (
        [string]$Title,
        [string]$UserId,
        [string]$Description,
        [string]$AcceptanceCriteria,
        [string]$State
        )
        
    $PlanId = "0PwDWSUrKUGfoXLDDE-XvpcADMgL"
    # Create vanilla task object
    $Task = [PSCustomObject]@{
        PlanId = $PlanId
        Title = $Title
    }

    # If the task is already assigned than add accordingly the member structure
    if ($UserId -ne ""){
        $Task | Add-Member -NotePropertyMembers @{
            assignments = @{
                $UserId = @{
                    "@odata.type" = "#microsoft.graph.plannerAssignment"
                    orderHint = " !"
                }
            }
        }
    }
    
    if ($State -eq "Done"){
        $Task | Add-Member -MemberType NoteProperty -Name "percentComplete" -Value "100"
    }

    if ($State -eq "Active"){
        $Task | Add-Member -MemberType NoteProperty -Name "percentComplete" -Value "50"
    } 

    try {
        # Try to create the task 
        $body = ConvertTo-HashtableFromPsCustomObject $Task
        $Response = $null
        $Response = New-MgPlannerTask -BodyParameter $body
        # Description can only be added with update-methode
        if ($Description -ne ""){
            $d = $Description + $AcceptanceCriteria
            $taskdetail = Get-MgPlannerTaskDetail -PlannerTaskId $Response.Id
            $etag = $taskdetail.AdditionalProperties.'@odata.etag'
            Update-MgPlannerTaskDetail -PlannerTaskId $Response.id -Description $d -IfMatch $etag
        }
    } catch {
        Write-Host "An error occurred:"
        Write-Host $_.ScriptStackTrace
    }
}

$ExcelData | foreach {
    try {
        CreatePlannerTask -Title $_.title -UserId $_.User_Object_ID -Description $_.Description -AcceptanceCriteria $_.'Acceptance Criteria' -State $_.State
    }
    catch {
        Write-Host "An error occurred:"
        Write-Host $_.ScriptStackTrace
        break
    }
}
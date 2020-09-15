
<#
.SYNOPSIS
     This script create collections based on each found Boundary groups in an environment.

.OUTPUTS
     New ConfigMan Device Collection. 

.NOTES
     Version:        1.0
     Author:         Richard Westerlund
     Creation Date:  09/08/2020
     Purpose/Change: Initial script development
  
#>
function New-BoundaryGroupCollection
{
     [CmdletBinding()]
     param(
         [Parameter(Mandatory=$true)]
         [String]$CollectionName,

         [Parameter(Mandatory=$true)]
         [String]$LimitingCollection,

         [Parameter(Mandatory=$true)]
         [String]$BoundaryGroupName,

         [Parameter(Mandatory=$false)]
         [String]$ConsolePath
     )
     
     write-host "Creating ConfigMan Collection: $($CollectionName)..." 
     write-host "   Limiting Collection: $LimitingCollection" -foregroundColor Yellow

     # Create all necessary internal Variables
     $schedule = New-CMSchedule -RecurInterval Days -recurCount 1
     $Query = "SELECT SMS_R_SYSTEM.ResourceID FROM SMS_R_System where SMS_R_System.ResourceId IN (select resourceid FROM SMS_CollectionMemberClientBaselineStatus WHERE SMS_CollectionMemberClientBaselineStatus.boundarygroups Like '%"+$BoundaryGroupName+"%') AND SMS_R_System.Name NOT IN ('Unknown') AND SMS_R_System.Client = '1'"

     # Create Device Collection
     try 
     {
          New-CMDeviceCollection -LimitingCollectionName $LimitingCollection -Name $CollectionName -RefreshType 2 -RefreshSchedule $schedule -errorAction Stop | Out-Null 
          write-Host "   Collection Created!" -ForegroundColor green
     }
     catch 
     {
          
          write-host "   Error Creating $CollectionName with the following exception:" -foregroundcolor Red
          Write-Host "$($_.Exception)" -ForegroundColor Red
          write-host ""

          continue
     }
     
     Write-Host "   Adding Membership Query Rule From BoundaryGroup: $($BoundaryGroupname)"

     try 
     {
          Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionName -QueryExpression $Query -RuleName "BoundaryGroup - $($BoundaryGroupName)" -errorAction Stop 
          Write-Host "   Membership Rule Added!" -foregroundColor Green
     }
     catch 
     {
          write-host "Error Adding Membership Query Rules for: $CollectionName"
          Write-host "$($_.Exception)" -ForegroundColor Red

          Continue
     }
     
     # Move Collection
     if($consolePath)
     {

          write-host "   Moving Collection to: AOP:\DeviceCollection\$($ConsolePath)"
          try 
          {
               Move-CMObject -folderPath "AOP:\DeviceCollection\$($ConsolePath)" -inputObject $(Get-CMDeviceCollection -Name $CollectionName) -ErrorAction Stop
          }
          catch 
          {
               Write-Host "Error Moving $CollectionName to: $ConsolePath" 
               Write-Host "$($_.Exception)" -ForegroundColor Red
               write-host ""
     
               continue
          }
     }
     
     Start-sleep 2

     # YES I AM WAITING 2 SECONDS FOR SEEMINGLY NO REASON.
     # I JUST DON'T WANT ~250 COLLECTIONS ALL UPDATING AT THE SAME TIME AND I DON'T WANT TO USE Get-Random -Max something TO RANDOMIZE A START TIME!
     
}

$BoundaryGroups = Get-CMBoundaryGroup

foreach($boundaryGroup in $BoundaryGroups)
{
     New-BoundaryGroupCollection -CollectionName "BG - $($BoundaryGroup.Name)" -LimitingCollection "ADM - BASE - All Workstations" -BoundaryGroupName $($BoundaryGroup.name) -ConsolePath "Test Collections\Richard's Collections\Boundary Group Collections"
}

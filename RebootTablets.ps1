Import-Module Microsoft.Graph.Identity.DirectoryManagement
Import-Module Microsoft.Graph.DeviceManagement
#TODO: WILL NEED SERVICE ACCOUNT OR APP WITH PERMISSIONS

try{
    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.PrivilegedOperations.All", "Group.Read.All" -NoWelcome
    Write-Output "Graph auth successful...`n"
}catch{
    Write-Output "ERROR connecting to graph- $_"
    exit
}


<#------------------------------------------------------------------- 
                    REBOOT ALL ANDROID TABLETS
--------------------------------------------------------------------#>

<# #get android devices/filter out non-androidEnterpriseDedicatedDevice (this removes a few manually enrolled Androids)
$androidDevices = Get-MgDeviceManagementManagedDevice -All -Filter "OperatingSystem eq 'Android'"
$androidDevices = $androidDevices | Where-Object DeviceEnrollmentType -eq "androidEnterpriseDedicatedDevice"

#loop through devices and send reboot
foreach($tablet in $androidDevices){
    try{
        Restart-MgDeviceManagementManagedDeviceNow -ManagedDeviceId $tablet.Id
        Write-Host "$($tablet.SerialNumber) - reboot initiated"
    }catch{
        Write-Host "Failed to reboot ${$tablet.SerialNumber} - $_"
    }
} #>


<#------------------------------------------------------------------- 
                    REBOOT ALL DEVICES IN A GROUP
--------------------------------------------------------------------#>

#get group info
$groupName = "TESTDEVICES_sg"
#$groupName = "sg-app-Pilot-SCD"
#$groupName = "sg-SCD-WEST"
$group = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$search=`"displayName:$groupName`"" -Headers @{ConsistencyLevel = "eventual"}).Value

#get members of group using group id
[System.Collections.ArrayList]$tabletsInGroup = @()
$url = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/members"
$morePages = $true

while ($morePages) {
    $members = Invoke-MgGraphRequest -Method GET -Uri $url
    $tabletsInGroup += $members.value

    if($members."@odata.nextLink"){
        $url = $members."@odata.nextLink"
        continue
    }
    $morePages = $false
}


foreach($tablet in $tabletsInGroup){
    #get Intune ID from AzureADDeviceID (this is needed for the reboot)
    $tabletIntuneId = (Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=azureADDeviceId+eq+'$($tablet.deviceId)'&`$select=id,serialNumber").Value

    if(-not $tabletIntuneId){
        Write-Host "$($tablet.displayName) - device not found... continuing to next device" -ForegroundColor Red
        continue
    }
    
    #reboot tablet
    #Invoke-GraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($tabletIntuneId.id)/rebootNow"
    Write-Output "$($tabletIntuneId.serialNumber) - reboot initiated"
}



#loop through members of group
<# foreach($tablet in $members){

    try{
        #get Intune ID from AzureADDeviceID (this is needed for the reboot)
        $tabletIntuneId = (Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=azureADDeviceId+eq+'$($tablet.deviceId)'&`$select=id,serialNumber").Value

        if(-not $tabletIntuneId){
            Write-Host "$($tablet.displayName) - device not found... continuing to next device" -ForegroundColor Red
            continue
        }
        
        #reboot tablet
        #Restart-MgDeviceManagementManagedDeviceNow -ManagedDeviceId $tabletIntuneId.id
        Write-Host "$($tabletIntuneId.serialNumber) - reboot initiated" -ForegroundColor Green
        $count++
    }catch{
        Write-Host "ERROR: - $_" -ForegroundColor Red
    }
} #>




#testing paging through to call endpoint with Invoke-MgGraphRequest

<# $url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=OperatingSystem+eq+'Android'"
$tablets
$isMoreTablets = $true
while($isMoreTablets){

    Write-Host $count
    $response = Invoke-MgGraphRequest -Method GET -uri $url
    $tablets += $response.value

    if($response."@odata.nextLink"){
        $url = $response."@odata.nextLink"
        continue
    }

    $isMoreTablets = $false
}
Write-Host "Test" -ForegroundColor Green
$tablets.Count
$tablets = $tablets | Where-Object DeviceEnrollmentType -eq "androidEnterpriseDedicatedDevice"
$tablets.Count #>

$accessToken = ConvertTo-SecureString 
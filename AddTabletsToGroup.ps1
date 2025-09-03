#TAKES ARRAYLIST OF SERIAL NUMBERS AS AN ARG
param(
    [System.Collections.ArrayList]$serialNumberList 
)

#CHECK ARGUMENT IS NOT EMPTY
if(!$serialNumberList){
    Write-Output "Missing serialNumberList arg"
    exit
}

#PROMPT FOR NAME OF GROUP
$groupName = Read-Host -Prompt "Group Name"

#CONNECT TO GRAPH
try{
    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All" -NoWelcome
}catch{
    Write-Output "Failed to connect to graph. Exiting..."
    exit
}

#GET GROUP INFO
$groupInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq `'$groupName`'" -SkipHttpErrorCheck

if(!$groupInfo.value){
    Write-Output "Failed to get group info - $groupName"
    exit
}

$groupId = $groupInfo.value.id

Write-Output "Group gound. Id: $groupId"
Write-Output "Number of tablets: $($serialNumberList.Count)"

[System.Collections.ArrayList]$fails = @()
[System.Collections.ArrayList]$duplicates =@()

#LOOP THROUGH LIST OF SERIALNUMBERS
foreach($serialNumber in $serialNumberList){

    Write-Output "`n"
    Write-Output "-------------------$serialNumber-------------------"
    
    try{
        #GET TABLET INFO (NEED AZURE AD DEVICE ID)
        $tabletInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=serialNumber eq '$serialNumber'" -SkipHttpErrorCheck
        
        $azureDeviceId = $null

        #IF MULTIPLE TABLETS ARE FOUND, FIND THE ONE WITH THE MOST RECENT CHECK IN. ADD DUPLICATES TO DUIPLICATE ARRAY
        if($tabletInfo."@odata.count" -gt 1){
            
            Write-Output "Multiple devices found. $($tabletInfo."@odata.count") Going to try the more recent device check in"
            
            [void]$duplicates.Add($serialNumber)
            [datetime]$mostRecentCheckIn = $tabletInfo.value[0].lastSyncDateTime
            $azureDeviceId = $tabletInfo.value[0].azureActiveDirectoryDeviceId

            #STARTING AT INDEX 1 SINCE I AM SETTING DEFAULT AS INDEX 0
            for($index = 1; $index -lt $tabletInfo."@odata.count"; $index++){

                #IF CURRENT MOST RECENT CHECK IN IS EARLIER THAT CURRENT INDEX, UPDATE DEVICE ID AND MOST RECENT CHECK IN 
                if($mostRecentCheckIn -lt $tabletInfo.value[$index].lastSyncDateTime){
                    $mostRecentCheckIn = $tabletInfo.value[$index].lastSyncDateTime
                    $azureDeviceId = $tabletInfo.value[0].azureActiveDirectoryDeviceId
                }
            }
        }else{
            $azureDeviceId = $tabletInfo.value.azureActiveDirectoryDeviceId
        }
        
        #CALL GRAPH AGAIN BECAUSE WE NEED THE OBJECT ID
        $objectInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$azureDeviceId'&$count=true"
        $objectId = $objectInfo.value.id

        $body = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/devices/$objectId"
        } | ConvertTo-Json

        #ADD TABLET TO THE GROUP
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members/`$ref" -Body $body
    }catch{
        Write-Output "Error - $_"
        [void]$fails.Add($serialNumber)
        continue
    }

    Write-Output "Added to group."
}

if($fails){
    Write-Output "$($fails.Count) failed... Serial Numbers: $falls"
}

if($duplicates){
    Write-Output "$($duplicates.Count) duplicates found... Serial Numbers: $duplicates"
}
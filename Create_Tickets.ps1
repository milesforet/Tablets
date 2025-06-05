#creds and sharepoint site info
$creds = Get-AbsCred -credName "service-equip@abskids.net"
$freshApiKey = (Get-AbsCred -credName "Miles_Freshservice").GetNetworkCredential().Password
$encodedKey = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($freshApiKey))
$site = "https://abs79.sharepoint.com/sites/centertabletsteam"

#inventory sharepoint list names
$inventoryLists = @{
    eastInventory = "EAST-Tablet Inventory"
    mtnInventory = "Mountain- tablet inventory"
    westInventory = "West - Tablet inventory"
}

#connect to sharepoint site, exit if fails
try{
    Connect-PnPOnline -url $site -Credentials $creds -ClientId "e41d925a-fe12-4c7f-9675-87a1e5a04e7d"
}catch{
    Write-Host "Failed to connect to sharepoint in Maintenance"
    exit
}

#query for all devices with maintenance in the status
$query = @"
    <View>
        <Query>
            <Where>
                <Contains><FieldRef Name='Status'/><Value Type='Text'>Maintenance</Value></Contains>
            </Where>
        </Query>
    </View>
"@

#loop through the inventory lists
foreach($spList in $inventoryLists.Keys){

    Write-Host $inventoryLists[$spList]
    
    #get maintenance tablets, continue if error or no maintenance items
    try {
        $data = Get-PnPListItem -List $inventoryLists[$spList] -Query $query
        if(!$data){
            continue
        }
    }
    catch {
        Write-Host "Error getting SP data - $spList"
        continue
    }

    #loop through maintenance tablets in this inventory SP list
    foreach($tablet in $data){

        $tablet
        Write-Host "_-________-____-_--------__-___-------"

        #get tablet info... if info not available, set as N/A
        $tabletInfo = @{
            serialNumber = $($tablet["Title"]) ?? "N/A"
            center = $($tablet["Center"]) ?? "N/A"
            status = $($tablet["Status"]) ?? "N/A"
            email = $($tablet["CheckedOutTo"].Email) ?? "N/A"
            userName = $($tablet["CheckedOutTo"].LookupValue) ?? "N/A"
        }

        #create request body
        $body = @{
            subject = "$($tabletInfo.serialNumber) - Tablet Maintenance"
            category = "BT Tablets"
            requester_id = 21001534735 #NEED TO UPDATE THIS TO A SERVICE ACCOUNT
            priority = 1
            source = 2
            status = 2
            type = "Incident"
            description = "Serial Number: $($tabletInfo.serialNumber)<br>Status: $($tabletInfo.status)<br>Center: $($tabletInfo.center)<br>User: $($tabletInfo.userName)<br> User Email: Status: $($tabletInfo.email)"
            custom_fields = @{
                phone_number = "-"
            }
        }

        #create request header
        $header = @{
            "Authorization" = "Basic $encodedKey"
            "Content-Type" = "application/json"
        }

        #create request
        $param = @{
            Method = "Get"
            uri = "https://abskids.freshservice.com/api/v2/assets?search=`"name%3A%27$($tablet["Title"])%27`""
            ContentType = "application/json"
            Headers = $header
        }


        #call api for device info (this is to get the device id to add asset to ticket)
        #$freshDeviceInfo
        try{
            $freshDeviceInfo = Invoke-WebRequest @param -SkipHttpErrorCheck
        }catch{
            Write-Host "Error getting device from freshservice"
        }
        $freshDeviceInfo = $freshDeviceInfo.Content | ConvertFrom-Json
        $freshDeviceInfo.assets


        #if asset exists, associate asset to ticket
        if($freshDeviceInfo.StatusCode -eq 200 -and $freshDeviceInfo.assets){
            Write-Host "Not empty!" -ForegroundColor Green
            $freshDeviceInfo = $freshDeviceInfo | ConvertFrom-Json
            $freshDeviceInfo = $freshDeviceInfo.assets.display_id 
            $body.Add("assets", @(@{"display_id"=$freshDeviceInfo}))

        }

        switch ($tabletInfo.status) {
            "Maintenance - Charging issues" {$body.sub_category = "Charging Issues"}
            "Maintenance - Broken screen" {$body.sub_category = "Broken Screen"}
            "Maintenance - Login issues" {$body.sub_category = "Login Issues"}
            "Maintenance - Wireless connectivity" {$body.sub_category = "WiFi Connectivity"}
            Default {$body.sub_category = "Other"}
        }

        $body = $body | ConvertTo-Json

        $url = "https://abskids.freshservice.com/api/v2/tickets"

        #$response = Invoke-WebRequest -Method Post -Uri $url -Headers $header -Body $body -SkipHttpErrorCheck
        
        if($response.StatusCode -ne 200){
            Write-Host "Couldn't create ticket for $($tabletInfo.serialNumber) - $($response.StatusDescription)"
            continue
        }
        
        #Set-PnPListItem -List $inventoryLists[$spList] -Identity $tablet.Id -Values @{"Status" = "Ticket Created"}
    
    }
}
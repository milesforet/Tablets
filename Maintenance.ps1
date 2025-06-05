try{
    $creds = Get-AbsCred -credName "service-equip@abskids.net"
}catch{
    Write-Host "Failed to get Creds"
    exit
}


$site = "https://abs79.sharepoint.com/sites/centertabletsteam"

$creds | Out-File -Append -FilePath "test_file.txt"

$eastInventory = "EAST-Tablet Inventory"
$mtnInventory = "Mountain- tablet inventory"
$westInventory = "West - Tablet inventory"


try{
    Connect-PnPOnline -url $site -Credentials $creds -ClientId "e41d925a-fe12-4c7f-9675-87a1e5a04e7d"
}catch{
    Write-Host "Failed to connect to sharepoint in Maintenance"
    exit
}

$query = @"
    <View>
        <Query>
            <Where>
                <Contains><FieldRef Name='Status'/><Value Type='Text'>Maintenance</Value></Contains>
            </Where>
        </Query>
    </View>
"@

$maintenanceTablets = @{

}

#EST Inventory
$sp_list = Get-PnPListItem -List $eastInventory -Query $query

foreach($tablet in $sp_list){
    $maintenanceTablets[$tablet["Center"]] += @("$($tablet["Title"]) - $($tablet["Status"])")
}


#MST Inventory
$sp_list = Get-PnPListItem -List $mtnInventory -Query $query

foreach($tablet in $sp_list){
    $maintenanceTablets[$tablet["Center"]] += @("$($tablet["Title"]) - $($tablet["Status"])")
}


#PST Inventory
$sp_list = Get-PnPListItem -List $westInventory -Query $query

foreach($tablet in $sp_list){
    $maintenanceTablets[$tablet["Center"]] += @("$($tablet["Title"]) - $($tablet["Status"])")
}

$message = "<body>"

$maintenanceTablets.Keys | ForEach-Object {
    $message += "<h3 style=`"margin-bottom: 10px`">$_</h3>"
    foreach($item in $maintenanceTablets[$_]){
        $message += $item + "</br>"
    }
    $message += "</br>"
}
$message += "</body>"
$smtpCred = Get-AbsCred -credName "miles-smtp"
$recipients = "mforet@abskids.com"

Send-MailMessage -Body $message -Subject "Maintenance Tablets" -SmtpServer 'mail.smtp2go.com' -Credential $smtpCred -From 'mforet@abskids.com' -To $recipients -BodyAsHtml
[System.Collections.ArrayList]$errors= @()

try{
    try{
        Connect-MgGraph -Scopes "DeviceManagementManagedDevices.PrivilegedOperations.All", "Group.Read.All" -NoWelcome
        Write-Output "Graph auth successful...`n"
    }catch{
        throw "ERROR connecting to graph- $_"
    }

    $url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=OperatingSystem eq 'Android'"

    [System.Collections.ArrayList]$tablets = @()
    $morePages = $true
    
    while($morePages){
        $androidDevices = Invoke-MgGraphRequest -Method GET -Uri $url
        
        if(!$androidDevices."@odata.nextLink"){
            
            $morePages = $false
        }else{
            $url = $androidDevices."@odata.nextLink"
        }
    

        $tablets += $androidDevices.value | Where-Object {$_."deviceEnrollmentType" -eq "androidEnterpriseDedicatedDevice"}

    }

    #loop through devices and send reboot
    foreach($tablet in $tablets){
        try{
            Invoke-GraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($tablet.Id)/rebootNow"
            #Restart-MgDeviceManagementManagedDeviceNow -ManagedDeviceId $tablet.Id
            Write-Output "$($tablet.SerialNumber) - reboot initiated"
        }catch{
            Write-Output "Failed to reboot ${$tablet.SerialNumber} - $_"
            $errors.Add("Failed to reboot ${$tablet.SerialNumber} - $_") | Out-Null
        }
    } 
}catch{
    $errors.Add($_) | Out-Null
}

$errors

exit

try{
    #SEND EMAIL WITH ANY ERRORS
    if(!$errors){
        "Completed with no errors!!"
        exit
    }

    $smtpCreds = Get-AbsCred -credName "smtp2go"

    $errorEmailBody = ""

    foreach($err in $errors){
        $errorEmailBody += "<p>$err</p><br>"   
    }

    $emailHeaders = @{
        "Content-Type" = "application/json"
        "X-Smtp2go-Api-Key" = $smtpCreds.GetNetworkCredential().Password
        "Accept" = "application/json"
    }
    
    $emailBody = @{
        "sender" = "mforet@abskids.com"
        "to" = @("mforet@abskids.com")
        "Subject" = "Error(s) in Weekly Tablet Reboot"
        "html_body" = $errorEmailBody
    } | ConvertTo-Json

    $sendEmailRes = Invoke-WebRequest -Method "POST" -Uri "https://api.smtp2go.com/v3/email/send" -Headers $emailHeaders -Body $emailBody -SkipHttpErrorCheck
    
    if($sendEmailRes.StatusCode -ne 200){
        throw "Failed to send errror email"
    }

    Write-Output "Email with errors sent!"

}catch{
    $_
    exit
}
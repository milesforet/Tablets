$sandboxBaseUrl = "https://apis-sandbox.fedex.com"
$productionUrl = "https://apis.fedex.com"

function GetFedexToken {
    param(
        [Parameter(Mandatory)][bool] $isProductionEnvironment
    )

    if($isProductionEnvironment){
        $environment = "Production"
    }else{
        $environment = "Sandbox"
    }

    Write-Host $isProductionEnvironment -ForegroundColor Blue

    #read file with token expiration datetime & get current time
    $fileInfo = Get-Content ".\Fedex$environment`BearerExpiration.txt"
    if($fileInfo){
        $expiration = ($fileInfo).Trim()
        $expirationDateTime = [datetime]::ParseExact($expiration, "MM/dd/yy hh:mm tt", $null)
        $currentTime = Get-Date
    }

    #check if current time is less than expiration time
    if($currentTime -lt $expirationDateTime){

        Write-Host "Fedex $environment token still good!"
        #token is still good, retrieves from the XML file and returns the token
        $token = (Get-AbsCred -credName "Fedex-$environment`-BearerToken").GetNetworkCredential().Password
        return $token
    }
    
    #gets the Dell app client ID and client secret
    $creds = Get-AbsCred -credName "Fedex-$environment`-WarrantyPickUps"
    $client_id = $creds.UserName
    $client_secret = $creds.GetNetworkCredential().Password

    try{

        $header = @{
            "content-type" = "application/x-www-form-urlencoded"
        }

        $body = @{
            "grant_type" = "client_credentials"
            "client_id" = $creds.UserName
            "client_secret" = $creds.GetNetworkCredential().Password
        }
        
        if($isProductionEnvironment){
            $response = Invoke-WebRequest -Method POST -Uri "$productionUrl/oauth/token" -Headers $header -Body $body -SkipHttpErrorCheck
        }else{
            $response = Invoke-WebRequest -Method POST -Uri "$sandboxBaseUrl/oauth/token" -Headers $header -Body $body -SkipHttpErrorCheck
        }


        if($response.StatusCode -eq 200){

            #create encrypted xml file with new token and return the token
            $response = $response.Content | ConvertFrom-Json
            Create-AbsCred -credName "Fedex-$environment`-BearerToken" -username "na" -password $response.access_token
            $currentTime = Get-date ((Get-Date).AddSeconds(3600)) -UFormat "%m/%d/%y %I:%M %p"
            $currentTime | Out-File ".\Fedex$environment`BearerExpiration.txt"


            return $response.access_token
        }

    }catch{
        Write-Host "Error getting Fedex auth token"
    }
}

function CreateFedexPickUp {
     param(
        [Parameter(Mandatory)][string] $contactName,
        [Parameter(Mandatory)][string] $contactPhoneNumber,
        [Parameter(Mandatory)][array] $streetAddress,
        [Parameter(Mandatory)][string] $city,
        [Parameter(Mandatory)][string] $state,
        [Parameter(Mandatory)][string] $zipCode,
        [Parameter(Mandatory)][string] $pickUpStartTime,
        [Parameter(Mandatory)][int] $packageCount,
        [Parameter(Mandatory)][string] $url,
        [Parameter(Mandatory)][string] $accountNumber,
        [Parameter(Mandatory)][string] $bearerToken,
        [Parameter(Mandatory)][array] $confirmationEmail
    )

    $headers = @{
        "content-type" = "application/json"
        "authorization" = "Bearer $bearerToken"
    }


    $body = @{
        "associatedAccountNumber" = @{
            "value" = $accountNumber
        }
        "originDetail" = @{
            "pickupLocation" = @{
                "contact" = @{
                    "companyName" = "ABS Kids"
                    "personName" = $contactName
                    "phoneNumber" = $contactPhoneNumber
                }
                "address" = @{
                    "streetLines" = $streetAddress
                    "city" = $city
                    "stateOrProvinceCode" = $state
                    "postalCode" = $zipCode
                    "countryCode" = "US"
                }
            }
            "readyDateTimestamp" = $pickUpStartTime #<[YYYY-MM-DDTHH:MM:SSZ]>
            "customerCloseTime" = "17:00:00" #
        }

        "carrierCode" = "FDXE"
        "packageCount" = $packageCount
        "pickUpNotificationDetail" = @{
            "emailDetails" = $confirmationEmail
            "format" = "TEXT"
        }

    } | ConvertTo-Json -Depth 5

    $response = Invoke-WebRequest -Method POST -Uri "$url/pickup/v1/pickups" -Headers $headers -Body $body -SkipHttpErrorCheck


    if($response.StatusCode -eq 200){
        $response = $response.Content | ConvertFrom-Json
        return @{status = 200; pickUpCode = $response.output.pickupConfirmationCode}
    }

}



function AddressValidation {
    param(
        [Parameter(Mandatory)][array] $streetAddress,
        [Parameter(Mandatory)][string] $city,
        [Parameter(Mandatory)][string] $state,
        [Parameter(Mandatory)][string] $zipCode,
        [Parameter(Mandatory)][string] $url,
        [Parameter(Mandatory)][string] $bearerToken
    )

    $headers = @{
        "content-type" = "application/json"
        "authorization" = "Bearer $bearerToken"
    }

    $headers | ConvertTo-Json

    $body =  @{
        "addressesToValidate" = @(
            @{
                "address" = @{
                    "streetLines" = $streetAddress
                    "city" = $city
                    "stateOrProvinceCode" = $state
                    "postalCode" = $zipCode
                    "countryCode" = "US"
                }
            }
        )
    } | ConvertTo-Json -Depth 4

    Invoke-WebRequest -Method POST -Uri "$url/address/v1/addresses/resolve" -Headers $headers -Body $body -SkipHttpErrorCheck
}

[string]$baseUrl = ""
[string]$account = ""
[string]$bearer = ""


$isTest = Read-Host -Prompt "Is this Production or Test? (P/T)"
switch ($isTest) {
    "P" {
        $baseUrl = $productionUrl
        $account = "722655630"
        $bearer = GetFedexToken -isProductionEnvironment $true
    }
    "T" {
        $baseUrl = $sandboxBaseUrl
        $account = "740561073"
        $bearer = GetFedexToken -isProductionEnvironment $false
    }
    Default {Write-Host "ERROR: INVALID OPTION... PLEASE ENTER P OR T"; exit}
}



$tomorrow = (Get-Date).AddDays(1)
$pickUpStart = $tomorrow.ToString("yyyy-MM-dd") + "T09:00:00"

$params = @{
    contactName = "LuMar Bennett"
    contactPhoneNumber = "9198934929"
    streetAddress = @("8300 Health Park", "Ste 10")
    city = "Raleigh"
    state = "NC"
    zipCode = "27615"
    #$pickUpStart = ((Get-Date).AddHours(12)).ToString("yyyy-MM-ddTHH:mm:ss")
    pickUpStartTime = $pickUpStart
    packageCount = 9
    url = $baseUrl
    accountNumber = $account
    bearerToken = $bearer
    confirmationEmail = @("mforet@abskids.com")
}


#AddressValidation -streetAddress @("456 Summer Ridge Drive") -city "Stanley" -state "NC" -zipCode "28164" -url $baseUrl -bearerToken $bearer
$result = CreateFedexPickUp @params
$result

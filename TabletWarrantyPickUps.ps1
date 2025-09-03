#VARIABLE TO DENOTE IF THIS IS A TEST
$isTest = $true

<#
CACHED USERS TO SAVE ON UNECESSAY API CALLS
Hashtable with string keys and hashtable values. Fomat example:
    "123456789" : {
        email : "example@abskids.com"
        name : "example@abskids.com"
    }
#>
$cachedUsers = @{}

#FRESHSERVICE API INFO
$freshServiceBaseUrl = "https://abskids.freshservice.com"
$freshApiKey = (Get-AbsCred -credName "service-fs-inventory@abskids.com").GetNetworkCredential().Password
$freshApiKey = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($freshApiKey))

#CREATE VARIABLES FOR FEDEX AND UPS INFO
[string]$fedexToken = ""
[string]$upsToken = ""

[string]$fedexBaseUrl = ""
[string]$upsBaseUrl = ""

$fedexCreds = $null
$upsCreds = Get-AbsCred -credName "UPS_TabletWarrantyPickUps"

#ASSIGN VARIABLES DEPENING ON IF IT'S A TEST OR NOT
if($isTest){
    Write-Output "`n*********************THIS IS A TEST**********************`n"
    $fedexAccountNumber = "740561073"
    $fedexBaseUrl = "https://apis-sandbox.fedex.com"
    $fedexCreds = Get-AbsCred -credName "Fedex-Sandbox-WarrantyPickUps"
    $upsBaseUrl = "https://wwwcie.ups.com"
    
}else{
    $fedexAccountNumber = "722655630"
    $fedexBaseUrl = "https://apis.fedex.com"
    $fedexCreds = Get-AbsCred -credName "Fedex-Production-WarrantyPickUps"
    $upsBaseUrl = "https://onlinetools.ups.com"
}


#FUNCTION TO UPDATE THE TICKET TO PENDING STATUS AND MAKE A PRIVATE NOTE
function AddNoteToTicket([string]$message, [int32]$ticketNumber, $agentEmail = $null){

    #BODY/HEADER INFO FOR API CALL
    $updateHeader = @{
        "Content-Type" = "application/json"
        "Authorization" = "Basic $freshApiKey"
    }

    $body = @{
        "private" = $true
        "body" = $message
    }

    if($agentEmail){
        $body.Add("notify_emails", @($agentEmail))
    }

    $body = $body | ConvertTo-Json

    #CREATE PRIVATE NOTE ON TICKET
    $response = Invoke-WebRequest -Method Post -Uri "$freshServiceBaseUrl/api/v2/tickets/$ticketNumber/notes" -Body $body -Headers $updateHeader -SkipHttpErrorCheck

    if($response.StatusCode -ne 201){
        Write-Output "Error creating private note - $($response.StatusDescription)"
        Write-Output $response.Content
    }
    Write-Output "Private note added"
}

function SetTicketToPending([int32]$ticketNumber){
    #BODY/HEADER INFO FOR API CALL
    $updateHeader = @{
        "Content-Type" = "application/json"
        "Authorization" = "Basic $freshApiKey"
    }

    $updateBody = @{
        "status" = 3
    } | ConvertTo-Json

    #SET STATUS TO PENDING
    $response = Invoke-WebRequest -Method Put -Uri "$freshServiceBaseUrl/api/v2/tickets/$ticketNumber" -Body $updateBody -Headers $updateHeader -SkipHttpErrorCheck

    if($response.StatusCode -ne 200){
        Write-Output "Error updating status to pending - $($response.StatusDescription)"
        Write-Output $response.Content
    }
    
    Write-Output "Status updated to Pending"

}


#FUNCTION TO GET FEDEX TOKEN
function GetFedExToken {
    try{

        #HEADER/BODY INFO
        $fedexTokenHeader = @{
            "content-type" = "application/x-www-form-urlencoded"
        }

        $body = @{
            "grant_type" = "client_credentials"
            "client_id" = $fedexCreds.UserName
            "client_secret" = $fedexCreds.GetNetworkCredential().Password
        }
        
        #CALL API TO GET TOKEN
        $fedexResponse = Invoke-WebRequest -Method POST -Uri "$fedexBaseUrl/oauth/token" -Headers $fedexTokenHeader -Body $body -SkipHttpErrorCheck

        if($fedexResponse.StatusCode -ne 200){
            Write-Output "Error getting FedEx OAuth Token. Exiting"
            exit
        }

        $fedexResponse = $fedexResponse.Content | ConvertFrom-Json
        return $fedexResponse.access_token

    }catch{
        Write-Output "Error getting Fedex OAuth token. Error: $_"
        exit
    }
}

function GetUpsToken {

    try{
        $url = "$upsBaseUrl/security/v1/oauth/token"
    
        $upsCreds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $upsCreds.UserName, $upsCreds.GetNetworkCredential().password)))
        
        $body = @{
            "grant_type"= "client_credentials"
        }
        
        $headers = @{
            "Content-Type" = "application/x-www-form-urlencoded"
            "x-merchant-id" = "string"
            "Authorization" = "Basic $upsCreds"
        }

        $upsAuthResponse = Invoke-WebRequest -Uri $url -Method Post -Headers $headers -Body $body -SkipHttpErrorCheck

        if($upsAuthResponse.StatusCode -ne 200){
            Write-Output "Error getting UPS OAuth token"
            exit
        }
        $upsAuthResponse = $upsAuthResponse.Content | ConvertFrom-Json

        return $upsAuthResponse.access_token
    }catch{
        Write-Output "Error getting UPS OAuth Token. Error: $_"
        exit
    }
}


###################################END OF FUNCTIONS########################################



#HEADER FOR FRESHSERVICE API CALLS
$freshServiceHeaders = @{
    "Content-Type" = "application/json"
    "Authorization" = "Basic $freshApiKey"
}

#GET TICKETS WITH "SCHEDULE FEDEX PICK UP STATUS"
$response = Invoke-WebRequest -Uri "$freshServiceBaseUrl/api/v2/tickets/filter?workspace_id=17&query=`"status:17`"" -Headers $freshServiceHeaders -SkipHttpErrorCheck

#EXIT IF STATUS CODE ISN'T SUCCESSFUL
if($response.StatusCode -ne 200){
    Write-Output "ERROR EXITING - $($response.StatusCode) - $($response.StatusDescription)"
    exit
}

$response = $response.Content | ConvertFrom-Json

Write-Output "Tickets with FedEx Pick Up status: $($response.total)"

#EXIT IF NO TICKETS ARE FOUND WITH THIS STATUS, EXIT
if($response.total -lt 1){
    Write-Output "No tickets with status. Exiting..."
    exit
}

#GET CUSTOM OBJECT WITH ADDRESSES
$addressCustomObject = Invoke-WebRequest -Uri "$freshServiceBaseUrl/api/v2/objects/21000041357/records?page_size=100" -Headers $freshServiceHeaders -SkipHttpErrorCheck

if($addressCustomObject.StatusCode -ne 200){
    Write-Output "ERROR GETTING CUSTOM OBJECT. EXITING... $($addressCustomObject.StatusDescription)"
    Write-Output $addressCustomObject.Content
    exit
}


$addressCustomObject = $addressCustomObject.Content | ConvertFrom-Json
$addressCustomObject = $addressCustomObject.records.data

#LOOP THROUGH ARRAY OF TICKETS WITH FEDEX PICK UP STATUS
foreach($ticket in $response.tickets){
    Write-Output "`n---------#INC-$($ticket.Id) - $($ticket.subject)-----------"

    
    $agentId = $ticket.responder_id
    $requesterId = $ticket.requester_id

    #CHECK IF TICKET IS ASSIGNED TO SOMEONE
    if(-not $agentId){
        $errorMessage = "Ticket has to be assigned to an agent for Pick Up to be scheduled."
        SetTicketToPending -ticketNumber $ticket.Id
        AddNoteToTicket -message $errorMessage -ticketNumber $ticket.Id
        Write-Output "Ticket not assigned to agent. Contiunuing to next ticket."
        continue
    }

    #GET AGENT INFO AND ADD TO CACHED USERS IF IT'S NOT ALREADY IN CACHED USER HASH
    if(-not $cachedUsers.ContainsKey($agentId)){
        $agentInfoResponse = Invoke-WebRequest -Method Get -Uri "$freshServiceBaseUrl/api/v2/agents/$agentId" -Headers $freshServiceHeaders -SkipHttpErrorCheck

        if($agentInfoResponse.StatusCode -ne 200){
            Write-Output "Error getting agent info $agentId"
            break
        }

        $agentInfoResponse = $agentInfoResponse.Content | ConvertFrom-Json

        #ADD USER TO CACHED USER HASH (THIS IS TO PREVENT UNESSESARY API CALLS TO GET USER INFO)
        
        $cachedUsers.Add($agentId, @{
            "email" = $agentInfoResponse.agent.email
            "name" = $agentInfoResponse.agent.first_name + " " + $agentInfoResponse.agent.last_name
        })
    }
    
    #CHECK IF FIELDS ARE BLANK
    [System.Collections.ArrayList]$missingFields = @()
    if(-not $ticket.custom_fields.phone_number_for_pick_up){$missingFields += "Phone Number"} #CHECK PHONE NUMBER FIELD iS NOT BLANK
    if(-not $ticket.custom_fields.fedex_pick_up_date){$missingFields += "Pick Up Date"} #CHECK PICK UP DATE FIELD IS NOT BLANK
    if(-not $ticket.custom_fields.lf_center_for_fedex_pick_up){$missingFields += "Pick Up Address"} #CHECK PICK UP ADDRESS IS NOT BLANK
    if(-not $ticket.custom_fields.cf_of_boxes_for_ups -and -not $ticket.custom_fields.cf_of_boxes_for_pick_up){$missingFields += "# of Boxes for FedEx and/or # of Boxes for UPS"}

    #IF THERE ARE MISSING FIELDS OUTPUT THEM AND UPDATE TICKET WITH THE MISSING FIELD INFO.
    if($missingFields.Length -gt 0){
        $errorMessage = "The following fields cannot be blank:<ul>"
        
        foreach($field in $missingFields){
            $errorMessage += "<li>$field</li>"
        }

        $errorMessage += "</ul>"

        SetTicketToPending -ticketNumber $ticket.Id
        AddNoteToTicket -message $errorMessage -ticketNumber $ticket.Id -agentEmail $cachedUsers[$agentId].email
        Write-Output "Missing fields: $($missingFields -join ", ")... Continuing to next ticket"
        continue
    }

    #CHECK IF # OF FEDEX BOXES IS <= 0 OR > 99#CHECK IF FEDEX BOX # IS NOT A VALUE 1-99. AND IGNORE IF THE VALUE IS NULL BC WE MAY NOT NEED A FEDEX PICK UP
    if(($ticket.custom_fields.cf_of_boxes_for_pick_up -lt 1 -or $ticket.custom_fields.cf_of_boxes_for_pick_up -gt 99) -and $null -ne $ticket.custom_fields.cf_of_boxes_for_pick_up){
        $errorMessage = "Invalid number of FedEx boxes. Number has to be 1-99. Number selected: $($ticket.custom_fields.cf_of_boxes_for_pick_up)"
        SetTicketToPending -ticketNumber $ticket.Id
        AddNoteToTicket -message $errorMessage -ticketNumber $ticket.Id -agentEmail $cachedUsers[$agentId].email
        Write-Output "Invalid number of FedEx boxes selected. # selected: $($ticket.custom_fields.cf_of_boxes_for_pick_up)"
        continue
    }

    #CHECK IF # OF UPS BOXES IS <= 0 OR > 99#CHECK IF FEDEX BOX # IS NOT A VALUE 1-99
    if(($ticket.custom_fields.cf_of_boxes_for_ups -lt 1 -or $ticket.custom_fields.cf_of_boxes_for_ups -gt 99) -and $null -ne $ticket.custom_fields.cf_of_boxes_for_ups){
        $errorMessage = "Invalid number of UPS boxes. Number has to be 1-99. Number selected: $($ticket.custom_fields.cf_of_boxes_for_ups)"
        SetTicketToPending -ticketNumber $ticket.Id
        AddNoteToTicket -message $errorMessage -ticketNumber $ticket.Id -agentEmail $cachedUsers[$agentId].email
        Write-Output "Invalid number of UP boxes selected. # selected: $($ticket.custom_fields.cf_of_boxes_for_ups)"
        continue
    }



    #WE ARE HERE CHECKING TO MAKE SURE THE 2 SEPARATE FUNCTIONS ARE WORKING CORRECTLY. WE ARE ALSO ADDNIG WRITE OUTPUTS FOR ALL 
    #THE IF STATEMENTS BECAUSE THE FUNCTIONS NO LONGER OUTPUT THE ERROR MESSAGE THAT IS ADDED TO THE TICKET BECAUSE THE HTML TAGS MADE THE ERROR MESSAGES MESSY AND HARD TO READ

    #CREATE DATETIME OBJECT FOR PICK UP DATE FIELD
    $pickUpDate = $null
    try{
        $pickUpDate = [datetime]$ticket.custom_fields.fedex_pick_up_date
    }catch{
        $errorMessage = "Issue with datetime object."
        SetTicketToPending -ticketNumber $ticket.Id
        AddNoteToTicket -message $errorMessage -ticketNumber $ticket.Id -agentEmail $cachedUsers[$agentId].email
        Write-Output "Error with converting pick up date with datetime object $_"
        continue
    }

    #CHECK IF PICK UP DATE IS IN THE PAST OR A WEEKEND
    if($pickUpDate -lt (Get-Date).Date){
        $errorMessage = "Pick up date is in the past. Selected Date: $($pickUpDate.ToString("MM/dd/yyyy")) Today: $(Get-Date -Format "MM/dd/yyyy")"
        SetTicketToPending -ticketNumber $ticket.Id
        AddNoteToTicket -message $errorMessage -ticketNumber $ticket.Id -agentEmail $cachedUsers[$agentId].email
        Write-Output $errorMessage
        continue
    }elseif($pickUpDate.DayOfWeek -eq "Saturday" -or $pickUpDate.DayOfWeek -eq "Sunday") {
        $errorMessage = "Pick up date cannot be a Saturday or Sunday. $($pickUpDate.ToString("MM/dd/yyyy")) is a $($pickUpDate.DayOfWeek)"
        SetTicketToPending -ticketNumber $ticket.Id
        AddNoteToTicket -message $errorMessage -ticketNumber $ticket.Id -agentEmail $cachedUsers[$agentId].email
        Write-Output $errorMessage
        continue
    }elseif($pickUpDate -ne (Get-Date).Date -and $pickUpDate -ne (Get-Date).AddDays(1).Date){
        Write-Output "Pick up is not today or tomorrow. We will schedule later. Scheduled date: $($pickUpDate.ToString("MM/dd/yyyy"))"
        continue
    }

    #CHECK REQUESTER ID
    if( -not $requesterId){
        $errorMessage = "Requester ID not found in ticket"
        AddNoteToTicket -message $errorMessage -ticketNumber $ticket.Id -agentEmail $cachedUsers[$agentId].email
        continue
    }

    #IF REQUESTER ID IS NOT IN CACHED USERS, CALL API TO GET USER INFO
    if(-not $cachedUsers.ContainsKey($requesterId)){
        $requesterInfo = Invoke-WebRequest -Method Get -Uri "$freshServiceBaseUrl/api/v2/requesters/$requesterId" -Headers $freshServiceHeaders -SkipHttpErrorCheck

        if($requesterInfo.StatusCode -ne 200){
            $errorMessage = "Error getting requester info"
            AddNoteToTicket -message $errorMessage -ticketNumber $ticket.Id -agentEmail $cachedUsers[$agentId].email
            continue
        }
        $requesterInfo = $requesterInfo.Content | ConvertFrom-Json

        $cachedUsers.Add($requesterId, @{
            "email" = $requesterInfo.requester.primary_email
            "name" = $requesterInfo.requester.first_name+" "+$requesterInfo.requester.last_name
        })
    }
    
    #ADDRESS INFO OF THE SELECTED CUSTOM OBJECT
    $address = $addressCustomObject | Where-Object {$_.bo_display_id -eq $ticket.custom_fields.lf_center_for_fedex_pick_up}
    
    [System.Collections.ArrayList]$streetAddress = @($address.addressline1)

    #SINCE STREET ADDRESS IS AN ARRAY, WE WANT TO CHECK IF THERE IS A SUITE/ADDRESS LINE 2
    if($address.addressline2){
        $streetAddress += $address.addressline2
    }

    #DISPLAY INFO ABOUT THE REQUEST FOR LOGGING/DEBUGGING
    Write-Output "Contact Name: $($cachedUsers[$requesterId].name)`nContact Phone: $($ticket.custom_fields.phone_number_for_pick_up)`nAddress: $streetAddress, $($address.city), $($address.state) $($address.zipcode)`nAgent: $($cachedUsers[$agentId].name)`nPick Up: $($pickUpDate.ToString("MM/dd/yyyy"))`nFedEx #: $($ticket.custom_fields.cf_of_boxes_for_pick_up)`nUPS #: $($ticket.custom_fields.cf_of_boxes_for_ups)`n"

    $pickUpStatusMessage = ""

    #IF FEDEX PICK UP IS REQUIRED
    if($ticket.custom_fields.cf_of_boxes_for_pick_up){

        #IF A FEDEX TOKEN HAS NOT BEEN RETRIEVED, WE GET IT AND ASSIGN TO fedexToken variable
        if(-not $fedexToken){
            $fedexToken = GetFedExToken
        }

        $fedexHeader = @{
        "content-type" = "application/json"
        "authorization" = "Bearer $fedexToken"
        }

        $fedexPickUpBody = @{
            "associatedAccountNumber" = @{
                "value" = $fedexAccountNumber
            }
            "originDetail" = @{
                "pickupLocation" = @{
                    "contact" = @{
                        "companyName" = "ABS Kids"
                        "personName" = $cachedUsers[$requesterId].name
                        "phoneNumber" = $ticket.custom_fields.phone_number_for_pick_up
                    }
                    "address" = @{
                        "streetLines" = $streetAddress
                        "city" = $address.city
                        "stateOrProvinceCode" = $address.state
                        "postalCode" = $address.zipcode
                        "countryCode" = "US"
                    }
                }
                "readyDateTimestamp" = ($pickUpDate.ToString("yyyy-MM-dd") + "T10:00:00")
                "customerCloseTime" = "17:00:00" 
            }

            "carrierCode" = "FDXE"
            "packageCount" = $ticket.custom_fields.cf_of_boxes_for_pick_up
            "pickUpNotificationDetail" = @{
                "emailDetails" = @(
                    @{
                        "address" = $cachedUsers[$agentId].email
                        "locale" = "en_US"
                    },
                    @{
                        "address" = $cachedUsers[$requesterId].email
                        "locale" = "en_US"
                    }
                )
                "format" = "HTML"
            }

        } | ConvertTo-Json -Depth 5

        $pickUpResponse = Invoke-WebRequest -Method POST -Uri "$fedexBaseUrl/pickup/v1/pickups" -Headers $fedexHeader -Body $fedexPickUpBody -SkipHttpErrorCheck

        if($pickUpResponse.StatusCode -ne 200){
            $pickUpStatusMessage += "Error scheduling FedEx pick up. $($pickUpResponse.StatusCode) - $($pickUpResponse.StatusDescription)"
            Write-Output "Error scheduling FedEx pick up.`n$($pickUpResponse.StatusCode) - $($pickUpResponse.StatusDescription)`nBody:`n$fedexPickUpBody`nStatus:`n$($pickUpResponse.Content)"
        }else{
            $pickUpStatusMessage += "FedEx pick up scheduled!<br><br>Contact Name: $($cachedUsers[$requesterId].name)<br>Contact Phone: $($ticket.custom_fields.phone_number_for_pick_up)<br>Address: $streetAddress, $($address.city), $($address.state) $($address.zipcode)<br>Pick Up: $($pickUpDate.ToString("MM/dd/yyyy")) 10AM-5PM<br># of FedEx boxes: $($ticket.custom_fields.cf_of_boxes_for_pick_up)"
            Write-Output "FedEx pick up scheduled`n"
        }
    }

    if($pickUpStatusMessage){
        $pickUpStatusMessage += "<br><br>-------------------------------------------------------------<br><br>"
    }


    #IF UPS PICK UP IS REQUIRED
    if($ticket.custom_fields.cf_of_boxes_for_ups){
        if(-not $upsToken){
            $upsToken = GetUpsToken
        }
        
        $upsPickUpHeader = @{
            "Authorization" = "Bearer $upsToken"
            "Content-Type" = "application/json"
        }

        $upsPickUpBody = @{
            "PickupCreationRequest" = @{
                "RatePickupIndicator" = "N"
                "PickupDateInfo" = @{
                    "CloseTime" = "1700"
                    "ReadyTime" = "1000"
                    "PickupDate" = $pickUpDate.ToString("yyyyMMdd") 
                }
                "PickupAddress" = @{
                    "CompanyName" = "ABS Kids"
                    "ContactName" = $cachedUsers[$requesterId].name
                    "AddressLine" = ($streetAddress -join " ")
                    "City" = $address.city
                    "PostalCode" = $address.zipcode
                    "StateProvince" = $address.state
                    "CountryCode" = "US"
                    "ResidentialIndicator" = "N"
                    "Phone" = @{
                        "Number" = $ticket.custom_fields.phone_number_for_pick_up
                    }
                }
                "Shipper" = @{
                    "Account" =  @{
                        "AccountNumber" = "3124Y8"
                        "AccountCountryCode" = "US"
                    }

                }
                "AlternateAddressIndicator" = "Y"
                "PickupPiece" = @(
                    @{
                        "ServiceCode" = "002"
                        "Quantity" = ($ticket.custom_fields.cf_of_boxes_for_ups).ToString()
                        "DestinationCountryCode" = "US"
                        "ContainerCode" = "01"
                    }
                )
                "PaymentMethod" = "01"
                "Confirmation" = @{
                    "ConfirmationEmailAddress" = $cachedUsers[$requesterId].email
                }
                "ReferenceNumber" = "TabletPickUp-$($ticket.Id)"
            }
        } | ConvertTo-Json -Depth 4

        $upsPickUpResponse = Invoke-WebRequest -Method Post -Uri "$upsBaseUrl/api/pickupcreation/v2409/pickup" -Headers $upsPickUpHeader -Body $upsPickUpBody -SkipHttpErrorCheck
        
        if($upsPickUpResponse.StatusCode -ne 200){
            $pickUpStatusMessage += "Error scheduling UPS pick up. $($upsPickUpResponse.StatusCode) - $($upsPickUpResponse.StatusDescription)"
            Write-Output "Error scheduling UPS pick up. `n$($upsPickUpResponse.StatusCode) - $($upsPickUpResponse.StatusDescription)`nBody:`n$upsPickUpBody`nResponse:`n$($upsPickUpResponse.Content)`n"
        }else{
            $upsPickUpResponse = $upsPickUpResponse.Content | ConvertFrom-Json
            $pickUpStatusMessage += "UPS pick up scheduled!<br><br>Contact Name: $($cachedUsers[$requesterId].name)<br>Contact Phone: $($ticket.custom_fields.phone_number_for_pick_up)<br>Address: $streetAddress, $($address.city), $($address.state) $($address.zipcode)<br>Pick Up: $($pickUpDate.ToString("MM/dd/yyyy")) 10AM-5PM<br># of UPS boxes: $($ticket.custom_fields.cf_of_boxes_for_ups)<br><br>Pickup Reference #: $($upsPickUpResponse.PickupCreationResponse.PRN)"
            Write-Output "UPS pick up scheduled. $($upsPickUpResponse.PickupCreationResponse.PRN)`n"
        }
    }

    SetTicketToPending -ticketNumber $ticket.Id
    AddNoteToTicket -ticketNumber $ticket.Id -message $pickUpStatusMessage -agentEmail $cachedUsers[$agentId].email   
    Write-Output "`nDone with $($ticket.Id)!"
}
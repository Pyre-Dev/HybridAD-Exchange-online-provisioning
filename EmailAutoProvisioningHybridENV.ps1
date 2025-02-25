####################################################################################################################################################################################
# _______ _     _                     _                                          _ _   _               _              _____           _                     
# |__   __| |   (_)                   | |                                        (_) | | |             | |            / ____|         | |                    
#    | |  | |__  _ ___    ___ ___   __| | ___  __      ____ _ ___  __      ___ __ _| |_| |_ ___ _ __   | |__  _   _  | |  __ _ __ __ _| |__   __ _ _ __ ___  
#    | |  | '_ \| / __|  / __/ _ \ / _` |/ _ \ \ \ /\ / / _` / __| \ \ /\ / / '__| | __| __/ _ \ '_ \  | '_ \| | | | | | |_ | '__/ _` | '_ \ / _` | '_ ` _ \ 
#    | |  | | | | \__ \ | (_| (_) | (_| |  __/  \ V  V / (_| \__ \  \ V  V /| |  | | |_| ||  __/ | | | | |_) | |_| | | |__| | | | (_| | | | | (_| | | | | | |
#    |_|  |_| |_|_|___/  \___\___/ \__,_|\___|   \_/\_/ \__,_|___/   \_/\_/ |_|  |_|\__|\__\___|_| |_| |_.__/ \__, |  \_____|_|  \__,_|_| |_|\__,_|_| |_| |_|
#                                                                                                              __/ |                                         
#                                                                                                             |___/                                          
#####################################################################################################################################################################################
# Import the necessary modules
Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement
Import-Module Microsoft.Graph
# Ensure you have the Microsoft.Graph module installed
#run  "Install-Module Microsoft.Graph -Scope CurrentUser" in your console.

# Connect to Exchange Online (This will change into a token)
Connect-ExchangeOnline -UserPrincipalName example@domain.com

# Connect to Microsoft Graph
Connect-MgGraph

#Target User (This will change)
$UID= Read-Host "Enter the UID of the user you're actively provisioning"
$user = Get-ADuser -identity $UID -Property SamAccountName, GivenName, Surname, TelephoneNumber
$domain = "domain.com"

# Construct the email address
$email = "$($user.GivenName).$($user.Surname)@$domain"

# Check if mailbox exists
if (-not (Get-Mailbox -Identity $email -ErrorAction SilentlyContinue)) {
    # Create the mailbox.
    try{
        Enable-RemoteMailbox $UID -RemoteRoutingAddress "$($user.GivenName).$($user.Surname)@domain.mail.onmicrosoft.com"
    }
    catch{
        Write-Host "Mailbox creation FAILED for $email"
        Exit 1
    }
    try{
    Enable-Mailbox -Identity $email -Archive
    }
    catch{
        Write-Host "Mailbox Archiving for $email FAILED"
        Exit 1
    }

    # Assign the E3 license (here is where we're using McGraph to interact with our licenses)
    try{
        $E3= Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'SPE_E3'
    }
    catch{
        Write-Host "Unable to gather E3 SKU's"
        Exit 1
    }

    try{
        $disabledPlans = $E3.ServicePlans | Where-Object ServicePlanName -in ("KAIZALA_O365_P3", "YAMMER_ENTERPRISE", "MCOSTANDARD") | Select-object -ExpandProperty ServicePlanId
    }
    catch{
        Write-Host "Unable to define unwanted service SKU's. You should check if your service-names are correct"
        Exit 1
    }
    try{
    # Graph likes using unique UserID's so we're pulling that from our tenant filtering by email.
    $MGuser = Get-MgUser -Filter "userPrincipalName eq $email"
    # the logic looks funky but if you remember E3 = all service sku's in our E3. So we're asking for all E3 service sku's and $disabledPlans is other sku's we don't want.
    $addLicenses = @(
        @{SkuId = $E3.SkuId
        DisabledPlans = $disabledPlans
        }
        )  
    Set-MgUserLicense -UserId $MGuser.Id -AddLicenses $addLicenses -RemoveLicenses @()
    }
    catch{
        Write-Host "something broke in the user filter or your sku assignment, more than likely the sku assignment"
        Exit 1
    }
    Write-Host "License assigned to $email"

} else {
    Write-Host "Mailbox already exists for $email"
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

# Disconnect from Microsoft Graph
Disconnect-MgGraph

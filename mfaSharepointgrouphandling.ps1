# Import the Active Directory module

Import-Module ActiveDirectory

 

# Import the SharePoint connect module

Import-Module PnP.PowerShell

 

function Get-DateOnly($inputDate) {

    if ($inputDate) {

        return [datetime]::Parse($inputDate).ToString("MM/dd/yyyy")

    }

    return $null

}

 

Function global:Connect-MicrosoftGraph() {

    $cert = Get-AutomationCertificate -Name 'YourAutomationAppCert'

    $AppID = Get-AutomationVariable -Name 'AutomationApplication-AppId'

    Import-Module Microsoft.Graph.Identity.SignIns

    Connect-MgGraph -ClientId $AppID -TenantId "yourtenantid.onmicrosoft.com" -Certificate $Cert

}

 

# Array of exempt users

$exemptUsers = Get-ADGroupMember -Identity "Mfagroup" | Where-Object { $_.objectClass -eq 'user' } | Get-ADUser | Select-Object -ExpandProperty UserPrincipalName

 

$allowedMfaGroups = "group-exempt"


 

# Fetch current date

$today = Get-DateOnly(Get-Date)

 

$SPsite = 'https://yourorg.sharepoint.com/sites/your/site'

$SPList = 'exempt-List'

 

# Define the CAML query

$CAMLQuery = @"

<View>

    <ViewFields>

        <FieldRef Name='UserPrincipalName' />

        <FieldRef Name='RemovedDate' />

        <FieldRef Name='ID' />

        <FieldRef Name='EndDate' />

        <FieldRef Name='StartDate' />

        <FieldRef Name='Status' />

        <FieldRef Name='MFAgroup' />

    </ViewFields>

</View>

"@

# Fetch all users from the SharePoint list

$allUsers = ./Get-SPListItems.ps1 -Url $SPsite -SPList $SPList -CAMLQuery $CAMLQuery

 

# Iterate through users

foreach ($user in $allUsers) {

    $removeddate = $user.RemovedDate

    $itemId = $user.ID

    $endDate = $user.EndDate

    $startDate = $user.StartDate

    $currentStatus = $user.Status

    $UPN = $user.UserPrincipalName

 

    # Get the user's MFA groups

    $MfaGroups = $allowedMfaGroups

    $statusMessages = @()

    $statusOutputMessages = @()

    $statusColumnValue = @()

 

    if (-not [string]::IsNullOrEmpty($removedDate)) {

        continue

    }

        # Parse the EndDate value

        if (![string]::IsNullOrEmpty($endDate)) {

            try {

                $parsedEndDate = [datetime]::ParseExact($endDate, "MM/dd/yyyy HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)

            }

            catch {

                $statusMessages += "Failed to parse EndDate"

            }

        }

 

        # Parse the StartDate value

        if (![string]::IsNullOrEmpty($startDate)) {

            try {

                $parsedStartDate = [datetime]::ParseExact($startDate, "MM/dd/yyyy HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)

            }

            catch {

                $statusMessages += "Failed to parse StartDate"

            }

        }

        

 

        # Check if EndDate is earlier than StartDate

        if (![string]::IsNullOrEmpty($endDate) -and ![string]::IsNullOrEmpty($startDate) -and $parsedEndDate -lt $parsedStartDate) {

            $statusMessages += "Failed EndDate is earlier than StartDate"

        }

 

        # Check if the user exists in Active Directory

        $adUser = Get-ADUser -Filter { UserPrincipalName -eq $UPN } -ErrorAction SilentlyContinue

 

        if (!$UPN) {

            $statusMessages += "Failed User $UPN does not exist in Active Directory"

        }

 

        # Check if the user is exempt

        if ($exemptUsers -contains $UPN) {

            $statusMessages += "Failed User $UPN is exempt from removal"

        }

 

        # add user to the group in ad

        $aadUser = $adUser.Name

 

        # Remove the user from each allowed MFA group in Azure AD

        foreach ($MfaGroupName in $MfaGroups) {

            # Check if the MFA group is allowed

 

            $matchingGroup = $MfaGroups | Where-Object { $_.ToLower() -eq $MfaGroupName.ToLower() }

 

            if ([string]::IsNullOrEmpty($matchingGroup)) {

                $statusMessages += "Failed Group $MfaGroups is not the allowed mfa group: $allowedMfaGroups"

                continue

            }

 

            if ($parsedStartDate.Date -eq $today -and $statusMessages.Count -eq 0 -and $currentStatus -notlike 'Successfully added*') {

                # add user to the group in ad

                $addedUsers = ./add-userstoADGroup.ps1 -usersList $aadUser -Group $matchingGroup

 

                #get all users in group

                $successfuladdedUsers = $addedUsers.successfulUsers

                if ($successfuladdedUsers -contains $aadUser) {

                    $statusOutputMessages += "Successfully added $UPN to $matchingGroup"

                } elseif ($successfuladdedUsers -notcontains $aadUser) {

                    $statusMessages += "Failed to add $UPN to $matchingGroup : $_"

                }

            } elseif ($parsedEndDate.Date -eq $today -and $statusMessages.Count -eq 0 -and $currentStatus -notlike 'Successfully removed*') {

                # Remove user from the group in AD

                $removedUsers = ./Nonsqlview-Removeadusers.ps1 -usersList $aadUser -Group $matchingGroup

 

                #get all users removed in group

                $successfulremovedUsers = $removedUsers.successfulUsers

                

                if ($successfulremovedUsers -contains $aadUser) {

                    $statusOutputMessages += "Successfully removed $UPN from $matchingGroup"

                    # Update via user.id

                    $SPListItemValues = @{

                        "RemovedDate" = $today

                    }

                    ./Set-AllSPlist-nolog.ps1 -Url $SPsite -SPList $SPList -SPListItemValues $SPListItemValues -SPListItemID $itemId

 

                } elseif ($successfulremovedUsers -notcontains $aadUser) {

                    $statusMessages += "Failed removing $UPN from $matchingGroup : $_"

                }

            }

        }

                # Display all output messages in the host output and updates status with errors

                if ($statusMessages.Count -gt 0) {

                    $statusMessages | ForEach-Object {

                        Write-Output "Output for Item ID $itemId : $_"

                        # Update SharePoint list with the status messages

                        $statusColumnValue = $statusMessages -join " | "

                        $SPListItemValues = @{

                        "Status" = $statusColumnValue

                        }

                        ./Set-SPlist-nl.ps1 -Url $SPsite -SPList $SPList -SPListItemValues $SPListItemValues -SPListItemID $itemId

                    }

                }

                

                # Display all output messages in the host output and updates status message.

                if ($statusMessages.Count -eq 0 -and $statusOutputMessages -ne $null) {

                    # Update SharePoint list with the status output messages

                    $statusColumnValue = $statusOutputMessages -join " | "

                     $SPListItemValues = @{

                        "Status" = $statusColumnValue

                    }

                    ./Set-SPlist-nl.ps1 -Url $SPsite -SPList $SPList -SPListItemValues $SPListItemValues -SPListItemID $itemId

                }

    }

Write-Output "complete"

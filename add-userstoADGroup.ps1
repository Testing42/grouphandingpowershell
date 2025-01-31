[CmdletBinding()]

param (

    [string[]] $usersList,

    [string] $Group

)

 

# Required modules

Import-Module ActiveDirectory

 

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

 

function Add-AllUsersToGroup {

    [CmdletBinding()]

    param (

        [string[]] $usersList,

        [string] $Group

    )

 

    $userSIDs = @()  # Array to hold user Security IDs (SIDs)

    $successfulUsers = @()  # Array to hold successfully added users

    $failedUsers = @()  # Array to hold users that failed to be added

 

    foreach ($User in $usersList) {

        try {

            # Get user details from AD

            $AdUser = Get-ADUser -Filter "Name -eq '$User'" -Properties SamAccountName

            if ($AdUser -and $AdUser.SamAccountName) {

                # Collect the user's SID if not null

                $userSIDs += $AdUser.SID

                $successfulUsers += $User

            } else {

                Write-Output "No valid AD user found for $User"

                $failedUsers += $User

            }

        } catch {

            # Handle errors

            Write-Output "Error dealing with user $($User): $_"

            $failedUsers += $User

        }

    }

 

    # Check if there are any valid users to add to the group

    if ($userSIDs.Count -gt 0) {

        try {

            # Add all collected users to the AD group in one go

            Add-ADGroupMember -Identity $Group -Members $userSIDs -ErrorAction Stop

            Write-Output "All users added to the group successfully."

        } catch {

            Write-Output "Error adding users to group: $_"

            $failedUsers += $successfulUsers  # Assume all failed if the batch add fails

            $successfulUsers = @()

        }

    } else {

        Write-Output "No valid users available to add to the group."

    }

 

    # Log the outcome

    if ($successfulUsers.Count -gt 0) {

        Write-Output "Successfully added users: $($successfulUsers -join ', ')"

    }

    if ($failedUsers.Count -gt 0) {

        Write-Output "Users with issues: $($failedUsers -join ', ')"

    }

    #return $successfulUsers

    return @{

        successfulUsers = $successfulUsers

        failedUsers = $failedUsers

    }

}




# Get the current members of the AD group

try {

    # Get the group members from AD

    $groupNames = Get-ADGroup -Identity $Group

 

    # Get the distinguished name of the group

    $groupDN = (Get-ADGroup -Identity $groupNames).DistinguishedName

 

    # Use LDAP filter to find users in the group

    $ldapFilter = "(memberOf=$groupDN)"

 

    # Build the LDAP filter to search for direct group members

    $groupMembers = Get-ADUser -LDAPFilter $ldapFilter -Properties Name

 

    $nameOfGroups = $groupMembers.name

 

} catch {

    # Handle the case when the AD group is empty or not found

    Write-Output "Error retrieving AD group members."

    $nameOfGroups = @()  # Initialize as empty array if error occurs

}

 

# Compare SQL users with AD group members and add discrepancies

if ($nameOfGroups.Count -gt 0) {

    $Differences = Compare-Object -ReferenceObject $nameOfGroups -DifferenceObject $usersList

 

    # Filter users that need to be added to the AD group

    if ($Differences) {

        $UsersToAdd = $Differences | Where-Object { $_.SideIndicator -eq "=>" } | Select-Object -ExpandProperty InputObject

 

        # Add missing users to the AD group

        Add-AllUsersToGroup -usersList $UsersToAdd -Group $Group

 

    } else {

        Write-Output "No new users need to be added."

    }

} else {

    # If AD group is initially empty, add all SQL users

    Add-AllUsersToGroup -usersList $usersList -Group $Group

}

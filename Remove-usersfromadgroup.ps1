[CmdletBinding()]

param (

    [string[]] $usersList,

    [string] $Group

)

 

# Required modules

Import-Module ActiveDirectory

 

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

 

function Remove-AllUsersFromGroup {

    [CmdletBinding()]

    param (

        [string[]] $usersList,

        [string] $Group

    )

 

    $userSIDs = @()  # Array to hold user Security IDs (SIDs)

    $successfulUsers = @()  # Array to hold successfully removed users

    $failedUsers = @()  # Array to hold users that failed to be removed

 

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

 

    # Check if there are any valid users to remove from the group

    if ($userSIDs.Count -gt 0) {

        try {

            # Remove all collected users from the AD group in one go

            Remove-ADGroupMember -Identity $Group -Members $userSIDs -Confirm:$false -ErrorAction Stop

            Write-Output "All users removed from the group successfully."

        } catch {

            Write-Output "Error removing users from group: $_"

            $failedUsers += $successfulUsers  # Assume all failed if the batch remove fails

            $successfulUsers = @()

        }

    } else {

        Write-Output "No valid users available to remove from the group."

    }

 

    # Log the outcome

    if ($successfulUsers.Count -gt 0) {

        Write-Output "Successfully removed users: $($successfulUsers -join ', ')"

    }

    if ($failedUsers.Count -gt 0) {

        Write-Output "Users with issues: $($failedUsers -join ', ')"

    }

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

 

# Compare SQL users with AD group members and remove discrepancies

if ($nameOfGroups.Count -gt 0) {

    $Differences = Compare-Object -ReferenceObject $usersList -DifferenceObject $nameOfGroups

 

    # Filter users that need to be removed from the AD group

    if ($Differences) {

        $UsersToRemove = $Differences | Where-Object { $_.SideIndicator -eq "=>" } | Select-Object -ExpandProperty InputObject

 

        # Remove users from the AD group

        Remove-AllUsersFromGroup -usersList $UsersToRemove -Group $Group

 

    } else {

        Write-Output "No users need to be removed."

    }

} else {

    Write-Output "The AD group is empty or not found."

}

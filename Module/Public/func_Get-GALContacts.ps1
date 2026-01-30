function Get-GALContacts {
    [CmdletBinding()]
    param (
        [bool]$ContactsWithoutPhoneNumber,
        [bool]$ContactsWithoutEmail
    )
    try {
        Write-VerboseEvent "Getting GAL contacts"
        $selectProps = @(
            "displayName",
            "givenName",
            "surname",
            "jobTitle",
            "department",
            "mail",
            "businessPhones",
            "mobilePhone",
            "onPremisesExtensionAttributes"
        ) -join ","
        $allContacts = if ($UseGraphSDK) {
            Get-MgUser -All -Property $selectProps
        } else {
            New-GraphRequest -Endpoint "/users?`$select=$selectProps&`$top=999" -Beta
        }
        if (-not $ContactsWithoutPhoneNumber) {
            $allContacts = $allContacts | Where-Object { $_.businessPhones -or $_.mobilePhone }
        }
        if (-not $ContactsWithoutEmail) {
            $allContacts = $allContacts | Where-Object { $_.mail }
        }
        $returnObject = @()
        $allContacts | ForEach-Object {
            $extensionAttribute1 = $null
            if ($_.onPremisesExtensionAttributes) {
                $extensionAttribute1 = $_.onPremisesExtensionAttributes.extensionAttribute1
            }
            elseif ($_.extensionAttribute1) {
                $extensionAttribute1 = $_.extensionAttribute1
            }
            $extensionAttribute1Trimmed = if ([string]::IsNullOrWhiteSpace($extensionAttribute1)) { $null } else { $extensionAttribute1.Trim() }
            $surnameWithChildren = $_.surname
            $displayNameWithChildren = $_.displayName
            if ($extensionAttribute1Trimmed) {
                $suffix = " ($extensionAttribute1Trimmed)"
                $surnameWithChildren = if ([string]::IsNullOrWhiteSpace($_.surname)) { $suffix.Trim() } else { "$($_.surname)$suffix" }
                $displayNameWithChildren = if ([string]::IsNullOrWhiteSpace($_.displayName)) { $suffix.Trim() } else { "$($_.displayName)$suffix" }
            }
            $returnObject += [pscustomobject]@{
                businessPhones = $_.businessPhones
                displayname    = $displayNameWithChildren
                givenName      = $_.givenName
                surname        = $surnameWithChildren
                jobTitle       = $_.jobTitle                
                department     = $_.department
                personalNotes  = $extensionAttribute1Trimmed
                # homePhones     = if (-not $_.homePhones) { @() } else { @($_.homePhones) }
                emailAddresses = @(@{
                        name    = $_.mail
                        address = $_.mail
                    })
            }
        }
        Write-VerboseEvent "$($returnObject.count) contacts found"
        return $returnObject
    }
    catch {
        throw (Format-ErrorCode $_).ErrorMessage
    }
}

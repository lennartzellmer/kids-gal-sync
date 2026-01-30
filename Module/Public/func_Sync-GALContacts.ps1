function Sync-GALContacts {
    [cmdletbinding()]
    param (
        [parameter(Mandatory)][string]$Mailbox,
        [parameter(Mandatory)][string]$ContactFolderName,
        [parameter(Mandatory)][array]$ContactList
    )
    Write-LogEvent -Level Info -Message "Beginning sync for $($Mailbox)" 

    # Get/create contact folder
    try {
        $contactFolder = Get-ContactFolder -Mailbox $Mailbox -ContactFolderName $ContactFolderName
        if ($contactFolder) { Write-LogEvent -Level Info -Message "Found folder $($ContactFolderName) for $($Mailbox)" }
        else {
            try {
                $contactFolder = New-ContactFolder -Mailbox $Mailbox -ContactFolderName $ContactFolderName 
                Write-LogEvent -Level Info -Message "Created folder $($ContactFolderName) for $($Mailbox)"
            }
            catch { throw "Something went wrong creating the contact folder" }
        }
        if (-not $contactFolder) { throw "No contact folder found or not able to create one" }
    }
    catch {
        throw "Something went wrong getting the contact folder for $($Mailbox)"
    }

    # get contacts in that folder
    try {
        $contactsInFolder = Get-FolderContact -ContactFolder $contactFolder
    }
    catch {
        throw "Failed to create contact folder $($ContactFolderName) in mailbox $($Mailbox)"
    }

    function Get-ContactPrimaryEmail {
        param (
            [parameter(Mandatory)][object]$Contact
        )
        if ($Contact.emailAddresses) {
            $address = $Contact.emailAddresses |
                ForEach-Object { $_.address } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -First 1
            if ($address) { return $address.ToLowerInvariant() }
        }
        return $null
    }

    function Set-ContactNotesFromExtensionAttribute {
        param (
            [parameter(Mandatory)][object]$Contact
        )
        $extensionAttribute1 = $null
        if ($Contact.onPremisesExtensionAttributes) {
            $extensionAttribute1 = $Contact.onPremisesExtensionAttributes.extensionAttribute1
        }
        elseif ($Contact.extensionAttribute1) {
            $extensionAttribute1 = $Contact.extensionAttribute1
        }
        if (-not [string]::IsNullOrWhiteSpace($extensionAttribute1)) {
            $Contact | Add-Member -MemberType NoteProperty -Name "personalNotes" -Value $extensionAttribute1 -Force
        }
    }

    if ($ContactList) {
        $ContactList | ForEach-Object { Set-ContactNotesFromExtensionAttribute -Contact $_ }
    }

    if ($contactsInFolder) {
        $contactListEmails = $ContactList | ForEach-Object { Get-ContactPrimaryEmail -Contact $_ } | Where-Object { $_ }
        $removeContacts = @()
        if ($contactListEmails) {
            $removeContacts += @(
                $contactsInFolder | Where-Object {
                    $folderEmail = Get-ContactPrimaryEmail -Contact $_
                    $folderEmail -and ($folderEmail -notin $contactListEmails)
                }
            )
        }
        # Remove contacts that have duplicate email addresses. This is the only way to correctly sync
        # contacts when using email address as the "primary key"
        $removeContacts += @(
            $contactsInFolder |
                Where-Object { Get-ContactPrimaryEmail -Contact $_ } |
                Group-Object { Get-ContactPrimaryEmail -Contact $_ } |
                Where-Object { $_.Count -gt 1 } |
                ForEach-Object { $_.Group }
        )

        if ($removeContacts) {
            foreach ($contact in $removeContacts) {
                try { 
                    Remove-FolderContact -Contact $contact -ContactFolder $contactFolder | Out-Null
                    Write-LogEvent -Level Info -Message "Removed contact $($contact.displayName)"
                }
                catch {
                    Write-LogEvent -Level Error -Message "Failed to remove contact $($contact.displayName)"
                }
            }
        }

        # Get contacts in that folder again (after we've possibly removed some of them)
        try {
            $contactsInFolder = Get-FolderContact -ContactFolder $contactFolder
        }
        catch {
            throw "Failed to create contact folder $($ContactFolderName) in mailbox $($Mailbox)"
        }

        # foreach loop over the contactlist to compare to contacts in folder
        $updateContacts = @()
        foreach ($contact in $ContactList) {
            # find matching contact by primary email
            $contactEmail = Get-ContactPrimaryEmail -Contact $contact
            if (-not $contactEmail) { continue }
            $folderContact = $contactsInFolder | Where-Object {
                (Get-ContactPrimaryEmail -Contact $_) -eq $contactEmail
            } | Select-Object -First 1
            if ($folderContact) {
                # Always update existing contacts to ensure changes are pushed each run
                $contact | Add-Member -MemberType NoteProperty -Name "id" -Value $folderContact.id -Force
                $updateContacts += $contact
            }
        }

        if ($updateContacts) {
            foreach ($contact in $updateContacts) {
                try { 
                    Update-FolderContact -Contact $contact -ContactFolder $contactFolder | Out-Null
                    Write-LogEvent -Level Info -Message "Updated contact $($contact.displayName)"
                }
                catch {
                    Write-LogEvent -Level Error -Message "Failed to update contact $($updatedContact.displayName)"
                }
            }
        }
        # Get contacts in that folder again (after we've possibly modified some of them)
        try {
            $contactsInFolder = Get-FolderContact -ContactFolder $contactFolder
        }
        catch {
            throw "Failed to create contact folder $($ContactFolderName) in mailbox $($Mailbox)"
        }
    }

    # compare lists of new contacts vs old.
    if (-not $contactsInFolder) {
        $newContacts = $ContactList
    }
    else {
        $folderEmails = $contactsInFolder | ForEach-Object { Get-ContactPrimaryEmail -Contact $_ } | Where-Object { $_ }
        $newContacts = $ContactList | Where-Object {
            $contactEmail = Get-ContactPrimaryEmail -Contact $_
            $contactEmail -and ($contactEmail -notin $folderEmails)
        }
    }

    if ($newContacts) {
        foreach ($contact in $newContacts) { 
            try { 
                New-FolderContact -Contact $contact -ContactFolder $contactFolder | Out-Null
                Write-LogEvent -Level Info -Message "Created contact $($contact.displayName)"
            }
            catch {
                Write-LogEvent -Level Error -Message "Failed to create contact $($contact.displayName)"
            }
        }
    }
    if (-not $newContacts -and -not $updateContacts -and -not $removeContacts) {
        Write-LogEvent -Level Info -Message "No contacts available to sync"
    }
}

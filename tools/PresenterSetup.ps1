# Uses the existing checked Graph, paging, sharing and private-state helpers in WorkIqShared.ps1.
function Select-DemoPresenter {
    param([ValidateSet('en', 'de')] [string]$Language = 'en')
    $text = if ($Language -eq 'de') {
        @{ heading = 'Demobenutzer auswaehlen: Dateien werden zur Bearbeitung freigegeben; Mails und Einladungen gehen an ihn.'
           licensing = 'Exchange-Postfach wird geprueft; Copilot-Lizenz und Funktionen vor der Demo separat pruefen.'
           prompt = 'Nummer des Demobenutzers (Q oder leer zum Abbrechen)'
           invalid = 'Eine Nummer aus der Liste eingeben oder mit Q abbrechen.'
           selected = 'Demo-Empfaenger' }
    } else {
        @{ heading = 'Choose the demo user. Their files will be shared for editing; mail and meeting invitations go to them.'
           licensing = 'Exchange eligibility is checked; Copilot licensing and feature availability still need rehearsal.'
           prompt = 'Demo user number (Q or empty to cancel)'
           invalid = 'Enter a number from the displayed list, or Q to cancel.'
           selected = 'Demo recipient' }
    }
    $users = @(Get-WiqAudience | Sort-Object displayName, userPrincipalName)
    if (-not $users.Count) { throw 'No enabled internal users with an Exchange mailbox were found.' }
    Write-Host ''
    Write-Host $text.heading -ForegroundColor Cyan
    Write-Host $text.licensing
    for ($i = 0; $i -lt $users.Count; $i++) {
        Write-Host ("  {0,3}. {1} <{2}>" -f ($i + 1), $users[$i].displayName, $users[$i].userPrincipalName)
    }
    while ($true) {
        $choice = Read-Host $text.prompt
        if ([string]::IsNullOrWhiteSpace($choice) -or $choice -eq 'q') { throw 'Demo setup cancelled before tenant changes.' }
        $number = 0
        if ([int]::TryParse($choice, [ref]$number) -and $number -ge 1 -and $number -le $users.Count) {
            $user = Invoke-WiqGraph GET "/users/$([uri]::EscapeDataString($users[$number - 1].id))?`$select=id,displayName,userPrincipalName,mail,accountEnabled,userType,assignedPlans"
            if ($user.id -ne $users[$number - 1].id -or $user.accountEnabled -ne $true -or
                $user.userType -ne 'Member' -or -not $user.mail -or
                -not @($user.assignedPlans | Where-Object { $_.service -eq 'exchange' -and $_.capabilityStatus -eq 'Enabled' }).Count) {
                throw 'The selected user is no longer an enabled internal mailbox user; no setup was performed.'
            }
            Write-Host ("{0}: {1} <{2}>" -f $text.selected, $user.displayName, $user.mail) -ForegroundColor Green
            return $user
        }
        Write-Warning $text.invalid
    }
}

function Grant-DemoPresenterAccess($Item, $Presenter, [switch]$WhatIf, [switch]$SendInvitation) {
    if (-not $Presenter.id -or -not $Item.id -or
        ($null -eq $Item.folder -and $Item.package.type -ne 'oneNote')) {
        throw 'Only an identified demo folder or OneNote notebook may be shared with the selected user.'
    }
    $path = "/me/drive/items/$([uri]::EscapeDataString($Item.id))"
    $permissions = @(Get-WiqPages "$path/permissions")
    if (@($permissions | Where-Object { 'write' -in $_.roles -and (Test-WiqReadPermission $_ $Presenter) }).Count) { return 'same' }
    if ($WhatIf) { Write-Host "    Would grant folder/notebook editing access to $($Presenter.mail): $($Item.webUrl)"; return 'new' }
    $grant = Invoke-WiqGraph POST "$path/invite" @{
        recipients = @(@{ objectId = $Presenter.id }); requireSignIn = $true; sendInvitation = [bool]$SendInvitation
        roles = @('write'); retainInheritedPermissions = $true
    }
    if (-not @($grant.value).Count -or
        @($grant.value | Where-Object { $_.error -or -not $_.id -or 'write' -notin $_.roles }).Count) {
        throw "Editing access was not fully granted to $($Presenter.mail); do not claim this demo is ready."
    }
    $confirmed = @(Get-WiqPages "$path/permissions")
    if (-not @($confirmed | Where-Object { 'write' -in $_.roles -and (Test-WiqReadPermission $_ $Presenter) }).Count) {
        throw "Editing permission for $($Presenter.mail) was not confirmed after sharing; do not claim this demo is ready."
    }
    'new'
}

function Set-DemoPresenterPaths($Items, $Configs, $Presenter) {
    if (-not $Presenter.id -or $Presenter.id -notmatch '^[a-zA-Z0-9-]+$' -or -not $Presenter.mail) {
        throw 'A verified selected user is required for the shared demo folder.'
    }
    $root = "DemoKit-Presenters/$($Presenter.id)/Demo-Kit"
    foreach ($item in $Items) {
        $config = @($Configs | Where-Object { $_.Item.id -eq $item.id })
        if (@($config | Where-Object { $_.Config.presenterAccount -or $_.Config.executiveEvent }).Count) {
            throw "Private/account-locked demo $($item.id) cannot enter the shared folder."
        }
        if (@($config | Where-Object { $_.Config.workIqShared }).Count) { continue }
        $relative = if ($item.drivePath.StartsWith("$root/", [StringComparison]::Ordinal)) {
            $item.drivePath.Substring($root.Length + 1)
        } elseif ($item.drivePath.StartsWith('Demo-Kit/', [StringComparison]::Ordinal)) {
            $item.drivePath.Substring('Demo-Kit/'.Length)
        } else { throw "Unexpected demo upload path for $($item.id)." }
        $parts = @($relative -split '/')
        if ($parts.Count -ne 2 -or $parts[1] -cne $item.id -or
            @($parts | Where-Object { -not $_ -or $_ -in '.', '..' -or $_ -match '[\\:\r\n]' }).Count) {
            throw "Unsafe demo upload path for $($item.id)."
        }
        $item.drivePath = "$root/$relative"
    }
    $root
}

function Get-DemoDriveFolder([string]$Path, [hashtable]$Cache, [switch]$WhatIf) {
    $parts = @($Path -split '/')
    if (@($parts | Where-Object { -not $_ -or $_ -in '.', '..' -or $_ -match '[\\:\r\n]' }).Count) {
        throw 'A safe relative demo folder path is required.'
    }
    $parent = '/me/drive/root'
    $key = ''
    foreach ($name in $parts) {
        $key = if ($key) { "$key/$name" } else { $name }
        if (-not $Cache.ContainsKey($key)) {
            $folder = if ($WhatIf) {
                Invoke-WiqGraph GET "$($parent):/$([uri]::EscapeDataString($name))" -AllowNotFound
            } else { Get-WiqFolder $parent $name }
            if (-not $folder -and $WhatIf) { return $null }
            if (-not $folder.id -or $null -eq $folder.folder) { throw "Demo folder $key could not be confirmed." }
            $Cache[$key] = $folder
        }
        $parent = "/me/drive/items/$([uri]::EscapeDataString($Cache[$key].id))"
    }
    $Cache[$key]
}

function Publish-DemoPresenterRoot([string]$Path, $Items, $Links, $Presenter, [hashtable]$Cache, [string]$Language = 'en', [switch]$WhatIf) {
    if (-not $Presenter.id -or $Path -cne "DemoKit-Presenters/$($Presenter.id)/Demo-Kit") {
        throw 'Only the selected user-specific Demo-Kit folder may be shared as the start folder.'
    }
    $ordinary = @($Items | Where-Object { $_.drivePath.StartsWith("$Path/", [StringComparison]::Ordinal) })
    if (-not $ordinary.Count) { return 'same' }
    $folder = Get-DemoDriveFolder $Path $Cache -WhatIf:$WhatIf
    if (-not $folder -and $WhatIf) {
        Write-Host "    Would share one Demo-Kit folder with $($Presenter.mail); the user adds one OneDrive shortcut."
        return 'pending'
    }
    if ($folder.webUrl -notmatch '^https://') { throw 'The shared demo start folder has no confirmed web URL.' }
    $result = Grant-DemoPresenterAccess $folder $Presenter -SendInvitation -WhatIf:$WhatIf
    if (-not $WhatIf) {
        $instruction = if ($Language -eq 'de') {
            "Als $($Presenter.mail) den freigegebenen Ordner oeffnen: $($folder.webUrl) . Einmal oben 'Verknuepfung zu Meine Dateien hinzufuegen' (Add shortcut to My files) waehlen. OneDrive mit diesem Benutzer anmelden, um den Ordner im Explorer zu synchronisieren, soweit Tenant-Richtlinien es erlauben. Dateien danach in Copilot ueber + > Attach cloud files auswaehlen; nicht erneut hochladen. Freigabe und Verknuepfung bestaetigen keine Copilot-Indizierung."
        } else {
            "As $($Presenter.mail), open the shared folder: $($folder.webUrl) . Once, select 'Add shortcut to My files' at the top. Sign the OneDrive sync client in as this user to sync the folder in File Explorer, subject to tenant policy. In Copilot, select files through + > Attach cloud files; do not upload another copy. Sharing and shortcuts do not establish Copilot indexing."
        }
        $Links.root = $folder.webUrl
        $Links.manual = @($Links.manual | Where-Object { $_ }) + @($instruction)
        foreach ($item in $ordinary) {
            if ($Links.demos[$item.id].web) { $Links.demos[$item.id].presenterShared = $true }
        }
        Write-Host $instruction -ForegroundColor Cyan
    }
    $result
}

function Share-DemoNotebook([string]$Name, $Presenter, [switch]$WhatIf) {
    $query = [uri]::EscapeDataString($Name.Replace("'", "''"))
    $matches = @(Get-WiqPages "/me/drive/root/search(q='$query')?`$select=id,name,package,webUrl" |
        Where-Object { $_.name -eq $Name -and $_.package.type -eq 'oneNote' })
    if ($matches.Count -ne 1) {
        throw "Cannot identify one exact OneNote notebook '$Name' in the publisher's drive yet. Rerun after it becomes discoverable; notebook sharing is not confirmed."
    }
    Grant-DemoPresenterAccess $matches[0] $Presenter -WhatIf:$WhatIf
}

function Sync-DemoPresenterInvitation($Payload, $Presenter, [switch]$WhatIf) {
    $from = [datetime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $to = [datetime]::UtcNow.AddDays(8).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $matches = @(Get-WiqPages "/me/calendarView?startDateTime=$from&endDateTime=$to&`$select=id,subject,attendees,isOrganizer&`$top=200" |
        Where-Object { $_.subject -eq $Payload.subject -and $_.isOrganizer -eq $true })
    if ($matches.Count -gt 1) { throw "Multiple organizer events match '$($Payload.subject)'; select the canonical meeting before rerunning." }
    $attendee = @{ type = 'required'; emailAddress = @{ address = $Presenter.mail; name = $Presenter.displayName } }
    if ($matches.Count) {
        $existing = $matches[0]
        if ($Presenter.mail -in @($existing.attendees | ForEach-Object { $_.emailAddress.address })) { return 'same' }
        if (-not $WhatIf) {
            Invoke-WiqGraph PATCH "/me/events/$([uri]::EscapeDataString($existing.id))" @{ attendees = @($existing.attendees) + @($attendee) } | Out-Null
        }
        return 'updated'
    }
    $Payload.attendees = @($Payload.attendees | Where-Object { $_.emailAddress.address -ne $Presenter.mail }) + @($attendee)
    if (-not $WhatIf) {
        $created = Invoke-WiqGraph POST '/me/events' $Payload
        if (-not $created.id) { throw 'Meeting creation did not return an ID; check the organizer calendar before rerunning.' }
        if ($Presenter.mail -notin @($created.attendees | ForEach-Object { $_.emailAddress.address })) {
            throw 'The event response did not confirm the selected attendee. Check the organizer calendar and rerun; invitation creation is not confirmed.'
        }
    }
    'new'
}

function Get-DemoDeliveryKey([string]$Value) {
    $hash = [Security.Cryptography.SHA256]::Create()
    try { -join ($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value)) | ForEach-Object { $_.ToString('x2') }) }
    finally { $hash.Dispose() }
}

function Send-DemoPresenterMessage($Payload, $Operator, $Presenter, [string]$Language, [int]$WaitMinutes = 20, [string]$Revision = '', [switch]$WhatIf) {
    if (@($Payload.message.toRecipients).Count -ne 1 -or $Payload.message.toRecipients[0].emailAddress.address -ne $Presenter.mail) {
        throw 'Demo message recipient does not match the selected mailbox.'
    }
    if (-not $Payload.message.subject -or -not $Payload.message.from.emailAddress.address) { throw 'Demo message subject and sender are required.' }
    if ($WhatIf) { return 'new' }
    $context = Get-MgContext
    if (-not $context.TenantId -or -not $Operator.id -or -not $Presenter.id) { throw 'Verified tenant, publisher and recipient identities are required.' }
    $private = Get-WiqFolder '/me/drive/root' "DemoKit-Presenter-State-$($context.TenantId)-$($Presenter.id)-$Language"
    $lock = $null
    try {
        $lock = Invoke-WiqGraph POST "/me/drive/items/$($private.id)/children" @{
            name = 'installer-lock'; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail'
        }
        $saved = Invoke-WiqGraph GET "/me/drive/items/$($private.id):/state.json" -AllowNotFound
        if ($saved) {
            $state = Read-WiqJson "/me/drive/items/$($private.id):/state.json:/content"
            if ($state.version -ne 1 -or $state.tenantId -ne $context.TenantId -or
                $state.publisherId -ne $Operator.id -or $state.recipientId -ne $Presenter.id -or $state.language -ne $Language -or
                $state.deliveries -isnot [array] -or
                @($state.deliveries | Where-Object { -not $_.key -or $_.status -notin @('pending', 'submitted') }).Count) {
                throw 'Demo delivery state does not belong to this tenant, publisher, recipient and language.'
            }
        } else {
            $state = [pscustomobject]@{ version = 1; tenantId = $context.TenantId; publisherId = $Operator.id
                recipientId = $Presenter.id; language = $Language; deliveries = @() }
        }
        $key = Get-DemoDeliveryKey "$($Payload.message.from.emailAddress.address.ToLowerInvariant())`n$($Payload.message.subject)`n$Revision"
        $entries = @($state.deliveries | Where-Object { $_.key -eq $key })
        if ($entries.Count -gt 1) { throw 'Duplicate demo delivery state; no automatic resend.' }
        if ($entries.Count -and $entries[0].status -eq 'submitted') { return 'same' }
        if ($entries.Count) {
            throw "Submission outcome unresolved for '$($Payload.message.subject)' to $($Presenter.mail). Check the publisher's Sent Items or Exchange trace before repairing the private pending entry; no automatic resend."
        }
        $state.deliveries = @($state.deliveries) + @([pscustomobject]@{ key = $key; status = 'pending'; at = [datetime]::UtcNow.ToString('o') })
        Save-WiqState $state $private.id
        try { Send-WiqMessage $Payload.message $WaitMinutes }
        catch {
            if ($_.Exception.Data['HttpStatus'] -in @(400, 401, 403, 404, 413, 429)) {
                $state.deliveries = @($state.deliveries | Where-Object { $_.key -ne $key })
                Save-WiqState $state $private.id
            }
            throw
        }
        ($state.deliveries | Where-Object { $_.key -eq $key }).status = 'submitted'
        Save-WiqState $state $private.id
        'new'
    } finally {
        if ($lock) { Invoke-WiqGraph DELETE "/me/drive/items/$($lock.id)" | Out-Null }
    }
}

# Shared synthetic evidence; delegated Graph access never implies access to somebody else's mailbox.
function Invoke-WiqGraph {
    param([string]$Method, [string]$Path, $Body, [switch]$AllowNotFound, [string]$InputFile)
    if ($Path -notmatch '^/(?!/)' -or $Path -match '[\r\n]') { throw 'WIQ26: invalid Graph path.' }
    $params = @{ Method = $Method; Uri = "https://graph.microsoft.com/v1.0$Path"; SkipHttpErrorCheck = $true; StatusCodeVariable = 'status'; OutputType = 'PSObject' }
    if ($InputFile) { $params.InputFilePath = $InputFile; $params.ContentType = 'application/octet-stream' }
    elseif ($null -ne $Body) { $params.Body = ConvertTo-Json $Body -Depth 30 -Compress; $params.ContentType = 'application/json' }
    $status = 0
    $result = Invoke-MgGraphRequest @params
    if ($status -eq 404 -and $AllowNotFound -and $Method -eq 'GET') { return $null }
    if ($status -lt 200 -or $status -ge 300 -or $status -eq 207) {
        $failure = New-Object InvalidOperationException "WIQ26: Graph $Method $Path returned HTTP $status. Inspect the installer log; no automatic write retry."
        $failure.Data['HttpStatus'] = [int]$status
        $failure.Data['GraphCode'] = [string]$result.error.code
        throw $failure
    }
    $result
}

function Send-WiqMessage($Message, [int]$WaitMinutes) {
    $deadline = [datetime]::UtcNow.AddMinutes([math]::Max(0, $WaitMinutes))
    $announced = $false
    while ($true) {
        try {
            Invoke-WiqGraph POST '/me/sendMail' @{ message = $Message; saveToSentItems = $true } | Out-Null
            return
        } catch {
            if ($_.Exception.Data['HttpStatus'] -eq 403 -and $_.Exception.Data['GraphCode'] -eq 'ErrorSendAsDenied' -and [datetime]::UtcNow -lt $deadline) {
                if (-not $announced) { Write-Host 'WIQ26: waiting for the existing Exchange Send As grant to replicate...'; $announced = $true }
                Start-Sleep -Seconds 30
            } else { throw }
        }
    }
}

function Get-WiqPages([string]$Path) {
    do {
        $page = Invoke-WiqGraph GET $Path
        @($page.value) | Where-Object { $_ }
        $next = $page.'@odata.nextLink'
        if ($next) {
            if (-not $next.StartsWith('https://graph.microsoft.com/v1.0/', [StringComparison]::Ordinal)) { throw 'WIQ26: untrusted pagination URL.' }
            $Path = $next.Substring('https://graph.microsoft.com/v1.0'.Length)
        }
    } while ($next)
}

function Get-WiqAudience {
    @(Get-WiqPages '/users?$select=id,displayName,userPrincipalName,mail,accountEnabled,userType,assignedPlans&$top=999') |
        Where-Object {
            $_.id -and $_.accountEnabled -eq $true -and $_.userType -eq 'Member' -and $_.mail -and
            $_.userPrincipalName -notmatch '#EXT#' -and
            @($_.assignedPlans | Where-Object { $_.service -eq 'exchange' -and $_.capabilityStatus -eq 'Enabled' }).Count
        } | Sort-Object { $_.id } -Unique
}

function Get-WiqFolder([string]$Parent, [string]$Name) {
    $folder = Invoke-WiqGraph GET "$($Parent):/$([uri]::EscapeDataString($Name))" -AllowNotFound
    if ($folder) {
        if (-not $folder.folder) { throw "WIQ26: $Name exists but is not a folder." }
        return $folder
    }
    Invoke-WiqGraph POST "$Parent/children" @{ name = $Name; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail' }
}

function Read-WiqJson([string]$Path) {
    $value = Invoke-WiqGraph GET $Path
    if ($value -is [string]) { return ($value | ConvertFrom-Json) }
    if ($value -is [byte[]]) { return ([Text.Encoding]::UTF8.GetString($value) | ConvertFrom-Json) }
    $value
}

function Save-WiqState($State, [string]$FolderId) {
    $file = [IO.Path]::GetTempFileName()
    try {
        [IO.File]::WriteAllText($file, ($State | ConvertTo-Json -Depth 30), (New-Object Text.UTF8Encoding $false))
        Invoke-WiqGraph PUT "/me/drive/items/$($FolderId):/state.json:/content" -InputFile $file | Out-Null
    } finally { Remove-Item -LiteralPath $file -Force }
}

function Test-WiqReadPermission($Permission, $User) {
    if (-not @($Permission.roles | Where-Object { $_ -in 'read', 'write', 'owner' }).Count) { return $false }
    $identities = @($Permission.grantedToV2, $Permission.grantedTo) + @($Permission.grantedToIdentitiesV2) + @($Permission.grantedToIdentities)
    foreach ($identity in $identities | Where-Object { $_ }) {
        foreach ($person in @($identity.user, $identity.siteUser) | Where-Object { $_ }) {
            if ($person.id -eq $User.id -or $person.email -eq $User.mail -or
                ($person.loginName -and ($person.loginName -split '\|')[-1] -in @($User.mail, $User.userPrincipalName))) { return $true }
        }
    }
    $false
}

function Invoke-WorkIqSharedSetup {
    param($Config, $Item, $Me, $Senders, $Mails, [string]$Language = 'en', [switch]$WhatIf, [int]$SendAsWaitMinutes = 20)
    if ($Config.scenario -ne 'WIQ26' -or $Config.folder -ne 'Demo-Kit-Work-IQ-WIQ26') { throw 'WIQ26: unsupported shared destination.' }
    $context = Get-MgContext
    if (-not $context.TenantId -or -not $Me.id) { throw 'WIQ26: tenant and publisher identity are required.' }
    $audience = @(Get-WiqAudience)
    if (-not $audience.Count) { throw 'WIQ26: no enabled internal Exchange-enabled users found.' }
    $drive = Invoke-WiqGraph GET '/sites/root/drive?$select=id,webUrl'
    if (-not $drive.id) { throw 'WIQ26: the existing tenant root document library is unavailable. No site or admin role will be created.' }
    $root = "/drives/$($drive.id)/root"
    foreach ($mail in $Mails) {
        if (-not $Senders[$mail.FromName] -or $Senders[$mail.FromName].address -ne $mail.FromAddress) {
            throw "WIQ26: sender $($mail.FromName) is not a verified prepared mailbox."
        }
    }
    if (@($Config.documents).Count -ne 48 -or @($Config.documents | Sort-Object -Unique).Count -ne 48) {
        throw 'WIQ26: the complete 24-document bilingual corpus is required.'
    }
    foreach ($path in $Config.documents) {
        if ($path -notmatch '^(en|de)/WIQ-D\d{2}_[a-z0-9_]+\.docx$' -or
            -not (Test-Path -LiteralPath (Join-Path $Item.dir ($path -replace '/', '\')))) { throw "WIQ26: missing or invalid corpus file $path." }
    }
    Write-Host ("WIQ26: {0} eligible internal users; {1} documents; {2} messages per user." -f $audience.Count, @($Config.documents).Count, @($Mails).Count)
    if ($WhatIf) {
        Write-Host 'WIQ26 WhatIf: would publish to the existing root SharePoint library, grant folder-only read access, submit primary-mailbox evidence and invite the audience. No tenant writes.'
        return
    }

    # The private OneDrive lock serializes runs on different PCs without expiring an in-flight send.
    $private = Get-WiqFolder '/me/drive/root' "DemoKit-WIQ26-State-$($context.TenantId)"
    $lock = $null
    $changed = $false
    $created = $false
    try {
        try {
            $lock = Invoke-WiqGraph POST "/me/drive/items/$($private.id)/children" @{ name = 'installer-lock'; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail' }
        } catch { throw "WIQ26: publisher lock could not be acquired. Close any other setup. After a crashed run, remove only installer-lock in the private state folder after confirming no run is active. $_" }
        $saved = Invoke-WiqGraph GET "/me/drive/items/$($private.id):/state.json" -AllowNotFound
        if ($saved) {
            $state = Read-WiqJson "/me/drive/items/$($private.id):/state.json:/content"
            if ($state.version -ne 1 -or $state.tenantId -ne $context.TenantId -or $state.publisherId -ne $Me.id -or $state.scenario -ne 'WIQ26') {
                throw 'WIQ26: private state ownership mismatch; no shared evidence changed.'
            }
        } else {
            $state = [pscustomobject]@{ version = 1; scenario = 'WIQ26'; tenantId = $context.TenantId; publisherId = $Me.id; deliveries = @(); event = $null }
            Save-WiqState $state $private.id
        }
        $folder = Invoke-WiqGraph GET "$($root):/$($Config.folder)" -AllowNotFound
        if ($folder) {
            $marker = Read-WiqJson "/drives/$($drive.id)/items/$($folder.id):/demo-kit-owner.json:/content"
            if (-not $folder.folder -or $marker.scenario -ne 'WIQ26' -or $marker.owner -ne 'DemoKit' -or $marker.version -ne 1) {
                throw 'WIQ26: shared folder ownership could not be verified; refusing to adopt it.'
            }
        } else {
            $folder = Get-WiqFolder $root $Config.folder
            $created = $true
            $markerFile = [IO.Path]::GetTempFileName()
            try {
                [IO.File]::WriteAllText($markerFile, '{"owner":"DemoKit","scenario":"WIQ26","version":1}')
                Invoke-WiqGraph PUT "/drives/$($drive.id)/items/$($folder.id):/demo-kit-owner.json:/content" -InputFile $markerFile | Out-Null
            } finally { Remove-Item -LiteralPath $markerFile -Force }
        }
        $base = "/drives/$($drive.id)/items/$($folder.id)"
        foreach ($lang in 'en', 'de') { Get-WiqFolder $base $lang | Out-Null }
        foreach ($path in $Config.documents) {
            $file = Join-Path $Item.dir ($path -replace '/', '\')
            $remote = Invoke-WiqGraph GET "$($base):/$($path)?`$select=id,file" -AllowNotFound
            if (-not $remote -or -not $remote.file.hashes.quickXorHash -or $remote.file.hashes.quickXorHash -ne (Get-QuickXor $file)) {
                Invoke-WiqGraph PUT "$($base):/$($path):/content" -InputFile $file | Out-Null
                $changed = $true
            }
        }
        $permissions = @(Get-WiqPages "$base/permissions")
        foreach ($user in $audience) {
            if (@($permissions | Where-Object { Test-WiqReadPermission $_ $user }).Count) { continue }
            $grant = Invoke-WiqGraph POST "$base/invite" @{
                recipients = @(@{ objectId = $user.id }); requireSignIn = $true; sendInvitation = $false
                roles = @('read'); retainInheritedPermissions = $true
            }
            if (-not @($grant.value).Count -or @($grant.value | Where-Object { $_.error -or -not $_.id -or 'read' -notin $_.roles }).Count) {
                throw "WIQ26: incomplete read grant for directory user $($user.id). No mail will be sent."
            }
            $changed = $true
        }
        Write-Host "WIQ26 shared evidence: $($folder.webUrl)"
        foreach ($user in $audience) {
            foreach ($mail in $Mails) {
                $key = "$Language/$($user.id)/$($mail.Subject)"
                $entries = @($state.deliveries | Where-Object { $_.key -eq $key })
                if ($entries.Count -gt 1) { throw 'WIQ26: duplicate private delivery state.' }
                if ($entries.Count -and $entries[0].status -eq 'submitted') { continue }
                if ($entries.Count) { throw "WIQ26: delivery outcome unresolved for $key. Check the sender's Sent Items/Exchange trace before clearing this pending entry; never automatically resend." }
                $state.deliveries = @($state.deliveries) + @([pscustomobject]@{ key = $key; status = 'pending'; at = [datetime]::UtcNow.ToString('o') })
                Save-WiqState $state $private.id
                $sender = $Senders[$mail.FromName]
                $message = @{
                    subject = $mail.Subject
                    body = @{ contentType = 'Text'; content = "$($mail.Body)`n`nWIQ26 shared evidence / Gemeinsame Quellen: $($folder.webUrl)`n" }
                    from = @{ emailAddress = @{ address = $sender.address; name = $mail.FromName } }
                    toRecipients = @(@{ emailAddress = @{ address = $user.mail } })
                }
                try {
                    Send-WiqMessage $message $SendAsWaitMinutes
                } catch {
                    if ($_.Exception.Data['HttpStatus'] -in @(400, 401, 403, 404, 413, 429)) {
                        # An explicit rejection is safe to retry later; timeouts/unknown outcomes are not.
                        $state.deliveries = @($state.deliveries | Where-Object { $_.key -ne $key })
                        Save-WiqState $state $private.id
                    }
                    throw
                }
                ($state.deliveries | Where-Object { $_.key -eq $key }).status = 'submitted'
                Save-WiqState $state $private.id
                $changed = $true
            }
        }
        $attendees = @($audience | Where-Object { $_.id -ne $Me.id } | ForEach-Object { @{ type = 'required'; emailAddress = @{ address = $_.mail; name = $_.displayName } } })
        $audienceIds = @($audience.id | Sort-Object)
        $day = (Get-Date).Date.AddDays(1)
        while ($day.DayOfWeek -ne [DayOfWeek]::Thursday) { $day = $day.AddDays(1) }
        $start = $day.AddHours(10)
        $times = @{
            start = @{ dateTime = $start.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }
            end = @{ dateTime = $start.AddMinutes(30).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }
        }
        if ($state.event -and $state.event.status -ne 'created') { throw 'WIQ26: calendar creation outcome unresolved. Verify the organizer calendar before repairing private event state.' }
        if ($state.event) {
            $event = Invoke-WiqGraph GET "/me/events/$($state.event.id)"
            if (-not $event -or $event.subject -ne $state.event.subject) { throw 'WIQ26: calendar ownership mismatch.' }
            if (-not $event.end.dateTime -or $event.end.timeZone -ne 'UTC' -or $event.isCancelled) { throw 'WIQ26: calendar is cancelled or its UTC end time is unavailable.' }
            $end = [datetime]::Parse($event.end.dateTime, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime()
            if ($end -le [datetime]::UtcNow) {
                Invoke-WiqGraph PATCH "/me/events/$($state.event.id)" $times | Out-Null
                $changed = $true
            }
            if (($audienceIds -join ',') -ne (@($state.event.audience) -join ',')) {
                # Preserve any attendees added manually; only add newly eligible users.
                $all = @($event.attendees)
                foreach ($attendee in $attendees) {
                    if ($attendee.emailAddress.address -notin @($all | ForEach-Object { $_.emailAddress.address })) { $all += $attendee }
                }
                Invoke-WiqGraph PATCH "/me/events/$($state.event.id)" @{ attendees = $all } | Out-Null
                $state.event.audience = $audienceIds
                Save-WiqState $state $private.id
                $changed = $true
            }
        } else {
            $subject = $Config.event.subject.$Language
            if (-not $subject -or -not $Config.event.body.$Language) { throw 'WIQ26: localized calendar evidence is missing.' }
            $state.event = [pscustomobject]@{ status = 'pending'; id = ''; subject = $subject; audience = $audienceIds; transactionId = [guid]::NewGuid().ToString() }
            Save-WiqState $state $private.id
            $event = Invoke-WiqGraph POST '/me/events' @{
                subject = $subject
                body = @{ contentType = 'Text'; content = "$($Config.event.body.$Language)`n$($folder.webUrl)" }
                start = $times.start
                end = $times.end
                attendees = $attendees; transactionId = $state.event.transactionId
            }
            if (-not $event.id) { throw 'WIQ26: event creation did not return an identifier.' }
            $state.event.id = $event.id; $state.event.status = 'created'
            Save-WiqState $state $private.id
            $changed = $true
        }
        Write-Host 'WIQ26: folder access granted; messages accepted for submission; invitations created. Delivery, invitation processing, Copilot licensing and indexing still require presenter checks.'
        @{ name = 'WIQ26 - shared SharePoint evidence / Gemeinsame SharePoint-Quellen'; url = $folder.webUrl
           result = $(if ($created) { 'new' } elseif ($changed) { 'updated' } else { 'same' }) }
    } finally {
        if ($lock) { Invoke-WiqGraph DELETE "/me/drive/items/$($lock.id)" | Out-Null }
    }
}

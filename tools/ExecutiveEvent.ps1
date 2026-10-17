# Loaded only for tenant.json executiveEvent configuration. No writes on import.
$script:executiveEventCloudSync = @{}
function Get-ExecutiveEventSettings($Config, $Settings = $null) {
    $path = Join-Path $env:LOCALAPPDATA 'DemoKit\executive-event-settings.json'
    if ($null -eq $Settings) {
        if (-not (Test-Path -LiteralPath $path)) {
            throw "LEF26 requires private settings at $path or in the signed-in demo admin's private OneDrive cache. No team/site is created."
        }
        $Settings = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    foreach ($key in @('tenantId', 'teamId')) {
        $guid = [guid]::Empty
        if (-not [guid]::TryParse([string]$settings.$key, [ref]$guid) -or $guid -eq [guid]::Empty) {
            throw "LEF26 settings: $key must be a real nonempty GUID."
        }
    }
    if ($settings.account -ne $Config.account -or
        ($settings.account -split '@')[-1] -ne $Config.domain) {
        throw 'LEF26 settings account must match the explicitly approved demo account/domain.'
    }
    if (-not $settings.channelId -or $settings.channelId -notmatch '^19:[^/\s]+@thread\.(tacv2|skype)$') {
        throw 'LEF26 settings require the actual Teams channel ID, not a URL or display name.'
    }
    $settings
}

function Assert-ExecutiveEventIdentity($Config, $Settings, $Context, $Me) {
    if (-not $Context -or $Context.TenantId -ne $Settings.tenantId -or
        $Context.Account -ne $Config.account -or $Me.userPrincipalName -ne $Config.account -or
        $Me.mail -ne $Config.account) {
        throw 'LEF26 blocked: Graph tenant, signed-in account and mailbox must exactly match the approved private settings.'
    }
}

function Invoke-ExecutiveEventGraph {
    param([string]$Method, [string]$Path, $Body, [string]$InputFile, [switch]$Optional)
    if ($Path -notmatch '^/' -and $Path -notmatch '^https://graph\.microsoft\.com/v1\.0/') {
        throw 'LEF26 rejected an unexpected Graph URL.'
    }
    $uri = if ($Path.StartsWith('/')) { "https://graph.microsoft.com/v1.0$Path" } else { $Path }
    $params = @{ Method = $Method; Uri = $uri; OutputType = 'Hashtable'; SkipHttpErrorCheck = $true; StatusCodeVariable = 'status' }
    $resource = ([uri]$uri).AbsolutePath -replace '^/v1\.0', ''
    if ($resource -match '^/me/(events|calendarView)') { $params.Headers = @{ Prefer = 'outlook.timezone="W. Europe Standard Time"' } }
    if ($resource -match '^/me/(messages|mailFolders)') { $params.Headers = @{ Prefer = 'IdType="ImmutableId"' } }
    if ($InputFile) {
        $params.InputFilePath = $InputFile
        $params.ContentType = 'application/octet-stream'
    } elseif ($null -ne $Body) {
        $params.Body = ConvertTo-Json -InputObject $Body -Depth 20 -Compress
        $params.ContentType = 'application/json; charset=utf-8'
    }
    $response = Invoke-MgGraphRequest @params
    if ($Optional -and [int]$status -eq 404) { return $null }
    if ([int]$status -lt 200 -or [int]$status -ge 300) {
        throw "LEF26 Graph $Method $Path returned HTTP $status : $($response | ConvertTo-Json -Depth 5 -Compress)"
    }
    $response
}

function Get-ExecutiveEventPages([string]$Path) {
    do {
        $page = Invoke-ExecutiveEventGraph GET $Path
        @($page.value) | Where-Object { $null -ne $_ }
        $Path = $page.'@odata.nextLink'
    } while ($Path)
}

function Get-ExecutiveEventText($Value, [string]$Language) {
    if ($Value -is [string]) { return $Value }
    $text = $Value.$Language
    if (-not $text) { throw "LEF26 is missing $Language text." }
    [string]$text
}

function Get-ExecutiveEventLedgerPath($Settings, [string]$Language) {
    Join-Path $env:LOCALAPPDATA "DemoKit\executive-event\$($Settings.tenantId)\$Language.json"
}

function Save-ExecutiveEventLedger($Ledger, [string]$Path) {
    New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force | Out-Null
    $json = ConvertTo-Json -InputObject $Ledger -Depth 25
    [IO.File]::WriteAllText("$Path.tmp", $json, (New-Object Text.UTF8Encoding $false))
    Move-Item -LiteralPath "$Path.tmp" -Destination $Path -Force
    if ($script:executiveEventCloudSync.ContainsKey($Path)) {
        $sync = $script:executiveEventCloudSync[$Path]
        Save-ExecutiveEventCloudState $sync.Config $sync.Settings $sync.Me $sync.Language $Ledger
    }
}

function Get-ExecutiveEventLedger($Settings, [string]$Language) {
    $path = Get-ExecutiveEventLedgerPath $Settings $Language
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $value = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($value.scenario -ne 'LEF26' -or $value.tenantId -ne $Settings.tenantId -or
        $value.account -ne $Settings.account -or $value.language -ne $Language -or
        $value.teamId -ne $Settings.teamId -or $value.channelId -cne $Settings.channelId) {
        throw 'LEF26 ledger does not match the approved tenant/account/channel/language. Refusing to adopt unrelated objects.'
    }
    $value
}

function Assert-ExecutiveEventCloudState($State, $Config, $Me, [string]$Language) {
    if ($State.schemaVersion -ne 1 -or $State.language -cne $Language -or -not $State.settings) {
        throw 'LEF26 private setup cache has an unsupported schema or language.'
    }
    $settings = Get-ExecutiveEventSettings $Config $State.settings
    Assert-ExecutiveEventIdentity $Config $settings (Get-MgContext) $Me
    $ledger = $State.ledger
    if ($null -ne $ledger -and ($ledger.scenario -ne 'LEF26' -or
        $ledger.tenantId -ne $settings.tenantId -or $ledger.account -ne $settings.account -or
        $ledger.language -cne $Language -or $ledger.teamId -ne $settings.teamId -or
        $ledger.channelId -cne $settings.channelId -or -not $ledger.driveId -or
        -not $ledger.parentId -or -not $ledger.rootId)) {
        throw 'LEF26 private setup cache contains a mismatched ownership ledger.'
    }
}

function Get-ExecutiveEventCloudStatePath([ValidateSet('en', 'de')][string]$Language) {
    "/me/drive/root:/DemoKit/executive-event-state-$Language-v1.json"
}

function Initialize-ExecutiveEventState($Config, $Me, [ValidateSet('en', 'de')][string]$Language, [switch]$WhatIf) {
    $context = Get-MgContext
    if (-not $context -or $context.Account -ne $Config.account -or
        $Me.userPrincipalName -ne $Config.account -or $Me.mail -ne $Config.account) {
        throw 'LEF26 blocked: private setup recovery requires the approved demo admin account.'
    }
    $state = $null
    $otherLanguage = if ($Language -eq 'de') { 'en' } else { 'de' }
    foreach ($cachedLanguage in @($Language, $otherLanguage)) {
        $item = Invoke-ExecutiveEventGraph GET (Get-ExecutiveEventCloudStatePath $cachedLanguage) -Optional
        if (-not $item) { continue }
        $download = [uri]$item.'@microsoft.graph.downloadUrl'
        if (-not $item.file -or -not $download.IsAbsoluteUri -or $download.Scheme -ne 'https') {
            throw 'LEF26 private setup cache is not a downloadable OneDrive file.'
        }
        $state = Invoke-RestMethod -Uri $download.AbsoluteUri -UseBasicParsing
        if ($state -is [string]) { $state = $state.TrimStart([char]0xFEFF) | ConvertFrom-Json }
        Assert-ExecutiveEventCloudState $state $Config $Me $cachedLanguage
        if ($cachedLanguage -cne $Language) {
            $state = [pscustomobject]@{
                schemaVersion = 1; language = $Language; settings = $state.settings
                ledger = (Get-ExecutiveEventLedger $state.settings $Language)
            }
        }
        break
    }
    $settingsPath = Join-Path $env:LOCALAPPDATA 'DemoKit\executive-event-settings.json'
    $local = if (Test-Path -LiteralPath $settingsPath) { Get-ExecutiveEventSettings $Config } else { $null }
    if (-not $state) {
        if (-not $local) { throw 'LEF26 approved private settings are missing locally and in OneDrive. An authorized operator must approve the destination once; no channel is guessed or created.' }
        Assert-ExecutiveEventIdentity $Config $local $context $Me
        $state = [pscustomobject]@{ schemaVersion = 1; language = $Language; settings = $local; ledger = (Get-ExecutiveEventLedger $local $Language) }
        Assert-ExecutiveEventCloudState $state $Config $Me $Language
    } elseif ($local -and ($local.tenantId -ne $state.settings.tenantId -or
        $local.account -ne $state.settings.account -or $local.teamId -ne $state.settings.teamId -or
        $local.channelId -cne $state.settings.channelId)) {
        throw 'LEF26 local settings differ from the private OneDrive cache. Nothing was overwritten.'
    }
    $ledgerPath = Get-ExecutiveEventLedgerPath $state.settings $Language
    $localLedger = Get-ExecutiveEventLedger $state.settings $Language
    if ($localLedger -and (($localLedger | ConvertTo-Json -Depth 25 -Compress) -cne
        ($state.ledger | ConvertTo-Json -Depth 25 -Compress))) {
        throw 'LEF26 local ownership ledger differs from the private OneDrive cache. Preserve both and review; nothing was overwritten.'
    }
    if (-not $WhatIf) {
        $restore = @(@{ Path = $settingsPath; Value = $state.settings })
        if ($state.ledger) { $restore += @{ Path = $ledgerPath; Value = $state.ledger } }
        foreach ($file in $restore) {
            if (Test-Path -LiteralPath $file.Path) { continue }
            [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($file.Path)) | Out-Null
            $bytes = [Text.Encoding]::UTF8.GetBytes(($file.Value | ConvertTo-Json -Depth 25))
            $stream = [IO.File]::Open($file.Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
            try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
            Write-Host "  LEF26 restored private setup state from OneDrive: $($file.Path)"
        }
    } elseif (-not $local -or ($state.ledger -and -not $localLedger)) {
        Write-Host '  LEF26 WhatIf: would restore approved private settings and ownership ledger from OneDrive. No writes.'
    }
    $state
}

function Save-ExecutiveEventCloudState($Config, $Settings, $Me, [ValidateSet('en', 'de')][string]$Language, $Ledger) {
    $state = [pscustomobject]@{ schemaVersion = 1; language = $Language; settings = $Settings; ledger = $Ledger }
    Assert-ExecutiveEventCloudState $state $Config $Me $Language
    $folder = Invoke-ExecutiveEventGraph GET '/me/drive/root:/DemoKit' -Optional
    if (-not $folder) {
        $folder = Invoke-ExecutiveEventGraph POST '/me/drive/root/children' @{
            name = 'DemoKit'; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail'
        }
    }
    if ($null -eq $folder.folder) { throw 'LEF26 private cache parent DemoKit is not a OneDrive folder.' }
    [void](Invoke-ExecutiveEventGraph PUT "$(Get-ExecutiveEventCloudStatePath $Language):/content" $state)
}

function Enable-ExecutiveEventCloudSync($Config, $Settings, $Me, [ValidateSet('en', 'de')][string]$Language) {
    $path = Get-ExecutiveEventLedgerPath $Settings $Language
    $script:executiveEventCloudSync[$path] = @{ Config = $Config; Settings = $Settings; Me = $Me; Language = $Language }
    Save-ExecutiveEventCloudState $Config $Settings $Me $Language (Get-ExecutiveEventLedger $Settings $Language)
}

function Test-ExecutiveEventDestination($Config, $Settings, $Me, [string]$Language, $LedgerOverride = $null) {
    Assert-ExecutiveEventIdentity $Config $Settings (Get-MgContext) $Me
    $team = [uri]::EscapeDataString($Settings.teamId)
    $channel = [uri]::EscapeDataString($Settings.channelId)
    $channelPath = "/teams/$team/channels/$channel"
    $channelInfo = Invoke-ExecutiveEventGraph GET $channelPath
    $folder = Invoke-ExecutiveEventGraph GET "$channelPath/filesFolder"
    if (-not $folder.folder -or -not $folder.parentReference.driveId -or
        $folder.webUrl -notmatch '^https://[^/]+\.sharepoint\.com/' -or $folder.webUrl -match '/personal/') {
        throw 'LEF26 requires the verified SharePoint files folder of the approved demo channel.'
    }
    # Read access is verified here; actual write permission is confirmed only by setup.
    $messages = @(Get-ExecutiveEventPages "$channelPath/messages?`$top=50")
    $ledger = if ($null -ne $LedgerOverride) { $LedgerOverride } else { Get-ExecutiveEventLedger $Settings $Language }
    if ($ledger -and ($ledger.driveId -cne $folder.parentReference.driveId -or $ledger.parentId -cne $folder.id)) {
        throw 'LEF26 channel files destination changed since the ledger was created.'
    }
    [pscustomobject]@{ settings = $Settings; folder = $folder; channel = $channelInfo; channelPath = $channelPath; ledger = $ledger; messageCount = $messages.Count }
}

function Get-ExecutiveEventTransaction([string]$Key) {
    $hash = [Security.Cryptography.SHA256]::Create()
    try { $bytes = $hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($Key)) }
    finally { $hash.Dispose() }
    (New-Object Guid (,[byte[]]$bytes[0..15])).ToString()
}

function Invoke-ExecutiveEventSetup {
    param($Config, $Item, $Destination, $Senders, [string]$Language, [switch]$WhatIf)
    foreach ($event in $Config.events) {
        foreach ($person in $event.attendees) {
            $sender = $Senders[$person]
            $expected = $Config.people.$person
            if (-not $sender -or $sender.created -or
                -not $expected -or $sender.address -ne $expected -or
                ($sender.address -split '@')[-1] -ne $Config.domain) {
                throw "LEF26 cannot invite unresolved, substitute or external attendee $person."
            }
        }
    }
    if ($WhatIf) {
        Write-Host "  LEF26 WhatIf: $(@($Config.sharedFiles).Count) shared documents, $(@($Config.events).Count) fixed-date calendar entries, $(@($Config.threads).Count) truthful Teams threads. No writes."
        return @()
    }
    $settings = $Destination.settings
    $ledgerPath = Get-ExecutiveEventLedgerPath $settings $Language
    $ledger = $Destination.ledger
    $drive = [uri]::EscapeDataString($Destination.folder.parentReference.driveId)
    $parent = [uri]::EscapeDataString($Destination.folder.id)
    $rootName = "Demo-Kit-LEF26-$Language"
    $root = Invoke-ExecutiveEventGraph GET "/drives/$drive/items/${parent}:/${rootName}" -Optional
    if (-not $ledger) {
        if ($root) { throw 'LEF26 folder already exists without a matching private ledger. Do not overwrite or guess ownership.' }
        $root = Invoke-ExecutiveEventGraph POST "/drives/$drive/items/$parent/children" @{
            name = $rootName; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail'
        }
        $ledger = [pscustomobject]@{
            scenario = 'LEF26'; tenantId = $settings.tenantId; account = $settings.account; language = $Language
            teamId = $settings.teamId; channelId = $settings.channelId; driveId = $Destination.folder.parentReference.driveId
            parentId = $Destination.folder.id; rootId = $root.id; files = @(); events = @(); threads = @(); mailReplies = @()
        }
        Save-ExecutiveEventLedger $ledger $ledgerPath
    } elseif (-not $root -or $root.id -cne $ledger.rootId) {
        throw 'LEF26 tracked SharePoint folder is missing or changed. Refusing an automatic replacement.'
    }
    $rootId = [uri]::EscapeDataString($root.id)
    $links = @([pscustomobject]@{ name = "LEF26 SharePoint ($Language)"; url = $root.webUrl })
    foreach ($template in $Config.sharedFiles) {
        $relative = $template.Replace('{lang}', $Language)
        if ($relative -notmatch '^inputs/(de|en)/before/[A-Za-z_]+\.docx$') { throw "LEF26 unsafe shared file $relative" }
        $file = Join-Path $Item.dir ($relative -replace '/', '\')
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "LEF26 missing shared input $relative" }
        $name = [IO.Path]::GetFileName($file)
        $encoded = [uri]::EscapeDataString($name)
        $entry = @($ledger.files | Where-Object { $_.name -eq $name })
        if ($entry.Count -gt 1) { throw "LEF26 duplicate file ledger entry: $name" }
        $remote = Invoke-ExecutiveEventGraph GET "/drives/$drive/items/${rootId}:/$encoded" -Optional
        if ($remote -and (-not $entry.Count -or $entry[0].id -cne $remote.id)) {
            throw "LEF26 refuses untracked shared file $name."
        }
        if ($entry.Count -and (-not $remote -or $entry[0].cTag -cne $remote.cTag)) {
            throw "LEF26 baseline $name was changed or removed outside setup. Preserve it and review before restoring."
        }
        $sha = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant()
        if (-not $entry.Count -or $entry[0].sha256 -ne $sha) {
            $remote = Invoke-ExecutiveEventGraph PUT "/drives/$drive/items/${rootId}:/${encoded}:/content" -InputFile $file
            $ledger.files = @($ledger.files | Where-Object { $_.name -ne $name }) + @([pscustomobject]@{
                name = $name; id = $remote.id; sha256 = $sha; cTag = $remote.cTag; webUrl = $remote.webUrl
            })
            Save-ExecutiveEventLedger $ledger $ledgerPath
        }
        $links += [pscustomobject]@{ name = "LEF26 $name"; url = $remote.webUrl }
    }
    foreach ($event in $Config.events) {
        $key = [string]$event.key
        $entry = @($ledger.events | Where-Object { $_.key -eq $key })
        if ($entry.Count -gt 1) { throw "LEF26 duplicate event ledger entry: $key" }
        $subject = Get-ExecutiveEventText $event.subject $Language
        $marker = "LEF26-$Language-$key"
        if ($entry.Count) {
            $actual = Invoke-ExecutiveEventGraph GET "/me/events/$([uri]::EscapeDataString($entry[0].id))"
            $expectedStart = [datetime]::ParseExact("$($event.date) $($event.start)", 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
            if ($actual.subject -ne $subject -or $actual.categories -notcontains $marker -or
                $actual.start.timeZone -ne $event.timeZone -or
                [datetime]::Parse($actual.start.dateTime, [Globalization.CultureInfo]::InvariantCulture) -ne $expectedStart -or
                [datetime]::Parse($actual.end.dateTime, [Globalization.CultureInfo]::InvariantCulture) -ne $expectedStart.AddMinutes([int]$event.minutes)) {
                throw "LEF26 calendar baseline changed: $key. Review instead of overwriting."
            }
        } else {
            $start = [datetime]::ParseExact("$($event.date) $($event.start)", 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
            $end = $start.AddMinutes([int]$event.minutes)
            if ($event.timeZone -ne 'W. Europe Standard Time') { throw "LEF26 unexpected time zone: $($event.timeZone)" }
            $from = "$($event.date)T00:00:00+02:00"
            $to = "$($event.date)T23:59:59+02:00"
            $query = "/me/calendarView?startDateTime=$([uri]::EscapeDataString($from))&endDateTime=$([uri]::EscapeDataString($to))&`$select=id,subject,categories&`$top=100"
            $collisions = @(Get-ExecutiveEventPages $query | Where-Object { $_.subject -eq $subject -or $_.categories -contains $marker })
            if ($collisions.Count) { throw "LEF26 found an untracked calendar collision: $subject" }
            $attendees = @(foreach ($person in $event.attendees) {
                @{ type = 'required'; emailAddress = @{ name = $person; address = $Senders[$person].address } }
            })
            $actual = Invoke-ExecutiveEventGraph POST '/me/events' @{
                subject = $subject; body = @{ contentType = 'text'; content = (Get-ExecutiveEventText $event.body $Language) }
                start = @{ dateTime = $start.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = $event.timeZone }
                end = @{ dateTime = $end.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = $event.timeZone }
                attendees = $attendees; categories = @($marker)
                transactionId = (Get-ExecutiveEventTransaction "$($settings.tenantId)-$marker-$($event.date)")
            }
            $ledger.events = @($ledger.events) + @([pscustomobject]@{ key = $key; id = $actual.id; webUrl = $actual.webLink })
            Save-ExecutiveEventLedger $ledger $ledgerPath
        }
        $links += [pscustomobject]@{ name = $subject; url = $actual.webLink }
    }
    foreach ($thread in $Config.threads) {
        $marker = "LEF26-$Language-$($thread.key)"
        $entry = @($ledger.threads | Where-Object { $_.key -eq $thread.key })
        if ($entry.Count -gt 1) { throw "LEF26 duplicate Teams ledger entry: $($thread.key)" }
        $text = [Net.WebUtility]::HtmlEncode((Get-ExecutiveEventText $thread.text $Language))
        if ($entry.Count) {
            $record = $entry[0]
            $message = Invoke-ExecutiveEventGraph GET "$($Destination.channelPath)/messages/$([uri]::EscapeDataString($record.id))"
            if ($message.body.content -notlike "*$marker*") { throw 'LEF26 tracked Teams message marker changed.' }
        } else {
            $linkHtml = ''
            if ($thread.phase -eq 'BEFORE') {
                foreach ($file in $ledger.files) {
                    $linkHtml += '<p><a href="' + [Net.WebUtility]::HtmlEncode($file.webUrl) + '">' + [Net.WebUtility]::HtmlEncode($file.name) + '</a></p>'
                }
            }
            $message = Invoke-ExecutiveEventGraph POST "$($Destination.channelPath)/messages" @{
                body = @{ contentType = 'html'; content = "<p><strong>$marker</strong></p><p>$text</p>$linkHtml" }
            }
            $record = [pscustomobject]@{ key = $thread.key; id = $message.id; webUrl = $message.webUrl; replies = @() }
            $ledger.threads = @($ledger.threads) + @($record)
            Save-ExecutiveEventLedger $ledger $ledgerPath
        }
        for ($i = 0; $i -lt @($thread.replies).Count; $i++) {
            $savedReply = @($record.replies | Where-Object { $_.index -eq $i })
            if ($savedReply.Count -gt 1) { throw "LEF26 duplicate Teams reply ledger entry: $marker/$i" }
            if ($savedReply.Count) {
                $existingReply = Invoke-ExecutiveEventGraph GET "$($Destination.channelPath)/messages/$([uri]::EscapeDataString($record.id))/replies/$([uri]::EscapeDataString($savedReply[0].id))"
                if ($existingReply.body.content -notlike "*$marker-reply-$i*") { throw 'LEF26 tracked Teams reply changed.' }
                continue
            }
            $replyText = [Net.WebUtility]::HtmlEncode((Get-ExecutiveEventText $thread.replies[$i] $Language))
            $reply = Invoke-ExecutiveEventGraph POST "$($Destination.channelPath)/messages/$([uri]::EscapeDataString($record.id))/replies" @{
                body = @{ contentType = 'html'; content = "<p>$marker-reply-$i</p><p>$replyText</p>" }
            }
            $record.replies = @($record.replies) + @([pscustomobject]@{ index = $i; id = $reply.id })
            Save-ExecutiveEventLedger $ledger $ledgerPath
        }
        $links += [pscustomobject]@{ name = "LEF26 Teams $($thread.phase)"; url = $message.webUrl }
    }
    $reportPath = Join-Path (Split-Path $ledgerPath -Parent) "$Language-links.json"
    [IO.File]::WriteAllText($reportPath, (ConvertTo-Json -InputObject @($links) -Depth 5), (New-Object Text.UTF8Encoding $false))
    Write-Host "  LEF26 setup links: $reportPath"
    $links
}

function Assert-ExecutiveEventReply($Draft, [string]$Recipient, [string]$Account, [string]$Marker) {
    if (@($Draft.toRecipients).Count -ne 1 -or $Draft.toRecipients[0].emailAddress.address -ne $Recipient -or
        @($Draft.ccRecipients).Count -or @($Draft.bccRecipients).Count -or
        $Draft.from.emailAddress.address -ne $Account -or $Draft.body.content -notlike "*$Marker*") {
        throw 'LEF26 setup reply has unexpected recipients, sender or content marker. No send is allowed.'
    }
}

function Invoke-ExecutiveEventMailReplies {
    param($Config, $Settings, [string]$Language, [switch]$WhatIf)
    if ($WhatIf) {
        Write-Host "  LEF26 WhatIf: $(@($Config.mailReplies).Count) internal presenter replies via createReply/send, preserving real Outlook conversations. No writes."
        return
    }
    $ledger = Get-ExecutiveEventLedger $Settings $Language
    if (-not $ledger) { throw 'LEF26 shared/calendar setup ledger required before creating mail conversations.' }
    $path = Get-ExecutiveEventLedgerPath $Settings $Language
    if (-not $ledger.PSObject.Properties['mailReplies']) { $ledger | Add-Member NoteProperty mailReplies @() }
    foreach ($reply in $Config.mailReplies) {
        $recipient = $Config.people.($reply.parentSender)
        if (-not $recipient -or ($recipient -split '@')[-1] -ne $Config.domain) { throw 'LEF26 setup reply recipient is not approved.' }
        $subject = Get-ExecutiveEventText $reply.parentSubject $Language
        $filter = [uri]::EscapeDataString("subject eq '$($subject -replace "'", "''")'")
        $parents = @()
        for ($attempt = 0; $attempt -lt 15; $attempt++) {
            $parents = @(Get-ExecutiveEventPages "/me/mailFolders/inbox/messages?`$filter=$filter&`$select=id,subject,from,conversationId&`$top=25" |
                Where-Object { $_.from.emailAddress.address -eq $recipient })
            if ($parents.Count) { break }
            Start-Sleep -Seconds 2
        }
        if ($parents.Count -ne 1) { throw "LEF26 needs exactly one delivered original from $recipient with subject '$subject'; found $($parents.Count). No reply created." }
        $record = @($ledger.mailReplies | Where-Object { $_.key -eq $reply.key })
        if (-not $record.Count) {
            $pending = [pscustomobject]@{ key = $reply.key; parentId = $parents[0].id; id = ''; state = 'creating' }
            $ledger.mailReplies = @($ledger.mailReplies) + @($pending)
            Save-ExecutiveEventLedger $ledger $path
            $draft = Invoke-ExecutiveEventGraph POST "/me/messages/$([uri]::EscapeDataString($parents[0].id))/createReply" @{
                comment = "[LEF26-SETUP-$($reply.key)]`n`n$(Get-ExecutiveEventText $reply.text $Language)"
            }
            $pending.id = $draft.id
            $pending.state = 'draft'
            Save-ExecutiveEventLedger $ledger $path
            $record = @($pending)
        }
        if ($record.Count -ne 1 -or -not $record[0].id -or $record[0].parentId -cne $parents[0].id) {
            throw 'LEF26 mail reply ledger is ambiguous or interrupted. Inspect the exact pending draft; do not recreate blindly.'
        }
        $draft = Invoke-ExecutiveEventGraph GET "/me/messages/$([uri]::EscapeDataString($record[0].id))"
        Assert-ExecutiveEventReply $draft $recipient $Config.account "LEF26-SETUP-$($reply.key)"
        if ($draft.conversationId -cne $parents[0].conversationId) { throw 'LEF26 reply is not in the original conversation.' }
        if ($draft.isDraft) {
            if ($record[0].state -eq 'submitted') {
                throw 'LEF26 reply was already submitted but still appears as a draft. Wait and verify delivery; do not resend automatically.'
            }
            # Journal submission before the send: an interrupted/uncertain response must never cause a duplicate send.
            $record[0].state = 'submitted'
            Save-ExecutiveEventLedger $ledger $path
            Invoke-ExecutiveEventGraph POST "/me/messages/$([uri]::EscapeDataString($draft.id))/send" | Out-Null
            $sent = $null
            for ($attempt = 0; $attempt -lt 15; $attempt++) {
                $sent = Invoke-ExecutiveEventGraph GET "/me/messages/$([uri]::EscapeDataString($draft.id))" -Optional
                if ($sent -and -not $sent.isDraft) { break }
                Start-Sleep -Seconds 2
            }
            if (-not $sent -or $sent.isDraft) { throw 'LEF26 reply submission is still pending. Verify delivery before presenting; no automatic resend.' }
            Assert-ExecutiveEventReply $sent $recipient $Config.account "LEF26-SETUP-$($reply.key)"
            if ($sent.conversationId -cne $parents[0].conversationId) { throw 'LEF26 sent reply lost the original conversation.' }
            $record[0].state = 'sent'
            Save-ExecutiveEventLedger $ledger $path
        } else {
            $record[0].state = 'sent'
            Save-ExecutiveEventLedger $ledger $path
        }
        Write-Host "  LEF26 internal setup reply: $subject -> $recipient (live demos remain draft-only)."
    }
}

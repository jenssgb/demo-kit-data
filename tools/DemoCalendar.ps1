# Fixed-date calendar inputs and isolated run helpers. No writes on import.
function Assert-DemoPresenter($ExpectedAccount, $Context, $Me) {
    if (-not $ExpectedAccount -or -not $Context -or -not $Context.TenantId -or
        $Context.Account -ne $ExpectedAccount -or $Me.userPrincipalName -ne $ExpectedAccount -or
        $Me.mail -ne $ExpectedAccount) {
        throw 'Demo presenter blocked: cached Graph account and actual mailbox must match the approved presenterAccount.'
    }
}

function Invoke-DemoGraph {
    param([string]$Method, [string]$Path, $Body, [string]$InputFile)
    if ($Path -notmatch '^/me(?:/|\?)') { throw 'Demo Graph only permits /me resources.' }
    $p = @{ Method = $Method; Uri = "https://graph.microsoft.com/v1.0$Path"; OutputType = 'Hashtable' }
    if ($InputFile) {
        $p.InputFilePath = $InputFile
        $p.ContentType = 'application/octet-stream'
    } elseif ($null -ne $Body) {
        $p.Body = ConvertTo-Json -InputObject $Body -Depth 15 -Compress
        $p.ContentType = 'application/json; charset=utf-8'
    }
    Invoke-MgGraphRequest @p
}

function Get-DemoEventPlan($Event, [string]$Language) {
    if ($Event.fixedDate -notmatch '^\d{4}-\d{2}-\d{2}$' -or $Event.start -notmatch '^\d{2}:\d{2}$' -or
        $Event.scenarioKey -notmatch '^[A-Z][A-Z0-9]*-[a-z0-9-]+$' -or
        $Language -notin 'en', 'de' -or -not $Event.timeZone -or
        [int]$Event.minutes -le 0 -or [int]$Event.minutes -gt 1440 -or
        @($Event.attendees | Where-Object { $_ }).Count) {
        throw 'Fixed demo event requires an exact date, time zone, positive duration, scenario key and no attendees.'
    }
    $zone = [TimeZoneInfo]::FindSystemTimeZoneById($Event.timeZone)
    $start = [datetime]::ParseExact("$($Event.fixedDate) $($Event.start)", 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
    $end = $start.AddMinutes([int]$Event.minutes)
    if ($zone.IsInvalidTime($start) -or $zone.IsAmbiguousTime($start) -or
        $zone.IsInvalidTime($end) -or $zone.IsAmbiguousTime($end)) {
        throw 'Fixed demo event falls in an invalid or ambiguous daylight-saving time.'
    }
    $subject = if ($Event.subject -is [string]) { $Event.subject } else { $Event.subject.$Language }
    $body = if ($Event.body -is [string]) { $Event.body } else { $Event.body.$Language }
    if (-not $subject -or -not $body) { throw 'Fixed demo event is missing localized subject/body.' }
    $key = "$($Event.scenarioKey):$Language"
    $property = 'String {E65C7B13-BA4B-4199-994B-01B34E302D20} Name DemoKitFixedEvent'
    $hash = [Security.Cryptography.SHA256]::Create()
    try { $bytes = $hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($key)) }
    finally { $hash.Dispose() }
    $payload = @{
        subject = $subject; body = @{ contentType = 'text'; content = $body }
        start = @{ dateTime = $start.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = $Event.timeZone }
        end = @{ dateTime = $end.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = $Event.timeZone }
        location = @{ displayName = [string]$Event.location }
        transactionId = (New-Object Guid (,[byte[]]$bytes[0..15])).ToString()
        singleValueExtendedProperties = @(@{ id = $property; value = $key })
    }
    [pscustomobject]@{
        key = $key; property = $property; payload = $payload; start = $start
        startUtc = [TimeZoneInfo]::ConvertTimeToUtc($start, $zone)
        endUtc = [TimeZoneInfo]::ConvertTimeToUtc($end, $zone)
    }
}

function Convert-DemoEventUtc($Value) {
    if (-not $Value.dateTime -or -not $Value.timeZone) { throw 'Saved fixed event has no exact time/timeZone.' }
    $date = [datetime]::Parse($Value.dateTime, [Globalization.CultureInfo]::InvariantCulture)
    $date = [datetime]::SpecifyKind($date, [DateTimeKind]::Unspecified)
    [TimeZoneInfo]::ConvertTimeToUtc($date, [TimeZoneInfo]::FindSystemTimeZoneById($Value.timeZone))
}

function Sync-DemoFixedEvent {
    param($Event, [string]$Language, [scriptblock]$Request, [switch]$WhatIf)
    $plan = Get-DemoEventPlan $Event $Language
    # Stable property lookup finds the owned event even if someone moved/renamed it.
    $filter = "singleValueExtendedProperties/Any(ep: ep/id eq '$($plan.property)' and ep/value eq '$($plan.key)')"
    $expand = "singleValueExtendedProperties(`$filter=id eq '$($plan.property)')"
    $path = "/me/events?`$filter=$([uri]::EscapeDataString($filter))&`$expand=$([uri]::EscapeDataString($expand))&`$top=100"
    $found = @()
    do {
        $page = & $Request GET $path
        $found += @($page.value | Where-Object { $null -ne $_ })
        $path = $page.'@odata.nextLink'
        if ($path) {
            if ($path -notmatch '^https://graph\.microsoft\.com/v1\.0/me/events\?') { throw 'Unexpected fixed-event pagination URL.' }
            $path = $path -replace '^https://graph\.microsoft\.com/v1\.0', ''
        }
    } while ($path)
    if ($found.Count -gt 1) { throw "Ambiguous fixed event $($plan.key); no event was changed." }
    if ($found.Count) {
        $actual = $found[0]
        $identity = @($actual.singleValueExtendedProperties | Where-Object { $_.id -ceq $plan.property -and $_.value -ceq $plan.key })
        if ($identity.Count -ne 1 -or -not $actual.id -or
            $actual.subject -cne $plan.payload.subject -or $actual.isCancelled -or
            @($actual.attendees | Where-Object { $_ }).Count -or
            (Convert-DemoEventUtc $actual.start) -ne $plan.startUtc -or
            (Convert-DemoEventUtc $actual.end) -ne $plan.endUtc -or
            $actual.location.displayName -cne $plan.payload.location.displayName -or
            $actual.body.contentType -ine 'text' -or
            $actual.body.content.TrimEnd() -cne $plan.payload.body.content.TrimEnd()) {
            throw "Changed fixed event $($plan.key); baseline differs. Review the existing appointment; it was not overwritten."
        }
        return 'same'
    }
    if (-not $WhatIf) { [void](& $Request POST '/me/events' $plan.payload) }
    'new'
}

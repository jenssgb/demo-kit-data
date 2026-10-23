$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'tenant.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($name in 'Decode-Words', 'Read-Eml', 'New-DemoSeedMessage') {
    $definition = $ast.Find({ param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    if (-not $definition) { throw "Missing function $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}
function Get-Resource($item, $path) {
    $manifest = Get-Content -LiteralPath (Join-Path $item.dir 'manifest.json') -Raw | ConvertFrom-Json
    $entry = @($manifest.shared) | Where-Object { $_.path -eq $path } | Select-Object -First 1
    $file = if ($entry) { Join-Path $root "$($entry.from)\$($entry.path)" } else { Join-Path $item.dir $path }
    [IO.File]::ReadAllBytes($file)
}
function Assert($condition, $message) { if (-not $condition) { throw $message } }
$item = [pscustomobject]@{ id = 'chat-shipping-workiq'; dir = (Join-Path $root 'chat-shipping-workiq') }
$cfg = Get-Content -LiteralPath (Join-Path $item.dir 'tenant.json') -Raw | ConvertFrom-Json
$m = Read-Eml ([IO.File]::ReadAllBytes((Join-Path $item.dir 'tenant\10_Teresa_Attachments.eml')))
$sender = [pscustomobject]@{ address = 'TeresaS@M365CPI98544940.OnMicrosoft.com' }
$recipient = [pscustomobject]@{ mail = 'admin@M365CPI98544940.OnMicrosoft.com'; displayName = 'MOD Administrator' }
$legacy = New-DemoSeedMessage $m $sender $recipient $item $null
Assert ($legacy.message.body.contentType -eq 'Text') 'Legacy mail must stay plain text.'
Assert (-not $legacy.message.ContainsKey('attachments')) 'Legacy mail acquired attachments.'
Assert ($legacy.message.body.content -ceq $m.Body) 'Legacy mail body changed.'
$payload = New-DemoSeedMessage $m $sender $recipient $item $cfg.mailOptions[0]
Assert ($payload.message.attachments.Count -eq 2) 'Expected two actual file attachments.'
foreach ($attachment in $payload.message.attachments) {
    Assert ($attachment.'@odata.type' -eq '#microsoft.graph.fileAttachment') 'Wrong attachment type.'
    $expected = Get-Resource $item $attachment.name
    Assert ($attachment.contentBytes -ceq [Convert]::ToBase64String($expected)) 'Attachment bytes changed.'
}
$payload | ConvertTo-Json -Depth 20 | ConvertFrom-Json | Out-Null
$htmlMail = [pscustomobject]@{ Subject = 'Test'; FromName = 'Teresa Sac'; Body = "Evidence <not markup> & text`nSecond line" }
$html = New-DemoSeedMessage $htmlMail $sender $recipient $item ([pscustomobject]@{ html = $true })
Assert ($html.message.body.contentType -eq 'HTML') 'HTML option not applied.'
Assert ($html.message.body.content -match '&lt;not markup&gt; &amp; text<br>Second line') 'HTML escaping failed.'
foreach ($path in '..\secret.txt', 'C:\secret.txt', '/secret.txt', 'not-in-manifest.pdf', '', $null, 0) {
    $rejected = $false
    try { New-DemoSeedMessage $m $sender $recipient $item ([pscustomobject]@{ attachments = @($path) }) | Out-Null }
    catch { $rejected = $true }
    Assert $rejected "Unsafe or unlisted attachment accepted: $path"
}
foreach ($options in @([pscustomobject]@{ html = 'false' }, [pscustomobject]@{ attachments = 'file.pdf' })) {
    $rejected = $false
    try { New-DemoSeedMessage $m $sender $recipient $item $options | Out-Null }
    catch { $rejected = $true }
    Assert $rejected 'Invalid mailOptions type accepted.'
}
$originalResource = (Get-Item function:Get-Resource).ScriptBlock
try {
    function Get-Resource($Item, $Path) { ,([byte[]]::new($script:testAttachmentSize)) }
    $script:testAttachmentSize = 2MB
    $boundary = New-DemoSeedMessage $m $sender $recipient $item ([pscustomobject]@{ attachments = @('Revised_ETA_ContosoStar.pdf') })
    Assert ($boundary.message.attachments.Count -eq 1) 'Exactly 2 MB must be accepted.'
    $script:testAttachmentSize = 2MB + 1
    $rejected = $false
    try { New-DemoSeedMessage $m $sender $recipient $item ([pscustomobject]@{ attachments = @('Revised_ETA_ContosoStar.pdf') }) | Out-Null }
    catch { $rejected = $_ -match '2 MB combined' }
    Assert $rejected 'A payload over 2 MB must be rejected.'
} finally {
    Set-Item function:Get-Resource $originalResource
    Remove-Variable testAttachmentSize -Scope Script
}
Write-Host 'PASS: legacy text, escaped HTML, exact PDF/XLSX bytes, JSON, invalid inputs and the exact 2 MB boundary.'

<#
.SYNOPSIS
Exports, previews, or applies Home Assistant entity names, entity aliases, and device names from a CSV.

.DESCRIPTION
Connects to the Home Assistant WebSocket API.

-Export     Writes a CSV with every entity and device. NewName and Aliases are pre-filled with
            the current values, so you edit them in place. Save the file as aliases.csv.
            DisplayedName, Area, and Device are for reference only and are ignored.
(default)   Dry run. Validates aliases.csv and shows the planned changes. Nothing is written.
-Apply      Writes a rollback CSV to Backups\, then applies the planned changes.

CSV columns: Type,Target,NewName,Aliases (other columns are ignored)
  Type      entity or device
  Target    entity_id for entities. Device id or device name for devices.
  NewName   Unchanged if it matches the current value or is empty. <clear> = remove the custom name (back to the default).
  Aliases   Entity only. Semicolon-separated. Unchanged if it matches the current value or is empty.
            <name> = the entity's friendly name. <clear> = remove all aliases.
            Home Assistant has no aliases for devices, so device Aliases are ignored with a warning.

Entity IDs do not change on rename, so automations keep working.
Changes apply live. No restart is needed.

.EXAMPLE
.\Set_HAAliases.ps1 -Export
.\Set_HAAliases.ps1
.\Set_HAAliases.ps1 -Apply
#>
[CmdletBinding()]
param(
    [string]$BaseUrl = 'http://homeassistant.local:8123',
    [string]$AccessToken,
    [string]$MappingFile = (Join-Path $PSScriptRoot 'aliases.csv'),
    [string]$ExportFile,
    [switch]$Export,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
$NameToken = '<name>'
$ClearToken = '<clear>'
$RequiredColumns = @('Type', 'Target', 'NewName', 'Aliases')
$script:MessageId = 1

function Get-PlainToken {
    if ($AccessToken) { return $AccessToken }
    if ($env:HA_ACCESS_TOKEN) { return $env:HA_ACCESS_TOKEN }
    $secure = Read-Host 'Home Assistant long-lived access token' -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

function Send-HAMessage {
    param($Socket, $Message)
    $json = ConvertTo-Json -InputObject $Message -Depth 6 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $segment = [ArraySegment[byte]]::new($bytes)
    $Socket.SendAsync($segment, [Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
}

function Receive-HAMessage {
    param($Socket)
    $buffer = New-Object byte[] 65536
    $stream = New-Object System.IO.MemoryStream
    try {
        do {
            $segment = [ArraySegment[byte]]::new($buffer, 0, $buffer.Length)
            $result = $Socket.ReceiveAsync($segment, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
            if ($result.MessageType -eq [Net.WebSockets.WebSocketMessageType]::Close) {
                throw 'Home Assistant closed the WebSocket connection.'
            }
            $stream.Write($buffer, 0, $result.Count)
        } while (-not $result.EndOfMessage)
        return ([Text.Encoding]::UTF8.GetString($stream.ToArray()) | ConvertFrom-Json)
    }
    finally { $stream.Dispose() }
}

function Connect-HAWebSocket {
    param([string]$Url, [string]$Token)
    $base = [Uri]$Url
    $builder = New-Object System.UriBuilder
    $builder.Scheme = if ($base.Scheme -eq 'https') { 'wss' } else { 'ws' }
    $builder.Host = $base.Host
    $builder.Port = $base.Port
    $builder.Path = $base.AbsolutePath.TrimEnd('/') + '/api/websocket'

    $socket = New-Object System.Net.WebSockets.ClientWebSocket
    $timeout = New-Object System.Threading.CancellationTokenSource 15000
    try {
        $socket.ConnectAsync($builder.Uri, $timeout.Token).GetAwaiter().GetResult()
        $hello = Receive-HAMessage $socket
        if ($hello.type -ne 'auth_required') { throw "Unexpected first message from Home Assistant: $($hello.type)" }
        Send-HAMessage $socket @{ type = 'auth'; access_token = $Token }
        $auth = Receive-HAMessage $socket
        if ($auth.type -eq 'auth_invalid') { throw 'Home Assistant rejected the access token.' }
        if ($auth.type -ne 'auth_ok') { throw "Unexpected authentication response: $($auth.type)" }
    }
    catch {
        $socket.Dispose()
        throw "Could not connect to Home Assistant at $($builder.Uri): $($_.Exception.Message)"
    }
    return $socket
}

function Close-HAWebSocket {
    param($Socket)
    if ($Socket.State -eq [Net.WebSockets.WebSocketState]::Open) {
        $Socket.CloseAsync([Net.WebSockets.WebSocketCloseStatus]::NormalClosure, 'done', [Threading.CancellationToken]::None).GetAwaiter().GetResult()
    }
    $Socket.Dispose()
}

function Invoke-HACommand {
    param($Socket, [string]$Type, [hashtable]$Body = @{})
    $id = $script:MessageId
    $script:MessageId++
    $payload = @{ id = $id; type = $Type }
    foreach ($key in $Body.Keys) { $payload[$key] = $Body[$key] }
    Send-HAMessage $Socket $payload
    while ($true) {
        $msg = Receive-HAMessage $Socket
        if ($msg.type -eq 'result' -and $msg.id -eq $id) {
            if (-not $msg.success) { throw "Home Assistant rejected '$Type': $($msg.error.message)" }
            return $msg.result
        }
    }
}

function ConvertTo-AliasList {
    param([string]$Text)
    if ($Text -eq $ClearToken) {
        $empty = [object[]]@()
        return , $empty
    }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($item in ($Text -split ';')) {
        $value = $item.Trim()
        if (-not $value) { continue }
        if ($value -eq $NameToken) { $list.Add($null) } else { $list.Add($value) }
    }
    return , $list.ToArray()
}

function Format-AliasList {
    param($Aliases)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($alias in @($Aliases)) {
        if ($null -eq $alias) { $parts.Add($NameToken) } else { $parts.Add([string]$alias) }
    }
    return ($parts -join ';')
}

function Show-Value {
    param([string]$Value)
    if ($Value) { return $Value }
    return '(default)'
}

function Get-DeviceDisplayName {
    param($Device)
    if ($Device.name_by_user) { return [string]$Device.name_by_user }
    return [string]$Device.name
}

function Read-Mapping {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Mapping file not found: $Path" }
    $headerLine = Get-Content -LiteralPath $Path -TotalCount 1
    if (-not $headerLine) { throw "Mapping file is empty: $Path" }
    $headers = @($headerLine -split ',' | ForEach-Object { $_.Trim().Trim('"') })
    $missing = @($RequiredColumns | Where-Object { $headers -notcontains $_ })
    if ($missing.Count -gt 0) { throw "Mapping file is missing column(s): $($missing -join ', ')" }

    $problems = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $clean = New-Object System.Collections.Generic.List[object]
    $line = 1
    foreach ($row in @(Import-Csv -LiteralPath $Path)) {
        $line++
        $type = ([string]$row.Type).Trim().ToLowerInvariant()
        $target = ([string]$row.Target).Trim()
        $newName = ([string]$row.NewName).Trim()
        $aliasText = ([string]$row.Aliases).Trim()
        if (-not ($type -or $target -or $newName -or $aliasText)) { continue }

        if ($type -notin @('entity', 'device')) { $problems.Add("Line $line : Type must be 'entity' or 'device' (got '$($row.Type)')."); continue }
        if (-not $target) { $problems.Add("Line $line : Target is empty."); continue }
        if ($type -eq 'entity' -and $target -notmatch '^[a-z_]+\.[a-z0-9_]+$') {
            $problems.Add("Line $line : '$target' is not a valid entity_id (domain.object)."); continue
        }
        $key = "$type|$($target.ToLowerInvariant())"
        if ($seen.ContainsKey($key)) { $problems.Add("Line $line : duplicate row for $type '$target' (first at line $($seen[$key]))."); continue }
        $seen[$key] = $line

        $clean.Add([pscustomobject]@{
            Line = $line; Type = $type; Target = $target; NewName = $newName; Aliases = $aliasText
        })
    }
    if ($problems.Count -gt 0) { throw ("Mapping file has errors:`n  " + ($problems -join "`n  ")) }
    return , $clean.ToArray()
}

if ($Export -and $Apply) { throw '-Export and -Apply cannot be used together.' }

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$rows = @()
if (-not $Export) { $rows = Read-Mapping -Path $MappingFile }

$socket = Connect-HAWebSocket -Url $BaseUrl -Token (Get-PlainToken)
try {
    $entities = @(Invoke-HACommand $socket 'config/entity_registry/list')
    $devices = @(Invoke-HACommand $socket 'config/device_registry/list')
    $areas = @(Invoke-HACommand $socket 'config/area_registry/list')

    $entityById = @{}
    foreach ($entity in $entities) { $entityById[$entity.entity_id] = $entity }
    $deviceById = @{}
    foreach ($device in $devices) { $deviceById[$device.id] = $device }
    $areaNameById = @{}
    foreach ($area in $areas) { $areaNameById[$area.area_id] = $area.name }

    if ($Export) {
        $outPath = $ExportFile
        if (-not $outPath) { $outPath = Join-Path $PSScriptRoot "aliases_export_$stamp.csv" }

        $friendlyById = @{}
        foreach ($state in @(Invoke-HACommand $socket 'get_states')) {
            $friendlyById[$state.entity_id] = [string]$state.attributes.friendly_name
        }

        $exportRows = New-Object System.Collections.Generic.List[object]
        $index = 0
        foreach ($entity in ($entities | Sort-Object entity_id)) {
            $index++
            Write-Progress -Activity 'Reading entity names and aliases' -Status $entity.entity_id -PercentComplete ([int](100 * $index / $entities.Count))
            $full = Invoke-HACommand $socket 'config/entity_registry/get' @{ entity_id = $entity.entity_id }
            $device = $null
            if ($entity.device_id) { $device = $deviceById[$entity.device_id] }
            $areaId = $entity.area_id
            if (-not $areaId -and $device) { $areaId = $device.area_id }
            $areaName = ''
            if ($areaId) { $areaName = [string]$areaNameById[$areaId] }
            $deviceName = ''
            if ($device) { $deviceName = Get-DeviceDisplayName $device }
            $shownName = $friendlyById[$entity.entity_id]
            if (-not $shownName) { $shownName = [string]$full.original_name }

            $exportRows.Add([pscustomobject]@{
                Type = 'entity'
                Target = $entity.entity_id
                NewName = [string]$full.name
                Aliases = (Format-AliasList $full.aliases)
                DisplayedName = $shownName
                Area = $areaName
                Device = $deviceName
            })
        }
        Write-Progress -Activity 'Reading entity names and aliases' -Completed

        foreach ($device in ($devices | Sort-Object name)) {
            $areaName = ''
            if ($device.area_id) { $areaName = [string]$areaNameById[$device.area_id] }
            $exportRows.Add([pscustomobject]@{
                Type = 'device'
                Target = $device.id
                NewName = [string]$device.name_by_user
                Aliases = ''
                DisplayedName = (Get-DeviceDisplayName $device)
                Area = $areaName
                Device = ''
            })
        }

        $exportRows | Export-Csv -LiteralPath $outPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exported $($exportRows.Count) rows to $outPath" -ForegroundColor Green
        Write-Host 'Next: save it as aliases.csv, edit NewName and Aliases, delete rows you do not want to check, then run without -Apply to preview.'
    }
    else {
        $plan = New-Object System.Collections.Generic.List[object]
        $problems = New-Object System.Collections.Generic.List[string]
        $warnings = New-Object System.Collections.Generic.List[string]

        foreach ($row in $rows) {
            try {
                if ($row.Type -eq 'entity') {
                    if (-not $entityById.ContainsKey($row.Target)) { throw "entity not found: $($row.Target)" }
                    $entry = Invoke-HACommand $socket 'config/entity_registry/get' @{ entity_id = $row.Target }
                    if ($entry.PSObject.Properties.Name -notcontains 'aliases') {
                        throw 'this Home Assistant version does not report entity aliases. Update Home Assistant and retry.'
                    }

                    $body = @{ entity_id = $row.Target }
                    $details = New-Object System.Collections.Generic.List[string]
                    $undo = @{ NewName = ''; Aliases = '' }

                    if ($row.NewName) {
                        $wantName = $null
                        if ($row.NewName -ne $ClearToken) { $wantName = $row.NewName }
                        $haveName = [string]$entry.name
                        if (([string]$wantName) -cne $haveName) {
                            $body['name'] = $wantName
                            $details.Add("name '$(Show-Value $haveName)' -> '$(Show-Value ([string]$wantName))'")
                            $undo.NewName = $ClearToken
                            if ($haveName) { $undo.NewName = $haveName }
                        }
                    }

                    if ($row.Aliases) {
                        $wantAliases = ConvertTo-AliasList $row.Aliases
                        $haveAliases = @($entry.aliases)
                        $wantText = Format-AliasList $wantAliases
                        $haveText = Format-AliasList $haveAliases
                        if ($wantText -cne $haveText) {
                            $body['aliases'] = $wantAliases
                            $details.Add("aliases '$haveText' -> '$wantText'")
                            $undo.Aliases = $haveText
                            if ($haveAliases.Count -eq 0) { $undo.Aliases = $ClearToken }
                        }
                    }

                    if ($body.Count -gt 1) {
                        $plan.Add([pscustomobject]@{
                            Type = 'entity'
                            Target = $row.Target
                            Method = 'config/entity_registry/update'
                            Body = $body
                            Details = ($details -join '; ')
                            Undo = $undo
                        })
                    }
                }
                else {
                    $device = $null
                    $byId = @($devices | Where-Object { $_.id -eq $row.Target })
                    if ($byId.Count -eq 1) {
                        $device = $byId[0]
                    }
                    else {
                        $byName = @($devices | Where-Object { $_.name_by_user -eq $row.Target -or $_.name -eq $row.Target })
                        if ($byName.Count -eq 0) { throw "device not found: $($row.Target)" }
                        if ($byName.Count -gt 1) { throw "device name '$($row.Target)' matches $($byName.Count) devices. Use the device id instead." }
                        $device = $byName[0]
                    }

                    if ($row.Aliases) {
                        $warnings.Add("Line $($row.Line): Home Assistant has no aliases for devices. Aliases ignored for device '$($device.id)'.")
                    }

                    if ($row.NewName) {
                        $wantName = $null
                        if ($row.NewName -ne $ClearToken) { $wantName = $row.NewName }
                        $haveName = [string]$device.name_by_user
                        if (([string]$wantName) -cne $haveName) {
                            $undoName = $ClearToken
                            if ($haveName) { $undoName = $haveName }
                            $plan.Add([pscustomobject]@{
                                Type = 'device'
                                Target = $device.id
                                Method = 'config/device_registry/update'
                                Body = @{ device_id = $device.id; name_by_user = $wantName }
                                Details = "name '$(Show-Value $haveName)' -> '$(Show-Value ([string]$wantName))'"
                                Undo = @{ NewName = $undoName; Aliases = '' }
                            })
                        }
                    }
                }
            }
            catch {
                $problems.Add("Line $($row.Line) ($($row.Target)): $($_.Exception.Message)")
            }
        }

        if ($problems.Count -gt 0) { throw ("Cannot continue. Fix these rows:`n  " + ($problems -join "`n  ")) }
        foreach ($warning in $warnings) { Write-Warning $warning }

        if ($plan.Count -eq 0) {
            Write-Host 'Nothing to change. Home Assistant already matches the mapping file.' -ForegroundColor Green
        }
        else {
            Write-Host "Planned changes ($($plan.Count)):"
            foreach ($item in $plan) {
                Write-Host ("  {0,-7} {1,-45} {2}" -f $item.Type, $item.Target, $item.Details)
            }

            if (-not $Apply) {
                Write-Host 'Dry run only. Nothing was changed. Re-run with -Apply to write these changes.' -ForegroundColor Yellow
            }
            else {
                $backupDir = Join-Path $PSScriptRoot 'Backups'
                New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
                $undoPath = Join-Path $backupDir "aliases_rollback_$stamp.csv"
                $plan | ForEach-Object {
                    [pscustomobject]@{ Type = $_.Type; Target = $_.Target; NewName = $_.Undo.NewName; Aliases = $_.Undo.Aliases }
                } | Export-Csv -LiteralPath $undoPath -NoTypeInformation -Encoding UTF8
                Write-Host "Rollback file written: $undoPath"

                $done = 0
                foreach ($item in $plan) {
                    try {
                        Invoke-HACommand $socket $item.Method $item.Body | Out-Null
                        $done++
                    }
                    catch {
                        throw "Stopped at $($item.Type) '$($item.Target)' after $done of $($plan.Count) change(s): $($_.Exception.Message). To undo the changes already made, run: .\Set_HAAliases.ps1 -MappingFile `"$undoPath`" -Apply"
                    }
                }
                Write-Host "Applied $done change(s)." -ForegroundColor Green
            }
        }
    }
}
finally {
    Close-HAWebSocket $socket
}

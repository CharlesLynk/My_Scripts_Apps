param(
    [Parameter(Mandatory = $false)]
    [string]$BaseUrl = "http://homeassistant.local:8123",

    [Parameter(Mandatory = $false)]
    [string]$AccessToken,

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot = ".\ha_inventory_exports"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ConvertTo-PlainTextFromSecureString {
    param(
        [Parameter(Mandatory = $true)]
        [Security.SecureString]$SecureString
    )

    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}

function Receive-WebSocketJson {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.WebSockets.ClientWebSocket]$Socket
    )

    $buffer = New-Object byte[] 8192
    $segment = [System.ArraySegment[byte]]::new($buffer)
    $stream = New-Object System.IO.MemoryStream

    do {
        $result = $Socket.ReceiveAsync($segment, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
        if ($result.Count -gt 0) {
            $stream.Write($buffer, 0, $result.Count)
        }
    } while (-not $result.EndOfMessage)

    if ($result.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
        throw "WebSocket closed by server."
    }

    $stream.Position = 0
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    $json = $reader.ReadToEnd()
    $reader.Dispose()
    $stream.Dispose()

    if ([string]::IsNullOrWhiteSpace($json)) {
        throw "Received empty WebSocket payload."
    }

    return $json | ConvertFrom-Json
}

function Send-WebSocketJson {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.WebSockets.ClientWebSocket]$Socket,

        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $json = $Payload | ConvertTo-Json -Depth 20 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $segment = [System.ArraySegment[byte]]::new($bytes)
    $Socket.SendAsync(
        $segment,
        [System.Net.WebSockets.WebSocketMessageType]::Text,
        $true,
        [Threading.CancellationToken]::None
    ).GetAwaiter().GetResult()
}

function Invoke-HAWebSocketBundle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseUrl,

        [Parameter(Mandatory = $true)]
        [string]$AccessToken
    )

    $wsUrl = $BaseUrl.TrimEnd("/")
    if ($wsUrl.StartsWith("https://")) {
        $wsUrl = "wss://" + $wsUrl.Substring(8)
    }
    elseif ($wsUrl.StartsWith("http://")) {
        $wsUrl = "ws://" + $wsUrl.Substring(7)
    }
    else {
        throw "BaseUrl must begin with http:// or https://"
    }
    $wsUrl = "$wsUrl/api/websocket"

    $socket = [System.Net.WebSockets.ClientWebSocket]::new()

    try {
        $socket.ConnectAsync([Uri]$wsUrl, [Threading.CancellationToken]::None).GetAwaiter().GetResult()

        $hello = Receive-WebSocketJson -Socket $socket
        if ($hello.type -ne "auth_required") {
            throw "Unexpected handshake response: $($hello | ConvertTo-Json -Depth 10 -Compress)"
        }

        Send-WebSocketJson -Socket $socket -Payload @{
            type = "auth"
            access_token = $AccessToken
        }

        $authReply = Receive-WebSocketJson -Socket $socket
        if ($authReply.type -ne "auth_ok") {
            throw "Authentication failed: $($authReply | ConvertTo-Json -Depth 10 -Compress)"
        }

        $requestTypes = @(
            "config/entity_registry/list",
            "config/device_registry/list",
            "config/area_registry/list"
        )

        $results = @{}
        $reqId = 1

        foreach ($reqType in $requestTypes) {
            Send-WebSocketJson -Socket $socket -Payload @{
                id = $reqId
                type = $reqType
            }

            while ($true) {
                $message = Receive-WebSocketJson -Socket $socket

                if ($message.type -eq "result" -and $message.id -eq $reqId) {
                    if (-not $message.success) {
                        throw ("Request failed for {0}: {1}" -f $reqType, ($message | ConvertTo-Json -Depth 20 -Compress))
                    }
                    $results[$reqType] = $message.result
                    break
                }
            }

            $reqId += 1
        }

        return [pscustomobject]@{
            EntityRegistry = $results["config/entity_registry/list"]
            DeviceRegistry = $results["config/device_registry/list"]
            AreaRegistry = $results["config/area_registry/list"]
        }
    }
    finally {
        if ($socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
            $socket.CloseAsync(
                [System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure,
                "done",
                [Threading.CancellationToken]::None
            ).GetAwaiter().GetResult()
        }
        $socket.Dispose()
    }
}

function Invoke-HARestGet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,

        [Parameter(Mandatory = $true)]
        [string]$AccessToken
    )

    return Invoke-RestMethod -Method Get -Uri $Url -Headers @{ Authorization = "Bearer $AccessToken" }
}

function Get-SafeCount {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    if ($null -eq $Value) {
        return 0
    }

    if ($Value -is [System.Array]) {
        return $Value.Length
    }

    if ($Value -is [System.Collections.ICollection]) {
        return $Value.Count
    }

    return 1
}

if ([string]::IsNullOrWhiteSpace($AccessToken)) {
    Write-Host "Access token not provided as a parameter."
    $secureToken = Read-Host -Prompt "Paste Home Assistant long-lived access token" -AsSecureString
    $AccessToken = ConvertTo-PlainTextFromSecureString -SecureString $secureToken
}

if ([string]::IsNullOrWhiteSpace($AccessToken)) {
    throw "No access token provided."
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputDir = Join-Path $OutputRoot "ha_inventory_$timestamp"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$base = $BaseUrl.TrimEnd("/")

Write-Host "Collecting Home Assistant registry data via WebSocket..."
$wsData = Invoke-HAWebSocketBundle -BaseUrl $base -AccessToken $AccessToken

if ($wsData -is [System.Array]) {
    $wsData = $wsData | Where-Object {
        $_ -and $_.PSObject -and $_.PSObject.Properties["EntityRegistry"]
    } | Select-Object -Last 1
}

if (-not $wsData -or -not $wsData.PSObject.Properties["EntityRegistry"]) {
    throw "WebSocket registry query did not return the expected structure."
}

Write-Host "Collecting Home Assistant REST data..."
$states = Invoke-HARestGet -Url "$base/api/states" -AccessToken $AccessToken
$services = Invoke-HARestGet -Url "$base/api/services" -AccessToken $AccessToken
$config = Invoke-HARestGet -Url "$base/api/config" -AccessToken $AccessToken

$entityRegistry = @($wsData.EntityRegistry)
$deviceRegistry = @($wsData.DeviceRegistry)
$areaRegistry = @($wsData.AreaRegistry)
$statesArray = @($states)
$servicesArray = @($services)

$entityRegistry | ConvertTo-Json -Depth 50 | Set-Content -Path (Join-Path $outputDir "entity_registry.json") -Encoding UTF8
$deviceRegistry | ConvertTo-Json -Depth 50 | Set-Content -Path (Join-Path $outputDir "device_registry.json") -Encoding UTF8
$areaRegistry | ConvertTo-Json -Depth 50 | Set-Content -Path (Join-Path $outputDir "area_registry.json") -Encoding UTF8
$statesArray | ConvertTo-Json -Depth 50 | Set-Content -Path (Join-Path $outputDir "states.json") -Encoding UTF8
$servicesArray | ConvertTo-Json -Depth 50 | Set-Content -Path (Join-Path $outputDir "services.json") -Encoding UTF8
$config | ConvertTo-Json -Depth 50 | Set-Content -Path (Join-Path $outputDir "config.json") -Encoding UTF8

$helperDomains = @(
    "input_boolean",
    "input_button",
    "input_datetime",
    "input_number",
    "input_select",
    "input_text",
    "counter",
    "timer",
    "schedule"
)

$helpers = @($statesArray | Where-Object {
    $domain = ($_.entity_id -split "\\.", 2)[0]
    $helperDomains -contains $domain
})
$helpers | ConvertTo-Json -Depth 50 | Set-Content -Path (Join-Path $outputDir "helpers.json") -Encoding UTF8

$deviceById = @{}
foreach ($d in $deviceRegistry) {
    $deviceById[$d.id] = $d
}

$areaById = @{}
foreach ($a in $areaRegistry) {
    $areaById[$a.area_id] = $a
}

$stateByEntityId = @{}
foreach ($s in $statesArray) {
    $stateByEntityId[$s.entity_id] = $s
}

$entityReport = foreach ($e in $entityRegistry) {
    $entityId = $e.entity_id
    $domain = ($entityId -split "\\.", 2)[0]
    $stateObj = $stateByEntityId[$entityId]
    $device = $null
    if ($e.device_id -and $deviceById.ContainsKey($e.device_id)) {
        $device = $deviceById[$e.device_id]
    }

    $entityAreaId = $e.area_id
    if (-not $entityAreaId -and $device -and $device.area_id) {
        $entityAreaId = $device.area_id
    }

    $areaName = $null
    if ($entityAreaId -and $areaById.ContainsKey($entityAreaId)) {
        $areaName = $areaById[$entityAreaId].name
    }

    [pscustomobject]@{
        entity_id = $entityId
        domain = $domain
        original_name = $e.original_name
        friendly_name = if ($stateObj) { $stateObj.attributes.friendly_name } else { $null }
        state = if ($stateObj) { $stateObj.state } else { $null }
        area_id = $entityAreaId
        area_name = $areaName
        device_id = $e.device_id
        device_name = if ($device) {
            if (-not [string]::IsNullOrWhiteSpace($device.name_by_user)) {
                $device.name_by_user
            }
            else {
                $device.name
            }
        } else { $null }
        integration = $e.platform
        disabled_by = $e.disabled_by
        hidden_by = $e.hidden_by
    }
}

$deviceReport = foreach ($d in $deviceRegistry) {
    $areaName = $null
    if ($d.area_id -and $areaById.ContainsKey($d.area_id)) {
        $areaName = $areaById[$d.area_id].name
    }

    [pscustomobject]@{
        device_id = $d.id
        name = if (-not [string]::IsNullOrWhiteSpace($d.name_by_user)) {
            $d.name_by_user
        }
        else {
            $d.name
        }
        area_id = $d.area_id
        area_name = $areaName
        manufacturer = $d.manufacturer
        model = $d.model
        sw_version = $d.sw_version
        via_device_id = $d.via_device_id
        disabled_by = $d.disabled_by
    }
}

$entityReport | Export-Csv -Path (Join-Path $outputDir "entities.csv") -NoTypeInformation -Encoding UTF8
$deviceReport | Export-Csv -Path (Join-Path $outputDir "devices.csv") -NoTypeInformation -Encoding UTF8
$areaRegistry | Select-Object area_id, name, aliases | Export-Csv -Path (Join-Path $outputDir "areas.csv") -NoTypeInformation -Encoding UTF8

$summary = [ordered]@{
    exported_at_utc = (Get-Date).ToUniversalTime().ToString("o")
    base_url = $base
    counts = [ordered]@{
        entities_registry = Get-SafeCount -Value $entityRegistry
        devices_registry = Get-SafeCount -Value $deviceRegistry
        areas_registry = Get-SafeCount -Value $areaRegistry
        states = Get-SafeCount -Value $statesArray
        helpers = Get-SafeCount -Value $helpers
        service_domains = Get-SafeCount -Value $servicesArray
    }
    files = @(
        "entity_registry.json",
        "device_registry.json",
        "area_registry.json",
        "states.json",
        "helpers.json",
        "services.json",
        "config.json",
        "entities.csv",
        "devices.csv",
        "areas.csv"
    )
}

$summary | ConvertTo-Json -Depth 20 | Set-Content -Path (Join-Path $outputDir "summary.json") -Encoding UTF8

$zipPath = Join-Path $OutputRoot "ha_inventory_$timestamp.zip"
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}
Compress-Archive -Path (Join-Path $outputDir "*") -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "Export complete."
Write-Host "Folder: $outputDir"
Write-Host "Zip:    $zipPath"
Write-Host ""
Write-Host "Share the zip (or folder) to provide a complete environment inventory."

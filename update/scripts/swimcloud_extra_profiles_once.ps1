$ErrorActionPreference = 'Stop'

function Write-Log([string]$Message) {
    $stamp = Get-Date -Format 'HH:mm:ss'
    $line = "[$stamp] $Message"
    Write-Host $line
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
}

function Get-BrowserPath {
    $candidates = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    )
    foreach ($p in $candidates) { if ($p -and (Test-Path -LiteralPath $p)) { return $p } }
    return $null
}

function Get-FreePort {
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $listener.Start(); $port = ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port; $listener.Stop(); return $port
}

function Connect-Cdp([string]$WsUrl) {
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $ws.Options.KeepAliveInterval = [TimeSpan]::FromSeconds(20)
    $ws.ConnectAsync([Uri]$WsUrl, [System.Threading.CancellationToken]::None).Wait()
    return $ws
}

function Receive-CdpMessage($Ws) {
    $ms = New-Object System.IO.MemoryStream
    try {
        do {
            $buffer = New-Object byte[] 65536
            $segment = New-Object System.ArraySegment[byte] -ArgumentList (, $buffer)
            $task = $Ws.ReceiveAsync($segment, [System.Threading.CancellationToken]::None)
            $task.Wait(); $res = $task.Result
            if ($res.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) { throw 'DevTools WebSocket closed unexpectedly.' }
            $ms.Write($buffer, 0, $res.Count)
        } while (-not $res.EndOfMessage)
        return [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
    } finally { $ms.Dispose() }
}

$script:CdpId = 0
function Invoke-Cdp($Ws, [string]$Method, $Params = $null) {
    $script:CdpId++; $id = $script:CdpId
    $msg = [ordered]@{ id=$id; method=$Method }; if ($null -ne $Params) { $msg.params=$Params }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes(($msg | ConvertTo-Json -Depth 60 -Compress))
    $segment = New-Object System.ArraySegment[byte] -ArgumentList (, $bytes)
    $Ws.SendAsync($segment,[System.Net.WebSockets.WebSocketMessageType]::Text,$true,[System.Threading.CancellationToken]::None).Wait()
    while ($true) {
        $raw = Receive-CdpMessage $Ws
        try { $obj = $raw | ConvertFrom-Json } catch { continue }
        if ($null -ne $obj.id -and [int]$obj.id -eq $id) { return $obj }
    }
}

function Eval-Js($Ws,[string]$Expression) {
    $r = Invoke-Cdp $Ws 'Runtime.evaluate' @{expression=$Expression;returnByValue=$true;awaitPromise=$true}
    if ($r.result.exceptionDetails) { throw ('JavaScript evaluation failed: ' + $r.result.exceptionDetails.text) }
    return $r.result.result.value
}

function Fetch-ApiText($Ws,[string]$Url) {
    $escaped = $Url.Replace('\','\\').Replace("'","\'")
    $expr = @"
(async () => {
  try {
    const response = await fetch('$escaped', {credentials:'include', cache:'no-store'});
    const body = await response.text();
    return JSON.stringify({ok:response.ok,status:response.status,url:response.url,contentType:response.headers.get('content-type')||'',body});
  } catch(e) {
    return JSON.stringify({ok:false,status:0,url:'$escaped',contentType:'',error:String(e),body:''});
  }
})()
"@
    return ((Eval-Js $Ws $expr) | ConvertFrom-Json)
}

function Get-PageText($Ws) {
    $expr = "(() => ({title:document.title||'',url:location.href||'',text:(document.body?document.body.innerText:'').slice(0,3000),ready:document.readyState||''}))()"
    return Eval-Js $Ws $expr
}

function Wait-NormalPage($Ws,[int]$TimeoutSeconds=60) {
    $deadline=(Get-Date).AddSeconds($TimeoutSeconds); $last=$null
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 800
        try { $last=Get-PageText $Ws } catch { continue }
        $challenge=(([string]$last.title -match 'Just a moment') -or ([string]$last.text -match 'Verify you are human') -or ([string]$last.text -match 'Performing security verification'))
        if (-not $challenge -and $last.ready -eq 'complete' -and $last.url -match 'swimcloud\.com') { return [pscustomobject]@{ok=$true;challenge=$false;state=$last} }
    }
    $challenge=$false
    if ($last) { $challenge=(([string]$last.title -match 'Just a moment') -or ([string]$last.text -match 'Verify you are human') -or ([string]$last.text -match 'Performing security verification')) }
    return [pscustomobject]@{ok=$false;challenge=$challenge;state=$last}
}

function Safe-FileName([string]$s) { return ($s -replace '[^A-Za-z0-9._-]','_') }

function Invoke-ApiWithRetry($Ws,[string]$Url,[string]$Label) {
    $waits = @(0,20,40,60)
    for ($attempt=0; $attempt -lt $waits.Count; $attempt++) {
        if ($waits[$attempt] -gt 0) {
            Write-Log ("      rate-limit retry in {0}s..." -f $waits[$attempt])
            Start-Sleep -Seconds $waits[$attempt]
        }
        $resp = Fetch-ApiText $Ws $Url
        if ([int]$resp.status -ne 429) { return $resp }
        Write-Log ("      {0}: HTTP 429 on attempt {1}/{2}" -f $Label,($attempt+1),$waits.Count)
    }
    return $resp
}

$base=$PSScriptRoot
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
$script:OutDir=Join-Path $base ("swimcloud_api_output_extra_"+$stamp)
New-Item -ItemType Directory -Path $script:OutDir -Force | Out-Null
$script:LogPath=Join-Path $script:OutDir 'acquisition_log.txt'
Write-Log 'One-off additional Australian swimmer + Ian Thorpe acquisition. No installs required.'

# Read only manually verified profile IDs. Never guess URLs or edit the regular acquisition population.
$allProfiles = @(Import-Csv -LiteralPath (Join-Path $base '..\input\new_swimcloud_profiles_pending.csv'))
$profiles = @($allProfiles | Where-Object { $_.SwimmerID -match '^\d{4,}$' -and $_.SwimCloudURL -match '^https://www\.swimcloud\.com/swimmer/\d+' -and ([regex]::Match($_.SwimCloudURL, '/swimmer/(\d+)')).Groups[1].Value -eq $_.SwimmerID })
foreach ($p in $allProfiles) { if ($p -notin $profiles) { Write-Host ('SKIPPING unverified/missing SwimCloud URL: ' + $p.Swimmer) -ForegroundColor Yellow } }
if ($profiles.Count -eq 0) { throw 'No valid IDs/URLs. Populate new_swimcloud_profiles_pending.csv first. No requests were made.' }
Write-Host ('Ready to acquire ' + $profiles.Count + ' validated profiles, not part of the regular update.') -ForegroundColor Cyan
$events=@(
    [pscustomobject]@{Name='100 Free LCM';Key='1|100|L|1'},
    [pscustomobject]@{Name='200 Free LCM';Key='1|200|L|1'},
    [pscustomobject]@{Name='400 Free LCM';Key='1|400|L|1'},
    [pscustomobject]@{Name='100 Fly LCM';Key='4|100|L|1'},
    [pscustomobject]@{Name='100 Free SCM';Key='1|100|S|1'},
    [pscustomobject]@{Name='200 Free SCM';Key='1|200|S|1'},
    [pscustomobject]@{Name='400 Free SCM';Key='1|400|S|1'},
    [pscustomobject]@{Name='100 Fly SCM';Key='4|100|S|1'}
)
Write-Log ("Profiles: {0}; target events: {1}" -f $profiles.Count,$events.Count)

$browser=Get-BrowserPath; if (-not $browser) { throw 'Chrome or Edge was not found.' }
$port=Get-FreePort
$tempProfile=Join-Path $env:TEMP ("SwimCloud_API_extra_"+$stamp)
New-Item -ItemType Directory -Path $tempProfile -Force | Out-Null
$startUrl=([string]$profiles[0].SwimCloudURL)
$args=@("--remote-debugging-port=$port",'--remote-allow-origins=*',"--user-data-dir=$tempProfile",'--no-first-run','--no-default-browser-check','--new-window',$startUrl)
$proc=Start-Process -FilePath $browser -ArgumentList $args -PassThru
Write-Log ('Opened '+$startUrl)

$devtools=$null; $deadline=(Get-Date).AddSeconds(30)
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 500
    try { $targets=Invoke-RestMethod -Uri ("http://127.0.0.1:$port/json/list") -TimeoutSec 2; $devtools=$targets|Where-Object{$_.type -eq 'page' -and $_.webSocketDebuggerUrl}|Select-Object -First 1; if ($devtools){break} } catch {}
}
if (-not $devtools) { throw 'Could not connect to Chrome/Edge remote debugging.' }
$ws=Connect-Cdp $devtools.webSocketDebuggerUrl
$null=Invoke-Cdp $ws 'Page.enable'; $null=Invoke-Cdp $ws 'Runtime.enable'

Write-Host ''
Write-Host 'A normal Chrome/Edge window is open on the first additional swimmer's Times page.' -ForegroundColor Cyan
Write-Host 'Clear any SwimCloud/Cloudflare verification. When the normal Times page is visible, press ENTER here.' -ForegroundColor Yellow
[void](Read-Host)
$page=Wait-NormalPage $ws 60
if ($page.challenge) { throw 'Cloudflare verification was not cleared.' }

# Warm-up request and a deliberate pause before the bulk run.
$warm=Invoke-ApiWithRetry $ws ('https://www.swimcloud.com/api/swimmers/'+$profiles[0].SwimmerID+'/profile_fastest_times/') 'warm-up'
Write-Log ("Warm-up HTTP {0}; waiting 10 seconds before bulk acquisition." -f $warm.status)
Start-Sleep -Seconds 10

$rows=New-Object System.Collections.Generic.List[object]
$profileNo=0
foreach ($p in $profiles) {
    $profileNo++
    Write-Log ("[{0}/{1}] {2} ({3})" -f $profileNo,$profiles.Count,$p.Swimmer,$p.SwimmerID)
    $dir=Join-Path $script:OutDir ((Safe-FileName $p.Swimmer)+'_'+$p.SwimmerID)
    New-Item -ItemType Directory -Path $dir -Force | Out-Null

    foreach ($e in $events) {
        $encoded=[Uri]::EscapeDataString($e.Key)
        $url="https://www.swimcloud.com/api/swimmers/$($p.SwimmerID)/times_by_event/?event=$encoded"
        try {
            $resp=Invoke-ApiWithRetry $ws $url $e.Name
            $bytes=[Text.Encoding]::UTF8.GetByteCount([string]$resp.body)
            $file=(Safe-FileName ($e.Name.Replace(' ','_')))+'.json'
            [System.IO.File]::WriteAllText((Join-Path $dir $file),($resp|ConvertTo-Json -Depth 30),(New-Object System.Text.UTF8Encoding($false)))

            $status='http_error'
            $count=''
            if ($resp.ok) {
                try {
                    $arr=@(([string]$resp.body) | ConvertFrom-Json)
                    $count=$arr.Count
                    if ($arr.Count -eq 0) { $status='not_available' } else { $status='captured' }
                } catch { $status='parse_error' }
            }
            Write-Log ("    {0}: HTTP {1}; status={2}; swims={3}; bytes={4}" -f $e.Name,$resp.status,$status,$count,$bytes)
            $rows.Add([pscustomobject]@{Swimmer=$p.Swimmer;SwimmerID=$p.SwimmerID;DOB=$p.DOB;Group=$p.Group;Event=$e.Name;EventKey=$e.Key;Status=$status;HttpStatus=$resp.status;Swims=$count;Bytes=$bytes})
        } catch {
            Write-Log ("    {0}: ERROR {1}" -f $e.Name,$_.Exception.Message)
            $rows.Add([pscustomobject]@{Swimmer=$p.Swimmer;SwimmerID=$p.SwimmerID;DOB=$p.DOB;Group=$p.Group;Event=$e.Name;EventKey=$e.Key;Status='error';HttpStatus='';Swims='';Bytes=0})
        }
        Start-Sleep -Milliseconds 2500
    }
    Write-Log '  Profile pause: 8 seconds.'
    Start-Sleep -Seconds 8
}

$rows|Export-Csv -LiteralPath (Join-Path $script:OutDir 'capture_index.csv') -NoTypeInformation -Encoding UTF8
$summary=[ordered]@{
    CreatedAt=(Get-Date).ToString('s'); Profiles=$profiles.Count; TargetEvents=$events.Count; EventRows=$rows.Count;
    Captured=@($rows|Where-Object{$_.Status -eq 'captured'}).Count;
    NotAvailable=@($rows|Where-Object{$_.Status -eq 'not_available'}).Count;
    HttpErrors=@($rows|Where-Object{$_.Status -eq 'http_error'}).Count;
    ParseErrors=@($rows|Where-Object{$_.Status -eq 'parse_error'}).Count;
    Errors=@($rows|Where-Object{$_.Status -eq 'error'}).Count
}
$summary|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $script:OutDir 'summary.json') -Encoding UTF8
try{$ws.Dispose()}catch{}
$zipPath=Join-Path $base ("SwimCloud_API_Output_extra_"+$stamp+'.zip')
if(Test-Path -LiteralPath $zipPath){Remove-Item -LiteralPath $zipPath -Force}
Compress-Archive -Path (Join-Path $script:OutDir '*') -DestinationPath $zipPath -CompressionLevel Optimal
Write-Log ('Created ZIP: '+$zipPath)
Write-Host ''
Write-Host 'ACQUISITION COMPLETE' -ForegroundColor Green
Write-Host ('Output ZIP: '+$zipPath) -ForegroundColor Green
Write-Host 'Send this ZIP for future profile integration; this one-off acquisition does not itself update the master workbook.' -ForegroundColor Green

Set-Content -LiteralPath (Join-Path $base '..\input\latest_extra_zip.txt') -Value $zipPath -Encoding UTF8

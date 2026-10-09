$ErrorActionPreference = 'Stop'

function W([string]$m) {
    $s = Get-Date -Format 'HH:mm:ss'
    $l = "[$s] $m"
    Write-Host $l
    Add-Content -LiteralPath $script:Log -Value $l -Encoding UTF8
}

function Browser {
    foreach ($p in @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    )) {
        if ($p -and (Test-Path $p)) { return $p }
    }
    return $null
}

function Port {
    $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $l.Start()
    $p = ([System.Net.IPEndPoint]$l.LocalEndpoint).Port
    $l.Stop()
    return $p
}

function Conn($u) {
    $w = New-Object System.Net.WebSockets.ClientWebSocket
    $w.ConnectAsync([Uri]$u, [Threading.CancellationToken]::None).Wait()
    return $w
}

function Recv($w) {
    $ms = New-Object IO.MemoryStream
    try {
        do {
            $b = New-Object byte[] 65536
            $seg = New-Object System.ArraySegment[byte] -ArgumentList (, $b)
            $cts = New-Object Threading.CancellationTokenSource
            $cts.CancelAfter(30000)
            $t = $w.ReceiveAsync($seg, $cts.Token)
            $t.Wait()
            $r = $t.Result
            if ($r.MessageType -eq [Net.WebSockets.WebSocketMessageType]::Close) { throw 'CDP closed' }
            $ms.Write($b, 0, $r.Count)
        } while (-not $r.EndOfMessage)
        return [Text.Encoding]::UTF8.GetString($ms.ToArray())
    }
    finally { $ms.Dispose() }
}

$script:id = 0
function CDP($w, [string]$m, $p = $null) {
    $script:id++
    $id = $script:id
    $o = [ordered]@{ id = $id; method = $m }
    if ($null -ne $p) { $o.params = $p }
    $bs = [Text.Encoding]::UTF8.GetBytes(($o | ConvertTo-Json -Depth 50 -Compress))
    $seg = New-Object System.ArraySegment[byte] -ArgumentList (, $bs)
    $w.SendAsync($seg, [Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).Wait()
    while ($true) {
        try { $x = (Recv $w) | ConvertFrom-Json } catch { continue }
        if ($null -ne $x.id -and [int]$x.id -eq $id) { return $x }
    }
}

function Eval($w, [string]$e) {
    $r = CDP $w 'Runtime.evaluate' @{ expression = $e; returnByValue = $true; awaitPromise = $true }
    if ($r.result.exceptionDetails) { throw 'JavaScript evaluation failed' }
    return $r.result.result.value
}

function Fetch($w, [string]$u) {
    $e = $u.Replace("'", "\'")
    $js = @"
(async()=>{
  const c = new AbortController();
  const t = setTimeout(()=>c.abort(), 25000);
  try {
    const r = await fetch('$e',{credentials:'include',cache:'no-store',signal:c.signal});
    const body = await r.text();
    return JSON.stringify({ok:r.ok,status:r.status,url:r.url,body:body});
  } catch(err) {
    return JSON.stringify({ok:false,status:0,url:'$e',body:'',error:String(err)});
  } finally { clearTimeout(t); }
})()
"@
    return ((Eval $w $js) | ConvertFrom-Json)
}

function Retry($w, [string]$u) {
    foreach ($wait in @(0,20,40,60)) {
        if ($wait) { W "Rate limit - waiting $wait sec"; Start-Sleep -Seconds $wait }
        $r = Fetch $w $u
        if ([int]$r.status -ne 429) { return $r }
    }
    return $r
}

function Safe([string]$s) { return ($s -replace '[^A-Za-z0-9._-]', '_') }

function PageState($w) {
    $js = "(()=>({url:location.href,title:document.title,body:(document.body?document.body.innerText.slice(0,2000):''),hasProgression:[...document.querySelectorAll('button')].some(x=>(x.textContent||'').trim().toUpperCase()==='EVENT PROGRESSION')}))()"
    return (Eval $w $js)
}

function Ensure-TimesPage($w, [string]$swimmerId, [string]$targetUrl) {
    while ($true) {
        Start-Sleep -Seconds 2
        $s = PageState $w
        $okUrl = ([string]$s.url -match "/swimmer/$swimmerId/times")
        $challenge = (([string]$s.title -match 'Just a moment') -or ([string]$s.body -match 'Verify you are human|Performing security verification|Cloudflare'))
        if ($okUrl -and -not $challenge -and $s.hasProgression) {
            W "Times page confirmed for profile $swimmerId."
            return
        }

        Write-Host ''
        Write-Host "The normal SwimCloud Times page is not ready for profile $swimmerId." -ForegroundColor Yellow
        Write-Host "Browser URL seen: $($s.url)" -ForegroundColor Yellow
        Write-Host "Please use the open browser to complete Cloudflare and/or navigate to:" -ForegroundColor Yellow
        Write-Host $targetUrl -ForegroundColor Cyan
        Write-Host "Wait until the normal Times page and EVENT PROGRESSION control are visible." -ForegroundColor Yellow
        [void](Read-Host 'Then press ENTER here to retry')
    }
}

$base = $PSScriptRoot
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$out = Join-Path $base ("swimcloud_grayson_all_events_v8_" + $stamp)
New-Item -ItemType Directory -Path $out -Force | Out-Null
$script:Log = Join-Path $out 'log.txt'
W 'Starting Grayson all-event acquisition v8.'

$rows = New-Object Collections.Generic.List[object]
$profiles = Import-Csv (Join-Path $base '..\input\grayson_profiles.csv')
$browser = Browser
if (-not $browser) { throw 'Chrome or Edge not found' }

$port = Port
$tmp = Join-Path $env:TEMP ("SwimCloud_Grayson_v8_" + $stamp)
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$startId = '3651023'
$start = "https://www.swimcloud.com/swimmer/$startId/times/"

W 'Opening browser. The script will now wait for you before doing anything else.'
$proc = Start-Process $browser -ArgumentList @(
    "--remote-debugging-port=$port",
    '--remote-allow-origins=*',
    "--user-data-dir=$tmp",
    '--no-first-run',
    '--new-window',
    $start
) -PassThru

Write-Host ''
Write-Host 'BROWSER STEP' -ForegroundColor Cyan
Write-Host '1. Complete any Cloudflare verification in the browser.' -ForegroundColor Yellow
Write-Host '2. Make sure Grayson''s normal SwimCloud TIMES page is fully visible.' -ForegroundColor Yellow
Write-Host '3. Take as long as you need. The script is NOT timing you.' -ForegroundColor Yellow
Write-Host '4. Only then return here and press ENTER.' -ForegroundColor Yellow
[void](Read-Host 'Press ENTER when the Times page is ready')

# Connect to the browser only after the user confirms the page is ready.
$dev = $null
while (-not $dev) {
    try {
        $ts = Invoke-RestMethod "http://127.0.0.1:$port/json/list" -TimeoutSec 3
        $dev = $ts | Where-Object { $_.type -eq 'page' -and $_.webSocketDebuggerUrl } | Select-Object -First 1
    }
    catch { }
    if (-not $dev) {
        Write-Host 'Could not connect to the browser yet. Leave it open.' -ForegroundColor Yellow
        [void](Read-Host 'Press ENTER to retry browser connection')
    }
}

$ws = Conn $dev.webSocketDebuggerUrl
$null = CDP $ws 'Page.enable'
$null = CDP $ws 'Runtime.enable'
Ensure-TimesPage $ws $startId $start

try {
    foreach ($p in $profiles) {
        $sid = [string]$p.SwimmerID
        $url = "https://www.swimcloud.com/swimmer/$sid/times/"
        W "Profile $sid"

        $current = [string](Eval $ws 'location.href')
        if ($current -notmatch "/swimmer/$sid/times") {
            $null = CDP $ws 'Page.navigate' @{ url = $url }
            Start-Sleep -Seconds 4
        }
        Ensure-TimesPage $ws $sid $url

        # Open Event Progression and discover every current option from the live page.
        $clicked = Eval $ws "(()=>{const b=[...document.querySelectorAll('button')].find(x=>(x.textContent||'').trim().toUpperCase()==='EVENT PROGRESSION');if(b){b.click();return true}return false})()"
        if (-not $clicked) {
            W "  Could not click Event Progression automatically."
            Write-Host 'Please click EVENT PROGRESSION manually in the browser.' -ForegroundColor Yellow
            [void](Read-Host 'Press ENTER once its event dropdown is visible')
        }

        $opts = $null
        for ($i=0; $i -lt 20; $i++) {
            try {
                $opts = Eval $ws "(()=>[...document.querySelectorAll('#swimmer-profile-times select option')].map(o=>({key:o.value,name:(o.textContent||'').trim()})).filter(x=>x.key&&x.name))()"
            } catch { $opts = $null }
            if ($opts -and $opts.Count -gt 0) { break }
            Start-Sleep -Seconds 1
        }

        if (-not $opts -or $opts.Count -eq 0) {
            W '  No event options found; pausing for manual confirmation.'
            Write-Host 'No event list was detected.' -ForegroundColor Yellow
            Write-Host 'Make sure EVENT PROGRESSION is open and the event dropdown is visible.' -ForegroundColor Yellow
            [void](Read-Host 'Press ENTER to try reading the event list again')
            $opts = Eval $ws "(()=>[...document.querySelectorAll('#swimmer-profile-times select option')].map(o=>({key:o.value,name:(o.textContent||'').trim()})).filter(x=>x.key&&x.name))()"
        }

        W ("  discovered " + $opts.Count + " event options")
        $dir = Join-Path $out ("Grayson_Coulter_" + $sid)
        New-Item -ItemType Directory -Path $dir -Force | Out-Null

        $n = 0
        foreach ($e in $opts) {
            $n++
            $enc = [Uri]::EscapeDataString([string]$e.key)
            $api = "https://www.swimcloud.com/api/swimmers/$sid/times_by_event/?event=$enc"
            $r = Retry $ws $api
            $file = (Safe(([string]$e.name).Replace(' ', '_'))) + '.json'
            [IO.File]::WriteAllText(
                (Join-Path $dir $file),
                ($r | ConvertTo-Json -Depth 30),
                (New-Object Text.UTF8Encoding($false))
            )

            $count = ''
            $status = 'http_error'
            if ($r.ok) {
                try {
                    $a = @(([string]$r.body) | ConvertFrom-Json)
                    $count = $a.Count
                    $status = if ($a.Count -eq 0) { 'empty' } else { 'captured' }
                }
                catch { $status = 'parse_error' }
            }
            W ("    [$n/$($opts.Count)] $($e.name): HTTP $($r.status); $status; swims=$count")
            $rows.Add([pscustomobject]@{
                Profile = $sid
                Event = $e.name
                EventKey = $e.key
                Status = $status
                HttpStatus = $r.status
                Swims = $count
            })
            Start-Sleep -Milliseconds 2500
        }
        Start-Sleep -Seconds 8
    }
}
catch {
    W ("ERROR: " + $_.Exception.Message)
    Write-Host ''
    Write-Host 'The acquisition stopped early, but a partial ZIP will still be created.' -ForegroundColor Yellow
}
finally {
    $rows | Export-Csv (Join-Path $out 'capture_index.csv') -NoTypeInformation -Encoding UTF8
    $zip = Join-Path $base ("SwimCloud_Grayson_AllEvents_Output_update_" + $stamp + '.zip')
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Compress-Archive -Path (Join-Path $out '*') -DestinationPath $zip -CompressionLevel Optimal
    W ("Created ZIP: $zip")
    Set-Content -LiteralPath (Join-Path $base '..\input\latest_grayson_zip.txt') -Value $zip -Encoding UTF8
    Write-Host ''
    Write-Host 'COMPLETE - send the output ZIP back to ChatGPT.' -ForegroundColor Green
}

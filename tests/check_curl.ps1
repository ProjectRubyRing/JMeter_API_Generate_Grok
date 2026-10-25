# Parses curl commands into api_definition.xlsx and checks the generated JMX.
# powershell -NoProfile -ExecutionPolicy Bypass -File tests\check_curl.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$tool = Join-Path $root 'jmeter_api_tool.ps1'
$script:fail = 0

function Assert-True([bool]$cond, [string]$msg) {
    if (-not $cond) {
        Write-Host ('FAIL: ' + $msg) -ForegroundColor Red
        $script:fail++
    }
}

function Write-Utf8([string]$path, [string]$text) {
    $enc = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($path, $text, $enc)
}

function Get-SamplerChunks([string]$text) {
    $parts = [regex]::Split($text, '(?=<HTTPSamplerProxy )')
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($p in $parts) { if ($p.StartsWith('<HTTPSamplerProxy ')) { $out.Add($p) } }
    return ,$out
}

function Get-Prop([string]$chunk, [string]$name) {
    $pattern = '<stringProp name="' + [regex]::Escape($name) + '">([^<]*)</stringProp>'
    $m = [regex]::Match($chunk, $pattern)
    if ($m.Success) { return $m.Groups[1].Value }
    return ''
}

function Find-Chunk($chunks, [string]$path) {
    foreach ($c in $chunks) {
        if ((Get-Prop $c 'HTTPSampler.path') -eq $path) { return $c }
    }
    return ''
}

$work = Join-Path $root 'tests\_curl'
if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
$in = Join-Path $work 'in'
New-Item -ItemType Directory -Path $in | Out-Null
Write-Utf8 (Join-Path $in 'curl_note.txt') 'hello-note'
Write-Utf8 (Join-Path $in 'curl_body.json') '{"id":9,"name":"sample"}'
$curl = @'
# login
curl -sS --connect-timeout 2.5 --max-time 9 -X POST "https://api.example.com/v1/login" \
  -H "Content-Type: application/json" \
  -H "X-Request-Id: abc" \
  --data-raw "{\"user\":\"taro\",\"pass\":\"secret\"}"

curl "https://api.example.com/v1/items?page=1&size=2"

curl -L -G "https://api.example.com/v1/search" \
  --data-urlencode "q=a b" \
  -H "Authorization: Bearer TOKEN" \
  -A "demo-agent"

curl -k -u "user:p@ss" https://api.example.com/v1/me

curl -X POST https://api.example.com/v1/form -d "a=1&b=2"

curl -X POST https://api.example.com/v1/raw -d "hello world"

curl.exe -X PUT "https://api.example.com/v1/item" -H "Content-Type: application/json" -d "{\"id\":1}"

curl -F "title=hello" -F "file=@curl_note.txt" https://api.example.com/v1/upload --next -X DELETE https://api.example.com/v1/items/9

curl --json @curl_body.json https://api.example.com/v1/fromfile

curl https://api.example.com/a https://api.example.com/b

curl -X GET ^
  "https://api.example.com/v1/caret"
'@
# PowerShell-style backtick continuation is appended without a here-string eating the backtick.
$tick = [string][char]96
$curl += "`ncurl $tick`n  -X GET $tick`n  https://api.example.com/v1/tick`n"
$curl += "FOO=bar curl https://api.example.com/v1/envassign`n"
Write-Utf8 (Join-Path $in 'sample.curl') $curl

$xlsx = Join-Path $work 'api_definition.xlsx'
$csv = Join-Path $work 'api_definition.csv'
Copy-Item -LiteralPath (Join-Path $root 'input\api_definition.xlsx') -Destination $xlsx
Copy-Item -LiteralPath (Join-Path $root 'input\api_definition.csv') -Destination $csv
$log = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool curl -i (Join-Path $in 'sample.curl') -o $xlsx 2>&1 | Out-String
Assert-True ($LASTEXITCODE -eq 0) ('curl import failed: ' + $log)
Assert-True (Test-Path -LiteralPath $xlsx) 'xlsx was not created'
$csvText = [System.IO.File]::ReadAllText($csv)
Assert-True ($csvText.Contains('https://api.example.com/v1/login')) 'csv missing added row'
Assert-True ($csvText.Contains('https://postman-echo.com/post')) 'csv lost the existing row'

$out = Join-Path $work 'jmx'
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Recurse -Force }
New-Item -ItemType Directory -Path $out | Out-Null
$gen = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool generate -i $xlsx -o $out 2>&1 | Out-String
Assert-True ($LASTEXITCODE -eq 0) ('generate failed: ' + $gen)
$jmxPath = Get-ChildItem -LiteralPath $out -Filter *.jmx | Select-Object -First 1
Assert-True ($null -ne $jmxPath) 'jmx missing'
if ($null -eq $jmxPath) { Write-Host ($script:fail.ToString() + ' failure(s)'); exit 1 }
$text = [System.IO.File]::ReadAllText($jmxPath.FullName)
$chunks = Get-SamplerChunks $text
Assert-True ($chunks.Count -eq 19) ('sampler count=' + $chunks.Count + "`n" + $log)
Assert-True ($text.Contains('postman-echo.com/post')) 'existing sample API was removed'
Assert-True ($text.Contains('CLIENT_ID')) 'template settings were replaced'
Assert-True ($text.Contains('follow_redirects">true<')) 'redirect setting was changed'

$login = Find-Chunk $chunks '/v1/login'
Assert-True ((Get-Prop $login 'HTTPSampler.method') -eq 'POST') 'login method'
Assert-True ($login.Contains('taro')) 'login body'
Assert-True ($login.Contains('X-Request-Id')) 'login header'
Assert-True ($login.Contains('application/json')) 'login content type'
Assert-True ($login.Contains('postBodyRaw">true<')) 'login raw body'

$items = Find-Chunk $chunks '/v1/items'
Assert-True ((Get-Prop $items 'HTTPSampler.method') -eq 'GET') 'items method'
Assert-True ($items.Contains('Argument.name">page<')) 'page query'
Assert-True ($items.Contains('Argument.name">size<')) 'size query'
Assert-True (-not (Get-Prop $items 'HTTPSampler.path').Contains('?')) 'query left on path'

$search = Find-Chunk $chunks '/v1/search'
Assert-True ((Get-Prop $search 'HTTPSampler.method') -eq 'GET') 'search method'
Assert-True ($search.Contains('Argument.value">a%20b<')) 'urlencoded query'
Assert-True ($search.Contains('always_encode">false<')) 'encoded query not encoded again'
Assert-True ($search.Contains('Bearer TOKEN')) 'bearer header'
Assert-True ($search.Contains('demo-agent')) 'user agent'

$me = Find-Chunk $chunks '/v1/me'
$basic = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('user:p@ss'))
Assert-True ($me.Contains($basic)) ('basic auth missing: ' + $basic)
Assert-True ($me.Contains('curl -k')) 'insecure note'

$form = Find-Chunk $chunks '/v1/form'
Assert-True ((Get-Prop $form 'HTTPSampler.method') -eq 'POST') 'form method'
Assert-True ($form.Contains('Argument.name">a<') -and $form.Contains('Argument.name">b<')) 'form fields'
Assert-True (-not $form.Contains('postBodyRaw">true<')) 'safe form should not be raw'

$raw = Find-Chunk $chunks '/v1/raw'
Assert-True ($raw.Contains('hello world')) 'raw body'
Assert-True ($raw.Contains('application/x-www-form-urlencoded')) 'raw content type'
Assert-True ($raw.Contains('postBodyRaw">true<')) 'raw flag'

$item = Find-Chunk $chunks '/v1/item'
Assert-True ((Get-Prop $item 'HTTPSampler.method') -eq 'PUT') 'put method'
Assert-True ($item.Contains('&quot;id&quot;:1') -or $item.Contains('&quot;id&quot;: 1')) 'put json'

$up = Find-Chunk $chunks '/v1/upload'
Assert-True ($up.Contains('DO_MULTIPART_POST">true<')) 'multipart'
Assert-True ($up.Contains('Argument.name">title<')) 'multipart title'
Assert-True ($up.Contains('hello-note')) 'multipart file text'

$del = Find-Chunk $chunks '/v1/items/9'
Assert-True ((Get-Prop $del 'HTTPSampler.method') -eq 'DELETE') 'delete method'
Assert-True (-not $del.Contains('hello-note')) 'next request kept the form body'

$fileApi = Find-Chunk $chunks '/v1/fromfile'
Assert-True ($fileApi.Contains('sample')) 'json file body'
Assert-True ($fileApi.Contains('Accept')) 'json accept'

Assert-True ((Get-Prop (Find-Chunk $chunks '/a') 'HTTPSampler.method') -eq 'GET') 'url a'
Assert-True ((Get-Prop (Find-Chunk $chunks '/b') 'HTTPSampler.method') -eq 'GET') 'url b'
Assert-True ((Get-Prop (Find-Chunk $chunks '/v1/caret') 'HTTPSampler.method') -eq 'GET') 'caret continuation'
Assert-True ((Get-Prop (Find-Chunk $chunks '/v1/tick') 'HTTPSampler.method') -eq 'GET') 'backtick continuation'
Assert-True ((Get-Prop (Find-Chunk $chunks '/v1/envassign') 'HTTPSampler.method') -eq 'GET') 'env assignment prefix'

Assert-True ($text.Contains('${__P(CONNECT_TIMEOUT,10000)}')) 'template connect timeout changed'

$again = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool curl -i (Join-Path $in 'sample.curl') -o $xlsx 2>&1 | Out-String
Assert-True ($LASTEXITCODE -eq 0) ('second append failed: ' + $again)
$outB = Join-Path $work 'jmx-b'
New-Item -ItemType Directory -Path $outB | Out-Null
$genB = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool generate -i $xlsx -o $outB 2>&1 | Out-String
Assert-True ($LASTEXITCODE -eq 0) ('second generate failed: ' + $genB)
$jmxB = Get-ChildItem -LiteralPath $outB -Filter *.jmx | Select-Object -First 1
$chunksB = Get-SamplerChunks ([System.IO.File]::ReadAllText($jmxB.FullName))
Assert-True ($chunksB.Count -eq 34) ('second append count=' + $chunksB.Count)

Write-Utf8 (Join-Path $in 'noloc.curl') "curl https://api.example.com/ping`n"
$xlsx2 = Join-Path $work 'noloc.xlsx'
$log2 = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool curl -i (Join-Path $in 'noloc.curl') -o $xlsx2 2>&1 | Out-String
Assert-True ($LASTEXITCODE -eq 0) ('noloc import failed: ' + $log2)
$out2 = Join-Path $work 'jmx2'
New-Item -ItemType Directory -Path $out2 | Out-Null
$gen2 = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool generate -i $xlsx2 -o $out2 2>&1 | Out-String
Assert-True ($LASTEXITCODE -eq 0) ('noloc generate failed: ' + $gen2)
$jmx2 = Get-ChildItem -LiteralPath $out2 -Filter *.jmx | Select-Object -First 1
$t2 = [System.IO.File]::ReadAllText($jmx2.FullName)
Assert-True ((Get-SamplerChunks $t2).Count -eq 5) 'missing file should start from the template plus one row'
Assert-True ($t2.Contains('/ping')) 'ping row missing'
$csv2 = [System.IO.File]::ReadAllText((Join-Path $work 'noloc.csv'))
Assert-True ($csv2.Contains('https://api.example.com/ping')) 'noloc csv missing the added row'

$py = $null
foreach ($cand in @('python', 'python3')) {
    $cmd = Get-Command $cand -ErrorAction SilentlyContinue
    if ($null -ne $cmd) { $py = $cmd.Source; break }
}
if ($py) {
    $sh = [System.IO.File]::ReadAllText((Join-Path $root 'jmeter_api_tool.sh'))
    $mark = "__PYTHON_CORE__"
    $a = $sh.IndexOf($mark)
    $b = $sh.LastIndexOf($mark)
    Assert-True (($a -ge 0) -and ($b -gt $a)) 'python core markers missing'
    $lineEnd = $sh.IndexOf("`n", $a)
    $code = $sh.Substring($lineEnd + 1, $b - ($lineEnd + 1)).Trim()
    $pyFile = Join-Path $work 'core.py'
    Write-Utf8 $pyFile ($code + "`n")
    & $py -m py_compile $pyFile
    if ($LASTEXITCODE -ne 0) { Assert-True $false 'python core failed to compile'; exit 1 }
    $pyXlsx = Join-Path $work 'py.xlsx'
    Copy-Item -LiteralPath (Join-Path $root 'input\api_definition.xlsx') -Destination $pyXlsx
    Copy-Item -LiteralPath (Join-Path $root 'input\api_definition.csv') -Destination (Join-Path $work 'py.csv')
    $pyLog = & $py -I $pyFile curl -i (Join-Path $in 'sample.curl') -o $pyXlsx 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 0) ('python curl failed: ' + $pyLog)
    $pyOut = Join-Path $work 'pyjmx'
    New-Item -ItemType Directory -Path $pyOut | Out-Null
    $pyGen = & $py -I $pyFile generate -i $pyXlsx -o $pyOut 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 0) ('python generate failed: ' + $pyGen)
    $pyJmx = Get-ChildItem -LiteralPath $pyOut -Filter *.jmx | Select-Object -First 1
    $pyText = [System.IO.File]::ReadAllText($pyJmx.FullName)
    $sig = {
        param($body)
        $sb = New-Object System.Text.StringBuilder
        foreach ($c in (Get-SamplerChunks $body)) {
            [void]$sb.AppendLine((Get-Prop $c 'HTTPSampler.method') + ' ' + (Get-Prop $c 'HTTPSampler.path'))
            $bodyValue = Get-Prop $c 'Argument.value'
            if ($bodyValue) { [void]$sb.AppendLine($bodyValue) }
            foreach ($hm in [regex]::Matches($c, 'Header.name">([^<]*)<')) { [void]$sb.AppendLine('H ' + $hm.Groups[1].Value) }
        }
        return $sb.ToString()
    }
    $left = & $sig $text
    $right = & $sig $pyText
    if ($left -ne $right) {
        Write-Utf8 (Join-Path $work 'sig-ps.txt') $left
        Write-Utf8 (Join-Path $work 'sig-py.txt') $right
    }
    Assert-True ($left -eq $right) 'ps and python sampler signatures differ (see tests/_curl/sig-ps.txt)'
} else {
    Write-Host 'WARN: python not found; skipped sh/ps parity'
}

if ($script:fail -gt 0) {
    Write-Host ($script:fail.ToString() + ' failure(s)') -ForegroundColor Red
    exit 1
}
Write-Host 'curl import checks passed'
exit 0

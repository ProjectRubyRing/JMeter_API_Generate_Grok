# Checks Excel/CSV and OpenAPI JMX generation.
# powershell -NoProfile -ExecutionPolicy Bypass -File tests\check_generate.ps1
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

function Get-Arg([string]$chunk, [string]$name) {
    $pattern = '<elementProp name="' + [regex]::Escape($name) + '" elementType="HTTPArgument">([\s\S]*?)</elementProp>'
    $m = [regex]::Match($chunk, $pattern)
    if (-not $m.Success) { return $null }
    $body = $m.Groups[1].Value
    $enc = [regex]::Match($body, '<boolProp name="HTTPArgument.always_encode">(true|false)</boolProp>').Groups[1].Value
    $val = [regex]::Match($body, '<stringProp name="Argument.value">([^<]*)</stringProp>').Groups[1].Value
    return @{ enc = $enc; value = $val }
}

function Invoke-Gen([string]$inputFile, [string]$outDir, [string[]]$extra) {
    if (Test-Path -LiteralPath $outDir) { Remove-Item -LiteralPath $outDir -Recurse -Force }
    New-Item -ItemType Directory -Path $outDir | Out-Null
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $tool, 'generate', '-i', $inputFile, '-o', $outDir) + $extra
    $log = & powershell.exe @argList 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw ('generate failed: ' + $inputFile + "`n" + $log) }
    $jmx = Get-ChildItem -LiteralPath $outDir -Filter *.jmx | Select-Object -First 1
    if ($null -eq $jmx) { throw ('jmx missing: ' + $outDir + "`n" + $log) }
    return @{ text = [System.IO.File]::ReadAllText($jmx.FullName); log = $log; path = $jmx.FullName }
}

function Get-Signature([string]$text) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in (Get-SamplerChunks $text)) {
        [void]$sb.AppendLine((Get-Prop $c 'HTTPSampler.method') + ' ' + (Get-Prop $c 'HTTPSampler.path'))
        foreach ($m in [regex]::Matches($c, '<stringProp name="Argument.name">([^<]*)</stringProp>')) {
            [void]$sb.AppendLine('  arg ' + $m.Groups[1].Value)
        }
        $body = Get-Prop $c 'Argument.value'
        if ($body) { [void]$sb.AppendLine('  body ' + ($body -replace '\s+', ' ')) }
    }
    return $sb.ToString()
}

function Find-Chunk($chunks, [string]$id) {
    foreach ($c in $chunks) { if ($c.Contains($id)) { return $c } }
    return ''
}

Write-Host '== Excel / CSV =='
$xlsx = Invoke-Gen (Join-Path $root 'input\api_definition.xlsx') (Join-Path $root 'tests\_out_xlsx') @()
$csv = Invoke-Gen (Join-Path $root 'input\api_definition.csv') (Join-Path $root 'tests\_out_csv') @('-s', (Join-Path $root 'input\api_definition_settings.csv'))
$sigX = Get-Signature $xlsx.text
$sigC = Get-Signature $csv.text
Assert-True ($sigX -eq $sigC) 'xlsx and csv sampler signatures differ'
Assert-True ($sigX.Contains('POST /post')) 'sample POST path missing'
Assert-True ($sigX.Contains('GET /get')) 'sample GET path missing'
Assert-True ($xlsx.text.Contains('source=jmeter')) 'POST query was not kept on the path'
Assert-True ($xlsx.text.Contains('postBodyRaw')) 'JSON body is not raw'
Assert-True ($xlsx.text.Contains('itemCode')) 'form field missing'

Write-Host '== Excel regression =='
$old = 'C:\Users\taka_\AppData\Local\Temp\JMeter_API_Generate_upstream\jmeter_api_tool.ps1'
if (Test-Path -LiteralPath $old) {
    $oldOut = Join-Path $root 'tests\_out_old'
    if (Test-Path -LiteralPath $oldOut) { Remove-Item -LiteralPath $oldOut -Recurse -Force }
    New-Item -ItemType Directory -Path $oldOut | Out-Null
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $old generate -i (Join-Path $root 'input\api_definition.xlsx') -o $oldOut | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $oldJmx = Get-ChildItem -LiteralPath $oldOut -Filter *.jmx | Select-Object -First 1
        $oldText = [System.IO.File]::ReadAllText($oldJmx.FullName)
        Assert-True ((Get-Signature $oldText) -eq $sigX) 'Excel signature changed from upstream'
    } else {
        Write-Host 'WARN: upstream generate failed; skipped regression'
    }
}

Write-Host '== OpenAPI =='
$oa = Invoke-Gen (Join-Path $root 'tests\fixtures\sample_api.json') (Join-Path $root 'tests\_out_oa') @('-Seed', '1', '-UseExamples')
$chunks = Get-SamplerChunks $oa.text
Assert-True ($chunks.Count -eq 7) ('OpenAPI count=' + $chunks.Count)

$pet = Find-Chunk $chunks 'createPet'
Assert-True ($pet.Contains('&quot;id&quot;: 7')) 'allOf example id missing'
Assert-True ($pet.Contains('&quot;name&quot;: &quot;n&quot;')) 'allOf example name missing'
Assert-True ($pet.Contains('Header.name">X-API-KEY<')) 'global api key header missing'
Assert-True ($pet.Contains('XXXXXXX')) 'api key placeholder missing'

$show = Find-Chunk $chunks 'showPet'
Assert-True ((Get-Prop $show 'HTTPSampler.path') -eq '/v1/pets/a%2Fb') ('showPet path=' + (Get-Prop $show 'HTTPSampler.path'))
$tags = Get-Arg $show 'tags'
Assert-True ($null -ne $tags -and $tags.value -eq 'red%20blue' -and $tags.enc -eq 'false') 'spaceDelimited encoding'
$filt = Get-Arg $show 'filter[name]'
Assert-True ($null -ne $filt -and $filt.value -eq 'x%20y' -and $filt.enc -eq 'false') 'deepObject brackets'
$limit = Get-Arg $show 'limit'
Assert-True ($null -ne $limit -and $limit.value -eq '10') 'shared $ref query'
Assert-True ($show.Contains('Cookie.value">abc<')) 'cookie manager value'
Assert-True (-not $show.Contains('Header.name">Cookie<')) 'cookie header should be absent'
Assert-True ($show.Contains('Header.value">application/xml<')) 'Accept xml was overridden'
Assert-True ($show.Contains('Header.name">X-Trace<')) 'header parameter missing'
Assert-True ($show.Contains('Transaction-ID')) 'default API_HEADERS were not added'

$up = Find-Chunk $chunks 'upload'
Assert-True ((Get-Prop $up 'HTTPSampler.path') -eq '/upload/files') ('upload path=' + (Get-Prop $up 'HTTPSampler.path'))
Assert-True ((Get-Prop $up 'HTTPSampler.domain') -eq '${HOST_2}') 'operation server host var'
Assert-True ((Get-Prop $up 'HTTPSampler.port') -eq '${PORT_2}') 'operation server port var'
Assert-True ($up.Contains('DO_MULTIPART_POST">true<')) 'multipart flag'
$title = Get-Arg $up 'title'
Assert-True ($null -ne $title -and $title.value -eq 'a b' -and $title.enc -eq 'false') 'multipart value encoded'

$xml = Find-Chunk $chunks 'echoXml'
Assert-True ($xml.Contains('&lt;a&gt;z&lt;/a&gt;')) 'xml body'
Assert-True ($xml.Contains('Header.value">application/xml<')) 'xml content type'
Assert-True ($xml.Contains('postBodyRaw">true<')) 'xml raw body'

$mx = Find-Chunk $chunks 'byColor'
Assert-True ((Get-Prop $mx 'HTTPSampler.path') -eq '/v1/matrix;color=blue') ('matrix path=' + (Get-Prop $mx 'HTTPSampler.path'))
$ak = Get-Arg $mx 'api_key'
Assert-True ($null -ne $ak -and $ak.value -eq 'XXXXXXX') 'operation apiKey query'
Assert-True (-not $mx.Contains('Header.name">X-API-KEY<')) 'operation security did not replace global'

$head = Find-Chunk $chunks 'headSku'
Assert-True ((Get-Prop $head 'HTTPSampler.method') -eq 'HEAD') 'HEAD method'
Assert-True ((Get-Prop $head 'HTTPSampler.path') -eq '/v1/label.A1') ('label path=' + (Get-Prop $head 'HTTPSampler.path'))
Assert-True (-not (Get-Prop $head 'HTTPSampler.path').Contains('?')) 'HEAD query embedded in path'
$q = Get-Arg $head 'q'
Assert-True ($null -ne $q -and $q.value -eq 'k v' -and $q.enc -eq 'true') 'HEAD query argument'

$opt = Find-Chunk $chunks 'optionalBody'
Assert-True ($opt.Contains('&quot;a&quot;: &quot;b&quot;')) 'optional body missing'

$oa2 = Invoke-Gen (Join-Path $root 'tests\fixtures\sample_api.json') (Join-Path $root 'tests\_out_oa_req') @('-Seed', '1', '-UseExamples', '-RequiredOnly')
$chunks2 = Get-SamplerChunks $oa2.text
$opt2 = Find-Chunk $chunks2 'optionalBody'
Assert-True (-not $opt2.Contains('postBodyRaw')) 'RequiredOnly kept optional body'
$up2 = Find-Chunk $chunks2 'upload'
Assert-True ($up2.Contains('DO_MULTIPART_POST">true<')) 'RequiredOnly dropped required body'

Write-Host '== PORT keeps original =='
$port = Invoke-Gen (Join-Path $root 'tests\fixtures\port_api.json') (Join-Path $root 'tests\_out_port') @('-c', (Join-Path $root 'tests\fixtures\port_host.csv'), '-Seed', '1')
$ping = (Get-SamplerChunks $port.text)[0]
$pingChunks = Get-SamplerChunks $port.text
Assert-True ($pingChunks.Count -eq 1) ('ping samplers=' + $pingChunks.Count)
$ping = $pingChunks[0]
Assert-True ((Get-Prop $ping 'HTTPSampler.domain') -eq '${HOST}') ('ping domain=' + (Get-Prop $ping 'HTTPSampler.domain'))
Assert-True ($port.text.Contains('${__P(HOST,api.example.com)}')) 'host override'
Assert-True ($port.text.Contains('${__P(PORT,8080)}')) 'empty PORT dropped 8080'
Assert-True ((Get-Prop $ping 'HTTPSampler.path') -eq '/v1/ping') ('ping path=' + (Get-Prop $ping 'HTTPSampler.path'))

if ($script:fail -eq 0) {
    Write-Host 'OK'
    exit 0
}
Write-Host ('FAILED ' + $script:fail)
exit 1

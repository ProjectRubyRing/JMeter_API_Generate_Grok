# Updates an existing JMX in place from CSV, Excel, and OpenAPI.
# powershell -NoProfile -ExecutionPolicy Bypass -File tests\check_update.ps1
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

function Get-Py {
    $sh = [System.IO.File]::ReadAllText((Join-Path $root 'jmeter_api_tool.sh'))
    $mark = '__PYTHON_CORE__'
    $a = $sh.IndexOf($mark)
    $b = $sh.LastIndexOf($mark)
    $lineEnd = $sh.IndexOf("`n", $a)
    return ($sh.Substring($lineEnd + 1, $b - ($lineEnd + 1)).Trim() + "`n")
}

function Get-NormText([string]$path) {
    $utf8 = New-Object System.Text.UTF8Encoding $false
    $text = [System.IO.File]::ReadAllText($path, $utf8)
    return [regex]::Replace($text, '\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}', 'DATE')
}

function Show-FirstDiff([string]$a, [string]$b, [string]$label) {
    $n = [Math]::Min($a.Length, $b.Length)
    for ($i = 0; $i -lt $n; $i++) {
        if ($a[$i] -ne $b[$i]) {
            $from = [Math]::Max(0, $i - 60)
            $lenA = [Math]::Min(120, $a.Length - $from)
            $lenB = [Math]::Min(120, $b.Length - $from)
            Write-Host ($label + ' diff at ' + $i)
            Write-Host $a.Substring($from, $lenA)
            Write-Host $b.Substring($from, $lenB)
            return
        }
    }
    Write-Host ($label + ' length ' + $a.Length + ' vs ' + $b.Length)
}

function Test-SameBytes([byte[]]$a, [byte[]]$b) {
    if ($a.Length -ne $b.Length) { return $false }
    for ($i = 0; $i -lt $a.Length; $i++) { if ($a[$i] -ne $b[$i]) { return $false } }
    return $true
}

function Add-UserElement([string]$jmxPath) {
    $enc = New-Object System.Text.UTF8Encoding $false
    $original = [System.IO.File]::ReadAllText($jmxPath)
    $marker = "        <Arguments guiclass=`"ArgumentsPanel`" testclass=`"Arguments`" testname=`"USER_KEEP_ME`" enabled=`"true`">`n          <stringProp name=`"Argument.value`">keep-me</stringProp>`n        </Arguments>`n        <hashTree/>`n"
    $needle = '<HTTPSamplerProxy guiclass="HttpTestSampleGui" testclass="HTTPSamplerProxy" testname="API'
    $inserted = $original.Replace($needle, $marker + $needle)
    if ($inserted -eq $original) { return $false }
    [System.IO.File]::WriteAllText($jmxPath, $inserted, $enc)
    return $true
}

function Get-OnlyJmx([string]$dir) {
    $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.jmx' -File)
    if ($files.Count -ne 1) { return '' }
    return $files[0].FullName
}

function Test-NoBak([string]$dir, [string]$label) {
    $baks = @(Get-ChildItem -LiteralPath $dir -Filter '*.bak' -File -ErrorAction SilentlyContinue)
    Assert-True ($baks.Count -eq 0) ($label + ' created a bak file')
    $jmx = @(Get-ChildItem -LiteralPath $dir -Filter '*.jmx' -File)
    Assert-True ($jmx.Count -eq 1) ($label + ' jmx count ' + $jmx.Count)
}

function Convert-Args([string]$engine, [string[]]$argList) {
    if ($engine -ceq 'py') { return $argList }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($a in $argList) {
        if ($a -ceq '--seed') { $out.Add('-Seed') }
        elseif ($a -ceq '--diff') { $out.Add('-Diff') }
        else { $out.Add($a) }
    }
    return ,$out.ToArray()
}
function Invoke-Engine([string]$engine, [string[]]$argList) {
    $argsFor = Convert-Args $engine $argList
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        if ($engine -ceq 'py') {
            $log = & $script:pyCmd -I $script:pyFile @argsFor 2>&1 | Out-String
        } else {
            $log = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool @argsFor 2>&1 | Out-String
        }
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prev
    }
    return @{ Code = $code; Log = [string]$log }
}

function Set-ZipText([string]$zipPath, [string]$old, [string]$new) {
    Add-Type -AssemblyName System.IO.Compression | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $tmp = $zipPath + '.rewritten'
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
    $inZip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    $outZip = [System.IO.Compression.ZipFile]::Open($tmp, [System.IO.Compression.ZipArchiveMode]::Create)
    $utf8 = New-Object System.Text.UTF8Encoding $false
    $hits = 0
    try {
        foreach ($entry in $inZip.Entries) {
            $dst = $outZip.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            if ($entry.FullName.EndsWith('/')) { continue }
            $src = $entry.Open()
            $ms = New-Object System.IO.MemoryStream
            $src.CopyTo($ms)
            $src.Dispose()
            $bytes = $ms.ToArray()
            $ms.Dispose()
            if ($entry.FullName.EndsWith('.xml') -or $entry.FullName.EndsWith('.rels')) {
                $text = $utf8.GetString($bytes)
                if ($text.Contains($old)) {
                    $hits += ([regex]::Matches($text, [regex]::Escape($old))).Count
                    $text = $text.Replace($old, $new)
                    $bytes = $utf8.GetBytes($text)
                }
            }
            $os = $dst.Open()
            $os.Write($bytes, 0, $bytes.Length)
            $os.Dispose()
        }
    } finally {
        $inZip.Dispose()
        $outZip.Dispose()
    }
    if ($hits -lt 1) {
        Remove-Item -LiteralPath $tmp -Force
        return $false
    }
    Move-Item -LiteralPath $tmp -Destination $zipPath -Force
    return $true
}

$pyCmd = $null
foreach ($cand in @('python', 'python3')) {
    $cmd = Get-Command $cand -ErrorAction SilentlyContinue
    if ($null -ne $cmd) { $pyCmd = $cmd.Source; break }
}
if (-not $pyCmd) { Write-Host 'python not found'; exit 1 }
$script:pyCmd = $pyCmd

$work = Join-Path $root 'tests\_update'
if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
New-Item -ItemType Directory -Path $work | Out-Null
$script:pyFile = Join-Path $work 'core.py'
$enc = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($script:pyFile, (Get-Py), $enc)
& $pyCmd -m py_compile $script:pyFile
if ($LASTEXITCODE -ne 0) { Write-Host 'compile failed'; exit 1 }

function Update-Pair([string]$caseName, [string]$inputFile, [string]$settings, [string[]]$extra, [scriptblock]$mutate, [string]$expect, [string]$stablePattern) {
    $case = Join-Path $work $caseName
    New-Item -ItemType Directory -Path $case | Out-Null
    $baseOut = Join-Path $case 'base'
    New-Item -ItemType Directory -Path $baseOut | Out-Null
    $genArgs = @('generate', '-i', $inputFile, '-o', $baseOut)
    if ($settings) { $genArgs += @('-s', $settings) }
    if ($extra) { $genArgs += $extra }
    $gen = Invoke-Engine 'py' $genArgs
    Assert-True ($gen.Code -eq 0) ($caseName + ' base generate failed: ' + $gen.Log)
    if ($gen.Code -ne 0) { return }
    $baseJmx = Get-OnlyJmx $baseOut
    Assert-True ([bool]$baseJmx) ($caseName + ' base jmx missing')
    if (-not $baseJmx) { return }
    Assert-True (Add-UserElement $baseJmx) ($caseName + ' failed to insert user element')
    if ($mutate) { & $mutate }
    $stableBefore = ''
    if ($stablePattern) {
        $stableBefore = [regex]::Match([System.IO.File]::ReadAllText($baseJmx), $stablePattern).Value
        Assert-True ([bool]$stableBefore) ($caseName + ' stable block missing')
    }
    foreach ($engine in @('py', 'ps')) {
        $out = Join-Path $case $engine
        Copy-Item -LiteralPath $baseOut -Destination $out -Recurse
        $jmx = Get-OnlyJmx $out
        $decoy = Join-Path $case ($engine + '-decoy')
        $args = @('generate', '-i', $inputFile, '-o', $decoy, '-t', $jmx)
        if ($settings) { $args += @('-s', $settings) }
        if ($extra) { $args += $extra }
        $beforeDiff = [System.IO.File]::ReadAllBytes($jmx)
        $diffArgs = @()
        foreach ($a in $args) { $diffArgs += $a }
        $diffArgs += '--diff'
        $diffRun = Invoke-Engine $engine $diffArgs
        Assert-True ($diffRun.Code -eq 0) ($caseName + ' ' + $engine + ' diff failed: ' + $diffRun.Log)
        $mid = [System.IO.File]::ReadAllBytes($jmx)
        Assert-True (Test-SameBytes $beforeDiff $mid) ($caseName + ' ' + $engine + ' diff rewrote the jmx')
        $ran = Invoke-Engine $engine $args
        Assert-True ($ran.Code -eq 0) ($caseName + ' ' + $engine + ' update failed: ' + $ran.Log)
        if ($ran.Code -ne 0) { continue }
        Assert-True (-not (Test-Path -LiteralPath $decoy)) ($caseName + ' ' + $engine + ' created a new output dir')
        $updated = [System.IO.File]::ReadAllText($jmx)
        Assert-True ($updated.Contains($expect)) ($caseName + ' ' + $engine + ' missing ' + $expect)
        Assert-True ($updated.Contains('USER_KEEP_ME')) ($caseName + ' ' + $engine + ' dropped user element')
        Assert-True ($updated.Contains('keep-me')) ($caseName + ' ' + $engine + ' dropped user value')
        if ($stableBefore) {
            $stableAfter = [regex]::Match($updated, $stablePattern).Value
            Assert-True ($stableAfter -eq $stableBefore) ($caseName + ' ' + $engine + ' rewrote an unchanged sampler')
        }
        $reportPath = $jmx.Substring(0, $jmx.Length - 4) + '_update.txt'
        Assert-True (Test-Path -LiteralPath $reportPath) ($caseName + ' ' + $engine + ' update txt missing')
        $utf8 = New-Object System.Text.UTF8Encoding $false
        $report = [System.IO.File]::ReadAllText($reportPath, $utf8)
        Assert-True ($report.Contains($expect)) ($caseName + ' ' + $engine + ' report missed ' + $expect)
        Assert-True ($ran.Log.Contains($expect)) ($caseName + ' ' + $engine + ' screen missed ' + $expect)
        $reportBytes = [System.IO.File]::ReadAllBytes($reportPath)
        $hasCr = $false
        foreach ($b in $reportBytes) { if ($b -eq 13) { $hasCr = $true; break } }
        Assert-True (-not $hasCr) ($caseName + ' ' + $engine + ' update txt contains CR')
        Test-NoBak $out ($caseName + ' ' + $engine)
        $again = [System.IO.File]::ReadAllBytes($jmx)
        $ran2 = Invoke-Engine $engine $args
        Assert-True ($ran2.Code -eq 0) ($caseName + ' ' + $engine + ' second update failed: ' + $ran2.Log)
        $report2 = [System.IO.File]::ReadAllText($reportPath, $utf8)
        Assert-True (-not $report2.Contains($expect)) ($caseName + ' ' + $engine + ' second report still has ' + $expect)
        $after = [System.IO.File]::ReadAllBytes($jmx)
        Assert-True (Test-SameBytes $again $after) ($caseName + ' ' + $engine + ' jmx changed when there was no diff')
        Test-NoBak $out ($caseName + ' ' + $engine + ' second')
    }
    $pyText = Get-NormText (Get-OnlyJmx (Join-Path $case 'py'))
    $psText = Get-NormText (Get-OnlyJmx (Join-Path $case 'ps'))
    if ($pyText -ne $psText) { Show-FirstDiff $pyText $psText $caseName }
    Assert-True ($pyText -eq $psText) ($caseName + ' ps and python jmx differ')
}

$csv = Join-Path $work 'api.csv'
Copy-Item (Join-Path $root 'input\api_definition.csv') $csv
Copy-Item (Join-Path $root 'input\api_definition_settings.csv') (Join-Path $work 'api_settings.csv')
$settings = Join-Path $work 'api_settings.csv'
Update-Pair 'csv' $csv $settings @() {
    $text = [System.IO.File]::ReadAllText($csv)
    $next = $text.Replace('https://postman-echo.com/post', 'https://postman-echo.com/changed')
    Assert-True ($next -ne $text) 'csv path was not changed'
    [System.IO.File]::WriteAllText($csv, $next, $enc)
} '/changed' '(?s)<HTTPSamplerProxy[^>]*testname="API02_[^"]*".*?</HTTPSamplerProxy>'

$xlsx = Join-Path $work 'api.xlsx'
Copy-Item (Join-Path $root 'input\api_definition.xlsx') $xlsx
Update-Pair 'xlsx' $xlsx '' @() {
    Assert-True (Set-ZipText $xlsx 'https://postman-echo.com/post' 'https://postman-echo.com/changed') 'xlsx did not contain the post url'
} '/changed' '(?s)<HTTPSamplerProxy[^>]*testname="API02_[^"]*".*?</HTTPSamplerProxy>'

$api = Join-Path $work 'sample_api.json'
Copy-Item (Join-Path $root 'tests\fixtures\sample_api.json') $api
Update-Pair 'openapi' $api '' @('--seed', '7') {
    $text = [System.IO.File]::ReadAllText($api)
    $next = $text.Replace('"/pets":', '"/pets-changed":')
    Assert-True ($next -ne $text) 'openapi path was not changed'
    [System.IO.File]::WriteAllText($api, $next, $enc)
} '/v1/pets-changed' '(?s)<HTTPSamplerProxy[^>]*testname="[^"]*showPet[^"]*".*?</HTTPSamplerProxy>'

$errJmx = Join-Path $work 'missing.jmx'
$err = Invoke-Engine 'py' @('generate', '-i', $csv, '-s', $settings, '-t', $errJmx)
Assert-True ($err.Code -ne 0) 'python missing jmx should fail'
$errPs = Invoke-Engine 'ps' @('generate', '-i', $csv, '-s', $settings, '-t', $errJmx)
Assert-True ($errPs.Code -ne 0) 'powershell missing jmx should fail'
$some = Get-OnlyJmx (Join-Path $work 'csv\py')
if ($some) {
    $beforeErr = [System.IO.File]::ReadAllBytes($some)
    $splitPy = Invoke-Engine 'py' @('generate', '-i', $csv, '-s', $settings, '-t', $some, '--split')
    Assert-True ($splitPy.Code -ne 0) 'python split+t should fail'
    $splitPs = Invoke-Engine 'ps' @('generate', '-i', $csv, '-s', $settings, '-t', $some, '-Split')
    Assert-True ($splitPs.Code -ne 0) 'powershell split+t should fail'
    $afterErr = [System.IO.File]::ReadAllBytes($some)
    Assert-True (Test-SameBytes $beforeErr $afterErr) 'failed split changed the jmx'
}

if ($script:fail -gt 0) { Write-Host ($script:fail.ToString() + ' failure(s)'); exit 1 }
Write-Host 'jmx update checks passed'
exit 0

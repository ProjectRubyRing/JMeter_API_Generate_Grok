# Checks the Linux JMeter CLI command text written beside a generated JMX.
# powershell -NoProfile -ExecutionPolicy Bypass -File tests\check_cli.ps1
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

$work = Join-Path $root 'tests\_cli'
if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
$psOut = Join-Path $work 'ps'
New-Item -ItemType Directory -Path $psOut | Out-Null
$log = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool generate -i (Join-Path $root 'input\api_definition.xlsx') -o $psOut 2>&1 | Out-String
Assert-True ($LASTEXITCODE -eq 0) ('generate failed: ' + $log)
$txtPath = Join-Path $psOut 'api_test_cli.txt'
Assert-True (Test-Path -LiteralPath $txtPath) 'cli text was not created'
$bytes = [System.IO.File]::ReadAllBytes($txtPath)
$hasCr = $false
foreach ($b in $bytes) { if ($b -eq 13) { $hasCr = $true; break } }
Assert-True (-not $hasCr) 'cli text contains CR'
$text = [System.IO.File]::ReadAllText($txtPath)
Assert-True ($text.Contains('JMeter CLI')) 'title missing'
Assert-True ($text.Contains('-n')) 'non-gui flag missing'
Assert-True ($text.Contains('-t "api_test.jmx"')) 'jmx argument missing'
Assert-True ($text.Contains('-e')) 'report flag missing'
Assert-True ($text.Contains('-o "$RUN_DIR/report"')) 'report directory missing'
Assert-True ($text.Contains('-JTHREADS=1')) 'threads property missing'
Assert-True ($text.Contains('-JRAMP_UP=1')) 'ramp property missing'
Assert-True ($text.Contains('-JLOOPS=1')) 'loops property missing'
Assert-True ($text.Contains('-JPROTOCOL=https')) 'protocol property missing'
Assert-True ($text.Contains('-JHOST=postman-echo.com')) 'host property missing'
Assert-True ($text.Contains('-JPORT=')) 'port property missing'
Assert-True ($text.Contains('-JCSV_FILE=data/testdata.csv')) 'csv property missing'
Assert-True ($text.Contains('-JEXCEL_FILE=data/testdata.xlsx')) 'excel property missing'
Assert-True ($text.Contains('-JCONNECT_TIMEOUT=10000')) 'connect timeout missing'
Assert-True ($text.Contains('-JRESPONSE_TIMEOUT=30000')) 'response timeout missing'
Assert-True ($text.Contains('-JENV=staging')) 'user variable ENV missing'
Assert-True ($text.Contains('-JCLIENT_ID=jmeter-client')) 'user variable CLIENT_ID missing'
Assert-True ($text.Contains('-JTHREADS=10')) 'load example threads missing'
Assert-True ($text.Contains('-JLOOPS=100')) 'load example loops missing'
Assert-True ($text.Contains('export HEAP="-Xms2g -Xmx2g -XX:MaxMetaspaceSize=256m"')) 'load example heap missing'
Assert-True ($text.Contains('export JAVA_HOME=/usr/lib/jvm/jre-17-openjdk')) 'java home missing'
Assert-True ($text.Contains('nohup env HEAP=')) 'background example missing'
Assert-True ($text.Contains('-g "$RUN_DIR/result.jtl"')) 'report regen missing'
Assert-True ($text.Contains('non-GUI') -or $text.Contains([string][char]0x975E + 'GUI')) 'parameter explanation missing'
Assert-True ($log.Contains('_cli.txt')) 'generate did not mention the cli text'

$py = $null
foreach ($cand in @('python', 'python3')) {
    $cmd = Get-Command $cand -ErrorAction SilentlyContinue
    if ($null -ne $cmd) { $py = $cmd.Source; break }
}
if ($py) {
    $sh = [System.IO.File]::ReadAllText((Join-Path $root 'jmeter_api_tool.sh'))
    $mark = '__PYTHON_CORE__'
    $a = $sh.IndexOf($mark)
    $b = $sh.LastIndexOf($mark)
    $lineEnd = $sh.IndexOf("`n", $a)
    $code = $sh.Substring($lineEnd + 1, $b - ($lineEnd + 1)).Trim() + "`n"
    $pyFile = Join-Path $work 'core.py'
    $enc = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($pyFile, $code, $enc)
    & $py -m py_compile $pyFile
    if ($LASTEXITCODE -ne 0) { Assert-True $false 'python core failed to compile'; exit 1 }
    $pyOut = Join-Path $work 'py'
    New-Item -ItemType Directory -Path $pyOut | Out-Null
    $pyLog = & $py -I $pyFile generate -i (Join-Path $root 'input\api_definition.xlsx') -o $pyOut 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 0) ('python generate failed: ' + $pyLog)
    $pyTxt = [System.IO.File]::ReadAllText((Join-Path $pyOut 'api_test_cli.txt'))
    if ($text -ne $pyTxt) {
        [System.IO.File]::WriteAllText((Join-Path $work 'ps.txt'), $text, $enc)
        [System.IO.File]::WriteAllText((Join-Path $work 'py.txt'), $pyTxt, $enc)
    }
    Assert-True ($text -eq $pyTxt) 'ps and python cli text differ (see tests/_cli/ps.txt)'
    $psJmx = Join-Path $psOut 'api_test.jmx'
    $pyJmx = Join-Path $pyOut 'api_test.jmx'
    $psCli = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool cli -t $psJmx 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 0) ('ps cli command failed: ' + $psCli)
    $pyCli = & $py -I $pyFile cli -t $pyJmx 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 0) ('python cli command failed: ' + $pyCli)
    $psAgain = [System.IO.File]::ReadAllText((Join-Path $psOut 'api_test_cli.txt'))
    $pyAgain = [System.IO.File]::ReadAllText((Join-Path $pyOut 'api_test_cli.txt'))
    Assert-True ($psAgain.Contains('nohup env HEAP=')) 'cli command dropped the background example'
    Assert-True ($psAgain.Contains('-JHOST=postman-echo.com')) 'cli command dropped the host'
    if ($psAgain -ne $pyAgain) {
        [System.IO.File]::WriteAllText((Join-Path $work 'ps-cli.txt'), $psAgain, $enc)
        [System.IO.File]::WriteAllText((Join-Path $work 'py-cli.txt'), $pyAgain, $enc)
    }
    Assert-True ($psAgain -eq $pyAgain) 'ps and python cli command output differ'
} else {
    Write-Host 'WARN: python not found; skipped parity'
}

if ($script:fail -gt 0) {
    Write-Host ($script:fail.ToString() + ' failure(s)')
    exit 1
}
Write-Host 'cli text checks passed'
exit 0

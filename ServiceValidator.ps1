$errors = 0

$recentError = $false

echo @"


[...] $(Get-Date) :: Initializing...  

"@

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if ( !($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) ) {
    Write-Warning "[.!.] $(Get-Date) :: Warning: This script is not being run in Administrator Mode. Some checks may not function correctly. Please close this program and re-open as a system administrator."
}

$doStartSessionTranscript = Read-Host "To get an output of this postinstall validator, specify a path like '.\log.txt' now:"

If ( !($doStartSessionTranscript -like $null) ) {
    Start-Transcript -Path $doStartSessionTranscript -Confirm -NoClobber
} else {
    Write-Host "[...] $(Get-Date) :: Session is not being transcribed."
} 

Start-Sleep -Seconds 1

Write-Host "[...] $(Get-Date) :: Getting running config"

Start-Sleep -Milliseconds 500

$cfgFileInterim = Read-Host "Specify Working Directory for sqlconnector. Default [C:\Users\Public\Downloads]:" 
If ( !($cfgFileInterim -like $null) ) {
    $ConfigFilePath = $cfgFileInterim
    Write-Host "Using $ConfigFilePath"
} else {
    $ConfigFilePath = "C:\Users\Public\Downloads"
}

$cfgDomainInterim = Read-Host "Specify your Fully-Qualified Domain Name. Test Domain is [cluster-3-ship.noble.niwc.navy.mil]:"
If ( !($cfgDomainInterim -like $null) ) {
    $ConfluentDomain = $cfgDomainInterim
    Write-Host "Using $ConfluentDomain"
} else {
    $ConfluentDomain = "cluster-3-ship.noble.niwc.navy.mil"
}

Write-Host "[...] $(Get-Date) :: Configuration Complete. Starting tests..."

Write-Host "[1/5] $(Get-Date) :: SQL Connector Launch Config Detected. Parsing configuration"

If (Test-Path "$ConfigFilePath\config.json") {
    Write-Host -foregroundcolor Green "[1/5] $(Get-Date) :: Config file installed at {$ConfigFilePath\}"
} else {
    Write-Host -foregroundcolor red "[1/5] $(Get-Date) :: Failed to find config file at $ConfigFilePath\config.json"
    $errors++
}

try {
    $config = Get-Content "$ConfigFilePath\config.json" | ConvertFrom-Json 
} catch {
    Write-Host -foregroundcolor red "[1/5] $(Get-Date) :: Config JSON failed validation"
    $errors++
    $recentError = $true
}

if (  !($recentError) ) {  Write-Host -foregroundcolor Green "[1/5] $(Get-Date) :: Config JSON passed validation" }
$recentError = $false

if (Test-Path "$ConfigFilePath\sqlconnector.ps1") {
    Write-Host -foregroundcolor Green "[1/5] $(Get-Date) :: Script installed at {$ConfigFilePath\}"
} else {
    Write-Host -foregroundcolor red "[!!!] $(Get-Date) :: Core Script not found in install location."
    $errors++
}

Write-Host -foregroundcolor white "[2/5] $(Get-Date) :: Processing Json..."

try {
    $config = Get-Content "$ConfigFilePath\config.json" | ConvertFrom-Json 
} catch {
    Write-Host -foregroundcolor red "[!!!] $(Get-Date) :: Failed to load json config. `n$_"
    $errors++
    $recentError = $true
}

if (  !($recentError) ) {  Write-Host -foregroundcolor Green "[1/5] $(Get-Date) :: Config JSON loaded successfully" }
$recentError = $false

Write-Host -foregroundcolor green "[2/5] $(Get-Date) :: Configuration file at {$ConfigFilePath\config.json} fully validated"

Write-Host -foregroundcolor white "[3/5] $(Get-Date) :: Collecting SFTP information, security keys, and testing login"

Start-Sleep -Seconds 1

try {
    $uploaddsk = $config.SFTPUploadDestination.split(":")[0]
    putty.exe -ssh $uploaddsk 30022 
} catch {
    Write-Host -foregroundcolor red "[!!!] $(Get-Date) :: Failed to connect to SFTP Destination `n$_"
    $errors++
    $recentError = $true
}

if (  !($recentError) ) {  Write-Host -foregroundcolor Green "[4/5] $(Get-Date) :: Successfully Connected to SFTP Destination" }
$recentError = $false

Write-Host -foregroundcolor white "[4/5] $(Get-Date) :: Connecting to host database..."

Start-Sleep -Seconds 1

Write-Host -foregroundcolor white "[4/5] $(Get-Date) :: Testing SMO Library..."
try { 
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null 
} catch {
     Write-Host -foregroundcolor red "[!!!] $(Get-Date) :: Unable to load SQL Server Management Studio library. `n $_"
     $errors++
     $recentError = $true
} 

if (  !($recentError) ) {  Write-Host -foregroundcolor Green "[4/5] $(Get-Date) :: Successfully tested SQL Server Management Studio automation library" }
$recentError = $false

Write-Host -foregroundcolor white "[4/5] $(Get-Date) :: Testing database connection..."
try { 
    New-Object ('Microsoft.SqlServer.Management.Smo.Server') "$($config.DatabaseConnectionString)" | Out-Null
} catch {
     Write-Host -foregroundcolor red "[!!!] $(Get-Date) :: Database connection failed. `n $_"
     $errors++
     $recentError = $true 
}

if (  !($recentError) ) {  Write-Host -foregroundcolor Green "[4/5] $(Get-Date) :: Successfully Connected to target configured database." }
$recentError = $false

Write-Host -foregroundcolor white "[4/5] $(Get-Date) :: Contacting SQL Agent..."

If ( (Get-Service -DisplayName "*SQL SERVER AGENT*").Status -eq "Running" ) {
    Write-Host -foregroundcolor green "[4/5] $(Get-Date) :: SQL Agent is alive and running"
} Else {
    Write-Host -foregroundcolor yellow "[4/5] $(Get-Date) :: SQL Agent does not appear to be properly configured"
}

Write-Host -foregroundcolor white "[4/5] $(Get-Date) :: Database configuration complete"

Start-Sleep -Milliseconds 300

Write-Host -foregroundcolor white "[5/5] $(Get-Date) :: Checking confluent access point..."

Start-Sleep -Seconds 1

$retryNum = 3 

for ($i = 0; $i -le $retryNum; $i++) {
    if (Test-Connection -ComputerName "controlcenter-confluent.apps.$ConfluentDomain" -Count 1 -Quiet) { 
        Write-Host -foregroundcolor green "[5/5] $(Get-Date) :: Confluent API Endpoint [controlcenter-confluent.apps.$ConfluentDomain] is most likely online and alive." 
        Break
    } elseif ($i -ne $retryNum)  {
        Write-Host -foregroundcolor yellow "[5/5] $(Get-Date) :: FAILED. Retrying $($i + 1)/$retryNum..."
    } else {
        Write-Host -foregroundcolor red "[5/5] $(Get-Date) :: Connection to Confluent API Endpoint [controlcenter-confluent.apps.$ConfluentDomain] has failed."
    }
}

if ($errors -eq 0) {
    Write-Host -foregroundcolor green "[...] $(Get-Date) :: Analysis: SQL Connector32 Integrator is Alive"
} else {
    Write-Host -foregroundcolor red "[...] $(Get-Date) :: Analysis: SQL Connector32 Integrator is Dead or Partially Alive"
}
Write-Host  "`n"

If ( !($doStartSessionTranscript -like $null) ) {
    Stop-Transcript
}

Read-Host "Press Enter to Exit:"

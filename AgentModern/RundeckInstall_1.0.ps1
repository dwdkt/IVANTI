$StrComputer = $env:COMPUTERNAME
$baseUrl = "https://ivanh0cms03.dkcorp.net/lan_retaildecathlon/ivanti_Agent2022_ModernVersion/"
$fileNames = @(
    "647a7ae7.0",
    "EPM_Manifest",
    "EPMAgentInstaller.exe",
    "WKS_CSH_CMS_2022.1.6_V1.0.txt"
)
$hostName = "xxxxxxxxxxx.dkcorp.net"
$ports = @(80, 443, 139, 445, 9593, 9594, 9595)
$destinationFolder = "C:\Windows\Temp\EBA"


# Check if curl.exe is available in C:\Windows\System32
$useCurlExe = Test-Path -Path "C:\Windows\System32\curl.exe"

if ($useCurlExe) {
    Write-Output "$StrComputer : Using curl.exe for file downloads"
} else {
    Write-Output "$StrComputer : curl.exe not found, defaulting to Invoke-WebRequest for file downloads"
}

# Test DNS resolution
try {
    $dnsResolution = Test-Connection -ComputerName $hostName -Count 1 -ErrorAction Stop
    Write-Output "$StrComputer : DNS resolution for $hostName succeeded. IP Address: $($dnsResolution.IPV4Address.IPAddressToString)"
} catch {
    Write-Output "$StrComputer : ERROR - DNS resolution for $hostName failed: $_"
    exit 1  # Exit the script if DNS resolution fails
}

# Test TCP port connectivity for specified ports
foreach ($port in $ports) {
    $tcpTest = Test-NetConnection -ComputerName $hostName -Port $port
    if ($tcpTest.TcpTestSucceeded) {
        Write-Output "$StrComputer : TCP connection to $hostName on port $port succeeded"
    } else {
        Write-Output "$StrComputer : ERROR - TCP connection to $hostName on port $port failed"
    }
}

$ServiceName = "Ivanti EPM Agent Update Service"
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($service) {
    Write-Output "$StrComputer : $ServiceName already exist"
    exit 0
}


# Delete the folder if it exists, then recreate it
if (Test-Path -Path $destinationFolder) {
    try {
        Remove-Item -Path $destinationFolder -Recurse -Force
        Write-Output "$StrComputer : Deleted existing folder $destinationFolder"
    } catch {
        Write-Output "$StrComputer : ERROR - Failed to delete $destinationFolder : $_"
    }
}

# Create the folder (if deletion failed, it will be recreated here)
if (!(Test-Path -Path $destinationFolder)) {
    New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
    Write-Output "$StrComputer : Created folder $destinationFolder"
}

# Download each file to the destination folder, overwriting if it exists
foreach ($fileName in $fileNames) {
    $fileUrl = $baseUrl + $fileName
    $destinationPath = Join-Path -Path $destinationFolder -ChildPath $fileName
    
    try {
        if ($useCurlExe) {
            # Use curl.exe to download the file
            & "C:\Windows\System32\curl.exe" -k -L -o "$destinationPath" "$fileUrl" > $null 2>&1
            Write-Output "$StrComputer : Downloaded $fileName to $destinationPath using curl.exe"
        } else {
            # Use Invoke-WebRequest as an alternative to curl.exe
            Invoke-WebRequest -Uri $fileUrl -OutFile $destinationPath -UseBasicParsing
            Write-Output "$StrComputer : Downloaded $fileName to $destinationPath using Invoke-WebRequest"
        }
    } catch {
        Write-Output "$StrComputer : ERROR - Failed to download $fileName from $fileUrl : $_"
    }
}

# List downloaded files and their sizes
Write-Output "$StrComputer : Listing downloaded files with sizes in $destinationFolder"
Get-ChildItem -Path $destinationFolder | ForEach-Object {
    Write-Output "$StrComputer : Downloaded File $destinationFolder : $($_.Name), Size: $($_.Length) bytes"
}

# Display the executable path before running it
$installerPath = Join-Path -Path $destinationFolder -ChildPath "EPMAgentInstaller.exe"
Write-Output "$StrComputer : Preparing to execute $installerPath"

# Check if the executable file exists before continuing
if (Test-Path -Path $installerPath) {
    # Use Start-Process with -PassThru to capture the exit code
    $process = Start-Process -FilePath $installerPath -Wait -WorkingDirectory $destinationFolder -PassThru

    # Retrieve the process exit code
    $exitCode = $process.ExitCode
    Write-Output "$StrComputer : Execution of $installerPath completed with exit code $exitCode"

    # Check if the exit code is non-zero (indicating an error)
    if ($exitCode -ne 0) {
        Write-Output "$StrComputer : ERROR - $installerPath failed with exit code $exitCode"
        
    } else {
        Write-Output "$StrComputer : $installerPath executed successfully."
    }
} else {
    Write-Output "$StrComputer : ERROR - Installer not found at $installerPath"
}

# Silent creation of the systools directory for NoStopService.log
$systemToolsDir = "$Env:SystemDrive\systools"
if (!(Test-Path -Path $systemToolsDir)) {
    New-Item -Path $systemToolsDir -ItemType Directory -Force | Out-Null
}

# Silent creation of the NoStopService.log file
$logFilePath = Join-Path -Path $systemToolsDir -ChildPath "NoStopService.log"
New-Item -Path $logFilePath -ItemType File -Force | Out-Null

# Monitor for the creation of the AgentUpdateService.log file
$monitorLogFilePath = "C:\ProgramData\Ivanti\EPM Agent\Logs\AgentUpdaterService.log"
$timeout1 = [datetime]::Now.AddMinutes(2)

Write-Output "$StrComputer : Waiting for $monitorLogFilePath"

while (-not (Test-Path -Path $monitorLogFilePath) -and ([datetime]::Now -lt $timeout1)) {
    Start-Sleep -Seconds 5
}

if (Test-Path -Path $monitorLogFilePath) {
    Write-Output "$StrComputer : $monitorLogFilePath found"
} else {
    # List files in the log directory if the log file is not found
    $logDirectory = Split-Path -Path $monitorLogFilePath -Parent
    Write-Output "$StrComputer : ERROR - $monitorLogFilePath not found within 2 minutes"
    Write-Output "$StrComputer : Listing files in $logDirectory"
    
    Get-ChildItem -Path $logDirectory | ForEach-Object {
        Write-Output "$StrComputer : $_"
    }
}

# Wait for the required "ivanti" services to start
$timeout2 = [datetime]::Now.AddMinutes(15)
$requiredServiceCount = 4

Write-Output "$StrComputer : Waiting for 4 services starting with 'ivanti'"

while (([datetime]::Now -lt $timeout2) -and ($foundServices -lt $requiredServiceCount)) {
    $services = Get-Service -DisplayName 'ivanti*' | Where-Object { $_.Status -eq 'Running' }
    $foundServices = $services.Count

    Write-Output "$StrComputer : Found $foundServices out of $requiredServiceCount 'ivanti' services running."

    Start-Sleep -Seconds 5
}

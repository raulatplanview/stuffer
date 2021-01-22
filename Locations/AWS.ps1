#Import-Module \\scripthost\modules\pvadmin
#Import-Module SQLSERVER

<# SETTING CREDENTIALS #>
Write-Host "Sign-in with your 'planviewcloud\aws-<admin>' account:" -ForegroundColor Magenta
$aAdmin = "a-$($admin)"
$awsAdmin = "aws-$($admin)"
$credentials = Get-Credential -Credential "planviewcloud\$($awsAdmin)"
$awsCredentials = "AWSEC2ReadCreds"

<# SETTING REGIONAL OPTIONS #>
switch($option) {
    1 {$jumpbox = "Jumpbox01fr.frankfurt.planviewcloud.net"; 
        $ad_server = "WIN-SDUR6J6Q8TH.frankfurt.planviewcloud.net"; 
        $dataCenterLocation = "fr"; 
        $awsRegion = "eu-central-1" 
        $reportFarm = "https://pbirsfarm01fr.pvcloud.com/reportserver"
        break}
    2 {$jumpbox = "Jumpbox01.sydney.planviewcloud.net"; 
        $ad_server = "WIN-O669CEBVH8N.sydney.planviewcloud.net"; 
        $dataCenterLocation = "au";
        $awsRegion = "ap-southeast-2" 
        $reportFarm = "https://pbirsfarm03au.pvcloud.com/reportserver"
        break}
}

<# FIND CUSTOMER RELATED VMs #>
Write-Host "Connecting to AWS and finding instances corralated with customer code ""$($customerCode.ToUpper())""" -ForegroundColor DarkGreen
$resourceIds = Get-EC2Tag -Region $awsRegion -Filter @{Name="tag:Cust_Id";Value="$($customerCode.ToUpper())"},@{Name="resource-type";Value="instance"} -ProfileName $awsCredentials
Write-Host "EC2 Instances found associated with customer code '$($customerCode.ToUpper())': $($resourceIds.Count)" -ForegroundColor Yellow

<# FILTER OUT NONACTIVE INSTANCES #>
$activeResourceIds = @()
foreach ($id in $resourceIds) {
    $instanceState = ((Get-EC2InstanceStatus -InstanceId $id.ResourceID -Region $awsRegion -ProfileName $awsCredentials).InstanceState | Select-Object -property Name).Name
    if ( "$($instanceState)" -eq "running"){
        $activeResourceIds += ($id)
    }
}
Write-Host "EC2 Instances actively running: $($activeResourceIds.Length)" -ForegroundColor Yellow
Write-Host "Starting production and sandbox environment data collection..." -ForegroundColor Red



<# CREATE MASTER SERVER ARRAY #>
$servers = @()
foreach ($id in $activeResourceIds){

    # PRODUCT METADATA: SERVER NAME, SERVER TYPE, CUSTOMER CODE, CUSTOMER NAME, CURRENT E1 VERSION, MAJOR VERSION, CUSTOMER URL, TIME ZONE, MAINTENANCE DAY # 
    $productMetadata = Get-EC2Tag -Region $awsRegion -Filter @{Name="resource-id";Value="$($id.ResourceID)"} -ProfileName $awsCredentials | 
        Where-Object {$_.Key -eq "Name" -or $_.Key -eq "Cust_Id" -or $_.Key -eq "Sub_Tier" -or $_.Key -eq "Cust_Name" -or 
        $_.Key -eq "Major" -or $_.Key -eq "CrVersion" -or $_.Key -eq "Maint_Window" -or $_.Key -eq "Cust_Url" -or $_.Key -eq "Tz_Id"} |
        Select-Object Key, Value
    $serverName = $productMetadata | Where-Object Key -eq "Name" | Select-Object Value
    Write-Host "Connected to $($serverName.Value)..." -ForegroundColor Green

    # INSTANCE METADATA (JSON): INSTANCE ID, INSTANCE TYPE, AVAILABILITY ZONE, LOCAL IP ADDRESS (IPV4), INSTANCE STATE. #
    $instanceMetadata = @()
    Write-Host "Gathering Instance Metadata..." -ForegroundColor Cyan
    
    $instanceMetadata = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {
        $metadataArray = @()

        $instanceId = Get-EC2InstanceMetadata -Category InstanceId
        $instanceType = Get-EC2InstanceMetadata -Category InstanceType
        $availabilityZone = Get-EC2InstanceMetadata -Category AvailabilityZone
        $localIpv4 = Get-EC2InstanceMetadata -Category LocalIpv4
        $instanceState = ((Get-EC2InstanceStatus -InstanceId $instanceId).InstanceState | Select-Object -property Name).Name

        $metadataArray = ($instanceId, $instanceState, $instanceType, $availabilityZone, $localIpv4)

        return $metadataArray
    }

    
    # HARDWARE METADATA: HDINFO, RAMINFO, CPUINFO #
    $hardwareMetadata = @()

        # HARDDRIVE LOGISTICS: DRIVE LETTER, SIZE (IN BYTES) #
        Write-Host "Gathering HD sizes..." -ForegroundColor Cyan
        $hdInfo = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {get-WmiObject win32_logicaldisk} | Select-Object DeviceID, Size

        # RAM LOGISTICS: PHYSICAL MEMORY ARRAY, SIZE (IN BYTES) #
        Write-Host "Gathering RAM sizes..." -ForegroundColor Cyan
        $ramInfo = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {Get-WmiObject win32_PhysicalMemoryArray} | Select-Object MaxCapacity

        # CPU LOGISTICS: PHYSICAL MEMORY ARRAY, SIZE (IN BYTES) #
        Write-Host "Gathering number of vCPUs..." -ForegroundColor Cyan
        $cpuInfo = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {Get-WmiObject Win32_Processor} | Select-Object DeviceID, Name

    $hardwareMetadata = ($hdinfo, $raminfo, $cpuInfo)
    
    # SCHEDULED TASKS: TASKS
    Write-Host "Gathering scheduled task information..." -ForegroundColor Cyan
    $tasks = Invoke-Command -computer $serverName.Value -ScriptBlock {
        Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | Where-Object TaskName -notlike "Op*" 
    } -Credential $credentials

    # POPULATE SERVER ARRAY #
    $servers += (,($productMetadata, $instanceMetadata, $hardwareMetadata, $tasks))

    Write-Host "================================================================================" -ForegroundColor Gray 
}

<# SORT MASTER SERVER ARRAY #>
$serverNames = @()
$productionComputers = @()
$sandboxComputers = @()
$undeclaredServers = @()
foreach ($server in $servers) {

    $environmentURL = ($server[0] | Where-Object Key -eq "Cust_Url" | Select-Object Value).Value
    $serverName = ($server[0] | Where-Object Key -eq "Name" | Select-Object Value).Value

    if ($null -eq $environmentURL) {
        $undeclaredServers += (,($server))
    }
    if ($environmentURL -like "*$($sandboxURLsuffix).*") {
        $serverNames += (,($serverName))
        $sandboxComputers += (,($server))
    }
    else {
        $serverNames += (,($serverName))
        $productionComputers += (,($server))
    }
    
}

Write-Host "Total number of active Production servers identified: $($productionComputers.Count)" -ForegroundColor yellow
Write-Host "Total number of active Sandbox servers identified: $($sandboxComputers.Count)" -ForegroundColor yellow
Write-Host "Total number of active non-Production or non-Sandbox servers identified: $($undeclaredServers.Count)" -ForegroundColor yellow

############################################################
# PRIMARY VARIABLES - IMOPRTANT JUNK YOU NEED TO FILL OUT
############################################################
Import-Module \\scripthost\modules\pvadmin
Import-Module SQLSERVER
Import-Module F5-LTM

$admin = "rreyes"
$domain = ".CORPORATE" # leave blank if your computer admin name does not have '.CORPORATE' or something of the such appended to it. (i.e. 'jsmith.CORPORATE')

Get-ChildItem -Path "C:\Users\$($admin)$($domain)\Planview, Inc\E1 Build Cutover - Documents\Customer Builds\1_FolderTemplate\18\" -Filter "InPlace*" | Copy-Item -Destination "C:\Users\$($admin)$($domain)\Desktop"
$stufferType = Read-Host "Where is the environment located: 1-AWS or 2-US/EU"

if ($stufferType -eq 1) {
############################################################
# PRE-REQUISITES:
# 
# AWS POWERSHELL TOOLS NEED TO BE INSTALLED,
# 
# AWS READ-ONLY CREDENTIALS NEED TO BE CONFIGURED AFTER
# INSTALLING AWS TOOLS.
############################################################

############################################################
# SECONDARY VARIABLES - LESSER IMPORTANT JUNK 
############################################################
Write-Host "Sign-in with your 'planviewcloud\aws-<admin>' account:" -ForegroundColor Magenta
$aAdmin = "a-$($admin)"
$awsAdmin = "aws-$($admin)"
$credentials = Get-Credential -Credential "planviewcloud\$($awsAdmin)"
$awsCredentials = "AWSEC2ReadCreds"

$customerCode = Read-Host "Enter the customer code"
$option = Read-Host "Select a region: 1-Frankfurt or 2-Sydney"

$excel_file = Get-ChildItem -Path "C:\Users\$($admin)$($domain)\Desktop\" -Filter "InPlace*"

$productionDatabase = "$($customerCode.ToUpper())PROD"
$ctmDatabase = "$($customerCode.ToUpper())CTM"
$sandboxDatabase = "$($customerCode.ToUpper())SANDBOX1"
$excelFilePath = ("C:\Users\$($admin)$($domain)\Desktop\$excel_file")
$sandboxURLsuffix = "-sb"

############################################################
# DEFINES THE JUMPBOX, ACTIVE DIRECTORY, VSHPERE SERVERS
# BASED ON THE DATACENTER LOCATION
############################################################
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


############################################################
# SEARCH AWS FOR CUSTOMER RELATED BOXES
############################################################
Write-Host "Connecting to AWS and finding instances corralated with customer code ""$($customerCode.ToUpper())""" -ForegroundColor DarkGreen
$resourceIds = Get-EC2Tag -Region $awsRegion -Filter @{Name="tag:Cust_Id";Value="$($customerCode.ToUpper())"},@{Name="resource-type";Value="instance"} -ProfileName $awsCredentials
Write-Host "EC2 Instances found associated with customer code '$($customerCode.ToUpper())': $($resourceIds.Count)" -ForegroundColor Yellow

############################################################
# WEEDS OUT INSTANCES THAT ARE NOT RUNNING
############################################################
$activeResourceIds = @()

foreach ($id in $resourceIds) {

    $instanceState = ((Get-EC2InstanceStatus -InstanceId $id.ResourceID -Region $awsRegion -ProfileName $awsCredentials).InstanceState | Select-Object -property Name).Name

    if ( "$($instanceState)" -eq "running"){
        $activeResourceIds += ($id)
    }
}
Write-Host "EC2 Instances actively running: $($activeResourceIds.Length)" -ForegroundColor Yellow
Write-Host "Starting production and sandbox environment data collection..." -ForegroundColor Red
############################################################
# DATA COLLECTION FOR RUNNING INSTANCES
############################################################
$servers = @()
foreach ($id in $activeResourceIds){

    ############################################################
    # PRODUCT METADATA:
    # SERVER NAME, SERVER TYPE, CUSTOMER CODE, CUSTOMER NAME,
    # CURRENT E1 VERSION, MAJOR VERSION, CUSTOMER URL,
    # TIME ZONE, MAINTENANCE DAY.  
    ############################################################
    $productMetadata = Get-EC2Tag -Region $awsRegion -Filter @{Name="resource-id";Value="$($id.ResourceID)"} -ProfileName $awsCredentials | 
        Where-Object {$_.Key -eq "Name" -or $_.Key -eq "Cust_Id" -or $_.Key -eq "Sub_Tier" -or $_.Key -eq "Cust_Name" -or 
        $_.Key -eq "Major" -or $_.Key -eq "CrVersion" -or $_.Key -eq "Maint_Window" -or $_.Key -eq "Cust_Url" -or $_.Key -eq "Tz_Id"} |
        Select-Object Key, Value
    $serverName = $productMetadata | Where-Object Key -eq "Name" | Select-Object Value
    Write-Host "Connected to $($serverName.Value)..." -ForegroundColor Green
#    $productMetadata.Value

    ############################################################
    # INSTANCE METADATA (JSON):
    # INSTANCE ID, INSTANCE TYPE, AVAILABILITY ZONE, 
    # LOCAL IP ADDRESS (IPV4), INSTANCE STATE.
    ############################################################  
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

#    foreach ($x in $instanceMetadata) {
#        Write-Host $x
#    }

    ############################################################
    # GETTING HARDWARE METADATA
    # HDINFO, RAMINFO, CPUINFO
    ############################################################
    $hardwareMetadata = @()

        # HARDDRIVE LOGISTICS:
        # DRIVE LETTER, SIZE (IN BYTES)
        Write-Host "Gathering HD sizes..." -ForegroundColor Cyan
        $hdInfo = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {get-WmiObject win32_logicaldisk} | Select-Object DeviceID, Size
#        foreach ($hd in $hdInfo){
#            Write-Host """$($hd.DeviceID)"" --- $($hd.Size / 1073741824)gb"       
#        }

        # RAM LOGISTICS:
        # PHYSICAL MEMORY ARRAY, SIZE (IN BYTES)
        Write-Host "Gathering RAM sizes..." -ForegroundColor Cyan
        $ramInfo = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {Get-WmiObject win32_PhysicalMemoryArray} | Select-Object MaxCapacity
#        Write-Host "RAM: $($ramInfo.MaxCapacity / 1048576)gb"

        # CPU LOGISTICS:
        # PHYSICAL MEMORY ARRAY, SIZE (IN BYTES)
        Write-Host "Gathering number of vCPUs..." -ForegroundColor Cyan
        $cpuInfo = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {Get-WmiObject Win32_Processor} | Select-Object DeviceID, Name
#        Write-Host "Total vCPUs found: $($cpuInfo.DeviceID.Count)"
#        foreach ($cpu in $cpuInfo){
#            Write-Host """$($cpu.DeviceID)"" --- $($cpu.Name)gb"       
#        }

    $hardwareMetadata = ($hdinfo, $raminfo, $cpuInfo)

    ############################################################
    # GETTING SCHEDULED TASKS
    # TASKS
    ############################################################
    Write-Host "Gathering scheduled task information..." -ForegroundColor Cyan
    $tasks = Invoke-Command -computer $serverName.Value -ScriptBlock {
        Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | where TaskName -notlike "Op*" 
    } -Credential $credentials
#    foreach ($task in $tasks) {
#        Write-Host "Scheduled Task: $($task.TaskName)"
#    }



    $servers += (,($productMetadata, $instanceMetadata, $hardwareMetadata, $tasks))

    Write-Host "================================================================================" -ForegroundColor Gray 
}

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

$serverNames | Out-File -FilePath "C:\Users\$($admin)$($domain)\Documents\RDCM Files\New Servers.txt" -Encoding UTF8
Write-Host ":::::::: New Servers Available for Importing in RDCM ::::::::" -ForegroundColor DarkGreen

<##################################################################################################>
<##################################################################################################>
##
##  INITIALIZING EXCEL DOCUMENT
##
<##################################################################################################>
<##################################################################################################>

############################################################
# APPLICATION LAYER                                
# INSTANCIATES EXCEL IN PS AND OPENS THE EXCEL FILE 
############################################################
$excel = New-Object -ComObject Excel.Application
$excelfile = $excel.Workbooks.Open($excelFilePath)

############################################################
# WORKSHEET LAYER                                
# CALLS A WORKSHEET FROM THE EXCEL FILE 
############################################################
# EXAMPLE #
# $buildData = $excelfile.sheets.item("Jenkins Inputs")
$buildData = $excelfile.sheets.item("MasterConfig")

############################################################
# RANGE LAYER - (FOR READING THE EXCEL FILE)
# -HASH TABLES (KEY/VALUE PAIR FORMAT)
# -ISSUE A VARIABLE, THEN ASSIGN IT A CELL VALUE.
############################################################
# EXAMPLE #
# $data_JenkinsInputs = @{
# "<us_all_pipe>" = $buildData.RANGE("C2").Text  
# "<target_server_name>" = $buildData.RANGE("C3").Text
# }

############################################################
# RANGE LAYER - (FOR WRITING TO THE EXCEL FILE)
# TARGET A CELL ($ROW, $COLUMN) AND ASSIGN IT A VALUE
############################################################
# EXAMPLE #
# $buildData.Cells.Item(2,5)= 'Hello'
#AWS BUILD
$buildData.Cells.Item(18,2)= "True" 
#SPLIT TIER
$buildData.Cells.Item(19,2)= "False"

$buildData.Cells.Item(23,2)= $productionComputers.Count
$buildData.Cells.Item(23,3)= $sandboxComputers.Count
$buildData.Cells.Item(9,2)= "$($dataCenterLocation)"


############################################################
# ITERATES THROUGH PRODUCTION AND SANDBOX SERVERS 
# PLACES DATA IN EXCEL SHEET CELLS
#
# '$servers' ARRAY: 
# 0: $productMetadata (SERVER NAME, SERVER TYPE, CUSTOMER CODE,
#       CUSTOMER NAME, CURRENT E1 VERSION, MAJOR VERSION, 
#       CUSTOMER URL, TIME ZONE, MAINTENANCE DAY)
# 
# 1: $instanceMetadata (ARRAY --> $instanceId, $instanceState, 
#       $instanceType, $availabilityZone, $localIpv4)
# 
# 2: $hardwareMetadata (ARRAY --> $hdinfo, $raminfo, $cpuInfo) 
#
# 3: $tasks
############################################################


#######################
# PRODUCTION SERVERS
#######################
Write-Host ":::::::: PRODUCTION ENVIRONMENT ::::::::" -ForegroundColor Yellow

$webServerCount = 0
foreach ( $server in $productionComputers){

    $serverName = $server[0] | Where-Object Key -eq "Name" | Select-Object Value

    ##########################
    # PRODUCTION APP SERVER 
    ##########################
    if ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "app") {
        Write-Host "THIS IS THE PRODUCTION APP SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# OPEN SUITE #>
        Write-Host "OpenSuite" -ForegroundColor Red
        $opensuite = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {
            if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {

                $software = "*Actian*";
                $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -like $software }) -ne $null

                if ($installed) {
                    return "Yes"
                }
                
            } else {
                return "No"
            }
        }
        Write-Host "OpenSuite Detected: $($opensuite)"

        <# INTEGRATIONS #>
        Write-Host "Integrations" -ForegroundColor Red
        $PPAdapter = "False"
        $LKAdapter = "False"
        $integrations = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {
            param ($database)
            if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*") {
                Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*" | Select-Object -Property PSChildName
            } else {
                return 0
            }
        } -ArgumentList $productionDatabase

        if ($integrations -eq 0) {
            Write-Host "No integrations found in 'HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($productionDatabase)\'"
        } else {
            Write-Host "Number of integrations found: $($integrations.PSChildName.Count)" -ForegroundColor Cyan
            foreach ($x in $integrations.PSChildName) {
                if ($x -like "*ProjectPlace*") {
                    Write-Host "PP ADAPTER FOUND: $($x)"
                    $PPAdapter = "True"
                }
                elseif ($x -like "*PRM_Adapter*") {
                    Write-Host "LK ADAPTER FOUND: $($x)"
                    $LKAdapter= "True"
                } else {
                    Write-Host "Other Integration Identified: $($x)"
                }
            }
        }

        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress        

        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(52,2)= "$($serverName.Value)"
        $buildData.Cells.Item(52,3)= "$($server[2][2].DeviceID.Count)"
        $buildData.Cells.Item(52,4)= "$($server[2][1].MaxCapacity / 1048576)"
        $buildData.Cells.Item(52,5)= $hdStringArray
        $buildData.Cells.Item(52,6)= $diskResize
        $buildData.Cells.Item(52,7)= $task_array
        $buildData.Cells.Item(52,8)= $instanceSize
        $buildData.Cells.Item(52,9)= $AvailabilityZone
        $buildData.Cells.Item(52,10)= $ipAddress
        $buildData.Cells.Item(52,11)= $instanceID

        $buildData.Cells.Item(31,2)= $PPAdapter
        $buildData.Cells.Item(32,2)= $LKAdapter

        $buildData.Cells.Item(36,2)= $opensuite

        Write-Host "================================================================================" -ForegroundColor Red
    }

    ##################################
    # PRODUCTION CTM SERVER (Troux) 
    ##################################
    if ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "ctm") {
        Write-Host "THIS IS THE PRODUCTION TROUX SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"    

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"      

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress        

        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(53,2)= "$($serverName.Value)"
        $buildData.Cells.Item(53,3)= "$($server[2][2].DeviceID.Count)"
        $buildData.Cells.Item(53,4)= "$($server[2][1].MaxCapacity / 1048576)"
        $buildData.Cells.Item(53,5)= $hdStringArray
        $buildData.Cells.Item(53,6)= $diskResize
        $buildData.Cells.Item(53,7)= $task_array
        $buildData.Cells.Item(53,8)= $instanceSize
        $buildData.Cells.Item(53,9)= $AvailabilityZone
        $buildData.Cells.Item(53,10)= $ipAddress
        $buildData.Cells.Item(53,11)= $instanceID

        Write-Host "================================================================================" -ForegroundColor Red  
    }
   
    ##########################
    # PRODUCTION WEB SERVER 
    ##########################
    elseif ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "web") {

        Write-Host "THIS IS THE PRODUCTION WEB SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)" 

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# CURRENT VERSION #>
        Write-Host "Current Environment Version" -ForegroundColor Red
        $crVersion = ($server[0] | Where-Object Key -eq "CrVersion" | Select-Object Value).Value
        $crVersion

        <# MAJOR VERSION #>
        Write-Host "Major Version" -ForegroundColor Red
        $majorVersion = ($server[0] | Where-Object Key -eq "Major" | Select-Object Value).Value
        $majorVersion

        # NEW RELIC #
        Write-Host "New Relic" -ForegroundColor Red
        $newRelic = Invoke-Command -ComputerName "$($serverName.Value)" -Credential $credentials -ScriptBlock {
            if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                Write-Host "New Relic has been detected on this server"
                return "Yes"
            } else {
                Write-Host "New Relic was not detected on this server"
                return "No"
            }
        }

            # GET WEB CONFIG #
            $webConfig = Invoke-Command -ComputerName "$($serverName.Value)" -Credential $credentials -ScriptBlock {
                return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
            }
            $webConfig = [xml] $webConfig

        <# PRODUCTION URL #>
        Write-Host "Production URL" -ForegroundColor Red
        $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
        $metadataURL = ($server[0] | Where-Object Key -eq "Cust_Url" | Select-Object Value).Value
        Write-Host "Instance Metadata: $($metadataURL)"
        Write-Host "Web config file: $($environmentURL.value)"

        <# DNS ALIAS #>
        Write-Host "Production DNS Alias" -ForegroundColor Red            
        $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
        $dnsAlias

        <# REPORT FARM URL #>
        Write-Host "Report Farm URL" -ForegroundColor Red
        $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
        Write-Host $reportfarmURL.value

        <# ENCRYPTED PVMASTER PASSWORD #>
        Write-Host "Encrypted PVMaster Password" -ForegroundColor Red
        $encryptedPVMasterPassword = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUserPassword"} | Select-Object -Property value
        Write-Host $encryptedPVMasterPassword.value
        
        <# UNENCRYPTED PVMASTER PASSWORD #>
        Write-Host "Unencrypted PVMaster Password" -ForegroundColor Red
        $unencryptedPVMasterPassword = Invoke-PassUtil -InputString $encryptedPVMasterPassword.value -Deobfuscation
        Write-Host $unencryptedPVMasterPassword

        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress        

        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(25,2)= $crVersion
        $buildData.Cells.Item(13,2)= $majorVersion
        $buildData.Cells.Item(17,2)= $encryptedPVMasterPassword.value
        $buildData.Cells.Item(16,2)= $unencryptedPVMasterPassword
        $buildData.Cells.Item(1,2)= $environmentURL.value
        $buildData.Cells.Item(7,2)= $dnsAlias
        $buildData.Cells.Item(46,2)= $reportfarmURL.value
        $buildData.Cells.Item(15,2)= $newRelic

        if ($webServerCount -gt 0){
            $buildData.Cells.Item(58 + ($webServerCount - 1),2)= "$($serverName.Value)"
            $buildData.Cells.Item(58 + ($webServerCount - 1),3)= "$($server[2][2].DeviceID.Count)"
            $buildData.Cells.Item(58 + ($webServerCount - 1),4)= "$($server[2][1].MaxCapacity / 1048576)"
            $buildData.Cells.Item(58 + ($webServerCount - 1),5)= $hdStringArray
            $buildData.Cells.Item(58 + ($webServerCount - 1),6)= $diskResize
            $buildData.Cells.Item(58 + ($webServerCount - 1),7)= $task_array
            $buildData.Cells.Item(58 + ($webServerCount - 1),8)= $instanceSize
            $buildData.Cells.Item(58 + ($webServerCount - 1),9)= $AvailabilityZone
            $buildData.Cells.Item(58 + ($webServerCount - 1),10)= $ipAddress
            $buildData.Cells.Item(58 + ($webServerCount - 1),11)= $instanceID

        }
        else {
            $buildData.Cells.Item(51,2)= "$($serverName.Value)"
            $buildData.Cells.Item(51,3)= "$($server[2][2].DeviceID.Count)"
            $buildData.Cells.Item(51,4)= "$($server[2][1].MaxCapacity / 1048576)"
            $buildData.Cells.Item(51,5)= $hdStringArray
            $buildData.Cells.Item(51,6)= $diskResize
            $buildData.Cells.Item(51,7)= $task_array
            $buildData.Cells.Item(51,8)= $instanceSize
            $buildData.Cells.Item(51,9)= $AvailabilityZone
            $buildData.Cells.Item(51,10)= $ipAddress
            $buildData.Cells.Item(51,11)= $instanceID
            
        }
        
        $webServerCount++
        $buildData.Cells.Item(24,2)= $webServerCount

        
        Write-Host "================================================================================" -ForegroundColor Red  

   }

    ##########################
    # PRODUCTION SAS SERVER 
    ##########################
    elseif ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "sas") {
        Write-Host "THIS IS THE PRODUCTION SAS SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress        
        
        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(55,2)= "$($serverName.Value)"
        $buildData.Cells.Item(55,3)= "$($server[2][2].DeviceID.Count)"
        $buildData.Cells.Item(55,4)= "$($server[2][1].MaxCapacity / 1048576)"
        $buildData.Cells.Item(55,5)= $hdStringArray
        $buildData.Cells.Item(55,6)= $diskResize
        $buildData.Cells.Item(55,7)= $task_array
        $buildData.Cells.Item(55,8)= $instanceSize
        $buildData.Cells.Item(55,9)= $AvailabilityZone
        $buildData.Cells.Item(55,10)= $ipAddress
        $buildData.Cells.Item(55,11)= $instanceID

        Write-Host "================================================================================" -ForegroundColor Red  
   }

    ##########################
    # PRODUCTION SQL SERVER 
    ##########################
    elseif ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "sql") {
        Write-Host "THIS IS THE PRODUCTION SQL SERVER" -ForegroundColor Cyan
        
        <# OU NAME #>
        Write-Host "OU Name" -ForegroundColor Red
        $ouName = ($server[0] | Where-Object Key -eq "Cust_Name" | Select-Object Value).Value
        Write-Host $ouName

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# DATABASE PROPERTIES #>
        Write-Host "$($productionDatabase) Properties" -ForegroundColor Red
        $sqlSession = New-PSSession -ComputerName "$($serverName.Value)" -Credential $credentials

            # MAXDOP/THRESHOLD
            Write-Host "Identifying MaxDOP/Threshold..." -ForegroundColor Cyan
            $database_maxdop_threshold = Invoke-Command  -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                '%parallel%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
            } -ArgumentList $serverName.Value
            $maxdop = $database_maxdop_threshold | Where-Object {$_.name -like "cost*"} | Select-Object -property value
            $cost_threshold = $database_maxdop_threshold | Where-Object {$_.name -like "max*"} | Select-Object -property value
            Write-Host "Max DOP --- $($maxdop.value) MB"
            Write-Host "Cost Threshold --- $($cost_threshold.value) MB"           

            # MIN/MAX MEMORY
            Write-Host "Identifying MIN/MAX Memory..." -ForegroundColor Cyan
            $database_memory = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                '%server memory%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
            } -ArgumentList $serverName.Value 
            $database_memory_max = $database_memory | where-Object {$_.name -like "max*"} | Select-Object -property value
            $database_memory_min = $database_memory | where-Object {$_.name -like "min*"} | Select-Object -property value
            Write-Host "Max Server Memory --- $($database_memory_max.value) MB"
            Write-Host "Min Server Memory --- $($database_memory_min.value) MB"

            # DATABASE ENCRYPTION
            Write-Host "Identifying Database Encryption..." -ForegroundColor Cyan
            $database_encryption = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT
                db.name,
                db.is_encrypted
                FROM
                sys.databases db
                LEFT OUTER JOIN sys.dm_database_encryption_keys dm
                    ON db.database_id = dm.database_id;
                GO" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value
            $dbEncryption = $database_encryption | Where-Object {$_.name -eq $productionDatabase}
            Write-Host "$($dbEncryption.name) --- $($dbEncryption.is_encrypted)"
            
            # DATABASE SIZE (MAIN)
            Write-Host "Calculating Database Size" -ForegroundColor Cyan
            $database_dbSize = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database)
                GO
                exec sp_spaceused
                GO" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$productionDatabase
            Write-Host "$($database_dbSize.database_name) --- $($database_dbSize.database_size)"

            # ALL DATABASES (NAMES AND SIZES in MB)
            Write-Host "Listing All Databases and Sizes (in MB)" -ForegroundColor Cyan
            $all_databases = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "SELECT d.name,
                ROUND(SUM(mf.size) * 8 / 1024, 0) Size_MB
                FROM sys.master_files mf
                INNER JOIN sys.databases d ON d.database_id = mf.database_id
                WHERE d.database_id > 4 -- Skip system databases
                GROUP BY d.name
                ORDER BY d.name" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$productionDatabase
            foreach ($database in $all_databases) {
                Write-Host "$($database.name) ---- $($database.Size_MB) MB"
            }

            # CUSTOM MODELS
            Write-Host "Calculating Custom Models..." -ForegroundColor Cyan
            $database_custom_models = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT * FROM ip.olap_properties 
                WHERE bism_ind ='N' 
                AND olap_obj_name 
                NOT like 'PVE%'" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$productionDatabase | Select-Object -property olap_obj_name
            foreach ($model in $database_custom_models.olap_obj_name) {
                Write-Host $model
            }  
            
            # INTERFACES
            Write-Host "Identifying Interfaces..." -ForegroundColor Cyan
            $database_interfaces = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT
                s.description JobStreamName,
                j.description JobName,
                j.job_order JobOrder,
                j.job_id JobID,
                p.name ParamName,
                p.param_value ParamValue,
                MIN(r.last_started) JobLastStarted,
                MAX(r.last_finished) JobLastFinished,
                MAX(CONVERT(CHAR(8), DATEADD(S,DATEDIFF(S,r.last_started,r.last_finished),'1900-1-1'),8)) Duration
                FROM ip.job_stream_job j
                INNER JOIN ip.job_stream s
                ON j.job_stream_id = s.job_stream_id
                INNER JOIN ip.job_stream_schedule ss
                ON ss.job_stream_id = s.job_stream_id
                INNER JOIN ip.job_run_status r
                ON s.job_stream_id = r.job_stream_id
                LEFT JOIN ip.job_param p
                ON j.job_id = p.job_id
                WHERE P.Name = 'Command'
                GROUP BY
                s.description,
                j.description,
                j.job_order,
                j.job_id,
                p.name,
                p.param_value;" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$productionDatabase
            $database_interfaces.ParamValue

            # LICENSE COUNT
            Write-Host "Calculating License Count..." -ForegroundColor Cyan
            $database_license_count = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT
                LicenseRole,
                COUNT(UserName) UserCount,
                r.seats LicenseCount
                FROM (
                SELECT
                s1.description LicenseRole,
                s1.structure_code LicenseCode,
                u.active_ind Active,
                u.full_name UserName
                FROM ip.ip_user u
                INNER JOIN ip.structure s
                ON u.role_code = s.structure_code
                INNER JOIN ip.structure s1
                ON s.father_code = s1.structure_code
                WHERE u.active_ind = 'Y'
                ) q
                INNER JOIN ip.ip_role r
                ON q.LicenseCode = r.role_code
                GROUP BY
                LicenseRole,
                LicenseCode,
                r.seats" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$productionDatabase
            $licenseProperties = $database_license_count | Select-Object -Property LicenseRole,LicenseCount
            $totalLicenseCount = 0
            foreach ($license in $licenseProperties){
                Write-Output "$($license.LicenseRole): $($license.LicenseCount)"
                $totalLicenseCount += $license.LicenseCount
            }
            Write-Output "Total License Count: $($totalLicenseCount)"
            
            # PROGRESSING WEB VERSION
            Write-Host "Identifying Progressing Web Version..." -ForegroundColor Cyan
            $database_progressing_web_version = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database); SELECT TOP 1 sub_release 
                FROM ip.pv_version 
                WHERE release = 'PROGRESSING_WEB'
                ORDER BY seq DESC;" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$productionDatabase
            $database_progressing_web_version.sub_release
            
        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress            

        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(11,2)= $serverName.Value.Substring(($serverName.Value.Length - 2), 2)
        $buildData.Cells.Item(44,2)= $database_dbSize.database_size
        $buildData.Cells.Item(43,2)= $database_memory_max.value
        $buildData.Cells.Item(42,2)= $database_memory_min.value

        $buildData.Cells.Item(54,2)= "$($serverName.Value)"
        $buildData.Cells.Item(54,3)= "$($server[2][2].DeviceID.Count)"
        $buildData.Cells.Item(54,4)= "$($server[2][1].MaxCapacity / 1048576)"
        $buildData.Cells.Item(54,5)= $hdStringArray
        $buildData.Cells.Item(54,6)= $diskResize
        $buildData.Cells.Item(54,7)= $task_array
        $buildData.Cells.Item(54,8)= $instanceSize
        $buildData.Cells.Item(54,9)= $AvailabilityZone
        $buildData.Cells.Item(54,10)= $ipAddress
        $buildData.Cells.Item(54,11)= $instanceID

        $buildData.Cells.Item(26,2)= $database_progressing_web_version.sub_release
        
        $buildData.Cells.Item(28,2)= $database_custom_models.Count
        $modelCount = 0;
        foreach ($model in $database_custom_models.olap_obj_name){
            $buildData.Cells.Item(91, (2 + $modelCount))= $model
            $modelCount++
        }

        $databaseCount = 0
        foreach ($database in $all_databases) {
            $buildData.Cells.Item(99, (2 + $databaseCount))= $database.name
            $buildData.Cells.Item(100, (2 + $databaseCount))= "Size: $($database.Size_MB) MB"
            $databaseCount++
        }

        $buildData.Cells.Item(30,2)= $database_interfaces.ParamValue.Count
        $interfaceCount = 0
        foreach ($interface in $database_interfaces.ParamValue) {
            $buildData.Cells.Item(95, (2 + $interfaceCount))= $interface
            $interfaceCount++
        }

        $buildData.Cells.Item(41,2)= $dbEncryption.is_encrypted
        $buildData.Cells.Item(22,2)= $totalLicenseCount

        $buildData.Cells.Item(14,2)= $ouName 
        $buildData.Cells.Item(10,2)= $customerCode.ToUpper() 
        $buildData.Cells.Item(3,2)= "N/A --- AWS Build"
        $buildData.Cells.Item(40,2)= $cost_threshold.value         
        $buildData.Cells.Item(39,2)= $maxdop.value

        Remove-PSSession -Session $sqlSession
            
        Write-Host "================================================================================" -ForegroundColor Red  
   }

    ##########################
    # PRODUCTION PVE SERVER 
    ##########################
    elseif ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "pve") {
        Write-Host "THIS IS THE PRODUCTION PVE SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"
            
        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# CURRENT VERSION #>
        Write-Host "Current Environment Version" -ForegroundColor Red
        $crVersion = ($server[0] | Where-Object Key -eq "CrVersion" | Select-Object Value).Value
        $crVersion

        <# MAJOR VERSION #>
        Write-Host "Major Version" -ForegroundColor Red
        $majorVersion = ($server[0] | Where-Object Key -eq "Major" | Select-Object Value).Value
        $majorVersion

        <# OPEN SUITE #>
        Write-Host "OpenSuite" -ForegroundColor Red
        $opensuite = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {
            if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {

                $software = "*Actian*";
                $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where { $_.DisplayName -like $software }) -ne $null

                if ($installed) {
                    return "Yes"
                }
                
            } else {
                return "No"
            }
        }
        Write-Host "OpenSuite Detected: $($opensuite)"

        <# INTEGRATIONS #>
        Write-Host "Integrations" -ForegroundColor Red
        $PPAdapter = "False"
        $LKAdapter = "False"
        $integrations = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {
            param ($database)
            if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*") {
                Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*" | Select-Object -Property PSChildName
            } else {
                return 0
            }
        } -ArgumentList $productionDatabase

        if ($integrations -eq 0) {
            Write-Host "No integrations found in 'HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($productionDatabase)\'"
        } else {
            Write-Host "Number of integrations found: $($integrations.PSChildName.Count)" -ForegroundColor Cyan
            foreach ($x in $integrations.PSChildName) {
                if ($x -like "*ProjectPlace*") {
                    Write-Host "PP ADAPTER FOUND: $($x)"
                    $PPAdapter = "True"
                }
                elseif ($x -like "*PRM_Adapter*") {
                    Write-Host "LK ADAPTER FOUND: $($x)"
                    $LKAdapter= "True"
                } else {
                    Write-Host "Other Integration Identified: $($x)"
                }
            }
        }

        # NEW RELIC #
        Write-Host "New Relic" -ForegroundColor Red
        $newRelic = Invoke-Command -ComputerName "$($serverName.Value)" -Credential $credentials -ScriptBlock {
            if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                Write-Host "New Relic has been detected on this server"
                return "Yes"
            } else {
                Write-Host "New Relic was not detected on this server"
                return "No"
            }
        }

            # GET WEB CONFIG #
            $webConfig = Invoke-Command -ComputerName "$($serverName.Value)" -Credential $credentials -ScriptBlock {
                return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
            }
            $webConfig = [xml] $webConfig

        <# PRODUCTION URL #>
        Write-Host "Production URL" -ForegroundColor Red
        $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
        $metadataURL = ($server[0] | Where-Object Key -eq "Cust_Url" | Select-Object Value).Value
        Write-Host "Instance Metadata: $($metadataURL)"
        Write-Host "Web config file: $($environmentURL.value)"

        <# DNS ALIAS #>
        Write-Host "Production DNS Alias" -ForegroundColor Red            
        $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
        $dnsAlias

        <# REPORT FARM URL #>
        Write-Host "Report Farm URL" -ForegroundColor Red
        $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
        Write-Host $reportfarmURL.value

        <# ENCRYPTED PVMASTER PASSWORD #>
        Write-Host "Encrypted PVMaster Password" -ForegroundColor Red
        $encryptedPVMasterPassword = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUserPassword"} | Select-Object -Property value
        Write-Host $encryptedPVMasterPassword.value
        
        <# UNENCRYPTED PVMASTER PASSWORD #>
        Write-Host "Unencrypted PVMaster Password" -ForegroundColor Red
        $unencryptedPVMasterPassword = Invoke-PassUtil -InputString $encryptedPVMasterPassword.value -Deobfuscation
        Write-Host $unencryptedPVMasterPassword

        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress
        
        <# EXCEL LOGIC AND VARIABLES#>
        $webServerCount++
        $buildData.Cells.Item(24,2)= $webServerCount
        $buildData.Cells.Item(1,2)= $environmentURL.value
        $buildData.Cells.Item(7,2)= $dnsAlias
        $buildData.Cells.Item(46,2)= $reportfarmURL.value
        $buildData.Cells.Item(15,2)= $newRelic
        $buildData.Cells.Item(36,2)= $opensuite
        $buildData.Cells.Item(25,2)= $crVersion
        $buildData.Cells.Item(13,2)= $majorVersion
        $buildData.Cells.Item(19,2)= "True"
        $buildData.Cells.Item(17,2)= $encryptedPVMasterPassword.value
        $buildData.Cells.Item(16,2)= $unencryptedPVMasterPassword

        $buildData.Cells.Item(31,2)= $PPAdapter
        $buildData.Cells.Item(32,2)= $LKAdapter

        $buildData.Cells.Item(56,2)= "$($serverName.Value)"
        $buildData.Cells.Item(56,3)= "$($server[2][2].DeviceID.Count)"
        $buildData.Cells.Item(56,4)= "$($server[2][1].MaxCapacity / 1048576)"
        $buildData.Cells.Item(56,5)= $hdStringArray
        $buildData.Cells.Item(56,6)= $diskResize
        $buildData.Cells.Item(56,7)= $task_array
        $buildData.Cells.Item(56,8)= $instanceSize
        $buildData.Cells.Item(56,9)= $AvailabilityZone
        $buildData.Cells.Item(56,10)= $ipAddress
        $buildData.Cells.Item(56,11)= $instanceID

        Write-Host "================================================================================" -ForegroundColor Red  
    }
    
}


#######################
# SANDBOX SERVERS
#######################
Write-Host ":::::::: SANDBOX ENVIRONMENT ::::::::" -ForegroundColor Yellow

$webServerCount = 0
foreach ($server in $sandboxComputers){

    $serverName = $server[0] | Where-Object Key -eq "Name" | Select-Object Value

    #######################
    # SANDBOX APP SERVER 
    #######################
    if ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "app") {

        Write-Host "THIS IS THE SANDBOX APP SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }
            
            <# OPEN SUITE #>
            Write-Host "OpenSuite" -ForegroundColor Red
            $opensuite = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {
                if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {

                    $software = "*Actian*";
                    $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where { $_.DisplayName -like $software }) -ne $null

                    if ($installed) {
                        return "Yes"
                    }
                    
                } else {
                    return "No"
                }
            }
            Write-Host "OpenSuite Detected: $($opensuite)"

            <# INTEGRATIONS #>
            Write-Host "Integrations" -ForegroundColor Red
            $PPAdapter = "False"
            $LKAdapter = "False"
            $integrations = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {
                param ($database)
                if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*") {
                    Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*" | Select-Object -Property PSChildName
                } else {
                    return 0
                }
            } -ArgumentList $sandboxDatabase

            if ($integrations -eq 0) {
                Write-Host "No integrations found in 'HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($sandboxDatabase)\'"
            } else {
                Write-Host "Number of integrations found: $($integrations.PSChildName.Count)" -ForegroundColor Cyan
                foreach ($x in $integrations.PSChildName) {
                    if ($x -like "*ProjectPlace*") {
                        Write-Host "PP ADAPTER FOUND: $($x)"
                        $PPAdapter = "True"
                    }
                    elseif ($x -like "*PRM_Adapter*") {
                        Write-Host "LK ADAPTER FOUND: $($x)"
                        $LKAdapter= "True"
                    } else {
                        Write-Host "Other Integration Identified: $($x)"
                    }
                }
            }

        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress            
            
        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(72,2)= "$($serverName.Value)"
        $buildData.Cells.Item(72,3)= "$($server[2][2].DeviceID.Count)"
        $buildData.Cells.Item(72,4)= "$($server[2][1].MaxCapacity / 1048576)"
        $buildData.Cells.Item(72,5)= $hdStringArray
        $buildData.Cells.Item(72,6)= $diskResize
        $buildData.Cells.Item(72,7)= $task_array
        $buildData.Cells.Item(72,8)= $instanceSize
        $buildData.Cells.Item(72,9)= $AvailabilityZone
        $buildData.Cells.Item(72,10)= $ipAddress
        $buildData.Cells.Item(72,11)= $instanceID

        $buildData.Cells.Item(31,3)= $PPAdapter
        $buildData.Cells.Item(32,3)= $LKAdapter
        
        $buildData.Cells.Item(36,3)= $opensuite

        Write-Host "================================================================================" -ForegroundColor Red
    }

    ###############################
    # SANDBOX CTM SERVER (Troux) 
    ###############################
    if ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "ctm") {
        Write-Host "THIS IS THE SANDBOX TROUX SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"    

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress        
            
        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(73,2)= "$($serverName.Value)"
        $buildData.Cells.Item(73,3)= "$($server[2][2].DeviceID.Count)"
        $buildData.Cells.Item(73,4)= "$($server[2][1].MaxCapacity / 1048576)"
        $buildData.Cells.Item(73,5)= $hdStringArray
        $buildData.Cells.Item(73,6)= $diskResize
        $buildData.Cells.Item(73,7)= $task_array
        $buildData.Cells.Item(73,8)= $instanceSize
        $buildData.Cells.Item(73,9)= $AvailabilityZone
        $buildData.Cells.Item(73,10)= $ipAddress
        $buildData.Cells.Item(73,11)= $instanceID

        Write-Host "================================================================================" -ForegroundColor Red  
    }

    #######################
    # SANDBOX WEB SERVER 
    #######################
    elseif ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "web") {
        Write-Host "THIS IS THE SANDBOX WEB SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# CURRENT VERSION #>
        Write-Host "Current Environment Version" -ForegroundColor Red
        $crVersion = ($server[0] | Where-Object Key -eq "CrVersion" | Select-Object Value).Value
        $crVersion

        <# MAJOR VERSION #>
        Write-Host "Major Version" -ForegroundColor Red
        $majorVersion = ($server[0] | Where-Object Key -eq "Major" | Select-Object Value).Value
        $majorVersion

        # NEW RELIC #
        Write-Host "New Relic" -ForegroundColor Red
        $newRelic = Invoke-Command -ComputerName "$($serverName.Value)" -Credential $credentials -ScriptBlock {
            if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                Write-Host "New Relic has been detected on this server"
                return "Yes"
            } else {
                Write-Host "New Relic was not detected on this server"
                return "No"
            }
        }

            # GET WEB CONFIG #
            $webConfig = Invoke-Command -ComputerName "$($serverName.Value)" -Credential $credentials -ScriptBlock {
                return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
            }
            $webConfig = [xml] $webConfig

        <# SANDBOX URL #>
        Write-Host "Sandbox URL" -ForegroundColor Red
        $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
        $metadataURL = ($server[0] | Where-Object Key -eq "Cust_Url" | Select-Object Value).Value
        Write-Host "Instance Metadata: $($metadataURL)"
        Write-Host "Web config file: $($environmentURL.value)"

        <# DNS ALIAS #>
        Write-Host "Sandbox DNS Alias" -ForegroundColor Red            
        $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
        $dnsAlias

        <# REPORT FARM URL #>
        Write-Host "Report Farm URL" -ForegroundColor Red
        $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
        Write-Host $reportfarmURL.value

        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress        

        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(25,3)= $crVersion
        $buildData.Cells.Item(2,2)= $environmentURL.value
        $buildData.Cells.Item(8,2)= $dnsAlias

        if ($webServerCount -gt 0){
            $buildData.Cells.Item(78 + ($webServerCount - 1),2)= "$($serverName.Value)"
            $buildData.Cells.Item(78 + ($webServerCount - 1),3)= "$($server[2][2].DeviceID.Count)"
            $buildData.Cells.Item(78 + ($webServerCount - 1),4)= "$($server[2][1].MaxCapacity / 1048576)"
            $buildData.Cells.Item(78 + ($webServerCount - 1),5)= $hdStringArray
            $buildData.Cells.Item(78 + ($webServerCount - 1),6)= $diskResize
            $buildData.Cells.Item(78 + ($webServerCount - 1),7)= $task_array
            $buildData.Cells.Item(78 + ($webServerCount - 1),8)= $instanceSize
            $buildData.Cells.Item(78 + ($webServerCount - 1),9)= $AvailabilityZone
            $buildData.Cells.Item(78 + ($webServerCount - 1),10)= $ipAddress
            $buildData.Cells.Item(78 + ($webServerCount - 1),11)= $instanceID
        }
        else {
            $buildData.Cells.Item(71,2)= "$($serverName.Value)"
            $buildData.Cells.Item(71,3)= "$($server[2][2].DeviceID.Count)"
            $buildData.Cells.Item(71,4)= "$($server[2][1].MaxCapacity / 1048576)"
            $buildData.Cells.Item(71,5)= $hdStringArray
            $buildData.Cells.Item(71,6)= $diskResize
            $buildData.Cells.Item(71,7)= $task_array
            $buildData.Cells.Item(71,8)= $instanceSize
            $buildData.Cells.Item(71,9)= $AvailabilityZone
            $buildData.Cells.Item(71,10)= $ipAddress
            $buildData.Cells.Item(71,11)= $instanceID
        }

        $webServerCount++
        $buildData.Cells.Item(24,3)= $webServerCount

        Write-Host "================================================================================" -ForegroundColor Red  

   }

    #######################
    # SANDBOX SAS SERVER 
    #######################
    elseif ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "sas") {
        Write-Host "THIS IS THE SANDBOX SAS SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress        
        
        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(75,2)= "$($serverName.Value)"
        $buildData.Cells.Item(75,3)= "$($server[2][2].DeviceID.Count)"
        $buildData.Cells.Item(75,4)= "$($server[2][1].MaxCapacity / 1048576)"
        $buildData.Cells.Item(75,5)= $hdStringArray
        $buildData.Cells.Item(75,6)= $diskResize
        $buildData.Cells.Item(75,7)= $task_array
        $buildData.Cells.Item(75,8)= $instanceSize
        $buildData.Cells.Item(75,9)= $AvailabilityZone
        $buildData.Cells.Item(75,10)= $ipAddress
        $buildData.Cells.Item(75,11)= $instanceID

        Write-Host "================================================================================" -ForegroundColor Red  
   }

    #######################
    # SANDBOX SQL SERVER 
    #######################
    elseif ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "sql") {
        Write-Host "THIS IS THE SANDBOX SQL SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }
        
        <# DATABASE PROPERTIES #>
        Write-Host "$($sandboxDatabase) Properties" -ForegroundColor Red
        $sqlSession = New-PSSession -ComputerName "$($serverName.Value)" -Credential $credentials

            # MAXDOP/THRESHOLD
            Write-Host "Identifying MaxDOP/Threshold..." -ForegroundColor Cyan
            $database_maxdop_threshold = Invoke-Command  -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                '%parallel%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
            } -ArgumentList $serverName.Value
            $maxdop = $database_maxdop_threshold | Where-Object {$_.name -like "cost*"} | Select-Object -property value
            $cost_threshold = $database_maxdop_threshold | Where-Object {$_.name -like "max*"} | Select-Object -property value
            Write-Host "Max DOP --- $($maxdop.value) MB"
            Write-Host "Cost Threshold --- $($cost_threshold.value) MB"
            
            # MIN/MAX MEMORY
            Write-Host "Identifying MIN/MAX Memory..." -ForegroundColor Cyan
            $database_memory = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                '%server memory%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
            } -ArgumentList $serverName.Value 
            $database_memory_max = $database_memory | where-Object {$_.name -like "max*"} | Select-Object -property value
            $database_memory_min = $database_memory | where-Object {$_.name -like "min*"} | Select-Object -property value
            Write-Host "Max Server Memory --- $($database_memory_max.value) MB"
            Write-Host "Min Server Memory --- $($database_memory_min.value) MB"
            
            # DATABASE ENCRYPTION
            Write-Host "Identifying Database Encryption..." -ForegroundColor Cyan
            $database_encryption = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT
                db.name,
                db.is_encrypted
                FROM
                sys.databases db
                LEFT OUTER JOIN sys.dm_database_encryption_keys dm
                    ON db.database_id = dm.database_id;
                GO" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value
            $dbEncryption = $database_encryption | Where-Object {$_.name -eq $sandboxDatabase}
            Write-Host "$($dbEncryption.name) --- $($dbEncryption.is_encrypted)"

            # DATABASE SIZE (MAIN)
            Write-Host "Calculating Database Size" -ForegroundColor Cyan
            $database_dbSize = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database)
                GO
                exec sp_spaceused
                GO" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$sandboxDatabase
            Write-Host "$($database_dbSize.database_name) --- $($database_dbSize.database_size)"

            # ALL DATABASES (NAMES AND SIZES in MB)
            Write-Host "Listing All Databases and Sizes (in MB)" -ForegroundColor Cyan
            $all_databases = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "SELECT d.name,
                ROUND(SUM(mf.size) * 8 / 1024, 0) Size_MB
                FROM sys.master_files mf
                INNER JOIN sys.databases d ON d.database_id = mf.database_id
                WHERE d.database_id > 4 -- Skip system databases
                GROUP BY d.name
                ORDER BY d.name" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$sandboxDatabase
            foreach ($database in $all_databases) {
                Write-Host "$($database.name) ---- $($database.Size_MB) MB"
            }

            # CUSTOM MODELS
            Write-Host "Calculating Custom Models..." -ForegroundColor Cyan
            $database_custom_models = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT COUNT(*) FROM ip.olap_properties 
                WHERE bism_ind ='N' 
                AND olap_obj_name 
                NOT like 'PVE%'" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$sandboxDatabase | Select-Object -property olap_obj_name
            foreach ($model in $database_custom_models.olap_obj_name) {
                Write-Host $model
            }  
            
            # INTERFACES
            Write-Host "Identifying Interfaces..." -ForegroundColor Cyan
            $database_interfaces = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT
                s.description JobStreamName,
                j.description JobName,
                j.job_order JobOrder,
                j.job_id JobID,
                p.name ParamName,
                p.param_value ParamValue,
                MIN(r.last_started) JobLastStarted,
                MAX(r.last_finished) JobLastFinished,
                MAX(CONVERT(CHAR(8), DATEADD(S,DATEDIFF(S,r.last_started,r.last_finished),'1900-1-1'),8)) Duration
                FROM ip.job_stream_job j
                INNER JOIN ip.job_stream s
                ON j.job_stream_id = s.job_stream_id
                INNER JOIN ip.job_stream_schedule ss
                ON ss.job_stream_id = s.job_stream_id
                INNER JOIN ip.job_run_status r
                ON s.job_stream_id = r.job_stream_id
                LEFT JOIN ip.job_param p
                ON j.job_id = p.job_id
                WHERE P.Name = 'Command'
                GROUP BY
                s.description,
                j.description,
                j.job_order,
                j.job_id,
                p.name,
                p.param_value;" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$sandboxDatabase
            $database_interfaces.ParamValue  
            
            # LICENSE COUNT
            Write-Host "Calculating License Count..." -ForegroundColor Cyan
            $database_license_count = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT
                LicenseRole,
                COUNT(UserName) UserCount,
                r.seats LicenseCount
                FROM (
                SELECT
                s1.description LicenseRole,
                s1.structure_code LicenseCode,
                u.active_ind Active,
                u.full_name UserName
                FROM ip.ip_user u
                INNER JOIN ip.structure s
                ON u.role_code = s.structure_code
                INNER JOIN ip.structure s1
                ON s.father_code = s1.structure_code
                WHERE u.active_ind = 'Y'
                ) q
                INNER JOIN ip.ip_role r
                ON q.LicenseCode = r.role_code
                GROUP BY
                LicenseRole,
                LicenseCode,
                r.seats" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$sandboxDatabase
            $licenseProperties = $database_license_count | Select-Object -Property LicenseRole,LicenseCount
            $totalLicenseCount = 0
            foreach ($license in $licenseProperties){
                Write-Output "$($license.LicenseRole): $($license.LicenseCount)"
                $totalLicenseCount += $license.LicenseCount
            }
            Write-Output "Total License Count: $($totalLicenseCount)"

            # PROGRESSING WEB VERSION
            Write-Host "Identifying Progressing Web Version..." -ForegroundColor Cyan
            $database_progressing_web_version = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database); SELECT TOP 1 sub_release 
                FROM ip.pv_version 
                WHERE release = 'PROGRESSING_WEB'
                ORDER BY seq DESC;" -ServerInstance $server.Name 
            } -ArgumentList $serverName.Value,$sandboxDatabase
            $database_progressing_web_version.sub_release
            
        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress            

        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(12,2)= $serverName.Value.Substring(($serverName.Value.Length - 2), 2)
        $buildData.Cells.Item(44,3)= $database_dbSize.database_size
        $buildData.Cells.Item(43,3)= $database_memory_max.value
        $buildData.Cells.Item(42,3)= $database_memory_min.value

        $buildData.Cells.Item(74,2)= "$($serverName.Value)"
        $buildData.Cells.Item(74,3)= "$($server[2][2].DeviceID.Count)"
        $buildData.Cells.Item(74,4)= "$($server[2][1].MaxCapacity / 1048576)"
        $buildData.Cells.Item(74,5)= $hdStringArray
        $buildData.Cells.Item(74,6)= $diskResize
        $buildData.Cells.Item(74,7)= $task_array
        $buildData.Cells.Item(74,8)= $instanceSize
        $buildData.Cells.Item(74,9)= $AvailabilityZone
        $buildData.Cells.Item(74,10)= $ipAddress
        $buildData.Cells.Item(74,11)= $instanceID
        
        $buildData.Cells.Item(26,3)= $database_progressing_web_version.sub_release

        $buildData.Cells.Item(28,2)= $database_custom_models.Count
        $modelCount = 0;
        foreach ($model in $database_custom_models.olap_obj_name){
            $buildData.Cells.Item(91, (2 + $modelCount))= $model
            $modelCount++
        }

        $databaseCount = 0
        foreach ($database in $all_databases) {
            $buildData.Cells.Item(101, (2 + $databaseCount))= $database.name
            $buildData.Cells.Item(102, (2 + $databaseCount))= "Size: $($database.Size_MB) MB"
            $databaseCount++
        }

        $buildData.Cells.Item(30,3)= $database_interfaces.ParamValue.Count
        $interfaceCount = 0
        foreach ($interface in $database_interfaces.ParamValue) {
            $buildData.Cells.Item(96, (2 + $interfaceCount))= $interface
            $interfaceCount++
        }

        $buildData.Cells.Item(41,3)= $dbEncryption.is_encrypted
        $buildData.Cells.Item(22,3)= $totalLicenseCount
        $buildData.Cells.Item(40,3)= $cost_threshold.value     
        $buildData.Cells.Item(39,3)= $maxdop.value

        Remove-PSSession -Session $sqlSession

        Write-Host "================================================================================" -ForegroundColor Red  
   }

    #######################
    # SANDBOX PVE SERVER 
    #######################
    elseif ($serverName.Value.Substring(($serverName.Value.Length - 5), 3) -eq "pve") {
        Write-Host "THIS IS THE SANDBOX PVE SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($serverName.Value)"
        Write-Host "Server CPUs: $($server[2][2].DeviceID.Count)"
        Write-Host "Server RAM: $($server[2][1].MaxCapacity / 1048576)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[2][0]) {
            $hdString = "$($hd.DeviceID): $($hd.Size / 1073741824)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString
            if (($hd.Size / 1073741824) -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# CURRENT VERSION #>
        Write-Host "Current Environment Version" -ForegroundColor Red
        $crVersion = ($server[0] | Where-Object Key -eq "CrVersion" | Select-Object Value).Value
        $crVersion

        <# MAJOR VERSION #>
        Write-Host "Major Version" -ForegroundColor Red
        $majorVersion = ($server[0] | Where-Object Key -eq "Major" | Select-Object Value).Value
        $majorVersion
        
        <# OPEN SUITE #>
        Write-Host "OpenSuite" -ForegroundColor Red
        $opensuite = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {
            if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {

                $software = "*Actian*";
                $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where { $_.DisplayName -like $software }) -ne $null

                if ($installed) {
                    return "Yes"
                }
                
            } else {
                return "No"
            }
        }
        Write-Host "OpenSuite Detected: $($opensuite)"

        <# INTEGRATIONS #>
        Write-Host "Integrations" -ForegroundColor Red
        $PPAdapter = "False"
        $LKAdapter = "False"
        $integrations = Invoke-Command -ComputerName $serverName.Value -Credential $credentials -ScriptBlock {
            param ($database)
            if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*") {
                Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*" | Select-Object -Property PSChildName
            } else {
                return 0
            }
        } -ArgumentList $sandboxDatabase

        if ($integrations -eq 0) {
            Write-Host "No integrations found in 'HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($sandboxDatabase)\'"
        } else {
            Write-Host "Number of integrations found: $($integrations.PSChildName.Count)" -ForegroundColor Cyan
            foreach ($x in $integrations.PSChildName) {
                if ($x -like "*ProjectPlace*") {
                    Write-Host "PP ADAPTER FOUND: $($x)"
                    $PPAdapter = "True"
                }
                elseif ($x -like "*PRM_Adapter*") {
                    Write-Host "LK ADAPTER FOUND: $($x)"
                    $LKAdapter= "True"
                } else {
                    Write-Host "Other Integration Identified: $($x)"
                }
            }
        }

        # NEW RELIC #
        Write-Host "New Relic" -ForegroundColor Red
        $newRelic = Invoke-Command -ComputerName "$($serverName.Value)" -Credential $credentials -ScriptBlock {
            if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                Write-Host "New Relic has been detected on this server"
                return "Yes"
            } else {
                Write-Host "New Relic was not detected on this server"
                return "No"
            }
        }

            # GET WEB CONFIG #
            $webConfig = Invoke-Command -ComputerName "$($serverName.Value)" -Credential $credentials -ScriptBlock {
                return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
            }
            $webConfig = [xml] $webConfig

        <# SANDBOX URL #>
        Write-Host "Sandbox URL" -ForegroundColor Red
        $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
        $metadataURL = ($server[0] | Where-Object Key -eq "Cust_Url" | Select-Object Value).Value
        Write-Host "Instance Metadata: $($metadataURL)"
        Write-Host "Web config file: $($environmentURL.value)"

        <# DNS ALIAS #>
        Write-Host "Sandbox DNS Alias" -ForegroundColor Red            
        $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
        $dnsAlias

        <# REPORT FARM URL #>
        Write-Host "Report Farm URL" -ForegroundColor Red
        $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
        Write-Host $reportfarmURL.value

        <# INSTANCE ID #>
        Write-Host "Instance ID" -ForegroundColor Red
        $instanceID = $server[1][0]
        Write-Host $instanceID

        <# INSTANCE SIZE #>
        Write-Host "Instance Size" -ForegroundColor Red
        $instanceSize = $server[1][2]
        Write-Host $instanceSize

        <# AVAILABILITY ZONE #>
        Write-Host "Availability Zone" -ForegroundColor Red
        $AvailabilityZone = $server[1][3]
        Write-Host $AvailabilityZone

        <# IP ADDRESS #>
        Write-Host "IP Address" -ForegroundColor Red
        $ipAddress = $server[1][4]
        Write-Host $ipAddress        
        
        <# EXCEL LOGIC AND VARIABLES#>
        $webServerCount++
        $buildData.Cells.Item(24,3)= $webServerCount
        $buildData.Cells.Item(2,2)= $environmentURL.value   
        $buildData.Cells.Item(8,2)= $dnsAlias
        $buildData.Cells.Item(36,3)= $opensuite
        $buildData.Cells.Item(25,3)= $crVersion
        $buildData.Cells.Item(19,2)= "True"

        $buildData.Cells.Item(76,2)= "$($serverName.Value)"
        $buildData.Cells.Item(76,3)= "$($server[2][2].DeviceID.Count)"
        $buildData.Cells.Item(76,4)= "$($server[2][1].MaxCapacity / 1048576)"
        $buildData.Cells.Item(76,5)= $hdStringArray
        $buildData.Cells.Item(76,6)= $diskResize
        $buildData.Cells.Item(76,7)= $task_array
        $buildData.Cells.Item(76,8)= $instanceSize
        $buildData.Cells.Item(76,9)= $AvailabilityZone
        $buildData.Cells.Item(76,10)= $ipAddress
        $buildData.Cells.Item(76,11)= $instanceID

        $buildData.Cells.Item(31,3)= $PPAdapter
        $buildData.Cells.Item(32,3)= $LKAdapter

        Write-Host "================================================================================" -ForegroundColor Red  
    }

}
############################################################
# SAVE AND CLOSE
# SAVES AND CLOSES THE EXCEL WORKBOOK  
############################################################
$excelfile.Save()
$excelfile.Close()
}
else {

############################################################
#                       
#
#
#                           US-EU
#
#
#
############################################################

############################################################
# SECONDARY VARIABLES - LESSER IMPORTANT JUNK 
############################################################
Write-Host "Sign-in with your 'Planview\<a-admin>' account:" -ForegroundColor Magenta
$aAdmin = "a-$($admin)"
$credentials = Get-Credential "Planview\$($aAdmin)"
$vSphereCredentials = New-Object System.Management.Automation.PSCredential ($aAdmin, $credentials.Password)
$f5Credentials = $vSphereCredentials

$customerName = Read-Host "Enter the customer OU name"
$customerCode = Read-Host "Enter the customer code"
$option = Read-Host "Select a region: 1-SG or 2-LN"

$productionOUName = "production"
$sandboxOUName = "sandbox"

$excel_file = Get-ChildItem -Path "C:\Users\$($admin)$($domain)\Desktop\" -Filter "InPlace*"

$productionDatabase = "$($customerCode.ToUpper())PROD"
$ctmDatabase = "$($customerCode.ToUpper())CTM"
$sandboxDatabase = "$($customerCode.ToUpper())SANDBOX1"
$excelFilePath = ("C:\Users\$($admin)$($domain)\Desktop\$excel_file")
$directoryPath = "C:\Users\$($admin)$($domain)\Planview, Inc\E1 Build Cutover - Documents\Customer Builds"

############################################################
# DEFINES THE JUMPBOX, ACTIVE DIRECTORY, VSHPERE SERVERS
# BASED ON THE DATACENTER LOCATION
############################################################
switch($option) {
    1 {$jumpbox = "sgcvmrdp02.us.planview.world"; 
        $ad_server = "SGCVMADC14.us.planview.world";
        $vSphereServer = "sgivmvcsvr06.us.planview.world"; 
        $dataCenterLocation = "sg"; 
        $reportFarm = "https://usreportfarm03.pvcloud.com/ReportServer";
        $f5ip = "10.132.81.2";
        break}
    2 {$jumpbox = "sgcvmrdp02.us.planview.world"; 
        $ad_server = "LNCVMADC05.eu.planview.world"; 
        $vSphereServer = "lnivmvcsvr06.eu.planview.world"; 
        $dataCenterLocation = "ln"; 
        $reportFarm = "https://eureportfarm03.pvcloud.com/ReportServer";
        $f5ip = "10.60.2.171";
        break}
}

############################################################
# INITIATE THE JUMPBOX PERSISTENT SESSION
############################################################
$session = New-PSSession -ComputerName $jumpbox -Authentication Credssp -Credential $credentials

############################################################
# FINDS THE ORGANIZATIONAL UNIT 
# FOR THE CUSTOMER NAME PROVIDED
############################################################
Write-Host "Connecting to Active Directory..." -ForegroundColor Gray
$AD_OU = Get-ADOrganizationalUnit -Filter { Name -like $customerName } -Server $ad_server

############################################################
# LOCATES THE COMPUTER OBJECTS WITHIN 
# THE ENABLED ORGANIZATIONAL UNITS
############################################################
try {
    $productionComputers = Get-ADComputer -Filter * -SearchBase  "OU=$($productionOUName),$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name 
    Write-Host "Total Servers in ""Production"" OU: $($productionComputers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $sandboxComputers = Get-ADComputer -Filter * -SearchBase  "OU=$($sandboxOUName),$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name
    Write-Host "Total Servers in ""Sandbox"" OU: $($sandboxComputers.Length)" -ForegroundColor Yellow
}
catch { }
<#
try {
    $testComputers = Get-ADComputer -Filter * -SearchBase  "OU=Test,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name
    Write-Host "Total Servers in ""Test"" OU: $($testComputers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $trouxComputers = Get-ADComputer -Filter * -SearchBase  "OU=Troux,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name 
    Write-Host "Total Servers in ""Troux"" OU: $($trouxComputers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $trouxTestComputers = Get-ADComputer -Filter * -SearchBase  "OU=TrouxTest,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name 
    Write-Host "Total Servers in ""TrouxTest"" OU: $($trouxTestComputers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $sandbox2Computers = Get-ADComputer -Filter * -SearchBase  "OU=Sandbox2,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name
    Write-Host "Total Servers in ""Sandbox2"" OU: $($sandbox2Computers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $sandbox3Computers = Get-ADComputer -Filter * -SearchBase  "OU=Sandbox3,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name
    Write-Host "Total Servers in ""Sandbox3"" OU: $($sandbox3Computers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $sandbox4Computers = Get-ADComputer -Filter * -SearchBase  "OU=Sandbox4,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name
    Write-Host "Total Servers in ""Sandbox4"" OU: $($sandbox4Computers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $sandbox5Computers = Get-ADComputer -Filter * -SearchBase  "OU=Sandbox5,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name
    Write-Host "Total Servers in ""Sandbox5"" OU: $($sandbox5Computers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $sandbox6Computers = Get-ADComputer -Filter * -SearchBase  "OU=Sandbox6,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name
    Write-Host "Total Servers in ""Sandbox6"" OU: $($sandbox6Computers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $sandboxDisabledComputers = Get-ADComputer -Filter * -SearchBase  "OU=Sandbox_Disabled,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name 
    Write-Host "Total Servers in ""Sandbox Disabled"" OU: $($sandboxDisabledComputers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $devComputers = Get-ADComputer -Filter * -SearchBase  "OU=Dev,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name 
    Write-Host "Total Servers in ""Dev"" OU: $($devComputers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $preprodComputers = Get-ADComputer -Filter * -SearchBase  "OU=PreProd,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name 
    Write-Host "Total Servers in ""PreProd"" OU: $($preprodComputers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $preSandboxComputers = Get-ADComputer -Filter * -SearchBase  "OU=PreSandbox,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name
    Write-Host "Total Servers in ""PreSandbox"" OU: $($preSandboxComputers.Length)" -ForegroundColor Yellow
}
catch { }

try {
    $retireComputers = Get-ADComputer -Filter * -SearchBase  "OU=Retire,$($AD_OU.DistinguishedName)" -Server $ad_server | Select-Object -Property Name 
    Write-Host "Total Servers in ""Retire"" OU: $($retireComputers.Length)" -ForegroundColor Yellow
}
catch { }
#>

############################################################
# TRY TO EXPORT SERVER NAMES TO IMPORT INTO RDCM
############################################################
try{
    $productionComputers.Name + $sandboxComputers.Name | Out-File -FilePath "C:\Users\$($admin)$($domain)\Documents\RDCM Files\New Servers.txt" -Encoding UTF8
    Write-Host ":::::::: New Servers Available for Importing in RDCM ::::::::" -ForegroundColor DarkGreen
}
catch{
    Write-Host "There was an error making the server list available for importing in RDCM." -ForegroundColor DarkGreen
}

############################################################
# SCRIPTBLOCK THAT CONNECTS TO VSPHERE 
# VIA THE JUMPBOX AND RETRIEVES VIRTUAL MACHINE DETAILS
############################################################
$scriptBlock = {
    
    param ($servers, $credentials, $customerCode, $vSphereServer, $vSphereCredentials, $database)
    
    $computerObjects = @()

    <# CONNECTS SCRIPT TO VSPHERE SERVER #>
    Connect-VIServer -Server "$($vSphereServer)" -Credential $vSphereCredentials
    
    
    <# LOOPS THROUGH SERVERS SPECIFIED IN $serverArray #>
    Foreach ($server in $servers) {
        Write-Host "Connected to: $($server.Name)" -ForegroundColor Green
        
        Write-Host "Collecting CPU and memory information..." -ForegroundColor Cyan 
        $specs = Get-VM -Name $server.Name | Select-Object -Property Name, NumCpu, MemoryGB

        Write-Host "Collecting disk information..." -ForegroundColor Cyan
        $disks = Get-VM -Name $server.Name | Get-Harddisk

        Write-Host "Identifying server cluster..." -ForegroundColor Cyan 
        $cluster = Get-Cluster -VM $server.Name | Select-Object -Property Name

        Write-Host "Gathering scheduled task on server..." -ForegroundColor Cyan
        $tasks = Invoke-Command -computer $server.Name -ScriptBlock {
            Get-ScheduledTask -TaskPath "\" | Select-Object -Property TaskName, LastRunTime | where TaskName -notlike "Op*" 
        } -Credential $credentials

        Write-Host "================================================================================" -ForegroundColor Gray

        <# STORES SERVER ATTRIBUTES COLLECTED IN $computerObjects (NESTED ARRAY) #>
        $computerObjects += @(, (($specs), ($disks), ($cluster), ($tasks)))
    } 
    

    return $computerObjects
}

############################################################
# EXTRACTS INFO FROM SERVERS - ORGANIZED AS A NESTED ARRAY 
# COMPUTER_OBJECTS ->
# ((SERVER 1), (SERVER 2), ...) -> 
# (X[0], X[1], X[2], X[3]) 
############################################################
Write-Host "Starting production environment data collection..." -ForegroundColor Red
$computerObjects_Production = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $productionComputers, $credentials, $customerCode, $vSphereServer, $vSphereCredentials, $productionDatabase 

Write-Host "Starting sandbox environment data collection..." -ForegroundColor Red
$computerObjects_Sandbox = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $sandboxComputers, $credentials, $customerCode, $vSphereServer, $vSphereCredentials, $sandboxDatabase

<##################################################################################################>
<##################################################################################################>
##
##  INITIALIZING EXCEL DOCUMENT
##
<##################################################################################################>
<##################################################################################################>

############################################################
# APPLICATION LAYER                                
# INSTANCIATES EXCEL IN PS AND OPENS THE EXCEL FILE 
############################################################
$excel = New-Object -ComObject Excel.Application
$excelfile = $excel.Workbooks.Open($excelFilePath)

############################################################
# WORKSHEET LAYER                                
# CALLS A WORKSHEET FROM THE EXCEL FILE 
############################################################
# EXAMPLE #
# $buildData = $excelfile.sheets.item("Jenkins Inputs")
$buildData = $excelfile.sheets.item("MasterConfig")

############################################################
# RANGE LAYER - (FOR READING THE EXCEL FILE)
# -HASH TABLES (KEY/VALUE PAIR FORMAT)
# -ISSUE A VARIABLE, THEN ASSIGN IT A CELL VALUE.
############################################################
# EXAMPLE #
# $data_JenkinsInputs = @{
# "<us_all_pipe>" = $buildData.RANGE("C2").Text  
# "<target_server_name>" = $buildData.RANGE("C3").Text
# }

############################################################
# RANGE LAYER - (FOR WRITING TO THE EXCEL FILE)
# TARGET A CELL ($ROW, $COLUMN) AND ASSIGN IT A VALUE
############################################################
# EXAMPLE #
# $buildData.Cells.Item(2,5)= 'Hello'
#AWS BUILD
$buildData.Cells.Item(18,2)= "False" 
#SPLIT TIER
$buildData.Cells.Item(19,2)= "False"

$buildData.Cells.Item(23,2)= $productionComputers.Count
$buildData.Cells.Item(23,3)= $sandboxComputers.Count
$buildData.Cells.Item(9,2)= "$($dataCenterLocation)"
$buildData.Cells.Item(14,2)= "$($AD_OU.Name)"
$buildData.Cells.Item(10,2)= $customerCode.ToUpper()
#$buildData.Cells.Item(46,2)= $reportFarm
$buildData.Cells.Item(3,2)= "http://saasinfo.planview.world/$($customerName.Split(':')[0]).htm"


############################################################
# ITERATES THROUGH PRODUCTION AND SANDBOX SERVERS 
# PLACES DATA IN EXCEL SHEET CELLS
############################################################

#######################
# PRODUCTION SERVERS
#######################
Write-Host ":::::::: PRODUCTION ENVIRONMENT ::::::::" -ForegroundColor Yellow

$webServerCount = 0
foreach ( $server in $computerObjects_Production){
    
    if ($server[0].Name.Substring(($server[0].Name.Length - 5), 3) -eq "app") {
      
        ##########################
        # PRODUCTION APP SERVER 
        ##########################
        if ($server[0].Name.Substring(3, 1) -ne 't') {
            Write-Host "THIS IS THE PRODUCTION APP SERVER" -ForegroundColor Cyan

            <# CPU/RAM #>
            Write-Host "Server CPU and RAM" -ForegroundColor Red
            Write-Host "Server Name: $($server[0].Name)"
            Write-Host "Server CPUs: $($server[0].NumCpu)"
            Write-Host "Server RAM: $($server[0].MemoryGB)"

            <# HARDDRIVES #>
            Write-Host "Disks and Disk Capacity" -ForegroundColor Red
            $diskResize = "Yes"
            $hdStringArray = ""
            foreach ($hd in $server[1]) {
                $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
                $hdStringArray += "$($hdString)`n"
                Write-Host $hdString  
                if ($hd.CapacityGB -gt 60) {
                    $diskResize = "No"  
                }
            }
            Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

            <# CLUSTER #>
#            Write-Host "Server Cluster" -ForegroundColor Red
#            Write-Host "Cluster Name: $($server[2].Name)"

            <# SCHEDULED TASKS #>
            Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
            $task_array = ""
            foreach ($task in $server[3]){
                Write-Host "Task Name: $($task.TaskName)"
                $task_array += "$($task.TaskName)`n"
            }
            
            <# OPEN SUITE #>
            Write-Host "OpenSuite" -ForegroundColor Red
            $opensuite = Invoke-Command -ComputerName $server[0].Name -Credential $credentials -ScriptBlock {
                if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {

                    $software = "*Actian*";
                    $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -like $software }) -ne $null

                    if ($installed) {
                        return "Yes"
                    }
                    
                } else {
                    return "No"
                }
            }
            Write-Host "OpenSuite Detected: $($opensuite)"

            <# INTEGRATIONS #>
            Write-Host "Integrations" -ForegroundColor Red
            $PPAdapter = "False"
            $LKAdapter = "False"
            $integrations = Invoke-Command -ComputerName $server[0].Name -Credential $credentials -ScriptBlock {
                param ($database)
                if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*") {
                    Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*" | Select-Object -Property PSChildName
                } else {
                    return 0
                }
            } -ArgumentList $productionDatabase

            if ($integrations -eq 0) {
                Write-Host "No integrations found in 'HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($productionDatabase)\'"
            } else {
                Write-Host "Number of integrations found: $($integrations.PSChildName.Count)" -ForegroundColor Cyan
                foreach ($x in $integrations.PSChildName) {
                    if ($x -like "*ProjectPlace*") {
                        Write-Host "PP ADAPTER FOUND: $($x)"
                        $PPAdapter = "True"
                    }
                    elseif ($x -like "*PRM_Adapter*") {
                        Write-Host "LK ADAPTER FOUND: $($x)"
                        $LKAdapter= "True"
                    } else {
                        Write-Host "Other Integration Identified: $($x)"
                    }
                }
            }

            <# EXCEL LOGIC AND VARIABLES#>
            $buildData.Cells.Item(52,2)= "$($server[0].Name)"
            $buildData.Cells.Item(52,3)= "$($server[0].NumCpu)"
            $buildData.Cells.Item(52,4)= "$($server[0].MemoryGB)"
            $buildData.Cells.Item(52,5)= $hdStringArray
            $buildData.Cells.Item(52,6)= $diskResize
            $buildData.Cells.Item(52,7)= $task_array

            $buildData.Cells.Item(31,2)= $PPAdapter
            $buildData.Cells.Item(32,2)= $LKAdapter

            $buildData.Cells.Item(36,2)= $opensuite

            Write-Host "================================================================================" -ForegroundColor Red
        } 

        ##################################
        # PRODUCTION CTM SERVER (Troux) 
        ##################################
        elseif ($server[0].Name.Substring(3, 1) -eq 't') {
            Write-Host "THIS IS THE PRODUCTION TROUX SERVER" -ForegroundColor Cyan

            <# CPU/RAM #>
            Write-Host "Server CPU and RAM" -ForegroundColor Red
            Write-Host "Server Name: $($server[0].Name)"
            Write-Host "Server CPUs: $($server[0].NumCpu)"
            Write-Host "Server RAM: $($server[0].MemoryGB)"     

            <# HARDDRIVES #>
            Write-Host "Disks and Disk Capacity" -ForegroundColor Red
            $diskResize = "Yes"
            $hdStringArray = ""
            foreach ($hd in $server[1]) {
                $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
                $hdStringArray += "$($hdString)`n"
                Write-Host $hdString  
                if ($hd.CapacityGB -gt 60) {
                    $diskResize = "No"  
                }
            }
            Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"       

            <# CLUSTER #>
#            Write-Host "Server Cluster" -ForegroundColor Red
#            Write-Host "Cluster Name: $($server[2].Name)"

            <# SCHEDULED TASKS #>
            Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
            $task_array = ""
            foreach ($task in $server[3]){
                Write-Host "Task Name: $($task.TaskName)"
                $task_array += "$($task.TaskName)`n"
            }

            <# EXCEL LOGIC AND VARIABLES#>
            $buildData.Cells.Item(53,2)= "$($server[0].Name)"
            $buildData.Cells.Item(53,3)= "$($server[0].NumCpu)"
            $buildData.Cells.Item(53,4)= "$($server[0].MemoryGB)"
            $buildData.Cells.Item(53,5)= $hdStringArray
            $buildData.Cells.Item(53,6)= $diskResize
            $buildData.Cells.Item(53,7)= $task_array

            Write-Host "================================================================================" -ForegroundColor Red  
        }
   }
   
    ##########################
    # PRODUCTION WEB SERVER 
    ##########################
    elseif ($server[0].Name.Substring(($server[0].Name.Length - 5), 3) -eq "web") {

        Write-Host "THIS IS THE PRODUCTION WEB SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($server[0].Name)"
        Write-Host "Server CPUs: $($server[0].NumCpu)"
        Write-Host "Server RAM: $($server[0].MemoryGB)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[1]) {
            $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString  
            if ($hd.CapacityGB -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# CURRENT VERSION #>
        Write-Host "Current Environment Version" -ForegroundColor Red
        $crVersion = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
            Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\WebServerPlatform"
        }
        Write-Host $crVersion.CrVersion

        <# MAJOR VERSION #>
        Write-Host "Major Version" -ForegroundColor Red
        $majorVersion = $crVersion.CrVersion.Split('.')[0]
        $majorVersion

        # NEW RELIC #
        Write-Host "New Relic" -ForegroundColor Red
        $newRelic = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
            if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                Write-Host "New Relic has been detected on this server"
                return "Yes"
            } else {
                Write-Host "New Relic was not detected on this server"
                return "No"
            }
        }

            # GET WEB CONFIG #
            $webConfig = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
                return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
            }
            $webConfig = [xml] $webConfig

        <# PRODUCTION URL #>
        Write-Host "Production URL" -ForegroundColor Red
        $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
        Write-Host $environmentURL.value

        <# DNS ALIAS #>
        Write-Host "Production DNS Alias" -ForegroundColor Red            
        $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
        $dnsAlias

        <# REPORT FARM URL #>
        Write-Host "Report Farm URL" -ForegroundColor Red
        $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
        Write-Host $reportfarmURL.value

        <# ENCRYPTED PVMASTER PASSWORD #>
        Write-Host "Encrypted PVMaster Password" -ForegroundColor Red
        $encryptedPVMasterPassword = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUserPassword"} | Select-Object -Property value
        Write-Host $encryptedPVMasterPassword.value
        
        <# UNENCRYPTED PVMASTER PASSWORD #>
        Write-Host "Unencrypted PVMaster Password" -ForegroundColor Red
        $unencryptedPVMasterPassword = Invoke-PassUtil -InputString $encryptedPVMasterPassword.value -Deobfuscation
        Write-Host $unencryptedPVMasterPassword

        <# IP RESTRICTIONS #>
        Write-Host "IP Restrictions on F5" -ForegroundColor Red
        $IPRestrictions = "No"
            
            # Authentication on the F5 #
            $websession =  New-Object Microsoft.PowerShell.Commands.WebRequestSession
            $jsonbody = @{username = $f5Credentials.UserName ; password = $f5Credentials.GetNetworkCredential().Password; loginProviderName='tmos'} | ConvertTo-Json
            $authResponse = Invoke-RestMethodOverride -Method Post -Uri "https://$($f5ip)/mgmt/shared/authn/login" -Credential $f5Credentials -Body $jsonbody -ContentType 'application/json'
            $token = $authResponse.token.token
            $websession.Headers.Add('X-F5-Auth-Token', $Token)

            # Calling data-group REST endpoint and parsing IPRestrictions list #
            $IPRestrictionsList = (Invoke-RestMethod  -Uri "https://$($f5ip)/mgmt/tm/ltm/data-group/internal" -WebSession $websession).Items | 
                Where-Object {$_.name -eq "IPRestrictions"} | Select-Object -Property records

            foreach ($record in $IPRestrictionsList.records) {
                if ($record.name -eq "$($dnsAlias).pvcloud.com") {
                    $IPRestrictions = "Yes"
                    Write-Host "IP restrctions were found for $($dnsAlias).pvcloud.com"
                }
            }

            if ($IPRestrictions -eq "No") {
                Write-Host "No IP restrictions found for $($dnsAlias).pvcloud.com"
            }

        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(25,2)= $crVersion.CrVersion
        $buildData.Cells.Item(35,2)= $IPRestrictions
        $buildData.Cells.Item(13,2)= $majorVersion
        $buildData.Cells.Item(17,2)= $encryptedPVMasterPassword.value
        $buildData.Cells.Item(16,2)= $unencryptedPVMasterPassword
        $buildData.Cells.Item(1,2)= $environmentURL.value
        $buildData.Cells.Item(7,2)= $dnsAlias
        $buildData.Cells.Item(46,2)= $reportfarmURL.value
        $buildData.Cells.Item(15,2)= $newRelic

        if ($webServerCount -gt 0){
            $buildData.Cells.Item(58 + ($webServerCount - 1),2)= "$($server[0].Name)"
            $buildData.Cells.Item(58 + ($webServerCount - 1),3)= "$($server[0].NumCpu)"
            $buildData.Cells.Item(58 + ($webServerCount - 1),4)= "$($server[0].MemoryGB)"
            $buildData.Cells.Item(58 + ($webServerCount - 1),5)= $hdStringArray
            $buildData.Cells.Item(58 + ($webServerCount - 1),6)= $diskResize
            $buildData.Cells.Item(58 + ($webServerCount - 1),7)= $task_array
        }
        else {
            $buildData.Cells.Item(51,2)= "$($server[0].Name)"
            $buildData.Cells.Item(51,3)= "$($server[0].NumCpu)"
            $buildData.Cells.Item(51,4)= "$($server[0].MemoryGB)"
            $buildData.Cells.Item(51,5)= $hdStringArray
            $buildData.Cells.Item(51,6)= $diskResize
            $buildData.Cells.Item(51,7)= $task_array
        }
        
        $webServerCount++
        $buildData.Cells.Item(24,2)= $webServerCount

        
        Write-Host "================================================================================" -ForegroundColor Red  

   }

    ##########################
    # PRODUCTION SAS SERVER 
    ##########################
    elseif ($server[0].Name.Substring(($server[0].Name.Length - 5), 3) -eq "sas") {
        Write-Host "THIS IS THE PRODUCTION SAS SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($server[0].Name)"
        Write-Host "Server CPUs: $($server[0].NumCpu)"
        Write-Host "Server RAM: $($server[0].MemoryGB)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[1]) {
            $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString  
            if ($hd.CapacityGB -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }
        
        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(55,2)= "$($server[0].Name)"
        $buildData.Cells.Item(55,3)= "$($server[0].NumCpu)"
        $buildData.Cells.Item(55,4)= "$($server[0].MemoryGB)"
        $buildData.Cells.Item(55,5)= $hdStringArray
        $buildData.Cells.Item(55,6)= $diskResize
        $buildData.Cells.Item(55,7)= $task_array

        Write-Host "================================================================================" -ForegroundColor Red  
   }

    ##########################
    # PRODUCTION SQL SERVER 
    ##########################
    elseif ($server[0].Name.Substring(($server[0].Name.Length - 5), 3) -eq "sql") {
        Write-Host "THIS IS THE PRODUCTION SQL SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($server[0].Name)"
        Write-Host "Server CPUs: $($server[0].NumCpu)"
        Write-Host "Server RAM: $($server[0].MemoryGB)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[1]) {
            $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString  
            if ($hd.CapacityGB -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# DATABASE PROPERTIES #>
        Write-Host "$($productionDatabase) Properties" -ForegroundColor Red
        $sqlSession = New-PSSession -ComputerName "$($server[0].Name)" -Credential $credentials

            # MAXDOP/THRESHOLD
            Write-Host "Identifying MaxDOP/Threshold..." -ForegroundColor Cyan
            $database_maxdop_threshold = Invoke-Command  -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                '%parallel%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
            } -ArgumentList $server[0].Name
            $maxdop = $database_maxdop_threshold | Where-Object {$_.name -like "cost*"} | Select-Object -property value
            $cost_threshold = $database_maxdop_threshold | Where-Object {$_.name -like "max*"} | Select-Object -property value
            Write-Host "Max DOP --- $($maxdop.value) MB"
            Write-Host "Cost Threshold --- $($cost_threshold.value) MB"
            
            # MIN/MAX MEMORY
            Write-Host "Identifying MIN/MAX Memory..." -ForegroundColor Cyan
            $database_memory = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                '%server memory%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
            } -ArgumentList $server[0].Name 
            $database_memory_max = $database_memory | where-Object {$_.name -like "max*"} | Select-Object -property value
            $database_memory_min = $database_memory | where-Object {$_.name -like "min*"} | Select-Object -property value
            Write-Host "Max Server Memory --- $($database_memory_max.value) MB"
            Write-Host "Min Server Memory --- $($database_memory_min.value) MB"
            
            # DATABASE ENCRYPTION
            Write-Host "Identifying Database Encryption..." -ForegroundColor Cyan
            $database_encryption = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT
                db.name,
                db.is_encrypted
                FROM
                sys.databases db
                LEFT OUTER JOIN sys.dm_database_encryption_keys dm
                    ON db.database_id = dm.database_id;
                GO" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name
            $dbEncryption = $database_encryption | Where-Object {$_.name -eq $productionDatabase}
            Write-Host "$($dbEncryption.name) --- $($dbEncryption.is_encrypted)"
            
            # DATABASE SIZE (MAIN)
            Write-Host "Calculating Database Size" -ForegroundColor Cyan
            $database_dbSize = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database)
                GO
                exec sp_spaceused
                GO" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$productionDatabase
            Write-Host "$($database_dbSize.database_name) --- $($database_dbSize.database_size)"

            # ALL DATABASES (NAMES AND SIZES in MB)
            Write-Host "Listing All Databases and Sizes (in MB)" -ForegroundColor Cyan
            $all_databases = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "SELECT d.name,
                ROUND(SUM(mf.size) * 8 / 1024, 0) Size_MB
                FROM sys.master_files mf
                INNER JOIN sys.databases d ON d.database_id = mf.database_id
                WHERE d.database_id > 4 -- Skip system databases
                GROUP BY d.name
                ORDER BY d.name" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$productionDatabase
            foreach ($database in $all_databases) {
                Write-Host "$($database.name) ---- $($database.Size_MB) MB"
            }

            # CUSTOM MODELS
            Write-Host "Calculating Custom Models..." -ForegroundColor Cyan
            $database_custom_models = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT * FROM ip.olap_properties 
                WHERE bism_ind ='N' 
                AND olap_obj_name 
                NOT like 'PVE%'" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$productionDatabase | Select-Object -property olap_obj_name
            foreach ($model in $database_custom_models.olap_obj_name) {
                Write-Host $model
            }          

            # INTERFACES
            Write-Host "Identifying Interfaces..." -ForegroundColor Cyan
            $database_interfaces = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT
                s.description JobStreamName,
                j.description JobName,
                j.job_order JobOrder,
                j.job_id JobID,
                p.name ParamName,
                p.param_value ParamValue,
                MIN(r.last_started) JobLastStarted,
                MAX(r.last_finished) JobLastFinished,
                MAX(CONVERT(CHAR(8), DATEADD(S,DATEDIFF(S,r.last_started,r.last_finished),'1900-1-1'),8)) Duration
                FROM ip.job_stream_job j
                INNER JOIN ip.job_stream s
                ON j.job_stream_id = s.job_stream_id
                INNER JOIN ip.job_stream_schedule ss
                ON ss.job_stream_id = s.job_stream_id
                INNER JOIN ip.job_run_status r
                ON s.job_stream_id = r.job_stream_id
                LEFT JOIN ip.job_param p
                ON j.job_id = p.job_id
                WHERE P.Name = 'Command'
                GROUP BY
                s.description,
                j.description,
                j.job_order,
                j.job_id,
                p.name,
                p.param_value;" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$productionDatabase
            $database_interfaces.ParamValue

            # LICENSE COUNT
            Write-Host "Calculating License Count..." -ForegroundColor Cyan
            $database_license_count = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT
                LicenseRole,
                COUNT(UserName) UserCount,
                r.seats LicenseCount
                FROM (
                SELECT
                s1.description LicenseRole,
                s1.structure_code LicenseCode,
                u.active_ind Active,
                u.full_name UserName
                FROM ip.ip_user u
                INNER JOIN ip.structure s
                ON u.role_code = s.structure_code
                INNER JOIN ip.structure s1
                ON s.father_code = s1.structure_code
                WHERE u.active_ind = 'Y'
                ) q
                INNER JOIN ip.ip_role r
                ON q.LicenseCode = r.role_code
                GROUP BY
                LicenseRole,
                LicenseCode,
                r.seats" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$productionDatabase
            $licenseProperties = $database_license_count | Select-Object -Property LicenseRole,LicenseCount
            $totalLicenseCount = 0
            foreach ($license in $licenseProperties){
                Write-Output "$($license.LicenseRole): $($license.LicenseCount)"
                $totalLicenseCount += $license.LicenseCount
            }
            Write-Output "Total License Count: $($totalLicenseCount)"
            
            # PROGRESSING WEB VERSION
            Write-Host "Identifying Progressing Web Version..." -ForegroundColor Cyan
            $database_progressing_web_version = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database); SELECT TOP 1 sub_release 
                FROM ip.pv_version 
                WHERE release = 'PROGRESSING_WEB'
                ORDER BY seq DESC;" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$productionDatabase
            $database_progressing_web_version.sub_release 

        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(11,2)= $server[0].Name.Substring(($server[0].Name.Length - 2), 2)
        $buildData.Cells.Item(44,2)= $database_dbSize.database_size
        $buildData.Cells.Item(43,2)= $database_memory_max.value
        $buildData.Cells.Item(42,2)= $database_memory_min.value

        $buildData.Cells.Item(54,2)= "$($server[0].Name)"
        $buildData.Cells.Item(54,3)= "$($server[0].NumCpu)"
        $buildData.Cells.Item(54,4)= "$($server[0].MemoryGB)"
        $buildData.Cells.Item(54,5)= $hdStringArray
        $buildData.Cells.Item(54,6)= $diskResize
        $buildData.Cells.Item(54,7)= $task_array

        $buildData.Cells.Item(26,2)= $database_progressing_web_version.sub_release

        $buildData.Cells.Item(28,2)= $database_custom_models.Count
        $modelCount = 0;
        foreach ($model in $database_custom_models.olap_obj_name){
            $buildData.Cells.Item(91, (2 + $modelCount))= $model
            $modelCount++
        }

        $databaseCount = 0
        foreach ($database in $all_databases) {
            $buildData.Cells.Item(99, (2 + $databaseCount))= $database.name
            $buildData.Cells.Item(100, (2 + $databaseCount))= "Size: $($database.Size_MB) MB"
            $databaseCount++
        }

        $buildData.Cells.Item(30,2)= $database_interfaces.ParamValue.Count
        $interfaceCount = 0
        foreach ($interface in $database_interfaces.ParamValue) {
            $buildData.Cells.Item(95, (2 + $interfaceCount))= $interface
            $interfaceCount++
        }

        $buildData.Cells.Item(41,2)= $dbEncryption.is_encrypted
        $buildData.Cells.Item(22,2)= $totalLicenseCount
        $buildData.Cells.Item(40,2)= $cost_threshold.value            
        $buildData.Cells.Item(39,2)= $maxdop.value
        
        Remove-PSSession -Session $sqlSession

        Write-Host "================================================================================" -ForegroundColor Red  
   }

    ##########################
    # PRODUCTION PVE SERVER 
    ##########################
    elseif ($server[0].Name.Substring(($server[0].Name.Length - 5), 3) -eq "pve") {
        Write-Host "THIS IS THE PRODUCTION PVE SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($server[0].Name)"
        Write-Host "Server CPUs: $($server[0].NumCpu)"
        Write-Host "Server RAM: $($server[0].MemoryGB)"  

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[1]) {
            $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString  
            if ($hd.CapacityGB -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# CURRENT VERSION #>
        Write-Host "Current Environment Version" -ForegroundColor Red
        $crVersion = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
            Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\WebServerPlatform"
        }
        Write-Host $crVersion.CrVersion

        <# MAJOR VERSION #>
        Write-Host "Major Version" -ForegroundColor Red
        $majorVersion = $crVersion.CrVersion.Split('.')[0]
        $majorVersion

        <# OPEN SUITE #>
        Write-Host "OpenSuite" -ForegroundColor Red
        $opensuite = Invoke-Command -ComputerName $server[0].Name -Credential $credentials -ScriptBlock {
            if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {

                $software = "*Actian*";
                $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where { $_.DisplayName -like $software }) -ne $null

                if ($installed) {
                    return "Yes"
                }
                
            } else {
                return "No"
            }
        }
        Write-Host "OpenSuite Detected: $($opensuite)"

        <# INTEGRATIONS #>
        Write-Host "Integrations" -ForegroundColor Red
        $PPAdapter = "False"
        $LKAdapter = "False"
        $integrations = Invoke-Command -ComputerName $server[0].Name -Credential $credentials -ScriptBlock {
            param ($database)
            if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*") {
                Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*" | Select-Object -Property PSChildName
            } else {
                return 0
            }
        } -ArgumentList $productionDatabase

        if ($integrations -eq 0) {
            Write-Host "No integrations found in 'HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($productionDatabase)\'"
        } else {
            Write-Host "Number of integrations found: $($integrations.PSChildName.Count)" -ForegroundColor Cyan
            foreach ($x in $integrations.PSChildName) {
                if ($x -like "*ProjectPlace*") {
                    Write-Host "PP ADAPTER FOUND: $($x)"
                    $PPAdapter = "True"
                }
                elseif ($x -like "*PRM_Adapter*") {
                    Write-Host "LK ADAPTER FOUND: $($x)"
                    $LKAdapter= "True"
                } else {
                    Write-Host "Other Integration Identified: $($x)"
                }
            }
        }

        # NEW RELIC #
        Write-Host "New Relic" -ForegroundColor Red
        $newRelic = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
            if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                Write-Host "New Relic has been detected on this server"
                return "Yes"
            } else {
                Write-Host "New Relic was not detected on this server"
                return "No"
            }
        }

            # GET WEB CONFIG #
            $webConfig = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
                return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
            }
            $webConfig = [xml] $webConfig

        <# PRODUCTION URL #>
        Write-Host "Production URL" -ForegroundColor Red
        $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
        Write-Host $environmentURL.value

        <# DNS ALIAS #>
        Write-Host "Production DNS Alias" -ForegroundColor Red            
        $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
        $dnsAlias

        <# REPORT FARM URL #>
        Write-Host "Report Farm URL" -ForegroundColor Red
        $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
        Write-Host $reportfarmURL.value

        <# ENCRYPTED PVMASTER PASSWORD #>
        Write-Host "Encrypted PVMaster Password" -ForegroundColor Red
        $encryptedPVMasterPassword = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUserPassword"} | Select-Object -Property value
        Write-Host $encryptedPVMasterPassword.value
        
        <# UNENCRYPTED PVMASTER PASSWORD #>
        Write-Host "Unencrypted PVMaster Password" -ForegroundColor Red
        $unencryptedPVMasterPassword = Invoke-PassUtil -InputString $encryptedPVMasterPassword.value -Deobfuscation
        Write-Host $unencryptedPVMasterPassword

        <# IP RESTRICTIONS #>
        Write-Host "IP Restrictions on F5" -ForegroundColor Red
        $IPRestrictions = "No"
            
            # Authentication on the F5 #
            $websession =  New-Object Microsoft.PowerShell.Commands.WebRequestSession
            $jsonbody = @{username = $f5Credentials.UserName ; password = $f5Credentials.GetNetworkCredential().Password; loginProviderName='tmos'} | ConvertTo-Json
            $authResponse = Invoke-RestMethodOverride -Method Post -Uri "https://$($f5ip)/mgmt/shared/authn/login" -Credential $f5Credentials -Body $jsonbody -ContentType 'application/json'
            $token = $authResponse.token.token
            $websession.Headers.Add('X-F5-Auth-Token', $Token)

            # Calling data-group REST endpoint and parsing IPRestrictions list #
            $IPRestrictionsList = (Invoke-RestMethod  -Uri "https://$($f5ip)/mgmt/tm/ltm/data-group/internal" -WebSession $websession).Items | 
                Where-Object {$_.name -eq "IPRestrictions"} | Select-Object -Property records

            foreach ($record in $IPRestrictionsList.records) {
                if ($record.name -eq "$($dnsAlias).pvcloud.com") {
                    $IPRestrictions = "Yes"
                    Write-Host "IP restrctions were found for $($dnsAlias).pvcloud.com"
                }
            }

            if ($IPRestrictions -eq "No") {
                Write-Host "No IP restrictions found for $($dnsAlias).pvcloud.com"
            }
        
        <# EXCEL LOGIC AND VARIABLES#>
        $webServerCount++
        $buildData.Cells.Item(24,2)= $webServerCount
        $buildData.Cells.Item(1,2)= $environmentURL.value
        $buildData.Cells.Item(7,2)= $dnsAlias
        $buildData.Cells.Item(46,2)= $reportfarmURL.value
        $buildData.Cells.Item(15,2)= $newRelic
        $buildData.Cells.Item(35,2)= $IPRestrictions
        $buildData.Cells.Item(36,2)= $opensuite
        $buildData.Cells.Item(25,2)= $crVersion.CrVersion
        $buildData.Cells.Item(13,2)= $majorVersion
        $buildData.Cells.Item(19,2)= "True"
        $buildData.Cells.Item(17,2)= $encryptedPVMasterPassword.value
        $buildData.Cells.Item(16,2)= $unencryptedPVMasterPassword
        

        $buildData.Cells.Item(31,2)= $PPAdapter
        $buildData.Cells.Item(32,2)= $LKAdapter
        
        

        $buildData.Cells.Item(56,2)= "$($server[0].Name)"
        $buildData.Cells.Item(56,3)= "$($server[0].NumCpu)"
        $buildData.Cells.Item(56,4)= "$($server[0].MemoryGB)"
        $buildData.Cells.Item(56,5)= $hdStringArray
        $buildData.Cells.Item(56,6)= $diskResize
        $buildData.Cells.Item(56,7)= $task_array

        Write-Host "================================================================================" -ForegroundColor Red  
    }
    
}

#######################
# SANDBOX SERVERS
#######################
Write-Host ":::::::: SANDBOX ENVIRONMENT ::::::::" -ForegroundColor Yellow

$webServerCount = 0
foreach ( $server in $computerObjects_Sandbox){
    

    if ($server[0].Name.Substring(($server[0].Name.Length - 5), 3) -eq "app") {

        #######################
        # SANDBOX APP SERVER 
        #######################
        if ($server[0].Name.Substring(3, 1) -ne 't') {
            Write-Host "THIS IS THE SANDBOX APP SERVER" -ForegroundColor Cyan

            <# CPU/RAM #>
            Write-Host "Server CPU and RAM" -ForegroundColor Red
            Write-Host "Server Name: $($server[0].Name)"
            Write-Host "Server CPUs: $($server[0].NumCpu)"
            Write-Host "Server RAM: $($server[0].MemoryGB)"

            <# HARDDRIVES #>
            Write-Host "Disks and Disk Capacity" -ForegroundColor Red
            $diskResize = "Yes"
            $hdStringArray = ""
            foreach ($hd in $server[1]) {
                $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
                $hdStringArray += "$($hdString)`n"
                Write-Host $hdString  
                if ($hd.CapacityGB -gt 60) {
                    $diskResize = "No"  
                }
            }
            Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

            <# CLUSTER #>
#            Write-Host "Server Cluster" -ForegroundColor Red
#            Write-Host "Cluster Name: $($server[2].Name)"

            <# SCHEDULED TASKS #>
            Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
            $task_array = ""
            foreach ($task in $server[3]){
                Write-Host "Task Name: $($task.TaskName)"
                $task_array += "$($task.TaskName)`n"
            }
            
            <# OPEN SUITE #>
            Write-Host "OpenSuite" -ForegroundColor Red
            $opensuite = Invoke-Command -ComputerName $server[0].Name -Credential $credentials -ScriptBlock {
                if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {

                    $software = "*Actian*";
                    $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where { $_.DisplayName -like $software }) -ne $null

                    if ($installed) {
                        return "Yes"
                    }
                    
                } else {
                    return "No"
                }
            }
            Write-Host "OpenSuite Detected: $($opensuite)"

            <# INTEGRATIONS #>
            Write-Host "Integrations" -ForegroundColor Red
            $PPAdapter = "False"
            $LKAdapter = "False"
            $integrations = Invoke-Command -ComputerName $server[0].Name -Credential $credentials -ScriptBlock {
                param ($database)
                if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*") {
                    Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*" | Select-Object -Property PSChildName
                } else {
                    return 0
                }
            } -ArgumentList $sandboxDatabase

            if ($integrations -eq 0) {
                Write-Host "No integrations found in 'HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($sandboxDatabase)\'"
            } else {
                Write-Host "Number of integrations found: $($integrations.PSChildName.Count)" -ForegroundColor Cyan
                foreach ($x in $integrations.PSChildName) {
                    if ($x -like "*ProjectPlace*") {
                        Write-Host "PP ADAPTER FOUND: $($x)"
                        $PPAdapter = "True"
                    }
                    elseif ($x -like "*PRM_Adapter*") {
                        Write-Host "LK ADAPTER FOUND: $($x)"
                        $LKAdapter= "True"
                    } else {
                        Write-Host "Other Integration Identified: $($x)"
                    }
                }
            }
            
            <# EXCEL LOGIC AND VARIABLES#>
            $buildData.Cells.Item(72,2)= "$($server[0].Name)"
            $buildData.Cells.Item(72,3)= "$($server[0].NumCpu)"
            $buildData.Cells.Item(72,4)= "$($server[0].MemoryGB)"
            $buildData.Cells.Item(72,5)= $hdStringArray
            $buildData.Cells.Item(72,6)= $diskResize
            $buildData.Cells.Item(72,7)= $task_array

            $buildData.Cells.Item(31,3)= $PPAdapter
            $buildData.Cells.Item(32,3)= $LKAdapter
            
            $buildData.Cells.Item(36,3)= $opensuite

            Write-Host "================================================================================" -ForegroundColor Red
        }

        ###############################
        # SANDBOX CTM SERVER (Troux) 
        ###############################
        elseif ($server[0].Name.Substring(3, 1) -eq 't') {
            Write-Host "THIS IS THE SANDBOX TROUX SERVER" -ForegroundColor Cyan

            <# CPU/RAM #>
            Write-Host "Server CPU and RAM" -ForegroundColor Red
            Write-Host "Server Name: $($server[0].Name)"
            Write-Host "Server CPUs: $($server[0].NumCpu)"
            Write-Host "Server RAM: $($server[0].MemoryGB)"     

            <# HARDDRIVES #>
            Write-Host "Disks and Disk Capacity" -ForegroundColor Red
            $diskResize = "Yes"
            $hdStringArray = ""
            foreach ($hd in $server[1]) {
                $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
                $hdStringArray += "$($hdString)`n"
                Write-Host $hdString  
                if ($hd.CapacityGB -gt 60) {
                    $diskResize = "No"  
                }
            }
            Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"       

            <# CLUSTER #>
#            Write-Host "Server Cluster" -ForegroundColor Red
#            Write-Host "Cluster Name: $($server[2].Name)"

            <# SCHEDULED TASKS #>
            Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
            $task_array = ""
            foreach ($task in $server[3]){
                Write-Host "Task Name: $($task.TaskName)"
                $task_array += "$($task.TaskName)`n"
            }
            
            <# EXCEL LOGIC AND VARIABLES#>
            $buildData.Cells.Item(73,2)= "$($server[0].Name)"
            $buildData.Cells.Item(73,3)= "$($server[0].NumCpu)"
            $buildData.Cells.Item(73,4)= "$($server[0].MemoryGB)"
            $buildData.Cells.Item(73,5)= $hdStringArray
            $buildData.Cells.Item(73,6)= $diskResize
            $buildData.Cells.Item(73,7)= $task_array

            Write-Host "================================================================================" -ForegroundColor Red  
        }
   }

    #######################
    # SANDBOX WEB SERVER 
    #######################
    elseif ($server[0].Name.Substring(($server[0].Name.Length - 5), 3) -eq "web") {
        Write-Host "THIS IS THE SANDBOX WEB SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($server[0].Name)"
        Write-Host "Server CPUs: $($server[0].NumCpu)"
        Write-Host "Server RAM: $($server[0].MemoryGB)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[1]) {
            $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString  
            if ($hd.CapacityGB -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# CURRENT VERSION #>
        Write-Host "Current Environment Version" -ForegroundColor Red
        $crVersion = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
            Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\WebServerPlatform"
        }
        Write-Host $crVersion.CrVersion

        <# MAJOR VERSION #>
        Write-Host "Major Version" -ForegroundColor Red
        $majorVersion = $crVersion.CrVersion.Split('.')[0]
        $majorVersion

        # NEW RELIC #
        Write-Host "New Relic" -ForegroundColor Red
        $newRelic = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
            if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                Write-Host "New Relic has been detected on this server"
                return "Yes"
            } else {
                Write-Host "New Relic was not detected on this server"
                return "No"
            }
        }

            # GET WEB CONFIG #
            $webConfig = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
                return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
            }
            $webConfig = [xml] $webConfig

        <# SANDBOX URL #>
        Write-Host "Sandbox URL" -ForegroundColor Red
        $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
        Write-Host $environmentURL.value

        <# DNS ALIAS #>
        Write-Host "Sandbox DNS Alias" -ForegroundColor Red            
        $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
        $dnsAlias

        <# REPORT FARM URL #>
        Write-Host "Report Farm URL" -ForegroundColor Red
        $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
        Write-Host $reportfarmURL.value

        <# IP RESTRICTIONS #>
        Write-Host "IP Restrictions on F5" -ForegroundColor Red
        $IPRestrictions = "No"
            
            # Authentication on the F5 #
            $websession =  New-Object Microsoft.PowerShell.Commands.WebRequestSession
            $jsonbody = @{username = $f5Credentials.UserName ; password = $f5Credentials.GetNetworkCredential().Password; loginProviderName='tmos'} | ConvertTo-Json
            $authResponse = Invoke-RestMethodOverride -Method Post -Uri "https://$($f5ip)/mgmt/shared/authn/login" -Credential $f5Credentials -Body $jsonbody -ContentType 'application/json'
            $token = $authResponse.token.token
            $websession.Headers.Add('X-F5-Auth-Token', $Token)

            # Calling data-group REST endpoint and parsing IPRestrictions list #
            $IPRestrictionsList = (Invoke-RestMethod  -Uri "https://$($f5ip)/mgmt/tm/ltm/data-group/internal" -WebSession $websession).Items | 
                Where-Object {$_.name -eq "IPRestrictions"} | Select-Object -Property records

            foreach ($record in $IPRestrictionsList.records) {
                if ($record.name -eq "$($dnsAlias).pvcloud.com") {
                    $IPRestrictions = "Yes"
                    Write-Host "IP restrctions were found for $($dnsAlias).pvcloud.com"
                }
            }

            if ($IPRestrictions -eq "No") {
                Write-Host "No IP restrictions found for $($dnsAlias).pvcloud.com"
            }

        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(25,3)= $crVersion.CrVersion
        $buildData.Cells.Item(35,3)= $IPRestrictions
        $buildData.Cells.Item(2,2)= $environmentURL.value
        $buildData.Cells.Item(8,2)= $dnsAlias
        

        if ($webServerCount -gt 0){
            $buildData.Cells.Item(78 + ($webServerCount - 1),2)= "$($server[0].Name)"
            $buildData.Cells.Item(78 + ($webServerCount - 1),3)= "$($server[0].NumCpu)"
            $buildData.Cells.Item(78 + ($webServerCount - 1),4)= "$($server[0].MemoryGB)"
            $buildData.Cells.Item(78 + ($webServerCount - 1),5)= $hdStringArray
            $buildData.Cells.Item(78 + ($webServerCount - 1),6)= $diskResize
            $buildData.Cells.Item(78 + ($webServerCount - 1),7)= $task_array
        }
        else {
            $buildData.Cells.Item(71,2)= "$($server[0].Name)"
            $buildData.Cells.Item(71,3)= "$($server[0].NumCpu)"
            $buildData.Cells.Item(71,4)= "$($server[0].MemoryGB)"
            $buildData.Cells.Item(71,5)= $hdStringArray
            $buildData.Cells.Item(71,6)= $diskResize
            $buildData.Cells.Item(71,7)= $task_array
        }

        $webServerCount++
        $buildData.Cells.Item(24,3)= $webServerCount

        Write-Host "================================================================================" -ForegroundColor Red  

   }

    #######################
    # SANDBOX SAS SERVER 
    #######################
    elseif ($server[0].Name.Substring(($server[0].Name.Length - 5), 3) -eq "sas") {
        Write-Host "THIS IS THE SANDBOX SAS SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($server[0].Name)"
        Write-Host "Server CPUs: $($server[0].NumCpu)"
        Write-Host "Server RAM: $($server[0].MemoryGB)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[1]) {
            $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString  
            if ($hd.CapacityGB -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }
        
        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(75,2)= "$($server[0].Name)"
        $buildData.Cells.Item(75,3)= "$($server[0].NumCpu)"
        $buildData.Cells.Item(75,4)= "$($server[0].MemoryGB)"
        $buildData.Cells.Item(75,5)= $hdStringArray
        $buildData.Cells.Item(75,6)= $diskResize
        $buildData.Cells.Item(75,7)= $task_array

        Write-Host "================================================================================" -ForegroundColor Red  
   }

    #######################
    # SANDBOX SQL SERVER 
    #######################
    elseif ($server[0].Name.Substring(($server[0].Name.Length - 5), 3) -eq "sql") {
        Write-Host "THIS IS THE SANDBOX SQL SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($server[0].Name)"
        Write-Host "Server CPUs: $($server[0].NumCpu)"
        Write-Host "Server RAM: $($server[0].MemoryGB)"

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[1]) {
            $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString  
            if ($hd.CapacityGB -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }
        
        <# DATABASE PROPERTIES #>
        Write-Host "$($sandboxDatabase) Properties" -ForegroundColor Red
        $sqlSession = New-PSSession -ComputerName "$($server[0].Name)" -Credential $credentials

            # MAXDOP/THRESHOLD
            Write-Host "Identifying MaxDOP/Threshold..." -ForegroundColor Cyan
            $database_maxdop_threshold = Invoke-Command  -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                '%parallel%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
            } -ArgumentList $server[0].Name
            $maxdop = $database_maxdop_threshold | Where-Object {$_.name -like "cost*"} | Select-Object -property value
            $cost_threshold = $database_maxdop_threshold | Where-Object {$_.name -like "max*"} | Select-Object -property value
            Write-Host "Max DOP --- $($maxdop.value) MB"
            Write-Host "Cost Threshold --- $($cost_threshold.value) MB"
            
            # MIN/MAX MEMORY
            Write-Host "Identifying MIN/MAX Memory..." -ForegroundColor Cyan
            $database_memory = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT name, value, [description] FROM sys.configurations WHERE name like
                '%server memory%' ORDER BY name OPTION (RECOMPILE);" -ServerInstance $server.Name
            } -ArgumentList $server[0].Name 
            $database_memory_max = $database_memory | where-Object {$_.name -like "max*"} | Select-Object -property value
            $database_memory_min = $database_memory | where-Object {$_.name -like "min*"} | Select-Object -property value
            Write-Host "Max Server Memory --- $($database_memory_max.value) MB"
            Write-Host "Min Server Memory --- $($database_memory_min.value) MB"
            
            # DATABASE ENCRYPTION
            Write-Host "Identifying Database Encryption..." -ForegroundColor Cyan
            $database_encryption = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server)
                Invoke-Sqlcmd -Query "SELECT
                db.name,
                db.is_encrypted
                FROM
                sys.databases db
                LEFT OUTER JOIN sys.dm_database_encryption_keys dm
                    ON db.database_id = dm.database_id;
                GO" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name
            $dbEncryption = $database_encryption | Where-Object {$_.name -eq $sandboxDatabase}
            Write-Host "$($dbEncryption.name) --- $($dbEncryption.is_encrypted)"

            # DATABASE SIZE (MAIN)
            Write-Host "Calculating Database Size" -ForegroundColor Cyan
            $database_dbSize = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database)
                GO
                exec sp_spaceused
                GO" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$sandboxDatabase
            Write-Host "$($database_dbSize.database_name) --- $($database_dbSize.database_size)"

            # ALL DATABASES (NAMES AND SIZES in MB)
            Write-Host "Listing All Databases and Sizes (in MB)" -ForegroundColor Cyan
            $all_databases = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "SELECT d.name,
                ROUND(SUM(mf.size) * 8 / 1024, 0) Size_MB
                FROM sys.master_files mf
                INNER JOIN sys.databases d ON d.database_id = mf.database_id
                WHERE d.database_id > 4 -- Skip system databases
                GROUP BY d.name
                ORDER BY d.name" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$sandboxDatabase
            foreach ($database in $all_databases) {
                Write-Host "$($database.name) ---- $($database.Size_MB) MB"
            }

            # CUSTOM MODELS
            Write-Host "Calculating Custom Models..." -ForegroundColor Cyan
            $database_custom_models = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT COUNT(*) FROM ip.olap_properties 
                WHERE bism_ind ='N' 
                AND olap_obj_name 
                NOT like 'PVE%'" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$sandboxDatabase | Select-Object -property olap_obj_name
            foreach ($model in $database_custom_models.olap_obj_name) {
                Write-Host $model
            }  
            
            # INTERFACES
            Write-Host "Identifying Interfaces..." -ForegroundColor Cyan
            $database_interfaces = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT
                s.description JobStreamName,
                j.description JobName,
                j.job_order JobOrder,
                j.job_id JobID,
                p.name ParamName,
                p.param_value ParamValue,
                MIN(r.last_started) JobLastStarted,
                MAX(r.last_finished) JobLastFinished,
                MAX(CONVERT(CHAR(8), DATEADD(S,DATEDIFF(S,r.last_started,r.last_finished),'1900-1-1'),8)) Duration
                FROM ip.job_stream_job j
                INNER JOIN ip.job_stream s
                ON j.job_stream_id = s.job_stream_id
                INNER JOIN ip.job_stream_schedule ss
                ON ss.job_stream_id = s.job_stream_id
                INNER JOIN ip.job_run_status r
                ON s.job_stream_id = r.job_stream_id
                LEFT JOIN ip.job_param p
                ON j.job_id = p.job_id
                WHERE P.Name = 'Command'
                GROUP BY
                s.description,
                j.description,
                j.job_order,
                j.job_id,
                p.name,
                p.param_value;" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$sandboxDatabase
            $database_interfaces.ParamValue  
            
            # LICENSE COUNT
            Write-Host "Calculating License Count..." -ForegroundColor Cyan
            $database_license_count = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database);
                SELECT
                LicenseRole,
                COUNT(UserName) UserCount,
                r.seats LicenseCount
                FROM (
                SELECT
                s1.description LicenseRole,
                s1.structure_code LicenseCode,
                u.active_ind Active,
                u.full_name UserName
                FROM ip.ip_user u
                INNER JOIN ip.structure s
                ON u.role_code = s.structure_code
                INNER JOIN ip.structure s1
                ON s.father_code = s1.structure_code
                WHERE u.active_ind = 'Y'
                ) q
                INNER JOIN ip.ip_role r
                ON q.LicenseCode = r.role_code
                GROUP BY
                LicenseRole,
                LicenseCode,
                r.seats" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$sandboxDatabase
            $licenseProperties = $database_license_count | Select-Object -Property LicenseRole,LicenseCount
            $totalLicenseCount = 0
            foreach ($license in $licenseProperties){
                Write-Output "$($license.LicenseRole): $($license.LicenseCount)"
                $totalLicenseCount += $license.LicenseCount
            }
            Write-Output "Total License Count: $($totalLicenseCount)"

            # PROGRESSING WEB VERSION
            Write-Host "Identifying Progressing Web Version..." -ForegroundColor Cyan
            $database_progressing_web_version = Invoke-Command -Session $sqlSession -ScriptBlock { 
                param ($server,$database)        
                Invoke-Sqlcmd -Query "USE $($database); SELECT TOP 1 sub_release 
                FROM ip.pv_version 
                WHERE release = 'PROGRESSING_WEB'
                ORDER BY seq DESC;" -ServerInstance $server.Name 
            } -ArgumentList $server[0].Name,$sandboxDatabase
            $database_progressing_web_version.sub_release

        <# EXCEL LOGIC AND VARIABLES#>
        $buildData.Cells.Item(12,2)= $server[0].Name.Substring(($server[0].Name.Length - 2), 2)
        $buildData.Cells.Item(44,3)= $database_dbSize.database_size
        $buildData.Cells.Item(43,3)= $database_memory_max.value
        $buildData.Cells.Item(42,3)= $database_memory_min.value

        $buildData.Cells.Item(74,2)= "$($server[0].Name)"
        $buildData.Cells.Item(74,3)= "$($server[0].NumCpu)"
        $buildData.Cells.Item(74,4)= "$($server[0].MemoryGB)"
        $buildData.Cells.Item(74,5)= $hdStringArray
        $buildData.Cells.Item(74,6)= $diskResize
        $buildData.Cells.Item(74,7)= $task_array
        
        $buildData.Cells.Item(26,3)= $database_progressing_web_version.sub_release

        $buildData.Cells.Item(28,2)= $database_custom_models.Count
        $modelCount = 0;
        foreach ($model in $database_custom_models.olap_obj_name){
            $buildData.Cells.Item(91, (2 + $modelCount))= $model
            $modelCount++
        }

        $databaseCount = 0
        foreach ($database in $all_databases) {
            $buildData.Cells.Item(101, (2 + $databaseCount))= $database.name
            $buildData.Cells.Item(102, (2 + $databaseCount))= "Size: $($database.Size_MB) MB"
            $databaseCount++
        }

        $buildData.Cells.Item(30,3)= $database_interfaces.ParamValue.Count
        $interfaceCount = 0
        foreach ($interface in $database_interfaces.ParamValue) {
            $buildData.Cells.Item(96, (2 + $interfaceCount))= $interface
            $interfaceCount++
        }

        $buildData.Cells.Item(41,3)= $dbEncryption.is_encrypted
        $buildData.Cells.Item(22,3)= $totalLicenseCount
        $buildData.Cells.Item(40,3)= $cost_threshold.value
        $buildData.Cells.Item(39,3)= $maxdop.value

        Remove-PSSession -Session $sqlSession

        Write-Host "================================================================================" -ForegroundColor Red  
   }

    #######################
    # SANDBOX PVE SERVER 
    #######################
    elseif ($server[0].Name.Substring(($server[0].Name.Length - 5), 3) -eq "pve") {
        Write-Host "THIS IS THE SANDBOX PVE SERVER" -ForegroundColor Cyan

        <# CPU/RAM #>
        Write-Host "Server CPU and RAM" -ForegroundColor Red
        Write-Host "Server Name: $($server[0].Name)"
        Write-Host "Server CPUs: $($server[0].NumCpu)"
        Write-Host "Server RAM: $($server[0].MemoryGB)"  

        <# HARDDRIVES #>
        Write-Host "Disks and Disk Capacity" -ForegroundColor Red
        $diskResize = "Yes"
        $hdStringArray = ""
        foreach ($hd in $server[1]) {
            $hdString = "$($hd.Name): $($hd.CapacityGB)gb"
            $hdStringArray += "$($hdString)`n"
            Write-Host $hdString  
            if ($hd.CapacityGB -gt 60) {
                $diskResize = "No"  
            }
        }
        Write-Host "Standard Size Disks (less than 60GB): $($diskResize)"

        <# CLUSTER #>
#        Write-Host "Server Cluster" -ForegroundColor Red
#        Write-Host "Cluster Name: $($server[2].Name)"

        <# SCHEDULED TASKS #>
        Write-Host "Scheduled Tasks on Server" -ForegroundColor Red
        $task_array = ""
        foreach ($task in $server[3]){
            Write-Host "Task Name: $($task.TaskName)"
            $task_array += "$($task.TaskName)`n"
        }

        <# CURRENT VERSION #>
        Write-Host "Current Environment Version" -ForegroundColor Red
        $crVersion = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
            Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\WebServerPlatform"
        }
        Write-Host $crVersion.CrVersion

        <# MAJOR VERSION #>
        Write-Host "Major Version" -ForegroundColor Red
        $majorVersion = $crVersion.CrVersion.Split('.')[0]
        $majorVersion
        
        <# OPEN SUITE #>
        Write-Host "OpenSuite" -ForegroundColor Red
        $opensuite = Invoke-Command -ComputerName $server[0].Name -Credential $credentials -ScriptBlock {
            if ((Test-Path -Path "C:\ProgramData\Actian" -PathType Container) -And (Test-Path -Path "F:\Planview\Interfaces\OpenSuite" -PathType Container)) {

                $software = "*Actian*";
                $installed = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where { $_.DisplayName -like $software }) -ne $null

                if ($installed) {
                    return "Yes"
                }
                
            } else {
                return "No"
            }
        }
        Write-Host "OpenSuite Detected: $($opensuite)"

        <# INTEGRATIONS #>
        Write-Host "Integrations" -ForegroundColor Red
        $PPAdapter = "False"
        $LKAdapter = "False"
        $integrations = Invoke-Command -ComputerName $server[0].Name -Credential $credentials -ScriptBlock {
            param ($database)
            if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*") {
                Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($database)\*" | Select-Object -Property PSChildName
            } else {
                return 0
            }
        } -ArgumentList $sandboxDatabase

        if ($integrations -eq 0) {
            Write-Host "No integrations found in 'HKLM:\SOFTWARE\WOW6432Node\Planview\Integrations\$($sandboxDatabase)\'"
        } else {
            Write-Host "Number of integrations found: $($integrations.PSChildName.Count)" -ForegroundColor Cyan
            foreach ($x in $integrations.PSChildName) {
                if ($x -like "*ProjectPlace*") {
                    Write-Host "PP ADAPTER FOUND: $($x)"
                    $PPAdapter = "True"
                }
                elseif ($x -like "*PRM_Adapter*") {
                    Write-Host "LK ADAPTER FOUND: $($x)"
                    $LKAdapter= "True"
                } else {
                    Write-Host "Other Integration Identified: $($x)"
                }
            }
        }

        # NEW RELIC #
        Write-Host "New Relic" -ForegroundColor Red
        $newRelic = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
            if (Test-Path -Path "C:\ProgramData\New Relic" -PathType Container ) {
                Write-Host "New Relic has been detected on this server"
                return "Yes"
            } else {
                Write-Host "New Relic was not detected on this server"
                return "No"
            }
        }

            # GET WEB CONFIG #
            $webConfig = Invoke-Command -ComputerName "$($server[0].Name)" -Credential $credentials -ScriptBlock {
                return Get-Content -Path "F:\Planview\MidTier\ODataService\Web.config"
            }
            $webConfig = [xml] $webConfig

        <# SANDBOX URL #>
        Write-Host "Sandbox URL" -ForegroundColor Red
        $environmentURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "PveUrl"} | Select-Object -Property value
        Write-Host $environmentURL.value

        <# DNS ALIAS #>
        Write-Host "Sandbox DNS Alias" -ForegroundColor Red            
        $dnsAlias = ($environmentURL.value.Split('//')[2]).Split('.')[0] 
        $dnsAlias

        <# REPORT FARM URL #>
        Write-Host "Report Farm URL" -ForegroundColor Red
        $reportfarmURL = $webConfig.configuration.appSettings.add | Where-Object {$_.key -eq "Report_Server_Web_Service_URL"} | Select-Object -Property value
        Write-Host $reportfarmURL.value

        <# IP RESTRICTIONS #>
        Write-Host "IP Restrictions on F5" -ForegroundColor Red
        $IPRestrictions = "No"
            
            # Authentication on the F5 #
            $websession =  New-Object Microsoft.PowerShell.Commands.WebRequestSession
            $jsonbody = @{username = $f5Credentials.UserName ; password = $f5Credentials.GetNetworkCredential().Password; loginProviderName='tmos'} | ConvertTo-Json
            $authResponse = Invoke-RestMethodOverride -Method Post -Uri "https://$($f5ip)/mgmt/shared/authn/login" -Credential $f5Credentials -Body $jsonbody -ContentType 'application/json'
            $token = $authResponse.token.token
            $websession.Headers.Add('X-F5-Auth-Token', $Token)

            # Calling data-group REST endpoint and parsing IPRestrictions list #
            $IPRestrictionsList = (Invoke-RestMethod  -Uri "https://$($f5ip)/mgmt/tm/ltm/data-group/internal" -WebSession $websession).Items | 
                Where-Object {$_.name -eq "IPRestrictions"} | Select-Object -Property records

            foreach ($record in $IPRestrictionsList.records) {
                if ($record.name -eq "$($dnsAlias).pvcloud.com") {
                    $IPRestrictions = "Yes"
                    Write-Host "IP restrctions were found for $($dnsAlias).pvcloud.com"
                }
            }

            if ($IPRestrictions -eq "No") {
                Write-Host "No IP restrictions found for $($dnsAlias).pvcloud.com"
            }
        
        <# EXCEL LOGIC AND VARIABLES#>
        $webServerCount++
        $buildData.Cells.Item(24,3)= $webServerCount
        $buildData.Cells.Item(2,2)= $environmentURL.value   
        $buildData.Cells.Item(8,2)= $dnsAlias
        $buildData.Cells.Item(35,3)= $IPRestrictions
        $buildData.Cells.Item(36,3)= $opensuite
        $buildData.Cells.Item(25,3)= $crVersion.CrVersion
        $buildData.Cells.Item(19,2)= "True"

        $buildData.Cells.Item(76,2)= "$($server[0].Name)"
        $buildData.Cells.Item(76,3)= "$($server[0].NumCpu)"
        $buildData.Cells.Item(76,4)= "$($server[0].MemoryGB)"
        $buildData.Cells.Item(76,5)= $hdStringArray
        $buildData.Cells.Item(76,6)= $diskResize
        $buildData.Cells.Item(76,7)= $task_array

        $buildData.Cells.Item(31,3)= $PPAdapter
        $buildData.Cells.Item(32,3)= $LKAdapter

        Write-Host "================================================================================" -ForegroundColor Red  
    }

}

############################################################
# CLOSES THE JUMPBOX CONNECTION
############################################################
Remove-PSSession -Session $session

############################################################
# SAVE AND CLOSE
# SAVES AND CLOSES THE EXCEL WORKBOOK  
############################################################
$excelfile.Save()
$excelfile.Close()
}
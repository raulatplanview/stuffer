Import-Module \\scripthost\modules\pvadmin
Import-Module SQLSERVER
Import-Module F5-LTM

<# EXCEL FILE #>
Get-ChildItem -Path "C:\Users\$($admin)$($domain)\Planview, Inc\E1 Build Cutover - Documents\Customer Builds\1_FolderTemplate\18\" -Filter "InPlace*" | Copy-Item -Destination "C:\Users\$($admin)$($domain)\Desktop"
$excelFilePath = Get-ChildItem -Path "C:\Users\$($admin)$($domain)\Desktop\" -Filter "InPlace*" | ForEach-Object {$_.FullName}


<# VSPHERE #>
$scriptBlock = {
    
    param ($environments, $vSphereServer, $vSphereCredentials)

    Connect-VIServer -Server $vSphereServer -Credential $vSphereCredentials
    
    $environments = $environments.GetEnumerator() | Sort-Object -Property Name

    # AT A GLANCE (ON THE CONSOLE)#
    Write-Host "`nEnvironments and Servers Found:" -ForegroundColor Red
    foreach ($e in $environments){

        Write-Host "$($e.Name)" -ForegroundColor Cyan
        Write-Host "$(($e.Value).TrimEnd(',').Split(',') | Sort-Object) `n" -ForegroundColor Yellow

    }

    $environmentsMaster = @()
    foreach ($e in $environments){
        
        $environmentName = $e.Name
        $servers = ($e.Value).TrimEnd(',').Split(',') | Sort-Object

        # ENVIRONMENT ARRAY #
        New-Variable -Name $environmentName -Value @() -Force
        $environment =  Get-Variable -Name $environmentName

        Write-Host "$($environmentName) Environment-------------------------" -ForegroundColor Red
        
        foreach ($serverName in $servers) {
            
            # SERVER ARRAY #
            New-Variable -Name $serverName -Value @() -force
            $server = Get-Variable -Name $serverName
        

            Write-Host "Connected to --- $serverName" -ForegroundColor Green
            
            Write-Host "Collecting CPU and memory information..." -ForegroundColor Cyan 
            $specs = Get-VM -Name $serverName | Select-Object -Property Name, NumCpu, MemoryGB

            Write-Host "Collecting disk information..." -ForegroundColor Cyan
            $disks = Get-VM -Name $serverName | Get-Harddisk

            Write-Host "Identifying server cluster...`n" -ForegroundColor Cyan 
            $cluster = Get-Cluster -VM $serverName | Select-Object -Property Name

            $server = (($specs), ($disks), ($cluster))

            $environment = $server

            Write-Host $environment[0][1]

        }

        #$environmentsMaster += $environment

        <# STORES SERVER ATTRIBUTES COLLECTED IN $computerObjects (NESTED ARRAY) #>
        #$ += @(($specs), ($disks), ($cluster))
    } 

    return $environments
}

$environmentsMaster = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $environments, $vSphereServer, $vSphereCredentials
Remove-PSSession -Session $session
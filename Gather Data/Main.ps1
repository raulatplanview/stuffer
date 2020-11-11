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
    
    <# SORTING ENVRIRONMENTS#>
    $environments = $environments.GetEnumerator() | Sort-Object -Property Name

    <# AT A GLANCE ON THE CONSOLE #>
    Write-Host "`nEnvironments and Servers Found:" -ForegroundColor Red

    foreach ($e in $environments){

        Write-Host "$($e.Name)" -ForegroundColor Cyan
        Write-Host "$(($e.Value).TrimEnd(',').Split(',') | Sort-Object) `n" -ForegroundColor Yellow

    }

    <# SORTING SERVERS #>
    foreach ($e in $environments){
        
        $environmentName = $e.Name
        $servers = ($e.Value).TrimEnd(',').Split(',') | Sort-Object
        
        Write-Host "$($environmentName) Environment-------------------------" -ForegroundColor Red
    
        
        foreach ($server in $servers) {
    
            Write-Host "Connected to --- $($server)" -ForegroundColor Green
            
            Write-Host "Collecting CPU and memory information..." -ForegroundColor Cyan 
            $specs = Get-VM -Name $server | Select-Object -Property Name, NumCpu, MemoryGB

            Write-Host "Collecting disk information..." -ForegroundColor Cyan
            $disks = Get-VM -Name $server | Get-Harddisk

            Write-Host "Identifying server cluster...`n" -ForegroundColor Cyan 
            $cluster = Get-Cluster -VM $server | Select-Object -Property Name

        }

        

        <# STORES SERVER ATTRIBUTES COLLECTED IN $computerObjects (NESTED ARRAY) #>
        #$ += @(($specs), ($disks), ($cluster))
    } 

    return $environments
}

$environmentsMaster = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $environments, $vSphereServer, $vSphereCredentials

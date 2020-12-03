#Import-Module \\scripthost\modules\pvadmin
#Import-Module SQLSERVER
#Import-Module F5-LTM

<# AT A GLANCE #> 
$environments = $environments.GetEnumerator() | Sort-Object -Property Name

Write-Host "`nEnvironments and Servers Found:" -ForegroundColor Red
foreach ($e in $environments){

    Write-Host "$($e.Name)" -ForegroundColor Cyan
    Write-Host "$(($e.Value).TrimEnd(',').Split(',') | Sort-Object) `n" -ForegroundColor Yellow

}

<# VSPHERE #>
$environmentsMaster = @()
Connect-VIServer -Server $vSphereServer -Credential $vSphereCredentials

foreach ($e in $environments){
    
    $environmentName = $e.Name
    $servers = ($e.Value).TrimEnd(',').Split(',') | Sort-Object

    # ENVIRONMENT ARRAY #
    New-Variable -Name $environmentName -Value @() -Force
    $environment =  Get-Variable -Name $environmentName
    $environment.Value += @(($environment.Name))

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

        $server = @(($specs), ($disks), ($cluster))
        $environment.Value += @(,($server))

    }

    $environmentsMaster += @(,($environment.Value))

} 

$environmentsMaster

<# EXCEL FILE #>
#Get-ChildItem -Path "C:\Users\$($admin)\Planview, Inc\E1 Build Cutover - Documents\Customer Builds\1_FolderTemplate\18\" -Filter "InPlace*" | Copy-Item -Destination "C:\Users\$($admin)\Desktop"
#$excelFilePath = Get-ChildItem -Path "C:\Users\$($admin)\Desktop\" -Filter "InPlace*" | ForEach-Object {$_.FullName}

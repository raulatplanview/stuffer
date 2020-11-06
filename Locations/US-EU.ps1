Import-Module \\scripthost\modules\pvadmin
Import-Module SQLSERVER
Import-Module F5-LTM

<# SETTING CREDENTIALS #>
Write-Host "Sign-in with your 'Planview\<a-admin>' account:" -ForegroundColor Magenta
$aAdmin = "a-$($admin)"
$credentials = Get-Credential "Planview\$($aAdmin)"
$vSphereCredentials = New-Object System.Management.Automation.PSCredential ($aAdmin, $credentials.Password)
$f5Credentials = $vSphereCredentials

<# SPECIFYING THE CUSTOMER #>
$option = "2" #Read-Host "Select a region: 1-SG or 2-LN"
$customerName = "BAES: BAE Systems" #Read-Host "Enter the customer OU name"
$customerCode = "bae" #Read-Host "Enter the customer code"

<# SETTING REGIONAL OPTIONS #>
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

<# COLLECTING AD OBJECTS #>
$session = New-PSSession -ComputerName $jumpbox -Authentication Credssp -Credential $credentials
Write-Host "Connecting to Active Directory..." -ForegroundColor Gray
$AD_OU = Get-ADOrganizationalUnit -Filter { Name -like $customerName } -Server $ad_server
$distinguishedNames = (Get-ADComputer -Filter * -SearchBase "$($AD_OU.DistinguishedName)" -Server $ad_server).DistinguishedName

<# SORTING ENVIRONMENTS AND RESPECTIVE SERVERS #>
$environments = @()
foreach ($server in $distinguishedNames) {
    
    $folders = $server.Split(',')
    $serverName = $folders[0].substring(3)

    foreach ($folder in $folders){

        $ouName = $folder.substring(3)

        if ($ouName -like "prod*" -Or $ouName -like "sand*" -Or $ouName -like "pre*"){
            
            if ($environments -NotContains (Get-Variable -Name $ouName)){
                
                New-Variable -Name $ouName -Value @() -Force
                $environments += @(Get-Variable -Name $ouName)

            }

            foreach ($e in $environments) {

                if ($e.Name -eq $ouName){

                    $e.Value += $serverName
                    
                }

            }

        }

    }

}

<# ENVIRONMENT/SERVERNAME CHECK #>
for ($x=0; $x -lt $environments.Length; $x++) {

    Write-Host $environments[$x].Name -ForegroundColor "Cyan"

    for ($y=0; $y -lt $environments[$x].Length; $y++) {
        
        Write-Host $environments[$x][$y].Value -ForegroundColor "Yellow"
    
    }

    Write-Host "-------------"

}

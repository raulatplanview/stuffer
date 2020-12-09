<# SETTING CREDENTIALS #>
Write-Host "Sign-in with your 'Planview\<a-admin>' account:" -ForegroundColor Magenta
$aAdmin = "a-$($admin)"
$credentials = Get-Credential "Planview\$($aAdmin)"
$vSphereCredentials = New-Object System.Management.Automation.PSCredential ($aAdmin, $credentials.Password)
$f5Credentials = $vSphereCredentials


<# SETTING REGIONAL OPTIONS #>
switch($option) {

    1 {
        $ad_server = "SGCVMADC12.us.planview.world";
        $vSphereServer = "sgivmvcsvr06.us.planview.world"; 
        $dataCenterLocation = "sg"; 
        $reportFarm = "https://usreportfarm03.pvcloud.com/ReportServer";
        $f5ip = "10.132.81.2";
        break
    } 

    2 {
        $ad_server = "LNCVMADC05.eu.planview.world"; 
        $vSphereServer = "lnivmvcsvr06.eu.planview.world"; 
        $dataCenterLocation = "ln"; 
        $reportFarm = "https://eureportfarm03.pvcloud.com/ReportServer";
        $f5ip = "10.60.2.171";
        break
    }

}

<# COLLECTING AD OBJECTS #>
Write-Host "Connecting to Active Directory..." -ForegroundColor Gray
$AD_OU = Get-ADOrganizationalUnit -Filter { Name -like $customerName } -Server $ad_server
$distinguishedNames = (Get-ADComputer -Filter * -SearchBase "$($AD_OU.DistinguishedName)" -Server $ad_server).DistinguishedName

<# SETTING ENVIRONMENTS/SERVERS #>
$environments = @{}
foreach ($server in $distinguishedNames) {
    
    $folders = $server.Split(',')

    foreach ($folder in $folders) {

        $ouName = $folder.substring(3)

        if ($ouName -like "prod*" -Or $ouName -like "sand*" -Or $ouName -like "pre*") {
        
            $serverName = $folders[0].substring(3)

            if (-not $environments.ContainsKey($ouName)) {
                
                # ENVIRONMENT NAME (KEY) #
                $environments.Add($ouName,"")

            }

            # SERVER NAME (VALUE) #
            $environments[$ouName] += "$($serverName),"

        } 

    }

}

<# TO 'Logic' #>
. "$($stufferDirectory)\Logic\US-EU Array.ps1" $environments $ad_server $dataCenterLocation $reportFarm $f5ip $f5Credentials $vSphereCredentials $vSphereServer $aAdmin $customerCode $AD_OU $customerName $credentials
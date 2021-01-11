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
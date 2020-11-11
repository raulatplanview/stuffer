<# ADMIN #>
$admin = "rreyes"
$domain = ".CORPORATE"

<# STUFFER PATH #>
$stufferDirectory = $MyInvocation.MyCommand.Path | Split-Path


<# TO LOCATION #>
$stufferType = "2" #Read-Host "Where is the environment located: 1-AWS or 2-US/EU"

if ($stufferType -eq 1) {

    . "$($stufferDirectory)\Locations\AWS.ps1" $stufferDirectory $admin $domain

}
else {

    $option = "2" #Read-Host "Select a region: 1-SG or 2-LN"
    $customerName = "BAES: BAE Systems" #Read-Host "Enter the customer OU name"
    $customerCode = "bae" #Read-Host "Enter the customer code"

    . "$($stufferDirectory)\Locations\US-EU.ps1" $stufferDirectory $admin $domain $option $customerName $customerCode
}
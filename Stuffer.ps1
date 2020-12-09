<# ADMIN #>
$admin = "rreyes"

<# STUFFER PATH #>
$stufferDirectory = $MyInvocation.MyCommand.Path | Split-Path


<# TO LOCATION #>
$stufferType = "2" #Read-Host "Where is the environment located: 1-AWS or 2-US/EU"

if ($stufferType -eq 1) {

    . "$($stufferDirectory)\Locations\AWS.ps1" $stufferDirectory $admin 

}
else {

    $option = "2" #Read-Host "Select a region: 1-SG or 2-LN"
    $customerName = "PRIM: Primark Ireland" #Read-Host "Enter the customer OU name"
    $customerCode = "pki" #Read-Host "Enter the customer code"

    . "$($stufferDirectory)\Locations\US-EU.ps1" $stufferDirectory $admin $option $customerName $customerCode
}

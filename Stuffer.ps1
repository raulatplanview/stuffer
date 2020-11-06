$admin = "rreyes"
$stufferDirectory = $MyInvocation.MyCommand.Path | Split-Path

$stufferType = "2" #Read-Host "Where is the environment located: 1-AWS or 2-US/EU"
if ($stufferType -eq 1) {
    . '.\Locations\AWS.ps1' $stufferDirectory $admin
}
else {
    . '.\Locations\US-EU.ps1' $stufferDirectory $admin
}
<# STUFFER PATH #>
$stufferDirectory = $MyInvocation.MyCommand.Path | Split-Path

<# PRIMARY CUSTOMER INFORMATION #>
$type = Read-Host "Select a build type: 1-InPlace 2-New Build/Upgrade"
$option = Read-Host "Select a region: 1-SG or 2-LN"
$customerName = Read-Host "Enter the customer OU name"
$customerCode = Read-Host "Enter the customer code"
$environmentsOption = Read-Host "Select Environments (Default is Production and Sandbox): 1-Default or 2-Specify"

switch ($environmentsOption) {
    1 {
        $slot1 = "Production";
        $slot2 = "Sandbox";
        break
    }
    2 {
        $slot1 = Read-Host "Environment to replace 'Production'";
        $slot2 = Read-Host "Environment to replace 'Sandbox'";
        break
    }
}
. "$($stufferDirectory)\Locations\US-EU.ps1" 


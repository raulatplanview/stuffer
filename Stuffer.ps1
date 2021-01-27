<# STUFFER PATH #>
$stufferDirectory = $MyInvocation.MyCommand.Path | Split-Path

<# PRIMARY CUSTOMER INFORMATION #>
$option = Read-Host "Select a region: 1-SG or 2-LN"
$customerName = Read-Host "Enter the customer OU name"
$customerCode = Read-Host "Enter the customer code"

. "$($stufferDirectory)\Locations\US-EU.ps1" 


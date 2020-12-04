
Write-Host "Hello"
<# EXCEL FILE #>
Get-ChildItem -Path "F:\Build Templates\InPlace" -Filter "InPlace*" | Copy-Item -Destination "C:\Users\$($aAdmin)\Desktop"
$excelFilePath = Get-ChildItem -Path "C:\Users\$($aAdmin)\Desktop\" -Filter "InPlace*" | ForEach-Object {$_.FullName}

Write-Host $excelFilePath 
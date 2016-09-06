Import-Module BitsTransfer

$url ="http://metroforsteam.com/"

$currentPath = (Get-Item -Path ".\" -Verbose).FullName

Write-Host "Querying for download link of latest version."

$r=Invoke-WebRequest $url

$downloadUrl = ($r.Links | Where-Object{$_.class -like "download"}).href

$output = $downloadUrl.Split("/")[-1]

Write-Host "Found $output."

Write-Host "Starting download to '$currentPath\$output'."

Start-BitsTransfer -Source $downloadUrl -Destination $output

Write-Host "Download complete. Commencing extraction."

$shell = new-object -com shell.application

$MetroForSteam = $shell.NameSpace("$currentPath\$output").items() | Where-Object{$_.Name -like "Metro for Steam"}

$shell.Namespace($currentPath).copyhere($MetroForSteam, 0x10)

Remove-Item $output

Write-Host "Metro For Steam updated, changes will apply with next Steam start."

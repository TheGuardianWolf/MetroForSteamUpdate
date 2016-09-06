# Metro for Steam Update
# Written by TheGuardianWolf

# Set url to home page of Metro for Steam skin
$url ="http://metroforsteam.com/"

# Get the current path of the script
$currentPath = (Get-Item -Path ".\" -Verbose).FullName

Write-Host "Querying for download link of latest version."

# Start a new webclient object
$web = New-Object Net.WebClient

# Store web page HTML into request object
$r = $web.DownloadString($url)

$matches

# Use regex to find the download link
$downloadMatch = $r -match '<a.*?href="(\S+\/(\S+))".*?class="download"[^>]*>'

# If we can't find a match, throw an error
if (!$downloadMatch)
{
    Throw "Cannot find download link."
}

# Get the captured regex from the string
$downloadUrl = $matches[1]

# Get the output file name from the captured regex
$output = $matches[2]

Write-Host "Found $output."

Write-Host "Starting download to '$currentPath\$output'."

$r = $web.DownloadFile($downloadURL, $output)

Write-Host "Download complete. Commencing extraction."

$shell = new-object -com shell.application

$MetroForSteam = $shell.NameSpace("$currentPath\$output").items() | Where-Object{$_.Name -like "Metro for Steam"}

$shell.Namespace($currentPath).copyhere($MetroForSteam, 0x10)

Remove-Item $output

Write-Host "Metro For Steam updated, changes will apply with next Steam start."

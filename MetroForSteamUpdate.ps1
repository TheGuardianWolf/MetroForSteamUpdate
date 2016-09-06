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

# Use regex to find the download link
$downloadMatch = $r -match '<a.*?href="(\S+\/((\S+)\.\S+))".*?class="download"[^>]*>'

# If we can't find a match, throw an error
if (!$downloadMatch)
{
    Throw "Cannot find download link."
}

# Get the captured regex from the string
$downloadUrl = $Matches[1]

# Get the output file name from the captured regex
$output = $Matches[2]

# Get the version
$version = $Matches[3]

# Check our current version
if (Test-Path "$currentPath\Metro for Steam\$version.version")
{
    Write-Host "Current version already up to date."
    exit
}

if (Test-Path "$currentPath\Metro for Steam\")
{
	Get-ChildItem "$currentPath\Metro for Steam\" -recurse -include *.version | foreach ($_) {remove-item $_.fullname}
}

$_newItem = New-Item "$currentPath\Metro for Steam\$version.version" -type file -force

Write-Host "Found $output."

Write-Host "Starting download to '$currentPath\$output'."

# Download file to output
$r = $web.DownloadFile($downloadURL, $output)

Write-Host "Download complete. Commencing extraction."

# Use explorer to extract zip
$shell = new-object -com shell.application

# Overwrite existing
$MetroForSteam = $shell.NameSpace("$currentPath\$output").items() | Where-Object{$_.Name -like "Metro for Steam"}

$shell.Namespace($currentPath).copyhere($MetroForSteam, 0x10)

Remove-Item $output

Write-Host "Metro For Steam updated, changes will apply with next Steam start."

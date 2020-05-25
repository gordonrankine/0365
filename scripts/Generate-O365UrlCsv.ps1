<#

.SYNOPSIS
Downloads the URL's from the json file supplied by Microsoft.

.DESCRIPTION
Downloads the URL's from the json file supplied by Microsoft then converts to a csv file. https://docs.microsoft.com/en-gb/Office365/Enterprise/office-365-ip-web-service

.PARAMETER outDir
This is the directory where the files are stored. If the directory doesn't exist it will be created.

.EXAMPLE
.\Generate-O365UrlCsv -outDir "c:\temp"
Generates a csv in c:\temp. The csv will be populated with the O365 Urls.

.LINK
https://github.com/gordonrankine/

.NOTES
License:            MIT License
Compatibility:      Windows 10
Author:             Gordon Rankine
Date:               25/05/2020
Version:            1.0
PSSscriptAnalyzer:  Pass

#>

[cmdletbinding()]

Param(

    [Parameter(Mandatory=$True, Position=1, HelpMessage="This is the directory for the output files. It will be created if it doesn't exist.")]
    [string]$outDir
)

function fnCreateDir {

<#

.SYNOPSIS
Creates a directory.

.DESCRIPTION
Creates a directory.

.PARAMETER outDir
This is the directory to be created.

.EXAMPLE
.\Create-Directory.ps1 -outDir "c:\test"
Creates a directory called "test" in c:\

.EXAMPLE
.\Create-Directory.ps1 -outDir "\\COMP01\c$\test"
Creates a directory called "test" in c:\ on COMP01

.LINK
https://github.com/gordonrankine/powershell

.NOTES
    License:            MIT License
    Compatibility:      Windows 7 or Server 2008 and higher
    Author:             Gordon Rankine
    Date:               13/01/2019
    Version:            1.1
    PSSscriptAnalyzer:  Pass

#>

    [CmdletBinding()]

        Param(

        # The directory to be created.
        [Parameter(Mandatory=$True, Position=0, HelpMessage='This is the directory to be created. E.g. C:\Temp')]
        [string]$outDir

        )

        # Create out directory if it doesnt exist
        if(!(Test-Path -path $outDir)){
            if(($outDir -notlike "*:\*") -and ($outDir -notlike "*\\*")){
            Write-Output "[ERROR]: $outDir is not a valid path. Script terminated."
            break
            }
                try{
                New-Item $outDir -type directory -Force -ErrorAction Stop | Out-Null
                Write-Output "[INFO] Created output directory $outDir"
                }
                catch{
                Write-Output "[ERROR]: There was an issue creating $outDir. Script terminated."
                Write-Output ($_.Exception.Message)
                Write-Output ""
                break
                }
        }
        # Directory already exists
        else{
        Write-Output "[INFO] $outDir already exists."
        }

} # end fnCreateDir

Clear-Host

fnCreateDir $outDir

$sw = [system.diagnostics.stopwatch]::StartNew()
$timestamp = Get-Date -UFormat %Y%m%d%H%M
$guid = ([guid]::NewGuid()).guid
Write-Output "[INFO] Session GUID: $guid"

    # Download json file
    try{
    $o365URL = "https://endpoints.office.com"
    $webRequest = "$o365Url/endpoints/worldwide?clientrequestid=$guid"
    Write-Output "[INFO] Downloading information from $webRequest"
    Invoke-WebRequest -uri ("$webRequest") -OutFile "$outDir\O365_Endpoints_$timestamp.json" -ErrorAction SilentlyContinue
    }
    catch{
    Write-Output "[ERROR]: Unable to download information. Script terminated."
    Write-Output "[ERROR]: $_.Exception.Message"
    break
    }

    # Import json file
    try{
    Write-Output "[INFO] Importing json file from $outDir\O365_Endpoints_$timestamp.json."
    $endpoints = Get-Content -Path "$outDir\O365_Endpoints_$timestamp.json" -ErrorAction SilentlyContinue -Force | ConvertFrom-Json -ErrorAction SilentlyContinue
    }
    catch{
    Write-Output "[ERROR]: Unable to download information. Script terminated."
    Write-Output "[ERROR]: $_.Exception.Message"
    break
    }

    # Create csv files
    try{
    Write-Output "[INFO] Creating csv files for URL's."
    "Service,URL,Express Route,Category,Required" | Out-File "$outDir\O365_URLs_$timestamp.csv" -Encoding ascii -Force
    }
    catch{
    Write-Output "[ERROR]: Unable to create csv files. Script terminated."
    Write-Output "[ERROR]: $_.Exception.Message"
    break
    }

    Write-Output "[INFO] Getting URL's. This can take a few moments, please wait... "
    foreach($endpoint in $endpoints){

        # Split out URLs and write to csv
        foreach($endPointUrl in $endpoint.urls){
        '"' + $endpoint.serviceAreaDisplayName + '"'  + ","  + $endPointUrl + ","  + $endpoint.expressRoute + ","  + $endpoint.category + ","  + `
        $endpoint.required | Out-File "$outDir\O365_URLs_$timestamp.csv" -Encoding ascii -Force -Append
        }

    }

Write-Output "[INFO] Output at: $outDir\O365_URLs_$timestamp.csv"
Write-Output "[INFO] Script complete in $($sw.Elapsed.Hours) hours, $($sw.Elapsed.Minutes) minutes, $($sw.Elapsed.Seconds) seconds."
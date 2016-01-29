<# 
    Download a file from a given URL and save it to disk. 
    This script is based on one written by Jourdan Templeton, April 2015, 
    Modified by Josh Smith for Demos

#>

$url = "http://prod.method52.casmconsulting.co.uk/output/irish-examiner/download/download-candidates-tweets/cantweets.csv"
$output = "D:\Qlikview\Projects\Irish Examiner\1_Data\data.csv"

# TODO: Make authenticating optional

function getFileFromURL{
    param(
        [System.Management.Automation.PSCredential] $webCredentials,
        [string] $authURL, # The URL to authenticate to, if needed.
        [string] $fileURL,
        [string] $outputPath
    )

    $start_time = Get-Date
    $client = New-Object System.Net.WebClient

    # Set up basic authentication for request
    $credentialCache = new-object System.Net.CredentialCache
    $credentialCache.Add($authURL, "Basic", $webCredentials)
    $client.Credentials = $credentialCache

    echo "Fetching data from URL: $fileURL"

    # Download the file
    $client.DownloadFile($fileURL, $outputPath)

    echo " - File downloaded to $outputPath."
    echo " - Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"  

}


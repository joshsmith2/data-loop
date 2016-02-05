<#

    Perform reload of Qlikview document, both conversion and app steps.
    At the moment this is configured to point to the Irish Examiner, but 
    project by project configuration is coming. 

#>


# Open a QVW file, reload it, and close it again
function reloadQVW{
    param (
        [string] $qvwPath,
        [string] $qvExe = 'C:\Program Files\QlikView\qv.exe'
    )
    & $qvExe $qvwPath /R
}

function checkFileUpdated {
    param (
        [string] $path, # Path to file to check
        [string] $qvd, # QVD being monitored, for debugging purposes.
        [int] $frequency = 1, # How often to check, in seconds
        [int] $timeout = 60, # Throw an exception if file is not updated before timeout checks.
        [System.DateTime] $since # Passes if file modified since this time.
    )
    $startTime = Get-Date
    for($i=0; $i -le $timeout; $i++){
        $time = $i * $frequency

        # See if that file exists. Don't throw an error if not.
        try {
            $lastModified = (Get-Item $path).LastWriteTime
        } catch [Exception] {
            echo "ERROR: No file exists at $pdfPath"
            echo "Full exception trace: "
            echo $_.Exception | format-list -force
        }
        if ($lastModified -gt $since){
            echo " - $path modified after $time seconds."
            # Add a second on the safe side, to allow QV to close the file and such.
            Start-Sleep -s 1
            break
        } 
        Start-Sleep -s $frequency
        # Throw an exception if the script is reaching timeout
        if ($i -eq $timeout - 1){
            throw " ERROR: The QVD at $qvd has been running for $time seconds without updating $path. You might want to have a look at it. The data loop process has been stopped."
        }
    }
}


# Press a button in a Qlikview document defined by ID. 
# Throw an exception if this not possible
function pressButton{
    param (
        [string] $buttonID,
        $qvDocument
    )
    try {
        $qvDocument.ClearAll()
        $button = $qvDocument.GetSheetObject($buttonID)
        $button.Press()
        echo "$buttonID pressed"
    } catch [Exception] {
        echo "ERROR: Couldn't find or press the button with ID: $buttonID"
        echo "Full exception trace: "
        echo $_.Exception | format-list -force
    }
}

# Once the conversion QVD has run, reload the App QVD
# and press the button to output a pdf.
function ExportPDF {
    param (
        [string] $QVDToExportFrom
    )
    
    $qvComObject = New-Object -comobject QlikTech.Qlikview
    $qvDoc = $qvComObject.OpenDoc($QVDToExportFrom)
    $qvDoc.Activate()
    $qvDoc.Reload()

    Start-Sleep -s 1

    pressButton -qvDocument $qvDoc -buttonID "PrintAllCandidates"
    pressButton -qvDocument $qvDoc -buttonID "PrintCork"
    pressButton -qvDocument $qvDoc -buttonID "PrintIndependents"


    # Tidy up
    $qvDoc.ClearAll()
    $qvDoc.Save()
    $qvDoc.CloseDoc()
    $qvComObject.Quit()
}

function reloadAndExport {
    param (
        [string] $QVDToReload,
        [int] $attempts
    )

    for ($i=1; $i -lt $attempts + 1; $i++){
        try{
            ExportPDF -QVDToExportFrom $QVDToReload
            break
        }catch [Exception] {
            echo " - Qlikview snafu encountered on try $i. Trying again in a second."
            Start-Sleep -s 1
            if ($i -eq $attempts){
                echo "ERROR: After $attempts tries this Qlikview document has refused to open. Give up."
                echo "Full exception trace: "
                echo $_.Exception | format-list -force
                throw
            }
        }
    }
}

# The main function of this module. Reloads conversion QVW, then prints sheet. 
# This function rather specific to a CASM project due to button names, 
# and may need reconfiguring - primarily ExportPDF.
function fullReloadAndPrint{
    param (
        [string] $conversionQVWPath,
        [string] $convertedQVDPath,
        [string] $appQVDPath,
        [string[]] $pdfPaths,
        [int] $exportAttempts = 5
    )

    # We want this script to stop on non terminating errors, so set it up accordingly
    $oldErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "stop"


    # Reload the conversion QVD, and allow 5 mins for the output QVD to be updated
    $timeBeforeConversion = Get-Date
    echo "Converting CSV to QVD"
    reloadQVW -qvwPath $conversionQVWPath

    echo "Checking QVD file"
    checkFileUpdated -path $convertedQVDPath `
                     -qvd $conversionQVWPath `
                     -timeout 60 `
                     -since $timeBeforeConversion

    # Reload the app QVD, and press buttons within it to generate a pdf. 
    # Try this a few times.
    $timeBeforeAppReload = Get-Date
    echo "Reloading app QVD at $appQVDPath"
    reloadAndExport -QVDToReload $appQVDPath `
                    -attempts $exportAttempts
    
    foreach ($PDF in $pdfPaths){                    
        echo "Checking the PDF at $PDF has been updated"
        checkFileUpdated -path $PDF `
                         -qvd $appQVDPath `
                         -timeout 600 `
                         -since $timeBeforeAppReload
    }
    $ErrorActionPreference = $oldErrorActionPreference
}
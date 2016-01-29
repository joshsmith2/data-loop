<#

    Some utilities for manipulating credentials etc

#> 

# Given a username, and the path to a file containing a secure string, return a PSCredential object.
function getCredentials{
    param(
        [string] $username,
        [string] $passwordFile
    )
    $credentialObject = New-Object -TypeName System.Management.Automation.PSCredential `
 -ArgumentList $username, (Get-Content $passwordFile | ConvertTo-SecureString)

    return $credentialObject
}

# Convert a string to a secure, hashed string, and save it in a .txt file
function makeSecureString{
    param (
        [string] $stringToConvert,
        [string] $outputFilePath
    )   

    $stringToConvert | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File $outputFilePath
}
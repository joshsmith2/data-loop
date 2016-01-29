<#

    Uses Gmail as an SMTP server to send an email from a defined path.

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

$emailCredentials = getCredentials -username $emailUsername -passwordFile $emailPasswordFile

# Assemble a system.Net.Mail.MailMessage object, and send it off.
function sendMessage{
    
    
    param( 
        [System.Management.Automation.PSCredential] $emailCredentials, 
        [string] $from,
        [string] $to,
        [string[]] $bcc,
        [string[]] $attachments,
        [string] $subject,
        [string] $body
    )

    # Create message and populate fields
    $message = New-Object system.Net.Mail.MailMessage $from, $to 
    $message.Subject = $subject
    $message.Body    = $body

    # Populate attachments
    foreach ($attachment in $attachments){
        $message.Attachments.Add($attachment)
    }

    # Add addresses to BCC
    foreach ($address in $bcc){
        $addressObject = New-Object system.Net.Mail.mailaddress $address[0], $address[1]
        $message.Bcc.Add($addressObject)
    }

    # Create SMTP Client 
    $client = New-Object system.Net.Mail.SmtpClient 
    $client.Host = "smtp.gmail.com" 
    $client.Port = 587
    $client.EnableSsl = $true
    $client.Credentials = $emailCredentials

    # send the message 
    try { 
        $client.Send($message); 
    }   
    catch { 
        "Exception caught while sending message: {0}" -f $Error[0] 
    } 
}






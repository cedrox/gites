# iterate through a csv file and send an email to each user via gmail
# install module 
function Send-ToEmail([string]$email, [string]$login, [string]$password) {
    
    $message = new-object Net.Mail.MailMessage
    $message.From = "info@gitedemer.net"
    $message.To.Add($email)
    $message.Subject = "Gite de mer: Infos Marité et Jacky"
    $message.Body = Get-Content -Path "mail.html" -Raw

    $smtp = new-object Net.Mail.SmtpClient("smtp.gmail.com", "587")
    $smtp.EnableSSL = $true
    $smtp.UseDefaultCredentials = $false
    # $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
    $smtp.Credentials = New-Object System.Net.NetworkCredential($login, $password)
    try {
        $smtp.send($message)

        write-host "Mail Sent"  
    }
    catch {
        write-host "Failed to send email"
        Write-Host $_.Exception.Message -ForegroundColor Red
        write-host $Error[0].Exception
    }
    
}

$username = "info@gitedemer.net"
$password = Get-Content -Path "C:\Users\cefollio\OneDrive\Projets\GiteDeMer\Backup\mailPassword.txt" 

# Set the csv file path
$csv = "C:\Users\cefollio\OneDrive\Projets\GiteDeMer\Backup\AllMailsTest.csv"

# Set the csv variable
$csv = Import-Csv $csv

# Iterate through the csv file
foreach ($line in $csv) {
    # Set the to variable
    $to = $line.mail
    Send-ToEmail  -email $to -login $username -password $password
    # Send the email
    # Send-MailMessage -From $from -To $to -Subject $subject -Body $body -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $credentials
}



# # Set the username and password variables
# $username = "info@gitedemer.net"

# # Set the subject and body variables
# $subject = " Gite de mer: Infos Marité et Jacky"
# # read content from file mail.html
# $body = Get-Content -Path "mail.html" -Raw

# # Set the from and to variables
# $from = "info@gitedemer.net"

# # Set the SMTP server and port variables
# $smtpServer = "smtp.gmail.com"
# $smtpPort = "587"

# # Set the credentials variable
# $credentials = New-Object System.Management.Automation.PSCredential -ArgumentList @($username,(ConvertTo-SecureString -String $password -AsPlainText -Force))

# iterate through a csv file and send an email to each user via gmail
# install module 
# don't forget to allow less secure apps in gmail: https://myaccount.google.com/lesssecureapps
function Send-ToEmail([string]$email, [string]$login, [string]$password, [string]$body) {
    
    $message = new-object Net.Mail.MailMessage
    $message.From = "info@gitedemer.net"
    $message.To.Add($email)
    $message.Subject = "Après travaux, nos appartements sont ouverts"
    $message.Body = $body

    $smtp = new-object Net.Mail.SmtpClient("smtp.gmail.com", "587")
    $smtp.EnableSSL = $true
    $smtp.UseDefaultCredentials = $false
    # $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
    $smtp.Credentials = New-Object System.Net.NetworkCredential($login, $password)
    try {
       # $smtp.send($message)
        # check if the recipicent is valid and still exists
        if ($smtp.Send($message).Status -eq "Failure") {
            write-host "Failed to send email to $email"
            return
        }


        write-host "Mail Sent to $email" -ForegroundColor Green -BackgroundColor Black
    }
    catch {
        write-host "Failed to send email to $email"
        Write-Host $_.Exception.Message -ForegroundColor Red
        write-host $Error[0].Exception
    }
    
}

$username = "info@gitedemer.net"
$password = Get-Content -Path "C:\Users\cefollio\OneDrive\Projets\GiteDeMer\Backup\mailPassword.txt" 
#Get the ps1 path
$scriptPath = $MyInvocation.MyCommand.Path.Substring(0, $MyInvocation.MyCommand.Path.LastIndexOf("\")+1)
$content = Get-Content -Path ($scriptPath +"mail.html") -Raw
# Set the csv file path
$csv = "C:\Users\cefollio\OneDrive\Projets\GiteDeMer\Backup\AllMails.csv"

# Set the csv variable
$csv = Import-Csv $csv

$i=0
$fromLine=651
$toLine=750
# Iterate through the csv file
foreach ($line in $csv) {
    # Send mail user x to user x only
    if ($i -lt $fromLine) {
        $i++
        continue
    }
    
    if ($i -gt $toLine) {
        break
    }
    
    Write-Host "Send Mail N° $i" -ForegroundColor Green -BackgroundColor Black
    $to = $line.mail
    Send-ToEmail  -email $to -login $username -password $password -body $content
    Start-Sleep -s 5
    $i ++
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

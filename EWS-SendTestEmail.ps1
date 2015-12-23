#Test that Exchange can receive email via EWS

import-module -name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

$Username = ""
$Password = ""
$Domain = ""
$Email = ""

$smtp = "" 

$to = "" 
$bcc = "" 
$to = "" 
$from = "" 
$subject = ""  

$mailCount=0
$mailCount2=0

$ExchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_sp2)
$ExchService.Credentials = New-Object System.Net.NetworkCredential -ArgumentList $Username, $Password, $Domain

$ExchService.AutodiscoverUrl($email)

$mBody = ''
$mBody2 = ''

$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Exchservice,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)

$PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::html;
#$SearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Isequal([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[system.DateTime]::Now.AddDays(-5))

$View =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000) 

$Mails = $ExchService.FindItems($Inbox.Id,$View)

$subject = 'test subject'
$mBody = 'test body'
send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $mbody -BodyAsHtml #-Priority high 



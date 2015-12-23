#Randy Coffland
#Not the best way to do this, but I was in a hurry.  Better way is to search inbox for subjects instead of opening each email
#
import-module -name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

$Username = "ChangeUserName"
$Password = "ChangePassword"
$Domain = "ChangeDomain"
$Email = "ChangeEmailAddressofMailBox"

$smtp = "ChangeServerName" 
$to = "ChangeToAddress" 
$bcc = "ChangeAnyBCC" 
$from = "ChangeFromAddress" 
$subject = "Conflicts Check and Add Parties Daily Summary"  

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

if ($mails.totalcount -ne 0) 
{

    do {  
        $ExchService.LoadPropertiesForItems($Mails,$PropertySet) | Out-Null

        foreach ($Mail in $Mails.Items) 
            {
                if ($mail.Subject -like "*Conflict Search Requested*")
                {
                 $mailCount = $mailCount + 1
                $mBody = $mBody + "<P>********************* Conflict # "  + $mailcount + " ********************* <P>"

                $mBody = $mBody + " " + $Mail.Body        
                 $mail.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
                
               }

     
             if ($mail.Subject -like "*Complete for Add Party*")
                {
                 $mailCount2 = $mailCount2 + 1
                 $mBody2 = $mBody2 + "<P>********************* Add Parties # "  + $mailcount2 + " ********************* <P>"

                 $mBody2 = $mBody2 + " " + $Mail.Body        
                $mail.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
                
                 }






        $View.Offset += $Mails.Items.Count
     
            }

        } while($Mails.MoreAvailable -eq $true)

 }     

 


if ($mailCount -ne 0 -or $mailCount2 -ne 0)
{
$subject = $subject + ' - ' + $mailCount + ' Conflicts / ' + $mailCount2 + ' Add Parties'
$mBody = $mBody + $mBody2 +  "<P> **************** End of Conflicts / Add Parties Summary ****************"
send-MailMessage -SmtpServer $smtp -To $to -bcc $bcc -From $from -Subject $subject -Body $mbody -BodyAsHtml #-Priority high 
}

#Set-ExecutionPolicy Unrestricted


#Start-Process Outlook

#$document = Import-Csv -Path 'C:\Users\Beju\Desktop\Book1.csv'

$document = Import-Csv -Path 'C:\Users\cbejinar\Desktop\incident.csv'

 


$document | foreach {

    $outlook = New-Object -com Outlook.Application

 

    # This is where the new mail object is created

    $mail = $outlook.CreateItem(0)

 

    # This next line is to set the importantce of the email

    # 2 = High importance

    $mail.importance = 2

 

    # The email address you want to send from, you need to have SendAs access to it on your account

    $mail.SentOnBehalfOfName = helpdesk@crfhealth.com

 

    # For multiple email, use semi-colon ; to separate

    $mail.To = $($_.email)

 

# Search email address in Outlook Global Address List

#Function Search-GAL {

#      param (

#             [string]$searchString

#      )

#      $ol = New-Object -ComObject Outlook.Application

#      $item = $ol.Session.GetGlobalAddressList().AddressEntries.Item($searchString)

#      Write-Host ($item.Name), ($item.Address)

#}

    # Email subject line
    $mail.subject = "First strike | $($_.number) - $($_.business_service) - $($_.cmdb_ci) - $($_.short_description)"


    # This is the body of the email, you can write whaever you want in here
    $mail.body = "Hi team,

                    We are in need of an update for a user reported incident that has the following details:

                    Ticket number: $($_.number)
                    Ticket Open Date: $($_.sys_created_on) PST 
                    Product: $($_.business_service)
                    study: $($_.cmdb_ci)
                    Study Protocol: 20170703
                    Site: 21003
                    Short Description: $($_.short_description)"

    # This is the fuction that executes itself in order to send the email

    $mail.Send()

    Write-Host "Email was sent to $($_.email)"

}

#Start-Process Outlook
$document = Import-Csv -Path 'C:\Users\Beju\Desktop\Book1.csv'


$document | foreach {
    $outlook = New-Object -com Outlook.Application

    # This is where the new mail object is created
    $mail = $outlook.CreateItem(0)

    # This next line is to set the importantce of the email
    # 2 = High importance
    $mail.importance = 2

    # The email address you want to send from, you need to have SendAs access to it on your account
    $mail.From = "something@gmail.com"

    # For multiple email, use semi-colon ; to separate
    $mail.To = $($_.email)

    # Email subject line
    $mail.subject = "First strike | (product name) - (subject code) - (actual description)"

    # This is the body of the email, you can write whaever you want in here
    $mail.body = "This shit rocks!"



    # This is the fuction that executes itself in order to send the email
    $mail.Send()

    Write-Host "Email was sent to $($_.email)"

}
Sub send_mail()

Dim objOutlook As Object
Dim objEmail As Object

Set objOutlook = CreateObject("Outlook.Application")
Set objEmail = objOutlook.CreateItem(0)

' Build the email structure:
objEmail.Subject = "Hello there!"
objEmail.Body = "This is a test message!"
objEmail.To = "firstname.surname@provider.com"
 
' Display and send methods:
'objEmail.Display
objEmail.Send

End Sub

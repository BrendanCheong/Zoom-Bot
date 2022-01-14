Dim OlApp
Dim Eml
Dim Arg
Dim EmlBody 
Set Arg = WScript.Arguments
Const olFormatHTML = 2

Dim ClientName
ClientName = Arg(0)

Dim ZoomLink
ZoomLink = Arg(1)

Dim EmailRecipient
EmailRecipient = Arg(2)

Dim Attachments
Attachments = Arg(3)

Dim Input
Input = Arg(4)

Set OlApp = CreateObject("Outlook.Application")
Set Eml = OlApp.CreateItemFromTemplate(Input)

ProcessMailBody()
ProcessMailSubject()
ProcessOthers()

Sub ProcessMailSubject()
    Eml.Subject = Replace(Eml.Subject, "Recipient_of_email", ClientName)
    Eml.Subject = Replace(Eml.Subject, "}", "")
    Eml.Subject = Replace(Eml.Subject, "{", "")
End Sub

Sub ProcessMailBody()
    Eml.HtmlBody = Replace(Eml.HtmlBody, "Zoom_link", "<a href=" & ZoomLink & ">Unique Zoom Link</a>")
    Eml.HtmlBody = Replace(Eml.HtmlBody, "Recipient_of_email", ClientName)
    Eml.HtmlBody = Replace(Eml.HtmlBody, "}", "")
    Eml.HtmlBody = Replace(Eml.HtmlBody, "{", "")
End Sub

Sub ProcessOthers()
    Eml.Recipients.Add(EmailRecipient)
    Eml.Attachments.Add(Attachments)
    Eml.BodyFormat = olFormatHTML
End Sub

Eml.Send
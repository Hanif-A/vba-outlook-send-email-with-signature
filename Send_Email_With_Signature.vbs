Sub Send_Email_With_Signature()

    ' Tested on Office 2007 - 2019

    Dim strHTMLBody As String, strSignature As String
    
    Dim objOutApp As Object: Set objOutApp = CreateObject("Outlook.Application")
    Dim objOutMail As Object: Set objOutMail = objOutApp.CreateItem(0)

    On Error Resume Next

    With objOutMail

        ' Set the email parameters
        .To = "email@address.com"
        .CC = ""
        .BCC = ""
        .Subject = "Your subject here"

        ' If you want to include an attachment, uncomment the next line
        '.Attachments.Add ("C:\Users\Desktop\File.docx)

        ' To send from another email account (e.g. team inbox), change here and uncomment
        '.SentOnBehalfOfName = "my_team_account@emailaddress.com"

        ' Ensures all recipients are validated against the address book or are valid email formats
        .Recipients.ResolveAll

        ' First popup the message box - this contains your default signature
        .Display

        ' Store the signature in a variable
        strSignature = .Htmlbody

        ' You can create your own HTML body here
        'HTML TAGS CAN BE INCLUDED HERE
        strHTMLBody = "<font face=Tahoma size=3> Add whatever you want your message to say here.</a></font>"

        ' Now replace the contents with your contents and your signature
        .Htmlbody = strHTMLBody & strSignature

        ' If you want the email to automatically send, turn this on (email will briefly display, this is unavoidable)
        '.Send

    End With

    On Error GoTo 0
    Set objOutMail = Nothing
    Set objOutApp = Nothing

End Sub

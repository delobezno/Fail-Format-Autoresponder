Attribute VB_Name = "Modul1"
Sub CheckAttachmentsAndReply(Item As Outlook.MailItem)
    On Error GoTo ErrorHandler ' Enable error handling

    Dim att As Attachment
    Dim response As MailItem
    Dim attachType As String
    Dim senderDomain As String
    Dim internalDomain As String

    attachType = ""
    internalDomain = "@pbeakk.de" ' Replace with your organization's domain

    ' Get the sender's domain
    senderDomain = LCase(Mid(Item.Sender.Address, InStrRev(Item.Sender.Address, "@") + 1))

    ' Check if the email is from an external domain
    If InStr(senderDomain, internalDomain) > 0 Then
        Debug.Print "Internal email, no reply needed."
        Exit Sub ' Skip the auto-reply for internal emails
    End If

    ' Check if the email has attachments
    If Item.Attachments.Count = 0 Then
        Debug.Print "No attachments found."
        Exit Sub
    End If

    ' Loop through each attachment and check the file extension
    For Each att In Item.Attachments
        Debug.Print "Attachment: " & att.FileName ' Debugging output
        If LCase(Right(att.FileName, 3)) = "png" Or _
           LCase(Right(att.FileName, 3)) = "xls" Or _
           LCase(Right(att.FileName, 4)) = "xlsx" Or _
           LCase(Right(att.FileName, 3)) = "odt" Then
            Debug.Print "Invalid attachment found: " & att.FileName
            attachType = "invalid"
            Exit For
        Else
            Debug.Print "Valid attachment: " & att.FileName ' Output for valid attachments
        End If
    Next att

    ' If an invalid attachment is found, send an automatic reply
    If attachType = "invalid" Then
        Debug.Print "Sending automatic reply."

        Set response = Application.CreateItem(0) ' Create a new email (olMailItem = 0)
        response.Subject = "Invalid Attachment Format"
        response.Body = "Thank you for your email. Unfortunately, we cannot process your message because it contains attachments with unsupported file formats (PNG, XLS, XLSX, or ODT). Please resend the email with an acceptable file format (e.g., PDF, DOCX)." & vbCrLf & vbCrLf & _
                        "Thank you for your understanding." & vbCrLf & _
                        "Best regards," & vbCrLf & "[Your Name or Company Name]"

        response.To = Item.Sender.Address ' Send the reply to the sender
        response.Send

        Debug.Print "Automatic reply sent to: " & Item.Sender.Address
    Else
        Debug.Print "No invalid attachments found, no reply sent."
    End If

    Exit Sub ' Normal exit

ErrorHandler: ' Error handling
    Debug.Print "Error: " & Err.Description
End Sub



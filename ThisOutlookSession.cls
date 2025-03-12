VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim objMail As Object
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    ' Get the mail item by EntryID
    Set objMail = ns.GetItemFromID(EntryIDCollection)

    ' Check if the object is a mail item
    If TypeName(objMail) = "MailItem" Then
        ' Call the function to check attachments and send the reply
        Call CheckAttachmentsAndReply(objMail)
    End If
End Sub


<div align="center">

## Easy Outlook E\-Mail w/Attachment


</div>

### Description

To send an e-mail with outlook with an attachment. Very easy to understand instructions.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dino Roger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dino-roger.md)
**Level**          |Beginner
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dino-roger-easy-outlook-e-mail-w-attachment__1-59230/archive/master.zip)





### Source Code

```
'This is a simple way to use the Outlook reference to send an e-mail with an attachment
'Set the Boolean to true at the end of the SendMail function call to display the e-mail
'instead of automatically sending it.
'INSTRUCTIONS:
'* Click PROJECT - REFERENCES
'* Check the box next to "Microsoft Outlook ## Object Library
'* Copy the below code into a form
'
Function SendMail(EM_TO, Em_CC, EM_BCC, EM_Subject, EM_Body, EM_Attachment As String, Display As Boolean)
 Dim objOA As Outlook.Application
 Dim objMI As Outlook.MailItem
 Dim obgAtt As Outlook.Attachments
 Set objOA = New Outlook.Application
 Set objMI = objOA.CreateItem(olMailItem)
 If EM_TO <> "" Then objMI.To = EM_TO
 If Em_CC <> "" Then objMI.CC = Em_CC
 If EM_BCC <> "" Then objMI.BCC = EM_BCC
 If EM_Subject <> "" Then objMI.Subject = EM_Subject
 If EM_Body <> "" Then objMI.Body = EM_Body
 If EM_Attachment <> "" Then objMI.Attachments.Add EM_Attachment, 1, , EM_Attachment
 If Display Then
  objMI.Display
   Else
    objMI.Send
 End If
 Set objOA = Nothing
 Set objMI = Nothing
End Function
Private Sub Form_Load()
 'How to call the SendMail function. If you do not want a function of the main just use two quotes and a comma
 'instead of filling the string variable. Example of a call with a To only:
 'SendMail "SendTo@Address.com", "", "", "", "", "", True
 'The code represented here will error unless a real attachment path is specified.
 SendMail "SendTo@Address.com", "CarbonCopy@Address.com", "BlindCC@Address.com", "SUBJECT", "BODY", "C:\Attachment.txt", False
End Sub
```


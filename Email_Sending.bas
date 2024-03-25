Attribute VB_Name = "Email_Sending"
Public Sub SendPentegraMessage(ByRef oTo() As String, ByVal oSubject As String, ByVal oBody As String, oImportance As Integer)

  Dim oTask As Outlook.TaskItem
  Dim oMail As Outlook.MailItem
  Dim oFld As Outlook.MAPIFolder

      Set oMail = Application.CreateItem(olMailItem)
      oMail.subject = oSubject
      oMail.Body = oBody
      
      For x = LBound(oTo) To UBound(oTo) 'define start and end of array
    
          oMail.Recipients.Add oTo(x)
      
        Next x ' Loop!

      
      oMail.Importance = oImportance
      oMail.Recipients.ResolveAll
      oMail.Send

End Sub

Sub ReplyAllDone()

Dim replyItem As MailItem

Set objItem = GetCurrentItem()

Set replyItem = objItem.ReplyAll
    
replyItem.Display

Set objDoc = Application.ActiveInspector.WordEditor
Set objSel = objDoc.Windows(1).Selection
objSel.TypeText Text:="This has been taken care of." + Chr(13) + Chr(13) + "Thanks," + Chr(13) + "Rob"

MsgBox ("Hello")

' replyItem.Send  ' This will send the email, and I had too many accidental sends

'objItem.ReplyAll(Response, false)
 
End Sub
Sub ReplyTaskList()

Dim replyItem As MailItem

Set objItem = GetCurrentItem()

Set replyItem = objItem.ReplyAll
    
replyItem.Display

Set objDoc = Application.ActiveInspector.WordEditor
Set objSel = objDoc.Windows(1).Selection
objSel.TypeText Text:="I just wanted to let you know I've seen this and it's been added to my queue." + Chr(13) + Chr(13) + "Thanks," + Chr(13) + "Rob"

' replyItem.Send  ' This will send the email, and I had too many accidental sends

'objItem.ReplyAll(Response, false)

End Sub
Sub ReplyResearching()

Dim replyItem As MailItem

Set objItem = GetCurrentItem()

Set replyItem = objItem.ReplyAll
    
replyItem.Display

Set objDoc = Application.ActiveInspector.WordEditor
Set objSel = objDoc.Windows(1).Selection
objSel.TypeText Text:="I'll need to look into this.  I'll let you know when I have more information." + Chr(13) + Chr(13) + "Thanks," + Chr(13) + "Rob"

' replyItem.Send  ' This will send the email, and I had too many accidental sends

'objItem.ReplyAll(Response, false)
 
End Sub

Sub SendEmail()

 Dim OutApp As Outlook.Application
       Dim OutMail As Outlook.MailItem
        Dim colFiles As New Collection

        Set OutApp = New Outlook.Application
        Set OutMail = OutApp.CreateItem(0)
        'OutMail.Parent.Display
        On Error Resume Next
        With OutMail
            .To = "rscholl@pentegra.com"
            .subject = "Reminder"
            .Body = "Dear Sir"
            .Send
         End With

End Sub

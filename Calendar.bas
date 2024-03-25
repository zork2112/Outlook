Attribute VB_Name = "Calendar"
Sub CreateAppt()
 
    Dim myItem As Object
 
    Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient
 
    Set myItem = Application.CreateItem(olAppointmentItem)
    myItem.MeetingStatus = olMeeting
    
    myItem.subject = "Rob Scholl Out of Office"
    
    Set myOptionalAttendee = myItem.Recipients.Add("InformationTechnology@pentegra.com")
    
    myItem.BusyStatus = 0
    myItem.Location = "n/a"
    myItem.Start = Now
    myItem.AllDayEvent = True
    myItem.ReminderSet = False
    
    myItem.ResponseRequested = False
    myItem.Display

End Sub


Sub CreateApptTryToGetInfoTechToResolve()
 Dim myItem As Object
 Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient
 
 Set myItem = Application.CreateItem(olAppointmentItem)
 myItem.MeetingStatus = olMeeting
 myItem.subject = "Rob Scholl Out of Office"
 myItem.BusyStatus = 0
 myItem.Location = "n/a"
 myItem.Start = Now
 myItem.AllDayEvent = True
 myItem.ReminderSet = False
 
 ' 20230124
 'Dim myRecipients As Outlook.Recipients
 'Dim myRecipient As Outlook.Recipient
 'Set myRecipients = myItem.Recipients
 'myRecipients.Add ("Information Technology")
 'myRecipient.Name = "Information Technology"
   
 'If Not myRecipients.ResolveAll Then
 
  '  For Each myRecipient In myRecipients
 
   '     If Not myRecipient.Resolved Then
 
    '        MsgBox myRecipient.Name
 
     '   End If
    'Next
 
 'End If
 
 'Set requiredRecipient = myItem.Recipients.Add("Information Technology")
 'requiredRecipient.Type = olRequired
 'olOptional
 'requiredRecipient.Resolve
 
 'Set requiredRecipient = myItem.Recipients.To("Information Technology")
   
  'If requiredRecipient.Resolved Then
 
   '     MsgBox myRecipient.Name
  'End If
  
 'Set myRequiredAttendee = myItem.Recipients.Add("Information Technology")
 'Set myRequiredAttendee = myItem.Recipients.Add("InformationTechnology@pentegra.com")
 Set myRequiredAttendee = myItem.Recipients.Add("zork2112@gmail.com")
 myItem.ResponseRequested = False
 myItem.Display

End Sub


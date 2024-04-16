Option Explicit

' ____________________________________________________________________________
Sub accept_outlook_appointment() 

Dim olNameSpace As Outlook.NameSpace 
Dim olInboxDirectory As Outlook.Folder 
Dim olMeetingRequest As Outlook.MeetingItem 
Dim olAppointmentItem As Outlook.AppointmentItem 
Dim olMeetingResponse As Outlook.MeetingItem 

Set olNameSpace = Application.GetNamespace("MAPI") 
Set olInboxDirectory = myNameSpace.GetDefaultFolder(olFolderInbox)
Set olMeetingRequest = myFolder.Items.Find("[MessageClass] = 'IPM.Schedule.Meeting.Request'")

If TypeName(olMeetingRequest) <> "Nothing" Then 
    Set olAppointmentItem = myMtgReq.GetAssociatedAppointment(True) 
    Set olMeetingResponse = myAppt.Respond(olResponseAccepted, True) 
    olMeetingResponse.Send 
End If 

End Sub

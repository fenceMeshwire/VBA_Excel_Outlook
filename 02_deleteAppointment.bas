Option Explicit

' Please refer to the program 01_makeAppointments.bas
' The name tag is created in the eMail subject
' Activate the reference library "Microsoft Outlook XY.0 Object Library", depending on your installed office version.

Sub deleteAppointments()

Dim appOutlook As Outlook.Application
Dim nspOutlookNameSpace As Outlook.Namespace
Dim appointmentDirectory As Outlook.MAPIFolder
Dim appointments As Outlook.Items
Dim lngIndex, lngCount As Long

On Error GoTo deleteAppointments_Error

Set appOutlook = New Outlook.Application
Set nspOutlookNameSpace = appOutlook.GetNamespace("MAPI")
Set appointmentDirectory = nspOutlookNameSpace.GetDefaultFolder(olFolderCalendar)
Set appointments = appointmentDirectory.Items

lngCount = appointments.Count

For lngIndex = lngCount To 1 Step -1
  If Right(appointments(lngIndex).Subject, 5) = "[L&H]" Then
    appointments(lngIndex).Delete
  End If
Next lngIndex

Exit Sub

deleteAppointments_Error:
MsgBox "An error occured. Program finished."

End Sub

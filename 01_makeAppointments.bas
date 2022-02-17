Option Explicit

Sub makeSprintAppointments()

Dim lngRow, lngRowMax As Long
Dim dteStart, dteEnd As Date
Dim strSprint As String

Dim wksSheet As Worksheet

'dteStart: Start date for sprint -> Sheet1.Cells(lngRow, 1).Value
'dteEnd: End date for sprint -> Sheet1.Cells(lngRow, 2).Value
'strSprint: Designation for sprint -> Sheet1.Cells(lngRow, 3).Value

Set wksSheet = Sheet1

With wksSheet
  lngRowMax = .UsedRange.Rows.Count
  For lngRow = 1 To lngRowMax
    dteStart = .Cells(lngRow, 1).Value
    dteEnd = .Cells(lngRow, 2).Value
    strSprint = .Cells(lngRow, 3).Value
    Call createAppointment(dteStart, dteEnd, strSprint)
  Next lngRow
End With

End Sub

'Create an appointment
Sub createAppointment(dteStart, dteEnd As Date, strSprint As String)
       
Dim appOutlook, objAppointment As Object

Set appOutlook = CreateObject("Outlook.Application")
Set objAppointment = appOutlook.CreateItem(1)

With objTermin
  .MeetingStatus = 1
  .AllDayEvent = True
  .Start = dteStart
  .End = dteEnd + 1
  .Subject = strSprint
  .Location = "workplace"
  .Save
End With

End Sub

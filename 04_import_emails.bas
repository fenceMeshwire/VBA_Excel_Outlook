Option Explicit

' ________________________________________________________________________________________________
Sub import_messages_from_outlook()

' Import Outlook library first.
Dim objOutFolder As Outlook.MAPIFolder

Dim dteMax As Date
Dim i As Integer
Dim lngMsg As Long, lngMsgGes As Long, lngRowFree As Long
Dim wksSheet as Worksheet

Set wksSheet = Sheet1

On Error GoTo Error_Handler

Set objOutFolder = GetObject("", "Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("Subfolder_To_Inbox")

lngMsgGes = objOutFolder.Items.Count

wksSheet.Cells(1, 1).Value = "Subject"
wksSheet.Cells(1, 2).Value = "Sender"
wksSheet.Cells(1, 3).Value = "Date"
wksSheet.Cells(1, 4).Value = "Body"
wksSheet.Cells(1, 5).Value = "Read"

lngRowFree = wksSheet.Cells(wksSheet.Rows.Count, 1).End(xlUp).Row + 1

For lngMsg = 1 To lngMsgGes
  With objOutFolder.Items(lngMsg)
      
      wksSheet.Cells(lngRowFree, 1).Value = .Subject
      wksSheet.Cells(lngRowFree, 2).Value = .SenderName
      wksSheet.Cells(lngRowFree, 3).Value = Format(.ReceivedTime, "dd.mm.yyyy hh:mm")
      wksSheet.Cells(lngRowFree, 4).Value = .Body
      ' Determine if the email has been read:
      If Not .UnRead = -1 Then
        wksSheet.Cells(lngRowFree, 5).Value = "Yes"
      Else
        wksSheet.Cells(lngRowFree, 5).Value = "No"
      End If
      wksSheet.Rows(lngRowFree).RowHeight = 15 ' Set row height due to the .Body import.

      lngRowFree = wksSheet.Cells(wksSheet.Rows.Count, 1).End(xlUp).Row + 1
      
  End With
Next lngMsg

For i = 1 To 5 ' Set the column width
  wksSheet.Columns(i).ColumnWidth = 25
Next i

If lngMsgGes > 0 Then
  MsgBox "Done! Successfully imported " & lngMsgGes & " new emails!"
  Exit Sub
Else
  MsgBox "No emails in the current directory!"
  Exit Sub
End If

Error_Handler:
MsgBox Err.Number & " " & Err.Description

End Sub

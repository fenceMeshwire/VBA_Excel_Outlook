Option Explicit

' ________________________________________________________________________________________________
Sub check_create_directory()

Dim myNameSpace As Object
Dim myFolderInbox As Outlook.Folder
Dim myFolder As Outlook.Folder

Dim strFolder As String

strFolder = "special_dir"

Set myNameSpace = GetObject("", "Outlook.Application").Application.GetNamespace("MAPI")
Set myFolderInbox = myNameSpace.GetDefaultFolder(olFolderInbox)

On Error GoTo ErrorHandler
Set myFolder = myFolderInbox.Folders.Add(strFolder, olFolderInbox)

Exit Sub

ErrorHandler:
  MsgBox "The Directory " & strFolder & " exists already!"
  Resume Next

End Sub

Option Explicit

' ________________________________________________________________________________________________
Sub check_create_rules()

Dim colRule As Object, colRules As Object
Dim oRule As Outlook.Rule
Dim colRuleActions As Outlook.RuleActions
Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction
Dim oFromCondition As Outlook.ToOrFromRuleCondition
Dim oExceptSubject As Outlook.TextRuleCondition
Dim oSPECIAL As Object
Dim oMoveTarget As Outlook.Folder

Dim strRuleName As String

strRuleName = "special_rule"

' Specify target directory:
On Error GoTo err_handler
Set oSPECIAL = GetObject("", "Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("special_dir")

' Obtain rules from Session.DefaultStore object
Set colRules = GetObject("", "Outlook.Application").Session.DefaultStore.GetRules()

' Check if the rule has been created already:
For Each colRule In colRules
  If colRule.Name = strRuleName Then
    MsgBox "The rule for moving the incoming emails has been created already!"
    Exit Sub
  End If
Next colRule

' Create the rule
Set oRule = colRules.Create("special_rule", olRuleReceive)

' Specify the conditions
Set oFromCondition = oRule.Conditions.From

With oFromCondition
    .Enabled = True
    .Recipients.Add ("sam.sample@sampleton.com")
    .Recipients.ResolveAll
End With

' Specify the action in a MoveOrCopyRuleAction object
Set oMoveRuleAction = oRule.Actions.MoveToFolder

With oMoveRuleAction
    .Enabled = True
    .Folder = oSPECIAL
End With

' Update the server and display progress dialog
colRules.Save

MsgBox "The rule has been created successfully."

Exit Sub

err_handler:
MsgBox "The directory special_dir does not exist. Please create the directory in order to proceed."

End Sub

Attribute VB_Name = "Módulo2"
Sub RunRules(Item As Outlook.MailItem)

Dim olRules As Outlook.Rules
Dim myRule As Outlook.Rule
Dim olRuleNames() As Variant
Dim name As Variant

' Enter the names of the rules you want to run
olRuleNames = Array("SetCategory", "SetFlag", "MoveClutter")

Set olRules = Application.Session.DefaultStore.GetRules()

For Each name In olRuleNames()
    For Each myRule In olRules
    ' Rules we want to run
        If myRule.name = name Then
        myRule.Execute ShowProgress:=True
        End If
    Next
Next
End Sub


'Sub RunRulesSecondary()
'
'Dim oStores As Outlook.Stores
'Dim oStore As Outlook.Store
'
'Dim olRules As Outlook.Rules
'Dim myRule As Outlook.Rule
'Dim olRuleNames() As Variant
'Dim name As Variant
'
'' Enter the names of the rules you want to run
'olRuleNames = Array("Rule 1 Name", "Rule 2 Name", "Rule 3 Name")
'
'Set oStores = Application.Session.Stores
'For Each oStore In oStores
'On Error Resume Next
'
'' use the display name as it appears in the navigation pane
'If oStore.DisplayName = "alias@domain.com" Then
'
'Set olRules = oStore.GetRules()
'
'For Each name In olRuleNames()
'
'    For Each myRule In olRules
'       Debug.Print "myrule " & myRule
'
'     If myRule.name = name Then
'
'' inbox belonging to oStore
'' need GetfolderPath functionhttp://slipstick.me/4eb2l
'        myRule.Execute ShowProgress:=True, Folder:=GetFolderPath(oStore.DisplayName & "\Inbox")
'
'' current folder
''      myRule.Execute ShowProgress:=True, Folder:=Application.ActiveExplorer.CurrentFolder
'
'       End If
'    Next
'Next
'
'End If
'Next
'End Sub


Sub RunAllInboxRules()
Dim st As Outlook.Store
Dim myRules As Outlook.Rules
Dim rl As Outlook.Rule
Dim runrule As String
Dim rulename As String
rulename = "Your Rule Name"
Set st = Application.Session.DefaultStore
Set myRules = st.GetRules
Set cf = Application.ActiveExplorer.CurrentFolder
For Each rl In myRules
If rl.RuleType = olRuleReceive Then
If rl.name = rulename Then
rl.Execute ShowProgress:=True, Folder:=cf
runrule = rl.name
End If
End If
Next
ruleList = "Rule was executed correctly:" & vbCrLf & runrule
MsgBox ruleList, vbInformation, "Macro: Whatever_Finished"
Set rl = Nothing
Set st = Nothing
Set myRules = Nothing
End Sub


Sub SCAN_OCR()

End Sub

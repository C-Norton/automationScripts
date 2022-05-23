VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub ruleButton()
    
    FilePath = Environ("APPDATA") & "\Cnorton_AHK\OutlookCreateRule.txt" '
    TextFile = FreeFile()
    Open FilePath For Input As #TextFile
    Line Input #TextFile, casenum
    Line Input #TextFile, rulename
    Line Input #TextFile, FolderName
 
    Close #TextFile
    Close TextFile
    Close
 
    
    
    Call CreateRule(casenum, rulename, FolderName)

    
End Sub


Public Sub CreateRule(casenumber, rulename, folder)
    Dim colRules As Outlook.rules
    Dim oRule As Outlook.rule
    Dim colRuleActions As Outlook.RuleActions
    Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction
    Dim oFromCondition As Outlook.ToOrFromRuleCondition
    Dim oCondSubject As Outlook.TextRuleCondition
    Dim oInbox As Outlook.folder
    Dim oMoveTarget As Outlook.folder
 
    'Specify target folder for rule move action
    Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox)
    'Assume that target folder already exists
    Set oMoveTarget = oInbox.Folders(folder)
     
    'Get Rules from Session.DefaultStore object
    Set colRules = Application.Session.DefaultStore.GetRules()
     
    'Create the rule by adding a Receive Rule to Rules collection
    Set oRule = colRules.Create(rulename, olRuleReceive)
    'Specify the action in a MoveOrCopyRuleAction object
    'Action is to move the message to the target folder
    Set oMoveRuleAction = oRule.Actions.MoveToFolder
    With oMoveRuleAction
        .Enabled = True
        .folder = oMoveTarget
    End With
 
    'Specify the exception condition for the subject in a TextRuleCondition object
    'Exception condition is if the subject contains "fun" or "chat"
    Set oCondSubject = oRule.Conditions.Subject
    With oCondSubject
        .Enabled = True
        .Text = Array(casenumber)
    End With
    
    oRule.Enabled = True
    
    'Update the server and display progress dialog
    colRules.Save
End Sub

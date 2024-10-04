Attribute VB_Name = "Rules"



Sub RunAllInboxRules()
    Dim st As Outlook.Store
    Dim myRules As Outlook.Rules
    Dim rl As Outlook.Rule
    Dim count As Integer
    Dim ruleList As String
    'On Error Resume Next
     
    'Start
    Beep
        
    ' get default store (where rules live)
    Set st = Application.Session.DefaultStore
    ' get rules
    Set myRules = st.GetRules
     
    ' iterate all the rules
    For Each rl In myRules
        ' determine if it's an Inbox rule
        If rl.RuleType = olRuleReceive Then
            ' if so, run it
            rl.Execute ShowProgress:=False
            count = count + 1
            ruleList = ruleList & vbCrLf & rl.Name
        End If
    Next
     
    ' tell the user what you did
    ruleList = "These rules were executed against the Inbox: " & vbCrLf & ruleList
    'MsgBox ruleList, vbInformation, "Macro: RunAllInboxRules"
     
    
    Set rl = Nothing
    Set st = Nothing
    Set myRules = Nothing

    
    'Stop
    Beep
    
    'Application.wait does not work
    'This loop lets it have 2 beeps to single end
    
    PauseTime = 0.5   ' Set duration.
    Start = Timer    ' Set start time.
    Do While Timer < Start + PauseTime
        DoEvents    ' Yield to other processes.
    Loop
    
    Beep
    
End Sub





Sub UpdateOutlookRule()
    Dim colRules As Outlook.Rules
    Dim oRule As Outlook.Rule
    Dim oAddressRuleCondition As Outlook.AddressRuleCondition
    Dim ruleName As String
    Dim specificWords As String

    ' Define the rule name and specific words to look for
    ruleName = "Sender Address Rule"
    specificWords = "example" ' Change this to the specific words you want to filter

    ' Get the rules collection
    Set colRules = Application.Session.DefaultStore.GetRules()

    ' Check if the rule already exists
    On Error Resume Next
    Set oRule = colRules(ruleName)
    On Error GoTo 0

    ' If the rule doesn't exist, create it
    If oRule Is Nothing Then
        Set oRule = colRules.Create(ruleName, olRuleReceive)
    End If

    ' Clear existing conditions
    oRule.Conditions.Clear

    ' Create a new AddressRuleCondition for the sender's address
    Set oAddressRuleCondition = oRule.Conditions.SenderAddress
    oAddressRuleCondition.Enabled = True
    oAddressRuleCondition.Address = specificWords

    ' Save the updated rules
    colRules.Save
End Sub


Attribute VB_Name = "ExtractEmailAddresses"
Sub AddEmailToDeleteRule()

Dim OlApp As Outlook.Application
Dim Mailobject As Object
Dim Email As String
Dim NS As NameSpace
Dim Folder As MAPIFolder
Dim specificWords As String
 
Set OlApp = CreateObject("Outlook.Application")
' Setup Namespace
Set NS = ThisOutlookSession.Session

' Display select folder dialog (need to change this to be fixed)
Set Folder = NS.PickFolder

' get default store (where rules live)
Set st = Application.Session.DefaultStore
' get rules
Set myRules = st.GetRules
'Set myRules = Application.Session.DefaultStore.GetRules() ' Could do in one step

Set thisRule = myRules.Item("Delete Spam and Junk")
'Set thisRule = myRules.Item("TestMe")

'Debug.Print (thisRule.Name)

Dim currentRuleCondition As Outlook.AddressRuleCondition

Set currentRuleCondition = thisRule.Conditions.SenderAddress

' Cannot add to existing condition - have to replace, so create a collection, add to it, convert to array

' Only add unique values

Dim col As New Collection
For Each a In currentRuleCondition.Address
   If Not (CollectionValueExists(col, a)) Then col.Add a
Next


' Iterate through the emails in the folder you want to use for exclusions (usually Unknown contacts for me)
For Each Mailobject In Folder.Items
   
   Email = Mailobject.SenderEmailAddress 'Properties: Mailobject.To, Mailobject.Sender, Mailobject.SenderEmailAddress, Mailobject.SenderName and Mailobject.Body, Mailobject.HTMLBody or Mailobject.RTFBody
   If Not (CollectionValueExists(col, Email)) Then col.Add Email
   
Next


Dim newList() As Variant
newList = toArray(col) 'convert collection to an array

'printArray newList

With currentRuleCondition
    .Enabled = True
    .Address = newList
End With

' Save the updated rules
myRules.Save

Set OlApp = Nothing
Set Mailobject = Nothing

End Sub

Sub TestRule()
Dim olRules As Outlook.Rules
Dim olRule As Outlook.Rule

Set olRules = Application.Session.DefaultStore.GetRules
Set olRule = olRules.Item("TestMe")

Debug.Print TypeName(olRule)

printArray olRule.Conditions.Body.Text
printArray olRule.Conditions.MessageHeader.Text
printArray olRule.Conditions.SenderAddress.Address


End Sub

Private Sub printArray(ByRef pArr As Variant)
    Dim readString As Variant

    If (IsArray(pArr)) Then             'check if the passed variable is an array

        For Each readString In pArr

            If TypeName(readString) = "String" Then 'check if the readString is a String variable
                Debug.Print readString
            End If

        Next

    End If

End Sub

Function toArray(col As Collection)
  Dim arr() As Variant
  ReDim arr(0 To col.count - 1) As Variant
  For i = 1 To col.count
      arr(i - 1) = col(i)
  Next
  toArray = arr
End Function

' Check to see if a value is in a collection.
' Functional approcah to mimic dicitonary `exists` method.
Public Function CollectionValueExists(ByRef target As Collection, value As Variant) As Boolean
    Dim index As Long
    For index = 1 To target.count
        If target(index) = value Then
            CollectionValueExists = True
            Exit For
        End If
    Next index
End Function



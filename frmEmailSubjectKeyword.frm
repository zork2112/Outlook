VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmailSubjectKeyword 
   Caption         =   "Email Lead"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   OleObjectBlob   =   "frmEmailSubjectKeyword.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEmailSubjectKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub SetSubject(subject As String)

    Dim oMail As Outlook.MailItem

    Set oMail = GetCurrentItem()
    oMail.subject = subject + ": " + txtSubjectLine.Value
    
    Unload Me

End Sub

Private Sub btnAction_Click()

    SetSubject ("ACTION")

End Sub

Private Sub btnClose_Click()

    Unload Me
    
'ACTION – Compulsory for the recipient to take some action
'SIGN – Requires the signature of the recipient
'INFO – For informational purposes only, and there is no response or action required
'DECISION – Requires a decision by the recipient
'REQUEST – Seeks permission or approval by the recipient
'COORD – Coordination by or with the recipient is needed

End Sub

Private Sub btnCoord_Click()

    SetSubject ("COORD")

End Sub

Private Sub btnDecision_Click()

    SetSubject ("DECISION")
    
End Sub

Private Sub btnInfo_Click()

    SetSubject ("INFO")

End Sub

Private Sub btnQuestion_Click()

    SetSubject ("QUESTION")

End Sub

Private Sub btnRequest_Click()

    SetSubject ("REQUEST")
    
End Sub

Private Sub btnSign_Click()

    SetSubject ("SIGN")
    
End Sub

Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
           
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.currentItem
    End Select
       
    Set objApp = Nothing
End Function

Private Sub txtSubjectLine_Change()

End Sub

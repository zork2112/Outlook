VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRVEmployeeList 
   Caption         =   "UserForm1"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   OleObjectBlob   =   "frmRVEmployeeList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRVEmployeeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub chkAllOrNone_Click()

If chkAllOrNone.Value = True Then

    chkShuler.Value = True
    chkJon.Value = True
    chkMamatha.Value = True
    chkSrini.Value = True
    chkShoven.Value = True
    chkStokes.Value = True
    chkOwen.Value = True
    chkJenn.Value = True
    chkFults.Value = True
    
Else

    chkJon.Value = False
    chkShuler.Value = False
    chkMamatha.Value = False
    chkSrini.Value = False
    chkShoven.Value = False
    chkStokes.Value = False
    chkOwen.Value = False
    chkJenn.Value = False
    chkFults.Value = False

End If



End Sub

Private Sub CommandButton1_Click()

    Unload Me
    
    Dim emailList As String
    Dim domain As String
    
    domain = "@relationshipvelocity.com;"
    
    emailList = ""
    
    If chkShuler.Value = True Then emailList = emailList + "ashuler" + domain
    If chkMamatha.Value = True Then emailList = emailList + "mamatha" + domain
    If chkJon.Value = True Then emailList = emailList + "jon" + domain
    If chkSrini.Value = True Then emailList = emailList + "SAluri" + domain
    If chkShoven.Value = True Then emailList = emailList + "Shoven" + domain
    If chkStokes.Value = True Then emailList = emailList + "sstokes" + domain
    If chkOwen.Value = True Then emailList = emailList + "OLevy" + domain
    If chkJenn.Value = True Then emailList = emailList + "Jenn" + domain
    If chkFults.Value = True Then emailList = emailList + "JFults" + domain
    
 '"ashuler@relationshipvelocity.com;mamatha@relationshipvelocity.com;raj@relationshipvelocity.com;SAluri@relationshipvelocity.com;Shoven@relationshipvelocity.com;sstokes@relationshipvelocity.com;OLevy@relationshipvelocity.com"
    
    Call ClipBoard2.SetClipboard(emailList)
    
    'Application.ScreenUpdating = True

End Sub

Private Sub UserForm_Click()

End Sub

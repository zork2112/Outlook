Attribute VB_Name = "Folders"
Function GetFolder(ByVal FolderPath As String) As Outlook.Folder
    Dim TestFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer
        
    On Error GoTo GetFolder_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set TestFolder = Application.Session.Folders.item(FoldersArray(0))
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = TestFolder.Folders
            Set TestFolder = SubFolders.item(FoldersArray(i))
            If TestFolder Is Nothing Then
                Set GetFolder = Nothing
            End If
        Next
    End If
    'Return the TestFolder
    Set GetFolder = TestFolder
    Exit Function
        
GetFolder_Error:
    Set GetFolder = Nothing
    Exit Function
End Function

Sub TestGetFolder()
    Dim Folder As Outlook.Folder
    Set Folder = GetFolder("\\Mailbox - Dan Wilson\Inbox\Customers")
    If Not (Folder Is Nothing) Then
    Folder.Display
    End If
End Sub


Private Sub ProcessFolder(ByVal oParent As Outlook.MAPIFolder)

        Dim oFolder As Outlook.MAPIFolder
        Dim oMail As Outlook.MailItem

        For Each oMail In oParent.Items

        MsgBox oMail.ReceivedTime
        

        Next

        If (oParent.Folders.count > 0) Then
            For Each oFolder In oParent.Folders
                ProcessFolder (oFolder)
            Next
        End If
End Sub


Public Sub CheckDBMailFolderDate()
        
'Private Sub folderMaxDate(ByVal oFolderName As String)

      '  Dim oFolder As Outlook.MAPIFolder
       ' Dim oMail As Outlook.MailItem
        
        Dim oBody As String
        Dim oImportance As Integer
        Dim latestWarningEmailDate As Date
        
        Dim oProperty As UserDefinedProperty
        
        Dim oMailTo() As String
        ReDim Preserve oMailTo(0) As String  ' need to use () to declare the array then redim otherwise compiler error
        
        
        Set oFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Parent.Folders("DB Mail Status")
        
        'Set oMails = oFolder.Items
        
        Set oStatusMails = oFolder.Items.Restrict("[Subject] = 'DB Mail Status'")
        
        oStatusMails.Sort "CreationTime", True
        
        oImportance = olImportanceHigh  'olImportanceNormal  -- all emails should be high if triggered
        
        ' The date of the latest warning email is stored in the folder
        ' This checks that date to control how often to send reminder email
        
        oFolderProperty = GetFolderProperty(oFolder, "Latest Warning Email Date")
                
        If IsDate(oFolderProperty) Then
        
            latestWarningEmailDate = CDate(oFolderProperty)
        Else
        
            latestWarningEmailDate = "01/01/1900" ' Default this for comparison below
    
        End If
        
        ' To keep from getting emails sent constantly, need to check the date: n = minute, h=hour, d=day
        ' for example DateAdd("h",-2, Now()) sends one every 2 hours
        ' 20150804 since we've been getting errors that resend successfully  I changed this to every 24 hours
        If latestWarningEmailDate >= DateAdd("h", -24, Now()) Then
        
            Exit Sub
        
        End If
        
        
        ' Need to have at least 1 email otherwise the InStr will error
                
        If oStatusMails.count >= 1 Then
                    
            Set oLatestMail = oStatusMails.item(1)
            
            lastError = oLatestMail.ReceivedTime
            
            ' Check if there is an error in it - then bump up importance and send to everyone
                        
            If InStr(1, oLatestMail.Body, "error", vbTextCompare) > 0 Then  ' use Cstr() to print this
            
                oSubject = "IMPORTANT - There were Errors with DB Mail"
                
                oBody = "There is an error being report in DB Mail"
                
                ReDim Preserve oMailTo(0 To 1) As String
                oMailTo(0) = "rscholl@pentegra.com"
                oMailTo(1) = "zork2112@hotmail.com"
                
                Call SendPentegraMessage(oMailTo, oSubject, oBody, oImportance)
                
                Call SetFolderProperty(oFolder, "Latest Warning Email Date", Now(), olDateTime)
                
            
            Else
                If oLatestMail.ReceivedTime < DateAdd("d", -1, Now()) Then

                    oSubject = "DB Mail Status Report"
                    
                    oBody = "DB Mail Archive Latest Email is out of Date.  Please check to see if DB Mail is functioning."
                    
                    oMailTo(0) = "rscholl@pentegra.com"
                    
                    Call SendPentegraMessage(oMailTo, oSubject, oBody, oImportance)
                    
                    Call SetFolderProperty(oFolder, "Latest Warning Email Date", Now(), olDateTime)
                
                End If
            
            End If
            
        ' this would mean there are no emails in the folder to check
        
        Else
        
            oSubject = "DB Mail Status Report"
            
            oBody = "There are no emails in the DB Mail Status Folder"
            
            oMailTo(0) = "rscholl@pentegra.com"
            
            Call SendPentegraMessage(oMailTo, oSubject, oBody, oImportance)
            
            Call SetFolderProperty(oFolder, "Latest Warning Email Date", Now(), olDateTime)
 
        End If
        
        
'   This would be the code to iterate through emails

'        For Each oMail In oFolder.Items
'        MsgBox latestMail.ReceivedTime
'       Next oMail

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


Private Sub SetFolderProperty(ByVal oFolder As Outlook.MAPIFolder, ByVal oProperty As String, ByVal oValue As String, ByVal oType As OlUserPropertyType)
 
 
    Dim oStorage As StorageItem
    Dim oPrivateProperty As UserProperty
 
    Set oStorage = oFolder.GetStorage("Pentegra", olIdentifyBySubject)
 
    If oStorage.UserProperties(oProperty) Is Nothing Then
    
        Set oPrivateProperty = oStorage.UserProperties.Add(oProperty, oType)
        
    Else
    
        Set oPrivateProperty = oStorage.UserProperties(oProperty)
        
    End If
 
    oPrivateProperty.Value = oValue
    
    oStorage.Save
    
End Sub

Public Function GetFolderProperty(ByVal oFolder As Outlook.MAPIFolder, ByVal oProperty As String) As String

    Dim oStorage As StorageItem

    Set oStorage = oFolder.GetStorage("Pentegra", olIdentifyBySubject)
 
    If oStorage.UserProperties(oProperty) Is Nothing Then
    
        GetFolderProperty = ""
    
    Else
    
        GetFolderProperty = oStorage.UserProperties(oProperty)
    
    End If

End Function


Public Function ResetWarningMessageDate()

    Set oFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Parent.Folders("DB Mail Status")
    Call SetFolderProperty(oFolder, "Latest Warning Email Date", "01/01/1900", olDateTime)

End Function


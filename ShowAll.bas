Attribute VB_Name = "ShowAll"
Sub ShowTotalInAllFolders()
Dim oStore As Outlook.Store


On Error Resume Next

'For Each oStore In Application.Session.Stores
'Set oRoot = oStore.GetRootFolder
'ShowTotalInFolders oRoot
'Next

Dim fld As Outlook.Folder
Set fld = Application.ActiveExplorer.CurrentFolder
'fld.ShowItemCount = olShowTotalItemCount
ShowTotalInFolders fld

End Sub

Private Sub ShowTotalInFolders(ByVal Root As Outlook.Folder)
Dim oFolder As Outlook.Folder

On Error Resume Next

If Root.Folders.count > 0 Then
    For Each oFolder In Root.Folders
        oFolder.ShowItemCount = olShowUnreadItemCount
        oFolder.ShowItemCount = olShowTotalItemCount
        ShowTotalInFolders oFolder
    Next
End If

End Sub

Private Sub ShowTotal()

Dim fld As Outlook.Folder
Set fld = Application.ActiveExplorer.CurrentFolder

fld.ShowItemCount = olShowUnreadItemCount
fld.ShowItemCount = olShowTotalItemCount
MsgBox fld

End Sub


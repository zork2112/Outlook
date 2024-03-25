Attribute VB_Name = "DeleteItems"
Sub DeleteOldItems()
    
    Dim outlookApp As Outlook.Application
    Dim deleteItemFolder As Outlook.MAPIFolder
    Dim filteredItems As Outlook.Items
    Dim currentItem As Object ' Outlook item
    Dim deleteBeforeDate As Date
    
    ' Calculate the date 3 years ago
    deleteBeforeDate = DateAdd("yyyy", -3, Date)
    
    deleteBeforeDate = DateValue("02/7/2023")
    
    deleteBeforeDate = DateAdd("s", -1, DateAdd("d", 1, deleteBeforeDate))
    
    Debug.Print deleteBeforeDate
    
    ' Get the Outlook Application object
    Set outlookApp = New Outlook.Application
    
    ' Get the Inbox folder (change the folder if needed)
    ' Set inboxFolder = outlookApp.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set deleteItemFolder = outlookApp.GetNamespace("MAPI").GetDefaultFolder(olFolderDeletedItems)
    
    ' Filter items last modified before 3 years ago
    Set filteredItems = deleteItemFolder.Items.Restrict("[LastModificationTime] <= '" & Format(deleteBeforeDate, "ddddd h:nn AMPM") & "'")
    
    Debug.Print filteredItems.count
    
    ' Loop through the filtered items and delete them
    For Each currentItem In filteredItems
        currentItem.Delete
    '    Debug.Print currentItem.subject
    Next currentItem
    
    ' Optionally, display a message when the deletion is complete
    MsgBox "Filtered items deleted successfully."
    
End Sub


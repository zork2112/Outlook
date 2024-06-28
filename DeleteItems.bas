Attribute VB_Name = "DeleteItems"
Sub DeleteOldItems()
    
    Dim outlookApp As Outlook.Application
    Dim deleteItemFolder As Outlook.MAPIFolder
    Dim filteredItems As Outlook.Items
    Dim currentItem As Object ' Outlook item
    Dim deleteBeforeDate As Date
    Dim Message, Title, Default, MyValue
    Dim deleteCount As Integer
    
    
    Dim currentFolder As Outlook.MAPIFolder
    Set currentFolder = Application.ActiveExplorer.currentFolder

    ' Calculate the date 3 years ago
    deleteBeforeDate = DateAdd("yyyy", -3, Date)
 
    Message = "Enter date to delete before (MM/DD/YYYY"    ' Set prompt.
        Title = "Delete Old Items"    ' Set title.
        Default = deleteBeforeDate
        ' Display message, title, and default value.
        MyValue = InputBox(Message, Title, Default)
        
        ' Use Helpfile and context. The Help button is added automatically.
        'MyValue = InputBox(Message, Title, , , , "DEMO.HLP", 10)
        
        ' Display dialog box at position 100, 100.
        'MyValue = InputBox(Message, Title, Default, 100, 100)

    'deleteBeforeDate = DateValue("02/7/2023")
    deleteBeforeDate = DateValue(MyValue)
    
    deleteBeforeDate = DateAdd("s", -1, DateAdd("d", 1, deleteBeforeDate))
    
    
    Debug.Print deleteBeforeDate
    
    ' Get the Outlook Application object
    Set outlookApp = New Outlook.Application
    
    ' Get the Inbox folder (change the folder if needed)
    ' Set inboxFolder = outlookApp.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set deleteItemFolder = outlookApp.GetNamespace("MAPI").GetDefaultFolder(olFolderDeletedItems)
    
    Set deleteItemFolder = currentFolder ' updated to run on any folder
    
    ' Filter items last modified before 3 years ago
    'Set filteredItems = deleteItemFolder.Items.Restrict("[LastModificationTime] <= '" & Format(deleteBeforeDate, "ddddd h:nn AMPM") & "'")
    
    ' Filter by Received Time
    Set filteredItems = deleteItemFolder.Items.Restrict("[ReceivedTime] <= '" & Format(deleteBeforeDate, "ddddd h:nn AMPM") & "'")
    
    Debug.Print filteredItems.count
    
    deleteCount = filteredItems.count
    
    'MsgBox (deleteItemFolder)
    MsgBox (Format(deleteBeforeDate, "ddddd h:nn AMPM") & " " & filteredItems.count)
    'MsgBox (filteredItems.count)
    
    ' Loop through the filtered items and delete them
    For Each currentItem In filteredItems
        currentItem.Delete
    '    Debug.Print currentItem.subject
    Next currentItem
    
    ' Optionally, display a message when the deletion is complete
    MsgBox (deleteCount & " Filtered items before " + MyValue + " deleted successfully.")
    
End Sub


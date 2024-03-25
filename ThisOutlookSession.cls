VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public WithEvents myOlFolders As Outlook.Folders
Attribute myOlFolders.VB_VarHelpID = -1

Sub Initialize_handler()
 
 Set myNS = Application.GetNamespace("MAPI")

 'Set myFolders = myNS.GetDefaultFolder(olFolderDeletedItems).Folders
 Set myOlFolders = myNS.ActiveExplorer.CurrentFolder
 
End Sub
 
Private Sub myOlFolders_FolderAdd(ByVal Folder As Outlook.Folder)

    MsgBox ("Made it")
    Folder.ShowItemCount = olShowTotalItemCount

End Sub



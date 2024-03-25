Attribute VB_Name = "Clipboard"
Option Explicit

Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As LongPtr
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As LongPtr
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr



Sub InfoToClipboard()
        
        Dim mText As DataObject
        Set mText = New DataObject
        Dim outlookItem As Object
        Dim mItem As Outlook.MailItem
        
        Dim objItems As ItemProperties
        Dim objItem As ItemProperty

        'Application.ActiveExplorer.Selection(1)
        
        Set outlookItem = GetCurrentItem()
        
        ' mText.SetText "Outlook:" + outlookItem.EntryID  yyyymmdd
        'mText.SetText " -- " + outlookItem.Subject + " [" + outlookItem.SenderName + " " + Format(outlookItem.ReceivedTime, "m/d/yy h:nn am/pm") + "]"
        
        mText.SetText "Goodbye"
        
'           OpenClipboard (0)
 '               EmptyClipboard
  '          CloseClipboard
       ' Call ClipBoard2.SetClipboard(" -- " + outlookItem.subject + " [" + outlookItem.SenderName + " " + Format(outlookItem.ReceivedTime, "m/d/yy h:nn am/pm") + "]")
       SetText (" -- " + outlookItem.subject + " [" + outlookItem.SenderName + " " + Format(outlookItem.ReceivedTime, "m/d/yy h:nn am/pm") + "]")
       
        'mText.PutInClipboard
        
        

End Sub
Sub InfoToClipboardFollowup()
        
        Dim mText As DataObject
        Set mText = New DataObject
        Dim outlookItem As Object
        Dim mItem As Outlook.MailItem
        
        Dim mRecipient As Outlook.Recipient
        Dim recipientName As String
        
        Dim objItems As ItemProperties
        Dim objItem As ItemProperty

        'Application.ActiveExplorer.Selection(1)
        
        Set outlookItem = GetCurrentItem()
        
        
        ' mText.SetText "Outlook:" + outlookItem.EntryID  yyyymmdd
        'mText.SetText " -- " + outlookItem.Subject + " [" + outlookItem.SenderName + " " + Format(outlookItem.ReceivedTime, "m/d/yy h:nn am/pm") + "]"
        Set mRecipient = outlookItem.Recipients.item(1)
        
        recipientName = mRecipient.Name
       
        
        mText.SetText "Goodbye"
        
'           OpenClipboard (0)
 '               EmptyClipboard
  '          CloseClipboard
        'Call ClipBoard2.SetClipboard("FU - " + fuckYouOutlook + " -- " + outlookItem.subject + " [" + outlookItem.SenderName + " " + Format(outlookItem.ReceivedTime, "m/d/yy h:nn am/pm") + "]")
        SetText ("FU - " + recipientName + " -- " + outlookItem.subject + " [" + outlookItem.SenderName + " " + Format(outlookItem.ReceivedTime, "m/d/yy h:nn am/pm") + "]")
        'mText.PutInClipboard
        
        

End Sub
'Adds a link to the currently selected message to the clipboard
Sub AddLinkToMessageInClipboard()

   Dim objMail As Outlook.MailItem
   Dim doClipboard As New DataObject

   'One and ONLY one message muse be selected
   If Application.ActiveExplorer.Selection.count <> 1 Then
       MsgBox ("Select one and ONLY one message.")
       Exit Sub
   End If

   Set objMail = Application.ActiveExplorer.Selection.item(1)
   'doClipboard.SetText "[[outlook:" + objMail.EntryID + "][MESSAGE: " + objMail.subject + " (" + objMail.SenderName + ")]]"
   doClipboard.SetText "outlook:" + objMail.EntryID
   doClipboard.PutInClipboard

End Sub


Sub RVAddressesToClipboard()
        
        Call Clipboard.SetText("ashuler@relationshipvelocity.com;mamatha@relationshipvelocity.com;raj@relationshipvelocity.com;SAluri@relationshipvelocity.com;Shoven@relationshipvelocity.com;sstokes@relationshipvelocity.com;OLevy@relationshipvelocity.com")
       

End Sub

Public Sub SetText(Text As String)

Dim hGlobalMemory As LongPtr
Dim lpGlobalMemory As LongPtr
Dim hClipMemory As LongPtr

Const GHND = &H42
Const CF_TEXT = 1

   ' Allocate moveable global memory.
   '-------------------------------------------
   hGlobalMemory = GlobalAlloc(GHND, Len(Text) + 1)

   ' Lock the block to get a far pointer
   ' to this memory.
   lpGlobalMemory = GlobalLock(hGlobalMemory)

   ' Copy the string to this global memory.
   lpGlobalMemory = lstrcpy(lpGlobalMemory, Text)

   ' Unlock the memory.
   If GlobalUnlock(hGlobalMemory) <> 0 Then
      MsgBox "Could not unlock memory location. Copy aborted."
      GoTo CloseClipboard
   End If

   ' Open the Clipboard to copy data to.
   If OpenClipboard(0&) = 0 Then
      MsgBox "Could not open the Clipboard. Copy aborted."
      Exit Sub
   End If

   ' Clear the Clipboard.
   Call EmptyClipboard

   ' Copy the data to the Clipboard.
   hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

CloseClipboard:

   If CloseClipboard() = 0 Then
      MsgBox "Could not close Clipboard."
   End If

End Sub

Public Property Get GetText()

Dim hClipMemory As LongPtr
Dim lpClipMemory As LongPtr

Dim MaximumSize As Long
Dim ClipText As String

Const CF_TEXT = 1

   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Property
   End If

   ' Obtain the handle to the global memory block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo CloseClipboard
   End If

   ' Lock Clipboard memory so we can reference the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)

   If Not IsNull(lpClipMemory) Then
      MaximumSize = 64

      Do
        MaximumSize = MaximumSize * 2

        ClipText = Space$(MaximumSize)
        Call lstrcpy(ClipText, lpClipMemory)
        Call GlobalUnlock(hClipMemory)

      Loop Until ClipText Like "*" & vbNullChar & "*"

      ' Peel off the null terminating character.
      ClipText = Left$(ClipText, InStrRev(ClipText, vbNullChar) - 1)

   Else
      MsgBox "Could not lock memory to copy string from."
   End If

CloseClipboard:

   Call CloseClipboard
   GetText = ClipText

End Property


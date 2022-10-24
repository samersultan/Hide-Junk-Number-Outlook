'define variables on startup

Dim WithEvents g_DelFolder As Outlook.Items
Dim WithEvents g_JunkFolder As Outlook.Items
Dim WithEvents myOlExp As Outlook.Explorer

'release variables on quit

Private Sub Application_Quit()
Set g_DelFolder = Nothing
Set g_JunkFolder = Nothing
End Sub

'Set variable values

Private Sub Application_Startup()
Set g_DelFolder = Session.GetDefaultFolder(olFolderDeletedItems).Items
Set g_JunkFolder = Session.GetDefaultFolder(olFolderJunk).Items
Set myOlExp = Application.ActiveExplorer
End Sub

'If an item is deleted, mark it Read.

Private Sub g_DelFolder_ItemAdd(ByVal Item As Object)
Item.UnRead = False
Item.Save
End Sub

'Call JunkRead sub (uses 'Fake' input because Outlook Rules requires it). Function input does nothing and returns nothging.
'Note, I couldn't get the above code to work for Junk Email, because Outlook sees sending items to Junk Mail as "Block Sender", and not "Item Add".
'I couldn't find an Outlook event for "Block Sender". That's why I created the SelectionChange Event Handlers below.

Sub JunkReadOnReceive(Fake As Outlook.MailItem)

JunkRead

End Sub

'Call JunkRead and DelRead upon each selection change (effectively eliminates the possiblility of any items in either Junk or Deleted Items to be marked UnRead)

Private Sub myOlExp_SelectionChange()

 JunkRead

 'Calling DelRead isn't strictly necessary. The above ItemAdd handler will always ensure anything that is deleted is marked read.
 'Calling it here will prevent anything in the delete folder from being marked unread in the future.
 DelRead

End Sub

'Checks whether any unread items exist in the Junk folder and marks them as read.

Private Sub JunkRead()

Dim objJunk As Outlook.MAPIFolder
Dim objOutlook As Object, objnSpace As Object, objMessage As Object

Set objOutlook = CreateObject("Outlook.Application")
Set objnSpace = objOutlook.GetNamespace("MAPI")
Set objJunk = objnSpace.GetDefaultFolder(olFolderJunk)

For Each objMessage In objJunk.Items
    If objMessage.UnRead = True Then
        objMessage.UnRead = False
    End If
Next

Set objOutlook = Nothing
Set objnSpace = Nothing
Set objJunk = Nothing

End Sub

'Checks whether any unread items exist in the Deleted Items folder and marks them as read.

Private Sub DelRead()

Dim objDel As Outlook.MAPIFolder
Dim objOutlook As Object, objnSpace As Object, objMessage As Object

Set objOutlook = CreateObject("Outlook.Application")
Set objnSpace = objOutlook.GetNamespace("MAPI")
Set objDel = objnSpace.GetDefaultFolder(olFolderDeletedItems)

For Each objMessage In objDel.Items
    If objMessage.UnRead = True Then
        objMessage.UnRead = False
    End If
Next

Set objOutlook = Nothing
Set objnSpace = Nothing
Set objDel = Nothing

End Sub


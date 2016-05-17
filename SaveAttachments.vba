' From http://www.slipstick.com/developer/save-attachments-to-the-hard-drive/
' Modified to save identically named attachements as `foo (<number>)'

' 1. Select all the emails you want to save attachments from
' 2. Run this macro
' 3. They'll be saved to ~\My Documents\OLAttachments\

Public Sub SaveAttachments()
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String

    ' Get the path to your My Documents folder
    strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)
    On Error Resume Next

    ' Instantiate an Outlook Application object.
    Set objOL = CreateObject("Outlook.Application")

    ' Get the collection of selected objects.
    Set objSelection = objOL.ActiveExplorer.Selection

' The attachment folder needs to exist
' You can change this to another folder name of your choice

    ' Set the Attachment folder.
    strFolderpath = strFolderpath & "\OLAttachments\"

    ' Check each selected item for attachments.
    For Each objMsg In objSelection

    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count

    If lngCount > 0 Then

    ' Use a count down loop for removing items
    ' from a collection. Otherwise, the loop counter gets
    ' confused and only every other item is removed.

    For i = lngCount To 1 Step -1

    ' Get the file name.
    strFile = objAttachments.Item(i).FileName

    ' Combine with the path to the Temp folder.
    strFile = strFolderpath & strFile
    
    Dim strBase As String, strExt As String
    strBase = CreateObject("Scripting.FileSystemObject").GetBaseName(strFile)
    strExt = CreateObject("Scripting.FileSystemObject").GetExtensionName(strFile)
    Dim fileno As Integer
    fileno = 2
    Do While (Not Dir(strFile) = vbNullString)
        strFile = strFolderpath & strBase & " (" & fileno & ")." & strExt
        fileno = fileno + 1
    Loop

    ' Save the attachment as a file.
    objAttachments.Item(i).SaveAsFile strFile

    Next i
    End If

    Next

ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Sub

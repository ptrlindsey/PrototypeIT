'''''''''''''''''''
'     AutoBCC     '
'''''''''''''''''''
' When the user replies to the target, automatically BCC another account
Private Sub Application_ItemSend(ByVal objItem As Object, Cancel As Boolean)
    Dim targetAddress As String
    Dim bccAddress As String
    targetAddress = <TARGET EMAIL ADDRESS>
    bccAddress = <BCC EMAIL ADDRESS>
    Dim mi As MailItem

    If TypeName(objItem) = "MailItem" Then
        Set mi = objItem

        Dim rc As Recipient
        For Each rc In mi.Recipients
            If StrComp(rc.address, targetAddress, vbTextCompare) = 0 Then
              ' Check length to see whether it's a new message or a reply (see http://stackoverflow.com/questions/36412152/)
                Dim indexLen
                indexLen = Len(mi.ConversationIndex)
                If indexLen = 44 Then ' New conversation
                    ' We don't want to interfere with new conversations
                    Exit Sub
                Else ' Reply
                    Dim bcc As Recipient
                    Set bcc = mi.Recipients.Add(bccAddress)
                    bcc.Type = olBCC
                    bcc.Resolve
                End If
                Exit For
            End If
        Next
    End If
End Sub

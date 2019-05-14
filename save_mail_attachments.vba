Sub CreateFolder(sFolder As String)
    If Len(Dir(sFolder, vbDirectory)) = 0 Then
        MkDir sFolder
    End If
End Sub

Private Sub Application_NewMail()
    Dim onamespace As Outlook.NameSpace
    Set onamespace = Outlook.GetNamespace("MAPI")
    
    Dim myfol As Outlook.Folder
    Set myfol = onamespace.GetDefaultFolder(olFolderInbox)
    
    Dim destfol As Outlook.Folder
    Set destfol = myfol.Folders("Pasta de Emails")

    Dim omail As Outlook.MailItem
    Set omail = Outlook.CreateItem(olMailItem)

    Dim atmt As Outlook.attachment
    Dim assunto As String

    For Each omail In myfol.Items
            If omail.SenderEmailType = "EX" Then
                If omail.sender.GetExchangeUser.PrimarySmtpAddress = "xxxxxxx@xxxxxxxxx.com.br" Then
                    For Each atmt In omail.Attachments
                        assunto = omail.Subject
                        CreateFolder ("P:\Pasta_compartilhada\" & assunto)
                        atmt.SaveAsFile "P:\Pasta_compartilhada" & assunto & "\" & atmt.fileName
                    Next
                    omail.Move destfol
                End If
            End If
    Next
End Sub

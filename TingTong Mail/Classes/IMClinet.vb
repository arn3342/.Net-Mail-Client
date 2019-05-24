Imports ImapX
Imports ImapX.Collections

Namespace IMClient
    Class ImpClient
        Public Shared Property Errormessage As String
        Public Shared Property ImapC As String
        Public Shared Property port As Integer
        Private Shared Property client As ImapX.ImapClient
        Class emailFolder
            Public Property Title As String
            Public Overrides Function ToString() As String
                Return Me.Title
            End Function
        End Class
        Public Shared Sub Initialize()
            If My.Computer.Network.IsAvailable = True Then
                client = New ImapClient(ImapC, port, True)
                If Not client.Connect() Then
                    Errormessage = "Error 8 : Couldn't connect to server."
                End If
            Else
                Errormessage = "Error 4 : Please make sure you're connected to the internet."
            End If
        End Sub
        Public Shared Function Login(ByVal u As String, ByVal p As String) As Boolean
            If Errormessage = "" Then
                Return client.Login(u, p)
            Else
                Errormessage = Errormessage + vbCrLf + "Error 23 : Please re-check your email and password."
            End If
        End Function
        Public Shared Sub Logout()
            If client.IsAuthenticated Then
                client.Logout()
            End If
        End Sub
        Public Shared Function GetFolders() As List(Of emailFolder)
            Dim folder = New List(Of emailFolder)
            Dim foldername = client.Folders
            For Each parentFolder In foldername
                Dim parentPath = parentFolder.Path
                If parentFolder.HasChildren Then
                    Dim subfolders = parentFolder.SubFolders
                    For Each subfolder In subfolders
                        folder.Add(New emailFolder With {.Title = subfolder.Name})
                    Next
                    ' If parent folder has not been added above, do it here.
                Else
                    folder.Add(New emailFolder With {.Title = parentFolder.Name})
                End If
            Next
            client.Folders.Inbox.StartIdling()
            AddHandler client.Folders.Inbox.OnNewMessagesArrived, AddressOf Inbox_OnNewMessagesArrived
            Return folder
        End Function
        Private Shared Sub Inbox_OnNewMessagesArrived(ByVal sender As Object, ByVal e As IdleEventArgs)
            MessageBox.Show($"A new message was downloaded in {e.Folder.Name} folder.")
        End Sub
        Public Shared Function GetMessagesForFolder(ByVal name As String) As MessageCollection
            client.Folders(name).Messages.Download()
            Return client.Folders(name).Messages
        End Function
    End Class
End Namespace
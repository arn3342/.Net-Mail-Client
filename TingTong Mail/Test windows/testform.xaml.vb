Imports MimeKit
Imports MailKit
Imports MailKit.Net.Imap
Imports System
Imports iclinet
Imports System.ComponentModel
Imports System.Threading
Imports System.IO
Imports WpfAnimatedGif
Imports System.Data

Public Class testform
    Public ItemList As New List(Of String)
    Public ItemList2 As New List(Of String)
    Public Sub btn_Add()
        ItemList.Add("StringHere")
    End Sub
    Public Sub mailkit()
        Using client = New ImapClient()
            client.ServerCertificateValidationCallback = Function(s, c, h, e) True
            client.Connect("imap.gmail.com", 993, True)
            client.Authenticate("nabilrashid44@gmail.com", "arn33423342")
            Dim folders = client.GetFolders(client.PersonalNamespaces(0))
            For Each fol In folders
                MsgBox(fol.Name)
            Next
        End Using
    End Sub
    Private Sub testform_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ItemList.Add("<body>")
        ItemList.Add("<lc:default name=CSS />")
        ItemList.Add("</body>")
        ItemList2.Add("<link rel=""icon"" type=""mage/png"" href=""/IMAGES/fav-icon.png"" />")
        ItemList2.Add("<link href=""https: //fonts.googleapis.com/css?family=Oswald"" rel=""stylesheet"">")
        ItemList2.Add("<link rel=""stylesheet"" href=""https: //maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css"" integrity=""sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm"" crossorigin=""anonymous"">")
        Dim index As Integer
        Dim foundWordr As String = ""
        For Each ite In ItemList
            If ite.Contains("<lc:default") Then
                index = ItemList.FindIndex(Function(a) a = ite)
                foundWordr = ite
            End If
        Next

        ItemList(index).Replace(foundWordr, "")
        ItemList.Add("<body>")
        ItemList.Add("</body>")

        For Each item In ItemList2
            ItemList(index) = item
            index = index + 1
        Next

        For Each item In ItemList
            MsgBox(item)
        Next
    End Sub
    Public Sub imap()
        Dim int As Integer = My.Settings.txtport
        IMClient.ImpClient.ImapC = "imap.mail.yahoo.com"
        IMClient.ImpClient.port = int
        IMClient.ImpClient.Initialize()
        If IMClient.ImpClient.Login("testmail3342@yahoo.com", "testpassword") Then
            For Each item In IMClient.ImpClient.GetFolders
                Dim FolderName As String = item.ToString
                'EmailFolderList.Add(FolderName)
            Next
        End If
    End Sub



End Class

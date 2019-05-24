Public Class smallEmailContainer
    Public WithEvents DetectUnSelected As New Forms.Timer
    Dim ttm As Startup
    Private Sub Main_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles Main.MouseDown
        DownloadMail()
        DetectUnSelected.Start()
        SelectedRect.Visibility = Visibility.Visible
    End Sub

    Private Sub DetectUnSelected_Tick(sender As Object, e As EventArgs) Handles DetectUnSelected.Tick
        If Not My.Settings.OpenedMailId = UniqueId.Text Then
            DetectUnSelected.Stop()
            SelectedRect.Visibility = Visibility.Hidden
        End If
    End Sub
    Private Sub DownloadMail()
        My.Settings.OpenedMailId = UniqueId.Text
        Dim Folder = ttm.MailKitClient.GetFolder(My.Settings.CurrentMailFolderName)
        Dim CurrentMailFolder As String = System.IO.Directory.GetCurrentDirectory() + "\Data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\"
        Dim Message = Folder.GetMessage(MailKit.UniqueId.Parse(UniqueId.Text))
        Dim emlPath As String = CurrentMailFolder + UniqueId.Text + "
"
        Message.WriteTo(emlPath)
    End Sub
End Class

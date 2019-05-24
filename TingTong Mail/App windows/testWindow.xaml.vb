Public Class testWindow
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        For i = 0 To 100
            Dim a As New MailContainer1
            a.EdateText = "10\12\13 2015/12/13"
            a.FromText = "jlakdjlkadkljasdkjkasd"
            a.UniqueIdText = "454875"
            a.SubjectText = "a quick brown fox jumps over a lazy dog and it starts barking like an asshole and bla bla bla"
            Teststack.Height += 50
            Teststack.Children.Add(a)
        Next
        For i = 0 To 10
            MessageFoldersList.stp.Height += 26
            MessageFoldersList.stp.Children.Add(New Button With {.Height = 26,
                                               .Content = "abc"})
        Next
    End Sub
End Class

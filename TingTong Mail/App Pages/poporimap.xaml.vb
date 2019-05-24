Class poporimap
    Dim sep As New SetUp_mail
    Private Sub imap_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles imap.MouseDown
        My.Settings.imappop = "IMAP"
        Me.NavigationService.Navigate(sep)
    End Sub

    Private Sub pop_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles pop.MouseDown
        My.Settings.imappop = "Pop"
        Me.NavigationService.Navigate(sep)
    End Sub

    Private Sub poporimap_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        grid1.Opacity = 0
    End Sub
End Class


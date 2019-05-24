Public Class LoadMoreEmails
    Private Sub LoadOlderMails_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles LoadOlderMails.MouseDown
        CType(Me.Parent, VirtualizingStackPanel).Children.Remove(Me)
    End Sub
End Class

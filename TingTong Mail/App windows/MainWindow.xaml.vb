Imports System.ComponentModel

Class MainWindow
    Dim dd As New LogIn

    Private Sub MainWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

    End Sub
    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        grid1.Opacity = 0
        frame.Navigate(dd)
        frame.NavigationUIVisibility = NavigationUIVisibility.Hidden
    End Sub
End Class

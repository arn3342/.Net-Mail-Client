Imports System.Windows.Controls.Primitives
Public Class SearchBox
    Inherits System.Windows.Controls.Control
    Public CloseBtn As Image
    Public Property Searchtextbox As TextBox
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        Searchtextbox = GetTemplateChild("Searchtextbox")
        CloseBtn = GetTemplateChild("CloseBtn")
    End Sub
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(SearchBox), New FrameworkPropertyMetadata(GetType(SearchBox)))
    End Sub
    Private Sub Searchtextbox_GotFocus(sender As Object, e As RoutedEventArgs)
        If Searchtextbox.Text = "Search here" Then
            Searchtextbox.Text = ""
        End If
    End Sub

    Private Sub Searchtextbox_LostFocus(sender As Object, e As RoutedEventArgs)
        If Searchtextbox.Text = "" Then
            Searchtextbox.Text = "Search here"
        End If
    End Sub
    Private Sub CloseBtn_MouseDown(sender As Object, e As MouseButtonEventArgs)
        Me.Visibility = Visibility.Hidden
    End Sub
End Class

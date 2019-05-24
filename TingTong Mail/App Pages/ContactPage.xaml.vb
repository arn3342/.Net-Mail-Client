Imports System.Windows.Media.Animation

Class ContactPage
    Dim MainParent As Frame
    Private Sub NoRecord()
        Main.Children.Add(New TextBlock With {.Style = FindResource("CustomTextStyle")})
    End Sub
    Private Sub CloseBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles CloseBtn.MouseDown
        Dim sb As New Storyboard
        sb = FindResource("HideMainframeForContact")
        Storyboard.SetTarget(sb, MainParent)
        AddHandler sb.Completed, AddressOf AnimationDone
        sb.Begin()
    End Sub
    Private Sub AnimationDone()
        MainParent.Width = Double.NaN
        MainParent.Visibility = Visibility.Hidden
        MainParent.Opacity = 1
        MainParent.Navigate("")
        ConContainer.Children.Clear()
    End Sub
    Private Sub ContactPage_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        MainParent = FindParent(Of Frame)(Me)
    End Sub
    Private Sub ContactPage_Unloaded(sender As Object, e As RoutedEventArgs) Handles Me.Unloaded
        ConContainer.Children.Clear()
    End Sub

    Public Shared Function FindParent(Of T As DependencyObject)(ByVal child As DependencyObject) As T
        Dim parentObject As DependencyObject = VisualTreeHelper.GetParent(child)
        If parentObject Is Nothing Then Return Nothing
        Dim parent As T = TryCast(parentObject, T)
        If parent IsNot Nothing Then Return parent Else Return FindParent(Of T)(parentObject)
    End Function
End Class

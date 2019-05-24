Imports System.Windows.Media.Animation
Public Class FolderCreation
    Inherits System.Windows.Controls.Control
    Dim myResourceDictionary As New ResourceDictionary
    Dim MainBorder As Border
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(FolderCreation), New FrameworkPropertyMetadata(GetType(FolderCreation)))
    End Sub
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        MainBorder = GetTemplateChild("MainBorder")
        myResourceDictionary.Source = New Uri("pack://application:,,,/TingTong Mail;component/Themes/Generic.xaml", UriKind.RelativeOrAbsolute)
    End Sub
    Private Sub FolderCreation_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim sb As New Storyboard
        sb = myResourceDictionary("BringUp")
        Storyboard.SetTarget(sb, MainBorder)
        sb.Begin()
    End Sub
End Class

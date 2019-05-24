Imports System.Windows.Media.Animation
Public Class FilteringOptions
    Inherits System.Windows.Controls.Control
    Private FilOpGrid As Grid
    Dim myResourceDictionary As New ResourceDictionary()
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(FilteringOptions), New FrameworkPropertyMetadata(GetType(FilteringOptions)))
    End Sub
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        FilOpGrid = GetTemplateChild("FilOpGrid")
    End Sub
    Private Sub FilteringOptions_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        myResourceDictionary.Source = New Uri("pack://application:,,,/TingTong Mail;component/Themes/Generic.xaml", UriKind.RelativeOrAbsolute)
        Dim sb As New Storyboard
        sb = CType(myResourceDictionary("Expdd2"), Storyboard)
        Storyboard.SetTarget(sb, FilOpGrid)
        sb.Begin()
    End Sub
End Class

Imports System.Windows.Media.Animation

Public Class CurUPanel
    Inherits System.Windows.Controls.Control
    Public UserEmail As TextBlock
    Public FirstChar As TextBlock
    Private MainGd As Grid
    Dim myResourceDictionary As ResourceDictionary
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        myResourceDictionary = New ResourceDictionary()
        myResourceDictionary.Source = New Uri("pack://application:,,,/TingTong Mail;component/Themes/Generic.xaml", UriKind.RelativeOrAbsolute)
        UserEmail = GetTemplateChild("UserEmail")
        FirstChar = GetTemplateChild("FirstChar")
        Dim ManageAcc As Rectangle = GetTemplateChild("ManageAcc")
        AddHandler ManageAcc.MouseDown, AddressOf ManageAcc_MouseDown

        MainGd = GetTemplateChild("FocRect")
    End Sub

    Private Sub ManageAcc_MouseDown(sender As Object, e As MouseButtonEventArgs)
        Dim manageacc As New ManageAcc
        manageacc.ShowDialog()
    End Sub
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(CurUPanel), New FrameworkPropertyMetadata(GetType(CurUPanel)))
    End Sub

    Private Sub CurUPanel_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        UserEmail.Text = My.Settings.email
        FirstChar.Text = My.Settings.email.ToString.Substring(0, 1).ToUpper
        Dim sb As New Storyboard
        sb = CType(myResourceDictionary("Expdd"), Storyboard)
        Storyboard.SetTarget(sb, MainGd)
        sb.Begin()
    End Sub
End Class

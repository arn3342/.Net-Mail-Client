Imports System.Windows.Media.Animation
Imports System.Windows.Controls.Primitives
Public Class Folders
    Inherits System.Windows.Controls.Control
    Public HdrCon As TextBlock
    Public OptionsBtn As SmallButtons
    Dim indicator As Path
    Dim bor As Border
    Dim MsEnterRect As Rectangle
    Dim myResourceDictionary As ResourceDictionary
    Public Shared HeaderContentDependency As DependencyProperty = DependencyProperty.Register("HeaderContent", GetType(String), GetType(Folders))
    Shared Sub New()
        'This OverrideMetadata call tells the system that this element wants to provide a style that is different than its base class.
        'This style is defined in themes\generic.xaml
        DefaultStyleKeyProperty.OverrideMetadata(GetType(Folders), New FrameworkPropertyMetadata(GetType(Folders)))
    End Sub
    Public Property HeaderText As String
        Get
            Return CStr(GetValue(HeaderContentDependency))
        End Get
        Set(value As String)
            SetValue(HeaderContentDependency, value)
        End Set
    End Property
    Public Overrides Sub OnApplyTemplate()
        myResourceDictionary = New ResourceDictionary()
        myResourceDictionary.Source = New Uri("pack://application:,,,/TingTong Mail;component/Themes/Generic.xaml", UriKind.RelativeOrAbsolute)
        MyBase.OnApplyTemplate()
        bor = GetTemplateChild("Border")
        MsEnterRect = GetTemplateChild("MsEnterRect")
        Dim Gdd As Grid = GetTemplateChild("Gdd")
        AddHandler Gdd.MouseEnter, AddressOf FolderCntrol_MouseEnter
        AddHandler Gdd.MouseLeave, AddressOf FolderCntrol_MouseLeave
        HdrCon = GetTemplateChild("Header")
        Dim HeaderContentBinding As Binding = New Binding("HeaderContent") With {
        .Source = Me,
          .Mode = BindingMode.TwoWay
          }
        HdrCon.SetBinding(TextBlock.TextProperty, HeaderContentBinding)
        OptionsBtn = GetTemplateChild("OptionsBtn")
        AddHandler OptionsBtn.MouseDown, AddressOf OptionsBtn_MouseDown
    End Sub
    Private Sub FolderCntrol_MouseEnter(sender As Object, e As MouseEventArgs)
        Dim sb As New Storyboard
        sb = CType(myResourceDictionary("Isfocused"), Storyboard)
        Storyboard.SetTarget(sb, MsEnterRect)
        sb.Begin()
    End Sub

    Private Sub FolderCntrol_MouseLeave(sender As Object, e As MouseEventArgs)
        Dim sb As New Storyboard
        sb = CType(myResourceDictionary("Notfocused"), Storyboard)
        Storyboard.SetTarget(sb, MsEnterRect)
        sb.Begin()
    End Sub
    Private Sub OptionsBtn_MouseDown()
    End Sub
End Class

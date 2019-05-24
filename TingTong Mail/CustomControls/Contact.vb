Imports System.Windows.Controls.Primitives
Public Class Contact
    Inherits System.Windows.Controls.Control
    Public MainGrid As Grid
    Private ConMail As TextBlock
    Private ConName As TextBlock
    Dim Ic As Integer = 0
    Public IsChecked As Boolean
    Dim bb As New BrushConverter
    Private CheckBx As CustomCheckbox
    Public Shared ReadOnly ContactNameDependency As DependencyProperty = DependencyProperty.Register("Contact Name", GetType(String), GetType(Contact))
    Public Shared ReadOnly ContactEmailDependency As DependencyProperty = DependencyProperty.Register("Contact Email", GetType(String), GetType(Contact))
    Public Property ContactName As String
        Get
            Return CStr(GetValue(ContactNameDependency))
        End Get
        Set(value As String)
            SetValue(ContactNameDependency, value)
        End Set
    End Property
    Public Property ContactEmail As String
        Get
            Return CStr(GetValue(ContactEmailDependency))
        End Get
        Set(value As String)
            SetValue(ContactEmailDependency, value)
        End Set
    End Property
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(Contact), New FrameworkPropertyMetadata(GetType(Contact)))
    End Sub
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        ConMail = GetTemplateChild("ConMail")
        ConName = GetTemplateChild("ConName")
        CheckBx = GetTemplateChild("ChkBx")
        AddHandler CheckBx.MouseLeftButtonDown, AddressOf CheckBx_MouseLeftButtonDown
        Dim ContactNameBinding As Binding = New Binding("Contact Name") With {
            .Source = Me,
            .Mode = BindingMode.TwoWay}
        ConName.SetBinding(TextBlock.TextProperty, ContactNameBinding)
        Dim ContactEmailBinding As Binding = New Binding("Contact Email") With {
         .Source = Me,
         .Mode = BindingMode.TwoWay}
        ConMail.SetBinding(TextBlock.TextProperty, ContactEmailBinding)
    End Sub
    Private Sub CheckBx_MouseLeftButtonDown()
        IsChecked = CheckBx.Checked
    End Sub
End Class

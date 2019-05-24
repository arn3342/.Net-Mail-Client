Imports System.Windows.Controls.Primitives
Imports System.Windows.Controls.Image
Public Class EmailDetails
    Inherits System.Windows.Controls.Control
    Public From As TextBlock
    Public ETo As TextBlock
    Public Cc As TextBlock
    Public Attach As Canvas
    Public Subject As TextBlock
    Public Shared ReadOnly FromDependency As DependencyProperty = DependencyProperty.Register("FromText", GetType(String), GetType(EmailDetails))
    Public Shared ReadOnly EToDependency As DependencyProperty = DependencyProperty.Register("EToText", GetType(String), GetType(EmailDetails))
    Public Shared ReadOnly CCDependency As DependencyProperty = DependencyProperty.Register("CCText", GetType(String), GetType(EmailDetails))
    Public Shared ReadOnly AttachDependency As DependencyProperty = DependencyProperty.Register("AttachVisibility", GetType(Visibility), GetType(EmailDetails))
    Public Shared ReadOnly SubjectDependency As DependencyProperty = DependencyProperty.Register("SubjectText", GetType(String), GetType(EmailDetails))

    Public Property AttachVisibility As Visibility
        Get
            Return CType(GetValue(AttachDependency), Visibility)
        End Get
        Set(ByVal value As Visibility)
            SetValue(AttachDependency, value)
        End Set
    End Property
    Public Property SubjectText As String
        Get
            Return CStr(GetValue(SubjectDependency))
        End Get

        Set(ByVal value As String)
            SetValue(SubjectDependency, value)
        End Set
    End Property
    Public Property EToText As String
        Get
            Return CStr(GetValue(EToDependency))
        End Get
        Set(ByVal value As String)
            SetValue(EToDependency, value)
        End Set
    End Property
    Public Property CcText As String
        Get
            Return CStr(GetValue(CCDependency))
        End Get
        Set(ByVal value As String)
            SetValue(CCDependency, value)
        End Set
    End Property
    Public Property FromText As String
        Get
            Return CStr(GetValue(FromDependency))
        End Get
        Set(ByVal value As String)
            SetValue(FromDependency, value)
        End Set
    End Property
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(EmailDetails), New FrameworkPropertyMetadata(GetType(EmailDetails)))
    End Sub
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        From = GetTemplateChild("From")

        Dim FromTextBinding As Binding = New Binding("FromText")
        FromTextBinding.Source = Me
        FromTextBinding.Mode = BindingMode.TwoWay
        From.SetBinding(TextBlock.TextProperty, FromTextBinding)

        ETo = GetTemplateChild("To")
        Dim EToTextBinding As Binding = New Binding("EToText")
        EToTextBinding.Source = Me
        EToTextBinding.Mode = BindingMode.TwoWay
        ETo.SetBinding(TextBlock.TextProperty, EToTextBinding)

        Cc = GetTemplateChild("Cc")
        Dim CcTextBinding As Binding = New Binding("CCText")
        CcTextBinding.Source = Me
        CcTextBinding.Mode = BindingMode.TwoWay
        Cc.SetBinding(TextBlock.TextProperty, CcTextBinding)

        Attach = GetTemplateChild("AttachmentCanvas")
        Dim VisibilityBinding As Binding = New Binding("AttachVisibility")
        VisibilityBinding.Source = Me
        VisibilityBinding.Mode = BindingMode.TwoWay
        Attach.SetBinding(Image.VisibilityProperty, VisibilityBinding)

        Subject = GetTemplateChild("Subject")
        Dim SubjectTextBinding As Binding = New Binding("SubjectText")
        SubjectTextBinding.Source = Me
        SubjectTextBinding.Mode = BindingMode.TwoWay
        Subject.SetBinding(TextBlock.TextProperty, SubjectTextBinding)
        AttachVisibility = Visibility.Hidden
    End Sub
End Class

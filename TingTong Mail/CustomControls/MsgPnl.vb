Imports System.Windows.Controls.Primitives
Public Class MsgPnl
    Inherits System.Windows.Controls.Control
    Public CurrentFolder As TextBlock
    Private CurFilter As Rectangle
    Public Shared HeaderContentDependency As DependencyProperty = DependencyProperty.Register("HdrText", GetType(String), GetType(FolderCntrol))
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(MsgPnl), New FrameworkPropertyMetadata(GetType(MsgPnl)))
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
        MyBase.OnApplyTemplate()
        CurrentFolder = GetTemplateChild("CurFolName")
        CurrentFolder.Text = "abc"
        Dim HeaderContentBinding As Binding = New Binding("HeaderText") With {
        .Source = Me,
         .Mode = BindingMode.TwoWay
          }
        CurrentFolder.SetBinding(TextBlock.TextProperty, HeaderContentBinding)

        CurFilter = GetTemplateChild("CurFilter")
        AddHandler CurFilter.MouseLeftButtonDown, AddressOf CurFilter_MouseLeftButtonDown
    End Sub
    Private Sub CurFilter_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim cm = ContextMenuService.GetContextMenu(TryCast(sender, DependencyObject))
        If cm Is Nothing Then
            Return
        Else
            cm.Placement = PlacementMode.Bottom
            cm.PlacementTarget = TryCast(sender, UIElement)
            cm.IsOpen = True
        End If
    End Sub

End Class

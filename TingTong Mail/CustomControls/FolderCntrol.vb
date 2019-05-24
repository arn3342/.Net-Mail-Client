Imports System.ComponentModel
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media.Animation

Public Class FolderCntrol
    Inherits System.Windows.Controls.Control
    Public HdrCon As TextBlock
    Public stp As StackPanel
    Public IsExpanded As Boolean
    Public ExpBtn As Rectangle
    Dim indicator As Path
    Dim bor As Border
    Dim MsEnterRect As Rectangle
    Dim myResourceDictionary As New ResourceDictionary
    Public FolderList As New List(Of String)
    Public Shared HeaderContentDependency As DependencyProperty = DependencyProperty.Register("HeaderContent", GetType(String), GetType(FolderCntrol))
    ''' <summary>
    ''' Gets or sets the header of the FolderControl
    ''' </summary>
    ''' <returns></returns>
    <Description("Gets or sets the header of the FolderControl")>
    Public Property HeaderText As String
        Get
            Return CStr(GetValue(HeaderContentDependency))
        End Get
        Set(value As String)
            SetValue(HeaderContentDependency, value)
        End Set
    End Property
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(FolderCntrol), New FrameworkPropertyMetadata(GetType(FolderCntrol)))
    End Sub
    Public Overrides Sub OnApplyTemplate()
        myResourceDictionary = New ResourceDictionary()
        myResourceDictionary.Source = New Uri("pack://application:,,,/TingTong Mail;component/Themes/Generic.xaml", UriKind.RelativeOrAbsolute)
        MyBase.OnApplyTemplate()
        bor = GetTemplateChild("Border")
        MsEnterRect = GetTemplateChild("MsEnterRect")
        indicator = GetTemplateChild("Indicator")
        Dim Gdd As Grid = GetTemplateChild("Gdd")
        AddHandler Gdd.MouseEnter, AddressOf FolderCntrol_MouseEnter
        AddHandler Gdd.MouseLeave, AddressOf FolderCntrol_MouseLeave


        HdrCon = GetTemplateChild("Header")
        Dim HeaderContentBinding As Binding = New Binding("HeaderContent") With {
        .Source = Me,
          .Mode = BindingMode.TwoWay
          }
        HdrCon.SetBinding(TextBlock.TextProperty, HeaderContentBinding)
        stp = GetTemplateChild("stp")
        ExpBtn = GetTemplateChild("ExpBtn")
        AddHandler ExpBtn.MouseDown, AddressOf ExpBtn_MouseDown
    End Sub
    Private Sub ExpBtn_MouseDown()
        If IsExpanded Then
            IsExpanded = False
            bor.Height = 30
            Dim sb As New Storyboard
            sb = CType(myResourceDictionary("Collapse"), Storyboard)
            Storyboard.SetTarget(sb, indicator)
            Dim sb1 As New Storyboard
            sb1 = CType(myResourceDictionary("Collapse1"), Storyboard)
            Storyboard.SetTarget(sb1, stp)
            AddHandler sb1.Completed, AddressOf ClearMsgFolders
            sb1.Begin()
            sb.Begin()
        Else
            IsExpanded = True
            bor.Height = 169
            LoadFolders()
            Dim sb As New Storyboard
            sb = FindResource("Expand")
            Storyboard.SetTarget(sb, indicator)
            Dim sb1 As New Storyboard
            sb1 = CType(myResourceDictionary("Expand1"), Storyboard)
            Storyboard.SetTarget(sb1, stp)
            sb1.Begin()
            sb.Begin()
        End If
    End Sub
    Private Sub ClearMsgFolders()
        stp.Children.Clear()
        stp.Height = 1
    End Sub
    Private Sub FolderCntrol_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        myResourceDictionary.Source = New Uri("pack://application:,,,/TingTong Mail;component/Themes/Generic.xaml", UriKind.RelativeOrAbsolute)
        If IsExpanded Then
            bor.Height = 169
            indicator.RenderTransform = New RotateTransform(-90)
            LoadFolders()
        Else
            indicator.RenderTransform = New RotateTransform(90)
            bor.Height = 30
        End If
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
    Private Sub LoadFolders()
        For Each Folder In FolderList
            Dim btn As New Button
            btn.Style = myResourceDictionary("EfolderBtns")
            btn.Content = Folder
            stp.Height += 26
            AddHandler btn.MouseDown, AddressOf Btn_Click
            stp.Children.Add(btn)
        Next
    End Sub
    Public Event ButtonClicked As EventHandler
    Private Sub Btn_Click(sender As Object, e As MouseButtonEventArgs)
        RaiseEvent ButtonClicked(sender, e)
    End Sub
End Class

Imports System.Windows.Media.Animation
Imports System.ComponentModel
Imports System.IO.Directory
Imports System.IO
Imports System.Diagnostics
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Text

Public Class MailContainer1
    Inherits System.Windows.Controls.Control
    Dim photoloader As New ImageLoader
    Public Main As System.Windows.Shapes.Rectangle
    Private SelectedRect As System.Windows.Shapes.Rectangle
    Public MainGrid As Grid
    Private Ic As Integer
    Public From As TextBlock
    Public EDate As TextBlock
    Public Subject As TextBlock
    Private FromToolTip As TextBlock
    Public IsSeen As Boolean
    Public AddedToFavourites As Boolean
    Private IsOpened As Boolean
    Private GetFromSender As Boolean
    Public HasAttachment As Boolean
    Public IsChecked As Boolean
    Private CheckBx As CustomCheckbox
    Public UniqueIdText As String
    Private bb As New BrushConverter
    Public Delegate Sub OnCh(a As String)
    Public Event Changed As OnCh
    Private WithEvents DownloadMail As New BackgroundWorker
    Private WithEvents DetectUnSelected As New Forms.Timer
    Public Shared ReadOnly FromDependency As DependencyProperty = DependencyProperty.Register("FromText", GetType(String), GetType(MailContainer1))
    Public Shared ReadOnly EDateDependency As DependencyProperty = DependencyProperty.Register("Date", GetType(String), GetType(MailContainer1))
    Public Shared ReadOnly SubjectDependency As DependencyProperty = DependencyProperty.Register("Subject", GetType(String), GetType(MailContainer1))
    Private ttm As Startup
    Dim CurrentMailFolder As String = AppDomain.CurrentDomain.BaseDirectory + "Data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\"
    Public Property SubjectText As String
        Get
            Return CStr(GetValue(SubjectDependency))
        End Get
        Set(ByVal value As String)
            SetValue(SubjectDependency, value)
            Using ()
        End Set
    End Property
    Public Property EdateText As String
        Get
            Return CStr(GetValue(EDateDependency))
        End Get
        Set(ByVal value As String)
            SetValue(EDateDependency, value)
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
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        MainGrid = GetTemplateChild("MainGrid")
        From = TryCast(GetTemplateChild("From"), TextBlock)
        CheckBx = GetTemplateChild("ChkBx")
        AddHandler CheckBx.MouseLeftButtonDown, AddressOf CheckBx_MouseLeftButtonDown
        Dim FromTextBinding As Binding = New Binding("FromText")
        FromTextBinding.Source = Me
        FromTextBinding.Mode = BindingMode.TwoWay
        From.SetBinding(TextBlock.TextProperty, FromTextBinding)
        AddHandler From.MouseLeftButtonDown, AddressOf From_MouseLeftButtonDown

        EDate = TryCast(GetTemplateChild("Date"), TextBlock)
        Dim EDateBinding As Binding = New Binding("Date")
        EDateBinding.Source = Me
        EDateBinding.Mode = BindingMode.TwoWay
        EDate.SetBinding(TextBlock.TextProperty, EDateBinding)

        Subject = TryCast(GetTemplateChild("Subject"), TextBlock)
        Dim SubjectBinding As Binding = New Binding("Subject")
        SubjectBinding.Source = Me
        SubjectBinding.Mode = BindingMode.TwoWay
        Subject.SetBinding(TextBlock.TextProperty, SubjectBinding)

        SelectedRect = GetTemplateChild("SelectedRect")

        Main = GetTemplateChild("Main")
        AddHandler Main.MouseLeftButtonDown, AddressOf Main_MouseLeftButtonDown

        FromToolTip = GetTemplateChild("FromToolTip")

        FromToolTip.Text = "Show all mails of " + From.Text
    End Sub
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(MailContainer1), New FrameworkPropertyMetadata(GetType(MailContainer1)))
    End Sub
    Private Sub DeleteMail_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        My.Settings.DeleteMailId = UniqueIdText
        CType(Me.Parent, VirtualizingStackPanel).Children.Remove(Me)
        CType(Me.Parent, VirtualizingStackPanel).Height = CType(Me.Parent, VirtualizingStackPanel).Height - 45
    End Sub
    Private Sub CheckBx_MouseLeftButtonDown()
        IsChecked = CheckBx.Checked
    End Sub
    Private Sub From_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        My.Settings.MailDownloadDone = False
        My.Settings.OpenedMailId = UniqueIdText
        GetFromSender = True
        If GetFromSender Then
            RaiseEvent TriggerGetFromSender(Me, e)
        Else
            RaiseEvent GetSingleMail(Me, e)
        End If
        IsOpened = True
        SelectedRect.Visibility = Visibility.Visible
        DetectUnSelected.Start()
    End Sub
    Private Sub Main_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        My.Settings.MailDownloadDone = False
        My.Settings.OpenedMailId = UniqueIdText
        GetFromSender = False
        If GetFromSender Then
            RaiseEvent TriggerGetFromSender(Me, e)
        Else
            RaiseEvent GetSingleMail(Me, e)
        End If
        IsOpened = True
        SelectedRect.Visibility = Visibility.Visible
        DetectUnSelected.Start()
    End Sub
    Private Sub MailContainer1_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If HasAttachment Then
            AttachmentLoad()
        End If
    End Sub
    Private Sub AttachmentLoad()
        Dim AttachMentBtn As New Image With {
            .Source = photoloader.LoadImage(New Uri("pack://application:,,,/TingTong Mail;component/icons/attachico.png"), 20),
            .Height = 18,
            .Width = 18,
            .VerticalAlignment = VerticalAlignment.Top
        }
        MainGrid.Children.Add(AttachMentBtn)
        Grid.SetColumn(AttachMentBtn, 3)
        AttachMentBtn.Margin = New Thickness(12, 16, 0, 16)
    End Sub
    Private Sub DetectUnSelected_Tick(sender As Object, e As EventArgs) Handles DetectUnSelected.Tick
        If Not My.Settings.OpenedMailId = UniqueIdText Then
            DetectUnSelected.Stop()
            SelectedRect.Visibility = Visibility.Hidden
        End If
    End Sub
    Public Event TriggerGetFromSender As EventHandler
    Public Event GetSingleMail As EventHandler
    Public WithEvents btn As New Button

    Private Sub btn_Click(sender As Object, e As RoutedEventArgs) Handles btn.Click

    End Sub
End Class
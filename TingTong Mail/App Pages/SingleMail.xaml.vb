Imports System.Text.RegularExpressions
Imports System.Data.OleDb
Imports System.Data
Imports System.IO.Directory
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media.Animation
Imports mshtml
Imports WpfAnimatedGif

Class SingleMail
    Dim documentEvents As HTMLDocumentEvents2_Event
    Public _wbhandler As WebBrowserHostUIHandler
    Dim GetGroup As New Groups
    Dim GroupName As String
    Public CurrentlyNavigted As String
    Dim MainParent As Frame
    Public Expanded As Boolean
    Dim PopOutWindow As New EmailWindow
    Public Shared Function FindParent(Of T As DependencyObject)(ByVal child As DependencyObject) As T
        Dim parentObject As DependencyObject = VisualTreeHelper.GetParent(child)
        If parentObject Is Nothing Then Return Nothing
        Dim parent As T = TryCast(parentObject, T)
        If parent IsNot Nothing Then Return parent Else Return FindParent(Of T)(parentObject)
    End Function
    Private Sub SingleMail_LostFocus(sender As Object, e As RoutedEventArgs) Handles Me.LostFocus

    End Sub
    Private Sub SingleMail_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        MainParent = FindParent(Of Frame)(Me)
        GetGroup.Initialize()
        LoadToolTips()
    End Sub
    Private Sub Browser1_LoadCompleted(sender As Object, e As NavigationEventArgs) Handles Browser1.LoadCompleted
        documentEvents = CType(Browser1.Document, HTMLDocumentEvents2_Event)
        AddHandler documentEvents.oncontextmenu, AddressOf webBrowser_ContextMenuOpening
    End Sub
    Private Function webBrowser_ContextMenuOpening(ByVal pEvtObj As IHTMLEventObj) As Boolean
        Return False
    End Function
    Private Sub CloseBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles CloseBtn.MouseDown
        Browser1.Navigate("about:blank")
        Browser1.Visibility = Visibility.Hidden
        Dim sb As New Storyboard
        sb = FindResource("HideMainframeForSingleMail")
        Storyboard.SetTarget(sb, MainParent)
        AddHandler sb.Completed, AddressOf AnimationDone
        sb.Begin()
    End Sub
    Private Sub AnimationDone()
        MainParent.Visibility = Visibility.Hidden
        MainParent.Opacity = 1
    End Sub
    Private Sub ExpandBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles expandBtn.MouseDown
        CurrentlyNavigted = Browser1.Source.ToString()
        Browser1.Navigate("about:blank")
        Browser1.Visibility = Visibility.Hidden
        Dim sb As New Storyboard
        If Expanded = False Then
            sb = Me.FindResource("ExpandFrame")
            Expanded = True
            expandBtn.ToolTipText = "Collapse"
        Else
            sb = Me.FindResource("CollapseFrame")
            Expanded = False
            expandBtn.ToolTipText = "Expand"
        End If
        Storyboard.SetTarget(sb, MainParent)
        AddHandler sb.Completed, AddressOf ReloadPage
        sb.Begin()
    End Sub
    Private Sub ReloadPage()
        Browser1.Navigate(CurrentlyNavigted)
        Browser1.Visibility = Visibility.Visible
        Dim bi As New BitmapImage
        bi.BeginInit()
        If Expanded = False Then
            bi.UriSource = New Uri("pack://application:,,,/TingTong Mail;component/icons/expand.png")
        Else
            bi.UriSource = New Uri("pack://application:,,,/TingTong Mail;component/icons/collapse.png")
        End If
        bi.EndInit()
        bi.Freeze()
        expandBtn.ImgSource = bi
    End Sub
    Private Sub LoadToolTips()
        CloseBtn.ToolTipText = "Close"
        expandBtn.ToolTipText = "Expand"
        replyBtn.ToolTipText = "Reply to the sender"
        replytoallBtn.ToolTipText = "Reply to all sender(s)"
        forwardBtn.ToolTipText = "Forward this E-mail"
        PopOutBtn.ToolTipText = "Open in separate window"
    End Sub

    Private Sub PopOutBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles PopOutBtn.MouseDown
        Browser1.Navigate("about:blank")
        Browser1.Visibility = Visibility.Hidden
        Dim sb As New Storyboard
        sb = FindResource("HideMainframeForSingleMail")
        Storyboard.SetTarget(sb, MainParent)
        AddHandler sb.Completed, AddressOf AnimationDone
        MainParent.NavigationService.RemoveBackEntry()
        MainParent.Content = ""
        CurrentlyNavigted = Browser1.Source.ToString
        With PopOutWindow
            .From.Text = EDetails.From.Text
            .To.Text = EDetails.ETo.Text
            .Cc.Text = EDetails.Cc.Text
            .Subject.Text = EDetails.Subject.Text
            .Browser1.Navigate(CurrentlyNavigted)
            .Show()
        End With
        MainParent.Visibility = Visibility.Hidden
        MainParent.Opacity = 1
    End Sub
    Private Sub replyBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles replyBtn.MouseDown
        Dim replyWindow As New sendTheEmail
        Dim sendingMailWindow As New SendMail
        replyWindow.Show()
        sendingMailWindow.To.Text = EDetails.From.Text
        replyWindow.ShowPage.Navigate(sendingMailWindow)
    End Sub
End Class

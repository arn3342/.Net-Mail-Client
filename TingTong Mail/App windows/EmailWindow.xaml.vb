Imports System.Data.OleDb
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media.Animation
Imports System.IO.Directory
Imports mshtml
Imports System.Text.RegularExpressions
Imports System.Data
Imports WpfAnimatedGif

Public Class EmailWindow
    Dim PhotoLoader As New ImageLoader
    Dim Image1 As New BitmapImage
    Dim Image2 As New BitmapImage
    Dim bb As New BrushConverter
    Dim documentEvents As HTMLDocumentEvents2_Event
    Public _wbhandler As WebBrowserHostUIHandler
    Dim saveCon As New SaveNewContact
    Dim WithEvents UndoSaveText As New Forms.Timer
    Dim WithEvents LoadGroups As New Forms.Timer
    Dim i As Integer
    Dim GetGroup As New Groups
    Dim GroupName As String
    Dim Expanded As Boolean
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetCurrentDirectory() & "\Data\" + My.Settings.email + "\_Cons.ttmdata;Jet OLEDB:Database Password=arn33423342;")
    Private Sub LoadLogos()
        Try
            Dim path = New Uri("pack://application:,,,/TingTong Mail;component/icons/replytosender.data")
            Dim Image As BitmapImage = PhotoLoader.LoadImage(path, 30)
            ImageBehavior.SetAnimatedSource(replyimg, Image)
        Catch ex As Exception
        End Try
        Try
            Dim path = New Uri("pack://application:,,,/TingTong Mail;component/icons/replytoall.data")
            Dim Image As BitmapImage = PhotoLoader.LoadImage(path, 30)
            ImageBehavior.SetAnimatedSource(replytoallimg, Image)
        Catch ex As Exception
        End Try
        Try
            Dim path = New Uri("pack://application:,,,/TingTong Mail;component/icons/down.png")
            Dim Image As BitmapImage = PhotoLoader.LoadImage(path, 30)
            Options.Source = Image
        Catch ex As Exception
        End Try
        Try
            Dim path = New Uri("pack://application:,,,/TingTong Mail;component/icons/forward.data")
            Dim Image As BitmapImage = PhotoLoader.LoadImage(path, 30)
            ImageBehavior.SetAnimatedSource(forwardImg, Image)
        Catch ex As Exception
        End Try
        Try
            Dim path = New Uri("pack://application:,,,/TingTong Mail;component/icons/usericon.png")
            Dim Image As BitmapImage = PhotoLoader.LoadImage(path, 30)
            UserPic.ImageSource = Image
        Catch ex As Exception
        End Try
        Try
            Dim path = New Uri("pack://application:,,,/TingTong Mail;component/icons/downloadBack.png")
            Image1 = PhotoLoader.LoadImage(path, 30)
        Catch ex As Exception
        End Try
        Try
            Dim path = New Uri("pack://application:,,,/TingTong Mail;component/icons/downloaddBlue.png")
            Image2 = PhotoLoader.LoadImage(path, 30)
        Catch ex As Exception
        End Try
        DownloadAtaachmentsBtn.Source = Image1
    End Sub
    Private Sub DownloadAtaachmentsBtn_MouseLeave(sender As Object, e As MouseEventArgs) Handles DownloadAtaachmentsBtn.MouseLeave
        DownloadAtaachmentsBtn.Source = Image1
    End Sub
    Private Sub CloseAttachmentDialog_MouseEnter(sender As Object, e As MouseEventArgs) Handles CloseAttachmentDialog.MouseEnter
        CloseAttachmentDialog.Opacity = 1
    End Sub

    Private Sub CloseAttachmentDialog_MouseLeave(sender As Object, e As MouseEventArgs) Handles CloseAttachmentDialog.MouseLeave
        CloseAttachmentDialog.Opacity = 0.5
    End Sub
    Private Sub CloseAttachmentDialog_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles CloseAttachmentDialog.MouseDown
        AttachmentMsg.Visibility = Visibility.Hidden
        TopGrid.Height = TopGrid.Height - 21
        Browser1.VerticalAlignment = VerticalAlignment.Stretch
    End Sub

    Private Sub DownloadAtaachmentsBtn_MouseEnter(sender As Object, e As MouseEventArgs) Handles DownloadAtaachmentsBtn.MouseEnter
        DownloadAtaachmentsBtn.Source = Image2
    End Sub
    Private Sub EmailWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        LoadLogos()
        con.Open()
        GetGroup.Initialize()
    End Sub
    Private Sub Browser1_LoadCompleted(sender As Object, e As NavigationEventArgs) Handles Browser1.LoadCompleted
        documentEvents = CType(Browser1.Document, HTMLDocumentEvents2_Event)
        AddHandler documentEvents.oncontextmenu, AddressOf webBrowser_ContextMenuOpening
    End Sub
    Private Function webBrowser_ContextMenuOpening(ByVal pEvtObj As IHTMLEventObj) As Boolean
        Return False
    End Function
    Private Sub Options_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles Options.MouseDown
        Dim cm = ContextMenuService.GetContextMenu(TryCast(sender, DependencyObject))
        If cm Is Nothing Then
            Return
        End If
        cm.Placement = PlacementMode.Bottom
        cm.PlacementTarget = TryCast(sender, UIElement)
        cm.IsOpen = True
        Dim template = ContactOptions.Template
        Dim SenderAddress As Match = Regex.Match(From.Text, "<(.*?)>")
        Dim email = CType(template.FindName("email", ContactOptions), TextBlock)
        If SenderAddress.Success Then
            email.Text = SenderAddress.Groups(1).Value.ToString
        End If
        GetCurrentUserContact()
    End Sub
    Private Sub AddName_KeyDown(sender As TextBox, e As KeyEventArgs)
        If e.Key = Key.Return Then
            If Not sender.Text = "" Then
                sender.Visibility = Visibility.Hidden
                Dim template = ContactOptions.Template
                Dim UserName = CType(template.FindName("name", ContactOptions), TextBlock)
                UserName.Text = sender.Text
            End If
        End If
    End Sub

    Private Sub Address_GotFocus(sender As TextBox, e As RoutedEventArgs)
        With sender
            .Foreground = Brushes.Black
            If .Text = " Address" Then
                .Text = ""
            End If
        End With
    End Sub
    Private Sub Address_LostFocus(sender As TextBox, e As RoutedEventArgs)
        With sender
            .Foreground = bb.ConvertFrom("#FFB0B0B0")
            If .Text = "" Then
                .Text = " Address"
            End If
        End With
    End Sub

    Private Sub SecEmail_GotFocus(sender As TextBox, e As RoutedEventArgs)
        With sender
            .Foreground = Brushes.Black
            If .Text = " Secondary email" Then
                .Text = ""
            End If
        End With
    End Sub

    Private Sub SecEmail_LostFocus(sender As Object, e As RoutedEventArgs)
        With sender
            .Foreground = bb.ConvertFrom("#FFB0B0B0")
            If .Text = "" Then
                .Text = " Secondary email"
            End If
        End With
    End Sub
    Private Sub UserName_MouseDown(sender As Object, e As MouseButtonEventArgs)
        If e.ClickCount = 2 Then
            Dim template = ContactOptions.Template
            Dim AddName = CType(template.FindName("AddName", ContactOptions), TextBox)
            AddName.Visibility = Visibility.Visible
            FocusManager.SetFocusedElement(Me, AddName)
        End If
    End Sub
    Private Sub SaveContact()
        Dim template = ContactOptions.Template
        Dim n = CType(template.FindName("name", ContactOptions), TextBlock).Text
        Dim Name As String = n
        If n = "Double click to add a name" Then
            Name = ""
        End If
        Dim e = CType(template.FindName("email", ContactOptions), TextBlock).Text
        Dim Email As String = e
        Dim s = CType(template.FindName("SecEmail", ContactOptions), TextBox).Text
        Dim SecMail As String = s
        If s = " Secondary email" Then
            SecMail = ""
        End If
        Dim a = CType(template.FindName("Address", ContactOptions), TextBox).Text
        Dim Address As String = a
        If a = " Address" Then
            Address = ""
        End If
        saveCon.SaveContact(Name, Email, SecMail, Address, "", My.Settings.ContactGroupName)
        Dim saveimg = CType(template.FindName("saveImg", ContactOptions), Image)
        Dim scale As ScaleTransform = saveimg.RenderTransform
        scale.ScaleX = 1
        scale.ScaleY = 1
        saveimg.Visibility = Visibility.Visible
        Dim sb As New Storyboard
        sb = FindResource("Saved")
        Storyboard.SetTarget(sb, saveimg)
        sb.Begin()
        UndoSaveText.Interval = 300
        UndoSaveText.Start()
    End Sub
    Private Sub GetCurrentUserContact()
        Dim template = ContactOptions.Template
        Dim email = CType(template.FindName("email", ContactOptions), TextBlock)
        Dim SenderAddress As Match = Regex.Match(From.Text, "<(.*?)>")
        Dim emailFrom = SenderAddress.Groups(1).Value.ToString
        Dim cmd As New OleDbCommand("Select * from contacts WHERE Email='" + emailFrom + "'", con)
        Dim table As New DataTable
        Dim adapter As New OleDbDataAdapter(cmd)
        adapter.Fill(table)
        Dim AddName = CType(template.FindName("AddName", ContactOptions), TextBox)
        Dim name = CType(template.FindName("name", ContactOptions), TextBlock)
        Dim secmail = CType(template.FindName("SecEmail", ContactOptions), TextBox)
        Dim address = CType(template.FindName("Address", ContactOptions), TextBox)
        If Not table.Rows.Count <= 0 Then
            name.Text = table.Rows(0)(0).ToString : AddName.Text = table.Rows(0)(0).ToString
            email.Text = table.Rows(0)(1).ToString
            secmail.Text = table.Rows(0)(2).ToString
            address.Text = table.Rows(0)(3).ToString
        Else
            name.Text = "Double click to add a name" : AddName.Text = ""
            email.Text = emailFrom
            secmail.Text = " Secondary email"
            address.Text = " Address"
        End If
    End Sub
End Class

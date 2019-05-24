Imports System.ComponentModel
Imports System.Windows.Media.Animation
Imports System.IO.Directory
Imports System.Windows.Controls.Primitives
Imports System.Data.OleDb
Imports System.IO
Imports MimeKit
Imports MailKit
Imports MailKit.Search
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Configuration
Imports System.Reflection
Imports System.Text

Public Class Home
    Dim conStr = AppDomain.CurrentDomain.BaseDirectory()
    Public Shared contact As OleDbConnection

    Dim err As String
    Dim sb As New Storyboard
    Dim sb2 As New Storyboard
    Dim bb As New BrushConverter
    Dim emaillist As New BindingList(Of UserControl)
    Dim EmailFolders As New List(Of String)
    Dim WithEvents LoadMessageFolders As New BackgroundWorker
    Dim Conpage As New ContactPage
    'Dim my.settings.Email As String = File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory() + "data\ctuser.data")


    Dim btn As New MailContainer1
    Dim ttm As Startup
    Public WithEvents MailkitInboxLoad As New BackgroundWorker
    Public WithEvents MailkitLogIn As New BackgroundWorker
    Public WithEvents LoadInboxExistingMails As New BackgroundWorker
    Dim WithEvents AllEmailOf As New BackgroundWorker
    Dim TotalEmails As String
    Dim FetchType As String
    Dim EmailSender As String
    Dim InboxMails As New List(Of Emails)
    Dim NewMailsInRange As New List(Of Emails)
    Dim OldMailsInRange As New List(Of Emails)
    Dim AllEmailFromSender As New List(Of Emails)
    Dim TotalDownloadedMails As Integer
    Dim TotalCurrentMails, TotalSeenMails, TotalUnSeenMails, TotalFlaggedMails, TotalFavouritesMails As Integer
    Dim emailRange As Integer = 0
    Dim LoadingType As String
    Dim LastMail As Integer = 0
    Dim FetchTill As Integer
    Dim SingleEmaillViewer As New SingleMail
    Dim AllEmailOfViewer As New EmailReader
    Dim LoadMoreMails As New LoadMoreEmails
    Dim WithEvents ViewSingleMail As New Forms.Timer
    Dim WithEvents HideSingleMail As New Forms.Timer
    Dim WithEvents ViewSingleMailOfSpecificUser As New Forms.Timer
    Dim WithEvents HideAllOfSpecificUser As New Forms.Timer
    Dim WithEvents StartAllMailProcess As New Forms.Timer
    Dim WithEvents HideSearchBox As New Forms.Timer
    Dim WithEvents SearchForMailInMailbox As New Forms.Timer
    Dim photoloader As New ImageLoader
    Public IsSingleMailShows As Boolean
    Dim IsListReversed As Boolean
    Dim IsFound As Boolean
    Dim MailExist As Boolean
    Private Shared resetEvent As AutoResetEvent = New AutoResetEvent(False)
    Private Shared WaitTillComplete As AutoResetEvent = New AutoResetEvent(False)
    Dim currentImapFolder As MailKit.IMailFolder = Nothing
    Private Sub Home_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        AllEmailsScrollViewer.HorizontalAlignment = HorizontalAlignment.Stretch
        AllEmailsScrollViewer.Width = Double.NaN
        MessagePnl.HeaderText = "Inbox"
        My.Settings.shutdowncommand = 0
        MessageFoldersList.IsExpanded = True
        My.Settings.CurrentMailFolderName = "Inbox"
        Dim allFolders As String = System.AppDomain.CurrentDomain.BaseDirectory() + "Data\" + My.Settings.email + "\EmailFolders\"
        Dim mailFolders = GetDirectories(allFolders)
        If mailFolders.Length <= 0 Then
            MailkitLogIn.RunWorkerAsync()
        Else
            If Not GetFiles(System.AppDomain.CurrentDomain.BaseDirectory() + "Data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\", "*.*", SearchOption.AllDirectories).Count <= 0 Then
                MailExist = True
                LoadMailFromFolder(My.Settings.CurrentMailFolderName)
            Else
                MailkitInboxLoad.RunWorkerAsync()
            End If
            For Each folderName As String In GetDirectories(allFolders)
                MessageFoldersList.FolderList.Add("    " + folderName.Remove(0, allFolders.Length))
            Next
            For Each Item As Button In MessageFoldersList.stp.Children
                AddHandler MessageFoldersList.ButtonClicked, AddressOf btn_click
            Next
        End If
        AddHandler ContactButton.MouseLeftButtonDown, AddressOf ShowContactsPage
    End Sub
    Private Sub ConnectToContacts()
        Dim currentDomain = AppDomain.CurrentDomain
        currentDomain.SetData("DataDirectory", currentDomain.BaseDirectory() + "Data")
        contact = New OleDbConnection(ConfigurationManager.ConnectionStrings("contactDb").ConnectionString)
        contact.Open()
    End Sub
    Private Sub LoadMailFromFolder(foldername As String)
        MailkitInboxLoad.RunWorkerAsync()
        resetEvent.WaitOne()
    End Sub
    Private Sub Home_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If e Is Nothing Then
            Throw New ArgumentNullException(NameOf(e))
        End If
        My.Settings.OpenMode = "Add"
        Application.Current.Shutdown()
    End Sub
    Class EmailFolder
        Public Property Title As String
        Public Overrides Function ToString() As String
            Return Me.Title
        End Function
    End Class
    Private Sub btn_click(sender As Object, e As RoutedEventArgs)
        TotalSeenMails = 0
        TotalUnSeenMails = 0
        TotalFlaggedMails = 0
        TotalFavouritesMails = 0
        Dim btnText = DirectCast(sender, Button).Content.Replace("    ", "")
        My.Settings.CurrentMailFolderName = btnText
        Teststack.Height = 1
        Teststack.Children.Clear()
        InboxMails.Clear()
        NewMailsInRange.Clear()
        OldMailsInRange.Clear()
        AllEmailFromSender.Clear()
        MessageFoldersList.HdrCon.Opacity = 0
        MessageFoldersList.HeaderText = btnText
        Dim sb As New Storyboard
        sb = FindResource("MailFolderChanged")
        Storyboard.SetTarget(sb, MessageFoldersList.HdrCon)
        sb.Begin()
        Dim fi = GetFiles(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\", "*.mail", SearchOption.AllDirectories)
        If Not fi.Count <= 0 Then
            LastMail = Path.GetFileNameWithoutExtension(fi(fi.Count - 1))
            FetchTill = LastMail + 20
        End If
        If Not GetFiles(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\", "*.mail", SearchOption.AllDirectories).Count <= 0 Then
            LoadMailFromFolder(My.Settings.CurrentMailFolderName)
        Else
            MailkitInboxLoad.RunWorkerAsync()
        End If
    End Sub
    Private Sub LoadMessageInRange_Working()
        If currentImapFolder Is Nothing Then
            If My.Settings.CurrentMailFolderName = "Drafts" Then
                currentImapFolder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Drafts)
            ElseIf My.Settings.CurrentMailFolderName.Contains("Sen") Then
                currentImapFolder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Sent)
            ElseIf My.Settings.CurrentMailFolderName = "Inbox" Then
                currentImapFolder = ttm.MailKitClient.GetFolder("Inbox")
            End If
            If currentImapFolder Is Nothing Then
                currentImapFolder = ttm.MailKitClient.GetFolder("Inbox")
            End If
            currentImapFolder.Open(FolderAccess.ReadWrite)
        End If
        Dim range = New UniqueIdRange(New UniqueId(LastMail), New UniqueId(FetchTill))
        Dim uid = currentImapFolder.Search(range, SearchQuery.All)
        For Each summary In currentImapFolder.Fetch(uid, MessageSummaryItems.Full)
            Dim Email As New Emails
            Dim From = If((summary.Envelope.From.ToString <> ""), summary.Envelope.From.ToString, "No sender")
            Dim match As Match = Regex.Match(From, """([^""]*)""")
            If match.Success Then
                Dim EmailFrom As String = match.Groups(1).Value
                Email.From = EmailFrom
            Else
                Email.From = summary.Envelope.From.ToString
            End If
            If Not summary.Attachments.Count <= 0 Then
                Email.HasAttachment = True
            End If
            Email.UniqueId = summary.UniqueId.ToString
            Email.EmailDate = summary.Envelope.Date.ToString
            If summary.Flags.Value.HasFlag(MessageFlags.Seen) Then
                Email.IsSeen = "Seen"
            Else
                Email.IsSeen = "Unseen"
            End If
            Try
                Email.Subject = summary.Envelope.Subject.ToString
            Catch ex As Exception
            End Try
            NewMailsInRange.Add(Email)
        Next
        My.Settings.Save()
        resetEvent.Set()
        FetchType = "Older mails"
    End Sub
    Private Sub LoadFolderMails_Working()
        If currentImapFolder Is Nothing Then
            If My.Settings.CurrentMailFolderName = "Drafts" Then
                currentImapFolder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Drafts)
            ElseIf My.Settings.CurrentMailFolderName.Contains("Sent") Then
                currentImapFolder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Sent)
            ElseIf My.Settings.CurrentMailFolderName = "Inbox" Then
                currentImapFolder = ttm.MailKitClient.GetFolder("Inbox")
            ElseIf currentImapFolder Is Nothing Then
                currentImapFolder = ttm.MailKitClient.GetFolder("Inbox")
            End If
            currentImapFolder.Open(FolderAccess.ReadWrite)
        End If
        Dim CurrentMailFolder As String = System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\"
        Dim inFolder = GetFiles(CurrentMailFolder, "*.mail", SearchOption.AllDirectories)
        TotalEmails = currentImapFolder.Count.ToString
        If currentImapFolder.Count = 0 Then
        Else
            If inFolder.Count <= 0 Then
                ''''If there are no existing emails in INBOX
                Dim lst As New List(Of String)
                Dim range As New UniqueIdRange(New UniqueId(currentImapFolder.Count), New UniqueId(currentImapFolder.Count + 9999))
                Dim uid = currentImapFolder.Search(range, SearchQuery.All)
                For Each summary In currentImapFolder.Fetch(uid, MessageSummaryItems.Full Or MessageSummaryItems.UniqueId)
                    Dim Email As New Emails
                    Dim From = If((summary.Envelope.From.ToString <> ""), summary.Envelope.From.ToString, "No sender")

                    Dim match As Match = Regex.Match(From, """([^""]*)""")
                    If match.Success Then
                        Dim EmailFrom As String = match.Groups(1).Value
                        Email.From = EmailFrom
                    Else
                        Email.From = summary.Envelope.From.ToString
                    End If
                    Email.UniqueId = summary.UniqueId.ToString
                    Email.EmailDate = summary.Envelope.Date.ToString
                    If summary.Flags.Value.HasFlag(MessageFlags.Seen) Then
                        Email.IsSeen = "Seen"
                    Else
                        Email.IsSeen = "Unseen"
                    End If
                    If Not summary.Attachments.Count <= 0 Then
                        Email.HasAttachment = True
                    End If
                    Email.Cc = summary.Envelope.Cc.ToString
                    Email.Subject = summary.Envelope.Subject.ToString
                    InboxMails.Add(Email)
                    Dim lss As String = summary.UniqueId.ToString
                    Dim topmailid As Integer = Convert.ToInt32(summary.UniqueId.ToString)
                Next
                My.Settings.Save()
                FetchType = "New"

            Else
                'If there are existing emails In inbox
                LoadMessageInRange_Working()
            End If
        End If

    End Sub
    Private Sub MailkitInboxLoad_DoWorkAsync(sender As Object, e As DoWorkEventArgs) Handles MailkitInboxLoad.DoWork
        Dim fi = GetFiles(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\", "*.mail", SearchOption.AllDirectories)
        If fi.Count <= 0 Then
        Else

            If LoadingType = "" Then
                LoadingType = "Newer"
                LastMail = Path.GetFileNameWithoutExtension(fi(fi.Count - 1))
                FetchTill = LastMail + 20
            End If
        End If
        LoadFolderMails_Working()
        resetEvent.Set()
    End Sub
    Private Sub MailkitInboxLoad_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles MailkitInboxLoad.RunWorkerCompleted
        If Not MailExist Then
            InboxMails.Reverse()
        End If
        NewMailsInRange.Reverse()
        If NewMailsInRange.Count <= 0 Then
            For Each email As Emails In InboxMails
                If Not File.Exists(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + email.UniqueId.ToString + ".mail") Then
                    File.WriteAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + email.UniqueId.ToString + ".mail", New String() {email.From, email.Subject, email.EmailDate, "No", email.IsSeen, email.HasAttachment.ToString, email.Cc})
                End If
                Dim btn As New MailContainer1 With {
                .UniqueIdText = email.UniqueId.ToString,
                .FromText = email.From,
                .EdateText = email.EmailDate,
                .SubjectText = email.Subject,
                .AddedToFavourites = False
                }
                If email.HasAttachment Then
                    btn.HasAttachment = True
                End If
                If email.IsSeen = "Seen" Then
                    btn.IsSeen = True
                    TotalSeenMails = TotalSeenMails + 1
                Else
                    btn.IsSeen = False
                    TotalUnSeenMails = TotalUnSeenMails + 1
                End If
                AddHandler btn.TriggerGetFromSender, AddressOf GetAllMailsFromSender
                AddHandler btn.GetSingleMail, AddressOf SingleMail
                'AddHandler btn.DeleteMail.MouseDown, AddressOf DeleteSelectedMail
                'AddHandler btn.SeenUnseenbtn.MouseDown, AddressOf MarkasSeenOrUnseen
                'AddHandler btn.favourites.MouseDown, AddressOf AddToFavourites
                Teststack.Height = Teststack.Height + 50
                Teststack.Children.Add(btn)
            Next
            'MsgBox(InboxMails.Count)
        Else
            For Each email As Emails In NewMailsInRange.GetRange(0, NewMailsInRange.Count)
                If Not File.Exists(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + email.UniqueId.ToString + ".mail") Then
                    File.WriteAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + email.UniqueId.ToString + ".mail", New String() {email.From, email.Subject, email.EmailDate, "No", email.IsSeen, email.HasAttachment.ToString, email.Cc})
                End If
                Dim btn As New MailContainer1 With {
                .UniqueIdText = email.UniqueId.ToString,
                .FromText = email.From,
                .EdateText = email.EmailDate,
                .SubjectText = email.Subject,
                .AddedToFavourites = False
                }
                If email.IsSeen = "Seen" Then
                    btn.IsSeen = True
                    TotalSeenMails = TotalSeenMails + 1
                Else
                    btn.IsSeen = False
                    TotalUnSeenMails = TotalUnSeenMails + 1
                End If
                AddHandler btn.TriggerGetFromSender, AddressOf GetAllMailsFromSender
                AddHandler btn.GetSingleMail, AddressOf SingleMail
                'AddHandler btn.DeleteMail.MouseDown, AddressOf DeleteSelectedMail
                'AddHandler btn.SeenUnseenBtn.MouseDown, AddressOf MarkasSeenOrUnseen
                'AddHandler btn.favourites.MouseDown, AddressOf AddToFavourites
                Teststack.Height = Teststack.Height + 45
                If LoadingType = "Newer" Then
                    Teststack.Children.Insert(0, btn)
                Else
                    Teststack.Children.Add(btn)
                End If
            Next
        End If
        AddHandler LoadMoreMails.LoadOlderMails.MouseDown, AddressOf LoadOlderMails_MouseDown
        If Not MailExist Then
            Teststack.Children.Add(LoadMoreMails)
        End If
        'CurrentMailFolder.Text = My.Settings.CurrentMailFolderName + "(" + TotalEmails + " mails)"
        emailRange = emailRange + 25
        If LoadingType = "Newer" Then
            Try
                Teststack.Children.RemoveAt(0)
                Teststack.Height = Teststack.Height - 45
            Catch ex As Exception
            End Try
        End If
        TotalCurrentMails = Teststack.Children.Count
        If MailExist Then
            LoadInboxExistingMails.RunWorkerAsync()
        End If
    End Sub
    Friend Sub MarkasSeenOrUnseen(sender As Object, e As EventArgs)
        If MailkitInboxLoad.IsBusy = True Then
            MsgBox("Please be patient.Emails are still loading.")
        Else
            Dim Folder As IMailFolder = Nothing '
            If My.Settings.CurrentMailFolderName = "Drafts" Then
                Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Drafts)
            ElseIf My.Settings.CurrentMailFolderName = "Send" Then
                Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Sent)
            ElseIf My.Settings.CurrentMailFolderName = "Inbox" Then
                Folder = ttm.MailKitClient.GetFolder("Inbox")
            ElseIf Folder Is Nothing Then
                Folder = ttm.MailKitClient.GetFolder("Inbox")
            End If
            If My.Settings.SeenUnseenId.ToString.Contains("-") Then
                Dim id As Integer = Convert.ToInt32(My.Settings.SeenUnseenId.ToString.Split("-")(1))
                Dim emailDetails As New List(Of String)(System.IO.File.ReadAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail"))
                File.Delete(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail")
                File.WriteAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail", New String() {emailDetails(0), emailDetails(1), emailDetails(2), "No", "UnSeen", emailDetails(5), emailDetails(6)})
                Folder.RemoveFlags(id, MessageFlags.Seen, True)
                TotalUnSeenMails = TotalUnSeenMails + 1
                TotalSeenMails = TotalSeenMails - 1
                Dim index As Integer = InboxMails.FindIndex(Function(a) a.UniqueId = id.ToString)
                InboxMails(index).IsSeen = "Unseen"
            Else
                'Folder.AddFlags(My.Settings.SeenUnseenId, MessageFlags.Seen, True)
                Dim emailDetails As New List(Of String)(System.IO.File.ReadAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + My.Settings.SeenUnseenId.ToString + ".mail"))
                File.Delete(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + My.Settings.SeenUnseenId.ToString + ".mail")
                File.WriteAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + My.Settings.SeenUnseenId.ToString + ".mail", New String() {emailDetails(0), emailDetails(1), emailDetails(2), "No", "Seen", emailDetails(5), emailDetails(6)})
                TotalSeenMails = TotalSeenMails + 1
                TotalUnSeenMails = TotalUnSeenMails - 1
                Dim index As Integer = InboxMails.FindIndex(Function(a) a.UniqueId = My.Settings.SeenUnseenId.ToString)
                InboxMails(index).IsSeen = "Seen"
            End If
        End If
    End Sub
    Private Sub LoadInboxExistingMails_DoWork(sender As Object, e As DoWorkEventArgs) Handles LoadInboxExistingMails.DoWork
        Dim FolderPath As String = System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\"
        For Each message As String In GetFiles(FolderPath, "*.mail", SearchOption.AllDirectories)
            Dim emailDetails As String() = (System.IO.File.ReadAllLines(message))
            Try
                Dim email As New Emails
                email.UniqueId = Path.GetFileNameWithoutExtension(message)
                email.From = emailDetails(0)
                email.Subject = emailDetails(1)
                email.EmailDate = emailDetails(2)
                email.Favourite = emailDetails(3)
                email.IsSeen = emailDetails(4)
                email.HasAttachment = Convert.ToBoolean(emailDetails(5))

                InboxMails.Add(email)
            Catch ex As Exception
                MsgBox(Path.GetFileNameWithoutExtension(message))
            End Try
        Next
    End Sub
    Private Sub LoadInboxExistingMails_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles LoadInboxExistingMails.RunWorkerCompleted
        InboxMails.Reverse()
        If InboxMails.Count > 27 Then
            emailRange = 25
        Else
            emailRange = 0
        End If
        Dim rangeTo As Integer
        If emailRange = 0 Then
            rangeTo = InboxMails.Count - 1
        Else
            rangeTo = InboxMails.Count - emailRange
        End If
        For Each email As Emails In InboxMails
            Dim btn As New MailContainer1
            btn.UniqueIdText = email.UniqueId.ToString
            btn.FromText = email.From
            btn.EdateText = email.EmailDate
            btn.SubjectText = email.Subject

            If email.Favourite = "Yes" Then
                btn.AddedToFavourites = True
            Else
                btn.AddedToFavourites = False
            End If
            If email.IsSeen = "Seen" Then
                btn.IsSeen = True
                TotalSeenMails = TotalSeenMails + 1
            Else
                btn.IsSeen = False
                TotalUnSeenMails = TotalUnSeenMails + 1
            End If
            btn.HasAttachment = email.HasAttachment
            AddHandler btn.TriggerGetFromSender, AddressOf GetAllMailsFromSender
            AddHandler btn.GetSingleMail, AddressOf SingleMail
            'AddHandler btn.DeleteMail.MouseDown, AddressOf DeleteSelectedMail
            'AddHandler btn.SeenUnseenbtn.MouseDown, AddressOf MarkasSeenOrUnseen
            'AddHandler btn.favourites.MouseDown, AddressOf AddToFavourites
            Teststack.Height = Teststack.Height + 45
            Teststack.Children.Add(btn)
        Next
        Teststack.Children.Add(LoadMoreMails)
    End Sub
    Friend Sub DeleteSelectedMail()
        Dim Folder As IMailFolder = ttm.MailKitClient.GetFolder(My.Settings.CurrentMailFolderName)
        If My.Settings.CurrentMailFolderName = "Drafts" Then
            Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Drafts)
        ElseIf My.Settings.CurrentMailFolderName = "Send" Then
            Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Sent)
        ElseIf My.Settings.CurrentMailFolderName = "Inbox" Then
            Folder = ttm.MailKitClient.GetFolder("Inbox")
        End If
        'Folder.AddFlags(My.Settings.DeleteMailId, MessageFlags.Deleted, True)
    End Sub
    Private Sub MailkitLogIn_DoWork(sender As Object, e As DoWorkEventArgs) Handles MailkitLogIn.DoWork
        Dim folders = ttm.MailKitClient.GetFolders(ttm.MailKitClient.PersonalNamespaces(0))
        For Each fol In folders
            EmailFolders.Add(fol.Name)
        Next
    End Sub

    Private Sub MailkitLogIn_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles MailkitLogIn.RunWorkerCompleted
        MessageFoldersList.stp.Children.Clear()
        MessageFoldersList.stp.Height = 1
        For Each Folder In EmailFolders
            Directory.CreateDirectory(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + Folder + "\")
            MessageFoldersList.FolderList.Add("    " + Folder)
        Next
        For Each Item As Button In MessageFoldersList.stp.Children
            AddHandler MessageFoldersList.ButtonClicked, AddressOf btn_click
        Next
        MailkitInboxLoad.RunWorkerAsync()
    End Sub

    Private Sub LoadOlderMails_MouseDown(sender As Object, e As MouseButtonEventArgs)
        NewMailsInRange.Clear()
        LoadingType = "Older"
        Dim fi = Directory.GetFiles(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\", "*.mail", SearchOption.AllDirectories)(0)
        LastMail = Convert.ToInt32(Path.GetFileNameWithoutExtension(fi)) - 1
        FetchTill = LastMail - 21
        If Not MailkitInboxLoad.IsBusy = True Then
            MailkitInboxLoad.RunWorkerAsync()
        End If
    End Sub
    Friend Sub SingleMail()
        With SingleEmaillViewer.EDetails
            .SubjectText = "Please wait"
            .FromText = "..."
            .EToText = "To : ..."
        End With
        SingleEmaillViewer.Expanded = False
        If Not My.Settings.OpenedMailId.ToString = "" Then
            Mainframe.Navigate(SingleEmaillViewer)
            Mainframe.Visibility = Visibility.Visible
            Mainframe.Opacity = 1
            LoadMoreMails.HorizontalAlignment = HorizontalAlignment.Left
            Dim sb As New Storyboard
            sb = FindResource("ViewMainframeForSingleMail")
            AddHandler sb.Completed, AddressOf ViewSingleMail_Done
            Storyboard.SetTarget(sb, Mainframe)
            sb.Begin()
        End If
        HideSingleMail.Start()
    End Sub
    Private Sub ViewSingleMail_Done()
        AllEmailsScrollViewer.HorizontalAlignment = HorizontalAlignment.Left
        AllEmailsScrollViewer.Width = 320
        Dim CurrentMailFolder As String = AppDomain.CurrentDomain.BaseDirectory + "Data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\"
        If File.Exists(CurrentMailFolder + My.Settings.OpenedMailId + "WI.MHT") Then
            SingleEmaillViewer.Browser1.Navigate(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + My.Settings.OpenedMailId + "WI.mht")
            Dim Message As String() = File.ReadAllLines(CurrentMailFolder + My.Settings.OpenedMailId + ".Mail")
            Dim match As Match = Regex.Match(Message(0), """([^""]*)""")
            With SingleEmaillViewer.EDetails
                .FromText = Message(0)
                SingleEmaillViewer.ContactPane.EmailFrom = Message(0)
                SingleEmaillViewer.ContactPane.From2.Text = match.Groups(1).Value
                .SubjectText = Message(1)
                .EToText = "To : " + My.Settings.email
                .CcText = Message(6).ToString
            End With
            Dim hasAttachment = Convert.ToBoolean(Message(5))
            If Not hasAttachment Then
                SingleEmaillViewer.EDetails.AttachVisibility = Visibility.Hidden
            Else
                SingleEmaillViewer.EDetails.AttachVisibility = Visibility.Visible
            End If
        Else
            If currentImapFolder Is Nothing Then
                If My.Settings.CurrentMailFolderName.Contains("Drafts") Then
                    currentImapFolder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Drafts)
                ElseIf My.Settings.CurrentMailFolderName.Contains("Sent") Then
                    currentImapFolder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Sent)
                ElseIf My.Settings.CurrentMailFolderName.Contains("Inbox") Then
                    currentImapFolder = ttm.MailKitClient.GetFolder("Inbox")
                ElseIf My.Settings.CurrentMailFolderName = "Trash" Or "Deleted" Or "Removed" Or "Recycle" Then
                    currentImapFolder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Trash)
                ElseIf My.Settings.CurrentMailFolderName = "Junk" Or "Harmful" Or "Malware" Or "Potential" Or "miscellaneous" Or My.Settings.CurrentMailFolderName.Contains("Spam") Then
                    currentImapFolder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Junk)
                ElseIf My.Settings.CurrentMailFolderName.Contains("all") Then
                    currentImapFolder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.All)
                ElseIf My.Settings.CurrentMailFolderName.Contains("all") Or My.Settings.CurrentMailFolderName.Contains("Everything") Then
                    currentImapFolder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.All)
                End If
                currentImapFolder.Open(FileAccess.ReadWrite)
            End If
            Dim Message = currentImapFolder.GetMessage(MailKit.UniqueId.Parse(My.Settings.OpenedMailId))
            Dim emlPath As String = CurrentMailFolder + My.Settings.OpenedMailId + ".MHT"
            Dim emlPath2 As String = CurrentMailFolder + My.Settings.OpenedMailId + "WI.MHT"
            Dim document = New HtmlAgilityPack.HtmlDocument()
            document.LoadHtml(Message.Body.ToString)
            If document.DocumentNode.InnerHtml.Contains("<img") Then
                For Each eachNode In document.DocumentNode.SelectNodes("//img")
                    eachNode.Attributes.Remove("src")
                Next
            End If
            Dim newhtml = document.DocumentNode.OuterHtml
            If Not File.Exists(emlPath) AndAlso Not File.Exists(emlPath2) Then
                Message.WriteTo(emlPath) : File.WriteAllText(emlPath2, newhtml)
            End If
            Dim match As Match = Regex.Match(Message.From.ToString, """([^""]*)""")
            With SingleEmaillViewer.EDetails
                .FromText = Message.From.ToString
                SingleEmaillViewer.ContactPane.EmailFrom = Message.From.ToString
                SingleEmaillViewer.ContactPane.From2.Text = match.Groups(1).Value
                .SubjectText = Message.Subject.ToString
                .EToText = "To : " + My.Settings.email
                If Message.Cc.ToString = "" Then
                    .CcText = ""
                Else
                    .CcText = "Cc : " + Message.Cc.ToString
                End If
            End With
            If Message.Attachments.Count <= 0 Then
                SingleEmaillViewer.EDetails.AttachVisibility = Visibility.Hidden
            Else
                SingleEmaillViewer.EDetails.AttachVisibility = Visibility.Visible
            End If
            SingleEmaillViewer.Browser1.Navigate(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + My.Settings.OpenedMailId + "WI.mht")
        End If
        SingleEmaillViewer.Browser1.Visibility = Visibility.Visible
        For Each btn In Teststack.Children.OfType(Of MailContainer1)
            With btn
                .From.Margin = New Thickness(7, 15, 0, 0)
                Grid.SetColumn(.From, 2)
                Grid.SetColumnSpan(.From, 2)
                .Subject.FontSize = 10
                .Subject.FontFamily = New FontFamily("Segoe UI Semibold")
                .Subject.Margin = New Thickness(7, 32, 0, -2)
                Grid.SetColumn(.Subject, 2)
                .EDate.Margin = New Thickness(7, 3, 0, 0)
                .EDate.FontSize = 9
                Grid.SetColumn(.EDate, 2)
                Grid.SetColumnSpan(.Subject, 3)
                .EDate.FontFamily = New FontFamily("Segoe UI Semibold")
                .From.Width = 100
                .Subject.Width = 200
            End With
        Next
    End Sub
    Private Sub Home_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        If Me.WindowState = WindowState.Maximized Then
            If Not Mainframe.Visibility = Visibility.Visible Then
                AllEmailsScrollViewer.HorizontalAlignment = HorizontalAlignment.Stretch
                AllEmailsScrollViewer.Width = Double.NaN
                For Each btn In Teststack.Children.OfType(Of MailContainer1)
                    With btn
                        .From.Margin = New Thickness(17, 15, 0, 0)
                        .Subject.FontSize = 14
                        .Subject.Margin = New Thickness(10.333, 23, 0, 0)
                        .Subject.FontFamily = New FontFamily("Segoe UI")
                        Grid.SetColumn(.Subject, 3)
                        Grid.SetColumnSpan(.Subject, 1)
                        .EDate.Margin = New Thickness(10.333, 6, 0, 0)
                        .EDate.FontSize = 11
                        Grid.SetColumn(.EDate, 3)
                    End With
                Next
            Else
                Dim CurrentPage = Mainframe.Content.GetType
                MsgBox(CurrentPage.ToString)
                If Not CurrentPage.ToString = "ContactPage" Then
                    AllEmailsScrollViewer.HorizontalAlignment = HorizontalAlignment.Left
                    AllEmailsScrollViewer.Width = 320
                    LoadMoreMails.Width = 320
                    LoadMoreMails.HorizontalAlignment = HorizontalAlignment.Left
                    For Each btn In Teststack.Children.OfType(Of MailContainer1)
                        With btn
                            .From.Margin = New Thickness(7, 15, 0, 0)
                            Grid.SetColumn(.From, 2)
                            Grid.SetColumnSpan(.From, 2)
                            .Subject.FontSize = 10
                            .Subject.FontFamily = New FontFamily("Segoe UI Semibold")
                            .Subject.Margin = New Thickness(7, 32, 0, -2)
                            Grid.SetColumn(.Subject, 2)
                            .EDate.Margin = New Thickness(7, 3, 0, 0)
                            .EDate.FontSize = 9
                            Grid.SetColumn(.EDate, 2)
                            Grid.SetColumnSpan(.Subject, 3)
                            .EDate.FontFamily = New FontFamily("Segoe UI Semibold")
                            .From.Width = 100
                            .Subject.Width = 200
                        End With
                    Next
                End If
            End If
        End If
    End Sub
    Private Sub HideSingleMail_Tick(sender As Object, e As EventArgs) Handles HideSingleMail.Tick
        If Mainframe.Visibility = Visibility.Hidden Then
            HideSingleMail.Stop()
            AllEmailsScrollViewer.HorizontalAlignment = HorizontalAlignment.Stretch
            AllEmailsScrollViewer.Width = Double.NaN
            For Each btn In Teststack.Children.OfType(Of MailContainer1)
                With btn
                    .From.Margin = New Thickness(17, 15, 0, 0)
                    .Subject.FontSize = 14
                    .Subject.Margin = New Thickness(10.333, 23, 0, 0)
                    .Subject.FontFamily = New FontFamily("Segoe UI")
                    Grid.SetColumn(.Subject, 3)
                    Grid.SetColumnSpan(.Subject, 1)
                    .EDate.Margin = New Thickness(10.333, 6, 0, 0)
                    .EDate.FontSize = 11
                    Grid.SetColumn(.EDate, 3)
                End With
            Next
        End If
    End Sub
    Public Sub GetAllMailsFromSender()
        AllEmailFromSender.Clear()
        Mainframe.Navigate(AllEmailOfViewer)
        Mainframe.Visibility = Visibility.Visible
        Mainframe.Opacity = 1
        Dim sb As New Storyboard
        sb = FindResource("ViewMainframe")
        Storyboard.SetTarget(sb, Mainframe)
        AddHandler sb.Completed, AddressOf AnimationDone
        sb.Begin()
    End Sub
    Friend Sub AnimationDone()
        StartAllMailProcess.Start()
    End Sub
    Private Sub StartAllMailProcess_Tick(sender As Object, e As EventArgs) Handles StartAllMailProcess.Tick
        If Not AllEmailOf.IsBusy = True Then
            If My.Settings.MailDownloadDone = True Then
                StartAllMailProcess.Stop()
                AllEmailOf.RunWorkerAsync()
                HideAllOfSpecificUser.Start()
            End If
        End If
    End Sub
    Private Sub AllEmailOf_DoWork(sender As Object, e As DoWorkEventArgs) Handles AllEmailOf.DoWork
        Dim Folder As IMailFolder = ttm.MailKitClient.GetFolder(My.Settings.CurrentMailFolderName) ' ttm.MailKitClient.GetFolder(My.Settings.CurrentMailFolderName)
        If My.Settings.CurrentMailFolderName = "Drafts" Then
            Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Drafts)
        ElseIf My.Settings.CurrentMailFolderName = "Send" Then
            Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Sent)
        ElseIf My.Settings.CurrentMailFolderName = "Inbox" Then
            Folder = ttm.MailKitClient.GetFolder("Inbox")
        End If
        Dim message = Folder.GetMessage(UniqueId.Parse(My.Settings.OpenedMailId))
        Dim SenderAddress As Match = Regex.Match(message.From.ToString, "<(.*?)>")
        If SenderAddress.Success Then
            EmailSender = SenderAddress.Groups(1).Value
        End If
        Dim eml As New Emails
        ''''Getting All mails from specific sender
        Dim uids = Folder.Search(SearchQuery.FromContains(EmailSender))
        For Each summary In Folder.Fetch(uids, MessageSummaryItems.Full)
            Dim Email As New Emails
            Dim From = If((summary.Envelope.From.ToString <> ""), summary.Envelope.From.ToString, "No sender")
            Dim match As Match = Regex.Match(From, """([^""]*)""")
            If match.Success Then
                Dim EmailFrom As String = match.Groups(1).Value
                Email.From = EmailFrom
            Else
                Email.From = summary.Envelope.From.ToString
            End If
            Email.UniqueId = summary.UniqueId.ToString
            Email.EmailDate = summary.Envelope.Date.ToString
            If summary.Flags.Value.HasFlag(MessageFlags.Seen) Then
                Email.IsSeen = "Seen"
            Else
                Email.IsSeen = "Unseen"
            End If
            'Email.EmailBody = summary.Body.ToString
            Try
                Email.Subject = summary.Envelope.Subject.ToString
            Catch ex As Exception
            End Try
            AllEmailFromSender.Add(Email)
        Next
    End Sub

    Private Sub AllEmailOf_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles AllEmailOf.RunWorkerCompleted
        AllEmailFromSender.Reverse()
        Try
            AllEmailOfViewer.Browser1.Navigate(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + My.Settings.OpenedMailId + ".mht")
        Catch ex As Exception
            MsgBox(ex)
        End Try
        AllEmailOfViewer.Browser1.Visibility = Visibility.Visible
        Mainframe.Visibility = Visibility.Visible
        Dim Folder As IMailFolder = ttm.MailKitClient.GetFolder(My.Settings.CurrentMailFolderName) ' ttm.MailKitClient.GetFolder(My.Settings.CurrentMailFolderName)
        If My.Settings.CurrentMailFolderName = "Drafts" Then
            Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Drafts)
        ElseIf My.Settings.CurrentMailFolderName = "Send" Then
            Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Sent)
        ElseIf My.Settings.CurrentMailFolderName = "Inbox" Then
            Folder = ttm.MailKitClient.GetFolder("Inbox")
        End If

        ''''''Displaying the current message
        Dim message = Folder.GetMessage(UniqueId.Parse(My.Settings.OpenedMailId))
        Dim SenderName As Match = Regex.Match(message.From.ToString, """([^""]*)""")
        Dim SenderAddress As Match = Regex.Match(message.From.ToString, "<(.*?)>")
        With AllEmailOfViewer
            If SenderName.Success Then
                Dim EmailFrom As String = SenderName.Groups(1).Value
                .From2.Text = EmailFrom
            Else
                .From2.Text = "No name"
            End If
            .From.Text = .From2.Text
            If SenderAddress.Success Then
                Dim EmailFrom As String = SenderAddress.Groups(1).Value
                .senderAddress.Text = EmailFrom
                EmailSender = EmailFrom
            End If
            .Subject.Text = message.Subject.ToString
            .To.Text = "To : " + message.To.ToString
            If message.Cc.ToString = "" Then
                .Cc.Text = ""
            Else
                For Each cc In message.Cc
                    .Cc.Text = "Cc : " + cc.ToString + ","
                Next
            End If
            If Not message.Attachments.Count <= 0 Then
                .AttachmentMsg.Visibility = Visibility.Visible
            End If
        End With
        ''''''Showing all emails from sender
        For Each email As Emails In AllEmailFromSender
            Dim btn As New smallEmailContainer
            btn.UniqueId.Text = email.UniqueId.ToString
            btn.Date.Text = email.EmailDate
            btn.Subject.Text = email.Subject
            AddHandler btn.MouseDown, AddressOf EachmailFromOneSender_Click
            AllEmailOfViewer.EmailHolder.Height = AllEmailOfViewer.EmailHolder.Height + 53
            AllEmailOfViewer.EmailHolder.Children.Add(btn)
        Next
    End Sub
    Private Sub EachmailFromOneSender_Click()
        ViewSingleMailOfSpecificUser.Start()
    End Sub
    Private Sub ViewSingleMailOfSpecificUser_Tick(sender As Object, e As EventArgs) Handles ViewSingleMailOfSpecificUser.Tick
        If Not My.Settings.OpenedMailId = "" Then
            ViewSingleMailOfSpecificUser.Stop()
            Dim Folder As IMailFolder = ttm.MailKitClient.GetFolder(My.Settings.CurrentMailFolderName) ' ttm.MailKitClient.GetFolder(My.Settings.CurrentMailFolderName)
            If My.Settings.CurrentMailFolderName = "Drafts" Then
                Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Drafts)
            ElseIf My.Settings.CurrentMailFolderName = "Send" Then
                Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Sent)
            ElseIf My.Settings.CurrentMailFolderName = "Inbox" Then
                Folder = ttm.MailKitClient.GetFolder("Inbox")
            End If
            Dim message = Folder.GetMessage(UniqueId.Parse(My.Settings.OpenedMailId))
            Dim match As Match = Regex.Match(message.From.ToString, """([^""]*)""")
            With AllEmailOfViewer
                If match.Success Then
                    Dim EmailFrom As String = match.Groups(1).Value
                    .From2.Text = EmailFrom
                Else
                    .From2.Text = "No name"
                End If
                .From.Text = message.From.ToString
                .Subject.Text = message.Subject.ToString
                .To.Text = "To : " + message.To.ToString
                If message.Cc.ToString = "" Then
                    .Cc.Text = ""
                Else
                    For Each cc In message.Cc
                        .Cc.Text = "Cc : " + cc.ToString + ","
                    Next
                End If
            End With
            If Not message.Attachments.Count <= 0 Then
                AllEmailOfViewer.AttachmentMsg.Visibility = Visibility.Visible
            Else
                AllEmailOfViewer.AttachmentMsg.Visibility = Visibility.Hidden
            End If
            Try
                AllEmailOfViewer.Browser1.Navigate(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + My.Settings.OpenedMailId + ".mht")
            Catch ex As Exception
                MsgBox(ex)
            End Try
            AllEmailOfViewer.Browser1.Visibility = Visibility.Visible
        End If
    End Sub
    Private Sub HideAllOfSpecificUser_Tick(sender As Object, e As EventArgs) Handles HideAllOfSpecificUser.Tick
        If Not Mainframe.Visibility = Visibility.Visible Then
            AllEmailsScrollViewer.HorizontalAlignment = HorizontalAlignment.Stretch
            AllEmailsScrollViewer.Width = Double.NaN
            AllEmailsScrollViewer.Visibility = Visibility.Visible
        End If
    End Sub
    Private Sub AddToFavourites()
        Dim Folder As IMailFolder = ttm.MailKitClient.GetFolder(My.Settings.CurrentMailFolderName) ' ttm.MailKitClient.GetFolder(My.Settings.CurrentMailFolderName)
        If My.Settings.CurrentMailFolderName = "Drafts" Then
            Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Drafts)
        ElseIf My.Settings.CurrentMailFolderName = "Send" Then
            Folder = ttm.MailKitClient.GetFolder(MailKit.SpecialFolder.Sent)
        ElseIf My.Settings.CurrentMailFolderName = "Inbox" Then
            Folder = ttm.MailKitClient.GetFolder("Inbox")
        End If
        If My.Settings.AddToFavouritesId.ToString.Contains("-") Then
            ''Removing from favourites
            Dim id As Integer = Convert.ToInt32(My.Settings.AddToFavouritesId.ToString.Split("-")(1))
            Dim emailDetails As New List(Of String)(System.IO.File.ReadAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail"))
            File.Delete(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail")
            If My.Settings.SeenUnseenId.ToString.Contains("-") Then
                File.WriteAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail", New String() {emailDetails(0), emailDetails(1), emailDetails(2), "No", "UnSeen"})
            Else
                File.WriteAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail", New String() {emailDetails(0), emailDetails(1), emailDetails(2), "No", "Seen"})
            End If
            TotalFavouritesMails = TotalFavouritesMails - 1
            Dim index As Integer = InboxMails.FindIndex(Function(a) a.UniqueId = id.ToString)
            InboxMails(index).Favourite = "No"
        Else
            TotalFavouritesMails = TotalFavouritesMails + 1
            ''''Adding to favourites
            Dim id = My.Settings.AddToFavouritesId
            Dim emailDetails As New List(Of String)(System.IO.File.ReadAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail"))
            File.Delete(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail")
            If My.Settings.SeenUnseenId.ToString.Contains("-") Then
                File.WriteAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail", New String() {emailDetails(0), emailDetails(1), emailDetails(2), "Yes", "UnSeen"})
            Else
                File.WriteAllLines(System.AppDomain.CurrentDomain.BaseDirectory() + "data\" + My.Settings.email + "\EmailFolders\" + My.Settings.CurrentMailFolderName + "\" + id.ToString + ".mail", New String() {emailDetails(0), emailDetails(1), emailDetails(2), "Yes", "Seen"})
            End If
            Dim index As Integer = InboxMails.FindIndex(Function(a) a.UniqueId = id.ToString)
            InboxMails(index).Favourite = "Yes"
        End If
    End Sub

    Private Sub ActsBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles ActsBtn.MouseDown
        Dim cm = ContextMenuService.GetContextMenu(TryCast(sender, DependencyObject))
        If cm Is Nothing Then
            Return
        Else
            cm.Placement = PlacementMode.Bottom
            cm.PlacementTarget = TryCast(sender, UIElement)
            cm.IsOpen = True
        End If
    End Sub

    Private Sub CreateFolderBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles CreateFolderBtn.MouseDown
        Dim cm = ContextMenuService.GetContextMenu(TryCast(sender, DependencyObject))
        If cm Is Nothing Then
            Return
        Else
            cm.Placement = PlacementMode.Top
            cm.PlacementTarget = TryCast(sender, UIElement)
            cm.IsOpen = True
        End If
    End Sub
    Private Sub CalenderBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles CalenderBtn.MouseDown
        Dim cm = ContextMenuService.GetContextMenu(TryCast(sender, DependencyObject))
        If cm Is Nothing Then
            Return
        Else
            cm.Placement = PlacementMode.Top
            cm.PlacementTarget = TryCast(sender, UIElement)
            cm.IsOpen = True
        End If
    End Sub
    Private Sub ShowContactsPage()
        Conpage.ConContainer.Children.Clear()
        Mainframe.Navigate(Conpage)
        Mainframe.Visibility = Visibility.Visible
        Mainframe.Opacity = 1
        Dim sb As New Storyboard
        sb = FindResource("ViewMainframeForSingleMail")
        AddHandler sb.Completed, AddressOf ViewContactPage_Done
        Storyboard.SetTarget(sb, Mainframe)
        sb.Begin()
        HideSingleMail.Start()
    End Sub
    Private Sub ViewContactPage_Done()
        For i = 0 To 10
            Conpage.ConContainer.Children.Add(New Contact With {.ContactName = "Aousaf Rashid",
                                              .ContactEmail = "nabilrashid44@gmail.com"})
        Next
    End Sub
End Class
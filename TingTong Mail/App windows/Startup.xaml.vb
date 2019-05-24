Imports WpfAnimatedGif
Imports System.Data.OleDb
Imports System.IO.Directory
Imports System.IO
Imports System.ComponentModel
Imports MailKit
Imports MailKit.Security

Public Class Startup
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetCurrentDirectory() & "\Data\09US_dt.ttmdata;Jet OLEDB:Database Password=arn33423342;")
    Public WithEvents bgw As New BackgroundWorker
    Dim mw As New MainWindow
    Dim i As Integer
    Dim err As String
    Public WithEvents tmr As New Forms.Timer
    Public WithEvents tmr2 As New Forms.Timer
    Dim home1 As New Home
    Public Shared Property MailKitClient As New MailKit.Net.Imap.ImapClient(New ProtocolLogger("imap.log"))
    Dim photoloader As New ImageLoader
    Friend Sub rt()
        My.Settings.Reset()
    End Sub
    Private Sub LoadLoggedInUsersDetails()
        con.Open()
        Dim cmd2 As New OleDbCommand("Select * from userinfo where Email='" & My.Settings.email & "'", con)
        Dim photo As Byte()
        Dim dr2 As OleDbDataReader = cmd2.ExecuteReader
        While dr2.Read
            If IsDBNull(dr2(11)) Then
            Else
                photo = CType(dr2(11), Byte())
                Dim strm As MemoryStream = New MemoryStream(photo)
                Dim bi As New BitmapImage
                bi.BeginInit()
                strm.Seek(0, SeekOrigin.Begin)
                bi.StreamSource = strm
                bi.EndInit()
                'home1.UserPic.ImageSource = bi
            End If
        End While
        con.Close()
    End Sub

    Private Sub TingTong_Mail_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim image As New BitmapImage
        image.BeginInit()
        image.UriSource = New Uri("pack://application:,,,/TingTong Mail;component/icons/LoadingAnim.data")
        image.EndInit()
        ImageBehavior.SetAnimatedSource(LoadingAnim, image)
        My.Settings.openmode = "Default"
        bgw.WorkerReportsProgress = True
        If My.Settings.imappop = "IMAP" Then
            If My.Settings.setupdone = "Done" Then
                bgw.RunWorkerAsync()
            End If
        Else
            tmr2.Start()
        End If
    End Sub
    Private Sub rect1_MouseDown(sender As Object, e As Input.MouseEventArgs) Handles rect1.MouseDown
        Application.Current.Shutdown()
    End Sub

    Private Sub Rect1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles rect1.MouseDown
        Me.DragMove()
    End Sub
    Private Sub tmr_Tick(sender As Object, e As EventArgs) Handles tmr.Tick
        i = i + 1
        If i >= 8 Then
            tmr.Stop()
            bgw.RunWorkerAsync()
        End If
    End Sub
    Private Sub bgw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgw.RunWorkerCompleted
        If err = "Logged in" Then
            Me.Close()
            LoadLoggedInUsersDetails()
            home1.Show()
        Else
            My.Settings.setupdone = ""
            My.Settings.Save()
            MsgBox(err + vbCrLf + "Error 49 : There was a problem logging in to your account." + vbCrLf + "You must log in again.", MsgBoxStyle.Exclamation)
            My.Settings.imappop = ""
            My.Settings.setupdone = ""
            My.Settings.Save()
            If CBool(MsgBoxResult.Ok) Then
                Me.Close()
                mw.Show()
            Else
                Application.Current.Shutdown()
            End If
        End If
    End Sub

    Private Sub bgw_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgw.ProgressChanged
        Progress.Text = e.UserState.ToString
    End Sub
    Private Sub bgw_DoWork(sender As Object, f As DoWorkEventArgs) Handles bgw.DoWork
        Try
            Dim port As Integer = CInt(My.Settings.txtport)
            MailKitClient.ServerCertificateValidationCallback = Function(s, c, h, e) True
            For i = 0 To 100
                bgw.ReportProgress(i, "Connecting to server")
            Next
            MailKitClient.Connect(My.Settings.txtserver, port, SecureSocketOptions.SslOnConnect)
            For i = 0 To 100
                bgw.ReportProgress(i, "Logging in to " + My.Settings.email)
            Next
            MailKitClient.Authenticate(My.Settings.email, My.Settings.password)
            err = "Logged in"
        Catch ex As Exception
            err = err + vbCrLf + ex.ToString
        End Try
    End Sub
    Private Sub tmr2_Tick(sender As Object, e As EventArgs) Handles tmr2.Tick
        i = i + 1
        If i >= 8 Then
            tmr2.Stop()
            Me.Close()
            mw.Show()
        End If
    End Sub
    Private Sub TingTong_Mail_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.Key = Key.[F3] Then
            rt()
        End If
    End Sub
End Class
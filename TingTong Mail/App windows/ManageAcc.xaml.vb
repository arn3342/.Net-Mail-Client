Imports System.Windows.Media.Animation
Imports System.Data
Imports System.Data.OleDb
Imports System.IO.Directory
Imports System.IO
Imports System.ComponentModel
Imports System.Drawing.Imaging

Public Class ManageAcc
    Dim accountsphotos As New List(Of BitmapImage)
    Dim CurrentUserPhoto As New BitmapImage
    Dim CurrentUser As String
    Dim AllUsers As New List(Of LoadAllUsersClass)
    Dim newUsers As New List(Of LoadAllUsersClass)
    Public WithEvents LoadCurrentAccount As New BackgroundWorker
    Public WithEvents LoadAllAccounts As New BackgroundWorker
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetCurrentDirectory() & "\Data\09US_dt.ttmdata;Jet OLEDB:Database Password=arn33423342;")
    Private Sub AddBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles AddBtn.MouseDown
        If AccountsPanel.Children.Count >= 5 Then
            MsgBox("Sorry,but you cannot add more than 5 accounts.Please remove an existing account to continue.")
        Else
            My.Settings.openmode = "Add"
            Dim mw As New MainWindow
            mw.ShowDialog()
            AccountsPanel.Children.Clear()
            con.Open()
            Dim cmd As New OleDbCommand("Select * from userinfo", con)
            Dim dr As OleDbDataReader = cmd.ExecuteReader
            Dim _photo As Byte()
            While dr.Read
                Dim accounts As New emailUserList
                With accounts
                    .UserName.Text = dr(9).ToString
                    .UserEmail.Text = dr(0).ToString
                    .LastLogin.Text = dr(10).ToString
                End With
                If IsDBNull(dr(11)) Then
                Else
                    _photo = CType(dr(11), Byte())
                    Dim strm As MemoryStream = New MemoryStream(_photo)
                    Dim bi As New BitmapImage
                    bi.BeginInit()
                    bi.CacheOption = BitmapCacheOption.OnLoad
                    bi.StreamSource = strm
                    bi.EndInit()
                    accounts.userphoto.ImageSource = bi
                End If
                AccountsPanel.Children.Add(accounts)
                If accounts.UserEmail.Text = UserEmail.Text Then
                    AccountsPanel.Children.Remove(accounts)
                End If
            End While
            con.Close()
            If Not AccountsPanel.Children.Count <= 0 Then
                NoAccounts.Visibility = Visibility.Hidden
            End If
        End If
    End Sub
    Private Sub AddBtn_MouseEnter(sender As Object, e As MouseEventArgs) Handles AddBtn.MouseEnter
        Dim sb As New Storyboard
        sb = FindResource("SlideRight")
        Storyboard.SetTarget(sb, AddTxt)
        sb.Begin()
        Dim sb2 As New Storyboard
        sb2 = FindResource("Show")
        Storyboard.SetTarget(sb2, addImg)
        sb2.Begin()
    End Sub

    Private Sub AddBtn_MouseLeave(sender As Object, e As MouseEventArgs) Handles AddBtn.MouseLeave
        Dim sb2 As New Storyboard
        sb2 = FindResource("Hide")
        Storyboard.SetTarget(sb2, addImg)
        sb2.Begin()
        Dim sb As New Storyboard
        sb = FindResource("SlideLeft")
        Storyboard.SetTarget(sb, AddTxt)
        sb.Begin()
    End Sub
    Private Sub AddName_KeyDown(sender As Object, e As KeyEventArgs) Handles AddName.KeyDown
        If e.Key = Key.Return Then
            If AddName.Text = "" Then
            Else
                CurrentUserName.Content = AddName.Text
                con.Open()
                Dim cmd As New OleDbCommand("UPDATE userinfo SET Uname='" & CurrentUserName.Content & "' WHERE Email='" & UserEmail.Text & "'", con)
                cmd.ExecuteNonQuery()
                con.Close()
                AddName.Visibility = Visibility.Hidden
            End If
        End If
    End Sub

    Private Sub CurrentUserName_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles CurrentUserName.MouseDoubleClick
        AddName.Visibility = Visibility.Visible
        AddName.Focus()
    End Sub

    Private Sub CurrentUserName_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles CurrentUserName.MouseDown

    End Sub

    Private Sub ManageAcc_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        UserEmail.Text = My.Settings.email
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
                userpic.ImageSource = bi
                CurrentUser = dr2(9).ToString
            End If
        End While
        'con.Close()
        If CurrentUser = "" Then
            CurrentUserName.Content = "Double click to add a name"
        Else
            CurrentUserName.Content = CurrentUser
        End If
        LoadAllAccounts.RunWorkerAsync()
    End Sub
    Friend Sub accounts_LoginBtn_MouseDown()
    End Sub
    Private Sub UserPic_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles usPic.MouseDown
        Dim ofpd As New System.Windows.Forms.OpenFileDialog
        ofpd.Filter = "Image Files(*.BMP;*.JPG;*.PNG;*.JPEG)|*.BMP;*.JPG;*.PNG;*.JPEG"
        ofpd.Title = "Add a photo"
        If ofpd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim fill = New ImageBrush(New BitmapImage(New Uri(ofpd.FileName)))
            RenderOptions.SetBitmapScalingMode(fill, BitmapScalingMode.Fant)
            usPic.Fill = fill
            con.Open()
            Dim cmd As New OleDbCommand("UPDATE userinfo SET Pic=@pic WHERE Email='" & My.Settings.email & "'", con)
            Dim data As Byte()
            Dim myimage As System.Drawing.Image = System.Drawing.Image.FromFile(ofpd.FileName)
            Using ms As MemoryStream = New MemoryStream()
                myimage.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp)
                data = ms.ToArray()
            End Using
            cmd.Parameters.AddWithValue("@pic", data)
            cmd.ExecuteNonQuery()
            con.Close()
        End If
    End Sub
    Private Sub LoadAllAccounts_DoWork(sender As Object, e As DoWorkEventArgs) Handles LoadAllAccounts.DoWork
        Dim cmd As New OleDbCommand("Select * from userinfo", con)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        Dim _photo As Byte()
        While dr.Read
            Dim allaccounts As New LoadAllUsersClass
            allaccounts.AllUserNames = (dr(9).ToString)
            allaccounts.AllUserEmails = (dr(0).ToString)
            allaccounts.UserLastLogIn = (dr(10).ToString)
            If IsDBNull(dr(11)) Then
            Else
                _photo = CType(dr(11), Byte())
                Dim strm As MemoryStream = New MemoryStream(_photo)
                Dim bi As New BitmapImage
                bi.BeginInit()
                bi.CacheOption = BitmapCacheOption.OnLoad
                bi.StreamSource = strm
                bi.EndInit()
                bi.Freeze()
                allaccounts.AllUserPhotos = bi
                strm.Close()
            End If
            newUsers.Add(allaccounts)
        End While
        con.Close()
    End Sub

    Private Sub LoadAllAccounts_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles LoadAllAccounts.RunWorkerCompleted
        For Each allUseraccount As LoadAllUsersClass In newUsers
            Dim accounts As New emailUserList
            With accounts
                .UserName.Text = allUseraccount.AllUserNames
                .UserEmail.Text = allUseraccount.AllUserEmails
                .LastLogin.Text = allUseraccount.UserLastLogIn
                Try
                    .userphoto.ImageSource = allUseraccount.AllUserPhotos
                Catch ex As Exception
                End Try
            End With
            AddHandler accounts.LogInBtn.MouseDown, AddressOf accounts_LoginBtn_MouseDown
            AccountsPanel.Children.Add(accounts)
            loadingaccounts.Visibility = Visibility.Hidden
            If accounts.UserEmail.Text = UserEmail.Text Then
                AccountsPanel.Children.Remove(accounts)
            End If
        Next
        If AccountsPanel.Children.Count <= 0 Then
            loadingaccounts.Visibility = Visibility.Hidden
            NoAccounts.Visibility = Visibility.Visible
        End If
    End Sub
End Class

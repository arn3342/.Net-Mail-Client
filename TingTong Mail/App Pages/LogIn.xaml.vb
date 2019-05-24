Imports System.Data.OleDb
Imports System.ComponentModel
Imports System.IO.Directory
Imports WpfAnimatedGif
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Class LogIn
    Dim err As String
    Dim i As Integer
    Public WithEvents tmr As New Forms.Timer
    Dim bb As New BrushConverter
    Public WithEvents manulaworker As New BackgroundWorker
    Public WithEvents autoworker As New BackgroundWorker
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetCurrentDirectory() & "\Data\sim.ttmdata;Jet OLEDB:Database Password=arn33423342;")
    Dim con2 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetCurrentDirectory() & "\Data\09US_dt.ttmdata;Jet OLEDB:Database Password=arn33423342;")
    'Dim mail As String = ""
    Private Sub WindowsFormsHost1_GotFocus(sender As Object, e As RoutedEventArgs) Handles newwindows1.GotFocus
        If txtEmail.Text = "Email" Then
            txtEmail.Text = ""
        End If
        txtEmail.ForeColor = System.Drawing.Color.Black
    End Sub
    Private Sub WindowsFormsHost1_LostFocus(sender As Object, e As RoutedEventArgs) Handles newwindows1.LostFocus
        If txtEmail.Text = "" Then
            txtEmail.Text = "Email"
        End If
        txtEmail.ForeColor = System.Drawing.ColorTranslator.FromHtml("#8a8b8c")
    End Sub
    Private Sub WindowsFormsHost_GotFocus(sender As Object, e As RoutedEventArgs) Handles passowrdwindow.GotFocus
        If txtpassword.Text = "Password" Then
            txtpassword.PasswordChar = "•"
            txtpassword.Text = ""
        End If
        txtpassword.ForeColor = System.Drawing.Color.Black
    End Sub
    Private Sub WindowsFormsHost_LostFocus(sender As Object, e As RoutedEventArgs) Handles passowrdwindow.LostFocus
        If txtpassword.Text = "" Then
            txtpassword.Text = "Password"
            txtpassword.PasswordChar = ""
        End If
        txtpassword.ForeColor = System.Drawing.ColorTranslator.FromHtml("#8a8b8c")
    End Sub
    Private Sub checkBox_Checked(sender As Object, e As RoutedEventArgs) Handles checkBox.Checked
        My.Settings.setup = "Manual"
    End Sub

    Private Sub checkBox_Unchecked(sender As Object, e As RoutedEventArgs) Handles checkBox.Unchecked
        My.Settings.setup = ""
    End Sub

    Public Sub loginprerequisite()
        con2.Open()
        Dim cmd As New OleDbCommand("Select * from userinfo WHERE Email=@email", con2)
        cmd.Parameters.AddWithValue("@email", txtEmail.Text)
        Dim table As New DataTable
        Dim ada As New OleDbDataAdapter(cmd)
        ada.Fill(table)
        If My.Settings.openmode = "Add" Then
            If txtEmail.Text = My.Settings.email Then
                MsgBox("You are currently logged in with this account.")
            Else
                If Not table.Rows.Count <= 0 Then
                    MsgBox("The account is already added.Please add another account.")
                Else
                    Login()
                End If
            End If
        Else
            Login()
        End If
        con2.Close()
    End Sub
    Private Sub loginrect_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles loginrect.MouseDown
        loginprerequisite()
    End Sub
    Public Sub Login()
        tmr.Interval = 400
        If manulaworker.IsBusy Then
        Else
            Buttons.Visibility = Visibility.Hidden
            Dim image As New BitmapImage
            image.BeginInit()
            image.UriSource = New Uri(GetCurrentDirectory() + "\LoadingAnim.data")
            image.EndInit()
            ImageBehavior.SetAnimatedSource(anim, image)
            anim.Visibility = Visibility.Visible
            If My.Settings.setup = "Manual" Then
                If My.Settings.openmode = "Add" Then
                    My.Settings.addAccountEmail = txtEmail.Text
                    My.Settings.AddAccountPassword = txtpassword.Text
                Else
                    My.Settings.email = txtEmail.Text
                    My.Settings.password = txtpassword.Text
                End If
                tmr.Start()
                Else
                    manulaworker.RunWorkerAsync()
            End If
        End If
    End Sub

    Private Sub LogIn_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        txtpassword.Font = New Font("Segoe UI SemiLight", 13)
        txtEmail.Font = New Font("Segoe UI SemiLight", 13)
        txtpassword.ForeColor = ColorTranslator.FromHtml("#8a8b8c")
        txtEmail.ForeColor = ColorTranslator.FromHtml("#8a8b8c")
        txtEmail.Text = "Email"
        txtpassword.Text = "Password"
    End Sub

    Private Sub bgw_DoWork(sender As Object, e As DoWorkEventArgs) Handles manulaworker.DoWork
        If txtEmail.Text = "Email" Or txtEmail.Text = "" Or txtpassword.Text = "" Or txtpassword.Text = "Password" Then
            err = "Error 12: Please enter a valid email address and password"
        Else
            If Not My.Settings.openmode = "Add" Then
                With My.Settings
                    .email = txtEmail.Text
                    .password = txtpassword.Text
                End With
            Else
                My.Settings.addAccountEmail = txtEmail.Text
                My.Settings.addAccountPassword = txtpassword.Text
            End If
            Try
                Dim mail As String
                mail = "@" + txtEmail.Text.Split("@")(1)
                Dim cmd As New OleDbCommand("Select * from impsmtp where name=@name", con)
                cmd.Parameters.Add("@name", OleDbType.VarChar).Value = mail
                Dim adap As New OleDbDataAdapter(cmd)
                Dim table As New DataTable
                adap.Fill(table)
                If table.Rows.Count <= 0 Then
                    My.Settings.setup = "Manual"
                Else
                    My.Settings.setup = "Auto"
                End If
            Catch ex As Exception
                MsgBox(ex)
                err = "Error 17 : " + ex.Message
                End Try
            End If
    End Sub
    Private Sub bgw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles manulaworker.RunWorkerCompleted
        If err = "" Then
            anim.Visibility = Visibility.Hidden
            My.Settings.Save()
            If My.Settings.setup = "Manual" Then
                NavigationService.Navigate(New poporimap)
            Else
                NavigationService.Navigate(New warning)
            End If
        Else
            MsgBox(err, MsgBoxStyle.Information)
        End If
    End Sub
    Private Sub tmr_Tick(sender As Object, e As EventArgs) Handles tmr.Tick
        i = i + 1
        If i >= 8 Then
            tmr.Stop()
            NavigationService.Navigate(New poporimap)
        End If
    End Sub
    Private Sub txtpassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtpassword.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.[Return]) Then
            If txtEmail.Text = "Email" Or txtEmail.Text = "" Then
            Else
                loginprerequisite()
            End If
        End If
        If e.KeyChar = Convert.ToChar(Keys.Tab) Then
            txtpassword.Select()
        End If
    End Sub

    Private Sub txtEmail_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtEmail.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.[Return]) Then
            If txtpassword.Text = "" Or txtpassword.Text = "Password" Then
            Else
                loginprerequisite()
            End If
        End If
    End Sub

    Private Sub LogIn_KeyDown(sender As Object, e As Input.KeyEventArgs) Handles Me.KeyDown

    End Sub
End Class

Imports System.IO.Directory
Imports WpfAnimatedGif
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Data


Class warning
    Dim err As String
    Dim i As Integer
    Public WithEvents tmr2 As New Timer
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetCurrentDirectory() & "\Data\sim.ttmdata;Jet OLEDB:Database Password=arn33423342;")
    Dim mail As String = ""
    Dim con2 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetCurrentDirectory() & "\Data\09US_dt.ttmdata;Jet OLEDB:Database Password=arn33423342;")
    Dim bgworker As New BackgroundWorker
    Public WithEvents backworker2 As New BackgroundWorker

    Private Sub tmr2_Tick(sender As Object, e As EventArgs) Handles tmr2.Tick
        i = i + 1
        If i >= 3 Then
            tmr2.Stop()
            backworker2.RunWorkerAsync()
        End If
    End Sub
    Private Sub Page_Loaded(sender As Object, e As RoutedEventArgs)
        tmr2.Interval = 250
        tmr2.Start()
        Dim image As New BitmapImage
        image.BeginInit()
        image.UriSource = New Uri(GetCurrentDirectory() + "\LoadingAnim.data")
        image.EndInit()
        ImageBehavior.SetAnimatedSource(anim, image)
        anim.Visibility = Visibility.Visible
        anim.Opacity = 0
        label.Opacity = 0
    End Sub
    Private Sub bgw2_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles backworker2.RunWorkerCompleted
        Dim currentuser As String
        If My.Settings.openmode = "Add" Then
            currentuser = My.Settings.addAccountEmail
        Else
            currentuser = My.Settings.email
        End If
        If err = "Logged in" Then
            My.Settings.setupdone = "Done"
            My.Settings.Save()
            label.Content = "  Ready to go!"
            Startgrid.Visibility = Visibility.Visible
            anim.Visibility = Visibility.Hidden
            If Not IO.Directory.Exists(GetCurrentDirectory() + "\Data\" + currentuser) Then
                IO.Directory.CreateDirectory(GetCurrentDirectory() + "\Data\" + currentuser)
                IO.Directory.CreateDirectory(GetCurrentDirectory() + "\Data\" + currentuser + "\Folders\")
                IO.Directory.CreateDirectory(GetCurrentDirectory() + "\Data\" + currentuser + "\Favourites\")
                IO.Directory.CreateDirectory(GetCurrentDirectory() + "\Data\" + currentuser + "\EmailFolders\")
                Dim tg As String = GetCurrentDirectory() & "\Data\" & currentuser & "\Userdata.ttmdata"
                Dim target As String = GetCurrentDirectory() & "\Data\" & currentuser & "\_userCons.ttmdata"
                Dim target2 As String = GetCurrentDirectory() & "\Data\" & currentuser & "\_Cons.ttmdata"
                IO.File.Copy(GetCurrentDirectory() + "\Data\Userdata.ttmdata", tg, True)
                IO.File.Copy(GetCurrentDirectory() + "\Data\_userCons.ttmdata", target, True)
                IO.File.Copy(GetCurrentDirectory() + "\Data\_Cons.ttmdata", target2, True)
                If My.Settings.openmode = "Add" Then
                Else
                    If IO.Directory.Exists(GetCurrentDirectory() + "\Data\ctuser.data") Then
                        My.Computer.FileSystem.DeleteFile(GetCurrentDirectory() + "\Data\ctuser.data")
                        IO.Directory.Delete(GetCurrentDirectory() + "\Data\" + currentuser + "\EmailFolders\")
                        IO.File.WriteAllText(GetCurrentDirectory() + "\Data\ctuser.data", My.Settings.email)
                    Else
                        IO.Directory.CreateDirectory(GetCurrentDirectory() + "\Data\" + currentuser + "\EmailFolders\")
                        IO.File.WriteAllText(GetCurrentDirectory() + "\Data\ctuser.data", My.Settings.email)
                    End If
                End If
                Dim con3 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & tg & ";Jet OLEDB:Database Password=arn33423342;")
                con3.Open()
                Dim cmd As New OleDbCommand("Insert into userdata(Email)values(@mail)", con3)
                cmd.Parameters.AddWithValue("@mail", currentuser)
                cmd.ExecuteNonQuery()
                con3.Close()
            End If
            My.Settings.addAccount = 0
            My.Settings.Save()
        Else
            label.Content = "       Oops!"
            err = err + vbCrLf + IMClient.ImpClient.Errormessage
            Startgrid.Visibility = Visibility.Hidden
            anim.Visibility = Visibility.Hidden
            failgrid.Visibility = Visibility.Visible
        End If
    End Sub
    Private Sub bgw2_DoWork(sender As Object, f As DoWorkEventArgs) Handles backworker2.DoWork
        mail = "@" + My.Settings.email.ToString.Split("@")(1)
        con.Open()
        con2.Open()
        Dim cmd As New OleDbCommand("Select * from impsmtp where name=@name", con)
        cmd.Parameters.Add("@name", OleDbType.VarChar).Value = mail
        Dim cmd2 As New OleDbCommand("INSERT INTO userinfo (Email,Pass,Server1,Server2,INPORT,OUTPORT) values (@email,@password,@server1,@server2,@inport,@outport)", con2)
        Dim adap As New OleDbDataAdapter(cmd)
        Dim table As New DataTable
        adap.Fill(table)

        'Checking if the user configured the server manually.....
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>...

        If My.Settings.setup = "Manual" Then
            Dim cmd3 As New OleDbCommand("Insert into userinfo(Email,Pass,IMAPPOP,SERVER1,SERVER2,INPORT,OUTPORT,SSL1,SSL2)values(@email,@password,@imappop,@server1,@server2,@inport,@outport,@ssl1,@ssl2)", con2)
            Try
                If Not My.Settings.openmode = "Add" Then
                    With cmd3.Parameters
                        .Add("@email", OleDbType.VarChar).Value = My.Settings.email
                        .Add("@password", OleDbType.VarChar).Value = My.Settings.password

                    End With
                Else
                    With cmd3.Parameters
                        .Add("@email", OleDbType.VarChar).Value = My.Settings.addAccountEmail
                        .Add("@password", OleDbType.VarChar).Value = My.Settings.AddAccountPassword
                    End With
                End If
                With cmd3.Parameters
                    .Add("@imappop", OleDbType.VarChar).Value = My.Settings.imappop
                    .Add("@server1", OleDbType.VarChar).Value = My.Settings.txtserver
                    .Add("@server2", OleDbType.VarChar).Value = My.Settings.txtserver1
                    .Add("@inport", OleDbType.VarChar).Value = My.Settings.txtport
                    .Add("@outport", OleDbType.VarChar).Value = My.Settings.txtport1
                    .Add("@ssl1", OleDbType.VarChar).Value = My.Settings.ssl1
                    .Add("@ssl2", OleDbType.VarChar).Value = My.Settings.ssl2
                End With
            Catch ex As Exception
                err = err + vbCrLf + "Error 41 " + vbCrLf + ex.Message
            End Try

            'Checking if the the server type is IMAP or not 
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>......

            If My.Settings.imappop = "IMAP" Then

                IMClient.ImpClient.ImapC = My.Settings.txtserver
                IMClient.ImpClient.port = My.Settings.txtport
                IMClient.ImpClient.Initialize()
                If Not My.Settings.openmode = "Add" Then

                    'WHen the user is logging in  with his account

                    If IMClient.ImpClient.Login(My.Settings.email, My.Settings.password) Then
                        'Checking if the user logged in successfully or not
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>........
                        err = "Logged in"
                        cmd3.ExecuteNonQuery()
                        My.Settings.imappop = "IMAP"


                    Else

                        err = "Error 44 : Failed to authenticate user.Please re-check your email and password."
                    End If
                Else

                    'When the user is adding an account and not logging in with that account

                    If IMClient.ImpClient.Login(My.Settings.addAccountEmail, My.Settings.AddAccountPassword) Then

                        'Checking if the user logged in successfully or not
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>........
                        My.Settings.txtserver = table.Rows(0)(1)
                        My.Settings.txtport = table.Rows(0)(3)
                        err = "Logged in"
                        cmd3.ExecuteNonQuery()
                        My.Settings.imappop = "IMAP"

                    Else

                        err = "Error 44 : Failed to authenticate user.Please re-check your email and password."
                    End If
                End If
            Else

                'When the server type is POP
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>.........

            End If

        Else

            'When the server was configured automatically
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>..........
            If table.Rows.Count <= 0 Then
                err = "There was a problem connection to the server"
            Else
                If Not My.Settings.openmode = "Add" Then

                    With cmd2.Parameters
                        .Add("@email", OleDbType.VarChar).Value = My.Settings.email
                        .Add("@password", OleDbType.VarChar).Value = My.Settings.password
                    End With
                Else
                    With cmd2.Parameters
                        .Add("@email", OleDbType.VarChar).Value = My.Settings.addAccountEmail
                        .Add("@password", OleDbType.VarChar).Value = My.Settings.AddAccountPassword
                    End With

                End If
                With cmd2.Parameters
                    .Add("@server1", OleDbType.VarChar).Value = table.Rows(0)(1)
                    .Add("@server2", OleDbType.VarChar).Value = table.Rows(0)(2)
                    .Add("@inport", OleDbType.VarChar).Value = table.Rows(0)(3)
                    .Add("@outport", OleDbType.VarChar).Value = table.Rows(0)(4)
                End With
                Dim int As Integer = table.Rows(0)(3)
                IMClient.ImpClient.ImapC = table.Rows(0)(1)
                IMClient.ImpClient.port = int
                IMClient.ImpClient.Initialize()

                'Checking if the user logged in successfully (when server was configured automatically)
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>.............


                'When the user is adding an account and not logging in with that account

                If Not My.Settings.openmode = "Add" Then
                    If IMClient.ImpClient.Login(My.Settings.email, My.Settings.password) Then
                        'When login is successful,adding the data to the database
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>..............
                        My.Settings.txtserver = table.Rows(0)(1)
                        My.Settings.txtport = table.Rows(0)(3)

                        Try
                            cmd2.ExecuteNonQuery()
                            ' MsgBox("Data Added 1")
                            con2.Close()
                            con.Close()
                        Catch ex As Exception
                            err = err + "Error 36 : " + ex.Message
                        End Try
                        'Checking if there was any previous error
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>..............
                        My.Settings.imappop = "IMAP"
                        err = "Logged in"

                    Else

                        'When the login was unsuccessful
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>...............

                        err = err + vbCrLf + "Error 44 : Failed to authenticate user.Please re-check your email and password."
                    End If
                Else
                    If IMClient.ImpClient.Login(My.Settings.addAccountEmail, My.Settings.AddAccountPassword) Then
                        'When login is successful,adding the data to the database
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>..............
                        My.Settings.txtserver = table.Rows(0)(1)
                        My.Settings.txtport = table.Rows(0)(3)

                        Try

                            cmd2.ExecuteNonQuery()
                            ' MsgBox("Data Added 2")
                            con2.Close()
                            con.Close()
                        Catch ex As Exception
                            err = err + "Error 36 : " + ex.Message
                        End Try
                        'Checking if there was any previous error
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>..............
                        My.Settings.imappop = "IMAP"
                        err = "Logged in"

                    Else

                        'When the login was unsuccessful
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>...............

                        err = err + vbCrLf + "Error 44 : Failed to authenticate user.Please re-check your email and password."
                    End If
                End If
            End If
        End If
        con.Close()
        con2.Close()
    End Sub
    Private Sub details_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles details.MouseDown
        MsgBox(err, MsgBoxStyle.Information)
    End Sub
    Private Sub back_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles back.MouseDown
        NavigationService.Navigate(New LogIn)
    End Sub
    Private Sub retry_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles retry.MouseDown
        failgrid.Visibility = Visibility.Hidden
        label.Content = "Just a moment.."
        backworker2.RunWorkerAsync()
    End Sub
    Private Sub next2_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles next2.MouseDown
        My.Settings.shutdowncommand = 1
        Dim mainwindow = Window.GetWindow(Me)
        If My.Settings.openmode = "Add" Then
            mainwindow.Close()
        Else
            My.Settings.openmode = ""
            Process.Start(IO.Directory.GetCurrentDirectory + "\TingTong Mail.exe")
            Application.Current.Shutdown()
        End If
    End Sub
End Class

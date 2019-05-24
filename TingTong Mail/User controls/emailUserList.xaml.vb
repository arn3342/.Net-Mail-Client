Imports System.Windows.Media.Animation
Imports System.IO.Directory
Imports System.Data.OleDb

Public Class emailUserList
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetCurrentDirectory() & "\Data\09US_dt.ttmdata;Jet OLEDB:Database Password=arn33423342;")

    Private Sub deletebtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles deletebtn.MouseDown
        warning.Text = "Are you sure you want to remove this email ? This will permanently delete the email."
        Dim sb As New Storyboard
        sb = FindResource("ShowDelete")
        Storyboard.SetTarget(sb, DeleteGrid)
        sb.Begin()
    End Sub

    Private Sub LogInBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles LogInBtn.MouseDown
        warning.Text = "Are you sure you want to log in with " + UserEmail.Text + " ?"
        Dim sb As New Storyboard
        sb = FindResource("ShowDelete")
        Storyboard.SetTarget(sb, DeleteGrid)
        sb.Begin()
    End Sub

    Private Sub LogInBtn_MouseEnter(sender As Object, e As MouseEventArgs) Handles LogInBtn.MouseEnter

    End Sub

    Private Sub NoBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles NoBtn.MouseDown
        Dim sb As New Storyboard
        sb = FindResource("HideDelete")
        Storyboard.SetTarget(sb, DeleteGrid)
        sb.Begin()
    End Sub

    Private Sub YesBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles YesBtn.MouseDown
        If warning.Text.Contains("@") Then
            con.Open()
            Dim cmd As New OleDbCommand("Select * from userinfo where Email='" & UserEmail.Text & "'", con)
            Dim dr2 As OleDbDataReader = cmd.ExecuteReader
            While dr2.Read
                With My.Settings
                    .txtport = dr2(5).ToString
                    .txtserver = dr2(3).ToString
                    .email = dr2(0).ToString
                    .password = dr2(1).ToString
                End With
                My.Settings.Save()
            End While
            If IO.Directory.Exists(GetCurrentDirectory() + "\Data\ctuser.data") Then
                My.Computer.FileSystem.DeleteFile(GetCurrentDirectory() + "\Data\ctuser.data")
                IO.File.WriteAllText(GetCurrentDirectory() + "\Data\ctuser.data", UserEmail.Text)
            Else
                IO.File.WriteAllText(GetCurrentDirectory() + "\Data\ctuser.data", UserEmail.Text)
            End If
            con.Close()
            Process.Start(IO.Directory.GetCurrentDirectory + "\TingTong Mail.exe")
            Application.Current.Shutdown()
        Else
            con.Open()
            Dim cmd As New OleDbCommand("Delete from userinfo WHERE Email='" & UserEmail.Text & "'", con)
            cmd.ExecuteNonQuery()
            con.Close()
            CType(Me.Parent, StackPanel).Children.Remove(Me)
            IO.Directory.Delete(GetCurrentDirectory() & "\Data\" & UserEmail.Text & "\", True)
        End If
    End Sub
End Class

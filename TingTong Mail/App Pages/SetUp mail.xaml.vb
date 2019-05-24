Imports System.Drawing
Imports WpfAnimatedGif
Imports System.IO
Imports System.Data.OleDb
Imports System.ComponentModel
Imports System.Windows.Forms

Class SetUp_mail
    Public WithEvents tmr2 As New Timer
    Dim err As String = ""
    Dim i As Integer
    Dim con2 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Directory.GetCurrentDirectory & "\Data\09US_dt.ttmdata;Jet OLEDB:Database Password=arn33423342;")
    Public WithEvents backworker2 As New BackgroundWorker

    Private Sub WindowsFormsHost1_GotFocus(sender As Object, e As RoutedEventArgs) Handles newwindows1.GotFocus
        If txtserver.Text = "Server" Then
            txtserver.Text = ""
        End If
        txtserver.ForeColor = System.Drawing.Color.Black
    End Sub
    Private Sub WindowsFormsHost1_LostFocus(sender As Object, e As RoutedEventArgs) Handles newwindows1.LostFocus
        If txtserver.Text = "" Then
            txtserver.Text = "Server"
        End If
        txtserver.ForeColor = System.Drawing.ColorTranslator.FromHtml("#8a8b8c")
    End Sub
    Private Sub server_GotFocus(sender As Object, e As RoutedEventArgs) Handles newwindows2.GotFocus
        If txtserver1.Text = "Server" Then
            txtserver1.Text = ""
        End If
        txtserver1.ForeColor = System.Drawing.Color.Black
    End Sub
    Private Sub server_LostFocus(sender As Object, e As RoutedEventArgs) Handles newwindows2.LostFocus
        If txtserver1.Text = "" Then
            txtserver1.Text = "Server"
        End If
        txtserver1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#8a8b8c")
    End Sub
    Private Sub Newport_GotFocus(sender As Object, e As RoutedEventArgs) Handles newport1.GotFocus
        If txtport.Text = "Port" Then
            txtport.Text = ""
        End If
        txtport.ForeColor = System.Drawing.Color.Black
    End Sub
    Private Sub Newport_LostFocus(sender As Object, e As RoutedEventArgs) Handles newport1.LostFocus
        If txtport.Text = "" Then
            txtport.Text = "Port"
        End If
        txtport.ForeColor = System.Drawing.ColorTranslator.FromHtml("#8a8b8c")
    End Sub
    Private Sub Newport1_GotFocus(sender As Object, e As RoutedEventArgs) Handles newport2.GotFocus
        If txtport1.Text = "Port" Then
            txtport1.Text = ""
        End If
        txtport1.ForeColor = System.Drawing.Color.Black
    End Sub
    Private Sub Newport1_LostFocus(sender As Object, e As RoutedEventArgs) Handles newport2.LostFocus
        If txtport1.Text = "" Then
            txtport1.Text = "Port"
        End If
        txtport1.ForeColor = ColorTranslator.FromHtml("#8a8b8c")
    End Sub
    Private Sub SetUp_mail_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        grid1.Opacity = 0
        tmr2.Interval = 350
        With txtserver
            .ForeColor = ColorTranslator.FromHtml("#8a8b8c") : .Font = New Font("Segoe UI SemiLight", 13) : .Text = "Server"
        End With
        With txtserver1
            .ForeColor = ColorTranslator.FromHtml("#8a8b8c") : .Font = New Font("Segoe UI SemiLight", 13) : .Text = "Server"
        End With
        With txtport
            .ForeColor = ColorTranslator.FromHtml("#8a8b8c") : .Font = New Font("Segoe UI SemiLight", 13) : .Text = "Port"
        End With
        With txtport1
            .ForeColor = ColorTranslator.FromHtml("#8a8b8c") : .Font = New Font("Segoe UI SemiLight", 13) : .Text = "Port"
        End With
    End Sub
    Private Sub txtport_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtport.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub txtport1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtport1.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub back_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles back.MouseDown
        NavigationService.Navigate(New LogIn)
    End Sub
    Private Sub next2_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles next2.MouseDown
        With My.Settings
            .txtserver = txtserver.Text
            .txtserver1 = txtserver1.Text
            .txtport = txtport.Text
            .txtport1 = txtport1.Text
            .ssl1 = ssl1.Text
            .ssl2 = ssl2.Text
        End With
        NavigationService.Navigate(New warning)
    End Sub
End Class

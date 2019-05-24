Imports System.Data
Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media.Animation

Public Class SaveContact
    Inherits System.Windows.Controls.Control
    Private Options As System.Windows.Controls.Image
    Private ContactOptions As ContextMenu
    Public EmailFrom As String
    Public From2 As TextBlock
    Dim saveCon As New SaveNewContact
    Private MainGrid As Grid

    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(SaveContact), New FrameworkPropertyMetadata(GetType(SaveContact)))
    End Sub
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        Options = GetTemplateChild("Options")
        AddHandler Options.MouseDown, AddressOf Options_MouseDown

        ContactOptions = GetTemplateChild("ContactOptions")
        MainGrid = GetTemplateChild("MainGrid")

        From2 = GetTemplateChild("From2")
    End Sub
    Private Sub Options_MouseDown(sender As Object, e As MouseButtonEventArgs)
        Dim cm = ContextMenuService.GetContextMenu(TryCast(MainGrid, DependencyObject))
        If cm Is Nothing Then
            Return
        End If
        cm.Placement = PlacementMode.Bottom
        cm.PlacementTarget = TryCast(MainGrid, UIElement)
        cm.IsOpen = True
        Dim template = ContactOptions.Template
        'GetCurrentUserContact()
    End Sub
    Private Sub SaveContact()
        Dim template = ContactOptions.Template
        Dim n = CType(template.FindName("Name", ContactOptions), TextBlock).Text
        Dim Name As String = n
        Dim s = CType(template.FindName("SecEmail", ContactOptions), TextBox).Text
        Dim SecMail As String = s
        Dim a = CType(template.FindName("Address", ContactOptions), TextBox).Text
        Dim Address As String = a
        Dim p = CType(template.FindName("Phone", ContactOptions), TextBox).Text
        Dim Phone As String = p
        saveCon.SaveContact(Name, EmailFrom, SecMail, Address, "", Phone)
    End Sub
    Private Sub GetCurrentUserContact()
        Dim template = ContactOptions.Template
        Dim SenderAddress As Match = Regex.Match(EmailFrom, "<(.*?)>")
        Dim eFrom = SenderAddress.Groups(1).Value.ToString
        Dim cmd As New OleDbCommand("Select * from contacts WHERE Email='" + eFrom + "'", Home.contact)
        Dim table As New DataTable
        Dim adapter As New OleDbDataAdapter(cmd)
        adapter.Fill(table)
        Dim Name = CType(template.FindName("Name", ContactOptions), TextBox)
        Dim secmail = CType(template.FindName("SecEmail", ContactOptions), TextBox)
        Dim address = CType(template.FindName("Address", ContactOptions), TextBox)
        If Not table.Rows.Count <= 0 Then
            Name.Text = table.Rows(0)(0).ToString
            secmail.Text = table.Rows(0)(2).ToString
            address.Text = table.Rows(0)(3).ToString
        Else
            Name.Text = ""
            secmail.Text = ""
            address.Text = ""
        End If
    End Sub
End Class

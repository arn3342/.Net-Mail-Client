Imports System.Data
Imports System.Data.OleDb
Imports System.IO.Directory
Public Class SaveNewContact
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetCurrentDirectory() & "\Data\" + My.Settings.email + "\_Cons.ttmdata;Jet OLEDB:Database Password=arn33423342;")
    Public Sub SaveContact(ByVal name As String, ByVal Email As String, ByVal SecondaryEmail As String, ByVal Address As String, ByVal LastContacted As String, ByVal Phone As String)
        Dim Exists = CheckIfContactExists(Email)
        If Exists = False Then
            AddNew(name, Email, SecondaryEmail, Address, LastContacted, Phone)
        Else
            UpdateExisting(name, Email, SecondaryEmail, Address, LastContacted, Phone)
        End If
    End Sub
    Public Function CheckIfContactExists(ByVal Email As String)
        Dim Exists As Boolean
        con.Open()
        Dim cmd As New OleDbCommand("Select * from contacts WHERE Email='" + Email + "'", con)
        Dim adapter As New OleDbDataAdapter(cmd)
        Dim table As New DataTable
        adapter.Fill(table)
        If table.Rows.Count <= 0 Then
            Exists = False
        Else
            Exists = True
        End If
        Return Exists
    End Function

    Public Sub AddNew(ByVal name As String, ByVal Email As String, ByVal SecondaryEmail As String, ByVal Address As String, ByVal LastContacted As String, ByVal GroupName As String)
        Dim InsertData As New OleDbCommand("Insert into contacts(UName,Email,SecondaryEmail,Address,Lastcon,Groups)values(@name,@Email,@secmail,@address,@lastcon,@groups)", con)
        InsertData.Parameters.Add("@name", OleDbType.VarChar).Value = name
        InsertData.Parameters.Add("@Email", OleDbType.VarChar).Value = Email
        InsertData.Parameters.Add("@secmail", OleDbType.VarChar).Value = SecondaryEmail
        InsertData.Parameters.Add("@address", OleDbType.VarChar).Value = Address
        InsertData.Parameters.Add("@lastcon", OleDbType.VarChar).Value = LastContacted
        InsertData.Parameters.Add("@groups", OleDbType.VarChar).Value = GroupName
        InsertData.ExecuteNonQuery()
        con.Close()
    End Sub
    Public Sub UpdateExisting(ByVal name As String, ByVal Email As String, ByVal SecondaryEmail As String, ByVal Address As String, ByVal LastContacted As String, ByVal GroupName As String)
        Dim UpdateData As New OleDbCommand("UPDATE contacts SET UName=@name,SecondaryEmail=@secmail,Address=@address,Lastcon=@lastcon,Groups=@groups", con)
        UpdateData.Parameters.Add("@name", OleDbType.VarChar).Value = name
        UpdateData.Parameters.Add("@secmail", OleDbType.VarChar).Value = SecondaryEmail
        UpdateData.Parameters.Add("@address", OleDbType.VarChar).Value = Address
        UpdateData.Parameters.Add("@lastcon", OleDbType.VarChar).Value = LastContacted
        UpdateData.Parameters.Add("@groups", OleDbType.VarChar).Value = GroupName
        UpdateData.ExecuteNonQuery()
        con.Close()
    End Sub
End Class

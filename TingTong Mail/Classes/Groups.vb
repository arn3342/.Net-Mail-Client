Imports System.Data.OleDb
Imports System.IO.Directory
Imports System.ComponentModel

Public Class Groups
    Public AllGroups As New List(Of GetGroupNames)
    Dim GroupNamesOnly As New List(Of String)
    Dim WithEvents bgw As New System.ComponentModel.BackgroundWorker
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetCurrentDirectory() & "\Data\" + My.Settings.email + "\_Cons.ttmdata;Jet OLEDB:Database Password=arn33423342;")
    Public Sub Initialize()
        My.Settings.GroupsLoaded = False
        bgw.RunWorkerAsync()
    End Sub

    Private Sub bgw_DoWork(sender As Object, f As DoWorkEventArgs) Handles bgw.DoWork
        con.Open()
        AllGroups.Clear()
        Dim cmd As New OleDbCommand("Select * from contacts", con)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        While dr.Read
            If Not GroupNamesOnly.Exists(Function(e) e.EndsWith(dr(5).ToString)) Then
                GroupNamesOnly.Add(dr(5).ToString)
            End If
        End While
        For Each GrpName In GroupNamesOnly
            Dim cmd2 As New OleDbCommand("Select * from contacts WHERE Groups='" + GrpName + "'", con)
            Dim reader As OleDbDataReader = cmd2.ExecuteReader
            Dim GroupsWidMembers As New GetGroupNames
            GroupsWidMembers.Name = GrpName
            While reader.Read
                If reader(0).ToString = "" Then
                    GroupsWidMembers.Member = reader(1).ToString + ","
                Else
                    GroupsWidMembers.Member = reader(0).ToString + ","
                End If
            End While
            AllGroups.Add(GroupsWidMembers)
        Next
        con.Close()
    End Sub

    Private Sub bgw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgw.RunWorkerCompleted
        My.Settings.GroupsLoaded = True
    End Sub
End Class

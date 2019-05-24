Imports System.Windows.Media.Animation
Public Class MyCustomCheckBoxxaml
    Public Checked As Boolean
    Private Sub MyCustomCheckBoxxaml_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseDown
        If Checked = False Then
            Checked = True
            Dim sb As New Storyboard
            sb = FindResource("Checked")
            Storyboard.SetTarget(sb, Toggle)
            sb.Begin()
        Else
            Checked = False
            Dim sb As New Storyboard
            sb = FindResource("Unchecked")
            Storyboard.SetTarget(sb, Toggle)
            sb.Begin()
        End If
    End Sub
End Class

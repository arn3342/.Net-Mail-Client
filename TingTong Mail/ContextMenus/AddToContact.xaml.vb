Class AddToContact
    Dim bb As New BrushConverter

    Private Sub AddName_KeyDown(sender As Object, e As KeyEventArgs) Handles AddName.KeyDown
        If e.Key = Key.Return Then
            If Not AddName.Text = "" Then
                AddName.Visibility = Visibility.Hidden
                UserName.Text = AddName.Text
            End If
        End If
    End Sub

    Private Sub Address_GotFocus(sender As Object, e As RoutedEventArgs) Handles Address.GotFocus
        With Address
            If .Text = "Address" Then
                .Text = ""
                .Foreground = Brushes.Black
            End If
        End With
    End Sub

    Private Sub Address_LostFocus(sender As Object, e As RoutedEventArgs) Handles Address.LostFocus
        With Address
            If .Text = "" Then
                .Text = "Address"
                .Foreground = bb.ConvertFrom("#FFB0B0B0")
            End If
        End With
    End Sub

    Private Sub SecEmail_GotFocus(sender As Object, e As RoutedEventArgs) Handles SecEmail.GotFocus
        With SecEmail
            If .Text = "Secondary email" Then
                .Text = ""
                .Foreground = Brushes.Black
            End If
        End With
    End Sub

    Private Sub SecEmail_LostFocus(sender As Object, e As RoutedEventArgs) Handles SecEmail.LostFocus
        With SecEmail
            If .Text = "" Then
                .Text = "Secondary email"
                .Foreground = bb.ConvertFrom("#FFB0B0B0")
            End If
        End With
    End Sub
    Private Sub UserName_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles UserName.MouseDown
        If e.ClickCount = 2 Then
            AddName.Visibility = Visibility.Visible
            FocusManager.SetFocusedElement(Me, AddName)
        End If
    End Sub
End Class

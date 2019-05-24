Public Class Window2
    Private Sub textBox_KeyDown(sender As Object, e As KeyEventArgs) Handles textBox.KeyDown

    End Sub

    Private Sub textBox_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles textBox.PreviewKeyDown
        If e.Key = Key.Space Then
            If lst1.Visibility = Visibility.Hidden Then
                lst1.Visibility = Visibility.Visible
            Else
                lst1.Visibility = Visibility.Hidden
            End If
        End If
        If lst1.Visibility = Visibility.Visible Then
            If e.Key = Key.Down Then
                lst1.SelectedIndex = lst1.SelectedIndex + 1
                textBlock1.Text = "Downkey pressed"
            End If
            If e.Key = Key.Up Then
                lst1.SelectedIndex = lst1.SelectedIndex - 1
                textBlock1.Text = "Upkey pressed"
            End If
        End If
    End Sub

    Private Sub textBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles textBox.TextChanged
        If textBox.IsFocused = True Then
            If lst1.Visibility = Visibility.Hidden Then
                lst1.Visibility = Visibility.Visible
            End If
        End If
    End Sub
End Class

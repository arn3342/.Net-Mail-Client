Imports System.Drawing.Text
Imports System.Windows.Controls.Primitives

Class SendMail
    Dim bb As New BrushConverter
    Dim CurrentAlignMent As String
    Public Shared BodyBox As New WordSuggestionsTextBoxLib.WordSuggestionsTextBox
    Private Sub EachFontSize_MouseDown()

    End Sub
    Private Sub FontSizeSelector_MouseDown(sender As Object, e As MouseButtonEventArgs)

    End Sub
    Private Sub SendMail_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        MainBody.Children.Add(BodyBox)
        BodyBox.SuggestionsList = IO.File.ReadAllLines(IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "words.txt"))
        FontSelector.LoadFontNameOnly = True
        FontSizeSelector.LoadFontSizeOnly = True
        Formattings.Formatting = True
        WordSuggestion.DisplayText.Text = "Word suggestion"
        AutoCorrect.DisplayText.Text = "Auto correct"
        HTMLorTEXT.DisplayText.Text = "HTML"
        BodyBox.SuggestionsEnabled = False
        BodyBox.AutomaticSpellCorrection = False
        FontSelector.MainTextBox.Text = System.Drawing.SystemFonts.DefaultFont.Name
        FontSizeSelector.MainTextBox.Text = 14
        BodyBox.UnderlyingRichTextBox.FontSize = 14
        BodyBox.UnderlyingRichTextBox.FontFamily = New FontFamily(System.Drawing.SystemFonts.DefaultFont.Name)
        QuickAdd.MainTextBox.Text = "Quick add"
    End Sub
    Private Sub AlignToCenter_MouseLeave(sender As Object, e As MouseEventArgs) Handles AlignToCenter.MouseLeave
        If CurrentAlignMent = AlignToCenter.Name Then
            e.Handled = True
        End If
    End Sub
    Private Sub AlignToCenter_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles AlignToCenter.MouseDown
        CurrentAlignMent = AlignToCenter.Name
        AlignToCenter.Opacity = 1
    End Sub

    Private Sub AlignToLeft_MouseLeave(sender As Object, e As MouseEventArgs) Handles AlignToLeft.MouseLeave
        If CurrentAlignMent = AlignToLeft.Name Then
            e.Handled = True
        End If
    End Sub
    Private Sub AlignToLeft_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles AlignToLeft.MouseDown
        CurrentAlignMent = AlignToLeft.Name
        AlignToLeft.Opacity = 1
    End Sub
    Private Sub AlignToRight_MouseLeave(sender As Object, e As MouseEventArgs) Handles AlignToRight.MouseLeave
        If CurrentAlignMent = AlignToRight.Name Then
            e.Handled = True
        End If
    End Sub
    Private Sub AlignToRight_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles AlignToRight.MouseDown
        CurrentAlignMent = AlignToRight.Name
        AlignToRight.Opacity = 1
    End Sub

    Private Sub WordSuggestion_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles WordSuggestion.MouseDown
        If WordSuggestion.Checked = True Then
            BodyBox.SuggestionsEnabled = True
        Else
            BodyBox.SuggestionsEnabled = False
        End If
    End Sub

    Private Sub AutoCorrect_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles AutoCorrect.MouseDown
        If AutoCorrect.Checked = True Then
            BodyBox.AutomaticSpellCorrection = True
        Else
            BodyBox.AutomaticSpellCorrection = False
        End If
    End Sub
    Public Shared Sub ChangeTextProperty(ByVal dp As DependencyProperty, ByVal value As String)
        If BodyBox.UnderlyingRichTextBox Is Nothing Then Return
        Dim ts As TextSelection = BodyBox.UnderlyingRichTextBox.Selection
        If ts IsNot Nothing Then ts.ApplyPropertyValue(dp, value)
        BodyBox.UnderlyingRichTextBox.Focus()
    End Sub
End Class

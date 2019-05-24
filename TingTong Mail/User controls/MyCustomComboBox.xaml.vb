
Imports System.Drawing.Text
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media.Animation

Public Class MyCustomComboBox
    Dim FontSize As New List(Of String)
    Dim Fonts As New List(Of UserFonts)
    Public LoadFontSizeOnly As Boolean
    Public LoadFontNameOnly As Boolean
    Public Formatting As Boolean

    Private Sub Main_LostFocus(sender As Object, e As RoutedEventArgs) Handles Main.LostFocus
        Me.Height = 29
    End Sub
    Private Sub ExpanderBtn_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles ExpanderBtn.MouseDown
        Dim cm = ContextMenuService.GetContextMenu(TryCast(Me, DependencyObject))
        If cm Is Nothing Then
            Return
        End If
        cm.Placement = PlacementMode.Bottom
        cm.PlacementTarget = TryCast(Me, UIElement)
        cm.IsOpen = True
    End Sub
    Private Sub LoadAllFontSize()
        FontSize.Add("8")
        FontSize.Add("9")
        FontSize.Add("10")
        FontSize.Add("12")
        FontSize.Add("14")
        FontSize.Add("16")
        FontSize.Add("18")
        FontSize.Add("20")
        FontSize.Add("22")
        FontSize.Add("24")
        FontSize.Add("26")
        FontSize.Add("28")
        FontSize.Add("36")
        FontSize.Add("48")
        FontSize.Add("72")
    End Sub
    Public Sub LoadFontSize()
        Dim template = List.Template
        Dim ListOfFontsSizes = CType(template.FindName("ItemHolder", List), VirtualizingStackPanel)
        If ListOfFontsSizes.Children.Count = 0 Then
            For Each fnt In FontSize
                Dim fontBtn As New UserFonts
                fontBtn.FontName.Text = fnt
                AddHandler fontBtn.MouseDown, AddressOf EachFontSize_MouseDown
                ListOfFontsSizes.Height = ListOfFontsSizes.Height + 24
                ListOfFontsSizes.Children.Add(fontBtn)
            Next
        End If
        Dim ContextGrid = CType(template.FindName("ContextGrid", List), Grid)
        ContextGrid.Height = ContextGrid.Height + 12
    End Sub

    Private Sub LoadFonts()
        Dim fontCol As InstalledFontCollection = New InstalledFontCollection()
        For x As Integer = 0 To fontCol.Families.Length - 1
            Dim EachFont As New UserFonts
            EachFont.FontName.Text = fontCol.Families(x).Name
            Fonts.Add(EachFont)
        Next
        Dim template = List.Template
        Dim ListOfFonts = CType(template.FindName("ItemHolder", List), VirtualizingStackPanel)
        If ListOfFonts.Children.Count = 0 Then
            For Each fnt In Fonts
                AddHandler fnt.MouseDown, AddressOf EachFont_MouseDown
                ListOfFonts.Height = ListOfFonts.Height + 24
                ListOfFonts.Children.Add(fnt)
            Next
        End If
    End Sub
    Private Sub EachFont_MouseDown(sender As Object, e As MouseButtonEventArgs)
        Dim SelectedFont = DirectCast(sender, UserFonts)
        MainTextBox.Text = SelectedFont.FontName.Text
        List.IsOpen = False
        Dim richTexta = New TextRange(SendMail.BodyBox.UnderlyingRichTextBox.Document.ContentStart, SendMail.BodyBox.UnderlyingRichTextBox.Document.ContentEnd).Text
        SendMail.BodyBox.UnderlyingRichTextBox.Focus()
        If richTexta.Count = 0 Then
            SendMail.BodyBox.UnderlyingRichTextBox.FontFamily = New FontFamily(MainTextBox.Text)
        Else
            SendMail.ChangeTextProperty(FontFamilyProperty, MainTextBox.Text)
        End If
    End Sub
    Private Sub EachFontSize_MouseDown(sender As Object, e As MouseButtonEventArgs)
        Dim SelectedFont = DirectCast(sender, UserFonts)
        MainTextBox.Text = SelectedFont.FontName.Text
        List.IsOpen = False
        Dim richTexta = New TextRange(SendMail.BodyBox.UnderlyingRichTextBox.Document.ContentStart, SendMail.BodyBox.UnderlyingRichTextBox.Document.ContentEnd).Text
        SendMail.BodyBox.UnderlyingRichTextBox.Focus()
        If richTexta.Length = 0 Then
            SendMail.BodyBox.UnderlyingRichTextBox.FontSize = MainTextBox.Text
        Else
            SendMail.ChangeTextProperty(FontSizeProperty, MainTextBox.Text)
        End If
    End Sub

    Private Sub MyCustomComboBox_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If LoadFontSizeOnly = True Then
            LoadAllFontSize()
        End If
    End Sub

    Private Sub List_Loaded(sender As Object, e As RoutedEventArgs) Handles List.Loaded
        If LoadFontNameOnly = True Then
            LoadFonts()
        End If
        If LoadFontSizeOnly = True Then
            LoadFontSize()
        End If
        If Formatting = True Then
            Dim template = List.Template
            Dim ConGrid = CType(template.FindName("ConGrid", List), Grid)
            ConGrid.Width = Me.Width + 7
            Dim OpHolder = CType(template.FindName("FormattingHolder", List), StackPanel)
            ConGrid.Height = 28
            Dim clrFrmtng As New UserFonts
            clrFrmtng.FontName.Text = "Clear formatting"
            OpHolder.Children.Add(clrFrmtng)
            OpHolder.Height = 25
        End If
    End Sub
End Class

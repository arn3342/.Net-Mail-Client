Imports System.Windows.Controls.Primitives
Public Class CustomCheckbox
    Inherits System.Windows.Controls.Control
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(CustomCheckbox), New FrameworkPropertyMetadata(GetType(CustomCheckbox)))
    End Sub
    Dim CheckPath As Path
    Public Checked As Boolean
    Private MainRect As Rectangle
    Private CheckBoxGrid As Grid
    Dim bb As New BrushConverter
    Dim myResourceDictionary As New ResourceDictionary
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        MainRect = GetTemplateChild("MainRect")
        CheckBoxGrid = GetTemplateChild("CheckBoxGrid")
        myResourceDictionary.Source = New Uri("pack://application:,,,/TingTong Mail;component/Themes/Generic.xaml", UriKind.RelativeOrAbsolute)
        CheckPath = New Path With {.Style = myResourceDictionary("CstmPathStyle")}
    End Sub
    Private Sub CustomCheckbox_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        If Checked = False Then
            Checked = True
            CheckBoxGrid.Children.Add(CheckPath)
            MainRect.Fill = bb.ConvertFrom("#FF008BFF")
        Else
            CheckBoxGrid.Children.Remove(CheckPath)
            Checked = False
            MainRect.Fill = bb.ConvertFrom("#FFC3C3C3")
        End If
    End Sub
End Class

' Follow steps 1a or 1b and then 2 to use this custom control in a XAML file.
'
' Step 1a) Using this custom control in a XAML file that exists in the current project.
' Add this XmlNamespace attribute to the root element of the markup file where it is 
' to be used:
'
'     xmlns:MyNamespace="clr-namespace:TingTong_Mail"
'
'
' Step 1b) Using this custom control in a XAML file that exists in a different project.
' Add this XmlNamespace attribute to the root element of the markup file where it is 
' to be used:
'
'     xmlns:MyNamespace="clr-namespace:TingTong_Mail;assembly=TingTong_Mail"
'
' You will also need to add a project reference from the project where the XAML file lives
' to this project and Rebuild to avoid compilation errors:
'
'     Right click on the target project in the Solution Explorer and
'     "Add Reference"->"Projects"->[Browse to and select this project]
'
'
' Step 2)
' Go ahead and use your control in the XAML file. Note that Intellisense in the
' XML editor does not currently work on custom controls and its child elements.
'
'     <MyNamespace:CalenderControl/>
'

Imports System.Windows.Controls.Primitives
Imports System.Windows.Media.Animation

Public Class CalenderControl
    Inherits System.Windows.Controls.Control
    Dim myResourceDictionary As New ResourceDictionary
    Dim MainBorder As Border
    Shared Sub New()
        'This OverrideMetadata call tells the system that this element wants to provide a style that is different than its base class.
        'This style is defined in themes\generic.xaml
        DefaultStyleKeyProperty.OverrideMetadata(GetType(CalenderControl), New FrameworkPropertyMetadata(GetType(CalenderControl)))
    End Sub
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        MainBorder = GetTemplateChild("MainBorder")
        myResourceDictionary.Source = New Uri("pack://application:,,,/TingTong Mail;component/Themes/Generic.xaml", UriKind.RelativeOrAbsolute)
    End Sub
    Private Sub CalenderControl_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim sb As New Storyboard
        sb = myResourceDictionary("BringUp")
        Storyboard.SetTarget(sb, MainBorder)
        sb.Begin()
    End Sub
End Class

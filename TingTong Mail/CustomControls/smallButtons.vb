Public Class SmallButtons
    Inherits Control
    Public img As Image
    Private MainGrid As Grid
    Dim MainRectangle As Rectangle
    Dim TextToolTip As TextBlock
    Public Shared ReadOnly ImgHeightDependency As DependencyProperty = DependencyProperty.Register("ImageHeight", GetType(Double), GetType(SmallButtons))
    Public Shared ReadOnly ImgWidthDependency As DependencyProperty = DependencyProperty.Register("ImageWidth", GetType(Double), GetType(SmallButtons))
    Public Shared ReadOnly ImgOpacityDependency As DependencyProperty = DependencyProperty.Register("ImageOpacity", GetType(Double), GetType(SmallButtons))
    Public Shared ReadOnly ImgMOOpacityDependency As DependencyProperty = DependencyProperty.Register("ImageMOOpacity", GetType(Double), GetType(SmallButtons))
    Public Shared ReadOnly ImgMLOpacityDependency As DependencyProperty = DependencyProperty.Register("ImageMLOpacity", GetType(Double), GetType(SmallButtons))
    Public Shared ReadOnly ImgSourceDependency As DependencyProperty = DependencyProperty.Register("ImageSource", GetType(ImageSource), GetType(SmallButtons))
    Public Shared ReadOnly ToolTipTextDependency As DependencyProperty = DependencyProperty.Register("ToolTipText", GetType(String), GetType(SmallButtons))
    Public Shared ReadOnly ImgHrAlignmentDependency As DependencyProperty = DependencyProperty.Register("ImageHorizontalAlignment", GetType(HorizontalAlignment), GetType(SmallButtons))
    Public Shared ReadOnly RectXRadiusDependency As DependencyProperty = DependencyProperty.Register("RadiuxX", GetType(Double), GetType(SmallButtons))
    Public Shared ReadOnly RectYRadiusDependency As DependencyProperty = DependencyProperty.Register("RadiuxY", GetType(Double), GetType(SmallButtons))
    Public Shared ReadOnly MouseEnterColorDependency As DependencyProperty = DependencyProperty.Register("MouseEnter Color", GetType(String), GetType(SmallButtons))
    Public Shared ReadOnly AllowToolTipDependency As DependencyProperty = DependencyProperty.Register("Allow ToolTip", GetType(Boolean), GetType(SmallButtons))
    Private bb As New BrushConverter
    Public Property ImgHorizontalAlignment As HorizontalAlignment
        Get
            Return CType(GetValue(ImgHrAlignmentDependency), HorizontalAlignment)
        End Get
        Set(ByVal value As HorizontalAlignment)
            SetValue(ImgHrAlignmentDependency, value)
        End Set
    End Property
    Public Property DisallowDefaultToolTip As Boolean
        Get
            Return Convert.ToBoolean(GetValue(AllowToolTipDependency))
        End Get
        Set(ByVal value As Boolean)
            SetValue(AllowToolTipDependency, value)
        End Set
    End Property
    Public Property MouseEnterColor As String
        Get
            Return CStr(GetValue(MouseEnterColorDependency))
        End Get
        Set(ByVal value As String)
            SetValue(MouseEnterColorDependency, value)
        End Set
    End Property
    Public Property ToolTipText As String
        Get
            Return CStr(GetValue(ToolTipTextDependency))
        End Get
        Set(ByVal value As String)
            SetValue(ToolTipTextDependency, value)
        End Set
    End Property
    Public Property RadiusX As Double
        Get
            Return Convert.ToDouble(GetValue(RectXRadiusDependency))
        End Get
        Set(ByVal value As Double)
            SetValue(RectXRadiusDependency, value)
        End Set
    End Property
    Public Property ImgOpacity As Double
        Get
            Return Convert.ToDouble(GetValue(ImgOpacityDependency))
        End Get
        Set(ByVal value As Double)
            SetValue(ImgOpacityDependency, value)
        End Set
    End Property
    Public Property ImgMOOpacity As Double
        Get
            Return Convert.ToDouble(GetValue(ImgMOOpacityDependency))
        End Get
        Set(ByVal value As Double)
            SetValue(ImgMOOpacityDependency, value)
        End Set
    End Property
    Public Property ImgMLOpacity As Double
        Get
            Return Convert.ToDouble(GetValue(ImgMLOpacityDependency))
        End Get
        Set(ByVal value As Double)
            SetValue(ImgMLOpacityDependency, value)
        End Set
    End Property
    Public Property RadiusY As Double
        Get
            Return Convert.ToDouble(GetValue(RectYRadiusDependency))
        End Get
        Set(ByVal value As Double)
            SetValue(RectYRadiusDependency, value)
        End Set
    End Property

    Public Property ImgHeight As Double
        Get
            Return Convert.ToDouble(GetValue(ImgHeightDependency))
        End Get
        Set(ByVal value As Double)
            SetValue(ImgHeightDependency, value)
        End Set
    End Property

    Public Property ImgWidth As Double
        Get
            Return Convert.ToDouble(GetValue(ImgWidthDependency))
        End Get
        Set(ByVal value As Double)
            SetValue(ImgWidthDependency, value)
        End Set
    End Property
    Public Property ImgSource As ImageSource
        Get
            Return CType(GetValue(ImgSourceDependency), ImageSource)
        End Get
        Set(ByVal value As ImageSource)
            SetValue(ImgSourceDependency, value)
        End Set
    End Property
    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(SmallButtons), New FrameworkPropertyMetadata(GetType(SmallButtons)))
    End Sub
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()

        MainGrid = GetTemplateChild("MainGrid")
        AddHandler MainGrid.MouseEnter, AddressOf MainGrid_MouseEnter
        AddHandler MainGrid.MouseLeave, AddressOf MainGrid_MouseLeave

        MainRectangle = GetTemplateChild("MainRectangle")
        Dim RectangleXRadiusBinding As Binding = New Binding("RadiusX") With {
            .Source = Me,
            .Mode = BindingMode.TwoWay
        }
        MainRectangle.SetBinding(Rectangle.RadiusXProperty, RectangleXRadiusBinding)
        Dim RectangleYRadiusBinding As Binding = New Binding("RadiusY") With {
            .Source = Me,
            .Mode = BindingMode.TwoWay
        }
        MainRectangle.SetBinding(Rectangle.RadiusYProperty, RectangleYRadiusBinding)
        AddHandler MainRectangle.MouseEnter, AddressOf MainRectangle_MouseEnter
        If DisallowDefaultToolTip = False Then
            TextToolTip = GetTemplateChild("ToolTip")
            Dim ToolTipTextBinding As Binding = New Binding("ToolTipText") With {
            .Source = Me,
            .Mode = BindingMode.TwoWay
        }
            TextToolTip.SetBinding(TextBlock.TextProperty, ToolTipTextBinding)
        End If
        img = TryCast(GetTemplateChild("Image"), System.Windows.Controls.Image)
        Dim ImageHeightBinding As Binding = New Binding("ImageHeight") With {
            .Source = Me,
            .Mode = BindingMode.TwoWay
        }
        img.SetBinding(Image.HeightProperty, ImageHeightBinding)

        Dim ImageWidthBinding As Binding = New Binding("ImageWidth") With {
            .Source = Me,
            .Mode = BindingMode.TwoWay
        }
        img.SetBinding(Image.WidthProperty, ImageWidthBinding)

        Dim ImageSourceBinding As Binding = New Binding("ImageSource") With {
            .Source = Me,
            .Mode = BindingMode.TwoWay
        }
        img.SetBinding(Image.SourceProperty, ImageSourceBinding)

        Dim ImageHorizontalAlignmentBinding As Binding = New Binding("ImageHorizontalAlignment") With {
           .Source = Me,
           .Mode = BindingMode.TwoWay
        }
        img.SetBinding(Image.HorizontalAlignmentProperty, ImageHorizontalAlignmentBinding)
        ImgHorizontalAlignment = HorizontalAlignment.Stretch

        Dim ImageDefaultOpacityHorizontalAlignmentBinding As Binding = New Binding("ImageOpacity") With {
          .Source = Me,
          .Mode = BindingMode.TwoWay
       }
        img.SetBinding(Image.OpacityProperty, ImageDefaultOpacityHorizontalAlignmentBinding)
        ImgMOOpacity = 0.6

    End Sub
    Private Sub MainGrid_MouseEnter()
        ImgOpacity = ImgMOOpacity
    End Sub
    Private Sub MainGrid_MouseLeave()
        ImgOpacity = ImgMLOpacity
    End Sub
    Private Sub MainRectangle_MouseEnter(sender As Object, e As MouseEventArgs)
        Dim Rect As Rectangle = CType(sender, Rectangle)
        Rect.Fill = bb.ConvertFrom(MouseEnterColor)
    End Sub
End Class

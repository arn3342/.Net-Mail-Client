
Public Class ImageLoader
    Public Function LoadImage(ByVal uri As Uri, ByVal Pixel As Integer) As BitmapImage
        Dim bi As New BitmapImage()
        bi.BeginInit()
        bi.CacheOption = BitmapCacheOption.OnLoad
        bi.UriSource = uri
        bi.DecodePixelWidth = Pixel
        bi.EndInit()
        bi.Freeze()
        Return bi
    End Function
End Class

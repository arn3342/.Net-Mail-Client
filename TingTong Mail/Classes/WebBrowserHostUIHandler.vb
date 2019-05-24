Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports TingTong_Mail

Public Class WebBrowserHostUIHandler
    Implements Native.IDocHostUIHandler

    Private Const E_NOTIMPL As UInteger = 2147500033UI

    Private Const S_OK As UInteger = 0

    Private Const S_FALSE As UInteger = 1

    Public Sub New(ByVal browser As System.Windows.Controls.WebBrowser)
        If browser Is Nothing Then Throw New ArgumentNullException("browser")
        browser = browser
        AddHandler browser.LoadCompleted, AddressOf OnLoadCompleted
        AddHandler browser.Navigated, AddressOf OnNavigated
        IsWebBrowserContextMenuEnabled = True
        Flags = Flags Or HostUIFlags.ENABLE_REDIRECT_NOTIFICATION
    End Sub

    Public Property Browser As WebBrowser

    Public Property Flags As HostUIFlags

    Public Property IsWebBrowserContextMenuEnabled As Boolean

    Public Property ScriptErrorsSuppressed As Boolean

    Private Sub OnNavigated(ByVal sender As Object, ByVal e As NavigationEventArgs)
        SetSilent(Browser, ScriptErrorsSuppressed)
    End Sub

    Private Sub OnLoadCompleted(ByVal sender As Object, ByVal e As NavigationEventArgs)
        Dim doc As Native.ICustomDoc = TryCast(Browser.Document, Native.ICustomDoc)
        If doc IsNot Nothing Then
            doc.SetUIHandler(Me)
        End If
    End Sub

    Private Function ShowContextMenu(ByVal dwID As Integer, ByVal pt As Native.POINT, ByVal pcmdtReserved As Object, ByVal pdispReserved As Object) As UInteger
        Return If(IsWebBrowserContextMenuEnabled, S_FALSE, S_OK)
    End Function

    Private Function GetHostInfo(ByRef info As Native.DOCHOSTUIINFO) As UInteger
        info.dwFlags = CInt(Flags)
        info.dwDoubleClick = 0
        Return S_OK
    End Function

    Private Function ShowUI(ByVal dwID As Integer, ByVal activeObject As Object, ByVal commandTarget As Object, ByVal frame As Object, ByVal doc As Object) As UInteger
        Return E_NOTIMPL
    End Function

    Private Function HideUI() As UInteger
        Return E_NOTIMPL
    End Function

    Private Function UpdateUI() As UInteger
        Return E_NOTIMPL
    End Function

    Private Function EnableModeless(ByVal fEnable As Boolean) As UInteger
        Return E_NOTIMPL
    End Function

    Private Function OnDocWindowActivate(ByVal fActivate As Boolean) As UInteger
        Return E_NOTIMPL
    End Function

    Private Function OnFrameWindowActivate(ByVal fActivate As Boolean) As UInteger
        Return E_NOTIMPL
    End Function

    Private Function ResizeBorder(ByVal rect As Native.COMRECT, ByVal doc As Object, ByVal fFrameWindow As Boolean) As UInteger
        Return E_NOTIMPL
    End Function

    Private Function TranslateAccelerator(ByRef msg As System.Windows.Forms.Message, ByRef group As Guid, ByVal nCmdID As Integer) As UInteger
        Return S_FALSE
    End Function

    Private Function GetOptionKeyPath(ByVal pbstrKey As String(), ByVal dw As Integer) As UInteger
        Return E_NOTIMPL
    End Function

    Private Function GetDropTarget(ByVal pDropTarget As Object, <Out> ByRef ppDropTarget As Object) As UInteger
        ppDropTarget = Nothing
        Return E_NOTIMPL
    End Function

    Private Function GetExternal(<Out> ByRef ppDispatch As Object) As UInteger
        ppDispatch = Browser.ObjectForScripting
        Return S_OK
    End Function

    Private Function TranslateUrl(ByVal dwTranslate As Integer, ByVal strURLIn As String, <Out> ByRef pstrURLOut As String) As UInteger
        pstrURLOut = Nothing
        Return E_NOTIMPL
    End Function

    Private Function FilterDataObject(ByVal pDO As IDataObject, <Out> ByRef ppDORet As IDataObject) As UInteger
        ppDORet = Nothing
        Return E_NOTIMPL
    End Function

    Public Shared Sub SetSilent(ByVal browser As WebBrowser, ByVal silent As Boolean)
        Dim sp As Native.IOleServiceProvider = TryCast(browser.Document, Native.IOleServiceProvider)
        If sp IsNot Nothing Then
            Dim IID_IWebBrowserApp As Guid = New Guid("0002DF05-0000-0000-C000-000000000046")
            Dim IID_IWebBrowser2 As Guid = New Guid("D30C1661-CDAF-11d0-8A3E-00C04FC9E26E")
            Dim webBrowser As Object
            sp.QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, webBrowser)
            If webBrowser IsNot Nothing Then
                webBrowser.[GetType]().InvokeMember("Silent", BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.PutDispProperty, Nothing, webBrowser, New Object() {silent})
            End If
        End If
    End Sub

    Private Function IDocHostUIHandler_ShowContextMenu(dwID As Integer, pt As POINT, <MarshalAs(UnmanagedType.Interface)> pcmdtReserved As Object, <MarshalAs(UnmanagedType.Interface)> pdispReserved As Object) As UInteger Implements IDocHostUIHandler.ShowContextMenu
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_GetHostInfo(ByRef info As DOCHOSTUIINFO) As UInteger Implements IDocHostUIHandler.GetHostInfo
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_ShowUI(dwID As Integer, <MarshalAs(UnmanagedType.Interface)> activeObject As Object, <MarshalAs(UnmanagedType.Interface)> commandTarget As Object, <MarshalAs(UnmanagedType.Interface)> frame As Object, <MarshalAs(UnmanagedType.Interface)> doc As Object) As UInteger Implements IDocHostUIHandler.ShowUI
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_HideUI() As UInteger Implements IDocHostUIHandler.HideUI
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_UpdateUI() As UInteger Implements IDocHostUIHandler.UpdateUI
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_EnableModeless(fEnable As Boolean) As UInteger Implements IDocHostUIHandler.EnableModeless
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_OnDocWindowActivate(fActivate As Boolean) As UInteger Implements IDocHostUIHandler.OnDocWindowActivate
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_OnFrameWindowActivate(fActivate As Boolean) As UInteger Implements IDocHostUIHandler.OnFrameWindowActivate
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_ResizeBorder(rect As COMRECT, <MarshalAs(UnmanagedType.Interface)> doc As Object, fFrameWindow As Boolean) As UInteger Implements IDocHostUIHandler.ResizeBorder
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_TranslateAccelerator(ByRef msg As Message, ByRef group As Guid, nCmdID As Integer) As UInteger Implements IDocHostUIHandler.TranslateAccelerator
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_GetOptionKeyPath(<MarshalAs(UnmanagedType.LPArray)> <Out> pbstrKey() As String, dw As Integer) As UInteger Implements IDocHostUIHandler.GetOptionKeyPath
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_GetDropTarget(<[In]> <MarshalAs(UnmanagedType.Interface)> pDropTarget As Object, <MarshalAs(UnmanagedType.Interface)> <Out> ByRef ppDropTarget As Object) As UInteger Implements IDocHostUIHandler.GetDropTarget
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_GetExternal(<MarshalAs(UnmanagedType.IDispatch)> <Out> ByRef ppDispatch As Object) As UInteger Implements IDocHostUIHandler.GetExternal
        Throw New NotImplementedException()
    End Function

    Private Function IDocHostUIHandler_TranslateUrl(dwTranslate As Integer, <MarshalAs(UnmanagedType.LPWStr)> strURLIn As String, <MarshalAs(UnmanagedType.LPWStr)> <Out> ByRef pstrURLOut As String) As UInteger Implements IDocHostUIHandler.TranslateUrl
        Throw New NotImplementedException()
    End Function

    ' Private Function IDocHostUIHandler_FilterDataObject(pDO As System.Windows.IDataObject, <Out> ByRef ppDORet As System.Windows.IDataObject) As UInteger Implements IDocHostUIHandler.FilterDataObject
    'Throw New NotImplementedException()
    'End Function

    Private Function IDocHostUIHandler_FilterDataObject1(pDO As IDataObject, <Out> ByRef ppDORet As IDataObject) As UInteger Implements IDocHostUIHandler.FilterDataObject
        Throw New NotImplementedException()
    End Function
End Class

Friend Module Native

    <ComImport, Guid("BD3F23C0-D43E-11CF-893B-00AA00BDCE1A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Friend Interface IDocHostUIHandler

        <PreserveSig>
        Function ShowContextMenu(ByVal dwID As Integer, ByVal pt As POINT, <MarshalAs(UnmanagedType.[Interface])> ByVal pcmdtReserved As Object, <MarshalAs(UnmanagedType.[Interface])> ByVal pdispReserved As Object) As UInteger

        <PreserveSig>
        Function GetHostInfo(ByRef info As DOCHOSTUIINFO) As UInteger

        <PreserveSig>
        Function ShowUI(ByVal dwID As Integer, <MarshalAs(UnmanagedType.[Interface])> ByVal activeObject As Object, <MarshalAs(UnmanagedType.[Interface])> ByVal commandTarget As Object, <MarshalAs(UnmanagedType.[Interface])> ByVal frame As Object, <MarshalAs(UnmanagedType.[Interface])> ByVal doc As Object) As UInteger

        <PreserveSig>
        Function HideUI() As UInteger

        <PreserveSig>
        Function UpdateUI() As UInteger

        <PreserveSig>
        Function EnableModeless(ByVal fEnable As Boolean) As UInteger

        <PreserveSig>
        Function OnDocWindowActivate(ByVal fActivate As Boolean) As UInteger

        <PreserveSig>
        Function OnFrameWindowActivate(ByVal fActivate As Boolean) As UInteger

        <PreserveSig>
        Function ResizeBorder(ByVal rect As COMRECT, <MarshalAs(UnmanagedType.[Interface])> ByVal doc As Object, ByVal fFrameWindow As Boolean) As UInteger

        <PreserveSig>
        Function TranslateAccelerator(ByRef msg As System.Windows.Forms.Message, ByRef group As Guid, ByVal nCmdID As Integer) As UInteger

        <PreserveSig>
        Function GetOptionKeyPath(<Out, MarshalAs(UnmanagedType.LPArray)> ByVal pbstrKey As String(), ByVal dw As Integer) As UInteger

        <PreserveSig>
        Function GetDropTarget(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pDropTarget As Object, <Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppDropTarget As Object) As UInteger

        <PreserveSig>
        Function GetExternal(<Out> <MarshalAs(UnmanagedType.IDispatch)> ByRef ppDispatch As Object) As UInteger

        <PreserveSig>
        Function TranslateUrl(ByVal dwTranslate As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal strURLIn As String, <Out> <MarshalAs(UnmanagedType.LPWStr)> ByRef pstrURLOut As String) As UInteger

        <PreserveSig>
        Function FilterDataObject(ByVal pDO As IDataObject, <Out> ByRef ppDORet As IDataObject) As UInteger

    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Friend Structure DOCHOSTUIINFO

        Public cbSize As Integer

        Public dwFlags As Integer

        Public dwDoubleClick As Integer

        Public dwReserved1 As IntPtr

        Public dwReserved2 As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Friend Structure COMRECT

        Public left As Integer

        Public top As Integer

        Public right As Integer

        Public bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Friend Class POINT

        Public x As Integer

        Public y As Integer
    End Class

    <ComImport, Guid("3050F3F0-98B5-11CF-BB82-00AA00BDCE0B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Friend Interface ICustomDoc

        <PreserveSig>
        Function SetUIHandler(ByVal pUIHandler As IDocHostUIHandler) As Integer

    End Interface

    <ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Friend Interface IOleServiceProvider

        <PreserveSig>
        Function QueryService(<[In]> ByRef guidService As Guid, <[In]> ByRef riid As Guid, <Out> <MarshalAs(UnmanagedType.IDispatch)> ByRef ppvObject As Object) As UInteger

    End Interface
End Module

<Flags>
Public Enum HostUIFlags
    DIALOG = 1
    DISABLE_HELP_MENU = 2
    NO3DBORDER = 4
    SCROLL_NO = 8
    DISABLE_SCRIPT_INACTIVE = 16
    OPENNEWWIN = 32
    DISABLE_OFFSCREEN = 64
    FLAT_SCROLLBAR = 128
    DIV_BLOCKDEFAULT = 256
    ACTIVATE_CLIENTHIT_ONLY = 512
    OVERRIDEBEHAVIORFACTORY = 1024
    CODEPAGELINKEDFONTS = 2048
    URL_ENCODING_DISABLE_UTF8 = 4096
    URL_ENCODING_ENABLE_UTF8 = 8192
    ENABLE_FORMS_AUTOCOMPLETE = 16384
    ENABLE_INPLACE_NAVIGATION = 65536
    IME_ENABLE_RECONVERSION = 131072
    THEME = 262144
    NOTHEME = 524288
    NOPICS = 1048576
    NO3DOUTERBORDER = 2097152
    DISABLE_EDIT_NS_FIXUP = 4194304
    LOCAL_MACHINE_ACCESS_CHECK = 8388608
    DISABLE_UNTRUSTEDPROTOCOL = 16777216
    HOST_NAVIGATES = 33554432
    ENABLE_REDIRECT_NOTIFICATION = 67108864
    USE_WINDOWLESS_SELECTCONTROL = 134217728
    USE_WINDOWED_SELECTCONTROL = 268435456
    ENABLE_ACTIVEX_INACTIVATE_MODE = 536870912
    DPI_AWARE = 1073741824
End Enum
'End Class

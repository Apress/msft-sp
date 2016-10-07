Imports System
Imports System.ComponentModel
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Xml.Serialization
Imports Microsoft.SharePoint
Imports Microsoft.SharePoint.Utilities
Imports Microsoft.SharePoint.WebPartPages

'Description for WebPart1.
<DefaultProperty("Text"), ToolboxData("<{0}:Client runat=server></{0}:Client>"), XmlRoot(Namespace:="SPSTerminal")> _
Public Class Client
    Inherits Microsoft.SharePoint.WebPartPages.WebPart

    Private strURL As String = ""
    Private strServer As String = ""
    Private blnFullScreen As Boolean = False
    Private strDisplayWidth As String = "100%"
    Private strDisplayHeight As String = "100%"
    Private intDesktopWidth As Integer = 800
    Private intDesktopHeight As Integer = 600
    Private blnRedirectDrives As Boolean = False
    Private blnRedirectPrinters As Boolean = True

    <Browsable(True), Category("Miscellaneous"), DefaultValue(""), _
    WebPartStorage(Storage.Shared), FriendlyName("URL"), _
    Description("The URL where the web client ASP.NET page is located.")> _
    Property URL() As String
        Get
            Return strURL
        End Get

        Set(ByVal Value As String)
            strURL = Value
        End Set
    End Property

    <Browsable(True), Category("Miscellaneous"), DefaultValue(""), _
    WebPartStorage(Storage.Shared), FriendlyName("Server"), _
    Description("The name of the Terminal Services server.")> _
    Property Server() As String
        Get
            Return strServer
        End Get

        Set(ByVal Value As String)
            strServer = Value
        End Set
    End Property

    <Browsable(True), Category("Miscellaneous"), DefaultValue(True), _
    WebPartStorage(Storage.Shared), FriendlyName("FullScreen"), _
    Description("Determines if the Terminal Services session runs in full screen mode.")> _
    Property FullScreen() As Boolean
        Get
            Return blnFullScreen
        End Get

        Set(ByVal Value As Boolean)
            blnFullScreen = Value
        End Set
    End Property

    <Browsable(True), Category("Miscellaneous"), DefaultValue("100%"), _
    WebPartStorage(Storage.Shared), FriendlyName("DisplayWidth"), _
    Description("Specifies the relative width of the session viewer.")> _
    Property DisplayWidth() As String
        Get
            Return strDisplayWidth
        End Get

        Set(ByVal Value As String)
            strDisplayWidth = Value
        End Set
    End Property

    <Browsable(True), Category("Miscellaneous"), DefaultValue("100%"), _
    WebPartStorage(Storage.Shared), FriendlyName("DisplayHeight"), _
    Description("Specifies the relative height of the session viewer.")> _
    Property DisplayHeight() As String
        Get
            Return strDisplayHeight
        End Get

        Set(ByVal Value As String)
            strDisplayHeight = Value
        End Set
    End Property

    <Browsable(True), Category("Miscellaneous"), DefaultValue(800), _
    WebPartStorage(Storage.Shared), FriendlyName("DesktopWidth"), _
    Description("Specifies the width of the Terminal Services desktop.")> _
    Property DesktopWidth() As Integer
        Get
            Return intDesktopWidth
        End Get

        Set(ByVal Value As Integer)
            intDesktopWidth = Value
        End Set
    End Property

    <Browsable(True), Category("Miscellaneous"), DefaultValue(600), _
    WebPartStorage(Storage.Shared), FriendlyName("DesktopHeight"), _
    Description("Specifies the height of the Terminal Services desktop.")> _
    Property DesktopHeight() As Integer
        Get
            Return intDesktopHeight
        End Get

        Set(ByVal Value As Integer)
            intDesktopHeight = Value
        End Set
    End Property

    <Browsable(True), Category("Miscellaneous"), DefaultValue(False), _
    WebPartStorage(Storage.Shared), FriendlyName("RedirectDrives"), _
    Description("Determines if the client drives are mapped to the Terminal Services session.")> _
    Property RedirectDrives() As Boolean
        Get
            Return blnRedirectDrives
        End Get

        Set(ByVal Value As Boolean)
            blnRedirectDrives = Value
        End Set
    End Property

    <Browsable(True), Category("Miscellaneous"), DefaultValue(True), _
    WebPartStorage(Storage.Shared), FriendlyName("RedirectPrinters"), _
    Description("Determines if the client printers are mapped to the Terminal Services session.")> _
    Property RedirectPrinters() As Boolean
        Get
            Return blnRedirectPrinters
        End Get

        Set(ByVal Value As Boolean)
            blnRedirectPrinters = Value
        End Set
    End Property


    Protected Overrides Sub RenderWebPart( _
    ByVal output As System.Web.UI.HtmlTextWriter)

        With output

            Dim strConnectURL As String = ""
            strConnectURL += URL
            strConnectURL += "?Server=" + Server
            strConnectURL += "&FullScreen=" + FullScreen.ToString
            strConnectURL += "&DeskTopWidth=" + DesktopWidth.ToString
            strConnectURL += "&DeskTopHeight=" + DesktopHeight.ToString
            strConnectURL += "&DisplayWidth=" + DisplayWidth
            strConnectURL += "&DisplayHeight=" + DisplayHeight
            strConnectURL += "&RedirectDrives=" + RedirectDrives.ToString
            strConnectURL += "&RedirectPrinters=" + RedirectPrinters.ToString
            strConnectURL += "&Title=" + Title

            .Write("<a href=""" + strConnectURL + """ Target=""_blank"">" & _
              "Connect to " + Server + "</a><br>" + vbCrLf)
        End With

    End Sub



End Class

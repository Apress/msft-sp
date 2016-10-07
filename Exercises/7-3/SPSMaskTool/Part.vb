Imports System
Imports System.ComponentModel
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Xml.Serialization
Imports Microsoft.SharePoint
Imports Microsoft.SharePoint.Utilities
Imports Microsoft.SharePoint.WebPartPages

<DefaultProperty("Text"), ToolboxData("<{0}:Part runat=server></{0}:Part>"), XmlRoot(Namespace:="SPSMaskTool")> _
 Public Class Part
    Inherits Microsoft.SharePoint.WebPartPages.WebPart

    Dim m_text As String = ""

    'Browsable must be False to hide the normal custom property tool part
    <Browsable(False), Category("Miscellaneous"), DefaultValue(""), WebPartStorage(Storage.Personal), FriendlyName("Text"), Description("Text Property")> _
    Property [Text]() As String
        Get
            Return m_text
        End Get

        Set(ByVal Value As String)
            m_text = Value
        End Set
    End Property

    Protected Overrides Sub RenderWebPart(ByVal output As _
System.Web.UI.HtmlTextWriter)

        output.Write(SPEncode.HtmlEncode([Text]))
    End Sub

    Public Overrides Function GetToolParts() As ToolPart()

        'This code is required because it was contained in the
        'method we are overriding.  We cannot simply call the base class
        'because we can only return a single array, so we have to rebuild it.
        Dim toolParts(3) As ToolPart
        Dim objWebToolPart As WebPartToolPart = New WebPartToolPart
        Dim objCustomProperty As CustomPropertyToolPart = New CustomPropertyToolPart
        toolParts(0) = objWebToolPart
        toolParts(1) = objCustomProperty

        'This is where we add our tool part
        toolParts(2) = New Tool

        Return toolParts
    End Function


End Class

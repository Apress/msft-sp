Public Class WebForm1
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load

        With Response
            .Write(vbCrLf)
            .Write("<script language=""VBScript"">" + vbCrLf)
            .Write("<!--" + vbCrLf)
            .Write("Sub StateChange" + vbCrLf)
            .Write("  set RDP = Document.getElementById(""MsRdpClient"")" + vbCrLf)
            .Write("  If RDP.ReadyState = 4 Then" + vbCrLf)
            .Write("    RDP.Server = """ + Request.QueryString("Server") + """" + vbCrLf)
            .Write("    RDP.FullScreen = " + Request.QueryString("FullScreen") + vbCrLf)
            .Write("    RDP.DesktopWidth = """ + Request.QueryString("DesktopWidth") _
          + """" + vbCrLf)
            .Write("    RDP.DesktopHeight = """ + Request.QueryString("DesktopHeight") _
          + """" + vbCrLf)
            .Write("    RDP.AdvancedSettings2.RedirectDrives = " _
           + Request.QueryString("RedirectDrives") + vbCrLf)
            .Write("    RDP.AdvancedSettings2.RedirectPrinters = " _
          + Request.QueryString("RedirectPrinters") + vbCrLf)
            .Write("    RDP.FullScreenTitle = """ + Request.QueryString("Title") + _
          """" + vbCrLf)
            .Write("    RDP.Connect" + vbCrLf)
            .Write("  End If" + vbCrLf)
            .Write("End Sub" + vbCrLf)
            .Write("-->" + vbCrLf)
            .Write("</script>" + vbCrLf)
            .Write(vbCrLf)

            .Write("<OBJECT ID=""MsRdpClient"" Language=""VBScript""" + vbCrLf)
            .Write("CLASSID=""CLSID:7584c670-2274-4efb-b00b-d6aaba6d3850""" + vbCrLf)
            .Write("CODEBASE=""msrdp.cab#version=5,2,3790,0""" + vbCrLf)
            .Write("OnReadyStateChange=""StateChange""" + vbCrLf)
            .Write("WIDTH=""" + Request.QueryString("DisplayWidth") + """" + vbCrLf)
            .Write("HEIGHT=""" + Request.QueryString("DisplayHeight") + """" + vbCrLf)
            .Write("</OBJECT>" + vbCrLf)
        End With

    End Sub


End Class

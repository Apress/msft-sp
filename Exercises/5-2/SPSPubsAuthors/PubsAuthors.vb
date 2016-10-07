Imports System
Imports System.ComponentModel
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Xml.Serialization
Imports Microsoft.SharePoint
Imports Microsoft.SharePoint.Utilities
Imports Microsoft.SharePoint.WebPartPages

<DefaultProperty(""), ToolboxData("<{0}:View runat=server></{0}:View>"), XmlRoot(Namespace:="SPSPubsAuthors")> _
Public Class View
    Inherits Microsoft.SharePoint.WebPartPages.WebPart

    Private strSQLserver As String = ""
    Private strDatabase As String = ""
    Private strUserName As String = ""
    Private strPassword As String = ""

    'SQL Server Name
    <Browsable(True), Category("Miscellaneous"), DefaultValue(""), _
    WebPartStorage(Storage.Shared), FriendlyName("SQLServer"), _
    Description("The server where eRelationship is installed.")> _
    Property SQLServer() As String
        Get
            Return strSQLserver
        End Get

        Set(ByVal Value As String)
            strSQLserver = Value
        End Set
    End Property

    'Database Name
    <Browsable(True), Category("Miscellaneous"), DefaultValue(""), _
    WebPartStorage(Storage.Shared), FriendlyName("Database"), _
    Description("The database where the Enterprise Data is located.")> _
    Property Database() As String
        Get
            Return strDatabase
        End Get

        Set(ByVal Value As String)
            strDatabase = Value
        End Set
    End Property

    'User Name
    <Browsable(True), Category("Miscellaneous"), DefaultValue(""), _
    WebPartStorage(Storage.Shared), FriendlyName("UserName"), _
    Description("The account to use to access the database.")> _
    Property UserName() As String
        Get
            Return strUserName
        End Get

        Set(ByVal Value As String)
            strUserName = Value
        End Set
    End Property

    'Password
    <Browsable(True), Category("Miscellaneous"), DefaultValue(""), _
    WebPartStorage(Storage.Shared), FriendlyName("Password"), _
    Description("The password to access the database.")> _
    Property Password() As String
        Get
            Return strPassword
        End Get

        Set(ByVal Value As String)
            strPassword = Value
        End Set
    End Property

    Protected WithEvents grdNames As DataGrid
    Protected WithEvents lblMessage As Label

    Protected Overrides Sub CreateChildControls()

        'Grid to display results
        grdNames = New DataGrid
        With grdNames
            .Width = Unit.Percentage(100)
            .HeaderStyle.Font.Name = "arial"
            .HeaderStyle.Font.Size = New FontUnit(FontSize.AsUnit).Point(10)
            .HeaderStyle.Font.Bold = True
            .HeaderStyle.ForeColor = System.Drawing.Color.Wheat
            .HeaderStyle.BackColor = System.Drawing.Color.DarkBlue
            .AlternatingItemStyle.BackColor = System.Drawing.Color.LightCyan
        End With
        Controls.Add(grdNames)

        'Label for error messages
        lblMessage = New Label
        With lblMessage
            .Width = Unit.Percentage(100)
            .Font.Name = "arial"
            .Font.Size = New FontUnit(FontSize.AsUnit).Point(10)
            .Text = ""
        End With
        Controls.Add(lblMessage)

    End Sub

    Protected Overrides Sub RenderWebPart _
(ByVal output As System.Web.UI.HtmlTextWriter)

        Dim objDataSet As System.Data.DataSet

        'Set up connection string from custom properties
        Dim strConnection As String
        strConnection += "Password=" & Password
        strConnection += ";Persist Security Info=True;User ID="
        strConnection += UserName + ";Initial Catalog=" + Database
        strConnection += ";Data Source=" + SQLServer

        'Query pubs database 
        Dim strSQL As String = "select * from authors"

        'Try to run the query
        Try
            With New System.Data.SqlClient.SqlDataAdapter
                objDataSet = New DataSet("root")

                .SelectCommand = _
                New System.Data.SqlClient.SqlCommand(strSQL, _
                New System.Data.SqlClient.SqlConnection(strConnection))

                .Fill(objDataSet, "authors")
            End With
        Catch ex As Exception
            lblMessage.Text = ex.Message
            Exit Sub
        End Try

        'Bind to grid
        Try
            With grdNames
                .DataSource = objDataSet
                .DataMember = "authors"
                .DataBind()
            End With
        Catch ex As Exception
            lblMessage.Text = ex.Message
            Exit Sub
        End Try

        'Draw the controls in an HTML table
        With output
            .Write("<TABLE BORDER=0 WIDTH=100%>")
            .Write("<TR>")
            .Write("<TD>")
            grdNames.RenderControl(output)
            .Write("</TD>")
            .Write("</TR>")
            .Write("<TR>")
            .Write("<TD>")
            lblMessage.RenderControl(output)
            .Write("</TD>")
            .Write("</TR>")
            .Write("</TABLE>")
        End With

    End Sub

End Class

Option Strict On
Option Explicit On 
Option Compare Text

Imports System
Imports System.ComponentModel
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Xml.Serialization
Imports Microsoft.SharePoint
Imports Microsoft.SharePoint.Utilities
Imports Microsoft.SharePoint.WebPartPages
Imports Microsoft.SharePoint.WebControls
Imports System.Security.Principal
Imports System.Runtime.InteropServices
Imports Microsoft.SharePoint.Portal.SingleSignon

<DefaultProperty("ShowAllSites"), _
ToolboxData("<{0}:Lister runat=server></{0}:Lister>"), _
XmlRoot(Namespace:="SPSSubSites")> _
 Public Class Lister
    Inherits Microsoft.SharePoint.WebPartPages.WebPart

    Protected blnShowAllSites As Boolean = False

    <Browsable(True), Category("Behavior"), DefaultValue(False), _
    WebPartStorage(Storage.Shared), FriendlyName("Show All Sites"), _
    Description("Show sites, to which the user does not belong.")> _
    Property ShowAllSites() As Boolean
        Get
            Return blnShowAllSites
        End Get

        Set(ByVal Value As Boolean)
            blnShowAllSites = Value
        End Set
    End Property

    Protected WithEvents grdSites As DataGrid
    Protected WithEvents lblMessage As Label

    Protected Overrides Sub CreateChildControls()

        'Grid to display results
        grdSites = New DataGrid
        With grdSites
            .AutoGenerateColumns = False
            .Width = Unit.Percentage(100)
            .HeaderStyle.Font.Name = "arial"
            .HeaderStyle.Font.Size = New FontUnit(FontSize.AsUnit).Point(10)
            .HeaderStyle.Font.Bold = True
            .HeaderStyle.ForeColor = System.Drawing.Color.Wheat
            .HeaderStyle.BackColor = System.Drawing.Color.DarkBlue
            .AlternatingItemStyle.BackColor = System.Drawing.Color.LightCyan
        End With

        Dim objBoundColumn As BoundColumn
        Dim objHyperColumn As HyperLinkColumn

        'Name Column
        objHyperColumn = New HyperLinkColumn
        With objHyperColumn
            .HeaderText = "Site Name"
            .DataTextField = "Name"
            .DataNavigateUrlField = "URL"
            grdSites.Columns.Add(objHyperColumn)
        End With

        'Membership Column
        objBoundColumn = New BoundColumn
        With objBoundColumn
            .HeaderText = "Your Role"
            .DataField = "Role"
            grdSites.Columns.Add(objBoundColumn)
        End With

        'Description Column
        objBoundColumn = New BoundColumn
        With objBoundColumn
            .HeaderText = "Site Description"
            .DataField = "Description"
            grdSites.Columns.Add(objBoundColumn)
        End With

        'Contact Column
        objHyperColumn = New HyperLinkColumn
        With objHyperColumn
            .HeaderText = "Site Contact"
            .DataTextField = "Author"
            .DataNavigateUrlField = "eMail"
            grdSites.Columns.Add(objHyperColumn)
        End With

        Controls.Add(grdSites)

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

    Protected Shared Function CreateIdentity(ByVal User As String, _
ByVal Domain As String, ByVal Password As String) As WindowsIdentity

        Dim objToken As New IntPtr(0)
        Dim ID As WindowsIdentity
        Const LOGON32_PROVIDER_DEFAULT As Integer = 0
        Const LOGON32_LOGON_NETWORK As Integer = 3

        'Initialize token object
        objToken = IntPtr.Zero

        ' Attempt to log on
        Dim blnReturn As Boolean = LogonUser(User, Domain, Password, _
        LOGON32_LOGON_NETWORK, LOGON32_PROVIDER_DEFAULT, objToken)

        'Check for failure
        If blnReturn = False Then
            Dim intCode As Integer = Marshal.GetLastWin32Error()
            Throw New Exception("Logon failed: " & intCode.ToString)
        End If

        'Return new token
        ID = New WindowsIdentity(objToken)
        CloseHandle(objToken)
        Return ID

    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> _
        Private Shared Function LogonUser(ByVal lpszUsername As String, _
    ByVal lpszDomain As String, _
    ByVal lpszPassword As String, ByVal dwLogonType As Integer, _
    ByVal dwLogonProvider As Integer, _
            ByRef phToken As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto)> _
        Private Shared Function CloseHandle(ByVal handle As IntPtr) As Boolean
    End Function


    'Render this Web Part to the output parameter specified.
    Protected Overrides Sub RenderWebPart( _
    ByVal output As System.Web.UI.HtmlTextWriter)

        'Get the site collection
        Dim objSite As SPSite = SPControl.GetContextSite(Context)
        Dim objMainSite As SPWeb = objSite.OpenWeb
        Dim objAllSites As SPWebCollection = objSite.AllWebs
        Dim objMemberSites As SPWebCollection = objMainSite.GetSubwebsForCurrentUser
        Dim objSubSite As SPWeb

        'Get the user identity
        Dim strUsername As String = objMainSite.CurrentUser.LoginName

        'Create a DataSet and DataTable for the site collection
        Dim objDataset As DataSet = New DataSet("root")
        Dim objTable As DataTable = objDataset.Tables.Add("Sites")

        'Context for the new identity
        Dim objContext As WindowsImpersonationContext
        Dim arrCredentials() As String
        Dim strUID As String
        Dim strDomain As String
        Dim strPassword As String

        Try

            'Try to get credentials
            Credentials.GetCredentials( _
            Convert.ToUInt32("0"), "SubSiteList", arrCredentials)
            strUID = arrCredentials(0)
            strDomain = arrCredentials(1)
            strPassword = arrCredentials(2)

            'Change the context
            Dim objIdentity As WindowsIdentity
            objIdentity = CreateIdentity(strUID, strDomain, strPassword)
            objContext = objIdentity.Impersonate

        Catch x As SingleSignonException
            lblMessage.Text += "No credentials available." + vbCrLf
        Catch x As Exception
            lblMessage.Text += x.Message + vbCrLf
        End Try

        Try
            'Design Table
            With objTable.Columns
                .Add("Role", Type.GetType("System.String"))
                .Add("Name", Type.GetType("System.String"))
                .Add("Description", Type.GetType("System.String"))
                .Add("URL", Type.GetType("System.String"))
                .Add("Author", Type.GetType("System.String"))
                .Add("eMail", Type.GetType("System.String"))
            End With

            'Fill the Table with Member Sites
            For Each objSubSite In objMemberSites
                Dim objRow As DataRow = objTable.NewRow()
                With objRow
                    Try
                        .Item("Role") = objSubSite.Users(strUsername).Roles(0).Name
                    Catch
                        .Item("Role") = "None"
                    End Try
                    .Item("Name") = objSubSite.Name
                    .Item("Description") = objSubSite.Description
                    .Item("URL") = objSubSite.Url
                    .Item("Author") = objSubSite.Author.Name
                    .Item("eMail") = "mailto:" + objSubSite.Author.Email
                End With
                objTable.Rows.Add(objRow)
            Next
        Catch x As Exception
            lblMessage.Text += x.Message + vbCrLf
        End Try

        Try
            'Fill the Table with non-member sites
            If ShowAllSites = True Then

                For Each objSubSite In objAllSites

                    'Get the user collection for each sub site
                    Dim objUsers As SPUserCollection = objSubSite.Users
                    Dim objUser As SPUser
                    Dim blnMember As Boolean

                    'Skip the parent site
                    If objMainSite.Name <> objSubSite.Name Then

                        blnMember = False

                        'Look through user list
                        For Each objUser In objUsers
                            If objUser.LoginName.Trim = strUsername.Trim Then
                                blnMember = True
                            End If
                        Next

                        If blnMember = False Then

                            'If the current user is not a member add a record
                            Dim objRow As DataRow = objTable.NewRow()
                            With objRow
                                .Item("Role") = "Not a Member!"
                                .Item("Name") = objSubSite.Name
                                .Item("Description") = objSubSite.Description
                                .Item("URL") = objSubSite.Url
                                .Item("Author") = objSubSite.Author.Name
                                .Item("eMail") = "mailto:" + objSubSite.Author.Email
                            End With
                            objTable.Rows.Add(objRow)

                        End If

                    End If

                    objSubSite.Close()

                Next

            End If

            'Close sites
            objMainSite.Close()
            objSite.Close()

            'Tear down context
            objContext.Undo()

        Catch x As Exception
            lblMessage.Text += x.Message + vbCrLf
        End Try

        'Bind dataset to grid
        output.Write("<p>Current user: " + objMainSite.CurrentUser.Name + "<br>" _
        + "Collection owner: <a href=""mailto:" + objSite.Owner.Email + """>" _
        + objSite.Owner.Name + "</a></p>")
        With grdSites
            .DataSource = objDataset
            .DataMember = "Sites"
            .DataBind()
        End With

        'Show grid
        grdSites.RenderControl(output)
        output.Write("<br>")
        lblMessage.RenderControl(output)

    End Sub

End Class

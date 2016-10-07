Option Explicit On 
Option Strict On
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
Imports Microsoft.SharePoint.Administration
Imports System.Security.Principal
Imports System.Runtime.InteropServices
Imports Microsoft.SharePoint.Portal.SingleSignon

<DefaultProperty(""), ToolboxData("<{0}:Lister runat=server></{0}:Lister>"), _
XmlRoot(Namespace:="SPSTasks")> _
 Public Class Lister
    Inherits Microsoft.SharePoint.WebPartPages.WebPart
    Protected WithEvents grdTasks As DataGrid
    Protected WithEvents lblMessage As Label

    Protected Overrides Sub CreateChildControls()

        'Grid to display results
        grdTasks = New DataGrid
        With grdTasks
            .AutoGenerateColumns = False
            .Width = Unit.Percentage(100)
            .HeaderStyle.Font.Name = "arial"
            .HeaderStyle.Font.Size = New FontUnit(FontSize.AsUnit).Point(10)
            .HeaderStyle.Font.Bold = True
            .HeaderStyle.ForeColor = System.Drawing.Color.Wheat
            .HeaderStyle.BackColor = System.Drawing.Color.DarkBlue
            .AlternatingItemStyle.BackColor = System.Drawing.Color.LightCyan
        End With

        Dim objHyperColumn As HyperLinkColumn

        'Host Site Name Column
        objHyperColumn = New HyperLinkColumn
        With objHyperColumn
            .HeaderText = "Host Site"
            .DataTextField = "SiteName"
            .DataNavigateUrlField = "SiteURL"
            grdTasks.Columns.Add(objHyperColumn)
        End With

        'Host Site Name Column
        objHyperColumn = New HyperLinkColumn
        With objHyperColumn
            .HeaderText = "Task"
            .DataTextField = "TaskTitle"
            .DataNavigateUrlField = "ListURL"
            grdTasks.Columns.Add(objHyperColumn)
        End With
        Controls.Add(grdTasks)

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

    Protected Function GetGlobalTasks(ByVal objUser As SPUser) As DataSet

        'Purpose: Walk all sites and collect pointers to the tasks

        'Context for the new identity
        Dim objContext As WindowsImpersonationContext
        Dim arrCredentials() As String
        Dim strUID As String
        Dim strDomain As String
        Dim strPassword As String
        Dim objDataSet As DataSet

        Try

            'Try to get credentials
            Credentials.GetCredentials( _
    Convert.ToUInt32("0"), "SPSAuthority", arrCredentials)
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
            'Create new DataTable for tasks
            objDataSet = New DataSet("root")
            Dim objTable As DataTable = objDataSet.Tables.Add("Tasks")

            'Design Table
            With objTable.Columns
                .Add("SiteName", Type.GetType("System.String"))
                .Add("SiteURL", Type.GetType("System.String"))
                .Add("TaskTitle", Type.GetType("System.String"))
                .Add("ListURL", Type.GetType("System.String"))
            End With

            'Fill DataTable with tasks for the current user
            Dim objAdmin As New SPGlobalAdmin
            Dim objServer As SPVirtualServer = objAdmin.VirtualServers(0)
            Dim objSites As SPSiteCollection = objServer.Sites
            Dim objSite As SPSite

            'Walk every site in the installation
            For Each objSite In objSites

                Dim objWeb As SPWeb = objSite.OpenWeb()
                Dim objLists As SPListCollection = objWeb.Lists
                Dim objList As SPList

                'Walk every list on a site
                For Each objList In objLists

                    If objList.BaseType = SPBaseType.GenericList _
                    OrElse objList.BaseType = SPBaseType.Issue Then

                        For i As Integer = 0 To objList.ItemCount - 1
                            Try
                                Dim objItem As SPListItem = objList.Items(i)

                                'Check to see if this task is assigned to the user
                                Dim strAssignedTo As String = _
                                UCase(objItem.Item("Assigned To").ToString)

                                If strAssignedTo.IndexOf( _
                                UCase(objUser.LoginName)) > -1 _
                                OrElse strAssignedTo.IndexOf( _
                                UCase(objUser.Name)) > -1 Then

                                    'If so, add it to the DataSet
                                    Dim objRow As DataRow = objTable.NewRow()
                                    With objRow
                                        .Item("SiteName") = objList.ParentWeb.Title
                                        .Item("SiteURL") = objList.ParentWeb.Url
                                        .Item("TaskTitle") = objItem("Title")
                                        .Item("ListURL") = objList.DefaultViewUrl
                                    End With
                                    objTable.Rows.Add(objRow)
                                End If

                            Catch
                            End Try

                        Next
                    End If
                Next

                objWeb.Close()

            Next

            'Tear down the context
            objContext.Undo()

        Catch x As Exception
            lblMessage.Text += x.Message + vbCrLf
        End Try

        Return objDataSet

    End Function


    Protected Overrides Sub RenderWebPart( _
    ByVal output As System.Web.UI.HtmlTextWriter)

        'Get the site collection
        Dim objSite As SPSite = SPControl.GetContextSite(Context)
        Dim objWeb As SPWeb = objSite.OpenWeb
        Dim objUser As SPUser = objWeb.CurrentUser

        'Get the DataSet of Tasks
        Dim objDataSet As DataSet = GetGlobalTasks(objUser)

        'Display Tasks
        With grdTasks
            .DataSource = objDataSet
            .DataMember = "Tasks"
            .DataBind()
        End With

        'Show grid
        grdTasks.RenderControl(output)
        output.Write("<br>")
        lblMessage.RenderControl(output)

    End Sub


End Class

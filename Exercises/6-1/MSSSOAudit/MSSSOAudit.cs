using System;
using System.ComponentModel;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml.Serialization;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebPartPages;
using Microsoft.SharePoint.Portal.SingleSignon;
using System.Data;
using System.Data.SqlClient;

namespace MSSSOAudit
{

	[DefaultProperty(""),
	ToolboxData("<{0}:Log runat=server></{0}:Log>"),
	XmlRoot(Namespace="MSSSOAudit")]
	public class Log : Microsoft.SharePoint.WebPartPages.WebPart
	{

		//PROPERTIES
		private string m_userName="";
		private string m_password="";

		[Browsable(false),Category("Miscellaneous"),
		DefaultValue(""),
		WebPartStorage(Storage.Shared),
		FriendlyName("UserName"),Description("The account name to access the SSO database")]
		public string userName
		{
			get
			{
				return m_userName;
			}

			set
			{
				m_userName = value;
			}
		}

		[Browsable(false),Category("Miscellaneous"),
		DefaultValue(""),
		WebPartStorage(Storage.Shared),
		FriendlyName("Password"),Description("The password to access the SSO database")]
		public string password
		{
			get
			{
				return m_password;
			}

			set
			{
				m_password = value;
			}
		}

		//CHILD CONTROLS
		protected DataGrid grdAudit;
		protected Label lblMessage;

		protected override void CreateChildControls()
		{
			//DataGrid
			grdAudit = new DataGrid();
			grdAudit.Width = Unit.Percentage(100);
			grdAudit.HeaderStyle.Font.Name = "arial";
			grdAudit.HeaderStyle.ForeColor = System.Drawing.Color.Wheat;
			grdAudit.HeaderStyle.BackColor = System.Drawing.Color.DarkBlue;
			grdAudit.AlternatingItemStyle.BackColor = System.Drawing.Color.LightCyan;
			Controls.Add(grdAudit);

			//Label
			lblMessage=new Label();
			lblMessage.Width = Unit.Percentage(100);
			lblMessage.Font.Name = "arial";
			lblMessage.Text = "";
			Controls.Add(lblMessage);
		}

		//RENDERING
		protected override void RenderWebPart(HtmlTextWriter output)
		{
			string[] strCredentials=null;
			string strConnection=null;
			SqlDataAdapter objAdapter;
			SqlCommand objCommand;
			SqlConnection objConnection;
			DataSet objDataSet;

			//Try to get credentials
			try
			{

				// Call MSSSO
				Credentials.GetCredentials(Convert.ToUInt32(1),"MSSSOAudit",ref strCredentials);

				//save credentials
				userName=strCredentials[0];
				password=strCredentials[1];

				//Create connection string
				strConnection += "Password=" + password;
				strConnection += ";Persist Security Info=True;";
				strConnection += "User ID=" + userName + ";Initial Catalog=SSO;";
				strConnection += "Data Source=(local)";

			}
			catch (SingleSignonException x)
			{
				if (x.LastErrorCode==SSOReturnCodes.SSO_E_CREDS_NOT_FOUND)
				{lblMessage.Text="Credentials not found!";}
				else
				{lblMessage.Text=x.Message;}
			}

			//Try to show the grid
			try
			{
				//query the SSO database
				objAdapter=new SqlDataAdapter();
				objDataSet=new DataSet("root");

				objConnection=new SqlConnection(strConnection);
				objCommand=new SqlCommand("Select * from SSO_Audit",objConnection);
				objAdapter.SelectCommand=objCommand;
				objAdapter.Fill(objDataSet,"audit");

				//bind to the grid
				grdAudit.DataSource=objDataSet;
				grdAudit.DataMember="audit";
				grdAudit.DataBind();
			}
			catch (Exception x)
			{
				lblMessage.Text+=x.Message;
			}
			finally
			{
				//draw grid
				grdAudit.RenderControl(output);
				output.Write("<BR>");
				lblMessage.RenderControl(output);
			}
		}
	}
}


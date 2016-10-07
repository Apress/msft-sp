using System;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.ComponentModel;
using System.Xml.Serialization;
using System.Security;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebPartPages;
using Microsoft.SharePoint.WebPartPages.Communication;
using Microsoft.SharePoint.Portal;
using Microsoft.SharePoint.Portal.SingleSignon;
using System.Data;
using System.Data.SqlClient;

namespace SPSMultiFace
{
	[DefaultProperty(""),
	ToolboxData("<{0}:Detail runat=server></{0}:Detail>"),
	XmlRoot(Namespace="SPSMultiFace")]

	//Inherits WebPart and implements ICellConsumer
	public class Detail : Microsoft.SharePoint.WebPartPages.WebPart ,
		ICellConsumer, IRowProvider
	{
		private string m_sqlServer="";
		private string m_database="";
		private string m_userName="";
		private string m_password="";

		[Browsable(false),Category("Miscellaneous"),DefaultValue(""),
		WebPartStorage(Storage.Shared),FriendlyName("sqlServer"),
		Description("The server name where SQL Server is located.")]
		public string sqlServer
		{
			get
			{
				return m_sqlServer;
			}

			set
			{
				m_sqlServer = value;
			}
		}

		[Browsable(false),Category("Miscellaneous"),DefaultValue(""),
		WebPartStorage(Storage.Shared),FriendlyName("database"),
		Description("The name of the database (i. e., 'pubs')")]
		public string database
		{
			get
			{
				return m_database;
			}

			set
			{
				m_database = value;
			}
		}

		[Browsable(false),Category("Miscellaneous"),DefaultValue(""),
		WebPartStorage(Storage.Shared),FriendlyName("userName"),
		Description("The user name to access the database.")]
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

		[Browsable(false),Category("Miscellaneous"),DefaultValue(""),
		WebPartStorage(Storage.Shared),FriendlyName("password"),
		Description("The password for the database.")]
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

		public event CellConsumerInitEventHandler CellConsumerInit;

		public event RowReadyEventHandler RowReady;
		public event RowProviderInitEventHandler RowProviderInit;

		//Child Controls
		protected DataGrid grdBooks;
		protected Label lblMessage;

		protected override void CreateChildControls()
		{

			//Purpose: draw the user interface
			grdBooks = new DataGrid();
			grdBooks.AutoGenerateColumns=false;
			grdBooks.Width=Unit.Percentage(100);
			grdBooks.HeaderStyle.Font.Name = "arial";
			grdBooks.HeaderStyle.Font.Name = "arial";
			grdBooks.HeaderStyle.Font.Bold = true;
			grdBooks.HeaderStyle.ForeColor = System.Drawing.Color.Wheat;
			grdBooks.HeaderStyle.BackColor = System.Drawing.Color.DarkBlue;
			grdBooks.AlternatingItemStyle.BackColor = System.Drawing.Color.LightCyan;
			grdBooks.SelectedItemStyle.BackColor=System.Drawing.Color.Blue;

			//Add a button to the grid for selection
			ButtonColumn objButtonColumn = new ButtonColumn();
			objButtonColumn.Text="Select";
			objButtonColumn.CommandName="Select";
			grdBooks.Columns.Add(objButtonColumn);

			//Add data columns
			BoundColumn objColumn = new BoundColumn();
			objColumn.DataField="title_id";
			objColumn.HeaderText="Title ID";
			grdBooks.Columns.Add(objColumn);

			objColumn = new BoundColumn();
			objColumn.DataField="title";
			objColumn.HeaderText="Title";
			grdBooks.Columns.Add(objColumn);

			objColumn = new BoundColumn();
			objColumn.DataField="price";
			objColumn.HeaderText="Price";
			grdBooks.Columns.Add(objColumn);

			objColumn = new BoundColumn();
			objColumn.DataField="ytd_sales";
			objColumn.HeaderText="2003 Sales";
			grdBooks.Columns.Add(objColumn);

			objColumn = new BoundColumn();
			objColumn.DataField="pubdate";
			objColumn.HeaderText="Published";
			grdBooks.Columns.Add(objColumn);

			Controls.Add(grdBooks);

			lblMessage = new Label();
			lblMessage.Width = Unit.Percentage(100);
			lblMessage.Font.Name = "arial";
			lblMessage.Text = "";
			Controls.Add(lblMessage);

		}

		//Private member variables
		private int m_rowConsumers = 0;
		private int m_cellProviders=0;

		public override void EnsureInterfaces()
		{
			//Tell the connection infrastructure what interfaces the web part supports
			try
			{
				RegisterInterface("PublisherConsumer_WPQ_",
					"ICellConsumer",
					WebPart.LimitOneConnection,
					ConnectionRunAt.Server,
					this,
					"",
					"Get an publisher name from...",
					"Receives a publisher name.");

				RegisterInterface("BookProvider_WPQ_",
					"IRowProvider",
					WebPart.UnlimitedConnections,
					ConnectionRunAt.Server,
					this,
					"",
					"Provide a row to...",
					"Provides book information as a row of data.");

			}
			catch(SecurityException e)
			{
				//Use "WSS_Minimal" or "WSS_Medium" to utilize connections
				lblMessage.Text+=e.Message+ "<br>";
			}
		}

		public override ConnectionRunAt CanRunAt()
		{
			//This Web Part runs on the server
			return ConnectionRunAt.Server;
		}

		public override void PartCommunicationConnect(string interfaceName,
			WebPart connectedPart,string connectedInterfaceName,ConnectionRunAt runAt)
		{
			//Make sure this is a server-side connection
			if (runAt==ConnectionRunAt.Server)
			{
				//Draw the controls for the web part
				EnsureChildControls(); 

				//Check if this is my particular cell interface
				if (interfaceName == "PublisherConsumer_WPQ_")
				{
					//Keep a count of the connections
					m_cellProviders++;
				}

				if (interfaceName == "BookProvider_WPQ_")
				{
					//Keep a count of the connections
					m_rowConsumers++;
				}
			}
		}

		public override InitEventArgs GetInitEventArgs(string interfaceName)
		{

			//Purpose: provide data for a transformer

			if (interfaceName == "PublisherConsumer_WPQ_")
			{
				EnsureChildControls();

				CellConsumerInitEventArgs initCellArgs = new CellConsumerInitEventArgs();

				//Field name metadata
				initCellArgs.FieldName = "pub_name";
				initCellArgs.FieldDisplayName = "Publisher name";

				//return the metadata
				return(initCellArgs);
			}

			else if (interfaceName == "BookProvider_WPQ_")
			{
				EnsureChildControls();

				RowProviderInitEventArgs initRowArgs = new RowProviderInitEventArgs();

				//Field names metadata
				char [] splitter =";".ToCharArray();
				string [] fieldNames =
					"title_id;title;price;ytd_sales;pubdate".Split(splitter);
				string [] fieldTitles =
					"Title ID;Title;Price;Sales;Date".Split(splitter);

				initRowArgs.FieldList = fieldNames;
				initRowArgs.FieldDisplayList=fieldTitles;

				//return the metadata
				return(initRowArgs);
			}

			else
			{
				return null;
			}
		}

		public override void PartCommunicationInit()
		{
			if(m_cellProviders > 0)
			{

				CellConsumerInitEventArgs initCellArgs = new CellConsumerInitEventArgs();

				//Field name metadata
				initCellArgs.FieldName = "pub_name";
				initCellArgs.FieldDisplayName = "Publisher name";

				//Fire the event to broadcast what field the web part can consume
				CellConsumerInit(this, initCellArgs);
			}

			if(m_rowConsumers>0)
			{
				RowProviderInitEventArgs initRowArgs = new RowProviderInitEventArgs();

				//Field names metadata
				char [] splitter =";".ToCharArray();
				string [] fieldNames =
					"title_id;title;price;ytd_sales;pubdate".Split(splitter);
				string [] fieldTitles =
					"Title ID;Title;Price;Sales;Date".Split(splitter);

				initRowArgs.FieldList = fieldNames;
				initRowArgs.FieldDisplayList=fieldTitles;

				//Fire event to broadcast what fields the web part can provide
				RowProviderInit(this,initRowArgs);

			}
		}

		public override void PartCommunicationMain()
		{
			if (m_rowConsumers>0)
			{
				string status = string.Empty;
				DataRow[] dataRows = new DataRow[1];

				if(grdBooks.SelectedIndex > -1)
				{
					//Send selected row
					DataSet dataSet = (DataSet)grdBooks.DataSource;
					DataTable objTable = dataSet.Tables["books"];
					dataRows[0] = objTable.Rows[grdBooks.SelectedIndex];
					status = "Standard";
				}
				else
				{
					//Sene a null row
					dataRows[0] = null;
					status = "None";
				}

				//Fire the event
				RowReadyEventArgs rowReadyArgs = new RowReadyEventArgs();
				rowReadyArgs.Rows=dataRows;
				rowReadyArgs.SelectionStatus=status;
				RowReady(this,rowReadyArgs);

			}
		}

		public void CellProviderInit(object sender, CellProviderInitEventArgs cellProviderInitArgs)
		{
			//Since we can only consume one kind of cell, nothing can be done here
		}

		public void CellReady(object sender, CellReadyEventArgs cellReadyArgs)
		{

			//Purpose: Run the query whenever a new cell value is provided
			EnsureChildControls();

			//Get Credentials (from SSO in production environment)
			userName="sa";
			password="";
			database="pubs";
			sqlServer="(local)";
			string strConn = "Password=" + password + ";Persist Security Info=True " +
				";User ID=" + userName + ";Initial Catalog=" + database + ";Data Source=" +
				sqlServer;

			//Build SQL statement
			string strSQL;
			DataSet dataSet = new DataSet("books");

			try
			{
				strSQL = "SELECT title_id, title, price, ytd_sales, pubdate ";
				strSQL += "FROM publishers INNER JOIN titles ";
				strSQL += "ON publishers.pub_id = titles.pub_id ";
				strSQL += "WHERE (pub_name = '" + cellReadyArgs.Cell.ToString() + " ')";
			}
			catch
			{
				lblMessage.Text="Select a value from a connected web part.";
				return;
			}

			//Run the query
			try
			{
				SqlConnection conn = new SqlConnection(strConn);
				SqlDataAdapter adapter = new SqlDataAdapter(strSQL,conn);
				adapter.Fill(dataSet,"books");
			}
			catch(Exception x)
			{
				lblMessage.Text += x.Message + "<br>";
			}

			//Bind to grid
			try
			{
				grdBooks.DataSource=dataSet;
				grdBooks.DataMember="books";
				grdBooks.DataBind();
			}
			catch(Exception ex)
			{
				lblMessage.Text += ex.Message + "<br>";
			}

		}

		protected override void RenderWebPart(HtmlTextWriter output)
		{

			//Draw the control
			grdBooks.RenderControl(output);
			output.Write("<br>");
			lblMessage.RenderControl(output);
		}

	}

}


using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Web;
using System.Web.Services;
using System.Configuration;
using System.IO;
using System.Xml;	

namespace ResearchPubs
{
	[WebService(Namespace="urn:Microsoft.Search")]
	public class Service1 : System.Web.Services.WebService
	{

		[WebMethod()]public string Registration(string registrationxml)
		{
			//Key path information
			string templatePath = 
				HttpContext.Current.Server.MapPath(".").ToString()
				+ "\\RegistrationResponse.xml";
			string servicePath = "http://spsportal:8888/ResearchPubs/Service1.asmx";

			//Load template
			XmlDocument outXML = new XmlDocument();
			outXML.Load(templatePath);

			//Prepare to modify template
			XmlNamespaceManager manager = 
				new XmlNamespaceManager(outXML.NameTable);
			manager.AddNamespace("ns", "urn:Microsoft.Search.Registration.Response");

			//Modify XML
			outXML.SelectSingleNode("//ns:QueryPath", manager).InnerText = servicePath;
			outXML.SelectSingleNode("//ns:RegistrationPath", manager).InnerText = servicePath;
			outXML.SelectSingleNode("//ns:AboutPath", manager).InnerText = servicePath;

			return outXML.InnerXml.ToString();

		}

		[WebMethod()]public string Query(string queryXml)
		{
			//The query text from the Research Library
			string queryText="";

			//Key path information
			string templatePath = 
				HttpContext.Current.Server.MapPath(".").ToString()
				+ "\\QueryResponse.xml";

			//Load incoming XML into a document
			XmlDocument inXMLDoc = new XmlDocument();

			try
			{
				if (queryXml.Length > 0)
				{
					inXMLDoc.LoadXml(queryXml.ToString());

					//Prepare to parse incoming XML
					XmlNamespaceManager inManager = 
						new XmlNamespaceManager(inXMLDoc.NameTable);
					inManager.AddNamespace("ns", "urn:Microsoft.Search.Query");
					inManager.AddNamespace("oc", "urn:Microsoft.Search.Query.Office.Context");

					//Parse out quert text
					queryText = inXMLDoc.SelectSingleNode("//ns:QueryText", inManager).InnerText;
				}
			}
			catch{queryText="";}

			//Load response template
			XmlDocument outXML = new XmlDocument();
			outXML.Load(templatePath);

			//Prepare to modify template
			XmlNamespaceManager outManager = 
				new XmlNamespaceManager(outXML.NameTable);
			outManager.AddNamespace("ns", "urn:Microsoft.Search.Response");

			//Add results
			outXML.SelectSingleNode("//ns:Range",outManager).InnerXml = getResults(queryText);
    
			//Return XML stream
			return outXML.InnerXml.ToString();
		}

		public string getResults(string queryText)
		{

			//Credentials
			string userName="sa";
			string password="";
			string database="pubs";
			string sqlServer="(local)";

			//Build connection string
			string strConn = "Password=" + password + 
				";Persist Security Info=True;User ID=" + userName + 
				";Initial Catalog=" + database + ";Data Source=" + sqlServer;

			//Build SQL statement
			string strSQL = "SELECT pub_name, city, state FROM Publishers " +
				"WHERE pub_name LIKE '" + queryText + "%'";

			DataSet dataSet = new DataSet("publishers");

			//Run the query
			SqlConnection conn = new SqlConnection(strConn);
			SqlDataAdapter adapter = new SqlDataAdapter(strSQL,conn);
			adapter.Fill(dataSet,"publishers");

			//Build the results
			StringWriter stringWriter = new StringWriter();
			XmlTextWriter textWriter = new XmlTextWriter(stringWriter);
			DataTable dataTable = dataSet.Tables["publishers"];
			DataRowCollection dataRows = dataTable.Rows;

			textWriter.WriteElementString("StartAt", "1");
			textWriter.WriteElementString("Count", dataRows.Count.ToString());

			textWriter.WriteElementString("TotalAvailable", dataRows.Count.ToString());
			textWriter.WriteStartElement("Results");
			textWriter.WriteStartElement("Content", "urn:Microsoft.Search.Response.Content");

			foreach(DataRow dataRow in dataRows)
			{
				///Heading
				textWriter.WriteStartElement("Heading");
				textWriter.WriteAttributeString("collapsible", "true");
				textWriter.WriteAttributeString("collapsed", "true");
				textWriter.WriteElementString("Text", dataRow["pub_name"].ToString());

				//City
				textWriter.WriteStartElement("P");
				textWriter.WriteElementString("Char", dataRow["city"].ToString());
				textWriter.WriteStartElement("Actions");
				textWriter.WriteElementString("Insert", "");
				textWriter.WriteElementString("Copy", "");
				textWriter.WriteEndElement();
				textWriter.WriteEndElement();
    
				//State
				textWriter.WriteStartElement("P");
				textWriter.WriteElementString("Char", dataRow["state"].ToString());
				textWriter.WriteStartElement("Actions");
				textWriter.WriteElementString("Insert", "");
				textWriter.WriteElementString("Copy", "");
				textWriter.WriteEndElement();
				textWriter.WriteEndElement();

				textWriter.WriteEndElement();
			}

			textWriter.WriteEndElement();
			textWriter.WriteEndElement();
			textWriter.Close();

			return stringWriter.ToString();

		}

	}
}

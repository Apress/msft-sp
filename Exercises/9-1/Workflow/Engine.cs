using System;
using System.Windows.Forms;
using System.Xml;
using System.Security.Principal;
using System.Runtime.InteropServices;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Portal.SingleSignon;
namespace WorkFlow
{
	public class Engine:IListEventSink
	{
		protected static WindowsIdentity CreateIdentity
			(string userName, string domain, string password)
		{

			IntPtr tokenHandle = new IntPtr(0);
			tokenHandle=IntPtr.Zero;

			const int LOGON32_PROVIDER_DEFAULT=0;
			const int LOGON32_LOGON_NETWORK=3;

			//Logon the new user
			bool returnValue = LogonUser(userName,domain,password,
				LOGON32_LOGON_NETWORK,LOGON32_PROVIDER_DEFAULT,
				ref tokenHandle);

			if(returnValue==false)
			{
				int returnError = Marshal.GetLastWin32Error();
				throw new Exception("Log on failed: " + returnError);
			}

			//return new identity
			WindowsIdentity id = new WindowsIdentity(tokenHandle);
			CloseHandle(tokenHandle);
			return id;

		}

		[DllImport("advapi32.dll", SetLastError=true)]
		private static extern bool LogonUser
			(String lpszUsername, String lpszDomain, String lpszPassword,
			int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

		[DllImport("kernel32.dll", CharSet=CharSet.Auto)]
		private extern static bool CloseHandle(IntPtr handle);

		#region IListEventSink Members

		public void OnEvent(SPListEvent listEvent)
		{
			// Call MSSSO
			string[] strCredentials=null;

			//Credentials.GetCredentials
			//	(Convert.ToUInt32(1),"Workflow",ref strCredentials);
			string userName = "sab"; //strCredentials[0];
			string domain = "DATALANNT"; //strCredentials[1];
			string password = "honey1"; //strCredentials[2];

			//Create new context
			WindowsImpersonationContext windowsContext =
				CreateIdentity(userName,domain,password).Impersonate();

			//Get event objects
			SPWeb eventSite = listEvent.Site.OpenWeb();
			SPFile eventFile = eventSite.GetFile(listEvent.UrlAfter);
			SPListItem eventItem = eventFile.Item;

			//Determine if an approval event fired
			if((listEvent.Type == SPListEventType.Update) &&
				((string)eventItem["Approval Status"]=="0"))
			{

				//Load the XML document
				string xmlPath ="C:\\Workflow.xml";
				XmlDocument xmlDoc = new XmlDocument();
				xmlDoc.Load(xmlPath);

				//Prepare to parse XML
				XmlNamespaceManager manager = new XmlNamespaceManager(xmlDoc.NameTable);
				manager.AddNamespace("ns","urn:DataLan.SharePoint.WorkFlow.Engine");

				//Find the target library for the move
				string targetPath = xmlDoc.SelectSingleNode
					("//ns:" + listEvent.SinkData,manager).InnerText;

				//Move the document
				eventFile.MoveTo(targetPath + "/" + eventFile.Name,false);

			}

			//Tear down context
			windowsContext.Undo();
		}

		#endregion
	}
}

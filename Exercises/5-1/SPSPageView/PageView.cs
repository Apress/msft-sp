using System;
using System.ComponentModel;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml.Serialization;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebPartPages;

namespace SPSPageView
{

	[DefaultProperty("URL"),
	ToolboxData("<{0}:Container runat=server></{0}:Container>"),
	XmlRoot(Namespace="SPSPageView")]
	public class Container : Microsoft.SharePoint.WebPartPages.WebPart
	{
		private string url="";
		private int pageHeight = 400;

		[Browsable(true),Category("Miscellaneous"),
		DefaultValue(""),
		WebPartStorage(Storage.Personal),
		FriendlyName("URL"),Description("The address of the page to display")]
		public string URL
		{
			get
			{
				return url;
			}

			set
			{
				url = value;
			}
		}

		[Browsable(true),Category("Miscellaneous"),
		DefaultValue(400),
		WebPartStorage(Storage.Personal),
		FriendlyName("Page Height"),Description("The height of the page in pixels.")]
		public int PageHeight
		{
			get
			{
				return pageHeight;
			}

			set
			{
				pageHeight = value;
			}
		}


		protected override void RenderWebPart(HtmlTextWriter output)
		{
			output.Write("<div><iframe height='" + PageHeight + "' width='100%' src='" + URL + "'></iframe></div>");
		}
	}
}

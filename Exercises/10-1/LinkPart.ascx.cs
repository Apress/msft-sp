using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

public partial class LinkPart : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

	[Personalizable(),WebBrowsable(), WebDisplayName("LinkText"), WebDescription("The text of the link.")]
	public string linkText
	{
		get { return HyperLink1.Text; }
		set { HyperLink1.Text = value; }
	}

	[Personalizable(), WebBrowsable(), WebDisplayName("LinkURL"), WebDescription("The URL for the link.")]
	public string linkUrl
	{
		get { return HyperLink1.NavigateUrl; }
		set { HyperLink1.NavigateUrl = value; }
	}
}

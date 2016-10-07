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

public partial class TextPart : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

	[Personalizable(), WebBrowsable(), WebDisplayName("CardText"), WebDescription("The text to display.")]
	public string cardText
	{
		get { return Label1.Text; }
		set { Label1.Text = value; }
	}
}

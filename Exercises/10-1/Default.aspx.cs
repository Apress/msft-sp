using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

public partial class _Default : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {
		if (WebPartManager1.Personalization.Scope == PersonalizationScope.Shared)
			WebPartManager1.Personalization.ToggleScope();
    }

	protected void Button1_Click(object sender, EventArgs e)
	{
		if(Button1.Text=="Design Card")
		{
			WebPartManager1.DisplayMode = WebPartManager.EditDisplayMode;
			Button1.Text="View Card";
		}
		else
		{
			WebPartManager1.DisplayMode = WebPartManager.BrowseDisplayMode;
			Button1.Text="Design Card";
		}

	}
}
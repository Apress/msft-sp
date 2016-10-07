using System;
using System.ComponentModel;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml.Serialization;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebPartPages;
using Microsoft.SharePoint.WebControls;


namespace SPSIdentity
{

	[DefaultProperty(""),
	ToolboxData("<{0}:Report runat=server></{0}:Report>"),
	XmlRoot(Namespace="SPSIdentity")]
	public class Reporter : Microsoft.SharePoint.WebPartPages.WebPart
	{
		protected Label userNameLabel;
		protected TextBox displayNameText;
		protected TextBox emailText;
		protected Button updateButton;
		protected Label messageLabel;

		protected override void CreateChildControls()
		{

			//UserName Label
			userNameLabel = new Label();
			userNameLabel.Width = Unit.Percentage(100);
			userNameLabel.Font.Size = FontUnit.Point(10);
			userNameLabel.Font.Name = "arial";
			Controls.Add(userNameLabel);

			//DisplayName Text
			displayNameText = new TextBox();
			displayNameText.Width = Unit.Percentage(100);
			displayNameText.Font.Name = "arial";
			displayNameText.Font.Size = FontUnit.Point(10);
			Controls.Add(displayNameText);

			//E-Mail Text
			emailText = new TextBox();
			emailText.Width = Unit.Percentage(100);
			emailText.Font.Name = "arial";
			emailText.Font.Size = FontUnit.Point(10);
			Controls.Add(emailText);

			//Submit Button
			updateButton = new Button();
			updateButton.Font.Name = "arial";
			updateButton.Font.Size = FontUnit.Point(10);
			updateButton.Text = "Change";
			Controls.Add(updateButton);
			updateButton.Click +=new EventHandler(update_Click);

			//Message Label
			messageLabel = new Label();
			messageLabel.Width = Unit.Percentage(100);
			messageLabel.Font.Size = FontUnit.Point(10);
			messageLabel.Font.Name = "arial";
			Controls.Add(messageLabel);
		}


		protected override void RenderWebPart(HtmlTextWriter output)
		{
			//Get current user information before the context is changed
			SPSite site = SPControl.GetContextSite(Context);
			SPWeb web = site.OpenWeb();
			SPUser user = web.CurrentUser;

			//Show user information
			userNameLabel.Text = user.LoginName;
			displayNameText.Text = user.Name;
			emailText.Text = user.Email;

			//Create output
			output.Write("<TABLE Border=0>");
			output.Write("<TR>");
			output.Write("<TD>User name: ");
			userNameLabel.RenderControl(output);
			output.Write("</TD>");
			output.Write("</TR>");
			output.Write("<TR>");
			output.Write("<TD> Display name: ");
			displayNameText.RenderControl(output);
			output.Write("</TD>");
			output.Write("</TR>");
			output.Write("<TR>");
			output.Write("<TD>e-Mail: ");
			emailText.RenderControl(output);
			output.Write("</TD>");
			output.Write("</TR>");
			output.Write("<TR>");
			output.Write("<TD>");
			updateButton.RenderControl(output);
			output.Write("</TD>");
			output.Write("</TR>");
			output.Write("<TR>");
			output.Write("<TD>");
			messageLabel.RenderControl(output);
			output.Write("</TD>");
			output.Write("</TR>");
			output.Write("</TABLE>");

			//close
			web.Close();
			site.Close();
		}

		private void update_Click(object sender, EventArgs e)
		{
			//Get current user information before the context is changed
			SPSite site = SPControl.GetContextSite(Context);
			SPWeb web = site.OpenWeb();
			SPUser user = web.CurrentUser;

			//Update current user information
			user.Email=emailText.Text;
			user.Name=displayNameText.Text;
			user.Update();
			web.Close();
			site.Close();

		}

	}
}

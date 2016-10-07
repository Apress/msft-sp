<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="Default.aspx.cs" Inherits="_Default" %>

<%@ Register Src="LogoPart.ascx" TagName="LogoPart" TagPrefix="uc1" %>
<%@ Register Src="LinkPart.ascx" TagName="LinkPart" TagPrefix="uc2" %>
<%@ Register Src="TextPart.ascx" TagName="TextPart" TagPrefix="uc3" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:WebPartManager ID="WebPartManager1" runat="server">
        </asp:WebPartManager>
    
    </div>
        <table BGColor="#EFF3FB">
            <tr>
                <td style="width: 201px" rowspan="2">
                    <asp:WebPartZone ID="WebPartZone1" runat="server" BorderColor="#CCCCCC" Font-Names="Verdana"
                        Padding="6" Width="207px">
                        <PartChromeStyle BackColor="#EFF3FB" BorderColor="#D1DDF1" Font-Names="Verdana" ForeColor="#333333" />
                        <MenuLabelHoverStyle ForeColor="#D1DDF1" />
                        <EmptyZoneTextStyle Font-Size="0.8em" />
                        <MenuLabelStyle ForeColor="White" />
                        <MenuVerbHoverStyle BackColor="#EFF3FB" BorderColor="#CCCCCC" BorderStyle="Solid"
                            BorderWidth="1px" ForeColor="#333333" />
                        <HeaderStyle Font-Size="0.7em" ForeColor="#CCCCCC" HorizontalAlign="Center" />
                        <ZoneTemplate>
                            <uc1:LogoPart ID="LogoPart1" runat="server" />
                        </ZoneTemplate>
                        <MenuVerbStyle BorderColor="#507CD1" BorderStyle="Solid" BorderWidth="1px" ForeColor="White" />
                        <PartStyle Font-Size="0.8em" ForeColor="#333333" />
                        <TitleBarVerbStyle Font-Size="0.6em" Font-Underline="False" ForeColor="White" />
                        <MenuPopupStyle BackColor="#507CD1" BorderColor="#CCCCCC" BorderWidth="1px" Font-Names="Verdana"
                            Font-Size="0.6em" />
                        <PartTitleStyle BackColor="#507CD1" Font-Bold="True" Font-Size="0.8em" ForeColor="White" />
                    </asp:WebPartZone>
                </td>
                <td style="width: 200px">
                    <asp:WebPartZone ID="WebPartZone3" runat="server" BorderColor="#CCCCCC" Font-Names="Verdana"
                        Padding="6" Width="202px">
                        <PartChromeStyle BackColor="#EFF3FB" BorderColor="#D1DDF1" Font-Names="Verdana" ForeColor="#333333" />
                        <MenuLabelHoverStyle ForeColor="#D1DDF1" />
                        <EmptyZoneTextStyle Font-Size="0.8em" />
                        <MenuLabelStyle ForeColor="White" />
                        <MenuVerbHoverStyle BackColor="#EFF3FB" BorderColor="#CCCCCC" BorderStyle="Solid"
                            BorderWidth="1px" ForeColor="#333333" />
                        <HeaderStyle Font-Size="0.7em" ForeColor="#CCCCCC" HorizontalAlign="Center" />
                        <ZoneTemplate>
                            <uc3:TextPart ID="TextPart1" runat="server" />
                        </ZoneTemplate>
                        <MenuVerbStyle BorderColor="#507CD1" BorderStyle="Solid" BorderWidth="1px" ForeColor="White" />
                        <PartStyle Font-Size="0.8em" ForeColor="#333333" />
                        <TitleBarVerbStyle Font-Size="0.6em" Font-Underline="False" ForeColor="White" />
                        <MenuPopupStyle BackColor="#507CD1" BorderColor="#CCCCCC" BorderWidth="1px" Font-Names="Verdana"
                            Font-Size="0.6em" />
                        <PartTitleStyle BackColor="#507CD1" Font-Bold="True" Font-Size="0.8em" ForeColor="White" />
                    </asp:WebPartZone>
                </td>
            </tr>
            <tr>
                <td style="width: 200px">
                    <asp:WebPartZone ID="WebPartZone4" runat="server" BorderColor="#CCCCCC" Font-Names="Verdana"
                        Padding="6" Width="204px">
                        <PartChromeStyle BackColor="#EFF3FB" BorderColor="#D1DDF1" Font-Names="Verdana" ForeColor="#333333" />
                        <MenuLabelHoverStyle ForeColor="#D1DDF1" />
                        <EmptyZoneTextStyle Font-Size="0.8em" />
                        <MenuLabelStyle ForeColor="White" />
                        <MenuVerbHoverStyle BackColor="#EFF3FB" BorderColor="#CCCCCC" BorderStyle="Solid"
                            BorderWidth="1px" ForeColor="#333333" />
                        <HeaderStyle Font-Size="0.7em" ForeColor="#CCCCCC" HorizontalAlign="Center" />
                        <ZoneTemplate>
                            <uc3:TextPart ID="TextPart2" runat="server" />
                        </ZoneTemplate>
                        <MenuVerbStyle BorderColor="#507CD1" BorderStyle="Solid" BorderWidth="1px" ForeColor="White" />
                        <PartStyle Font-Size="0.8em" ForeColor="#333333" />
                        <TitleBarVerbStyle Font-Size="0.6em" Font-Underline="False" ForeColor="White" />
                        <MenuPopupStyle BackColor="#507CD1" BorderColor="#CCCCCC" BorderWidth="1px" Font-Names="Verdana"
                            Font-Size="0.6em" />
                        <PartTitleStyle BackColor="#507CD1" Font-Bold="True" Font-Size="0.8em" ForeColor="White" />
                    </asp:WebPartZone>
                </td>
            </tr>
            <tr>
                <td style="width: 201px" rowspan="2">
                    <asp:WebPartZone ID="WebPartZone2" runat="server" BorderColor="#CCCCCC" Font-Names="Verdana"
                        Padding="6" Width="206px">
                        <PartChromeStyle BackColor="#EFF3FB" BorderColor="#D1DDF1" Font-Names="Verdana" ForeColor="#333333" />
                        <MenuLabelHoverStyle ForeColor="#D1DDF1" />
                        <EmptyZoneTextStyle Font-Size="0.8em" />
                        <MenuLabelStyle ForeColor="White" />
                        <MenuVerbHoverStyle BackColor="#EFF3FB" BorderColor="#CCCCCC" BorderStyle="Solid"
                            BorderWidth="1px" ForeColor="#333333" />
                        <HeaderStyle Font-Size="0.7em" ForeColor="#CCCCCC" HorizontalAlign="Center" />
                        <ZoneTemplate>
                            <uc3:TextPart ID="TextPart4" runat="server" />
                        </ZoneTemplate>
                        <MenuVerbStyle BorderColor="#507CD1" BorderStyle="Solid" BorderWidth="1px" ForeColor="White" />
                        <PartStyle Font-Size="0.8em" ForeColor="#333333" />
                        <TitleBarVerbStyle Font-Size="0.6em" Font-Underline="False" ForeColor="White" />
                        <MenuPopupStyle BackColor="#507CD1" BorderColor="#CCCCCC" BorderWidth="1px" Font-Names="Verdana"
                            Font-Size="0.6em" />
                        <PartTitleStyle BackColor="#507CD1" Font-Bold="True" Font-Size="0.8em" ForeColor="White" />
                    </asp:WebPartZone>
                </td>
                <td style="width: 200px">
                    <asp:WebPartZone ID="WebPartZone5" runat="server" BorderColor="#CCCCCC" Font-Names="Verdana"
                        Padding="6" Width="206px">
                        <PartChromeStyle BackColor="#EFF3FB" BorderColor="#D1DDF1" Font-Names="Verdana" ForeColor="#333333" />
                        <MenuLabelHoverStyle ForeColor="#D1DDF1" />
                        <EmptyZoneTextStyle Font-Size="0.8em" />
                        <MenuLabelStyle ForeColor="White" />
                        <MenuVerbHoverStyle BackColor="#EFF3FB" BorderColor="#CCCCCC" BorderStyle="Solid"
                            BorderWidth="1px" ForeColor="#333333" />
                        <HeaderStyle Font-Size="0.7em" ForeColor="#CCCCCC" HorizontalAlign="Center" />
                        <ZoneTemplate>
                            <uc3:TextPart ID="TextPart3" runat="server" />
                        </ZoneTemplate>
                        <MenuVerbStyle BorderColor="#507CD1" BorderStyle="Solid" BorderWidth="1px" ForeColor="White" />
                        <PartStyle Font-Size="0.8em" ForeColor="#333333" />
                        <TitleBarVerbStyle Font-Size="0.6em" Font-Underline="False" ForeColor="White" />
                        <MenuPopupStyle BackColor="#507CD1" BorderColor="#CCCCCC" BorderWidth="1px" Font-Names="Verdana"
                            Font-Size="0.6em" />
                        <PartTitleStyle BackColor="#507CD1" Font-Bold="True" Font-Size="0.8em" ForeColor="White" />
                    </asp:WebPartZone>
                </td>
            </tr>
            <tr>
                <td style="width: 200px">
                    <asp:WebPartZone ID="WebPartZone6" runat="server" BorderColor="#CCCCCC" Font-Names="Verdana"
                        Padding="6" Width="206px">
                        <PartChromeStyle BackColor="#EFF3FB" BorderColor="#D1DDF1" Font-Names="Verdana" ForeColor="#333333" />
                        <MenuLabelHoverStyle ForeColor="#D1DDF1" />
                        <EmptyZoneTextStyle Font-Size="0.8em" />
                        <MenuLabelStyle ForeColor="White" />
                        <MenuVerbHoverStyle BackColor="#EFF3FB" BorderColor="#CCCCCC" BorderStyle="Solid"
                            BorderWidth="1px" ForeColor="#333333" />
                        <HeaderStyle Font-Size="0.7em" ForeColor="#CCCCCC" HorizontalAlign="Center" />
                        <ZoneTemplate>
                            <uc2:LinkPart ID="LinkPart1" runat="server" />
                        </ZoneTemplate>
                        <MenuVerbStyle BorderColor="#507CD1" BorderStyle="Solid" BorderWidth="1px" ForeColor="White" />
                        <PartStyle Font-Size="0.8em" ForeColor="#333333" />
                        <TitleBarVerbStyle Font-Size="0.6em" Font-Underline="False" ForeColor="White" />
                        <MenuPopupStyle BackColor="#507CD1" BorderColor="#CCCCCC" BorderWidth="1px" Font-Names="Verdana"
                            Font-Size="0.6em" />
                        <PartTitleStyle BackColor="#507CD1" Font-Bold="True" Font-Size="0.8em" ForeColor="White" />
                    </asp:WebPartZone>
                </td>
            </tr>
        </table>
        <asp:EditorZone ID="EditorZone1" runat="server" BackColor="#EFF3FB" BorderColor="#CCCCCC"
            BorderWidth="1px" Font-Names="Verdana" Padding="6">
            <LabelStyle Font-Size="0.8em" ForeColor="#333333" />
            <HeaderStyle BackColor="#D1DDF1" Font-Bold="True" Font-Size="0.8em" ForeColor="#333333" />
            <ZoneTemplate>
                <asp:PropertyGridEditorPart ID="PropertyGridEditorPart1" runat="server" />
                <asp:AppearanceEditorPart ID="AppearanceEditorPart1" runat="server" />
            </ZoneTemplate>
            <HeaderVerbStyle Font-Bold="False" Font-Size="0.8em" Font-Underline="False" ForeColor="#333333" />
            <PartChromeStyle BorderColor="#D1DDF1" BorderStyle="Solid" BorderWidth="1px" />
            <PartStyle BorderColor="#EFF3FB" BorderWidth="5px" />
            <FooterStyle BackColor="#D1DDF1" HorizontalAlign="Right" />
            <EditUIStyle Font-Names="Verdana" Font-Size="0.8em" ForeColor="#333333" />
            <InstructionTextStyle Font-Size="0.8em" ForeColor="#333333" />
            <ErrorStyle Font-Size="0.8em" />
            <VerbStyle Font-Names="Verdana" Font-Size="0.8em" ForeColor="#333333" />
            <EmptyZoneTextStyle Font-Size="0.8em" ForeColor="#333333" />
            <PartTitleStyle Font-Bold="True" Font-Size="0.8em" ForeColor="#333333" />
        </asp:EditorZone>
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Design Card" />
    </form>
</body>
</html>

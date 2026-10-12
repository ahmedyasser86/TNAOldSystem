<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="TNASettings_WP.ascx.cs" Inherits="NileSugar.HrTNA.Webparts.TNASettings_WP.TNASettings_WP" %>

<style type="text/css">
    .labelMsg {
        font-size: 20px;
    }

    .backColor {
        background-color: #c8b16a;
    }

    .Labels {
        display: inline-block;
        width: 30%;
        font-weight: 900;
        text-align: left;
        font-size: 50px;
    }

    .smallLabels {
        display: inline-block;
        width: 12%;
        font-weight: 900;
        text-align: left;
        font-size: 50px;
    }

    .reqVal {
        width: 12%;
        font-weight: 600;
        text-align: left;
        font-size: 20px;
    }

    .attach {
        display: inline-block;
        width: 50%;
        font-weight: 100;
        text-align: left;
        font-size: 15px;
    }

    .mybtnstyle {
        width: 100px;
        font-size: 20px;
    }

    .mybtnLargestyle {
        font-size: 20px;
    }

    .labelMsg {
        font-size: 20px;
    }
</style>

<div style="background-image: url('http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/bgImg.jpg');">
    <div style="font-size: 20px; text-align: right; color: black"><b>F-7-3.1   REV-0/2</b> </div>
    <br />
    <div style="width: 50%">

        <fieldset>
            <legend style="font-size: 15px">TNA Setting</legend>
            <br />
            <table style="width: 60%" cellpadding="7" border="1">
                <tr>
                    <td>TNA For Year :</td>

                    <td>
                        <asp:DropDownList ID="drpLstTNAYear" runat="server">
                            <%--<asp:ListItem>2020</asp:ListItem>--%>
                        </asp:DropDownList>
                    </td>
                </tr>
            </table>
        </fieldset>
        <br />
        <table style="width: 100%" cellpadding="7">
            <tr>
                <td colspan="2" style="text-align: center">
                    <br />
                    <asp:Button ID="btnSendEmail" runat="server" Text="Send Start TNA Email" OnClick="btnSendEmail_OnClick" CssClass="mybtnLargestyle" />
                    <br />
                    <br />
                    <asp:TextBox ID="TextBox1" runat="server" Visible="false"></asp:TextBox>
                    <br />
                    <asp:Button ID="Button1" runat="server" Text="Save Name" OnClick="Button1_Click" Visible="false" />
                    <br />
                    <asp:Label ID="lblAllEmps" runat="server" Text="Label"></asp:Label>
                              <br />
                    <asp:Label ID="lblMgrs" runat="server" Text="Label"></asp:Label>
                </td>
            </tr>
        </table>

    </div>

    <div id="div_bgLarge" style="height: 250px;" runat="server"></div>
    <div style="font-size: 20px; text-align: right; color: black"><b>F-7-3.1   REV-0/2</b> </div>
</div>





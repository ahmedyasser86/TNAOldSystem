<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Update_WP.ascx.cs" Inherits="NileSugar.HrTNA.Webparts.Update_WP.Update_WP" %>

Update EmpTechnicalCoursesLst :
<asp:Button ID="btn_UpdateEmpTechnicalCoursesLst" runat="server" Text="Update" OnClick="btn_UpdateEmpTechnicalCoursesLst_Click" />
<br />
<br />

Update EmpDevSkillsOtherCourses Lst :
<asp:Button ID="btn_updateEmpDevSkillsOtherCourses" runat="server" Text="Update" OnClick="btn_updateEmpDevSkillsOtherCourses_Click" />
<br />
<p>
    <asp:Label ID="lblMsg" runat="server"></asp:Label>
</p>

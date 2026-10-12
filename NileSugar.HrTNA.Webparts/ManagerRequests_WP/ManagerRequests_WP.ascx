<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ManagerRequests_WP.ascx.cs" Inherits="NileSugar.HrTNA.Webparts.ManagerRequests_WP.ManagerRequests_WP" %>

<%-- <body background="http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/bgImg.jpg">--%>

<div style="background-image: url('http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/bgImg.jpg');">
     <div style="font-size: 20px; text-align: right; color: black"><b>F-7-3.1   REV-0/2</b> </div>
        <br />
    <div style="text-align: center">

        <h1 style="color: black">
            <u>View Your Requests</u></h1>
    </div>
    <br />
    <div style="font-size: 20px; color: black" id="div_noCourses" runat="server" visible="False">
        There is no Courses 
    </div>
    <br />
    <div style="width: 100%" id="div_SupervisorReqs" runat="server">

        <h3><u>Requests By Supervisor </u></h3>
        <asp:GridView ID="GV_SupervisorCourses" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
            BorderColor="#DEBA84" BorderWidth="1px" CellPadding="3" EnableModelValidation="True"
            EmptyDataText="There is no Courses" BorderStyle="None" CellSpacing="2">

            <Columns>

                <asp:TemplateField HeaderText="i">
                    <ItemTemplate>
                        <%# Container.DataItemIndex + 1 %>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Emp. Code">
                    <ItemTemplate>
                        <asp:Label ID="lbl_BS_Other_EmpCode" runat="server" Text='<%# Eval("EmpCode") %>'></asp:Label>

                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                
                <asp:TemplateField HeaderText="Emp. Name">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpCode" runat="server" Text='<%# Eval("EmpName") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Division">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpDivision" runat="server" Text='<%# Eval("Division") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Department">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpDept" runat="server" Text='<%# Eval("Department") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Request Status">
                    <ItemTemplate>
                        <asp:Label ID="lbl_Req_Status" runat="server" Text='<%# Eval("RequestStatus")%>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Details">
                    <ItemTemplate>
                        <asp:Label ID="lbl_CourseID" runat="server" Text='<%# Eval("ID") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_SupervisorCode" runat="server" Text='<%# Eval("SupervisorCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_MgrCode" runat="server" Text='<%# Eval("MgrCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_DHCode" runat="server" Text='<%# Eval("DHCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_DVMgrCode" runat="server" Text='<%# Eval("DVMgrCode") %>' Visible="False"></asp:Label>
                        <asp:LinkButton ID="lnkBtnDetails" runat="server" OnClick="lnkBtnDetailsSuperviors_OnClick">Details</asp:LinkButton>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>


            </Columns>

            <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
            <HeaderStyle BackColor="#A55129" Font-Bold="True" Height="30px" ForeColor="White" />
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
            <SelectedRowStyle BackColor="#738A9C" ForeColor="White" Font-Bold="True" />
        </asp:GridView>
        <br />
    </div>

    <div style="width: 100%" id="div_MgrReqs" runat="server">

        <h3><u>Requests By Manager</u></h3>

        <asp:GridView ID="GV_MgrCourses" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
            BorderColor="#DEBA84" BorderWidth="1px" CellPadding="3" EnableModelValidation="True"
            EmptyDataText="There is no Courses" BorderStyle="None" CellSpacing="2">

            <Columns>

                <asp:TemplateField HeaderText="i">
                    <ItemTemplate>

                        <asp:Label ID="lbl_index" runat="server" Text='<%# Container.DataItemIndex + 1 %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Emp. Code">
                    <ItemTemplate>
                        <asp:Label ID="lbl_BS_Other_EmpCode" runat="server" Text='<%# Eval("EmpCode") %>'></asp:Label>

                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Emp. Name">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpCode" runat="server" Text='<%# Eval("EmpName") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>


                <asp:TemplateField HeaderText="Division">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpDivision" runat="server" Text='<%# Eval("Division") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Department">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpDept" runat="server" Text='<%# Eval("Department") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Request Status">
                    <ItemTemplate>
                        <asp:Label ID="lbl_Req_Status" runat="server" Text='<%# Eval("RequestStatus")%>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Details">
                    <ItemTemplate>
                        <asp:Label ID="lbl_CourseID" runat="server" Text='<%# Eval("ID") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_SupervisorCode" runat="server" Text='<%# Eval("SupervisorCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_MgrCode" runat="server" Text='<%# Eval("MgrCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_DHCode" runat="server" Text='<%# Eval("DHCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_DVMgrCode" runat="server" Text='<%# Eval("DVMgrCode") %>' Visible="False"></asp:Label>
                        <asp:LinkButton ID="lnkBtnDetails" runat="server" OnClick="lnkBtnDetailsMgrs_OnClick" Target="_blank">Details</asp:LinkButton>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

            </Columns>

            <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
            <HeaderStyle BackColor="#A55129" Font-Bold="True" Height="30px" ForeColor="White" />
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
            <SelectedRowStyle BackColor="#738A9C" ForeColor="White" Font-Bold="True" />
        </asp:GridView>
        <br />
    </div>

    <div style="width: 100%" id="div_DHReqs" runat="server">

        <h3><u>Requests By Dept. Head</u></h3>

        <asp:GridView ID="GV_DHCourses" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
            BorderColor="#DEBA84" BorderWidth="1px" CellPadding="3" EnableModelValidation="True"
            EmptyDataText="There is no Courses" BorderStyle="None" CellSpacing="2">

            <Columns>

                <asp:TemplateField HeaderText="i">
                    <ItemTemplate>
                        <%# Container.DataItemIndex + 1 %>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Emp. Code">
                    <ItemTemplate>
                        <asp:Label ID="lbl_BS_Other_EmpCode" runat="server" Text='<%# Eval("EmpCode") %>'></asp:Label>

                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Emp. Name">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpCode" runat="server" Text='<%# Eval("EmpName") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Division">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpDivision" runat="server" Text='<%# Eval("Division") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Department">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpDept" runat="server" Text='<%# Eval("Department") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Request Status">
                    <ItemTemplate>
                        <asp:Label ID="lbl_Req_Status" runat="server" Text='<%# Eval("RequestStatus")%>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Details">
                    <ItemTemplate>
                        <asp:Label ID="lbl_CourseID" runat="server" Text='<%# Eval("ID") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_SupervisorCode" runat="server" Text='<%# Eval("SupervisorCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_MgrCode" runat="server" Text='<%# Eval("MgrCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_DHCode" runat="server" Text='<%# Eval("DHCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_DVMgrCode" runat="server" Text='<%# Eval("DVMgrCode") %>' Visible="False"></asp:Label>
                        <asp:LinkButton ID="lnkBtnDetails" runat="server" OnClick="lnkBtnDetailsDH_OnClick" Target="_blank">Details</asp:LinkButton>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

            </Columns>

            <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
            <HeaderStyle BackColor="#A55129" Font-Bold="True" Height="30px" ForeColor="White" />
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
            <SelectedRowStyle BackColor="#738A9C" ForeColor="White" Font-Bold="True" />
        </asp:GridView>
        <br />
    </div>

    <div style="width: 100%" id="div_DVMgrReqs" runat="server">

        <h3><u>Requests By Division Manager</u></h3>

        <asp:GridView ID="GV_DVMgrCourses" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
            BorderColor="#DEBA84" BorderWidth="1px" CellPadding="3" EnableModelValidation="True"
            EmptyDataText="There is no Courses" BorderStyle="None" CellSpacing="2">

            <Columns>

                <asp:TemplateField HeaderText="i">
                    <ItemTemplate>
                        <%# Container.DataItemIndex + 1 %>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Emp. Code">
                    <ItemTemplate>
                        <asp:Label ID="lbl_BS_Other_EmpCode" runat="server" Text='<%# Eval("EmpCode") %>'></asp:Label>

                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Emp. Name">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpCode" runat="server" Text='<%# Eval("EmpName") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Division">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpDivision" runat="server" Text='<%# Eval("Division") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Department">
                    <ItemTemplate>
                        <asp:Label ID="lbl_EmpDept" runat="server" Text='<%# Eval("Department") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Request Status">
                    <ItemTemplate>
                        <asp:Label ID="lbl_Req_Status" runat="server" Text='<%# Eval("RequestStatus")%>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Details">
                    <ItemTemplate>
                        <asp:Label ID="lbl_CourseID" runat="server" Text='<%# Eval("ID") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_SupervisorCode" runat="server" Text='<%# Eval("SupervisorCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_MgrCode" runat="server" Text='<%# Eval("MgrCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_DHCode" runat="server" Text='<%# Eval("DHCode") %>' Visible="False"></asp:Label>
                        <asp:Label ID="lbl_DVMgrCode" runat="server" Text='<%# Eval("DVMgrCode") %>' Visible="False"></asp:Label>
                        <asp:LinkButton ID="lnkBtnDetails" runat="server" OnClick="lnkBtnDetailsDVMgr_OnClick">Details</asp:LinkButton>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

            </Columns>

            <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
            <HeaderStyle BackColor="#A55129" Font-Bold="True" Height="30px" ForeColor="White" />
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
            <SelectedRowStyle BackColor="#738A9C" ForeColor="White" Font-Bold="True" />
        </asp:GridView>
        <br />
    </div>
    
    <div id="div_bgLarge" style="height: 500px;" runat="server" visible="False"></div>
    
        <div style="font-size: 20px; text-align: right; color: black"><b>F-7-3.1   REV-0/2</b> </div>
        <br /> 
</div>

<%--</body> --%>

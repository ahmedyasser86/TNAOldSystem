<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="EmployeeCoursesReport.ascx.cs" Inherits="NileSugar.HrTNA.Webparts.EmployeeCoursesReport.EmployeeCoursesReport" %>
<style type="text/css">
    @page {
        size: auto;
        margin: 0mm;
    }

    .dropdown {
        width: 280px;
        font-size: medium;
    }

    .lableHeader {
        Font-Size: 13px;
        font-weight: bold;
    }

    .tr {
        border: solid thin;
    }

    .labelTitle {
        font-size: large;
    }

    .textBox {
        width: 270px;
    }
</style>
<div id="Form_Content" runat="server">
    <div style="text-align: Center">
        <h1>
            <u>Employee Cources Report
            </u>
        </h1>
    </div>

    <div id="applyCSS">

        <table style="width: 65%;">
            <tr>
                <td style="padding: 8px;">

                    <asp:Label ID="lbl_Devision" runat="server" Text="Devision" CssClass="labelTitle"></asp:Label>

                </td>
                <td style="padding: 8px;">
                    <asp:DropDownList ID="ddl_Devision" runat="server" CssClass="dropdown" AutoPostBack="true" OnSelectedIndexChanged="ddl_Devision_SelectedIndexChanged">
                    </asp:DropDownList>
                </td>

            </tr>
            <tr>
                <td style="padding: 8px;">
                    <asp:Label ID="lbl_Department" runat="server" Text="Department" CssClass="labelTitle"></asp:Label>

                </td>
                <td style="padding: 8px;">
                    <asp:DropDownList ID="ddl_Department" runat="server" CssClass="dropdown" AutoPostBack="true" OnSelectedIndexChanged="ddl_Department_SelectedIndexChanged"></asp:DropDownList>
                </td>

            </tr>
            <tr>
                <td style="padding: 8px;">
                    <asp:Label ID="lbl_EmployeeName" runat="server" Text="Employee Name" CssClass="labelTitle"></asp:Label>

                </td>

                <td style="padding: 8px;">
                    <asp:DropDownList ID="ddl_EmpNameSearch" runat="server" CssClass="dropdown">
                    </asp:DropDownList>


                </td>

            </tr>

            <tr>
                <td style="padding: 8px;">
                    <asp:Label ID="lbl_EmployeeCode" runat="server" Text="Employee Code" CssClass="labelTitle"></asp:Label>

                </td>
                <td style="padding: 8px;">
                    <asp:TextBox ID="txt_EmpCodeSearch" runat="server" CssClass="textBox"></asp:TextBox>

                </td>
            </tr>
            <tr>
                <td style="padding: 8px;">
                    <asp:Label ID="lbl_Categories" runat="server" Text="Category of courses" CssClass="labelTitle"></asp:Label>

                </td>
                <td style="padding: 8px;">
                    <asp:DropDownList ID="ddl_CoursesCategory" CssClass="dropdown" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddl_CoursesCategory_SelectedIndexChanged">
                        <asp:ListItem Value="-1">--Select--</asp:ListItem>
                        <asp:ListItem Value="0">Basic Skills</asp:ListItem>
                        <asp:ListItem Value="1">Development Skills</asp:ListItem>
                        <asp:ListItem Value="2">Technical Skills</asp:ListItem>

                    </asp:DropDownList>
                </td>


            </tr>
            <tr>
                <td style="padding: 8px;">
                    <asp:Label ID="lbl_CourseName" runat="server" Text="Course Name" CssClass="labelTitle"></asp:Label>

                </td>
                <td style="padding: 8px;">
                    <asp:DropDownList ID="ddl_CourseName" runat="server" CssClass="dropdown">
                    </asp:DropDownList>
                </td>

            </tr>
            <tr>
                <td style="padding: 8px;">
                    <asp:Label ID="lbl_ExpectedDate" runat="server" Text="Expected Date" CssClass="labelTitle"></asp:Label>

                </td>
                <%--<td style="padding: 8px;">
                    <%--<SharePoint:DateTimeControl ID="DateTimePicker_ExpectedDate" runat="server" />--%>
                          <%--<asp:TextBox ID="txt_DateTimePicker_ExpectedDate" runat="server" CssClass="textBox"></asp:TextBox>
                </td>--%>
                   <td style="padding: 8px;">
                    <SharePoint:DateTimeControl ID="DateTimePicker_ExpectedDate" runat="server" />
                          <%--<asp:TextBox ID="txt_DateTimePicker_ExpectedDate" runat="server" CssClass="textBox"></asp:TextBox>--%>
                </td>

            </tr>
            <tr>
                <td colspan="2" style="text-align: center">
                    <asp:Button ID="btn_LoadData" runat="server" Text="Search" OnClick="btn_LoadData_Click" Font-Bold="true"
                        Width="103px" Height="35px" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Button ID="btn_Exceel" runat="server" Text="Export To Excel" Width="157px" Height="35px" Visible="false" OnClick="btn_Exceel_Click" Font-Bold="true" />
                </td>
            </tr>
            <tr>
                <td colspan="2" style="text-align: center">&nbsp;</td>
            </tr>
        </table>
    </div>

    <div align="left" id="div_print" style="width: 100%" runat="server">
        <div id="NoRecords" runat="server" visible="false">
            No records are available.
        </div>
        <asp:Repeater ID="Rpt_EmpCources" runat="server" OnItemDataBound="Rpt_EmpCources_ItemDataBound">
            <ItemTemplate>
                <table>
                    <tr>
                        <td>
                            <asp:Label ID="lbl_EmpName" runat="server" Text=" Courses for Employee: " CssClass="lableHeader"></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lbl_EmpNameValue" runat="server" Text='<%# Eval("EmpName") %>' CssClass="lableBind"></asp:Label>
                        </td>

                        <td>
                            <asp:Label ID="lbl_EmpCode" runat="server" Text="Employee Code : " CssClass="lableHeader"></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lbl_EmpCodeValue" runat="server" Text='<%#Eval("EmpCode") %>' CssClass="lableBind"></asp:Label>
                        </td>


                    </tr>
                    <tr>
                        <td colspan="4">
                            <br />
                            <asp:GridView ID="Gv_EmpData" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
                                BorderColor="#DEBA84" BorderWidth="1px" CellPadding="3" EnableModelValidation="True"
                                EmptyDataText="There is no courses" BorderStyle="None" CellSpacing="2">


                                <Columns>

                                    <asp:BoundField DataField="Position" ReadOnly="True" HeaderText="Employee Position">
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>


                                    <asp:BoundField DataField="Division" ReadOnly="True" HeaderText="Employee Division">
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>



                                    <asp:TemplateField HeaderText="Employee Department">
                                        <ItemTemplate>
                                            <asp:Label ID="lbl_BS_Other_CourseName" runat="server" Text='<%# Eval("Department") %>'></asp:Label>

                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:TemplateField>



                                    <asp:BoundField DataField="Section" ReadOnly="True" HeaderText="Employee Section">
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>

                                    <asp:BoundField DataField="RequestStatus" ReadOnly="True" HeaderText="Request Status">
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>

                                </Columns>
                                <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
                                <HeaderStyle BackColor="#A55129" Font-Bold="True" Height="30px" ForeColor="White" />
                                <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
                                <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
                                <SelectedRowStyle BackColor="#738A9C" ForeColor="White" Font-Bold="True" />
                            </asp:GridView>

                        </td>

                    </tr>

                </table>

                <br />
                <u>
                    <asp:Label ID="Header1" runat="server" Text="Basic Skills" CssClass="lableHeader"></asp:Label></u>

                <br />
                <br />

                <asp:GridView ID="Gv_EmpCources_BasicSkills" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
                    BorderColor="#DEBA84" BorderWidth="1px" CellPadding="3" EnableModelValidation="True"
                    EmptyDataText="There is no courses" BorderStyle="None" CellSpacing="2">
                    <AlternatingRowStyle BackColor="PaleGoldenrod" />
                    <Columns>

                        <asp:TemplateField HeaderText="i">
                            <ItemTemplate>
                                <asp:Label ID="lbl_index" runat="server" Text='<%# Container.DataItemIndex + 1 %>'></asp:Label>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:TemplateField>



                        <asp:BoundField DataField="CourseName" ReadOnly="True" HeaderText="Course Name">
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>


                        <asp:BoundField DataField="ExpectedDate" ReadOnly="True" HeaderText="Expected Date" DataFormatString="{0:dd/MM/yyyy }" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>


                        <asp:BoundField DataField="Reason" ReadOnly="True" HeaderText="Reason">
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>


                    </Columns>
                    <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
                    <HeaderStyle BackColor="#A55129" Font-Bold="True" Height="30px" ForeColor="White" />
                    <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
                    <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
                    <SelectedRowStyle BackColor="#738A9C" ForeColor="White" Font-Bold="True" />
                </asp:GridView>

                <br />
                <u>
                    <asp:Label ID="Header2" runat="server" Text="Development Skills" CssClass="lableHeader"></asp:Label></u>

                <br />
                <br />

                <asp:GridView ID="Gv_EmpCources_DevelopmentSkills" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
                    BorderColor="#DEBA84" BorderWidth="1px" CellPadding="3" EnableModelValidation="True"
                    EmptyDataText="There is no courses" BorderStyle="None" CellSpacing="2">
                    <Columns>
                        <asp:TemplateField HeaderText="i">
                            <ItemTemplate>
                                <%# Container.DataItemIndex + 1 %>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:TemplateField>

                        <asp:BoundField DataField="CourseName" ReadOnly="True" HeaderText="Course Name">
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>


                        <asp:BoundField DataField="ExpectedDate" ReadOnly="True" HeaderText="Expected Date" DataFormatString="{0:dd/MM/yyyy }" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>


                        <asp:BoundField DataField="Reason" ReadOnly="True" HeaderText="Reason">
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>

                    </Columns>
                    <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
                    <HeaderStyle BackColor="#A55129" Font-Bold="True" Height="30px" ForeColor="White" />
                    <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
                    <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
                    <SelectedRowStyle BackColor="#738A9C" ForeColor="White" Font-Bold="True" />

                </asp:GridView>
                <br />

                <u>
                    <asp:Label ID="Header3" runat="server" Text="Technical Skills" CssClass="lableHeader"></asp:Label></u>

                <br />
                <br />

                <asp:GridView ID="Gv_EmpCources_TechnicalSkills" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
                    BorderColor="#DEBA84" BorderWidth="1px" CellPadding="3" EnableModelValidation="True"
                    EmptyDataText="There is no courses" BorderStyle="None" CellSpacing="2">

                    <Columns>
                        <asp:TemplateField HeaderText="i">
                            <ItemTemplate>
                                <%# Container.DataItemIndex + 1 %>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:TemplateField>

                        <asp:BoundField DataField="CourseName" ReadOnly="True" HeaderText="Course Name">
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>


                        <asp:BoundField DataField="ExpectedDate" ReadOnly="True" HeaderText="Expected Date" DataFormatString="{0:dd/MM/yyyy }" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>


                        <asp:BoundField DataField="Reason" ReadOnly="True" HeaderText="Reason">
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>

                    </Columns>
                    <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
                    <HeaderStyle BackColor="#A55129" Font-Bold="True" Height="30px" ForeColor="White" />
                    <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
                    <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
                    <SelectedRowStyle BackColor="#738A9C" ForeColor="White" Font-Bold="True" />
                </asp:GridView>
                <br />
                <hr />
                <br />
            </ItemTemplate>
        </asp:Repeater>
    </div>
</div>

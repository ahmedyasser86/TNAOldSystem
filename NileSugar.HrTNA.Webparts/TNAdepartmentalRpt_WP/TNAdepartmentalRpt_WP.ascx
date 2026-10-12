<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="TNAdepartmentalRpt_WP.ascx.cs" Inherits="NileSugar.HrTNA.Webparts.TNAdepartmentalRpt_WP.TNAdepartmentalRpt_WP" %>
<html>

<head>
    <style type="text/css">
        #projectTable td {
            padding: 7px;
            text-align: left;
            font-size: 15px;
        }

        h1, h3 {
            color: black;
        }

        .DropDownList {
            width: 80%;
        }

        .mybtnstyle {
            width: 200px;
            font-size: 20px;
        }
    </style>


    <link rel="stylesheet" type="text/css" href="/SiteAssets/DateTimePicker/jquery.datetimepicker.css" />
    <script src="/SiteAssets/DateTimePicker/jquery.js"></script>
    <script src="/SiteAssets/DateTimePicker/jquery.datetimepicker.js"></script>

    <%--<script type="text/javascript">
        jQuery(function () {
            jQuery('#' + '<%=txt_DateTimePicker_From.ClientID%>').datetimepicker({
               format: 'Y/m/d',
               timepicker: true
           });

           jQuery('#' + '<%=txt_DateTimePicker_To.ClientID%>').datetimepicker({
               format: 'Y/m/d',
               timepicker: true
           });

          
       });
</script>--%>
</head>

<body>
    <div style="background-image: url('http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/bgImg.jpg');">
        <div>
            <div style="font-size: 20px; text-align: right; color: black"><b>F-7-3.1   REV-0/2</b> </div>
            <h1 align="center"><u>TNA Departmental Report
            </u>

            </h1>
            <br />

            <table border="1" id="projectTable" style="color: black;" width="50%">

                <tr style="background-color: #c8b16a;">
                    <td colspan="4">

                        <h3><u>Search By</u></h3>
                    </td>
                </tr>

                <tr>
                    <td colspan="2">Department :</td>
                    <td colspan="2">
                        <asp:DropDownList ID="ddl_Department" runat="server" CssClass="DropDownList"></asp:DropDownList>
                    </td>
                </tr>

                <tr>
                    <td colspan="2">Employee Code:</td>
                    <td colspan="2">
                        <asp:TextBox ID="txt_EmpCode" runat="server" Width="179px"></asp:TextBox>
                    </td>
                </tr>

                <tr>
                    <td colspan="2">Category of courses :</td>
                    <td colspan="2">
                        <asp:DropDownList ID="ddl_CoursesCategory" CssClass="DropDownList" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddl_CoursesCategory_SelectedIndexChanged">
                            <asp:ListItem Value="-1">--Select--</asp:ListItem>
                            <asp:ListItem Value="0">Basic Skills</asp:ListItem>
                            <asp:ListItem Value="1">Development Skills</asp:ListItem>
                            <asp:ListItem Value="2">Technical Skills</asp:ListItem>

                        </asp:DropDownList>
                    </td>
                </tr>

                <tr>
                    <td colspan="2">Course Name :</td>
                    <td colspan="2">
                        <asp:DropDownList ID="ddl_CourseName" runat="server" CssClass="DropDownList" Enabled="False">
                        </asp:DropDownList>
                    </td>
                </tr>

                <tr>
                    <td colspan="2">TNA Year :</td>
                    <td colspan="2">
                        <asp:DropDownList ID="ddl_TNAYear" runat="server" CssClass="DropDownList">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr runat="server" visible="false">
                    <td>From :</td>
                    <td>

                        <SharePoint:DateTimeControl ID="cal_From" runat="server" DateOnly="True" AutoPostBack="true" OnDateChanged="cal_From_DateChanged" />

                        <asp:Label ID="lbl_Fr0mMsg" runat="server" Text=""></asp:Label>

                    </td>
                    <td>To :</td>
                    <td>

                        <SharePoint:DateTimeControl ID="cal_To" runat="server" DateOnly="True" AutoPostBack="true" OnDateChanged="cal_To_DateChanged" />

                        <asp:Label ID="lbl_ToMsg" runat="server" Text=""></asp:Label>

                    </td>
                </tr>

                <tr>
                    <td colspan="4" style="text-align: center">
                        <asp:Button ID="btn_Search" runat="server" Text="Search" CssClass="mybtnexpstyle" OnClick="btn_Search_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                       <asp:Button ID="btn_Excel" runat="server" Text="Export To Excel" CssClass="mybtnexpstyle" OnClick="btn_Excel_Click" Visible="False" />
                    </td>
                </tr>

            </table>
            <br />
            <div align="left" id="div_print" style="width: 100%" runat="server">
                <div id="NoRecords" runat="server" visible="false">
                    <h3>No records are available.</h3>
                </div>

                <asp:Repeater ID="Rpt_EmpCources" runat="server" OnItemDataBound="Rpt_EmpCources_ItemDataBound">
                    <ItemTemplate>
                        <h3><u>
                            <asp:Label ID="lbl_EmpDepartment" runat="server" Text='<%# Eval("Department") %>'></asp:Label></u></h3>
                        <br />
                        <asp:GridView ID="GV_DepartmenCources" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
                            BorderColor="#DEBA84" BorderWidth="1px" CellPadding="3" EnableModelValidation="True"
                            EmptyDataText="There is no courses" BorderStyle="None" CellSpacing="2" AllowPaging="false">
                            <AlternatingRowStyle BackColor="white" />
                            <Columns>

                                <asp:TemplateField HeaderText="i">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_index" runat="server" Text='<%# Container.DataItemIndex + 1 %>'></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Employee Code">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_EmpCode" runat="server" Text='<%# Eval("EmpCode").ToString() %>'></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Name">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_EmpName" runat="server" Text='<%# Eval("EmpName").ToString() %>'></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Position">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_EmpPOSITION" runat="server" Text='<%# Eval("POSITION").ToString() %>'></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Courses Category">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_CoursesCategory" runat="server" Text='<%# Eval("Category").ToString() %>'></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>


                                <asp:TemplateField HeaderText="Course Name">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_CourseName" runat="server" Text='<%# Eval("CourseName").ToString() %>'></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>

                                <asp:BoundField DataField="CreatedDate" ReadOnly="True" HeaderText="Created Date" DataFormatString="{0:dd/MM/yyyy }" HtmlEncode="False">
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>

                                <%--        <asp:BoundField DataField="ExpectedQuarter" ReadOnly="True" HeaderText="Expected Date" DataFormatString="{0:dd/MM/yyyy }" HtmlEncode="False">
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>--%>

                                <%--             <asp:TemplateField HeaderText="Expected Quarter">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_ExpectedQuarter" runat="server"><%# Eval("ExpectedQuarter").ToString() %></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>--%>

                                <asp:TemplateField HeaderText="Expected Quarter">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_ExpectedQuarter" runat="server"><%# Eval("ExpectedQuarter") %></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Request Status">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_RequestStatus" runat="server"><%# GetRequestStatus(Eval("EmpCode").ToString()) %></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Priority">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_RequestPriority" runat="server"><%# Eval("Priority") %></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Provider if any">
                                    <ItemTemplate>
                                        <asp:Label ID="lbl_TrainingProvider" runat="server" Text='<%# Eval("TrainingProvider").ToString() %>'></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>


                                <%--  <asp:Label ID="lblTrainingStatus" runat="server" Text='<%# Eval("TrainingStatus").ToString() %>' Visible="False"></asp:Label>--%>
                                <%--     <asp:TemplateField HeaderText="Training Status">
                                    <ItemTemplate>
                                       
                                        <asp:DropDownList ID="drpLstItemStatus" runat="server">
                                            <asp:ListItem Value="Requested">Requested</asp:ListItem>
                                            <asp:ListItem Value="Done">Done</asp:ListItem>
                                            <asp:ListItem Value="Postponed">Postponed</asp:ListItem>
                                            <asp:ListItem Value="Canceled">Canceled</asp:ListItem>
                                        </asp:DropDownList>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Scheduled Date By HR">
                                    <ItemTemplate>
                                        <SharePoint:DateTimeControl ID="cal_HRScheduledDate" runat="server" DateOnly="True" />
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>--%>
                            </Columns>

                            <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
                            <HeaderStyle BackColor="#A55129" Font-Bold="True" Height="30px" ForeColor="White" />
                            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
                            <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
                            <SelectedRowStyle BackColor="#738A9C" ForeColor="White" Font-Bold="True" />
                        </asp:GridView>
                    </ItemTemplate>
                </asp:Repeater>

            </div>
            <br />
            <br />
            <div style="text-align: center">
                <asp:Button ID="btn_Save" runat="server" Text="Save" CssClass="mybtnstyle" OnClick="btn_Save_Click" Visible="false" />
            </div>

            <div style="font-size: 20px; text-align: right; color: black"><b>F-7-3.1   REV-0/2</b> </div>
        </div>
    </div>
</body>
</html>

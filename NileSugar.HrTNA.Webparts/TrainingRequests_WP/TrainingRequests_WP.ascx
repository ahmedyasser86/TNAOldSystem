<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="TrainingRequests_WP.ascx.cs" Inherits="NileSugar.HrTNA.Webparts.TrainingRequests_WP.TrainingRequests_WP" %>


<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Trainings Request form</title>

<style type="text/css">
    .mybtnstyle {
        width: 200px;
        font-size: 20px;
    }

    .auto-style1 {
        height: 75px;
    }

    .auto-style2 {
        width: 32px;
    }
</style>

<!-- Latest compiled and minified CSS -->
<%-- <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">--%>

<!-- Optional theme -->
<%--    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css">
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
    <link rel="stylesheet" href="http://nshosplab01:45238/TNA/SiteAssets/css/form.css">
    <script src="http://nshosplab01:45238/TNA/SiteAssets/js/form.js"></script>--%>



<%--   <script type="text/javascript" src="http://nshosplab01:45238/TNA/SiteAssets/jquery-1.7.js"></script>--%>


<asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>



        <script type="text/javascript" src="http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/jquery-1.7.js"></script>
        <script src="https://code.jquery.com/ui/1.8.1/jquery-ui.min.js"></script>
        <script type="text/javascript">
            function ConfirmCancel() {
                return confirm("Are you sure from deleting this course ?");
            }
        </script>

        <script>
            $(document).ready(function () {

                // Bind Basic Text Other
                $(document.getElementById('<%= txt_BS_OtherC_Name.ClientID %>')).autocomplete({

                    source: function (request, response) {

                        var term = request.term;

                        var restUrl = "http://sharepoint.ns.local/forms/hr/TNA/_vti_bin/ListData.svc/BasicSkillsOtherCourses?$filter=substringof('" + term + "',CourseName)";
                        var coll = new Array();
                        $.ajax({
                            cache: true,
                            type: "GET",
                            async: false,
                            dataType: "json",
                            url: restUrl,
                            success: function (data) {

                                var results = data.d.results;
                                for (att in results) {
                                    var object = results[att];
                                    for (itm in object) {
                                        if (itm == "CourseName") {
                                            coll.push(object[itm]);
                                        }
                                    }
                                }
                            }
                        });

                        response(coll);
                    }
                });

                // Bind Development Text Other
                $(document.getElementById('<%= txt_DevS_OtherC_Name.ClientID %>')).autocomplete({

                    source: function (request, response) {

                        var term = request.term;

                        var restUrl = "http://sharepoint.ns.local/forms/hr/TNA/_vti_bin/ListData.svc/DevSkillsOtherCourses?$filter=substringof('" + term + "',CourseName)";
                        var coll = new Array();
                        $.ajax({
                            cache: true,
                            type: "GET",
                            async: false,
                            dataType: "json",
                            url: restUrl,
                            success: function (data) {

                                var results = data.d.results;
                                for (att in results) {
                                    var object = results[att];
                                    for (itm in object) {
                                        if (itm == "CourseName") {
                                            coll.push(object[itm]);
                                        }
                                    }
                                }
                            }
                        });

                        response(coll);
                    }
                });

                // Bind Technical Text Other
                $(document.getElementById('<%= txt_Tech_OtherC_Name.ClientID %>')).autocomplete({

                    source: function (request, response) {

                        var term = request.term;

                        var restUrl = "http://sharepoint.ns.local/forms/hr/TNA/_vti_bin/ListData.svc/TechnicalCoursesLst?$filter=substringof('" + term + "',CourseName)";
                        var coll = new Array();
                        $.ajax({
                            cache: true,
                            type: "GET",
                            async: false,
                            dataType: "json",
                            url: restUrl,
                            success: function (data) {

                                var results = data.d.results;
                                for (att in results) {
                                    var object = results[att];
                                    for (itm in object) {
                                        if (itm == "CourseName") {
                                            coll.push(object[itm]);
                                        }
                                    }
                                }
                            }
                        });

                        response(coll);
                    }
                });

            });

        </script>


        <div style="background-image: url('http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/bgImg.jpg');">

            <div>
                <div style="font-size: 20px; text-align: right; color: black"><b>F-7-3.1   REV-0/2</b> </div>
                <br />
                <div style="text-align: center">
                    <h1 style="color: black"><u>Learning and Development Need for Staff and Specialist Level </u></h1>
                </div>
                <br />
                <div id="div1" style="width: 100%; text-align: left;" runat="server" dir="ltr">
                    <table style="width: 90%; color: black" cellpadding="7">

                        <tr>
                            <td style="width: 15%">
                                <h3>Name :  
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLstEmpNames" runat="server" AutoPostBack="True" Font-Size="11"
                                    Width="90%" OnSelectedIndexChanged="drpLstEmpNames_SelectedIndexChanged">
                                </asp:DropDownList>
                                &nbsp

            <asp:RequiredFieldValidator InitialValue="-1" ID="rfv_drpLstEmpNames" Display="Dynamic"
                ValidationGroup="save" runat="server" ControlToValidate="drpLstEmpNames"
                ErrorMessage="*"></asp:RequiredFieldValidator>
                            </td>
                            <td style="width: 15%">
                                <h3>Code : 
                      
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txtEmpCode" runat="server" ReadOnly="True" Width="90%" Font-Size="11"></asp:TextBox>
                            </td>
                        </tr>

                        <tr>
                            <td>
                                <h3>Division :
                  
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txtEmpDivision" runat="server" ReadOnly="True" Width="90%" Font-Size="11"></asp:TextBox>

                            </td>
                            <td>
                                <h3>Section :
                   
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txtEmpSection" runat="server" ReadOnly="True" Width="90%" Font-Size="11"></asp:TextBox>

                            </td>
                        </tr>

                        <tr>
                            <td>
                                <h3>Department : </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txtEmpDept" runat="server" ReadOnly="True" Width="90%" Font-Size="11"></asp:TextBox>
                            </td>

                            <td>
                                <h3>Position :
                   
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txtEmpPosition" runat="server" ReadOnly="True" Width="90%" Font-Size="11"></asp:TextBox>
                            </td>
                        </tr>

                    </table>
                </div>
                <br />

                <div style="text-align: center">

                    <h1 style="color: black">
                        <u>Basic Skills</u></h1>
                </div>
                <div>
                    <table style="width: 100%; color: black" cellpadding="7">

                        <tr runat="server" visible="false">
                            <td style="width: 20%">
                                <h3 style="color: black">English Courses :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLst_BS_EngC_Name" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>General English Course</asp:ListItem>
                                    <asp:ListItem>Business English Course</asp:ListItem>

                                </asp:DropDownList>
                            </td>

                            <td>
                                <h3 style="color: black">Expected Quarter :
                                </h3>
                            </td>
                            <td>
                                <%--<SharePoint:DateTimeControl ID="cal_BS_EngC_ExpectedDate" runat="server" DateOnly="True" />--%>

                                <asp:DropDownList ID="drp_BS_EngC_ExpectedQuarter" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Quarter 1 </asp:ListItem>
                                    <asp:ListItem>Quarter 2</asp:ListItem>
                                    <asp:ListItem>Quarter 3 </asp:ListItem>
                                    <asp:ListItem>Quarter 4 </asp:ListItem>
                                </asp:DropDownList>

                            </td>
                            <td>
                                <h3 style="color: black">Reason :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_BS_EngC_Reason" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>

                                <%--<asp:RequiredFieldValidator ID="rfv_txt_BS_EngC_Reason" runat="server" ErrorMessage="*" ControlToValidate="txt_BS_EngC_Reason"
                            ValidationGroup="save" Display="Dynamic"></asp:RequiredFieldValidator>--%>
                            </td>
                            <td>
                                <h3 style="color: black">Priority :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLst_BS_EngC_priority" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Priority 1</asp:ListItem>
                                    <asp:ListItem>Priority 2</asp:ListItem>
                                    <asp:ListItem>Priority 3</asp:ListItem>
                                    <asp:ListItem>Priority 4</asp:ListItem>
                                    <asp:ListItem>Priority 5</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>

                        <tr>
                            <td>
                                <h3 style="color: black">Computer Skills :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLst_BS_CompSkillsC_Name" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Word (Basic Level)</asp:ListItem>
                                    <asp:ListItem>Word (Advanced Level)</asp:ListItem>
                                    <asp:ListItem>Excel (Basic Level)</asp:ListItem>
                                    <asp:ListItem>Excel (Advanced Level)</asp:ListItem>
                                    <asp:ListItem>PowerPoint (Basic Level)</asp:ListItem>
                                    <asp:ListItem>PowerPoint (Advanced Level)</asp:ListItem>
                                </asp:DropDownList>
                            </td>

                            <td>
                                <h3 style="color: black">Expected Quarter :
                                </h3>
                            </td>
                            <td>
                                <%--<SharePoint:DateTimeControl ID="cal_BS_CompSkillsC_ExpectedDate" runat="server" DateOnly="True" />--%>

                                <asp:DropDownList ID="drp_BS_CompSkillsC_ExpectedQuarter" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Quarter 1 </asp:ListItem>
                                    <asp:ListItem>Quarter 2</asp:ListItem>
                                    <asp:ListItem>Quarter 3 </asp:ListItem>
                                    <asp:ListItem>Quarter 4 </asp:ListItem>
                                </asp:DropDownList>

                            </td>
                            <td>
                                <h3 style="color: black">Reason :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_BS_CompSkillsC_Reason" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>

                                <%--<asp:RequiredFieldValidator ID="rfv_txt_BS_CompSkillsC_Reason" runat="server" ErrorMessage="*" ControlToValidate="txt_BS_CompSkillsC_Reason"
                                    ValidationGroup="save" Display="Dynamic" Enabled="false"></asp:RequiredFieldValidator>--%>
                            </td>
                            <td>
                                <h3 style="color: black">Priority :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLst_BS_CompSkillsC_priority" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Priority 1</asp:ListItem>
                                    <asp:ListItem>Priority 2</asp:ListItem>
                                    <asp:ListItem>Priority 3</asp:ListItem>
                                    <asp:ListItem>Priority 4</asp:ListItem>
                                    <asp:ListItem>Priority 5</asp:ListItem>
                                </asp:DropDownList>

                                <%--                <asp:RequiredFieldValidator InitialValue="-1" ID="rfv_drpLst_BS_CompSkillsC_priority" Display="Dynamic"
                                    ValidationGroup="save" runat="server" ControlToValidate="drpLst_BS_CompSkillsC_priority"
                                    ErrorMessage="*" Enabled="false"></asp:RequiredFieldValidator>--%>

                            </td>
                        </tr>

                        <tr>
                            <td>
                                <h3 style="color: black">Organization Skills :
                                </h3>
                            </td>

                            <td>
                                <asp:DropDownList ID="drpLst_BS_BEPC_Name" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Communication</asp:ListItem>
                                    <asp:ListItem>Teamwork</asp:ListItem>
                                    <asp:ListItem>Presentation Skills</asp:ListItem>
                                    <asp:ListItem>Time Management </asp:ListItem>
                                </asp:DropDownList>
                            </td>

                            <td>
                                <h3 style="color: black">Expected Quarter :
                                </h3>
                            </td>
                            <td>
                                <%--<SharePoint:DateTimeControl ID="cal_BS_BEPC_ExpectedDate" runat="server" DateOnly="True" />--%>

                                <asp:DropDownList ID="drp_BS_BEPC_ExpectedQuarter" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Quarter 1 </asp:ListItem>
                                    <asp:ListItem>Quarter 2</asp:ListItem>
                                    <asp:ListItem>Quarter 3 </asp:ListItem>
                                    <asp:ListItem>Quarter 4 </asp:ListItem>
                                </asp:DropDownList>

                            </td>
                            <td>
                                <h3 style="color: black">Reason :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_BS_BEPC_Reason" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>

                                <%--   <asp:RequiredFieldValidator ID="rfv_txt_BS_BEPC_Reason" runat="server" ErrorMessage="*" ControlToValidate="txt_BS_BEPC_Reason"
                                    ValidationGroup="save" Display="Dynamic" Enabled="false"></asp:RequiredFieldValidator>--%>
                            </td>
                            <td>
                                <h3 style="color: black">Priority :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLst_BS_BEPC_priority" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Priority 1</asp:ListItem>
                                    <asp:ListItem>Priority 2</asp:ListItem>
                                    <asp:ListItem>Priority 3</asp:ListItem>
                                    <asp:ListItem>Priority 4</asp:ListItem>
                                    <asp:ListItem>Priority 5</asp:ListItem>
                                </asp:DropDownList>

                                <%--                <asp:RequiredFieldValidator InitialValue="-1" ID="rfv_drpLst_BS_BEPC_priority" Display="Dynamic"
                                    ValidationGroup="save" runat="server" ControlToValidate="drpLst_BS_BEPC_priority"
                                    ErrorMessage="*" Enabled="false"></asp:RequiredFieldValidator>--%>
                            </td>
                        </tr>

                        <tr style="background-color: #c8b16a;">
                            <td colspan="8">
                                <h2 style="color: black"><u>Other Courses : </u></h2>
                            </td>
                        </tr>

                        <tr>
                            <td>
                                <h3 style="color: black">Other :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_BS_OtherC_Name" runat="server" Width="90%" TextMode="MultiLine" Height="50px" Font-Size="11"></asp:TextBox>
                            </td>

                            <td>
                                <h3 style="color: black">Expected Quarter :
                                </h3>
                            </td>
                            <td>
                                <%--<SharePoint:DateTimeControl ID="cal_BS_OtherC_ExpectedDate" runat="server" DateOnly="True" />--%>

                                <asp:DropDownList ID="drp_BS_OtherC_ExpectedQuarter" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Quarter 1 </asp:ListItem>
                                    <asp:ListItem>Quarter 2</asp:ListItem>
                                    <asp:ListItem>Quarter 3 </asp:ListItem>
                                    <asp:ListItem>Quarter 4 </asp:ListItem>
                                </asp:DropDownList>

                            </td>
                            <td>
                                <h3 style="color: black">Reason :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_BS_OtherC_Reason" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>

                                <%--  <asp:RequiredFieldValidator ID="rfv_txt_BS_OtherC_Reason" runat="server" ErrorMessage="*" ControlToValidate="txt_BS_OtherC_Reason"
                                    ValidationGroup="save_BasicOther" Display="Dynamic" Enabled="false"></asp:RequiredFieldValidator>--%>
                            </td>
                            <td>
                                <h3 style="color: black">Priority :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLst_BS_OtherC_priority" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Priority 1</asp:ListItem>
                                    <asp:ListItem>Priority 2</asp:ListItem>
                                    <asp:ListItem>Priority 3</asp:ListItem>
                                    <asp:ListItem>Priority 4</asp:ListItem>
                                    <asp:ListItem>Priority 5</asp:ListItem>
                                </asp:DropDownList>

                                <%--        <asp:RequiredFieldValidator InitialValue="-1" ID="rfv_drpLst_BS_OtherC_priority" Display="Dynamic"
                                    ValidationGroup="save_BasicOther" runat="server" ControlToValidate="drpLst_BS_OtherC_priority"
                                    ErrorMessage="*" Enabled="false"></asp:RequiredFieldValidator>--%>

                            </td>
                        </tr>

                        <tr>
                            <td colspan="8" style="text-align: center">
                                <asp:ImageButton ID="btn_AddOtherCourse_BasicSkills" runat="server" Width="40px" Height="40px" ImageUrl="http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/AddIcon.png"
                                    OnClick="btn_AddOtherCourse_BasicSkills_OnClick" AlternateText="Add" ValidationGroup="save_BasicOther" />
                                &nbsp;
                        <h3 style="color: brown;">Save Other Course</h3>
                            </td>
                        </tr>
                    </table>
                </div>

                <div style="width: 100%">
                    <h3 style="color: black"><u>Basic Skills Courses</u></h3>

                    <asp:GridView ID="GV_BS_OtherCourses" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
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

                            <asp:TemplateField HeaderText="Emp. Code">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_BS_Other_EmpCode" runat="server" Text='<%# Eval("EmpCode") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Course Name">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_BS_Other_CourseName" runat="server" Text='<%# Eval("CourseName") %>'></asp:Label>
                                    <asp:Label ID="lbl_BS_CourseID" runat="server" Text='<%# Eval("ID") %>' Visible="False"></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>


                            <%-- <asp:TemplateField HeaderText="Expected Date">
                        <ItemTemplate>
                            <asp:Label ID="lbl_BS_Other_ExpectedDate" runat="server" Text='<%# Eval("ExpectedDate")%>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>--%>


                            <%-- <asp:BoundField DataField="ExpectedDate" ReadOnly="True" HeaderText="Expected Date" DataFormatString="{0:dd/MM/yyyy }" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>--%>


                            <asp:TemplateField HeaderText="Expected Quarter">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_BS_Other_Quarter" runat="server" Text='<%# Eval("ExpectedQuarter")%>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Reason">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_BS_Other_Reason" runat="server" Text='<%# Eval("Reason")%>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Priority">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_BS_Other_Priority" runat="server" Text='<%# Eval("Priority")%>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Delete">
                                <ItemTemplate>
                                    <asp:ImageButton ID="btnDeleteBSOtherCourse" runat="server" ImageUrl='<%# GetDeleteIconUrl() %>' Width="30px" Height="30px"
                                        OnClientClick="if (!ConfirmCancel()) return false;" OnClick="btnDeleteBSOtherCourse_OnClick" />
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

                <div style="text-align: center">

                    <h1 style="color: black">
                        <u>Development Skills</u></h1>
                </div>

                <div>
                    <table style="width: 100%; color: black" cellpadding="7">
                        <tr>
                            <td>
                                <h3 style="color: black">Behavioral Skills :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="DrpLst_DevS_CommSC_Name" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Emotional Intelligence</asp:ListItem>
                                    <asp:ListItem>Conflict Management </asp:ListItem>
                                    <asp:ListItem>Creative Thinking </asp:ListItem>
                                    <asp:ListItem>Internal Customer Service </asp:ListItem>
                                    <asp:ListItem>Negotiation Skills </asp:ListItem>
                                    <asp:ListItem>Problem Solving and Decision </asp:ListItem>
                                </asp:DropDownList>
                            </td>

                            <td>
                                <h3 style="color: black">Expected Quarter :
                                </h3>
                            </td>
                            <td>
                                <%--<SharePoint:DateTimeControl ID="cal_DevS_CommSC_ExpectedDate" runat="server" DateOnly="True" />--%>

                                <asp:DropDownList ID="drp_DevS_CommSC_ExpectedQuarter" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Quarter 1 </asp:ListItem>
                                    <asp:ListItem>Quarter 2</asp:ListItem>
                                    <asp:ListItem>Quarter 3 </asp:ListItem>
                                    <asp:ListItem>Quarter 4 </asp:ListItem>
                                </asp:DropDownList>

                            </td>
                            <td>
                                <h3 style="color: black">Reason :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_DevS_CommSC_Reason" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>

                                <%--          <asp:RequiredFieldValidator ID="rfv_txt_DevS_CommSC_Reason" runat="server" ErrorMessage="*" ControlToValidate="txt_DevS_CommSC_Reason"
                                    ValidationGroup="save" Display="Dynamic" Enabled="false"></asp:RequiredFieldValidator>--%>
                            </td>
                            <td>
                                <h3 style="color: black">Priority :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLst_DevS_CommSC_priority" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Priority 1</asp:ListItem>
                                    <asp:ListItem>Priority 2</asp:ListItem>
                                    <asp:ListItem>Priority 3</asp:ListItem>
                                    <asp:ListItem>Priority 4</asp:ListItem>
                                    <asp:ListItem>Priority 5</asp:ListItem>
                                </asp:DropDownList>

                                <%--<asp:RequiredFieldValidator InitialValue="-1" ID="rfv_drpLst_DevS_CommSC_priority" Display="Dynamic"
                                    ValidationGroup="save" runat="server" ControlToValidate="drpLst_DevS_CommSC_priority"
                                    ErrorMessage="*" Enabled="false"></asp:RequiredFieldValidator>--%>

                            </td>
                        </tr>

                        <tr>
                            <td>
                                <h3 style="color: black">Management Skills  :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="DrpLst_DevS_InterPersSC_Name" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Middle Management Program </asp:ListItem>
                                    <asp:ListItem>Leadership Workshop</asp:ListItem>
                                    <asp:ListItem>Coaching Skills</asp:ListItem>
                                    <asp:ListItem>Change Management </asp:ListItem>
                                    <asp:ListItem>Decision Making </asp:ListItem>
                                    <asp:ListItem>Strategic Management</asp:ListItem>
                                </asp:DropDownList>
                            </td>

                            <td>
                                <h3 style="color: black">Expected Quarter :
                                </h3>
                            </td>
                            <td>
                                <%--<SharePoint:DateTimeControl ID="cal_DevS_InterPersSC_ExpectedDate" runat="server" DateOnly="True" />--%>

                                <asp:DropDownList ID="drp_DevS_InterPersSC_ExpectedQuarter" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Quarter 1 </asp:ListItem>
                                    <asp:ListItem>Quarter 2</asp:ListItem>
                                    <asp:ListItem>Quarter 3 </asp:ListItem>
                                    <asp:ListItem>Quarter 4 </asp:ListItem>
                                </asp:DropDownList>
                            </td>
                            <td>
                                <h3 style="color: black">Reason :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_DevS_InterPersSC_Reason" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>

                                <%--  <asp:RequiredFieldValidator ID="rfv_txt_DevS_InterPersSC_Reason" runat="server" ErrorMessage="*" ControlToValidate="txt_DevS_InterPersSC_Reason"
                                    ValidationGroup="save" Display="Dynamic" Enabled="false"></asp:RequiredFieldValidator>--%>
                            </td>
                            <td>
                                <h3 style="color: black">Priority :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLst_DevS_InterPersSC_priority" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Priority 1</asp:ListItem>
                                    <asp:ListItem>Priority 2</asp:ListItem>
                                    <asp:ListItem>Priority 3</asp:ListItem>
                                    <asp:ListItem>Priority 4</asp:ListItem>
                                    <asp:ListItem>Priority 5</asp:ListItem>
                                </asp:DropDownList>

                                <%--             <asp:RequiredFieldValidator InitialValue="-1" ID="rfv_drpLst_DevS_InterPersSC_priority" Display="Dynamic"
                                    ValidationGroup="save" runat="server" ControlToValidate="drpLst_DevS_InterPersSC_priority"
                                    ErrorMessage="*" Enabled="false"></asp:RequiredFieldValidator>--%>

                            </td>
                        </tr>

                        <tr style="background-color: #c8b16a;">
                            <td colspan="8">
                                <h2 style="color: black"><u>Other Courses : </u></h2>
                            </td>
                        </tr>

                        <tr>
                            <td>
                                <h3 style="color: black">Other :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_DevS_OtherC_Name" runat="server" Width="90%" TextMode="MultiLine" Height="50px" Font-Size="11"></asp:TextBox>
                            </td>
                            <td>
                                <h3 style="color: black">Expected Quarter :
                                </h3>
                            </td>
                            <td>
                                <%--<SharePoint:DateTimeControl ID="cal_DevS_OtherC_ExpectedDate" runat="server" DateOnly="True" />--%>

                                <asp:DropDownList ID="drp_DevS_OtherC_ExpectedQuarter" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Quarter 1 </asp:ListItem>
                                    <asp:ListItem>Quarter 2</asp:ListItem>
                                    <asp:ListItem>Quarter 3 </asp:ListItem>
                                    <asp:ListItem>Quarter 4 </asp:ListItem>
                                </asp:DropDownList>

                            </td>
                            <td>
                                <h3 style="color: black">Reason :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_DevS_OtherC_Reason" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>

                                <%--       <asp:RequiredFieldValidator ID="rfv_txt_DevS_OtherC_Reason" runat="server" ErrorMessage="*" ControlToValidate="txt_DevS_OtherC_Reason"
                                    ValidationGroup="save_DevOther" Display="Dynamic" Enabled="false"></asp:RequiredFieldValidator>--%>
                            </td>
                            <td>
                                <h3 style="color: black">Priority :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLst_DevS_OtherC_priority" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Priority 1</asp:ListItem>
                                    <asp:ListItem>Priority 2</asp:ListItem>
                                    <asp:ListItem>Priority 3</asp:ListItem>
                                    <asp:ListItem>Priority 4</asp:ListItem>
                                    <asp:ListItem>Priority 5</asp:ListItem>
                                </asp:DropDownList>

                                <%--          <asp:RequiredFieldValidator InitialValue="-1" ID="rfv_drpLst_DevS_OtherC_priority" Display="Dynamic"
                                    ValidationGroup="save_DevOther" runat="server" ControlToValidate="drpLst_DevS_OtherC_priority"
                                    ErrorMessage="*" Enabled="false"></asp:RequiredFieldValidator>--%>
                            </td>

                        </tr>
                        <tr>
                            <td colspan="8" style="text-align: center">
                                <asp:ImageButton ID="btn_AddOtherCourse_Dev" runat="server" Width="40px" Height="40px" ImageUrl="http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/AddIcon.png"
                                    OnClick="btn_AddOtherCourse_Dev_OnClick" AlternateText="Add" ValidationGroup="save_DevOther" />
                                &nbsp;
                        <h3 style="color: brown;">Save Other Course</h3>
                            </td>
                        </tr>
                    </table>
                </div>

                <div style="width: 100%">
                    <h3 style="color: black"><u>Development Skills Courses</u></h3>

                    <asp:GridView ID="GV_Dev_OtherCourses" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
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

                            <asp:TemplateField HeaderText="Emp. Code">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_Dev_Other_EmpCode" runat="server" Text='<%# Eval("EmpCode") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Course Name">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_Dev_Other_CourseName" runat="server" Text='<%# Eval("CourseName") %>'></asp:Label>
                                    <asp:Label ID="lbl_Dev_CourseID" runat="server" Text='<%# Eval("ID") %>' Visible="False"></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>


                            <%-- <asp:TemplateField HeaderText="Expected Date">
                        <ItemTemplate>
                            <asp:Label ID="lbl_Dev_Other_ExpectedDate" runat="server" Text='<%# Eval("ExpectedDate")%>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>--%>

                            <%--          <asp:BoundField DataField="ExpectedDate" ReadOnly="True" HeaderText="Expected Date" DataFormatString="{0:dd/MM/yyyy }" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>--%>

                            <asp:TemplateField HeaderText="Expected Quarter">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_BS_Other_Reason" runat="server" Text='<%# Eval("ExpectedQuarter")%>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Reason">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_Dev_Other_Reason" runat="server" Text='<%# Eval("Reason")%>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Priority">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_Dev_Other_Priority" runat="server" Text='<%# Eval("Priority")%>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Delete">
                                <ItemTemplate>
                                    <asp:ImageButton ID="btnDeleteDevOtherCourse" runat="server" ImageUrl='<%# GetDeleteIconUrl() %>' Width="30px" Height="30px"
                                        OnClientClick="if (!ConfirmCancel()) return false;" OnClick="btnDeleteDevOtherCourse_OnClick" />
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

                <div style="text-align: center">

                    <h1 style="color: black">
                        <u>Technical Skills</u></h1>
                </div>
                <div>
                    <table style="width: 100%; color: black" cellpadding="7">

                        <%--<tr >
                    <td>
                        <h3 style="color: black">Course Name :
                        </h3>
                    </td>
                    <td>
                        <asp:DropDownList ID="DrpLst_Tech_OtherC_Names" runat="server" Width="90%" Font-Size="11">
                        </asp:DropDownList>
                    </td>

                    <td>
                        <h3 style="color: black">Expected Date :
                        </h3>
                    </td>
                    <td>
                        <SharePoint:DateTimeControl ID="cal_TechCourse_ExpectedDate" runat="server" DateOnly="True" />
                    </td>

                    <td>
                        <h3 style="color: black">Training Provider :
                        </h3>
                    </td>
                    <td>
                        <asp:TextBox ID="txt_TechCourse_TrainingProvider" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>
                    </td>

                    <td>
                        <h3 style="color: black">Reason :
                        </h3>
                    </td>
                    <td>
                        <asp:TextBox ID="txt_TechCourse_Reason" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>
                    </td>
                    <td></td>
                </tr>--%>

                        <tr>
                            <td>
                                <h3 style="color: black">Course Name :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_Tech_OtherC_Name" runat="server" Width="90%" TextMode="MultiLine"
                                    Height="50px" Font-Size="11"></asp:TextBox>
                            </td>

                            <td>
                                <h3 style="color: black">Expected Quarter :
                                </h3>
                            </td>
                            <td>
                                <%--<SharePoint:DateTimeControl ID="cal_Tech_OtherC_ExpectedDate" runat="server" DateOnly="True" />--%>

                                <asp:DropDownList ID="drp_Tech_OtherC_ExpectedQuarter" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Quarter 1 </asp:ListItem>
                                    <asp:ListItem>Quarter 2</asp:ListItem>
                                    <asp:ListItem>Quarter 3 </asp:ListItem>
                                    <asp:ListItem>Quarter 4 </asp:ListItem>
                                </asp:DropDownList>

                            </td>

                            <td>
                                <h3 style="color: black">Training Provider :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_Tech_OtherC_TrainingProvider" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>
                            </td>

                            <td>
                                <h3 style="color: black">Reason :
                                </h3>
                            </td>
                            <td>
                                <asp:TextBox ID="txt_Tech_OtherC_Reason" runat="server" Width="90%" TextMode="MultiLine" Height="60px" Font-Size="11"></asp:TextBox>

                                <%--<asp:RequiredFieldValidator ID="rfv_txt_Tech_OtherC_Reason" runat="server" ErrorMessage="*" ControlToValidate="txt_Tech_OtherC_Reason"
                                    ValidationGroup="save_TechOther" Display="Dynamic" Enabled="false"></asp:RequiredFieldValidator>--%>
                            </td>
                            <td>
                                <h3 style="color: black">Priority :
                                </h3>
                            </td>
                            <td>
                                <asp:DropDownList ID="drpLst_Tech_OtherC_priority" runat="server" Width="90%" Font-Size="11">
                                    <asp:ListItem Value="-1">-- Select --</asp:ListItem>
                                    <asp:ListItem>Priority 1</asp:ListItem>
                                    <asp:ListItem>Priority 2</asp:ListItem>
                                    <asp:ListItem>Priority 3</asp:ListItem>
                                    <asp:ListItem>Priority 4</asp:ListItem>
                                    <asp:ListItem>Priority 5</asp:ListItem>
                                </asp:DropDownList>

                                <%--   <asp:RequiredFieldValidator InitialValue="-1" ID="rfv_drpLst_Tech_OtherC_priority" Display="Dynamic"
                                    ValidationGroup="save_TechOther" runat="server" ControlToValidate="drpLst_Tech_OtherC_priority"
                                    ErrorMessage="*" Enabled="false"></asp:RequiredFieldValidator>--%>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="8" style="text-align: center">
                                <asp:ImageButton ID="btn_AddOther_TechCourse" runat="server" Width="40px" Height="40px" ImageUrl="http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/AddIcon.png"
                                    OnClick="btn_AddOther_TechCourse_OnClick" AlternateText="Add" ValidationGroup="save_TechOther" />
                                &nbsp;
                        <h3 style="color: brown;">Save Technical Course</h3>
                            </td>
                        </tr>
                    </table>
                </div>
                <br />

                <div style="width: 100%">
                    <h3 style="color: black"><u>Technical Courses</u></h3>

                    <asp:GridView ID="GV_Tech_OtherCourses" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" Width="100%"
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

                            <asp:TemplateField HeaderText="Dept. Name">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_Tech_Other_DeptName" runat="server" Text='<%# Eval("DeptName") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Emp. Code">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_Tech_Other_EmpCode" runat="server" Text='<%# Eval("EmpCode") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Course Name">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_Tech_Other_CourseName" runat="server" Text='<%# Eval("CourseName") %>'></asp:Label>
                                    <asp:Label ID="lbl_Tech_CourseID" runat="server" Text='<%# Eval("ID") %>' Visible="False"></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>


                            <%-- <asp:TemplateField HeaderText="Expected Date">
                        <ItemTemplate>
                            <asp:Label ID="lbl_Tech_Other_ExpectedDate" runat="server" Text='<%# Eval("ExpectedDate")%>' DataFormatString="{0:dd/MM/yyyy }" HtmlEncode="False"></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>--%>

                            <%--        <asp:BoundField DataField="ExpectedDate" ReadOnly="True" HeaderText="Expected Date" DataFormatString="{0:dd/MM/yyyy }" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>--%>

                            <asp:TemplateField HeaderText="Expected Quarter">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_BS_Other_Quarter" runat="server" Text='<%# Eval("ExpectedQuarter")%>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Training Provider">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_Tech_Other_TrainingProvider" runat="server" Text='<%# Eval("TrainingProvider")%>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Reason">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_Tech_Other_Reason" runat="server" Text='<%# Eval("Reason")%>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Priority">
                                <ItemTemplate>
                                    <asp:Label ID="lbl_Tech_Other_Priority" runat="server" Text='<%# Eval("Priority")%>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Delete">
                                <ItemTemplate>
                                    <asp:ImageButton ID="btnDeleteTechOtherCourse" runat="server" ImageUrl='<%# GetDeleteIconUrl() %>' Width="30px" Height="30px"
                                        OnClientClick="if (!ConfirmCancel()) return false;" OnClick="btnDeleteTechOtherCourse_OnClick" />
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


                <table style="width: 100%;">
                    <tr runat="server" id="tr_supervisor">
                        <td style="text-align: center">
                            <br />
                            <asp:Button ID="btnSaveTNAReq" runat="server" Text="Save" CssClass="mybtnstyle"
                                OnClick="btnSaveTNAReq_OnClick" ValidationGroup="save" />

                            <%--<asp:ValidationSummary ID="ValidationSummary1" runat="server" ForeColor="Red" />--%>

                            <br />
                            <br />
                        </td>
                    </tr>

                    <tr runat="server" id="tr_mgr" visible="false">
                        <td style="text-align: center" class="auto-style1">

                            <asp:Button ID="btnApproveTNAReqByMgr" runat="server" Text="Approve" CssClass="mybtnstyle"
                                OnClick="btnApproveTNAReqByMgr_OnClick" />

                            &nbsp; &nbsp; &nbsp; &nbsp;

                          <asp:Button ID="btnRejectTNAReqByMgr" runat="server" Text="Reject" CssClass="mybtnstyle"
                              OnClick="btnRejectTNAReqByMgr_OnClick" />

                            <br />
                            <br />
                        </td>
                    </tr>

                    <tr runat="server" id="tr_DH" visible="false">
                        <td style="text-align: center">

                            <asp:Button ID="btnApproveTNAReqByDH" runat="server" Text="Approve" CssClass="mybtnstyle"
                                OnClick="btnApproveTNAReqByDH_OnClick" />

                            &nbsp; &nbsp; &nbsp; &nbsp;

                          <asp:Button ID="btnRejectTNAReqByDH" runat="server" Text="Reject" CssClass="mybtnstyle"
                              OnClick="btnRejectTNAReqByDH_OnClick" />

                            <br />
                            <br />
                        </td>
                    </tr>

                    <tr runat="server" id="tr_TopMgr" visible="false">
                        <td style="text-align: center">

                            <asp:Button ID="btnApproveTNAReqByTopMgr" runat="server" Text="Approve" CssClass="mybtnstyle"
                                OnClick="btnApproveTNAReqByTopMgr_OnClick" />

                            &nbsp; &nbsp; &nbsp; &nbsp;

                          <asp:Button ID="btnRejectTNAReqByTopMgr" runat="server" Text="Reject" CssClass="mybtnstyle"
                              OnClick="btnRejectTNAReqByTopMgr_OnClick" />

                            <br />
                            <br />
                        </td>
                    </tr>

                    <tr runat="server" id="tr_Print" visible="false">
                        <td style="text-align: center">

                            <asp:Button ID="btn_Print" runat="server" Text="Print" CssClass="mybtnstyle" />


                            <br />
                            <br />
                        </td>
                    </tr>
                </table>
                <br />
                <div style="font-size: 20px; text-align: right; color: black"><b>F-7-3.1   REV-0/2</b> </div>
                <br />
            </div>

        </div>


    </ContentTemplate>
</asp:UpdatePanel>

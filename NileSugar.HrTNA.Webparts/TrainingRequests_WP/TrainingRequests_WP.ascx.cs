using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;


namespace NileSugar.HrTNA.Webparts.TrainingRequests_WP
{
    [ToolboxItemAttribute(false)]
    public partial class TrainingRequests_WP : WebPart
    {

        #region Vars

        public Hashtable MailTemplates;
        public static Employee_INFO_ConDataContext EmpDB = new Employee_INFO_ConDataContext();

        #endregion

        // Uncomment the following SecurityPermission attribute only when doing Performance Profiling on a farm solution
        // using the Instrumentation method, and then remove the SecurityPermission attribute when the code is ready
        // for production. Because the SecurityPermission attribute bypasses the security check for callers of
        // your constructor, it's not recommended for production purposes.
        // [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Assert, UnmanagedCode = true)]
        public TrainingRequests_WP()
        {
        }

        protected override void OnInit(EventArgs e)
        {
            base.OnInit(e);
            InitializeControl();
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
            if (currentAppUrl.Contains("http://factory.ns.local/"))
                HttpContext.Current.Response.Redirect("http://sharepoint.ns.local/forms/hr/TNA/SitePages/TrainingNeedsRequest.aspx");


            if (!Page.IsPostBack)
            {
                try
                {
                    //Redirect to Acees denied Page
                    /*Rehab 24-01-2021*/
                    if (CommonMethods.GetLookups("AppStatus") == "Closed")
                    {
                        string accessDenied = CommonMethods.GetLookups("AccessDeniedPage");
                        HttpContext.Current.Response.Redirect(accessDenied);
                    }

                    //*******************************************************

                    string empId = CommonMethods.GetLoginEmpCode();
                    bool isMgr = CommonMethods.CheckIfLoginEmpIsMgr(empId);
                    bool isHr = CheckIfLoginUserIsHr(empId);
                    if (isMgr || isHr || empId == "01071")
                    {
                        if (HttpContext.Current.Request.QueryString["pageTo"] != null && HttpContext.Current.Request.QueryString["TNAId"] != null)
                        {
                            DataTable dt_Emps = CommonMethods.GetEmpsByTopMgrCode(empId);
                            BindDrpLstEmpNames(dt_Emps);
                            drpLstEmpNames.Enabled = false;
                            rfv_drpLstEmpNames.Enabled = false;

                            // DisableAllControls(this);

                            string TNAId = HttpContext.Current.Request.QueryString["TNAId"];

                            if (HttpContext.Current.Request.QueryString["pageTo"] == "supervisor")
                            {
                                LoadRequestData(TNAId);

                                tr_supervisor.Visible = true;
                                tr_mgr.Visible = false;
                                tr_DH.Visible = false;
                                tr_TopMgr.Visible = false;
                            }

                            else if (HttpContext.Current.Request.QueryString["pageTo"] == "mgr")
                            {
                                LoadRequestData(TNAId);

                                tr_supervisor.Visible = false;
                                tr_mgr.Visible = true;
                                tr_DH.Visible = false;
                                tr_TopMgr.Visible = false;
                            }

                            else if (HttpContext.Current.Request.QueryString["pageTo"] == "DH")
                            {
                                DisableAllControls(this);
                                LoadRequestData(TNAId);

                                tr_supervisor.Visible = false;
                                tr_mgr.Visible = false;
                                tr_DH.Visible = true;
                                tr_TopMgr.Visible = false;
                            }

                            else if (HttpContext.Current.Request.QueryString["pageTo"] == "DVMgr")
                            {
                                DisableAllControls(this);
                                LoadRequestData(TNAId);

                                tr_supervisor.Visible = false;
                                tr_mgr.Visible = false;
                                tr_DH.Visible = false;
                                tr_TopMgr.Visible = true;
                            }

                            else if (HttpContext.Current.Request.QueryString["pageTo"] == "hr")
                            {
                                DisableAllControls(this);
                                LoadRequestData(TNAId);

                                tr_supervisor.Visible = false;
                                tr_mgr.Visible = false;
                                tr_DH.Visible = false;
                                tr_TopMgr.Visible = false;
                            }
                        }
                        else
                        {
                            // DataTable dt_Emps = CommonMethods.GetEmpsBySupervisorCode("01125");
                            DataTable dt_Emps = CommonMethods.GetEmpsBySupervisorCode(empId);
                            BindDrpLstEmpNames(dt_Emps);
                        }


                        //BindDrpLst_GeneralCourses("Basic Skills", "English Courses", drpLst_BS_EngC_Name);
                        //BindDrpLst_GeneralCourses("Basic Skills", "Computer Skills", drpLst_BS_CompSkillsC_Name);
                        //BindDrpLst_GeneralCourses("Basic Skills", "Business Etiquette", drpLst_BS_BEPC_Name);

                        //BindDrpLst_GeneralCourses("Development Skills", "Communication skills", DrpLst_DevS_CommSC_Name);
                        //BindDrpLst_GeneralCourses("Development Skills", "Interpersonnel skills", DrpLst_DevS_InterPersSC_Name);
                        //BindDrpLst_GeneralCourses("Development Skills", "Business Intelligence", DrpLst_DevS_BIC_Name);
                    }
                    else
                    {
                        //Redirect to Acees denied Page
                        string accessDeniedPage = CommonMethods.GetLookups("AccessDeniedPage");
                        HttpContext.Current.Response.Redirect(accessDeniedPage);
                    }
                }
                catch (Exception)
                {
                }
            }
        }


        public bool CheckIfLoginUserIsHr(string empCode)
        {
            bool isHrEmp = false;
            if ((empCode == CommonMethods.GetLookups("HrEmpCode1")) || (empCode == CommonMethods.GetLookups("HrEmpCode2"))
                 || (empCode == CommonMethods.GetLookups("HrEmpCode3")))
            {
                isHrEmp = true;
            }
            return isHrEmp;
        }

        private void DisableAllControls(Control control)
        {
            foreach (var c in control.Controls)
            {
                if (c is TextBox) ((TextBox)c).ReadOnly = true;
            }

            foreach (var d in control.Controls)
            {
                if (d is DropDownList) ((DropDownList)d).Enabled = false;
            }

            foreach (var cal in control.Controls)
            {
                if (cal is DateTimeControl) ((DateTimeControl)cal).Enabled = false;
            }

            GV_BS_OtherCourses.Columns[5].Visible = false;
            GV_Dev_OtherCourses.Columns[5].Visible = false;
            GV_Tech_OtherCourses.Columns[7].Visible = false;

            btn_AddOtherCourse_BasicSkills.Visible = false;
            btn_AddOtherCourse_Dev.Visible = false;
            btn_AddOther_TechCourse.Visible = false;
        }

        private void LoadRequestData(string TNAId)
        {
            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                if (EmpCoursesLst != null)
                {
                    SPListItem empCourseItem = EmpCoursesLst.GetItemById(int.Parse(TNAId));
                    string empId = empCourseItem["EmpCode"].ToString();

                    if (!string.IsNullOrEmpty(empId))
                    {
                        drpLstEmpNames.SelectedItem.Text = empCourseItem["EmpName"].ToString().Trim();
                        DisplayEmployeeData(empId);

                        BindGV_BS_OtherCourses(txtEmpCode.Text.Trim());
                        BindGV_Dev_OtherCourses(txtEmpCode.Text.Trim());
                        BindGV_Tech_OtherCourses(txtEmpDept.Text.Trim(), txtEmpCode.Text.Trim());

                        if (empCourseItem["BS_EngC_Name"] != null && !string.IsNullOrEmpty(empCourseItem["BS_EngC_Name"].ToString()))
                            drpLst_BS_EngC_Name.SelectedValue = empCourseItem["BS_EngC_Name"].ToString();

                        //if (empCourseItem["BS_EngC_ExpectedDate"] != null && !string.IsNullOrEmpty(empCourseItem["BS_EngC_ExpectedDate"].ToString()))
                        //    cal_BS_EngC_ExpectedDate.SelectedDate =
                        //        Convert.ToDateTime(empCourseItem["BS_EngC_ExpectedDate"].ToString()).Date;

                        if (empCourseItem["BS_EngC_ExpectedQuarter"] != null && !string.IsNullOrEmpty(empCourseItem["BS_EngC_ExpectedQuarter"].ToString()))
                            drp_BS_EngC_ExpectedQuarter.SelectedValue = empCourseItem["BS_EngC_ExpectedQuarter"].ToString();

                        if (empCourseItem["BS_EngC_Reason"] != null && !string.IsNullOrEmpty(empCourseItem["BS_EngC_Reason"].ToString()))
                            txt_BS_EngC_Reason.Text = empCourseItem["BS_EngC_Reason"].ToString();

                        if (empCourseItem["BS_EngC_priority"] != null && !string.IsNullOrEmpty(empCourseItem["BS_EngC_priority"].ToString()))
                            drpLst_BS_EngC_priority.SelectedValue = empCourseItem["BS_EngC_priority"].ToString();
                        //**
                        if (empCourseItem["BS_CompSkillsC_Name"] != null && !string.IsNullOrEmpty(empCourseItem["BS_CompSkillsC_Name"].ToString()))
                            drpLst_BS_CompSkillsC_Name.Text = empCourseItem["BS_CompSkillsC_Name"].ToString();

                        if (empCourseItem["BS_CompSkillsC_ExpectedQuarter"] != null && !string.IsNullOrEmpty(empCourseItem["BS_CompSkillsC_ExpectedQuarter"].ToString()))
                            drp_BS_CompSkillsC_ExpectedQuarter.SelectedValue = empCourseItem["BS_CompSkillsC_ExpectedQuarter"].ToString();

                        if (empCourseItem["BS_CompSkillsC_Reason"] != null && !string.IsNullOrEmpty(empCourseItem["BS_CompSkillsC_Reason"].ToString()))
                            txt_BS_CompSkillsC_Reason.Text = empCourseItem["BS_CompSkillsC_Reason"].ToString();

                        if (empCourseItem["BS_CompSkillsC_priority"] != null && !string.IsNullOrEmpty(empCourseItem["BS_CompSkillsC_priority"].ToString()))
                            drpLst_BS_CompSkillsC_priority.SelectedValue = empCourseItem["BS_CompSkillsC_priority"].ToString();

                        //**
                        if (empCourseItem["BS_BEPC_Name"] != null &&
                            !string.IsNullOrEmpty(empCourseItem["BS_BEPC_Name"].ToString()))
                            drpLst_BS_BEPC_Name.SelectedValue = empCourseItem["BS_BEPC_Name"].ToString();

                        if (empCourseItem["BS_BEPC_ExpectedQuarter"] != null && !string.IsNullOrEmpty(empCourseItem["BS_BEPC_ExpectedQuarter"].ToString()))
                            drp_BS_BEPC_ExpectedQuarter.SelectedValue = empCourseItem["BS_BEPC_ExpectedQuarter"].ToString();

                        if (empCourseItem["BS_BEPC_Reason"] != null && !string.IsNullOrEmpty(empCourseItem["BS_BEPC_Reason"].ToString()))
                            txt_BS_BEPC_Reason.Text = empCourseItem["BS_BEPC_Reason"].ToString();

                        if (empCourseItem["BS_BEPC_priority"] != null && !string.IsNullOrEmpty(empCourseItem["BS_BEPC_priority"].ToString()))
                            drpLst_BS_BEPC_priority.SelectedValue = empCourseItem["BS_BEPC_priority"].ToString();


                        //**

                        if (empCourseItem["DevS_CommSC_Name"] != null &&
                            !string.IsNullOrEmpty(empCourseItem["DevS_CommSC_Name"].ToString()))
                            DrpLst_DevS_CommSC_Name.SelectedValue = empCourseItem["DevS_CommSC_Name"].ToString();


                        //if (empCourseItem["DevS_CommSC_ExpectedDate"] != null &&
                        //  !string.IsNullOrEmpty(empCourseItem["DevS_CommSC_ExpectedDate"].ToString()))
                        //    cal_DevS_CommSC_ExpectedDate.SelectedDate = Convert.ToDateTime(empCourseItem["DevS_CommSC_ExpectedDate"].ToString()).Date;

                        if (empCourseItem["DevS_CommSC_ExpectedQuarter"] != null && !string.IsNullOrEmpty(empCourseItem["DevS_CommSC_ExpectedQuarter"].ToString()))
                            drp_DevS_CommSC_ExpectedQuarter.SelectedValue = empCourseItem["DevS_CommSC_ExpectedQuarter"].ToString();

                        if (empCourseItem["DevS_CommSC_Reason"] != null && !string.IsNullOrEmpty(empCourseItem["DevS_CommSC_Reason"].ToString()))
                            txt_DevS_CommSC_Reason.Text = empCourseItem["DevS_CommSC_Reason"].ToString();


                        if (empCourseItem["BS_CompSkillsC_priority"] != null && !string.IsNullOrEmpty(empCourseItem["BS_CompSkillsC_priority"].ToString()))
                            drpLst_DevS_CommSC_priority.SelectedValue = empCourseItem["BS_CompSkillsC_priority"].ToString();
                        //**

                        if (empCourseItem["DevS_InterPersSC_Name"] != null &&
                            !string.IsNullOrEmpty(empCourseItem["DevS_InterPersSC_Name"].ToString()))
                            DrpLst_DevS_InterPersSC_Name.SelectedValue =
                                empCourseItem["DevS_InterPersSC_Name"].ToString();

                        //if (empCourseItem["DevS_InterPersSC_ExpectedDate"] != null &&
                        //!string.IsNullOrEmpty(empCourseItem["DevS_InterPersSC_ExpectedDate"].ToString()))
                        //    cal_DevS_InterPersSC_ExpectedDate.SelectedDate = Convert.ToDateTime(empCourseItem["DevS_InterPersSC_ExpectedDate"].ToString()).Date;

                        if (empCourseItem["DevS_InterPersSC_ExpectedQuarter"] != null && !string.IsNullOrEmpty(empCourseItem["DevS_InterPersSC_ExpectedQuarter"].ToString()))
                            drp_DevS_InterPersSC_ExpectedQuarter.SelectedValue = empCourseItem["DevS_InterPersSC_ExpectedQuarter"].ToString();


                        if (empCourseItem["DevS_InterPersSC_Reason"] != null && !string.IsNullOrEmpty(empCourseItem["DevS_InterPersSC_Reason"].ToString()))
                            txt_DevS_InterPersSC_Reason.Text = empCourseItem["DevS_InterPersSC_Reason"].ToString();

                        if (empCourseItem["DevS_InterPersSC_priority"] != null && !string.IsNullOrEmpty(empCourseItem["DevS_InterPersSC_priority"].ToString()))
                            drpLst_DevS_InterPersSC_priority.SelectedValue = empCourseItem["DevS_InterPersSC_priority"].ToString();

                        //**

                        //if (empCourseItem["DevS_BIC_Name"] != null &&
                        //    !string.IsNullOrEmpty(empCourseItem["DevS_BIC_Name"].ToString()))
                        //    DrpLst_DevS_BIC_Name.SelectedValue = empCourseItem["DevS_BIC_Name"].ToString();


                        //if (empCourseItem["DevS_BIC_ExpectedDate"] != null && !string.IsNullOrEmpty(empCourseItem["DevS_BIC_ExpectedDate"].ToString()))
                        //    cal_DevS_BIC_ExpectedDate.SelectedDate = Convert.ToDateTime(empCourseItem["DevS_BIC_ExpectedDate"].ToString()).Date;

                        //if (empCourseItem["DevS_BIC_Reason"] != null && !string.IsNullOrEmpty(empCourseItem["DevS_BIC_Reason"].ToString()))
                        //    txt_DevS_BIC_Reason.Text = empCourseItem["DevS_BIC_Reason"].ToString();


                        //if (empCourseItem["DevS_BIC_priority"] != null && !string.IsNullOrEmpty(empCourseItem["DevS_BIC_priority"].ToString()))
                        //    drpLst_DevS_BIC_priority.SelectedValue = empCourseItem["DevS_BIC_priority"].ToString();
                    }
                }
            }


        }

        private void BindDrpLst_TechCourseNames(string loginedDeptName)
        {
            //DataTable dt_Courses = CommonMethods.GetTechnicalCoursesByDeptName(loginedDeptName);

            //if (dt_Courses != null && dt_Courses.Rows.Count > 0)
            //{
            //    DrpLst_Tech_OtherC_Names.DataSource = dt_Courses;
            //    DrpLst_Tech_OtherC_Names.DataTextField = "CourseName";
            //    DrpLst_Tech_OtherC_Names.DataValueField = "CourseName";
            //    DrpLst_Tech_OtherC_Names.DataBind();

            //}
            //DrpLst_Tech_OtherC_Names.Items.Insert(0, new ListItem("-- Select --", "-1"));
        }

        private void BindDrpLst_GeneralCourses(string CourseCategory, string CourseType, DropDownList drpLst)
        {
            DataTable dt_Courses = CommonMethods.GetGeneralCourses(CourseCategory, CourseType);
            if (dt_Courses != null && dt_Courses.Rows.Count > 0)
            {
                drpLst.DataSource = dt_Courses;
                drpLst.DataTextField = "CourseName";
                drpLst.DataValueField = "CourseName";
                drpLst.DataBind();
            }

            drpLst.Items.Insert(0, new ListItem("-- Select --", "-1"));
        }

        private void BindDrpLstEmpNames(DataTable dt_Emps)
        {
            if (dt_Emps != null && dt_Emps.Rows.Count > 0)
            {
                drpLstEmpNames.DataSource = dt_Emps;
                drpLstEmpNames.DataTextField = "FULL_NAME";
                drpLstEmpNames.DataValueField = "EMPLOYEE_NUMBER";
                drpLstEmpNames.DataBind();
            }
            drpLstEmpNames.Items.Insert(0, new ListItem("-- Select --", "-1"));
        }


        protected void drpLstEmpNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (drpLstEmpNames.SelectedValue != "-1")
            {
                DisplayEmployeeData(drpLstEmpNames.SelectedValue);

                BindGV_BS_OtherCourses(txtEmpCode.Text.Trim());
                BindGV_Dev_OtherCourses(txtEmpCode.Text.Trim());
                BindGV_Tech_OtherCourses(txtEmpDept.Text.Trim(), txtEmpCode.Text.Trim());

                // to Get Last Self Appraisal Data for login employee

                //int EmpApprId = CommonMethods.CheckIfEmpAppraisalIsExist(drpLstEmpNames.SelectedValue); //CommonMethods.GetLoginEmpCode()
                //if (EmpApprId > 0)
                //{
                //    GetUserMgrApprData(EmpApprId.ToString());
                //}
            }
            else
            {
                txtEmpCode.Text = string.Empty;
                txtEmpPosition.Text = string.Empty;
                txtEmpDivision.Text = string.Empty;
                txtEmpDept.Text = string.Empty;
                txtEmpSection.Text = string.Empty;
            }
        }


        private void DisplayEmployeeData(string empCode)
        {
            var emp = (from emps in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                       where emps.EMPLOYEE_NUMBER == empCode
                       select emps).ToList();

            DataTable dt_emp = emp.ToDataTable();
            if (dt_emp != null && dt_emp.Rows.Count > 0)
            {
                txtEmpCode.Text = UpdateEmpNumber(dt_emp.Rows[0]["EMPLOYEE_NUMBER"].ToString());
                txtEmpPosition.Text = dt_emp.Rows[0]["POSITION"].ToString();
                txtEmpDivision.Text = dt_emp.Rows[0]["Division"].ToString();
                txtEmpDept.Text = dt_emp.Rows[0]["Department"].ToString();
                txtEmpSection.Text = dt_emp.Rows[0]["Section"].ToString();
            }
        }

        public string UpdateEmpNumber(string empNo)
        {
            if (empNo.Length == 1)
                empNo = "0000" + empNo;

            else if (empNo.Length == 2)
                empNo = "000" + empNo;

            else if (empNo.Length == 3)
                empNo = "00" + empNo;

            else if (empNo.Length == 4)
                empNo = "0" + empNo;

            return empNo;
        }

        public void SaveOther_TechCourse()
        {
            SPSecurity.RunWithElevatedPrivileges(
                  delegate()
                  {
                      using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                      using (SPWeb web = site.OpenWeb())
                      {
                          web.AllowUnsafeUpdates = true;
                          if (string.IsNullOrEmpty(txt_Tech_OtherC_Name.Text))
                          {
                              return;
                          }

                          if (!CheckCourseNameIsExist("TechnicalCoursesLst", txt_Tech_OtherC_Name.Text))
                          {
                              SPList TechnicalCoursesLst = web.Lists.TryGetList("TechnicalCoursesLst");
                              if (TechnicalCoursesLst != null)
                              {
                                  SPListItem item = TechnicalCoursesLst.Items.Add();

                                  item["DeptName"] = CommonMethods.GetLoginEmpDept().Trim();

                                  if (!string.IsNullOrEmpty(txt_Tech_OtherC_Name.Text))
                                      item["CourseName"] = txt_Tech_OtherC_Name.Text.Trim();

                                  item.Update();
                              }
                          }
                          if (!string.IsNullOrEmpty(txtEmpCode.Text.Trim()))
                          {
                              SPList EmpTechnicalCoursesLst = web.Lists.TryGetList("EmpTechnicalCoursesLst");
                              if (EmpTechnicalCoursesLst != null)
                              {
                                  SPListItem item = EmpTechnicalCoursesLst.Items.Add();

                                  item["EmpCode"] = txtEmpCode.Text.Trim();

                                  item["DeptName"] = txtEmpDept.Text.Trim();

                                  if (!string.IsNullOrEmpty(txt_Tech_OtherC_Name.Text))
                                      item["CourseName"] = txt_Tech_OtherC_Name.Text.Trim();

                                  if (drp_Tech_OtherC_ExpectedQuarter.SelectedValue != "-1")
                                      item["ExpectedQuarter"] = drp_Tech_OtherC_ExpectedQuarter.SelectedValue;

                                  if (!string.IsNullOrEmpty(txt_Tech_OtherC_TrainingProvider.Text))
                                      item["TrainingProvider"] = txt_Tech_OtherC_TrainingProvider.Text.Trim();

                                  if (!string.IsNullOrEmpty(txt_Tech_OtherC_Reason.Text))
                                      item["Reason"] = txt_Tech_OtherC_Reason.Text.Trim();

                                  item["TNAYear"] = CommonMethods.GetLookups("TNAYear");

                                  item.Update();
                              }
                          }
                          web.AllowUnsafeUpdates = false;
                      }

                  });
        }

        public void SaveOtherCourse_BasicSkills()
        {
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        web.AllowUnsafeUpdates = true;
                        if (string.IsNullOrEmpty(txt_BS_OtherC_Name.Text))
                        {
                            return;
                        }
                        if (!CheckCourseNameIsExist("BasicSkillsOtherCourses", txt_BS_OtherC_Name.Text))
                        {
                            SPList BasicSkillsOtherCoursesLst = web.Lists.TryGetList("BasicSkillsOtherCourses");
                            if (BasicSkillsOtherCoursesLst != null)
                            {
                                SPListItem BS_item = BasicSkillsOtherCoursesLst.Items.Add();
                                if (!string.IsNullOrEmpty(txt_BS_OtherC_Name.Text))
                                    BS_item["CourseName"] = txt_BS_OtherC_Name.Text.Trim();

                                BS_item.Update();
                            }
                        }


                        if (!string.IsNullOrEmpty(txtEmpCode.Text.Trim()))
                        {
                            SPList EmpBasicSkillsOtherCourses = web.Lists.TryGetList("EmpBasicSkillsOtherCourses");
                            if (EmpBasicSkillsOtherCourses != null)
                            {
                                SPListItem item = EmpBasicSkillsOtherCourses.Items.Add();

                                item["EmpCode"] = txtEmpCode.Text.Trim();

                                /*mary*/
                                item["Department"] = txtEmpDept.Text.Trim();

                                if (!string.IsNullOrEmpty(txt_BS_OtherC_Name.Text))
                                    item["CourseName"] = txt_BS_OtherC_Name.Text.Trim();

                                if (drp_BS_OtherC_ExpectedQuarter.SelectedValue != "-1")
                                    item["ExpectedQuarter"] = drp_BS_OtherC_ExpectedQuarter.SelectedValue;

                                if (!string.IsNullOrEmpty(txt_BS_OtherC_Reason.Text))
                                    item["Reason"] = txt_BS_OtherC_Reason.Text.Trim();

                                item["TNAYear"] = CommonMethods.GetLookups("TNAYear");

                                item.Update();
                            }
                        }

                        web.AllowUnsafeUpdates = false;
                    }
                });
        }
        public void SaveOtherCourse_Dev()
        {
            SPSecurity.RunWithElevatedPrivileges(
           delegate()
           {
               using (SPSite site = new SPSite(SPContext.Current.Web.Url))
               using (SPWeb web = site.OpenWeb())
               {
                   web.AllowUnsafeUpdates = true;
                   if (string.IsNullOrEmpty(txt_DevS_OtherC_Name.Text))
                   {
                       return;
                   }

                   if (!CheckCourseNameIsExist("DevSkillsOtherCourses", txt_DevS_OtherC_Name.Text))
                   {
                       SPList DevSkillsOtherCoursesLst = web.Lists.TryGetList("DevSkillsOtherCourses");
                       if (DevSkillsOtherCoursesLst != null)
                       {
                           SPListItem BS_item = DevSkillsOtherCoursesLst.Items.Add();
                           if (!string.IsNullOrEmpty(txt_BS_OtherC_Name.Text))
                               BS_item["CourseName"] = txt_BS_OtherC_Name.Text.Trim();

                           BS_item.Update();
                       }
                   }

                   if (!string.IsNullOrEmpty(txtEmpCode.Text.Trim()))
                   {
                       SPList EmpDevSkillsOtherCourses = web.Lists.TryGetList("EmpDevSkillsOtherCourses");
                       if (EmpDevSkillsOtherCourses != null)
                       {
                           SPListItem item = EmpDevSkillsOtherCourses.Items.Add();

                           item["EmpCode"] = txtEmpCode.Text.Trim();
                           /*mary*/
                           item["Department"] = txtEmpDept.Text.Trim();

                           if (!string.IsNullOrEmpty(txt_DevS_OtherC_Name.Text))
                               item["CourseName"] = txt_DevS_OtherC_Name.Text.Trim();

                           if (drp_DevS_OtherC_ExpectedQuarter.SelectedValue != "-1")
                               item["ExpectedQuarter"] = drp_DevS_OtherC_ExpectedQuarter.SelectedValue;

                           if (!string.IsNullOrEmpty(txt_DevS_OtherC_Reason.Text))
                               item["Reason"] = txt_DevS_OtherC_Reason.Text.Trim();

                           item["TNAYear"] = CommonMethods.GetLookups("TNAYear");

                           item.Update();
                       }
                   }
                   web.AllowUnsafeUpdates = false;
               }

           });
        }


        protected void btn_AddOther_TechCourse_OnClick(object sender, ImageClickEventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {
                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            web.AllowUnsafeUpdates = true;
                            if (string.IsNullOrEmpty(txt_Tech_OtherC_Name.Text))
                            {
                                return;
                            }

                            if (!CheckCourseNameIsExist("TechnicalCoursesLst", txt_Tech_OtherC_Name.Text))
                            {
                                SPList TechnicalCoursesLst = web.Lists.TryGetList("TechnicalCoursesLst");
                                if (TechnicalCoursesLst != null)
                                {
                                    SPListItem item = TechnicalCoursesLst.Items.Add();

                                    item["DeptName"] = CommonMethods.GetLoginEmpDept().Trim();

                                    if (!string.IsNullOrEmpty(txt_Tech_OtherC_Name.Text))
                                        item["CourseName"] = txt_Tech_OtherC_Name.Text.Trim();

                                    item.Update();
                                }
                            }
                            if (!string.IsNullOrEmpty(txtEmpCode.Text.Trim()))
                            {
                                SPList EmpTechnicalCoursesLst = web.Lists.TryGetList("EmpTechnicalCoursesLst");
                                if (EmpTechnicalCoursesLst != null)
                                {
                                    SPListItem item = EmpTechnicalCoursesLst.Items.Add();

                                    item["EmpCode"] = txtEmpCode.Text.Trim();

                                    item["DeptName"] = txtEmpDept.Text.Trim();

                                    if (!string.IsNullOrEmpty(txt_Tech_OtherC_Name.Text))
                                        item["CourseName"] = txt_Tech_OtherC_Name.Text.Trim();

                                    if (drp_Tech_OtherC_ExpectedQuarter.SelectedValue != "-1")
                                        item["ExpectedQuarter"] = drp_Tech_OtherC_ExpectedQuarter.SelectedValue;

                                    if (!string.IsNullOrEmpty(txt_Tech_OtherC_TrainingProvider.Text))
                                        item["TrainingProvider"] = txt_Tech_OtherC_TrainingProvider.Text.Trim();

                                    if (!string.IsNullOrEmpty(txt_Tech_OtherC_Reason.Text))
                                        item["Reason"] = txt_Tech_OtherC_Reason.Text.Trim();

                                    item["TNAYear"] = CommonMethods.GetLookups("TNAYear");

                                    if (drpLst_Tech_OtherC_priority.SelectedValue != "-1")
                                        item["Priority"] = drpLst_Tech_OtherC_priority.SelectedValue.Trim();

                                    item.Update();
                                }
                            }
                            web.AllowUnsafeUpdates = false;
                        }

                        BindGV_Tech_OtherCourses(txtEmpDept.Text.Trim(), txtEmpCode.Text.Trim());

                        //******mary******/

                        // clear other dev cource data from control o prever duplictae saving
                        txt_Tech_OtherC_Name.Text = string.Empty;
                        drp_Tech_OtherC_ExpectedQuarter.SelectedValue = "-1";
                        txt_Tech_OtherC_TrainingProvider.Text = string.Empty;
                        txt_Tech_OtherC_Reason.Text = string.Empty;
                        drpLst_Tech_OtherC_priority.SelectedValue = "-1";

                        //************/
                    });
            }
            catch (Exception)
            {

            }
        }

        //        private void BindGV_Tech_OtherCourses(string deptName, string empCode)
        //        {
        //            //SPSecurity.RunWithElevatedPrivileges(
        //            //    delegate()
        //            //    {

        //            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
        //            using (SPWeb web = site.OpenWeb())
        //            {

        //                SPList EmpTechnicalCoursesLst = web.Lists.TryGetList("EmpTechnicalCoursesLst");
        //                if (EmpTechnicalCoursesLst != null)
        //                {
        //                    string TNAYear = CommonMethods.GetLookups("TNAYear").ToString();
        //                    SPQuery qry = new SPQuery();

        //                    qry.Query = String.Format(
        //                        @" 
        //                       <Where>
        //                              <And>
        //                                 <Eq>
        //                                    <FieldRef Name='DeptName' />
        //                                    <Value Type='Text'>{0}</Value>
        //                                 </Eq>
        //                                 <And>
        //                                    <Eq>
        //                                       <FieldRef Name='EmpCode' />
        //                                       <Value Type='Text'>{1}</Value>
        //                                    </Eq>
        //                                    <Eq>
        //                                       <FieldRef Name='TNAYear' />
        //                                       <Value Type='Text'>{2}</Value>
        //                                    </Eq>
        //                                 </And>
        //                              </And>
        //                           </Where>", deptName, empCode, TNAYear);


        //                    SPListItemCollection collItems = EmpTechnicalCoursesLst.GetItems(qry);
        //                    if (collItems != null && collItems.Count > 0)
        //                    {
        //                        GV_Tech_OtherCourses.DataSource = collItems.GetDataTable();
        //                        GV_Tech_OtherCourses.DataBind();
        //                    }
        //                    else
        //                    {
        //                        GV_Tech_OtherCourses.DataSource = null;
        //                        GV_Tech_OtherCourses.DataBind();
        //                    }
        //                }
        //            }
        //            // });

        //        }

        private void BindGV_Tech_OtherCourses(string deptName, string empCode)
        {
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList EmpTechnicalCoursesLst = web.Lists.TryGetList("EmpTechnicalCoursesLst");
                        if (EmpTechnicalCoursesLst != null)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = String.Format(
                                @" 
                                  <Where>
                                      <And>
                                         <Eq>
                                            <FieldRef Name='EmpCode' />
                                            <Value Type='Text'>{0}</Value>
                                         </Eq>
                                         <And>
                                            <Eq>
                                               <FieldRef Name='DeptName' />
                                               <Value Type='Text'>{1}</Value>
                                            </Eq>
                                            <Eq>
                                               <FieldRef Name='TNAYear' />
                                               <Value Type='Text'>{2}</Value>
                                            </Eq>
                                         </And>
                                      </And>
                                   </Where>
                                       ", empCode, deptName, CommonMethods.GetLookups("TNAYear"));

                            SPListItemCollection collItems = EmpTechnicalCoursesLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                GV_Tech_OtherCourses.DataSource = collItems.GetDataTable();
                                GV_Tech_OtherCourses.DataBind();
                            }
                            else
                            {
                                GV_Tech_OtherCourses.DataSource = null;
                                GV_Tech_OtherCourses.DataBind();
                            }
                        }
                    }
                });
        }


        public string GetDeleteIconUrl()
        {
            return CommonMethods.GetLookups("DeleteImgUrl");
        }

        public string GetAddIconUrl()
        {
            return CommonMethods.GetLookups("AddImgUrl");
        }


        protected void btn_AddOtherCourse_BasicSkills_OnClick(object sender, ImageClickEventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {
                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            web.AllowUnsafeUpdates = true;
                            if (string.IsNullOrEmpty(txt_BS_OtherC_Name.Text))
                            {
                                return;
                            }
                            if (!CheckCourseNameIsExist("BasicSkillsOtherCourses", txt_BS_OtherC_Name.Text))
                            {
                                SPList BasicSkillsOtherCoursesLst = web.Lists.TryGetList("BasicSkillsOtherCourses");
                                if (BasicSkillsOtherCoursesLst != null)
                                {
                                    SPListItem BS_item = BasicSkillsOtherCoursesLst.Items.Add();
                                    if (!string.IsNullOrEmpty(txt_BS_OtherC_Name.Text))
                                        BS_item["CourseName"] = txt_BS_OtherC_Name.Text.Trim();

                                    BS_item.Update();
                                }
                            }

                            if (!string.IsNullOrEmpty(txtEmpCode.Text.Trim()))
                            {
                                SPList EmpBasicSkillsOtherCourses = web.Lists.TryGetList("EmpBasicSkillsOtherCourses");
                                if (EmpBasicSkillsOtherCourses != null)
                                {
                                    SPListItem item = EmpBasicSkillsOtherCourses.Items.Add();

                                    item["EmpCode"] = txtEmpCode.Text.Trim();

                                    /*mary*/
                                    item["Department"] = txtEmpDept.Text.Trim();

                                    if (!string.IsNullOrEmpty(txt_BS_OtherC_Name.Text))
                                        item["CourseName"] = txt_BS_OtherC_Name.Text.Trim();

                                    if (drp_BS_OtherC_ExpectedQuarter.SelectedValue != "-1")
                                        item["ExpectedQuarter"] = drp_BS_OtherC_ExpectedQuarter.SelectedValue;

                                    if (!string.IsNullOrEmpty(txt_BS_OtherC_Reason.Text))
                                        item["Reason"] = txt_BS_OtherC_Reason.Text.Trim();

                                    item["TNAYear"] = CommonMethods.GetLookups("TNAYear");

                                    if (drpLst_BS_OtherC_priority.SelectedValue != "-1")
                                        item["Priority"] = drpLst_BS_OtherC_priority.SelectedValue.Trim();

                                    item.Update();
                                }
                            }

                            web.AllowUnsafeUpdates = false;
                        }

                        BindGV_BS_OtherCourses(txtEmpCode.Text.Trim());

                        //******mary******/
                        // clear other basic cource data from control o prever duplictae saving

                        txt_BS_OtherC_Name.Text = string.Empty;
                        drp_BS_OtherC_ExpectedQuarter.SelectedValue = "-1";
                        txt_BS_OtherC_Reason.Text = string.Empty;
                        drpLst_BS_OtherC_priority.SelectedValue = "-1";

                        //************/

                   });
            }
            catch (Exception)
            {
            }
        }


        private void BindGV_BS_OtherCourses(string empCode)
        {
            SPSecurity.RunWithElevatedPrivileges(
               delegate()
               {
                   using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                   using (SPWeb web = site.OpenWeb())
                   {
                       SPList EmpBasicSkillsOtherCoursesLst = web.Lists.TryGetList("EmpBasicSkillsOtherCourses");
                       if (EmpBasicSkillsOtherCoursesLst != null)
                       {
                           SPQuery qry = new SPQuery();

                           qry.Query = String.Format(
                               @" 
                                  <Where>
                                      <And>
                                         <Eq>
                                            <FieldRef Name='EmpCode' />
                                            <Value Type='Text'>{0}</Value>
                                         </Eq>
                                         <Eq>
                                            <FieldRef Name='TNAYear' />
                                            <Value Type='Text'>{1}</Value>
                                         </Eq>
                                      </And>
                                   </Where>", empCode, CommonMethods.GetLookups("TNAYear"));

                           SPListItemCollection collItems = EmpBasicSkillsOtherCoursesLst.GetItems(qry);
                           if (collItems.Count > 0)
                           {
                               GV_BS_OtherCourses.DataSource = collItems.GetDataTable();
                               GV_BS_OtherCourses.DataBind();
                           }
                           else
                           {
                               GV_BS_OtherCourses.DataSource = null;
                               GV_BS_OtherCourses.DataBind();
                           }
                       }
                   }
               });
        }


        public bool CheckCourseNameIsExist(string LstName, string courseName)
        {
            bool Exist = false;

            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList HrLookupLst = web.Lists.TryGetList(LstName);
                if (HrLookupLst != null)
                {
                    SPQuery qry = new SPQuery();
                    qry.Query = String.Format(
                        @" 
                                <Where>
                                  <Eq>
                                     <FieldRef Name='CourseName' />
                                     <Value Type='Text'>{0}</Value>
                                  </Eq>
                               </Where>
                                       ", courseName);

                    SPListItemCollection collItems = HrLookupLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        Exist = true;
                    }
                }
            }
            return Exist;
        }


        protected void btn_AddOtherCourse_Dev_OnClick(object sender, ImageClickEventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
                  delegate()
                  {
                      using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                      using (SPWeb web = site.OpenWeb())
                      {
                          web.AllowUnsafeUpdates = true;
                          if (string.IsNullOrEmpty(txt_DevS_OtherC_Name.Text))
                          {
                              return;
                          }

                          if (!CheckCourseNameIsExist("DevSkillsOtherCourses", txt_DevS_OtherC_Name.Text))
                          {
                              SPList DevSkillsOtherCoursesLst = web.Lists.TryGetList("DevSkillsOtherCourses");
                              if (DevSkillsOtherCoursesLst != null)
                              {
                                  SPListItem BS_item = DevSkillsOtherCoursesLst.Items.Add();
                                  if (!string.IsNullOrEmpty(txt_BS_OtherC_Name.Text))
                                      BS_item["CourseName"] = txt_BS_OtherC_Name.Text.Trim();

                                  BS_item.Update();
                              }
                          }

                          if (!string.IsNullOrEmpty(txtEmpCode.Text.Trim()))
                          {
                              SPList EmpDevSkillsOtherCourses = web.Lists.TryGetList("EmpDevSkillsOtherCourses");
                              if (EmpDevSkillsOtherCourses != null)
                              {
                                  SPListItem item = EmpDevSkillsOtherCourses.Items.Add();

                                  item["EmpCode"] = txtEmpCode.Text.Trim();
                                  /*mary*/
                                  item["Department"] = txtEmpDept.Text.Trim();

                                  if (!string.IsNullOrEmpty(txt_DevS_OtherC_Name.Text))
                                      item["CourseName"] = txt_DevS_OtherC_Name.Text.Trim();

                                  if (drp_DevS_OtherC_ExpectedQuarter.SelectedValue != "-1")
                                      item["ExpectedQuarter"] = drp_DevS_OtherC_ExpectedQuarter.SelectedValue;

                                  if (!string.IsNullOrEmpty(txt_DevS_OtherC_Reason.Text))
                                      item["Reason"] = txt_DevS_OtherC_Reason.Text.Trim();

                                  item["TNAYear"] = CommonMethods.GetLookups("TNAYear");

                                  if (drpLst_DevS_OtherC_priority.SelectedValue != "-1")
                                      item["Priority"] = drpLst_DevS_OtherC_priority.SelectedValue.Trim();

                                  item.Update();
                              }
                          }
                          web.AllowUnsafeUpdates = false;
                      }

                      BindGV_Dev_OtherCourses(txtEmpCode.Text.Trim());

                      //******mary******/
                      // clear other dev cource data from control o prever duplictae saving
                      txt_DevS_OtherC_Name.Text = string.Empty;
                      drp_DevS_OtherC_ExpectedQuarter.SelectedValue = "-1";
                      txt_DevS_OtherC_Reason.Text = string.Empty;
                      drpLst_DevS_OtherC_priority.SelectedValue = "-1";

                      //************/
                  });
            }
            catch (Exception)
            {
            }
        }


        private void BindGV_Dev_OtherCourses(string empCode)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
          delegate()
          {
              using (SPSite site = new SPSite(SPContext.Current.Web.Url))
              using (SPWeb web = site.OpenWeb())
              {
                  SPList EmpDevSkillsOtherCoursesLst = web.Lists.TryGetList("EmpDevSkillsOtherCourses");
                  if (EmpDevSkillsOtherCoursesLst != null)
                  {
                      SPQuery qry = new SPQuery();

                      qry.Query = String.Format(
                          @" 
                                     <Where>
                                          <And>
                                             <Eq>
                                                <FieldRef Name='EmpCode' />
                                                <Value Type='Text'>{0}</Value>
                                             </Eq>
                                             <Eq>
                                                <FieldRef Name='TNAYear' />
                                                <Value Type='Text'>{1}</Value>
                                             </Eq>
                                          </And>
                                       </Where>
                                       ", empCode, CommonMethods.GetLookups("TNAYear"));

                      SPListItemCollection collItems = EmpDevSkillsOtherCoursesLst.GetItems(qry);
                      if (collItems.Count > 0)
                      {
                          GV_Dev_OtherCourses.DataSource = collItems.GetDataTable();
                          GV_Dev_OtherCourses.DataBind();
                      }
                      else
                      {
                          GV_Dev_OtherCourses.DataSource = null;
                          GV_Dev_OtherCourses.DataBind();
                      }
                  }
              }
          });
            }
            catch (Exception)
            {
            }
        }

        protected void btnSaveTNAReq_OnClick(object sender, EventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {

                if (!string.IsNullOrEmpty(txtEmpCode.Text.Trim()))
                {
                    if (!CheckEmpHaveInsertedCourses())
                    {
                        return;
                    }
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        Guid SiteId = site.ID;
                        Guid WebId = web.ID;
                        SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                        SPListItem empCourseItem = null;

                        if (EmpCoursesLst != null)
                        {
                            web.AllowUnsafeUpdates = true;
                            if (!CheckIfCurrentYearRequestIsExist(txtEmpCode.Text.Trim(), CommonMethods.GetLookups("TNAYear"))) // Add New Courses
                            {
                                empCourseItem = EmpCoursesLst.Items.Add();
                            }
                            else // Edit Courses
                            {
                                if (HttpContext.Current.Request.QueryString["pageTo"] != null &&
                                    HttpContext.Current.Request.QueryString["TNAId"] != null)
                                {
                                    string Id = HttpContext.Current.Request.QueryString["TNAId"];
                                    empCourseItem = EmpCoursesLst.GetItemById(int.Parse(Id));
                                }
                                else
                                {
                                    //mary 07-02-2021
                                    //empCourseItem = GetCurrentYearCourseItemByEmpCode(txtEmpCode.Text.Trim(), CommonMethods.GetLookups("TNAYear"));
                                    empCourseItem = EmpCoursesLst.GetItemById(GetCurrentYearCourseItemIDByEmpCode(txtEmpCode.Text.Trim(), CommonMethods.GetLookups("TNAYear")));
                                }
                            }

                            if (drpLstEmpNames.SelectedValue != "-1")
                                empCourseItem["EmpName"] = drpLstEmpNames.SelectedItem.Text.Trim();

                            empCourseItem["EmpCode"] = txtEmpCode.Text.Trim();
                            empCourseItem["Position"] = txtEmpPosition.Text.Trim();
                            empCourseItem["Division"] = txtEmpDivision.Text.Trim();
                            empCourseItem["Department"] = txtEmpDept.Text.Trim();
                            empCourseItem["Section"] = txtEmpSection.Text.Trim();


                            // *********************************************************************** //


                            if (drpLst_BS_EngC_Name.SelectedValue != "-1")
                                empCourseItem["BS_EngC_Name"] = drpLst_BS_EngC_Name.SelectedValue.Trim();

                            if (drp_BS_EngC_ExpectedQuarter.SelectedValue != "-1")
                                empCourseItem["BS_EngC_ExpectedQuarter"] = drp_BS_EngC_ExpectedQuarter.SelectedValue;
                            else
                            {
                                empCourseItem["BS_EngC_ExpectedQuarter"] = null;
                            }

                            empCourseItem["BS_EngC_Reason"] = txt_BS_EngC_Reason.Text.Trim();

                            if (drpLst_BS_EngC_priority.SelectedValue != "-1")
                                empCourseItem["BS_EngC_priority"] = drpLst_BS_EngC_priority.SelectedValue.Trim();

                            //**

                            if (drpLst_BS_CompSkillsC_Name.SelectedValue != "-1")
                                empCourseItem["BS_CompSkillsC_Name"] = drpLst_BS_CompSkillsC_Name.Text.Trim();

                            if (drp_BS_CompSkillsC_ExpectedQuarter.SelectedValue != "-1")
                                empCourseItem["BS_CompSkillsC_ExpectedQuarter"] = drp_BS_CompSkillsC_ExpectedQuarter.SelectedValue;
                            else
                            {
                                empCourseItem["BS_CompSkillsC_ExpectedQuarter"] = null;
                            }

                            empCourseItem["BS_CompSkillsC_Reason"] = txt_BS_CompSkillsC_Reason.Text.Trim();

                            if (drpLst_BS_CompSkillsC_priority.SelectedValue != "-1")
                                empCourseItem["BS_CompSkillsC_priority"] = drpLst_BS_CompSkillsC_priority.SelectedValue.Trim();

                            //** 

                            if (drpLst_BS_BEPC_Name.SelectedValue != "-1")
                                empCourseItem["BS_BEPC_Name"] = drpLst_BS_BEPC_Name.SelectedValue.Trim();

                            if (drp_BS_BEPC_ExpectedQuarter.SelectedValue != "-1")
                                empCourseItem["BS_BEPC_ExpectedQuarter"] = drp_BS_BEPC_ExpectedQuarter.SelectedValue;
                            else
                            {
                                empCourseItem["BS_BEPC_ExpectedQuarter"] = null;
                            }

                            empCourseItem["BS_BEPC_Reason"] = txt_BS_BEPC_Reason.Text.Trim();

                            if (drpLst_BS_BEPC_priority.SelectedValue != "-1")
                                empCourseItem["BS_BEPC_priority"] = drpLst_BS_BEPC_priority.SelectedValue.Trim();

                            //**

                            if (DrpLst_DevS_CommSC_Name.SelectedValue != "-1")
                                empCourseItem["DevS_CommSC_Name"] = DrpLst_DevS_CommSC_Name.SelectedValue.Trim();

                            if (drp_DevS_CommSC_ExpectedQuarter.SelectedValue != "-1")
                                empCourseItem["DevS_CommSC_ExpectedQuarter"] = drp_DevS_CommSC_ExpectedQuarter.SelectedValue;
                            else
                            {
                                empCourseItem["DevS_CommSC_ExpectedDate"] = null;
                            }

                            empCourseItem["DevS_CommSC_Reason"] = txt_DevS_CommSC_Reason.Text.Trim();

                            if (drpLst_DevS_CommSC_priority.SelectedValue != "-1")
                                empCourseItem["BS_CompSkillsC_priority"] = drpLst_DevS_CommSC_priority.SelectedValue.Trim();

                            //**

                            if (DrpLst_DevS_InterPersSC_Name.SelectedValue != "-1")
                                empCourseItem["DevS_InterPersSC_Name"] = DrpLst_DevS_InterPersSC_Name.SelectedValue.Trim();

                            if (drp_DevS_InterPersSC_ExpectedQuarter.SelectedValue != "-1")
                                empCourseItem["DevS_InterPersSC_ExpectedQuarter"] = drp_DevS_InterPersSC_ExpectedQuarter.SelectedValue;
                            else
                            {
                                empCourseItem["DevS_InterPersSC_ExpectedQuarter"] = null;
                            }


                            empCourseItem["DevS_InterPersSC_Reason"] = txt_DevS_InterPersSC_Reason.Text.Trim();



                            if (drpLst_DevS_InterPersSC_priority.SelectedValue != "-1")
                                empCourseItem["DevS_InterPersSC_priority"] = drpLst_DevS_InterPersSC_priority.SelectedValue.Trim();

                            //**

                            //if (DrpLst_DevS_BIC_Name.SelectedValue != "-1")
                            //    empCourseItem["DevS_BIC_Name"] = DrpLst_DevS_BIC_Name.SelectedValue.Trim();
                            //if (!cal_DevS_BIC_ExpectedDate.IsDateEmpty)
                            //    empCourseItem["DevS_BIC_ExpectedDate"] = cal_DevS_BIC_ExpectedDate.SelectedDate;
                            //else
                            //{
                            //    empCourseItem["DevS_BIC_ExpectedDate"] = null;
                            //}

                            //empCourseItem["DevS_BIC_Reason"] = txt_DevS_BIC_Reason.Text.Trim();

                            //if (drpLst_DevS_BIC_priority.SelectedValue != "-1")
                            //    empCourseItem["DevS_BIC_priority"] = drpLst_DevS_BIC_priority.SelectedValue.Trim();
                            //**

                            if (!CommonMethods.CheckIfLoginEmpIsDivisionMgr(CommonMethods.GetLoginEmpCode()))
                            {
                                empCourseItem["RequestStatus"] = "Requested By Supervisor";
                            }
                            else
                            {
                                // If login user is DH
                                empCourseItem["RequestStatus"] = "Sent to HR";
                            }

                            empCourseItem["SupervisorCode"] = CommonMethods.GetLoginEmpCode();
                            empCourseItem["MgrCode"] =
                                CommonMethods.GetMgrCodeByEmpCode(CommonMethods.GetLoginEmpCode());

                            empCourseItem["DHUser"] = "Supervisor";
                            empCourseItem["TNAYear"] = CommonMethods.GetLookups("TNAYear");

                            empCourseItem.Update();

                            web.AllowUnsafeUpdates = false;


                            // Other Basic Skills //

                            SaveOtherCourse_BasicSkills();

                            //**********************


                            // Other Dev. Skills //

                            SaveOtherCourse_Dev();

                            //***************************


                            // Tech. Skills //

                            SaveOther_TechCourse();

                            //***************************


                            int TNAId = empCourseItem.ID;


                            /////////////// Send Hr TNA Mail To Manager ///////////////

                            if (!CommonMethods.CheckIfLoginEmpIsDivisionMgr(CommonMethods.GetLoginEmpCode()))
                            {
                                string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
                                string employeeName = drpLstEmpNames.SelectedItem.Text.Trim();
                                string pageName =
                                    System.IO.Path.GetFileName(HttpContext.Current.Request.PhysicalPath);


                                // "administrator@nslab.lab";

                                string mgrMail = CommonMethods.GetMgrMailByEmpCode(CommonMethods.GetLoginEmpCode());
                                MailTemplates = CommonMethods.LoadMailTemplates();

                                if (MailTemplates["SupervisorToMgrTemplate"] != null)
                                {
                                    if (CommonMethods.SendMail(mgrMail.Trim(),
                                        "TNA for Employee :" + employeeName,
                                        string.Format(MailTemplates["SupervisorToMgrTemplate"].ToString(),
                                            txtEmpCode.Text.Trim(), employeeName, txtEmpPosition.Text.Trim(),
                                            currentAppUrl + pageName, TNAId), SiteId, WebId))
                                    {
                                        //ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                                        //    "Success Message",
                                        //    "alert('Your Request Saved Sucessfully')", true);

                                        ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                            "alert('Your Request Saved Sucessfully')", true);

                                    }
                                    else
                                    {
                                        viewErrorMsg();
                                    }
                                }
                            }
                            else
                            {
                                //ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                                //          "Success Message",
                                //          "alert('Your Request Saved Sucessfully')", true);

                                ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                          "alert('Your Request Saved Sucessfully')", true);

                                // If Login is DVMgr
                                /*
                                string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
                                string employeeName = drpLstEmpNames.SelectedItem.Text.Trim();
                                string pageName =
                                    System.IO.Path.GetFileName(HttpContext.Current.Request.PhysicalPath);

                                string hrEmail = CommonMethods.GetLookups("hrEmail");

                                MailTemplates = CommonMethods.LoadMailTemplates();

                                if (MailTemplates["ToHrTemplate"] != null)
                                {
                                    if (CommonMethods.SendMail(hrEmail.Trim(),
                                        "TNA for Employee :" + employeeName,
                                        string.Format(MailTemplates["ToHrTemplate"].ToString(),
                                            txtEmpCode.Text.Trim(), employeeName, txtEmpPosition.Text.Trim(),
                                            currentAppUrl + pageName, TNAId), SiteId, WebId))
                                    {
                                        ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                                            "Success Message",
                                            "alert('Your Request Sent To HR Sucessfully')", true);
                                    }
                                    else
                                    {
                                        viewErrorMsg();
                                    }
                                }
                         */

                            }


                        }

                    }
                }
                 });
            }
            catch (Exception)
            {
            }
        }

        public bool CheckEmpHaveInsertedCourses()
        {
            bool haveCourses = false;
            if (drpLst_BS_EngC_Name.SelectedValue != "-1" ||
                drpLst_BS_CompSkillsC_Name.SelectedValue != "-1" ||
                drpLst_BS_BEPC_Name.SelectedValue != "-1" ||
                DrpLst_DevS_CommSC_Name.SelectedValue != "-1" ||
                //DrpLst_DevS_BIC_Name.SelectedValue != "-1" ||
                DrpLst_DevS_InterPersSC_Name.SelectedValue != "-1" ||
                GV_BS_OtherCourses.Rows.Count > 0 ||
                 GV_Dev_OtherCourses.Rows.Count > 0 ||
                 GV_Tech_OtherCourses.Rows.Count > 0
                )
            {
                haveCourses = true;
            }
            return haveCourses;
        }

        private bool CheckIfCurrentYearRequestIsExist(string empCode, string currentYear)
        {
            bool Exist = false;

            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                if (EmpCoursesLst != null)
                {
                    SPQuery qry = new SPQuery();

                    qry.Query = String.Format(
                        @" 
                                   <Where>
                                      <And>
                                         <Eq>
                                            <FieldRef Name='EmpCode' />
                                            <Value Type='Text'>{0}</Value>
                                         </Eq>
                                         <Eq>
                                            <FieldRef Name='TNAYear' />
                                            <Value Type='Text'>{1}</Value>
                                         </Eq>
                                      </And>
                                   </Where>
                                       ", empCode, currentYear);

                    SPListItemCollection collItems = EmpCoursesLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        Exist = true;
                    }
                }
            }
            return Exist;
        }

        private SPListItem GetCurrentYearCourseItemByEmpCode(string empCode, string year)
        {
            SPListItem item = null;

            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList HrLookupLst = web.Lists.TryGetList("EmpCourses");
                if (HrLookupLst != null)
                {
                    SPQuery qry = new SPQuery();

                    qry.Query = String.Format(
                        @" 
                                   <Where>
                                      <And>
                                         <Eq>
                                            <FieldRef Name='EmpCode' />
                                            <Value Type='Text'>{0}</Value>
                                         </Eq>
                                         <Eq>
                                            <FieldRef Name='TNAYear' />
                                            <Value Type='Text'>{1}</Value>
                                         </Eq>
                                      </And>
                                   </Where>
                                       ", empCode, year);

                    SPListItemCollection collItems = HrLookupLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        item = collItems[0];
                    }
                }
            }
            return item;
        }
        private int GetCurrentYearCourseItemIDByEmpCode(string empCode, string year)
        {
            int itemID = 0;

            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList HrLookupLst = web.Lists.TryGetList("EmpCourses");
                if (HrLookupLst != null)
                {
                    SPQuery qry = new SPQuery();

                    qry.Query = String.Format(
                        @" 
                                   <Where>
                                      <And>
                                         <Eq>
                                            <FieldRef Name='EmpCode' />
                                            <Value Type='Text'>{0}</Value>
                                         </Eq>
                                         <Eq>
                                            <FieldRef Name='TNAYear' />
                                            <Value Type='Text'>{1}</Value>
                                         </Eq>
                                      </And>
                                   </Where>
                                       ", empCode, year);

                    SPListItemCollection collItems = HrLookupLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        itemID = int.Parse((collItems.GetDataTable()).Rows[0]["ID"].ToString());
                    }
                }
            }
            return itemID;
        }
        private void viewErrorMsg()
        {
            ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Error Message", "alert('Error in saving')", true);
            //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "Error Message", "alert('Error in saving')", true);
        }

        protected void btnDeleteBSOtherCourse_OnClick(object sender, ImageClickEventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    var DelBtnLink = (ImageButton)sender;
                    var row = (GridViewRow)DelBtnLink.NamingContainer;
                    var courseId = (Label)row.FindControl("lbl_BS_CourseID");

                    if (!string.IsNullOrEmpty(courseId.Text))
                    {
                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            web.AllowUnsafeUpdates = true;
                            SPList EmpBasicSkillsOtherCoursesLst = web.Lists.TryGetList("EmpBasicSkillsOtherCourses");
                            if (EmpBasicSkillsOtherCoursesLst != null)
                            {
                                SPListItem itm = EmpBasicSkillsOtherCoursesLst.GetItemById(int.Parse(courseId.Text));
                                itm.Delete();
                                web.AllowUnsafeUpdates = false;
                                BindGV_BS_OtherCourses(txtEmpCode.Text.Trim());

                                ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                    "alert('Course Deleted Sucessfully')", true);

                                //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "Success Message",
                                //    "alert('Course Deleted Sucessfully')", true);

                            }
                        }
                    }
                });
            }
            catch (Exception)
            {
            }
        }

        protected void btnDeleteDevOtherCourse_OnClick(object sender, ImageClickEventArgs e)
        {
            try
            {


                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    var DelBtnLink = (ImageButton)sender;
                    var row = (GridViewRow)DelBtnLink.NamingContainer;
                    var courseId = (Label)row.FindControl("lbl_Dev_CourseID");

                    if (!string.IsNullOrEmpty(courseId.Text))
                    {
                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            SPList EmpDevSkillsOtherCoursesLst = web.Lists.TryGetList("EmpDevSkillsOtherCourses");
                            if (EmpDevSkillsOtherCoursesLst != null)
                            {
                                SPListItem itm = EmpDevSkillsOtherCoursesLst.GetItemById(int.Parse(courseId.Text));
                                itm.Delete();

                                BindGV_Dev_OtherCourses(txtEmpCode.Text.Trim());

                                //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "Success Message",
                                //    "alert('Course Deleted Sucessfully')", true);

                                ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                    "alert('Course Deleted Sucessfully')", true);
                            }
                        }
                    }
                });
            }
            catch (Exception)
            {
            }
        }

        protected void btnDeleteTechOtherCourse_OnClick(object sender, ImageClickEventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    var DelBtnLink = (ImageButton)sender;
                    var row = (GridViewRow)DelBtnLink.NamingContainer;
                    var courseId = (Label)row.FindControl("lbl_Tech_CourseID");

                    if (!string.IsNullOrEmpty(courseId.Text))
                    {
                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            SPList EmpTechnicalCoursesLst = web.Lists.TryGetList("EmpTechnicalCoursesLst");
                            if (EmpTechnicalCoursesLst != null)
                            {
                                SPListItem itm = EmpTechnicalCoursesLst.GetItemById(int.Parse(courseId.Text));
                                itm.Delete();

                                BindGV_Tech_OtherCourses(txtEmpDept.Text.Trim(), txtEmpCode.Text.Trim());

                                ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                    "alert('Course Deleted Sucessfully')", true);

                                //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "Success Message",
                                //    "alert('Course Deleted Sucessfully')", true);
                            }
                        }
                    }
                });
            }
            catch (Exception)
            {
            }
        }

        protected void btnApproveTNAReqByMgr_OnClick(object sender, EventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {

                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            Guid SiteId = site.ID;
                            Guid WebId = web.ID;
                            SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                            if (EmpCoursesLst != null)
                            {
                                if (HttpContext.Current.Request.QueryString["pageTo"] != null &&
                                    HttpContext.Current.Request.QueryString["TNAId"] != null)
                                {
                                    web.AllowUnsafeUpdates = true;
                                    string TNAId = HttpContext.Current.Request.QueryString["TNAId"];
                                    SPListItem empCourseItem = EmpCoursesLst.GetItemById(int.Parse(TNAId));
                                    string empSupervisorCode = empCourseItem["MgrCode"].ToString();
                                    if (!CommonMethods.CheckIfLoginEmpIsDivisionMgr(CommonMethods.GetLoginEmpCode()))
                                    {
                                        empCourseItem["RequestStatus"] = "Approved By Manger";
                                    }
                                    else
                                    {
                                        empCourseItem["RequestStatus"] = "Sent to HR";
                                    }
                                    empCourseItem["MgrCode"] = CommonMethods.GetLoginEmpCode();
                                    empCourseItem["DHCode"] =
                                        CommonMethods.GetMgrCodeByEmpCode(CommonMethods.GetLoginEmpCode());
                                    empCourseItem["DHUser"] = "Manger";

                                    empCourseItem.Update();

                                    web.AllowUnsafeUpdates = false;

                                    // If Login User is DVMgr // To skip Step of DVMgr approval
                                    if (!CommonMethods.CheckIfLoginEmpIsDivisionMgr(CommonMethods.GetLoginEmpCode()))
                                    {
                                        string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
                                        string employeeName = drpLstEmpNames.SelectedItem.Text.Trim();
                                        string pageName =
                                            System.IO.Path.GetFileName(HttpContext.Current.Request.PhysicalPath);

                                        string DHMail = CommonMethods.GetDHMailByEmpCode(txtEmpCode.Text.Trim());
                                        /* "mgeady@nslab.lab"; */
                                        MailTemplates = CommonMethods.LoadMailTemplates();

                                        if (MailTemplates["MgrToDHTemplate"] != null)
                                        {
                                            if (CommonMethods.SendMail(DHMail.Trim(),
                                                "TNA for Employee :" + employeeName,
                                                string.Format(MailTemplates["MgrToDHTemplate"].ToString(),
                                                    txtEmpCode.Text.Trim(), employeeName, txtEmpPosition.Text.Trim(),
                                                    currentAppUrl + pageName, TNAId), SiteId, WebId))
                                            {
                                                //ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                                                //    "Success Message",
                                                //    "alert('Your Request Saved Sucessfully')", true);

                                                ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                                    "alert('Your Request Saved Sucessfully')", true);
                                            }
                                            else
                                            {
                                                viewErrorMsg();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // If Login is DVMgr

                                        //string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
                                        //string employeeName = drpLstEmpNames.SelectedItem.Text.Trim();
                                        //string pageName =
                                        //    System.IO.Path.GetFileName(HttpContext.Current.Request.PhysicalPath);

                                        //string hrEmail = CommonMethods.GetLookups("hrEmail");

                                        //MailTemplates = CommonMethods.LoadMailTemplates();

                                        //if (MailTemplates["ToHrTemplate"] != null)
                                        //{
                                        //    if (CommonMethods.SendMail(hrEmail.Trim(),
                                        //        "TNA for Employee :" + employeeName,
                                        //        string.Format(MailTemplates["ToHrTemplate"].ToString(),
                                        //            txtEmpCode.Text.Trim(), employeeName, txtEmpPosition.Text.Trim(),
                                        //            currentAppUrl + pageName, TNAId), SiteId, WebId))
                                        //    {
                                        //        ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                                        //            "Success Message",
                                        //            "alert('Your Request Sent To HR Sucessfully')", true);
                                        //    }
                                        //    else
                                        //    {
                                        //        viewErrorMsg();
                                        //    }
                                        //}
                                    }

                                    //*********************************
                                }
                            }
                        }
                    });
            }
            catch (Exception)
            {
            }
        }

        protected void btnRejectTNAReqByMgr_OnClick(object sender, EventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {

                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            Guid SiteId = site.ID;
                            Guid WebId = web.ID;
                            SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                            if (EmpCoursesLst != null)
                            {
                                if (HttpContext.Current.Request.QueryString["pageTo"] != null &&
                                    HttpContext.Current.Request.QueryString["TNAId"] != null)
                                {
                                    string TNAId = HttpContext.Current.Request.QueryString["TNAId"];
                                    SPListItem empCourseItem = EmpCoursesLst.GetItemById(int.Parse(TNAId));

                                    empCourseItem["RequestStatus"] = "Rejected By Manger";
                                    string empSupervisorCode = empCourseItem["MgrCode"].ToString();
                                    empCourseItem.Update();

                                    // Send rejection mail to Supervisor
                                    string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
                                    string employeeName = drpLstEmpNames.SelectedItem.Text.Trim();
                                    string pageName =
                                        System.IO.Path.GetFileName(HttpContext.Current.Request.PhysicalPath);

                                    string supervisorMail = CommonMethods.GetEmpMailByEmpCode(empSupervisorCode);
                                    /*  "mgeady@nslab.lab";*/
                                    MailTemplates = CommonMethods.LoadMailTemplates();

                                    if (MailTemplates["MgrToSupervisorRejectTemplate"] != null)
                                    {
                                        if (CommonMethods.SendMail(supervisorMail.Trim(),
                                            "TNA for Employee :" + employeeName,
                                            string.Format(MailTemplates["MgrToSupervisorRejectTemplate"].ToString(),
                                                txtEmpCode.Text.Trim(), employeeName, txtEmpPosition.Text.Trim()),
                                            SiteId, WebId))
                                        {
                                            //ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                                            //    "Success Message",
                                            //    "alert('Your Request Rejected Sucessfully')", true);

                                            ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                                "alert('Your Request Rejected Sucessfully')", true);
                                        }
                                        else
                                        {
                                            viewErrorMsg();
                                        }
                                    }
                                }
                            }
                        }
                    });

            }
            catch (Exception)
            {
            }
        }

        protected void btnApproveTNAReqByDH_OnClick(object sender, EventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {

                        //Rehab
                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            Guid SiteId = site.ID;
                            Guid WebId = web.ID;
                            SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                            if (EmpCoursesLst != null)
                            {
                                if (HttpContext.Current.Request.QueryString["pageTo"] != null &&
                                    HttpContext.Current.Request.QueryString["TNAId"] != null)
                                {
                                    web.AllowUnsafeUpdates = true;
                                    string TNAId = HttpContext.Current.Request.QueryString["TNAId"];
                                    SPListItem empCourseItem = EmpCoursesLst.GetItemById(int.Parse(TNAId));

                                    empCourseItem["RequestStatus"] = "Approved By Dept. Head";
                                    empCourseItem["DHCode"] = CommonMethods.GetLoginEmpCode();
                                    empCourseItem["DVMgrCode"] =
                                        CommonMethods.GetMgrCodeByEmpCode(CommonMethods.GetLoginEmpCode());
                                    empCourseItem["DHUser"] = "Dept. Head";
                                    empCourseItem.Update();

                                    web.AllowUnsafeUpdates = false;

                                    // If Login User is DVMgr // To skip Step of DVMgr approval
                                    if (!CommonMethods.CheckIfLoginEmpIsDivisionMgr(txtEmpCode.Text.Trim()))
                                    {
                                        string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
                                        string employeeName = drpLstEmpNames.SelectedItem.Text.Trim();
                                        string pageName =
                                            System.IO.Path.GetFileName(HttpContext.Current.Request.PhysicalPath);
                                        string empDHCode = empCourseItem["DHCode"].ToString();
                                        string DVMgrMail = CommonMethods.GetDVMgrMailByEmpCode(txtEmpCode.Text.Trim());
                                        /*"mgeady@nslab.lab";*/
                                        MailTemplates = CommonMethods.LoadMailTemplates();

                                        if (MailTemplates["DHToDVMgrTemplate"] != null)
                                        {
                                            if (CommonMethods.SendMail(DVMgrMail.Trim(),
                                                "TNA for Employee :" + employeeName,
                                                string.Format(MailTemplates["DHToDVMgrTemplate"].ToString(),
                                                    txtEmpCode.Text.Trim(), employeeName, txtEmpPosition.Text.Trim(),
                                                    currentAppUrl + pageName, TNAId), SiteId, WebId))
                                            {
                                                //ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                                                //    "Success Message",
                                                //    "alert('Your Request Saved Sucessfully')", true);

                                                ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                                    "alert('Your Request Saved Sucessfully')", true);
                                            }
                                            else
                                            {
                                                viewErrorMsg();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // If Login is DVMgr

                                        //string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
                                        //string employeeName = drpLstEmpNames.SelectedItem.Text.Trim();
                                        //string pageName =
                                        //    System.IO.Path.GetFileName(HttpContext.Current.Request.PhysicalPath);

                                        //string hrEmail = CommonMethods.GetLookups("hrEmail");

                                        //MailTemplates = CommonMethods.LoadMailTemplates();

                                        //if (MailTemplates["ToHrTemplate"] != null)
                                        //{
                                        //    if (CommonMethods.SendMail(hrEmail.Trim(),
                                        //        "TNA for Employee :" + employeeName,
                                        //        string.Format(MailTemplates["ToHrTemplate"].ToString(),
                                        //            txtEmpCode.Text.Trim(), employeeName, txtEmpPosition.Text.Trim(),
                                        //            currentAppUrl + pageName, TNAId), SiteId, WebId))
                                        //    {
                                        //        ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                                        //            "Success Message",
                                        //            "alert('Your Request Sent To HR Sucessfully')", true);
                                        //    }
                                        //    else
                                        //    {
                                        //        viewErrorMsg();
                                        //    }
                                        //}
                                    }


                                    //*********************************
                                }
                            }
                        }
                    });
            }
            catch (Exception)
            {
            }
        }

        protected void btnRejectTNAReqByDH_OnClick(object sender, EventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {
                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            Guid SiteId = site.ID;
                            Guid WebId = web.ID;
                            SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                            if (EmpCoursesLst != null)
                            {
                                if (HttpContext.Current.Request.QueryString["pageTo"] != null &&
                                    HttpContext.Current.Request.QueryString["TNAId"] != null)
                                {
                                    string TNAId = HttpContext.Current.Request.QueryString["TNAId"];
                                    SPListItem empCourseItem = EmpCoursesLst.GetItemById(int.Parse(TNAId));

                                    empCourseItem["RequestStatus"] = "Rejected By Dept. Head";
                                    string empSupervisorCode = empCourseItem["SupervisorCode"].ToString();
                                    empCourseItem.Update();

                                    // Send rejection mail to Supervisor
                                    string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
                                    string employeeName = drpLstEmpNames.SelectedItem.Text.Trim();
                                    string pageName = System.IO.Path.GetFileName(HttpContext.Current.Request.PhysicalPath);

                                    string supervisorMail = CommonMethods.GetEmpMailByEmpCode(empSupervisorCode);
                                    /*"mgeady@nslab.lab";*/
                                    MailTemplates = CommonMethods.LoadMailTemplates();

                                    if (MailTemplates["DHToSupervisorRejectTemplate"] != null)
                                    {
                                        if (CommonMethods.SendMail(supervisorMail.Trim(),
                                            "TNA for Employee :" + employeeName,
                                            string.Format(MailTemplates["DHToSupervisorRejectTemplate"].ToString(),
                                                txtEmpCode.Text.Trim(), employeeName, txtEmpPosition.Text.Trim()), SiteId, WebId))
                                        {
                                            //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "Success Message",
                                            //    "alert('Your Request Rejected Sucessfully')", true);

                                            ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                                "alert('Your Request Rejected Sucessfully')", true);
                                        }
                                        else
                                        {
                                            viewErrorMsg();
                                        }
                                    }
                                }
                            }
                        }
                    });
            }
            catch (Exception)
            {
            }
        }

        protected void btnApproveTNAReqByTopMgr_OnClick(object sender, EventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {

                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            Guid SiteId = site.ID;
                            Guid WebId = web.ID;
                            SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                            if (EmpCoursesLst != null)
                            {
                                if (HttpContext.Current.Request.QueryString["pageTo"] != null &&
                                    HttpContext.Current.Request.QueryString["TNAId"] != null)
                                {
                                    web.AllowUnsafeUpdates = true;
                                    string TNAId = HttpContext.Current.Request.QueryString["TNAId"];
                                    SPListItem empCourseItem = EmpCoursesLst.GetItemById(int.Parse(TNAId));

                                    empCourseItem["RequestStatus"] = "Sent to HR";
                                    empCourseItem["DVMgrCode"] = CommonMethods.GetLoginEmpCode();
                                    empCourseItem["DHUser"] = "DV Mgr";
                                    empCourseItem.Update();

                                    web.AllowUnsafeUpdates = false;

                                    //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "Success Message",
                                    //           "alert('Your Request Saved Sucessfully')", true);

                                    ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                               "alert('Your Request Saved Sucessfully')", true);

                                    //string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
                                    //string employeeName = drpLstEmpNames.SelectedItem.Text.Trim();
                                    //string pageName = System.IO.Path.GetFileName(HttpContext.Current.Request.PhysicalPath);

                                    //string hrEmail = CommonMethods.GetLookups("hrEmail");

                                    //MailTemplates = CommonMethods.LoadMailTemplates();

                                    //if (MailTemplates["ToHrTemplate"] != null)
                                    //{
                                    //    if (CommonMethods.SendMail(hrEmail.Trim(),
                                    //        "TNA for Employee :" + employeeName,
                                    //        string.Format(MailTemplates["ToHrTemplate"].ToString(),
                                    //            txtEmpCode.Text.Trim(), employeeName, txtEmpPosition.Text.Trim(),
                                    //            currentAppUrl + pageName, TNAId), SiteId, WebId))
                                    //    {
                                    //        ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "Success Message",
                                    //            "alert('Your Request Sent To HR Sucessfully')", true);
                                    //    }
                                    //    else
                                    //    {
                                    //        viewErrorMsg();
                                    //    }
                                    //}



                                    //*********************************
                                }
                            }
                        }
                    });
            }
            catch (Exception)
            {
            }
        }


        protected void btnRejectTNAReqByTopMgr_OnClick(object sender, EventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {
                        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                        using (SPWeb web = site.OpenWeb())
                        {
                            Guid SiteId = site.ID;
                            Guid WebId = web.ID;
                            SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                            if (EmpCoursesLst != null)
                            {
                                if (HttpContext.Current.Request.QueryString["pageTo"] != null &&
                                    HttpContext.Current.Request.QueryString["TNAId"] != null)
                                {
                                    string TNAId = HttpContext.Current.Request.QueryString["TNAId"];
                                    SPListItem empCourseItem = EmpCoursesLst.GetItemById(int.Parse(TNAId));

                                    empCourseItem["RequestStatus"] = "Rejected By Division Manger";
                                    string empSupervisorCode = empCourseItem["SupervisorCode"].ToString();
                                    empCourseItem.Update();

                                    // Send rejection mail to Supervisor
                                    string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
                                    string employeeName = drpLstEmpNames.SelectedItem.Text.Trim();
                                    string pageName = System.IO.Path.GetFileName(HttpContext.Current.Request.PhysicalPath);

                                    string mgrMail = CommonMethods.GetEmpMailByEmpCode(empSupervisorCode);
                                    /* "mgeady@nslab.lab";*/
                                    MailTemplates = CommonMethods.LoadMailTemplates();

                                    if (MailTemplates["TopMgrToSupervisorRejectTemplate"] != null)
                                    {
                                        if (CommonMethods.SendMail(mgrMail.Trim(),
                                            "TNA for Employee :" + employeeName,
                                            string.Format(MailTemplates["TopMgrToSupervisorRejectTemplate"].ToString(),
                                                txtEmpCode.Text.Trim(), employeeName, txtEmpPosition.Text.Trim()), SiteId, WebId))
                                        {
                                            //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "Success Message",
                                            //    "alert('Your Request Saved Sucessfully')", true);

                                            ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), "Success Message",
                                                "alert('Your Request Saved Sucessfully')", true);

                                        }
                                        else
                                        {
                                            viewErrorMsg();
                                        }
                                    }
                                }
                            }
                        }
                    });
            }
            catch (Exception)
            {

            }
        }





    }
}

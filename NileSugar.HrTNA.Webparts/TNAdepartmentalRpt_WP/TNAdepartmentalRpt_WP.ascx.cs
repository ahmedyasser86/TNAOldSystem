using Microsoft.SharePoint;
using NileSugar.HrTNA.Webparts.TrainingRequests_WP;
using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace NileSugar.HrTNA.Webparts.TNAdepartmentalRpt_WP
{
    [ToolboxItemAttribute(false)]
    public partial class TNAdepartmentalRpt_WP : WebPart
    {
        // Uncomment the following SecurityPermission attribute only when doing Performance Profiling on a farm solution
        // using the Instrumentation method, and then remove the SecurityPermission attribute when the code is ready
        // for production. Because the SecurityPermission attribute bypasses the security check for callers of
        // your constructor, it's not recommended for production purposes.
        // [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Assert, UnmanagedCode = true)]

        #region Variables
        private int EmptyDepartmentCourses = 0;
        #endregion

        #region Functions

        private void LoadDepartment()
        {
            try
            {

                DataTable Department_dt = CommonMethods.GetAllDepartments();
                Department_dt.DefaultView.Sort = "Department ASC";
                ddl_Department.DataSource = Department_dt;
                ddl_Department.DataTextField = "Department";
                ddl_Department.DataValueField = "Department";
                ddl_Department.DataBind();
                ddl_Department.Items.Insert(0, new ListItem("-- Select --", "-1"));
            }
            catch { }
        }

        private void loadTNAYear()
        {
            try
            {
                for (int i = 0; i < 5; i++)
                {
                    ddl_TNAYear.Items.Insert(i, DateTime.Now.AddYears(i - 5).Year.ToString());
                }
                ddl_TNAYear.SelectedItem.Text = CommonMethods.GetLookups("TNAYear");
            }
            catch (Exception ex)
            {
                CommonMethods.UpdateTestLst(ex.ToString());
            }

        }
        private DataTable Get_BasicSkills_Courses()
        {
            DataTable EmpCoursesBasic_DT = new DataTable();
            try
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                using (SPWeb web = site.OpenWeb())
                {
                    SPList EmpBasicSkillsOtherCoursesLst = web.Lists.TryGetList("EmpBasicSkillsOtherCourses");
                    if (EmpBasicSkillsOtherCoursesLst != null)
                    {
                        SPQuery qry = new SPQuery();

                        qry.Query = String.Format(
                            @"  <OrderBy>
                                  <FieldRef Name='CourseName' Ascending='True' />
                               </OrderBy> ");

                        SPListItemCollection collItems = EmpBasicSkillsOtherCoursesLst.GetItems(qry);
                        if (collItems.Count > 0)
                        {
                            EmpCoursesBasic_DT = collItems.GetDataTable();
                        }
                        else
                        {
                            EmpCoursesBasic_DT.Columns.Add("CourseName", typeof(String));
                            EmpCoursesBasic_DT.Columns.Add("ExpectedQuarter", typeof(String));
                        }
                    }

                    SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                    if (EmpCoursesLst != null)
                    {
                        SPQuery qry = new SPQuery();
                        qry.Query = string.Format(@"<Where>
                                                              <Eq>
                                                                 <FieldRef Name='TNAYear' />
                                                                 <Value Type='Text'>{0}</Value>
                                                              </Eq>
                                                           </Where>", DateTime.Now.Year);
                        SPListItemCollection coll = EmpCoursesLst.GetItems(qry);
                        DataTable EmpCourses_DT = coll.GetDataTable();

                        for (int i = 0; i < EmpCourses_DT.Rows.Count; i++)
                        {
                            DataRow dr1 = EmpCoursesBasic_DT.NewRow();
                            dr1["CourseName"] = EmpCourses_DT.Rows[i]["BS_EngC_Name"];
                            dr1["ExpectedQuarter"] = EmpCourses_DT.Rows[i]["BS_EngC_ExpectedQuarter"];

                            DataRow dr2 = EmpCoursesBasic_DT.NewRow();
                            dr2["CourseName"] = EmpCourses_DT.Rows[i]["BS_CompSkillsC_Name"];
                            dr2["ExpectedQuarter"] = EmpCourses_DT.Rows[i]["BS_CompSkillsC_ExpectedQuarter"];

                            DataRow dr3 = EmpCoursesBasic_DT.NewRow();
                            dr3["CourseName"] = EmpCourses_DT.Rows[i]["BS_BEPC_Name"];
                            dr3["ExpectedQuarter"] = EmpCourses_DT.Rows[i]["BS_BEPC_ExpectedQuarter"];

                            if (dr1 != null && dr1["CourseName"].ToString() != "" && dr1["CourseName"].ToString() != null)
                                EmpCoursesBasic_DT.Rows.Add(dr1);

                            if (dr2 != null && dr2["CourseName"].ToString() != "" && dr2["CourseName"].ToString() != null)
                                EmpCoursesBasic_DT.Rows.Add(dr2);

                            if (dr3 != null && dr3["CourseName"].ToString() != "" && dr3["CourseName"].ToString() != null)
                                EmpCoursesBasic_DT.Rows.Add(dr3);
                        }

                    }
                    EmpCoursesBasic_DT = CommonMethods.RemoveDuplicates(EmpCoursesBasic_DT, "CourseName");
                }

            }
            catch { }
            return EmpCoursesBasic_DT;
        }


        private DataTable Get_DevelopmentSkills_Courses()
        {
            DataTable EmpCoursesDev_DT = new DataTable();
            try
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
                                   <OrderBy>
                                      <FieldRef Name='CourseName' Ascending='True' />
                                   </OrderBy>  ");

                        SPListItemCollection collItems = EmpDevSkillsOtherCoursesLst.GetItems(qry);
                        if (collItems.Count > 0)
                        {
                            EmpCoursesDev_DT = collItems.GetDataTable();
                        }
                        else
                        {
                            EmpCoursesDev_DT.Columns.Add("CourseName", typeof(String));
                            EmpCoursesDev_DT.Columns.Add("ExpectedQuarter", typeof(String));
                        }
                    }

                    SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                    if (EmpCoursesLst != null)
                    {
                        SPQuery qry = new SPQuery();
                        qry.Query = string.Format(@"<Where>
                                                              <Eq>
                                                                 <FieldRef Name='TNAYear' />
                                                                 <Value Type='Text'>{0}</Value>
                                                              </Eq>
                                                           </Where>", DateTime.Now.Year);
                        SPListItemCollection coll = EmpCoursesLst.GetItems(qry);

                        DataTable EmpCourses_DT = coll.GetDataTable();

                        for (int i = 0; i < EmpCourses_DT.Rows.Count; i++)
                        {
                            DataRow dr1 = EmpCoursesDev_DT.NewRow();
                            dr1["CourseName"] = EmpCourses_DT.Rows[i]["DevS_CommSC_Name"];
                            dr1["ExpectedQuarter"] = EmpCourses_DT.Rows[i]["DevS_CommSC_ExpectedQuarter"];

                            DataRow dr2 = EmpCoursesDev_DT.NewRow();
                            dr2["CourseName"] = EmpCourses_DT.Rows[i]["DevS_InterPersSC_Name"];
                            dr2["ExpectedQuarter"] = EmpCourses_DT.Rows[i]["DevS_InterPersSC_ExpectedQuarter"];

                            DataRow dr3 = EmpCoursesDev_DT.NewRow();
                            dr3["CourseName"] = EmpCourses_DT.Rows[i]["DevS_BIC_Name"];
                            dr3["ExpectedQuarter"] = EmpCourses_DT.Rows[i]["DevS_BIC_ExpectedQuarter"];

                            if (dr1 != null && dr1["CourseName"].ToString() != "" && dr1["CourseName"].ToString() != null)
                                EmpCoursesDev_DT.Rows.Add(dr1);

                            if (dr2 != null && dr2["CourseName"].ToString() != "" && dr2["CourseName"].ToString() != null)
                                EmpCoursesDev_DT.Rows.Add(dr2);

                            if (dr3 != null && dr3["CourseName"].ToString() != "" && dr3["CourseName"].ToString() != null)
                                EmpCoursesDev_DT.Rows.Add(dr3);
                        }
                    }
                    EmpCoursesDev_DT = CommonMethods.RemoveDuplicates(EmpCoursesDev_DT, "CourseName");
                }

            }
            catch { }
            return EmpCoursesDev_DT;
        }
        private DataTable Get_TechnicalSkills_Courses()
        {
            DataTable EmpCoursesBasic_DT = new DataTable();

            try
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                using (SPWeb web = site.OpenWeb())
                {

                    SPList EmpTechnicalCoursesLst = web.Lists.TryGetList("EmpTechnicalCoursesLst");
                    if (EmpTechnicalCoursesLst != null)
                    {
                        SPQuery qry = new SPQuery();

                        qry.Query = String.Format(@" <OrderBy>
                                      <FieldRef Name='CourseName' Ascending='True' />
                                   </OrderBy>");

                        SPListItemCollection collItems = EmpTechnicalCoursesLst.GetItems(qry);
                        if (collItems.Count > 0)
                        {
                            EmpCoursesBasic_DT = CommonMethods.RemoveDuplicates(collItems.GetDataTable(), "CourseName");
                        }
                    }


                }

            }

            catch { }
            return EmpCoursesBasic_DT;


        }

        private string FormateDate(DateTime ConvertedDate)
        {
            string Date_Formating = ConvertedDate.Day + "/" + ConvertedDate.Month + "/" + ConvertedDate.Year;
            return Date_Formating;
        }
        private DataTable GetEmployeeCourses(string Emp_Department, string TNAYear)
        {
            // Emp_Department = "Technical Department"; // will be deleted

            #region Create Binding Table

            DataTable EmpCourses_Dt = new DataTable();
            EmpCourses_Dt.Columns.Add("EmpCode", typeof(String));
            EmpCourses_Dt.Columns.Add("EmpName", typeof(String));
            EmpCourses_Dt.Columns.Add("POSITION", typeof(String));
            EmpCourses_Dt.Columns.Add("Category", typeof(String));
            EmpCourses_Dt.Columns.Add("CourseName", typeof(String));
            EmpCourses_Dt.Columns.Add("ExpectedQuarter", typeof(String));
            EmpCourses_Dt.Columns.Add("TrainingProvider", typeof(String));
            EmpCourses_Dt.Columns.Add("CreatedDate", typeof(DateTime));
            EmpCourses_Dt.Columns.Add("Priority", typeof(String));

            #endregion

            try
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                using (SPWeb web = site.OpenWeb())
                {

                    #region Get EmpCourses

                    SPList EmpCourses = web.Lists.TryGetList("EmpCourses");
                    if (EmpCourses != null)
                    {
                        SPQuery qry = new SPQuery();

                        qry.Query = String.Format(@"<Where>
                                                          <And>
                                                             <Eq>
                                                                <FieldRef Name='Department' />
                                                                <Value Type='Text'>{0}</Value>
                                                             </Eq>
                                                             <Eq>
                                                                <FieldRef Name='TNAYear' />
                                                                <Value Type='Text'>{1}</Value>
                                                             </Eq>
                                                          </And>
                                                       </Where>
                                                       <OrderBy>
                                                          <FieldRef Name='EmpCode' Ascending='True' />
                                                       </OrderBy>", Emp_Department.Trim(), TNAYear);

                        SPListItemCollection collItems = EmpCourses.GetItems(qry);
                        if (collItems.Count > 0)
                        {
                            DataTable EmpCoursesBasic_DT = collItems.GetDataTable();

                            if (EmpCoursesBasic_DT != null && EmpCoursesBasic_DT.Rows.Count > 0)
                            {
                                for (int i = 0; i < EmpCoursesBasic_DT.Rows.Count; i++)
                                {
                                    if (ddl_CoursesCategory.SelectedValue == "-1" || ddl_CoursesCategory.SelectedValue == "0")//case get all categories or basic only
                                    {
                                        #region Basic
                                        if (EmpCoursesBasic_DT.Rows[i]["BS_EngC_Name"] != null && EmpCoursesBasic_DT.Rows[i]["BS_EngC_Name"].ToString() != "")
                                        {
                                            DataRow dr1 = EmpCourses_Dt.NewRow();
                                            dr1["EmpCode"] = EmpCoursesBasic_DT.Rows[i]["EmpCode"];
                                            dr1["EmpName"] = EmpCoursesBasic_DT.Rows[i]["EmpName"];
                                            dr1["POSITION"] = EmpCoursesBasic_DT.Rows[i]["POSITION"];
                                            dr1["Category"] = "Basic Skills";
                                            dr1["CourseName"] = EmpCoursesBasic_DT.Rows[i]["BS_EngC_Name"];
                                            dr1["ExpectedQuarter"] = EmpCoursesBasic_DT.Rows[i]["BS_EngC_ExpectedQuarter"];
                                            dr1["CreatedDate"] = Convert.ToDateTime(EmpCoursesBasic_DT.Rows[i]["Created"]);
                                            dr1["Priority"] = EmpCoursesBasic_DT.Rows[i]["BS_EngC_priority"];
                                            EmpCourses_Dt.Rows.Add(dr1);
                                        }

                                        if (EmpCoursesBasic_DT.Rows[i]["BS_CompSkillsC_Name"] != null && EmpCoursesBasic_DT.Rows[i]["BS_CompSkillsC_Name"].ToString() != "")
                                        {
                                            DataRow dr2 = EmpCourses_Dt.NewRow();
                                            dr2["EmpCode"] = EmpCoursesBasic_DT.Rows[i]["EmpCode"];
                                            dr2["EmpName"] = EmpCoursesBasic_DT.Rows[i]["EmpName"];
                                            dr2["POSITION"] = EmpCoursesBasic_DT.Rows[i]["POSITION"];
                                            dr2["Category"] = "Basic Skills";
                                            dr2["CourseName"] = EmpCoursesBasic_DT.Rows[i]["BS_CompSkillsC_Name"];
                                            dr2["ExpectedQuarter"] = EmpCoursesBasic_DT.Rows[i]["BS_CompSkillsC_ExpectedQuarter"];
                                            dr2["CreatedDate"] = Convert.ToDateTime(EmpCoursesBasic_DT.Rows[i]["Created"]);
                                            dr2["Priority"] = EmpCoursesBasic_DT.Rows[i]["BS_CompSkillsC_priority"];
                                            EmpCourses_Dt.Rows.Add(dr2);
                                        }

                                        if (EmpCoursesBasic_DT.Rows[i]["BS_BEPC_Name"] != null && EmpCoursesBasic_DT.Rows[i]["BS_BEPC_Name"].ToString() != "")
                                        {
                                            DataRow dr3 = EmpCourses_Dt.NewRow();
                                            dr3["EmpCode"] = EmpCoursesBasic_DT.Rows[i]["EmpCode"];
                                            dr3["EmpName"] = EmpCoursesBasic_DT.Rows[i]["EmpName"];
                                            dr3["POSITION"] = EmpCoursesBasic_DT.Rows[i]["POSITION"];
                                            dr3["Category"] = "Basic Skills";
                                            dr3["CourseName"] = EmpCoursesBasic_DT.Rows[i]["BS_BEPC_Name"];
                                            dr3["ExpectedQuarter"] = EmpCoursesBasic_DT.Rows[i]["BS_BEPC_ExpectedQuarter"];
                                            dr3["CreatedDate"] = Convert.ToDateTime(EmpCoursesBasic_DT.Rows[i]["Created"]);
                                            dr3["Priority"] = EmpCoursesBasic_DT.Rows[i]["BS_BEPC_priority"];
                                            EmpCourses_Dt.Rows.Add(dr3);
                                        }
                                        #endregion
                                    }

                                    if (ddl_CoursesCategory.SelectedValue == "-1" || ddl_CoursesCategory.SelectedValue == "1")//case get all categories or Development
                                    {
                                        #region Development
                                        if (EmpCoursesBasic_DT.Rows[i]["DevS_CommSC_Name"] != null && EmpCoursesBasic_DT.Rows[i]["DevS_CommSC_Name"].ToString() != "")
                                        {
                                            DataRow dr1 = EmpCourses_Dt.NewRow();
                                            dr1["EmpCode"] = EmpCoursesBasic_DT.Rows[i]["EmpCode"];
                                            dr1["EmpName"] = EmpCoursesBasic_DT.Rows[i]["EmpName"];
                                            dr1["POSITION"] = EmpCoursesBasic_DT.Rows[i]["POSITION"];
                                            dr1["Category"] = "Development Skills";
                                            dr1["CourseName"] = EmpCoursesBasic_DT.Rows[i]["DevS_CommSC_Name"];
                                            dr1["ExpectedQuarter"] = EmpCoursesBasic_DT.Rows[i]["DevS_CommSC_ExpectedQuarter"];
                                            dr1["CreatedDate"] = Convert.ToDateTime(EmpCoursesBasic_DT.Rows[i]["Created"]);
                                            //Rehab
                                            //dr1["Priority"] = EmpCoursesBasic_DT.Rows[i]["DevS_CommSC_priority"];
                                            EmpCourses_Dt.Rows.Add(dr1);
                                        }

                                        if (EmpCoursesBasic_DT.Rows[i]["DevS_InterPersSC_Name"] != null && EmpCoursesBasic_DT.Rows[i]["DevS_InterPersSC_Name"].ToString() != "")
                                        {
                                            DataRow dr1 = EmpCourses_Dt.NewRow();
                                            dr1["EmpCode"] = EmpCoursesBasic_DT.Rows[i]["EmpCode"];
                                            dr1["EmpName"] = EmpCoursesBasic_DT.Rows[i]["EmpName"];
                                            dr1["POSITION"] = EmpCoursesBasic_DT.Rows[i]["POSITION"];
                                            dr1["Category"] = "Development Skills";
                                            dr1["CourseName"] = EmpCoursesBasic_DT.Rows[i]["DevS_InterPersSC_Name"];
                                            dr1["ExpectedQuarter"] = EmpCoursesBasic_DT.Rows[i]["DevS_InterPersSC_ExpectedQuarter"];
                                            dr1["CreatedDate"] = Convert.ToDateTime(EmpCoursesBasic_DT.Rows[i]["Created"]);
                                            dr1["Priority"] = EmpCoursesBasic_DT.Rows[i]["DevS_InterPersSC_priority"];
                                            EmpCourses_Dt.Rows.Add(dr1);
                                        }

                                        if (EmpCoursesBasic_DT.Rows[i]["DevS_BIC_Name"] != null && EmpCoursesBasic_DT.Rows[i]["DevS_BIC_Name"].ToString() != "")
                                        {
                                            DataRow dr1 = EmpCourses_Dt.NewRow();
                                            dr1["EmpCode"] = EmpCoursesBasic_DT.Rows[i]["EmpCode"];
                                            dr1["EmpName"] = EmpCoursesBasic_DT.Rows[i]["EmpName"];
                                            dr1["POSITION"] = EmpCoursesBasic_DT.Rows[i]["POSITION"];
                                            dr1["Category"] = "Development Skills";
                                            dr1["CourseName"] = EmpCoursesBasic_DT.Rows[i]["DevS_BIC_Name"];
                                            dr1["ExpectedQuarter"] = EmpCoursesBasic_DT.Rows[i]["DevS_BIC_ExpectedQuarter"];
                                            dr1["CreatedDate"] = Convert.ToDateTime(EmpCoursesBasic_DT.Rows[i]["Created"]);
                                            dr1["Priority"] = EmpCoursesBasic_DT.Rows[i]["DevS_BIC_priority"];
                                            EmpCourses_Dt.Rows.Add(dr1);
                                        }

                                        #endregion
                                    }
                                }
                            }

                        }
                    }

                    #endregion

                    if (ddl_CoursesCategory.SelectedValue == "-1" || ddl_CoursesCategory.SelectedValue == "0")//case get all categories or basic only
                    {
                        #region Get EmpBasicSkillsOtherCourses

                        SPList EmpBasicSkillsOtherCourses = web.Lists.TryGetList("EmpBasicSkillsOtherCourses");
                        if (EmpCourses != null)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = String.Format(@" <Where>
                                                          <And>
                                                             <Eq>
                                                                <FieldRef Name='Department' />
                                                                <Value Type='Text'>{0}</Value>
                                                             </Eq>
                                                             <Eq>
                                                                <FieldRef Name='TNAYear' />
                                                                <Value Type='Text'>{1}</Value>
                                                             </Eq>
                                                          </And>
                                                       </Where>
                                                       <OrderBy>
                                                          <FieldRef Name='EmpCode' Ascending='True' />
                                                       </OrderBy>", Emp_Department.Trim(), TNAYear);

                            SPListItemCollection collItems2 = EmpBasicSkillsOtherCourses.GetItems(qry);
                            if (collItems2 != null)
                            {
                                if (collItems2.Count > 0)
                                {
                                    DataTable EmpBasicSkillsOtherCourses_DT = collItems2.GetDataTable();
                                    for (int i = 0; i < EmpBasicSkillsOtherCourses_DT.Rows.Count; i++)
                                    {
                                        DataRow dr = EmpCourses_Dt.NewRow();
                                        dr["EmpCode"] = EmpBasicSkillsOtherCourses_DT.Rows[i]["EmpCode"];
                                        dr["EmpName"] = CommonMethods.GetEmpNameByCode(EmpBasicSkillsOtherCourses_DT.Rows[i]["EmpCode"].ToString());
                                        dr["POSITION"] = CommonMethods.GetEmpPossitionByCode(EmpBasicSkillsOtherCourses_DT.Rows[i]["EmpCode"].ToString());
                                        dr["Category"] = "Basic Skills";
                                        dr["CourseName"] = EmpBasicSkillsOtherCourses_DT.Rows[i]["CourseName"];
                                        dr["ExpectedQuarter"] = EmpBasicSkillsOtherCourses_DT.Rows[i]["ExpectedQuarter"];
                                        dr["CreatedDate"] = Convert.ToDateTime(EmpBasicSkillsOtherCourses_DT.Rows[i]["Created"]);
                                        dr["Priority"] = EmpBasicSkillsOtherCourses_DT.Rows[i]["Priority"];
                                        EmpCourses_Dt.Rows.Add(dr);
                                    }

                                }
                            }
                        }

                        #endregion
                    }

                    if (ddl_CoursesCategory.SelectedValue == "-1" || ddl_CoursesCategory.SelectedValue == "1")//case get all categories or Development only
                    {
                        #region Get EmpDevSkillsOtherCourses

                        SPList EmpDevSkillsOtherCourses = web.Lists.TryGetList("EmpDevSkillsOtherCourses");
                        if (EmpCourses != null)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = String.Format(@" <Where>
                                                          <And>
                                                             <Eq>
                                                                <FieldRef Name='Department' />
                                                                <Value Type='Text'>{0}</Value>
                                                             </Eq>
                                                             <Eq>
                                                                <FieldRef Name='TNAYear' />
                                                                <Value Type='Text'>{1}</Value>
                                                             </Eq>
                                                          </And>
                                                       </Where>
                                                       <OrderBy>
                                                          <FieldRef Name='EmpCode' Ascending='True' />
                                                       </OrderBy>", Emp_Department, TNAYear);


                            SPListItemCollection collItems3 = EmpDevSkillsOtherCourses.GetItems(qry);
                            if (collItems3 != null)
                            {
                                if (collItems3.Count > 0)
                                {
                                    DataTable EmpDevSkillsOtherCourses_DT = collItems3.GetDataTable();
                                    for (int i = 0; i < EmpDevSkillsOtherCourses_DT.Rows.Count; i++)
                                    {
                                        DataRow dr = EmpCourses_Dt.NewRow();
                                        dr["EmpCode"] = EmpDevSkillsOtherCourses_DT.Rows[i]["EmpCode"];
                                        dr["EmpName"] = CommonMethods.GetEmpNameByCode(EmpDevSkillsOtherCourses_DT.Rows[i]["EmpCode"].ToString());
                                        dr["POSITION"] = CommonMethods.GetEmpPossitionByCode(EmpDevSkillsOtherCourses_DT.Rows[i]["EmpCode"].ToString());
                                        dr["Category"] = "Development Skills";
                                        dr["CourseName"] = EmpDevSkillsOtherCourses_DT.Rows[i]["CourseName"];
                                        dr["ExpectedQuarter"] = EmpDevSkillsOtherCourses_DT.Rows[i]["ExpectedQuarter"];
                                        dr["CreatedDate"] = Convert.ToDateTime(EmpDevSkillsOtherCourses_DT.Rows[i]["Created"]);
                                        dr["Priority"] = EmpDevSkillsOtherCourses_DT.Rows[i]["Priority"];
                                        EmpCourses_Dt.Rows.Add(dr);
                                    }
                                }
                            }
                        }

                        #endregion
                    }

                    if (ddl_CoursesCategory.SelectedValue == "-1" || ddl_CoursesCategory.SelectedValue == "2")//case get all categories or Technical only
                    {
                        #region Get EmpTechnicalCoursesLst

                        SPList EmpTechnicalCoursesLst = web.Lists.TryGetList("EmpTechnicalCoursesLst");
                        if (EmpCourses != null)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = String.Format(@" <Where>
                                                      <And>
                                                         <Eq>
                                                            <FieldRef Name='DeptName' />
                                                            <Value Type='Text'>{0}</Value>
                                                         </Eq>
                                                         <Eq>
                                                            <FieldRef Name='TNAYear' />
                                                            <Value Type='Text'>{1}</Value>
                                                         </Eq>
                                                      </And>
                                                   </Where>
                                                   <OrderBy>
                                                      <FieldRef Name='EmpCode' Ascending='True' />
                                                   </OrderBy>", Emp_Department, TNAYear);

                            SPListItemCollection collItems4 = EmpTechnicalCoursesLst.GetItems(qry);
                            if (collItems4 != null)
                            {
                                if (collItems4.Count > 0)
                                {
                                    DataTable EmpTechnicalCourses_DT = collItems4.GetDataTable();
                                    for (int i = 0; i < EmpTechnicalCourses_DT.Rows.Count; i++)
                                    {
                                        DataRow dr = EmpCourses_Dt.NewRow();
                                        dr["EmpCode"] = EmpTechnicalCourses_DT.Rows[i]["EmpCode"];
                                        dr["EmpName"] = CommonMethods.GetEmpNameByCode(EmpTechnicalCourses_DT.Rows[i]["EmpCode"].ToString());
                                        dr["POSITION"] = CommonMethods.GetEmpPossitionByCode(EmpTechnicalCourses_DT.Rows[i]["EmpCode"].ToString());
                                        dr["Category"] = "Technical Skills";
                                        dr["CourseName"] = EmpTechnicalCourses_DT.Rows[i]["CourseName"];
                                        dr["ExpectedQuarter"] = EmpTechnicalCourses_DT.Rows[i]["ExpectedQuarter"];
                                        dr["TrainingProvider"] = EmpTechnicalCourses_DT.Rows[i]["TrainingProvider"];
                                        dr["CreatedDate"] = Convert.ToDateTime(EmpTechnicalCourses_DT.Rows[i]["Created"]);
                                        dr["Priority"] = EmpTechnicalCourses_DT.Rows[i]["Priority"];
                                        EmpCourses_Dt.Rows.Add(dr);
                                    }
                                }
                            }
                        }

                        #endregion
                    }
                }
            }
            catch (Exception ex) { CommonMethods.UpdateTestLst(ex.ToString()); }

            return EmpCourses_Dt;
        }


        public string GetRequestStatus(string Emp_Code)
        {
            string RequestStatus = string.Empty;
            try
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                using (SPWeb web = site.OpenWeb())
                {

                    string TNAYear = string.Empty;
                    if (ddl_TNAYear.SelectedItem.Text != CommonMethods.GetLookups("TNAYear"))
                        TNAYear = ddl_TNAYear.SelectedItem.Text;
                    else
                        TNAYear = CommonMethods.GetLookups("TNAYear");


                    SPList EmpCourses = web.Lists.TryGetList("EmpCourses");
                    if (EmpCourses != null)
                    {
                        SPQuery qry = new SPQuery();

                        qry.Query = String.Format(@"<Where>
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
                                                   </Where>", Emp_Code, TNAYear);
                        SPListItemCollection collItems = EmpCourses.GetItems(qry);
                        if (collItems.Count > 0)
                        {
                            RequestStatus = (collItems.GetDataTable()).Rows[0]["RequestStatus"].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CommonMethods.UpdateTestLst(ex.ToString());
            }
            return RequestStatus;
        }
        private void CheckTNAYearSearch()
        {

            lbl_Fr0mMsg.ForeColor = Color.Red;
            lbl_Fr0mMsg.Font.Size = 8;
            lbl_ToMsg.ForeColor = Color.Red;
            lbl_ToMsg.Font.Size = 8;

            #region cal_From With ddl_TNAYear
            if (!cal_From.IsDateEmpty)
            {
                if (ddl_TNAYear.SelectedItem.Text != cal_From.SelectedDate.Year.ToString())
                {
                    lbl_Fr0mMsg.Visible = true;
                    lbl_Fr0mMsg.Text = "TNA year must like calender year";

                    return;
                }
                else
                { lbl_Fr0mMsg.Visible = false; }
            }
            else
            { lbl_Fr0mMsg.Visible = false; }
            #endregion

            #region cal_To With ddl_TNAYear
            if (!cal_To.IsDateEmpty)
            {
                if (ddl_TNAYear.SelectedItem.Text != cal_To.SelectedDate.Year.ToString())
                {
                    lbl_ToMsg.Visible = true;
                    lbl_ToMsg.Text = "TNA year must like calender year";

                    return;
                }
                else
                { lbl_ToMsg.Visible = false; }
            }
            else
            { lbl_ToMsg.Visible = false; }

            #endregion

            #region cal_To With cal_From
            if (cal_To.SelectedDate < cal_From.SelectedDate)
            {
                lbl_ToMsg.Visible = true;
                lbl_ToMsg.Text = "Date From must be less than date to";

                return;
            }
            else
            { lbl_ToMsg.Visible = false; }
            #endregion

        }
        #endregion

        public TNAdepartmentalRpt_WP()
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
                HttpContext.Current.Response.Redirect("http://sharepoint.ns.local/forms/hr/TNA/SitePages/TNAReport.aspx");

            if (!Page.IsPostBack)
            {
                LoadDepartment();
                loadTNAYear();
                ddl_CourseName.Items.Insert(0, new ListItem("-- Select --", "-1"));
            }
        }

        protected void btn_Search_Click(object sender, EventArgs e)
        {
            CheckTNAYearSearch();
            DataTable Departments_DT = CommonMethods.GetAllDepartments();
            if (Departments_DT != null && Departments_DT.Rows.Count > 0)
            {
                NoRecords.Visible = false;
                Rpt_EmpCources.Visible = true;
                btn_Excel.Visible = true;

                #region filter by department

                if (ddl_Department.SelectedItem.Text != "-- Select --")
                {
                    DataRow[] rows = Departments_DT.Select("Department= '" + ddl_Department.SelectedItem.Text.Trim() + "'");
                    if (rows != null && rows.Length > 0)
                    {
                        NoRecords.Visible = false;
                        Rpt_EmpCources.Visible = true;
                        btn_Excel.Visible = true;
                        Rpt_EmpCources.DataSource = rows.CopyToDataTable();
                        Rpt_EmpCources.DataBind();


                        //case there isnot courses for selected department
                        if (EmptyDepartmentCourses == rows.Length)
                        {
                            NoRecords.Visible = true;
                            Rpt_EmpCources.Visible = false;
                            btn_Excel.Visible = false;
                        }
                    }
                    else
                    {
                        NoRecords.Visible = true;
                        Rpt_EmpCources.Visible = false;
                        btn_Excel.Visible = false;
                    }


                }
                #endregion

                else
                {
                    Rpt_EmpCources.DataSource = Departments_DT;
                    Rpt_EmpCources.DataBind();
                    //case there isnot courses for all department 
                    if (EmptyDepartmentCourses == Departments_DT.Rows.Count)
                    {
                        NoRecords.Visible = true;
                        Rpt_EmpCources.Visible = false;
                        btn_Excel.Visible = false;
                    }
                    else
                    {
                        NoRecords.Visible = false;
                        Rpt_EmpCources.Visible = true;
                        btn_Excel.Visible = true;

                    }
                }
            }
            else
            {
                NoRecords.Visible = true;
                Rpt_EmpCources.Visible = false;
                btn_Excel.Visible = false;
            }
        }

        protected void btn_Excel_Click(object sender, EventArgs e)
        {
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", "Employees Courses Report.xls"));
            HttpContext.Current.Response.ContentType = "application/ms-excel";
            HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.Unicode;
            HttpContext.Current.Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());
            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    Table RenderTable = new Table();
                    Table tableSpacing = new Table();
                    TableRow space_tr = new TableRow();

                    #region ISO_tr
                    TableRow ISO_tr = new TableRow();
                    TableCell ISO_tr_cell1 = new TableCell();
                    ISO_tr_cell1.Text = "F-7-3.1   REV-0/2";
                    ISO_tr_cell1.ColumnSpan = 6;
                    ISO_tr.Cells.Add(ISO_tr_cell1);
                    ISO_tr.Font.Bold = true;
                    ISO_tr_cell1.Font.Size = 15;
                    ISO_tr.HorizontalAlign = HorizontalAlign.Right;
                    RenderTable.Rows.Add(ISO_tr);
                    #endregion
                    #region Title_tr
                    TableRow Titl_tr = new TableRow();
                    TableCell Titl_tr_cell1 = new TableCell();
                    Titl_tr_cell1.Text = "Training Request (TNA ) Form";
                    Titl_tr_cell1.ColumnSpan = 5;
                    Titl_tr.Cells.Add(Titl_tr_cell1);
                    Titl_tr.Font.Bold = true;
                    Titl_tr.Font.Underline = true;
                    Titl_tr_cell1.Font.Size = 15;
                    Titl_tr.HorizontalAlign = HorizontalAlign.Center;
                    RenderTable.Rows.Add(Titl_tr);
                    #endregion
                    #region space
                    TableCell space_tr_cell1 = new TableCell();
                    space_tr_cell1.Controls.Add(tableSpacing);
                    space_tr.Cells.Add(space_tr_cell1);
                    RenderTable.Rows.Add(space_tr);
                    #endregion
                    #region repeater
                    foreach (RepeaterItem ri in Rpt_EmpCources.Items)
                    {

                        GridView GV_DepartmenCources = (GridView)ri.FindControl("GV_DepartmenCources");
                        if (GV_DepartmenCources != null && GV_DepartmenCources.Rows.Count > 0)
                        {
                            TableRow tr1 = new TableRow();
                            TableCell tr1_cell = new TableCell();
                            tr1_cell.Text = ((Label)ri.FindControl("lbl_EmpDepartment")).Text;
                            tr1_cell.Font.Bold = true;
                            tr1_cell.Font.Underline = true;
                            tr1_cell.ForeColor = Color.Blue;
                            tr1_cell.Font.Size = 12;
                            tr1.Cells.Add(tr1_cell);

                            TableRow tr2 = new TableRow();

                            Table DepartmenCources_tbl = GridViewExportUtil.ExportGridtoTable(GV_DepartmenCources);
                            TableCell tr2_cell = new TableCell();
                            tr2_cell.Controls.Add(DepartmenCources_tbl);
                            tr2_cell.ColumnSpan = 6;
                            tr2.Cells.Add(tr2_cell);

                            TableRow GridSpace_tr = new TableRow();
                            TableCell space_tr_cell = new TableCell();
                            space_tr_cell.Controls.Add(tableSpacing);
                            GridSpace_tr.Cells.Add(space_tr_cell);



                            RenderTable.Rows.Add(tr1);
                            RenderTable.Rows.Add(tr2);
                            RenderTable.Rows.Add(GridSpace_tr);
                        }
                    }
                    #endregion
                    #region ISO_tr
                    TableRow ISO_tr2 = new TableRow();
                    TableCell ISO_tr_cell2 = new TableCell();
                    ISO_tr_cell2.Text = "F-7-3.1   REV-0/2";
                    ISO_tr_cell2.ColumnSpan = 6;
                    ISO_tr_cell2.Font.Size = 15;

                    ISO_tr2.Cells.Add(ISO_tr_cell2);
                    ISO_tr2.Font.Bold = true;
                    ISO_tr2.HorizontalAlign = HorizontalAlign.Right;
                    RenderTable.Rows.Add(ISO_tr2);
                    #endregion

                    HttpContext.Current.Response.Write("<meta http-equiv=Content-Type content=\"text/html; charset=utf-8\">\n");
                    RenderTable.RenderControl(htw);
                    //  render the htmlwriter into the response  
                    string style = @"<style> td { mso-number-format:\@;} </style>";
                    HttpContext.Current.Response.Write(style);
                    HttpContext.Current.Response.Write(sw.ToString());
                    HttpContext.Current.Response.Write("</body>");
                    HttpContext.Current.Response.Write("</html>");
                    HttpContext.Current.Response.End();
                }
            }

        }

        protected void Rpt_EmpCources_ItemDataBound(object sender, System.Web.UI.WebControls.RepeaterItemEventArgs e)
        {
            try
            {
                GridView GV_DepartmenCources = (GridView)e.Item.FindControl("GV_DepartmenCources");

                string Emp_Department = ((DataRowView)e.Item.DataItem).Row["Department"].ToString();
                DataTable EmployeeCourses_dt = new DataTable();

                string TNAYear = string.Empty;
                if (ddl_TNAYear.SelectedItem.Text != CommonMethods.GetLookups("TNAYear"))
                    TNAYear = ddl_TNAYear.SelectedItem.Text;
                else
                    TNAYear = CommonMethods.GetLookups("TNAYear");

                EmployeeCourses_dt = GetEmployeeCourses(Emp_Department.Trim(), TNAYear);


                DataTable Tem_Data = new DataTable();
                Tem_Data = EmployeeCourses_dt.Clone();

                #region Filter By EmpCode

                if (!string.IsNullOrEmpty(txt_EmpCode.Text))
                {
                    if (EmployeeCourses_dt != null && EmployeeCourses_dt.Rows.Count > 0)
                    {
                        DataRow[] rows = EmployeeCourses_dt.Select("EmpCode='" + txt_EmpCode.Text.Trim() + "'");
                        if (rows != null && rows.Length > 0)
                        {
                            Tem_Data = rows.CopyToDataTable();
                            EmployeeCourses_dt.Clear();
                            EmployeeCourses_dt = Tem_Data.Copy();
                            Tem_Data.Clear();
                        }
                        else
                        {
                            EmployeeCourses_dt.Clear();
                            EmployeeCourses_dt = null;
                        }
                    }
                }

                #endregion

                #region Filter By CourseName

                if (ddl_CourseName.SelectedItem.Text != "-- Select --")
                {
                    if (EmployeeCourses_dt != null && EmployeeCourses_dt.Rows.Count > 0)
                    {
                        DataRow[] rows = EmployeeCourses_dt.Select("CourseName='" + ddl_CourseName.SelectedItem.Text.Trim() + "'");
                        if (rows != null && rows.Length > 0)
                        {
                            Tem_Data = rows.CopyToDataTable();
                            EmployeeCourses_dt.Clear();
                            EmployeeCourses_dt = Tem_Data.Copy();
                            Tem_Data.Clear();
                        }
                        else
                        {
                            EmployeeCourses_dt.Clear();
                            EmployeeCourses_dt = null;
                        }

                    }
                }

                #endregion
                #region Filter By Date From
                if (!cal_From.IsDateEmpty)
                {
                    if (EmployeeCourses_dt != null && EmployeeCourses_dt.Rows.Count > 0)
                    {

                        DataRow[] rows = EmployeeCourses_dt.Select("CreatedDate >= #" + CommonMethods.GetFromatDateString(cal_From.SelectedDate) + " #");
                        if (rows != null && rows.Length > 0)
                        {
                            Tem_Data = rows.CopyToDataTable();
                            EmployeeCourses_dt.Clear();
                            EmployeeCourses_dt = Tem_Data.Copy();
                            Tem_Data.Clear();

                        }
                        else
                        {
                            EmployeeCourses_dt.Clear();
                            EmployeeCourses_dt = null;
                        }
                    }
                }

                #endregion

                #region Filter By Date TO
                if (!cal_To.IsDateEmpty)
                {
                    if (EmployeeCourses_dt != null && EmployeeCourses_dt.Rows.Count > 0)
                    {
                        DataRow[] rows = EmployeeCourses_dt.Select("CreatedDate <= #" + CommonMethods.GetFromatDateString(cal_To.SelectedDate) + " #");
                        if (rows != null && rows.Length > 0)
                        {
                            Tem_Data = rows.CopyToDataTable();
                            EmployeeCourses_dt.Clear();
                            EmployeeCourses_dt = Tem_Data.Copy();
                            Tem_Data.Clear();

                        }
                        else
                        {
                            EmployeeCourses_dt.Clear();
                            EmployeeCourses_dt = null;
                        }

                    }
                }

                #endregion


                if (EmployeeCourses_dt != null && EmployeeCourses_dt.Rows.Count > 0)
                {
                    EmployeeCourses_dt.DefaultView.Sort = "EmpCode Asc";
                    GV_DepartmenCources.DataSource = EmployeeCourses_dt;
                    GV_DepartmenCources.DataBind();
                }
                else
                {
                    GV_DepartmenCources.DataSource = null;
                    GV_DepartmenCources.DataBind();
                    e.Item.Visible = false;
                    EmptyDepartmentCourses += 1;
                }

            }
            catch (Exception ex)
            { CommonMethods.UpdateTestLst(ex.ToString()); }

        }

        protected void ddl_CoursesCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddl_CoursesCategory.SelectedValue != "-1")
            {
                DataTable CategoryCourses_DT = new DataTable();
                if (ddl_CoursesCategory.SelectedValue == "0")
                {
                    CategoryCourses_DT = Get_BasicSkills_Courses();
                }
                else if (ddl_CoursesCategory.SelectedValue == "1")
                {
                    CategoryCourses_DT = Get_DevelopmentSkills_Courses();
                }
                else if (ddl_CoursesCategory.SelectedValue == "2")
                {
                    CategoryCourses_DT = Get_TechnicalSkills_Courses();
                }


                CategoryCourses_DT = CommonMethods.RemoveDuplicates(CategoryCourses_DT, "CourseName");
                ddl_CourseName.DataSource = CategoryCourses_DT;
                ddl_CourseName.DataTextField = "CourseName";
                ddl_CourseName.DataBind();
                ddl_CourseName.Items.Insert(0, new ListItem("-- Select --", "-1"));
                ddl_CourseName.Enabled = true;
            }
            else
            {
                ddl_CourseName.Enabled = false;
                ddl_CourseName.DataSource = null;
                ddl_CourseName.DataBind();
                ddl_CourseName.Items.Insert(0, new ListItem("-- Select --", "-1"));
            }
        }

        protected void cal_From_DateChanged(object sender, EventArgs e)
        {
            CheckTNAYearSearch();
        }



        protected void cal_To_DateChanged(object sender, EventArgs e)
        {
            CheckTNAYearSearch();
        }

        protected void btn_Save_Click(object sender, EventArgs e)
        {
            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                Guid SiteId = site.ID;
                Guid WebId = web.ID;

                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {
                        //if (GV_ReqsItems.Rows.Count > 0)
                        //{
                        //    bool isUpdated = false;
                        //    for (int i = 0; i < GV_ReqsItems.Rows.Count; i++)
                        //    {
                        //        string status = ((DropDownList)GV_ReqsItems.Rows[i].FindControl("drpLstItemStatus")).SelectedValue;

                        //        string notes = string.Empty;
                        //        if (((TextBox)GV_ReqsItems.Rows[i].FindControl("txtNotes")) != null && !string.IsNullOrEmpty(((TextBox)GV_ReqsItems.Rows[i].FindControl("txtNotes")).Text))
                        //        {
                        //            notes = ((TextBox)GV_ReqsItems.Rows[i].FindControl("txtNotes")).Text;
                        //        }
                        //    }
                        //}
                    });
            }
        }
    }
}

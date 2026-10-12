using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using NileSugar.HrTNA.Webparts.TrainingRequests_WP;

namespace NileSugar.HrTNA.Webparts.TNAReport
{
    [ToolboxItemAttribute(false)]
    public partial class TNAReport : WebPart
    {
        // Uncomment the following SecurityPermission attribute only when doing Performance Profiling on a farm solution
        // using the Instrumentation method, and then remove the SecurityPermission attribute when the code is ready
        // for production. Because the SecurityPermission attribute bypasses the security check for callers of
        // your constructor, it's not recommended for production purposes.
        // [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Assert, UnmanagedCode = true)]
        public TNAReport()
        {
        }

        #region Function
        private DataTable GetEmpCourses()
        {
            DataTable EmpCourses_DT = new DataTable();
            try
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                using (SPWeb web = site.OpenWeb())
                {

                    SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                    if (EmpCoursesLst != null)
                    {
                        SPQuery qry1 = new SPQuery();
                        qry1.Query = string.Format(@" <Where>
                                                        <Eq>
                                                            <FieldRef Name='TNAYear' />
                                                            <Value Type='Text'>{0}</Value>
                                                        </Eq>
                                                    </Where>", DateTime.Now.Year);
                        SPListItemCollection coll1 = EmpCoursesLst.GetItems(qry1);
                        DataTable QryResult_DT = coll1.GetDataTable();
                        if (QryResult_DT != null && QryResult_DT.Rows.Count > 0)
                        {
                            EmpCourses_DT = SearchCases(QryResult_DT);
                        }
                    }
                }

            }
            catch { }
            return EmpCourses_DT;

        }

        private DataTable SearchCases(DataTable QryResult_DT)
        {
            DataTable EmployeeCources = new DataTable();
            EmployeeCources = QryResult_DT.Clone();
            #region Filter By Devision
            if (ddl_Devision.SelectedValue != "-1" && ddl_Devision.SelectedValue != "")
            {
                DataRow[] rows = QryResult_DT.Select("Division='" + ddl_Devision.SelectedItem.Text.Trim() + "'");
                foreach (DataRow row in rows)
                {
                    EmployeeCources.ImportRow(row);
                }
                QryResult_DT.Rows.Clear();
                QryResult_DT = EmployeeCources.Copy();
                EmployeeCources.Rows.Clear();
            }
            #endregion

            #region Filter By Department
            if (ddl_Department.SelectedValue != "-1" && ddl_Department.SelectedValue != "")
            {
                DataRow[] rows = QryResult_DT.Select("Department='" + ddl_Department.SelectedItem.Text.Trim() + "'");
                foreach (DataRow row in rows)
                {
                    EmployeeCources.ImportRow(row);
                }
                QryResult_DT.Rows.Clear();
                QryResult_DT = EmployeeCources.Copy();
                EmployeeCources.Rows.Clear();
            }
            #endregion

            #region Filter By Employee NAme
            if (ddl_EmpNameSearch.SelectedValue != "-1" && ddl_EmpNameSearch.SelectedValue != "")
            {
                DataRow[] rows = QryResult_DT.Select("EmpName='" + ddl_EmpNameSearch.SelectedItem.Text.Trim() + "'");
                foreach (DataRow row in rows)
                {
                    EmployeeCources.ImportRow(row);
                }
                QryResult_DT.Rows.Clear();
                QryResult_DT = EmployeeCources.Copy();
                EmployeeCources.Rows.Clear();
            }
            #endregion

            #region Filter By Employee code
            if (txt_EmpCodeSearch.Text != "")
            {
                DataRow[] rows = QryResult_DT.Select("EmpCode='" + txt_EmpCodeSearch.Text + "'");
                foreach (DataRow row in rows)
                {
                    EmployeeCources.ImportRow(row);
                }
                QryResult_DT.Rows.Clear();
                QryResult_DT = EmployeeCources.Copy();
                EmployeeCources.Rows.Clear();
            }
            #endregion

            #region Filter by Request Status
            if (ddl_RequestStatus.SelectedValue != "-1" && ddl_RequestStatus.SelectedValue != "")
            {
                DataRow[] rows = QryResult_DT.Select("RequestStatus='" + ddl_RequestStatus.SelectedItem.Text.Trim() + "'");
                foreach (DataRow row in rows)
                {
                    EmployeeCources.ImportRow(row);
                }
                QryResult_DT.Rows.Clear();
                QryResult_DT = EmployeeCources.Copy();
                EmployeeCources.Rows.Clear();
            }
            #endregion

            return QryResult_DT;
        }

        private DataTable Get_BasicSkills_Courses(string empCode)
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
                            @" 
                                  <Where>
                                      <Eq>
                                         <FieldRef Name='EmpCode' />
                                         <Value Type='Text'>{0}</Value>
                                      </Eq>
                                   </Where>
                                       ", empCode);

                        SPListItemCollection collItems = EmpBasicSkillsOtherCoursesLst.GetItems(qry);
                        if (collItems.Count > 0)
                        {
                            EmpCoursesBasic_DT = collItems.GetDataTable();
                        }
                        else
                        {
                            EmpCoursesBasic_DT.Columns.Add("CourseName", typeof(String));
                            //EmpCoursesBasic_DT.Columns.Add("ExpectedDate", typeof(DateTime)); 
                            EmpCoursesBasic_DT.Columns.Add("ExpectedQuarter", typeof(String));
                            EmpCoursesBasic_DT.Columns.Add("Reason", typeof(String));
                        }
                    }

                    SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                    if (EmpCoursesLst != null)
                    {
                        SPQuery qry = new SPQuery();
                        qry.Query = string.Format(@"  <Where>
                                                          <And>
                                                             <Eq>
                                                                <FieldRef Name='TNAYear' />
                                                                <Value Type='Text'>{0}</Value>
                                                             </Eq>
                                                             <Eq>
                                                                <FieldRef Name='EmpCode' />
                                                                <Value Type='Text'>{1}</Value>
                                                             </Eq>
                                                          </And>
                                                       </Where>", DateTime.Now.Year, empCode);
                        SPListItemCollection coll = EmpCoursesLst.GetItems(qry);
                        DataTable EmpCourses_DT = coll.GetDataTable();

                        if (EmpCourses_DT.Rows[0]["BS_EngC_Name"] != null && EmpCourses_DT.Rows[0]["BS_EngC_Name"].ToString() != "")
                        {
                            DataRow dr1 = EmpCoursesBasic_DT.NewRow();
                            dr1["CourseName"] = EmpCourses_DT.Rows[0]["BS_EngC_Name"];
                            dr1["ExpectedQuarter"] = EmpCourses_DT.Rows[0]["BS_EngC_ExpectedQuarter"];
                            dr1["Reason"] = EmpCourses_DT.Rows[0]["BS_EngC_Reason"];
                            EmpCoursesBasic_DT.Rows.Add(dr1);
                        }

                        if (EmpCourses_DT.Rows[0]["BS_CompSkillsC_Name"] != null && EmpCourses_DT.Rows[0]["BS_CompSkillsC_Name"].ToString() != "")
                        {
                            DataRow dr2 = EmpCoursesBasic_DT.NewRow();
                            dr2["CourseName"] = EmpCourses_DT.Rows[0]["BS_CompSkillsC_Name"];
                            dr2["ExpectedQuarter"] = EmpCourses_DT.Rows[0]["BS_CompSkillsC_ExpectedQuarter"];
                            dr2["Reason"] = EmpCourses_DT.Rows[0]["BS_CompSkillsC_Reason"];
                            EmpCoursesBasic_DT.Rows.Add(dr2);
                        }

                        if (EmpCourses_DT.Rows[0]["BS_BEPC_Name"] != null && EmpCourses_DT.Rows[0]["BS_BEPC_Name"].ToString() != "")
                        {
                            DataRow dr3 = EmpCoursesBasic_DT.NewRow();
                            dr3["CourseName"] = EmpCourses_DT.Rows[0]["BS_BEPC_Name"];
                            dr3["ExpectedQuarter"] = EmpCourses_DT.Rows[0]["BS_BEPC_ExpectedQuarter"];
                            dr3["Reason"] = EmpCourses_DT.Rows[0]["BS_BEPC_Reason"];
                            EmpCoursesBasic_DT.Rows.Add(dr3);
                        }
                    }
                }

            }
            catch { }
            return EmpCoursesBasic_DT;

        }
        private DataTable Get_DevelopmentSkills_Courses(string empCode)
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
                                  <Where>
                                      <Eq>
                                         <FieldRef Name='EmpCode' />
                                         <Value Type='Text'>{0}</Value>
                                      </Eq>
                                   </Where>
                                       ", empCode);

                        SPListItemCollection collItems = EmpDevSkillsOtherCoursesLst.GetItems(qry);
                        if (collItems.Count > 0)
                        {
                            EmpCoursesDev_DT = collItems.GetDataTable();
                        }
                        else
                        {
                            EmpCoursesDev_DT.Columns.Add("CourseName", typeof(String));
                            // EmpCoursesDev_DT.Columns.Add("ExpectedDate", typeof(DateTime));
                            EmpCoursesDev_DT.Columns.Add("ExpectedQuarter", typeof(String));
                            EmpCoursesDev_DT.Columns.Add("Reason", typeof(String));
                        }
                    }

                    SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                    if (EmpCoursesLst != null)
                    {
                        SPQuery qry = new SPQuery();
                        qry.Query = string.Format(@"  <Where>
                                                          <And>
                                                             <Eq>
                                                                <FieldRef Name='TNAYear' />
                                                                <Value Type='Text'>{0}</Value>
                                                             </Eq>
                                                             <Eq>
                                                                <FieldRef Name='EmpCode' />
                                                                <Value Type='Text'>{1}</Value>
                                                             </Eq>
                                                          </And>
                                                       </Where>", DateTime.Now.Year, empCode);
                        SPListItemCollection coll = EmpCoursesLst.GetItems(qry);
                        DataTable EmpCourses_DT = coll.GetDataTable();

                        if (EmpCourses_DT.Rows[0]["DevS_CommSC_Name"] != null && EmpCourses_DT.Rows[0]["DevS_CommSC_Name"].ToString() != "")
                        {
                            DataRow dr1 = EmpCoursesDev_DT.NewRow();
                            dr1["CourseName"] = EmpCourses_DT.Rows[0]["DevS_CommSC_Name"];
                            dr1["ExpectedQuarter"] = EmpCourses_DT.Rows[0]["DevS_CommSC_ExpectedQuarter"];
                            dr1["Reason"] = EmpCourses_DT.Rows[0]["DevS_CommSC_Reason"];
                            EmpCoursesDev_DT.Rows.Add(dr1);
                        }

                        if (EmpCourses_DT.Rows[0]["DevS_InterPersSC_Name"] != null && EmpCourses_DT.Rows[0]["DevS_InterPersSC_Name"].ToString() != "")
                        {

                            DataRow dr2 = EmpCoursesDev_DT.NewRow();
                            dr2["CourseName"] = EmpCourses_DT.Rows[0]["DevS_InterPersSC_Name"];
                            dr2["ExpectedQuarter"] = EmpCourses_DT.Rows[0]["DevS_InterPersSC_ExpectedQuarter"];
                            dr2["Reason"] = EmpCourses_DT.Rows[0]["DevS_InterPersSC_Reason"];
                            EmpCoursesDev_DT.Rows.Add(dr2);
                        }

                        if (EmpCourses_DT.Rows[0]["DevS_BIC_Name"] != null && EmpCourses_DT.Rows[0]["DevS_BIC_Name"].ToString() != "")
                        {

                            DataRow dr3 = EmpCoursesDev_DT.NewRow();
                            dr3["CourseName"] = EmpCourses_DT.Rows[0]["DevS_BIC_Name"];
                            dr3["ExpectedQuarter"] = EmpCourses_DT.Rows[0]["DevS_BIC_ExpectedDate"];
                            dr3["Reason"] = EmpCourses_DT.Rows[0]["DevS_BIC_Reason"];
                            EmpCoursesDev_DT.Rows.Add(dr3);
                        }



                    }
                }

            }
            catch { }
            return EmpCoursesDev_DT;


        }

        private DataTable Get_TechnicalSkills_Courses(string empCode, string EmployeeDepartment)
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

                        qry.Query = String.Format(
                            @" 
                                  <Where>
                                          <And>
                                             <Eq>
                                                <FieldRef Name='DeptName' />
                                                <Value Type='Text'>{0}</Value>
                                             </Eq>
                                             <Eq>
                                                <FieldRef Name='EmpCode' />
                                                <Value Type='Text'>{1}</Value>
                                             </Eq>
                                          </And>
                                       </Where>
                                       ", EmployeeDepartment, empCode);

                        SPListItemCollection collItems = EmpTechnicalCoursesLst.GetItems(qry);
                        if (collItems.Count > 0)
                        {
                            EmpCoursesBasic_DT = collItems.GetDataTable();
                        }
                    }


                }

            }

            catch { }
            return EmpCoursesBasic_DT;


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
                            //EmpCoursesBasic_DT.Columns.Add("ExpectedDate", typeof(DateTime));
                            EmpCoursesBasic_DT.Columns.Add("ExpectedQuarter", typeof(String));
                            EmpCoursesBasic_DT.Columns.Add("Reason", typeof(String));
                        }
                    }

                    SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                    if (EmpCoursesLst != null)
                    {
                        SPQuery qry = new SPQuery();
                        qry.Query = string.Format(@"    <Where>
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
                            dr1["Reason"] = EmpCourses_DT.Rows[i]["BS_EngC_Reason"];

                            DataRow dr2 = EmpCoursesBasic_DT.NewRow();
                            dr2["CourseName"] = EmpCourses_DT.Rows[i]["BS_CompSkillsC_Name"];
                            dr2["ExpectedQuarter"] = EmpCourses_DT.Rows[i]["BS_CompSkillsC_ExpectedQuarter"];
                            dr2["Reason"] = EmpCourses_DT.Rows[i]["BS_CompSkillsC_Reason"];

                            DataRow dr3 = EmpCoursesBasic_DT.NewRow();
                            dr3["CourseName"] = EmpCourses_DT.Rows[i]["BS_BEPC_Name"];
                            dr3["ExpectedQuarter"] = EmpCourses_DT.Rows[i]["BS_BEPC_ExpectedQuarter"];
                            dr3["Reason"] = EmpCourses_DT.Rows[i]["BS_BEPC_Reason"];

                            if (dr1 != null && dr1["CourseName"].ToString() != "" && dr1["CourseName"].ToString() != null)
                                EmpCoursesBasic_DT.Rows.Add(dr1);

                            if (dr2 != null && dr2["CourseName"].ToString() != "" && dr2["CourseName"].ToString() != null)
                                EmpCoursesBasic_DT.Rows.Add(dr2);

                            if (dr3 != null && dr3["CourseName"].ToString() != "" && dr3["CourseName"].ToString() != null)
                                EmpCoursesBasic_DT.Rows.Add(dr3);
                        }

                    }
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
                            //EmpCoursesDev_DT.Columns.Add("ExpectedDate", typeof(DateTime));
                            EmpCoursesDev_DT.Columns.Add("ExpectedQuarter", typeof(String));
                            EmpCoursesDev_DT.Columns.Add("Reason", typeof(String));
                        }
                    }

                    SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                    if (EmpCoursesLst != null)
                    {
                        SPQuery qry = new SPQuery();
                        qry.Query = string.Format(@"    <Where>
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
                            dr1["Reason"] = EmpCourses_DT.Rows[i]["DevS_CommSC_Reason"];

                            DataRow dr2 = EmpCoursesDev_DT.NewRow();
                            dr2["CourseName"] = EmpCourses_DT.Rows[i]["DevS_InterPersSC_Name"];
                            dr2["ExpectedQuarter"] = EmpCourses_DT.Rows[i]["DevS_InterPersSC_ExpectedQuarter"];
                            dr2["Reason"] = EmpCourses_DT.Rows[i]["DevS_InterPersSC_Reason"];

                            DataRow dr3 = EmpCoursesDev_DT.NewRow();
                            dr3["CourseName"] = EmpCourses_DT.Rows[i]["DevS_BIC_Name"];
                            dr3["ExpectedQuarter"] = EmpCourses_DT.Rows[i]["DevS_BIC_ExpectedDate"];
                            dr3["Reason"] = EmpCourses_DT.Rows[i]["DevS_BIC_Reason"];

                            if (dr1 != null && dr1["CourseName"].ToString() != "" && dr1["CourseName"].ToString() != null)
                                EmpCoursesDev_DT.Rows.Add(dr1);

                            if (dr2 != null && dr2["CourseName"].ToString() != "" && dr2["CourseName"].ToString() != null)
                                EmpCoursesDev_DT.Rows.Add(dr2);

                            if (dr3 != null && dr3["CourseName"].ToString() != "" && dr3["CourseName"].ToString() != null)
                                EmpCoursesDev_DT.Rows.Add(dr3);
                        }
                    }
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
                            EmpCoursesBasic_DT = collItems.GetDataTable();
                        }
                    }


                }

            }

            catch { }
            return EmpCoursesBasic_DT;


        }


        private DataTable Get_EmployeeData(string EmployeeCode)
        {
            DataTable Emp_Data_DT = new DataTable();
            Emp_Data_DT.Columns.Add("Position", typeof(String));
            Emp_Data_DT.Columns.Add("Division", typeof(String));
            Emp_Data_DT.Columns.Add("Department", typeof(String));
            Emp_Data_DT.Columns.Add("Section", typeof(String));
            Emp_Data_DT.Columns.Add("RequestStatus", typeof(String));

            try
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                using (SPWeb web = site.OpenWeb())
                {
                    SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
                    if (EmpCoursesLst != null)
                    {
                        SPQuery qry = new SPQuery();
                        qry.Query = string.Format(@"  <Where>
                                                          <And>
                                                             <Eq>
                                                                <FieldRef Name='TNAYear' />
                                                                <Value Type='Text'>{0}</Value>
                                                             </Eq>
                                                             <Eq>
                                                                <FieldRef Name='EmpCode' />
                                                                <Value Type='Text'>{1}</Value>
                                                             </Eq>
                                                          </And>
                                                       </Where>", DateTime.Now.Year, EmployeeCode);
                        SPListItemCollection coll = EmpCoursesLst.GetItems(qry);
                        DataTable Emp_Dt = coll.GetDataTable();
                        if (Emp_Dt != null && Emp_Dt.Rows.Count > 0)
                        {
                            DataRow dr = Emp_Data_DT.NewRow();
                            dr["Position"] = Emp_Dt.Rows[0]["Position"];
                            dr["Division"] = Emp_Dt.Rows[0]["Division"];
                            dr["Department"] = Emp_Dt.Rows[0]["Department"];
                            dr["Section"] = Emp_Dt.Rows[0]["Section"];
                            dr["RequestStatus"] = Emp_Dt.Rows[0]["RequestStatus"];
                            Emp_Data_DT.Rows.Add(dr);
                        }

                    }
                }
            }
            catch { }
            return Emp_Data_DT;
        }
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
        private void LoadDevisions()
        {
            try
            {

                DataTable Devision_dt = CommonMethods.GetAllDevisions();
                Devision_dt.DefaultView.Sort = "Division ASC";
                ddl_Devision.DataSource = Devision_dt;
                ddl_Devision.DataTextField = "Division";
                ddl_Devision.DataValueField = "Division";
                ddl_Devision.DataBind();
                ddl_Devision.Items.Insert(0, new ListItem("-- Select --", "-1"));
            }
            catch { }
        }
        private void LoadEmployees()
        {
            try
            {

                DataTable Employees_dt = CommonMethods.GetAllEmployees();
                Employees_dt.DefaultView.Sort = "FULL_NAME ASC";
                ddl_EmpNameSearch.DataSource = Employees_dt;
                ddl_EmpNameSearch.DataTextField = "FULL_NAME";
                ddl_EmpNameSearch.DataValueField = "FULL_NAME";
                ddl_EmpNameSearch.DataBind();
                ddl_EmpNameSearch.Items.Insert(0, new ListItem("-- Select --", "-1"));
            }
            catch { }
        }
        #endregion

        protected override void OnInit(EventArgs e)
        {
            base.OnInit(e);
            InitializeControl();
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                LoadDepartment();
                LoadEmployees();
                LoadDevisions();
                ddl_CourseName.Items.Insert(0, new ListItem("-- Select --", "-1"));
            }
        }

        protected void ddl_Devision_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable Devision_Departments = CommonMethods.Get_Devision_Departments(ddl_Devision.SelectedItem.Text);
            Devision_Departments.DefaultView.Sort = "Department ASC";
            ddl_Department.DataSource = Devision_Departments;
            ddl_Department.DataTextField = "Department";
            ddl_Department.DataBind();
            ddl_Department.Items.Insert(0, new ListItem("-- Select --", "-1"));

            DataTable Devision_Employees = CommonMethods.Get_Devision_Employees(ddl_Devision.SelectedItem.Text);
            Devision_Employees.DefaultView.Sort = "FULL_NAME ASC";
            ddl_EmpNameSearch.DataSource = Devision_Employees;
            ddl_EmpNameSearch.DataTextField = "FULL_NAME";
            ddl_EmpNameSearch.DataBind();
            ddl_EmpNameSearch.Items.Insert(0, new ListItem("-- Select --", "-1"));
        }

        protected void ddl_Department_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable DeptEmployees = CommonMethods.GetDepartmentEmployees(ddl_Department.SelectedItem.Text);
            DeptEmployees.DefaultView.Sort = "FULL_NAME ASC";
            ddl_EmpNameSearch.DataSource = DeptEmployees;
            ddl_EmpNameSearch.DataTextField = "FULL_NAME";
            ddl_EmpNameSearch.DataBind();
            ddl_EmpNameSearch.Items.Insert(0, new ListItem("-- Select --", "-1"));
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
            }

        }

        protected void btn_LoadData_Click(object sender, EventArgs e)
        {

            DataTable distinctValues = new DataTable();
            DataTable ALLEmployeeCources = GetEmpCourses();
            if (ALLEmployeeCources != null && ALLEmployeeCources.Rows.Count > 0)
                distinctValues = CommonMethods.RemoveDuplicates(ALLEmployeeCources, "EmpCode");

            if (distinctValues != null && distinctValues.Rows.Count > 0)
            {
                NoRecords.Visible = false;
                Rpt_EmpCources.Visible = true;
                btn_Exceel.Visible = true;
                Rpt_EmpCources.DataSource = distinctValues;
                Rpt_EmpCources.DataBind();

            }
            else
            {
                NoRecords.Visible = true;
                Rpt_EmpCources.Visible = false;
                btn_Exceel.Visible = false;
            }
        }

        protected void btn_Exceel_Click(object sender, EventArgs e)
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

                    TableRow ISO_tr = new TableRow();
                    TableRow Titl_tr = new TableRow();
                    TableRow space_tr = new TableRow();

                    #region ISO_tr
                    TableCell ISO_tr_cell1 = new TableCell();
                    ISO_tr_cell1.Text = "F-7-3.1   REV-0/2";
                    ISO_tr_cell1.ColumnSpan = 5;
                    ISO_tr.Cells.Add(ISO_tr_cell1);
                    ISO_tr.Font.Bold = true;
                    ISO_tr_cell1.Font.Size = 15;
                    ISO_tr.HorizontalAlign = HorizontalAlign.Right;
                    RenderTable.Rows.Add(ISO_tr);
                    #endregion

                    #region Title_tr
                    TableCell Titl_tr_cell1 = new TableCell();
                    Titl_tr_cell1.Text = "Employees Courses Report";
                    Titl_tr_cell1.ColumnSpan = 5;
                    Titl_tr.Cells.Add(Titl_tr_cell1);
                    Titl_tr.Font.Bold = true;
                    Titl_tr_cell1.Font.Size = 15;
                    Titl_tr.HorizontalAlign = HorizontalAlign.Center;
                    RenderTable.Rows.Add(Titl_tr);
                    #endregion

                    #region
                    TableCell space_tr_cell1 = new TableCell();
                    space_tr_cell1.Controls.Add(tableSpacing);
                    space_tr.Cells.Add(space_tr_cell1);
                    RenderTable.Rows.Add(space_tr);
                    #endregion

                    foreach (RepeaterItem ri in Rpt_EmpCources.Items)
                    {

                        #region tr1
                        TableRow tr1 = new TableRow();
                        TableCell tr1_cell1 = new TableCell();
                        tr1_cell1.Text = "Courses for Employee: ";
                        tr1_cell1.Font.Bold = true;


                        TableCell tr1_cell2 = new TableCell();
                        tr1_cell2.Text = ((Label)ri.FindControl("lbl_EmpNameValue")).Text;

                        TableCell tr1_cell5 = new TableCell();
                        tr1_cell5.Text = "";

                        TableCell tr1_cell3 = new TableCell();
                        tr1_cell3.Text = "Employee Code: ";
                        tr1_cell3.Font.Bold = true;

                        TableCell tr1_cell4 = new TableCell();
                        tr1_cell4.Text = ((Label)ri.FindControl("lbl_EmpCodeValue")).Text;





                        tr1.Cells.Add(tr1_cell1);
                        tr1.Cells.Add(tr1_cell2);
                        tr1.Cells.Add(tr1_cell3);
                        tr1.Cells.Add(tr1_cell4);
                        tr1.Cells.Add(tr1_cell5);

                        RenderTable.Rows.Add(tr1);
                        #endregion

                        #region tr2
                        TableRow tr2 = new TableRow();

                        GridView Gv_EmpData = (GridView)ri.FindControl("Gv_EmpData");
                        Table EmpData_tbl = GridViewExportUtil.ExportGridtoTable(Gv_EmpData);
                        TableCell tr2_cell1 = new TableCell();
                        tr2_cell1.Controls.Add(EmpData_tbl);
                        tr2_cell1.ColumnSpan = 5;
                        tr2.Cells.Add(tr2_cell1);
                        RenderTable.Rows.Add(tr2);
                        #endregion

                        #region tr3
                        GridView Gv_EmpCources_BasicSkills = (GridView)ri.FindControl("Gv_EmpCources_BasicSkills");
                        if (Gv_EmpCources_BasicSkills.Rows.Count > 0)
                        {
                            TableRow tr3_Title = new TableRow();
                            TableCell tr3_Title_cell1 = new TableCell();
                            tr3_Title_cell1.Text = "Basic Skills";
                            tr3_Title.Cells.Add(tr3_Title_cell1);
                            tr3_Title.Font.Size = 9;
                            tr3_Title.ForeColor = Color.Blue;
                            tr3_Title.Font.Bold = true;
                            tr3_Title.HorizontalAlign = HorizontalAlign.Left;
                            RenderTable.Rows.Add(tr3_Title);


                            TableRow tr3 = new TableRow();
                            TableCell tr3_cell1 = new TableCell();

                            Table EmpCources_BasicSkills_tbl = GridViewExportUtil.ExportGridtoTable(Gv_EmpCources_BasicSkills, "templatField");
                            EmpCources_BasicSkills_tbl.HideColumn(0);
                            EmpCources_BasicSkills_tbl.HorizontalAlign = HorizontalAlign.Center;
                            tr3_cell1.Controls.Add(EmpCources_BasicSkills_tbl);
                            tr3_cell1.ColumnSpan = 5;
                            tr3_cell1.HorizontalAlign = HorizontalAlign.Center;
                            tr3.Cells.Add(tr3_cell1);
                            RenderTable.Rows.Add(tr3);
                        }
                        #endregion

                        #region tr4
                        GridView Gv_EmpCources_DevelopmentSkills = (GridView)ri.FindControl("Gv_EmpCources_DevelopmentSkills");
                        if (Gv_EmpCources_DevelopmentSkills.Rows.Count > 0)
                        {

                            TableRow tr4_Title = new TableRow();
                            TableCell tr4_Title_cell1 = new TableCell();
                            tr4_Title_cell1.Text = "Development Skills";
                            tr4_Title.Cells.Add(tr4_Title_cell1);
                            tr4_Title.Font.Size = 9;
                            tr4_Title.ForeColor = Color.Blue;
                            tr4_Title.Font.Bold = true;
                            tr4_Title.HorizontalAlign = HorizontalAlign.Left;
                            RenderTable.Rows.Add(tr4_Title);



                            TableRow tr4 = new TableRow();
                            TableCell tr4_cell1 = new TableCell();

                            Table EmpCources_DevelopmentSkills_tbl = GridViewExportUtil.ExportGridtoTable(Gv_EmpCources_DevelopmentSkills, "templatField");
                            EmpCources_DevelopmentSkills_tbl.HideColumn(0);

                            tr4_cell1.Controls.Add(EmpCources_DevelopmentSkills_tbl);
                            tr4_cell1.ColumnSpan = 5;
                            tr4_cell1.HorizontalAlign = HorizontalAlign.Center;
                            tr4.Cells.Add(tr4_cell1);
                            RenderTable.Rows.Add(tr4);
                        }
                        #endregion


                        #region tr5
                        GridView Gv_EmpCources_TechnicalSkills = (GridView)ri.FindControl("Gv_EmpCources_TechnicalSkills");
                        if (Gv_EmpCources_TechnicalSkills.Rows.Count > 0)
                        {
                            TableRow tr5_Title = new TableRow();
                            TableCell tr5_Title_cell1 = new TableCell();
                            tr5_Title_cell1.Text = "Technical Skills";
                            tr5_Title.Cells.Add(tr5_Title_cell1);
                            tr5_Title.Font.Size = 9;
                            tr5_Title.ForeColor = Color.Blue;
                            tr5_Title.Font.Bold = true;
                            tr5_Title.HorizontalAlign = HorizontalAlign.Left;
                            tr5_Title.VerticalAlign = VerticalAlign.Middle;
                            RenderTable.Rows.Add(tr5_Title);

                            TableRow tr5 = new TableRow();
                            TableCell tr5_cell1 = new TableCell();

                            Table EmpCources_TechnicalSkills_tbl = GridViewExportUtil.ExportGridtoTable(Gv_EmpCources_TechnicalSkills, "templatField");
                            EmpCources_TechnicalSkills_tbl.HideColumn(0);
                            tr5_cell1.Controls.Add(EmpCources_TechnicalSkills_tbl);
                            tr5_cell1.ColumnSpan = 5;
                            tr5_cell1.HorizontalAlign = HorizontalAlign.Center;
                            tr5.Cells.Add(tr5_cell1);
                            RenderTable.Rows.Add(tr5);
                        }
                        #endregion


                        #region
                        TableRow trSpace = new TableRow();
                        TableCell trSpace_cell = new TableCell();
                        trSpace_cell.Controls.Add(tableSpacing);

                        trSpace.Cells.Add(trSpace_cell);
                        RenderTable.Rows.Add(trSpace);



                        TableRow trLine = new TableRow();
                        TableCell trSLine_cell = new TableCell();
                        trSLine_cell.Controls.Add(new LiteralControl("<br />"));

                        trLine.Cells.Add(trSLine_cell);
                        RenderTable.Rows.Add(trLine);


                        //Pen redPen = new Pen(Color.Red, 1);
                        //float x1 = 20.0F, y1 = 20.0F;
                        //float x2 = 200.0F, y2 = 20.0F;

                        //TableRow trLine = new TableRow();
                        //TableCell trLine_cell = new TableCell();
                        //trLine_cell.Controls.Add(Graphics.DrawLine(redPen, x1, y1, x2, y2));

                        //trLine.Cells.Add(trSpace_cell);
                        //RenderTable.Rows.Add(trLine);

                        #endregion
                    }
                    HttpContext.Current.Response.Write("<meta http-equiv=Content-Type content=\"text/html; charset=utf-8\">\n");
                    RenderTable.RenderControl(htw);
                    //  render the htmlwriter into the response  
                    HttpContext.Current.Response.Write(sw.ToString());
                    HttpContext.Current.Response.Write("</body>");
                    HttpContext.Current.Response.Write("</html>");
                    HttpContext.Current.Response.End();
                }
            }
        }

        protected void Rpt_EmpCources_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            GridView Gv_EmpData = (GridView)e.Item.FindControl("Gv_EmpData");
            GridView Gv_EmpCources_BasicSkills = (GridView)e.Item.FindControl("Gv_EmpCources_BasicSkills");
            GridView Gv_EmpCources_DevelopmentSkills = (GridView)e.Item.FindControl("Gv_EmpCources_DevelopmentSkills");
            GridView Gv_EmpCources_TechnicalSkills = (GridView)e.Item.FindControl("Gv_EmpCources_TechnicalSkills");

            string EmployeeCode = ((DataRowView)e.Item.DataItem).Row["EmpCode"].ToString();


            DataTable EmployeeData = Get_EmployeeData(EmployeeCode);
            string EmployeeDepartment = EmployeeData.Rows[0]["Department"].ToString();
            Gv_EmpData.DataSource = EmployeeData;
            Gv_EmpData.DataBind();


            DataTable BasicSkills_DT = Get_BasicSkills_Courses(EmployeeCode);
            DataTable DevelopmentSkills_DT = Get_DevelopmentSkills_Courses(EmployeeCode);
            DataTable TechnicalSkills_DT = Get_TechnicalSkills_Courses(EmployeeCode, EmployeeDepartment);

            #region Filter By Course Name And expected Date


            if (drp_ExpectedQuarter.SelectedValue != "-1" && ddl_CourseName.SelectedValue != "-1")
            {
                // string DateExpected = CommonMethods.GetFromatDateString(DateTimePicker_ExpectedDate.SelectedDate.ToShortDateString());

                string ExpectedQuarter = drp_ExpectedQuarter.SelectedValue;
                if (BasicSkills_DT != null && BasicSkills_DT.Rows.Count > 0)
                {
                    DataRow[] rows1 = BasicSkills_DT.Select("CourseName='" + ddl_CourseName.SelectedItem.Text.Trim() + "' ANd ExpectedQuarter=" + ExpectedQuarter);
                    DataTable BasicSkillsFilter_DT = BasicSkills_DT.Clone();

                    foreach (DataRow row in rows1)
                    {
                        BasicSkillsFilter_DT.ImportRow(row);
                    }
                    Gv_EmpCources_BasicSkills.DataSource = BasicSkillsFilter_DT;
                    Gv_EmpCources_BasicSkills.DataBind();
                }




                if (DevelopmentSkills_DT != null && DevelopmentSkills_DT.Rows.Count > 0)
                {
                    DataRow[] rows2 = DevelopmentSkills_DT.Select("CourseName='" + ddl_CourseName.SelectedItem.Text.Trim() + "' ANd ExpectedDate=" + ExpectedQuarter);
                    DataTable DevelopmentSkillsFilter_DT = DevelopmentSkills_DT.Clone();

                    foreach (DataRow row in rows2)
                    {
                        DevelopmentSkillsFilter_DT.ImportRow(row);
                    }
                    Gv_EmpCources_DevelopmentSkills.DataSource = DevelopmentSkillsFilter_DT;
                    Gv_EmpCources_DevelopmentSkills.DataBind();
                }



                if (TechnicalSkills_DT != null && TechnicalSkills_DT.Rows.Count > 0)
                {
                    DataRow[] rows3 = TechnicalSkills_DT.Select("CourseName='" + ddl_CourseName.SelectedItem.Text.Trim() + "' ANd ExpectedDate=" + ExpectedQuarter);
                    DataTable TechnicalSkillsFilter_DT = TechnicalSkills_DT.Clone();

                    foreach (DataRow row in rows3)
                    {
                        TechnicalSkillsFilter_DT.ImportRow(row);
                    }
                    Gv_EmpCources_TechnicalSkills.DataSource = TechnicalSkillsFilter_DT;
                    Gv_EmpCources_TechnicalSkills.DataBind();
                }
            }
            #endregion

            #region  expected Date

            else if (drp_ExpectedQuarter.SelectedValue != "-1")
            {
                //string DateExpected = CommonMethods.GetFromatDateString(DateTimePicker_ExpectedDate.SelectedDate.ToShortDateString());
                string ExpectedQuarter = drp_ExpectedQuarter.SelectedValue;
                if (BasicSkills_DT != null && BasicSkills_DT.Rows.Count > 0)
                {
                    DataRow[] rows1 = BasicSkills_DT.Select("ExpectedDate=" + ExpectedQuarter );
                    DataTable BasicSkillsFilter_DT = BasicSkills_DT.Clone();

                    foreach (DataRow row in rows1)
                    {
                        BasicSkillsFilter_DT.ImportRow(row);
                    }
                    Gv_EmpCources_BasicSkills.DataSource = BasicSkillsFilter_DT;
                    Gv_EmpCources_BasicSkills.DataBind();
                }


                if (DevelopmentSkills_DT != null && DevelopmentSkills_DT.Rows.Count > 0)
                {
                    DataRow[] rows2 = DevelopmentSkills_DT.Select("ExpectedDate=" + ExpectedQuarter);
                    DataTable DevelopmentSkillsFilter_DT = DevelopmentSkills_DT.Clone();

                    foreach (DataRow row in rows2)
                    {
                        DevelopmentSkillsFilter_DT.ImportRow(row);
                    }
                    Gv_EmpCources_DevelopmentSkills.DataSource = DevelopmentSkillsFilter_DT;
                    Gv_EmpCources_DevelopmentSkills.DataBind();
                }

                if (TechnicalSkills_DT != null && TechnicalSkills_DT.Rows.Count > 0)
                {
                    DataRow[] rows3 = TechnicalSkills_DT.Select("ExpectedDate=" + ExpectedQuarter);
                    DataTable TechnicalSkillsFilter_DT = TechnicalSkills_DT.Clone();

                    foreach (DataRow row in rows3)
                    {
                        TechnicalSkillsFilter_DT.ImportRow(row);
                    }
                    Gv_EmpCources_TechnicalSkills.DataSource = TechnicalSkillsFilter_DT;
                    Gv_EmpCources_TechnicalSkills.DataBind();
                }
            }
            #endregion

            #region Filter By Course Name
            else if (ddl_CourseName.SelectedValue != "-1")
            {
                DataTable BasicSkillsFilter_DT = BasicSkills_DT.Clone();
                if (BasicSkills_DT != null && BasicSkills_DT.Rows.Count > 0)
                {
                    DataRow[] rows1 = BasicSkills_DT.Select("CourseName='" + ddl_CourseName.SelectedItem.Text.Trim() + "'");


                    foreach (DataRow row in rows1)
                    {
                        BasicSkillsFilter_DT.ImportRow(row);
                    }
                }

                Gv_EmpCources_BasicSkills.DataSource = BasicSkillsFilter_DT;
                Gv_EmpCources_BasicSkills.DataBind();



                DataTable DevelopmentSkillsFilter_DT = DevelopmentSkills_DT.Clone();
                if (DevelopmentSkills_DT != null && DevelopmentSkills_DT.Rows.Count > 0)
                {
                    DataRow[] rows2 = DevelopmentSkills_DT.Select("CourseName='" + ddl_CourseName.SelectedItem.Text.Trim() + "'");


                    foreach (DataRow row in rows2)
                    {
                        DevelopmentSkillsFilter_DT.ImportRow(row);
                    }
                }
                Gv_EmpCources_DevelopmentSkills.DataSource = DevelopmentSkillsFilter_DT;
                Gv_EmpCources_DevelopmentSkills.DataBind();



                DataTable TechnicalSkillsFilter_DT = TechnicalSkills_DT.Clone();
                if (TechnicalSkills_DT != null && TechnicalSkills_DT.Rows.Count > 0)
                {
                    DataRow[] rows3 = TechnicalSkills_DT.Select("CourseName='" + ddl_CourseName.SelectedItem.Text.Trim() + "'");

                    foreach (DataRow row in rows3)
                    {
                        TechnicalSkillsFilter_DT.ImportRow(row);
                    }
                }
                Gv_EmpCources_TechnicalSkills.DataSource = TechnicalSkillsFilter_DT;
                Gv_EmpCources_TechnicalSkills.DataBind();
            }
            #endregion
            
            else
            {
                Gv_EmpCources_BasicSkills.DataSource = BasicSkills_DT;
                Gv_EmpCources_BasicSkills.DataBind();

                Gv_EmpCources_DevelopmentSkills.DataSource = DevelopmentSkills_DT;
                Gv_EmpCources_DevelopmentSkills.DataBind();

                Gv_EmpCources_TechnicalSkills.DataSource = TechnicalSkills_DT;
                Gv_EmpCources_TechnicalSkills.DataBind();

            }

            #region Filter By Category of courses
            if (ddl_CoursesCategory.SelectedValue != "-1")
            {
                if (ddl_CoursesCategory.SelectedValue == "0")
                {

                    Gv_EmpCources_BasicSkills.Visible = true;
                    Gv_EmpCources_DevelopmentSkills.Visible = false;
                    Gv_EmpCources_TechnicalSkills.Visible = false;


                    ((Label)e.Item.FindControl("header1")).Visible = true;
                    ((Label)e.Item.FindControl("header2")).Visible = false;
                    ((Label)e.Item.FindControl("header3")).Visible = false;
                }
                else if (ddl_CoursesCategory.SelectedValue == "1")
                {
                    Gv_EmpCources_BasicSkills.Visible = false;
                    Gv_EmpCources_DevelopmentSkills.Visible = true;
                    Gv_EmpCources_TechnicalSkills.Visible = false;



                    ((Label)e.Item.FindControl("header1")).Visible = false;
                    ((Label)e.Item.FindControl("header2")).Visible = true;
                    ((Label)e.Item.FindControl("header3")).Visible = false;
                }
                else if (ddl_CoursesCategory.SelectedValue == "2")
                {
                    Gv_EmpCources_BasicSkills.Visible = false;
                    Gv_EmpCources_DevelopmentSkills.Visible = false;
                    Gv_EmpCources_TechnicalSkills.Visible = true;


                    ((Label)e.Item.FindControl("header1")).Visible = false;
                    ((Label)e.Item.FindControl("header2")).Visible = false;
                    ((Label)e.Item.FindControl("header3")).Visible = true;
                }
            }
            else
            {
                if (Gv_EmpCources_BasicSkills.Rows.Count == 0)
                {
                    ((Label)e.Item.FindControl("header1")).Visible = false;
                    Gv_EmpCources_BasicSkills.Visible = false;
                }
                if (Gv_EmpCources_DevelopmentSkills.Rows.Count == 0)
                {
                    ((Label)e.Item.FindControl("header2")).Visible = false;
                    Gv_EmpCources_DevelopmentSkills.Visible = false;
                }
                if (Gv_EmpCources_TechnicalSkills.Rows.Count == 0)
                {
                    ((Label)e.Item.FindControl("header3")).Visible = false;
                    Gv_EmpCources_TechnicalSkills.Visible = false;
                }
            }
            #endregion

            if (Gv_EmpCources_BasicSkills.Rows.Count == 0 && Gv_EmpCources_DevelopmentSkills.Rows.Count == 0 && Gv_EmpCources_TechnicalSkills.Rows.Count == 0)
            { e.Item.Visible = false; }
        }

    }
}

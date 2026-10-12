using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Web.Hosting;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;
using Microsoft.SharePoint.EmailIntegration;
using Microsoft.SharePoint.Utilities;


namespace NileSugar.HrTNA.Webparts.TrainingRequests_WP
{
    public static class CommonMethods
    {

        #region Vars

        // public static EmployeesBalanceDBDataContext NS_SpechialData_DB = new EmployeesBalanceDBDataContext(); // NS_SpechailData

        public static Employee_INFO_ConDataContext EmpDB = new Employee_INFO_ConDataContext();


        #endregion

        public static Hashtable LoadMailTemplates()
        {
            Hashtable MailTemplates;

            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList list = web.Lists.TryGetList("HrTNAMailsTemplates");
                MailTemplates = new Hashtable();
                if (list != null)
                {
                    foreach (SPListItem itm in list.Items)
                    {
                        MailTemplates.Add(itm["Title"].ToString(), itm["TemplateBody"].ToString());
                    }
                }
            }
            return MailTemplates;
        }


        /// <summary>
        ///  run sendEmail method with Elevated Privileges
        /// </summary>
        /// <param name="mailTo">user's email</param>
        /// <param name="mailTitle">title</param>
        /// <param name="textToBesend">body</param>
        /// <returns>true if email was sent s, otherwise false</returns>

        public static Boolean SendMail(string mailTo, string mailTitle, string textToBesend, Guid SiteId, Guid webId)
        {
            //string text = System.IO.File.ReadAllText("/Data"+"/emailTemplate.htm");
            //textToBesend= text.Replace("ReplaceMe", textToBesend);  
            StringDictionary headers = new StringDictionary
            {
                //{"from", "dalzoubi@the-e.net"},
                {"to", mailTo},
                {"subject", mailTitle},
                {"content-type", "text/html"}
            };
            try
            {
                bool flag = false;
                SPSecurity.RunWithElevatedPrivileges(
                    delegate()
                    {
                        using (SPSite site = new SPSite(
                            SiteId))
                        {
                            using (SPWeb web = site.OpenWeb(webId))
                            {
                                flag = SPUtility.SendEmail(web, headers,
                                    textToBesend);
                            }
                        }
                    });
                return flag;
            }
            catch (Exception exp)
            {
                return false;
            }
        }


        //******* Mail Methods

        public static string GetUserMail(string name)
        {
            string mail = string.Empty;
            using (HostingEnvironment.Impersonate())
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                using (SPWeb web = site.OpenWeb())
                {
                    try
                    {
                        PrincipalContext domain = new PrincipalContext(ContextType.Domain);
                        UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(domain, name);

                        if (userPrincipal != null)
                        {
                            mail = userPrincipal.GetMail();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
            return mail;

           // return "test01@nslab.lab";

        }

        public static string GetUserMailByEmpNo(string empNo)
        {
            string mail = string.Empty;

            #region Get UserMail from DB

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.EMPLOYEE_NUMBER == empNo
                      select em;
            if (emp.Any())
            {
                mail = emp.FirstOrDefault().EMAIL_ADDRESS;
            }
            return mail;

            #endregion

           // return "test01@nslab.lab";

        }

        public static string GetMgrMailByEmpCode(string empCode)
        {
            string mgrMail = string.Empty;

            #region Get UserMail from DB

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.EMPLOYEE_NUMBER == empCode
                      select em;
            if (emp.Any())
            {
                mgrMail = emp.FirstOrDefault().SUPERVISOR_EMAIL;
            }
            return mgrMail;

            #endregion

           // return "test01@nslab.lab";

        }

        public static string GetDHMailByEmpCode(string empCode)
        {
            string mgrMail = string.Empty;

            #region Get UserMail from DB

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.EMPLOYEE_NUMBER == empCode
                      select em;
            if (emp.Any())
            {
                mgrMail = emp.FirstOrDefault().DEP_MANAGER_MAIL;
            }
            return mgrMail;

            #endregion


          //  return "test01@nslab.lab";

        }

        public static string GetDVMgrMailByEmpCode(string empCode)
        {
            string mgrMail = string.Empty;

            #region Get UserMail from DB

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.EMPLOYEE_NUMBER == empCode
                      select em;
            if (emp.Any())
            {
                mgrMail = emp.FirstOrDefault().Division_MANAGER_MAIL;
            }
            return mgrMail;

            #endregion

           // return "test01@nslab.lab";

        }

        public static string GetEmpMailByEmpCode(string empCode)
        {
            string mgrMail = string.Empty;

            #region Get UserMail from DB

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.EMPLOYEE_NUMBER == empCode
                      select em;
            if (emp.Any())
            {
                mgrMail = emp.FirstOrDefault().EMAIL_ADDRESS;
            }
            return mgrMail;

            #endregion

           // return "test01@nslab.lab";

        }

        public static string GetDeptHeadMailByEmpCode(string empCode)
        {
            string deptHeadMail = string.Empty;
            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                          where emp.EMPLOYEE_NUMBER == empCode
                          select emp.DEP_MANAGER_MAIL;

            if (empData.Any())
            {
                deptHeadMail = empData.FirstOrDefault();
            }
            return deptHeadMail.Trim();

          //  return "test01@nslab.lab";

        }

        //**************

        public static string GetDivisionNameByDVMgrID(string DVMgrId)
        {
            string mgrMail = string.Empty;

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.EMPLOYEE_NUMBER == DVMgrId
                      select em;
            if (emp.Any())
            {
                mgrMail = emp.FirstOrDefault().Division;
            }
            return mgrMail;

        }

        public static string GetDivisionNameByDHCode(string DHCode)
        {
            string divisionName = string.Empty;



            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.DEP_MANAGER_NUM == DHCode
                      select em;
            if (emp.Any())
            {
                divisionName = emp.FirstOrDefault().Division;
            }
            return divisionName;

        }
        public static int GetNoOfEmpsInDivisionByDivisionName(string DvMgrCode)
        {
            int divisionCount = 0;

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.Division_MANAGER_NUM == DvMgrCode
                      select em;
            if (emp.Any())
            {
                divisionCount = emp.Count();
            }
            return divisionCount;

        }
        public static string GetLookups(string key)
        {
            string value = String.Empty;

            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList HrLookupLst = web.Lists.TryGetList("LookupsLst");
                if (HrLookupLst != null)
                {
                    SPQuery qry = new SPQuery();

                    qry.Query = String.Format(
                        @" <Where>
                                  <Eq>
                                     <FieldRef Name='Title' />
                                     <Value Type='Text'>{0}</Value>
                                  </Eq>
                               </Where> ", key);

                    SPListItemCollection collItems = HrLookupLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        foreach (SPListItem item in collItems)
                        {
                            value = item["Value"].ToString();
                        }
                    }
                }
            }
            return value;
        }

        public static DataTable ToDataTable<T>(this List<T> iList)
        {
            DataTable dataTable = new DataTable();
            PropertyDescriptorCollection propertyDescriptorCollection =
                TypeDescriptor.GetProperties(typeof(T));
            for (int i = 0; i < propertyDescriptorCollection.Count; i++)
            {
                PropertyDescriptor propertyDescriptor = propertyDescriptorCollection[i];
                Type type = propertyDescriptor.PropertyType;

                if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    type = Nullable.GetUnderlyingType(type);


                dataTable.Columns.Add(propertyDescriptor.Name, type);
            }
            object[] values = new object[propertyDescriptorCollection.Count];
            foreach (T iListItem in iList)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = propertyDescriptorCollection[i].GetValue(iListItem);
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }

        public static ArrayList ConvertDataTableToArrayList(DataTable dtTable)
        {
            ArrayList myArrayList = new ArrayList();

            for (int i = 0; i <= dtTable.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dtTable.Columns.Count - 1; j++)
                {
                    myArrayList.Add(dtTable.Rows[i][j].ToString());
                }
            }
            return myArrayList;
        }


        public static string GetLoginEmpDept()
        {
            string loginedEmpDept = string.Empty;
            try
            {
                //using (HostingEnvironment.Impersonate())
                //{
                //    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                //    using (SPWeb web = site.OpenWeb())
                //    {
                //        PrincipalContext domain = new PrincipalContext(ContextType.Domain);

                //        UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(domain,
                //                                        System.Web.HttpContext.Current.User.Identity.Name);

                //        empDept = userPrincipal.GetDepartment();

                //        return empDept;
                //    }
                //}


                var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s

                              where emp.EMPLOYEE_NUMBER == GetLoginEmpCode()
                              select emp.DEPARTMENT_NAME; 


                if (empData.Any())
                {
                    loginedEmpDept = empData.FirstOrDefault();
                }
            }
            catch (Exception)
            {
            }
            return loginedEmpDept;
        }


        public static string GetLoginEmpCode()
        {
            string empCode = string.Empty;
            try
            {
                using (HostingEnvironment.Impersonate())
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        PrincipalContext domain = new PrincipalContext(ContextType.Domain);

                        UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(domain,
                            System.Web.HttpContext.Current.User.Identity.Name);

                        empCode = userPrincipal.EmployeeId;
                    }
                }
            }
            catch (Exception)
            {
            }

            return empCode;

            // return "00318";
        }


        public static string GetMgrCodeByEmpCode(string empCode)
        {
            string mgrCode = string.Empty;

            #region Get UserMail from DB

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.EMPLOYEE_NUMBER == empCode
                      select em;
            if (emp.Any())
            {
                mgrCode = emp.FirstOrDefault().SUPERVISOR_NUMBER;
            }
            return mgrCode;

            #endregion
        }

        public static void CenterTextForAllTextbox(Control control)
        {
            foreach (var c in control.Controls)
            {
                if (c is TextBox) ((TextBox)c).Style["text-align"] = "center";
            }
        }

        public static void DeleteAllItems(string site, string list)
        {
            using (SPSite spSite = new SPSite(site))
            {
                using (SPWeb spWeb = spSite.OpenWeb())
                {
                    StringBuilder deletebuilder = BatchCommand(spWeb.Lists[list]);
                    spSite.RootWeb.ProcessBatchData(deletebuilder.ToString());
                }
            }
        }

        private static StringBuilder BatchCommand(SPList spList)
        {
            StringBuilder deletebuilder = new StringBuilder();
            deletebuilder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Batch>");
            string command = "<Method><SetList Scope=\"Request\">" + spList.ID +
                             "</SetList><SetVar Name=\"ID\">{0}</SetVar><SetVar Name=\"owsfileref\">{1}</SetVar><SetVar Name=\"Cmd\">Delete</SetVar></Method>";

            foreach (SPListItem item in spList.Items)
            {
                deletebuilder.Append(string.Format(command, item.ID.ToString(), item["FileRef"].ToString()));
            }
            deletebuilder.Append("</Batch>");
            return deletebuilder;
        }

        public static string GetAppraiserName(string empCode)
        {
            string appraiserName = string.Empty;
            var emp = (from emps in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                       where emps.EMPLOYEE_NUMBER == empCode
                       select emps).ToList();

            DataTable dt_emp = emp.ToDataTable();

            if (dt_emp != null && dt_emp.Rows.Count > 0)
            {
                appraiserName = dt_emp.Rows[0]["SUPERVISOR_NAME"].ToString();
            }
            return appraiserName;
        }

        public static string GetAppraiserOfNonCompUser(string empCode)
        {
            string appraiserName = string.Empty;
            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {

                SPList nonComUsersLst = web.Lists.TryGetList("NonCompUsers");
                if (nonComUsersLst != null)
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

                    SPListItemCollection collItems = nonComUsersLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        foreach (SPListItem item in collItems)
                        {
                            appraiserName = item["AppraiserName"].ToString();
                        }
                    }
                }

                //*********************
                if (string.IsNullOrEmpty(appraiserName))
                {
                    SPList ConvertedNonComUsersLst = web.Lists.TryGetList("ConvertedToNonCompEmps");
                    if (ConvertedNonComUsersLst != null)
                    {
                        SPQuery qry2 = new SPQuery();

                        qry2.Query = String.Format(
                            @" 
                                   <Where>
                                      <Eq>
                                         <FieldRef Name='EmpCode' />
                                         <Value Type='Text'>{0}</Value>
                                      </Eq>
                                   </Where>
                                       ", empCode);

                        SPListItemCollection collItems = ConvertedNonComUsersLst.GetItems(qry2);
                        if (collItems.Count > 0)
                        {
                            foreach (SPListItem item in collItems)
                            {
                                appraiserName = item["AppraiserName"].ToString();
                            }
                        }
                    }
                }

            }

            return appraiserName;
        }

        public static void UpdateRateSettings(string rateName, int rateValue)
        {
            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                web.AllowUnsafeUpdates = true;
                SPList HrReqLookupLst = web.Lists.TryGetList("HrAppraisalSettingsRates");
                if (HrReqLookupLst != null)
                {
                    SPQuery qry = new SPQuery();

                    qry.Query = String.Format(
                        @" 
                                <Where>
                                  <Eq>
                                     <FieldRef Name='Title' />
                                     <Value Type='Text'>{0}</Value>
                                  </Eq>
                               </Where>
                                       ", rateName);

                    SPListItemCollection collItems = HrReqLookupLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        foreach (SPListItem item in collItems)
                        {
                            item["RateValue"] = rateValue.ToString();
                            item.Update();
                        }
                    }
                }

                web.AllowUnsafeUpdates = false;
            }

        }

        public static string GetRateValue(string rateName)
        {
            string value = String.Empty;

            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList HrAppraisalValuesLst = web.Lists.TryGetList("HrAppraisalSettingsRates");
                if (HrAppraisalValuesLst != null)
                {
                    SPQuery qry = new SPQuery();

                    qry.Query = String.Format(
                        @" 
                                <Where>
                                  <Eq>
                                     <FieldRef Name='Title' />
                                     <Value Type='Text'>{0}</Value>
                                  </Eq>
                               </Where>
                                       ", rateName);

                    SPListItemCollection collItems = HrAppraisalValuesLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        foreach (SPListItem item in collItems)
                        {
                            value = item["RateValue"].ToString();
                        }
                    }
                }
            }
            return value;
        }

        public static string GetAppraisalRate(double rate)
        {
            //string scale = string.Empty;

            //if (rate >= 105)
            //    scale = "Outstanding";
            //if (rate > 100 && rate < 105)
            //    scale = "Excellent";
            //if (rate >= 80 && rate <= 100)
            //    scale = "Effective";
            //if (rate >= 50 && rate < 80)
            //    scale = "Needs Development";
            //if (rate < 50)
            //    scale = "Unsatisfactory";
            //return scale;

            string scale = string.Empty;

            if (rate > 100)
                scale = "Outstanding";
            if (rate >= 80 && rate <= 100)
                scale = "Excellent";
            if (rate >= 60 && rate < 80)
                scale = "Effective";
            if (rate >= 44 && rate < 60)
                scale = "Needs Development";
            if (rate < 44)
                scale = "Unsatisfactory";
            return scale;
        }


        public static string GetArAppraisalRate(double rate)
        {
            string scale = string.Empty;

            if (rate > 100)
                scale = "أداء استثنائي متميز";
            if (rate >= 80 && rate <= 100)
                scale = "أداء يفوق التوقعات";
            if (rate >= 60 && rate < 80)
                scale = "أداء ملاق للتوقعات";
            if (rate >= 44 && rate < 60)
                scale = "أداء يحتاج إلى تحسين";
            if (rate < 44)
                scale = "أداء غير مرضى";
            return scale;
        }

        public static string GetEmpDeptByEmpCode(string EmployeeID)
        {
            string loginedEmpDept = string.Empty;

            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                          where emp.EMPLOYEE_NUMBER == EmployeeID.ToString()
                          select emp.DEPARTMENT_NAME; // select emp.DEPARTMENT_NAME;


            if (empData.Any())
            {
                loginedEmpDept = empData.FirstOrDefault();
            }
            return loginedEmpDept;
        }

        public static string GetEmpDHCodeByEmpCode(string EmployeeID)
        {
            string DHCode = string.Empty;

            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                          where emp.EMPLOYEE_NUMBER == EmployeeID.ToString()
                          select emp.DEP_MANAGER_NUM;


            if (empData.Any())
            {
                DHCode = empData.FirstOrDefault();
            }
            return DHCode;
        }
        public static string GetEmpDVMgrCodeByEmpCode(string EmployeeID)
        {
            string DVCode = string.Empty;

            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                          where emp.EMPLOYEE_NUMBER == EmployeeID.ToString()
                          select emp.Division_MANAGER_NUM;


            if (empData.Any())
            {
                DVCode = empData.FirstOrDefault();
            }
            return DVCode;
        }
        public static string GetEmpCodeByEmpFullName(string empFullName)
        {
            string empCode = string.Empty;
            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                          where emp.FULL_NAME == empFullName
                          select emp.EMPLOYEE_NUMBER;


            if (empData.Any())
            {
                empCode = empData.FirstOrDefault();
            }
            return empCode;
        }

        public static string GetLoginEmpDisplayName()
        {
            //try
            //{
            //    using (HostingEnvironment.Impersonate())
            //    {
            //        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            //        using (SPWeb web = site.OpenWeb())
            //        {
            //            PrincipalContext domain = new PrincipalContext(ContextType.Domain);

            //            string str = System.Web.HttpContext.Current.User.Identity.Name;

            //            UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(domain,
            //                                            System.Web.HttpContext.Current.User.Identity.Name);

            //            empDisplayName = userPrincipal.GetEmployeeDisplayName();

            //            return empDisplayName;
            //        }
            //    }
            //}
            //catch (Exception)
            //{

            //}
            //return empDisplayName;

            string empDisplayName = string.Empty;
            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                          where emp.EMPLOYEE_NUMBER == GetLoginEmpCode()
                          select emp.FULL_NAME;


            if (empData.Any())
            {
                empDisplayName = empData.FirstOrDefault();
            }
            return empDisplayName;
        }

        public static string GetDHCodeByDeptName(string deptName)
        {
            string DHCode = string.Empty;
            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                          where emp.DEPARTMENT_NAME == deptName
                          select emp.DEP_MANAGER_NUM;

            if (empData.Any())
            {
                DHCode = empData.FirstOrDefault();
            }
            return DHCode;
        }

        public static string FormartEmpCode(string empCode)
        {
            if (empCode.Length < 5)
            {
                if (empCode.Length == 1)
                    empCode = "0000" + empCode;

                if (empCode.Length == 2)
                    empCode = "000" + empCode;

                if (empCode.Length == 3)
                    empCode = "00" + empCode;

                if (empCode.Length == 4)
                    empCode = "0" + empCode;
            }

            return empCode;
        }

        //public static DataTable GetDeptHeads()
        //{
        //    DataTable DtDept = new DataTable();
        //    DtDept.Columns.Add("DeptName", typeof(String));
        //    DtDept.Columns.Add("DHName", typeof(String));
        //    DtDept.Columns.Add("DHEmail", typeof(String));
        //    using (HostingEnvironment.Impersonate())
        //    {
        //        PrincipalContext domain = new PrincipalContext(ContextType.Domain);
        //        using (var DepartMentHeadgroup = GroupPrincipal.FindByIdentity(domain, "DL_DepartmentsHeads"))
        //        {
        //            if (DepartMentHeadgroup != null)
        //            {
        //                var Users = DepartMentHeadgroup.GetMembers(true);

        //                // List<UserPrincipal> Users2 = Users;
        //                foreach (UserPrincipal user in Users)
        //                {
        //                    DataRow row = DtDept.NewRow();
        //                    row["DeptName"] = NileSugar.Shared.Utilities.CommonMethodes.GetEmpData(user, "Department");
        //                    row["DHName"] = NileSugar.Shared.Utilities.CommonMethodes.GetEmpData(user, "name");
        //                    row["DHEmail"] = NileSugar.Shared.Utilities.CommonMethodes.GetEmpData(user, "mail");
        //                    row["DeptHeadCode"] = NileSugar.Shared.Utilities.CommonMethodes.GetEmpData(user, "employeeID");
        //                    DtDept.Rows.Add(row);
        //                }
        //            }
        //        }
        //    }
        //    return DtDept;
        //}

        public static string ConvertToArabicNumerals(string input)
        {
            System.Text.UTF8Encoding utf8Encoder = new UTF8Encoding();
            System.Text.Decoder utf8Decoder = utf8Encoder.GetDecoder();
            System.Text.StringBuilder convertedChars = new System.Text.StringBuilder();
            char[] convertedChar = new char[1];
            byte[] bytes = new byte[] { 217, 160 };
            char[] inputCharArray = input.ToCharArray();
            foreach (char c in inputCharArray)
            {
                if (char.IsDigit(c))
                {
                    bytes[1] = Convert.ToByte(160 + char.GetNumericValue(c));
                    utf8Decoder.GetChars(bytes, 0, 2, convertedChar, 0);
                    convertedChars.Append(convertedChar[0]);
                }
                else
                {
                    convertedChars.Append(c);
                }
            }
            return convertedChars.ToString();
        }

        public static string GetAppraisalYear()
        {
            string apprYear = string.Empty;
            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList SettingLst = web.Lists.TryGetList("HrAppraisalSettings");
                if (SettingLst != null && SettingLst.Items.Count > 0)
                {
                    foreach (SPListItem item in SettingLst.Items)
                    {
                        apprYear = item["HrApprYear"].ToString();
                    }
                }
            }
            return apprYear;
        }

        public static int GetTotalCountOfEmpsByDept(string divisionName, string deptName)
        {
            int totalCount = 0;
            if (divisionName == "All Divisions")
            {
                var AllEmps = (from emps in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                               select emps).ToList();

                totalCount = AllEmps.Count;
            }
            else
            {
                if (deptName == "All Department")
                {
                    var AlldeptEmps = (from emps in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                                       where emps.Division == divisionName
                                       select emps).ToList();

                    totalCount = AlldeptEmps.Count;
                }
                else if (divisionName == string.Empty)
                {
                    var deptEmps = (from emps in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                                    where emps.DEPARTMENT_NAME == deptName
                                    select emps).ToList();

                    totalCount = deptEmps.Count;
                }
                else
                {
                    var deptEmps = (from emps in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                                    where emps.DEPARTMENT_NAME == deptName && emps.Division == divisionName
                                    select emps).ToList();

                    totalCount = deptEmps.Count;
                }
            }

            return totalCount;
        }

        public static bool CheckDHDeptLikeHisEmps(string loginedDHCode, string deptName)
        {
            bool isExist = false;
            DataTable DHDepts = GetDHDeptsNames(loginedDHCode);
            isExist = DHDepts.Select().ToList().Exists(row => row["Department"].ToString() == deptName);

            return isExist;
        }

        public static string GetEmpAppraisalFormWithGrade(string GradeName)
        {
            string appraisalForm = string.Empty;
            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList GradesAppraisalsFormsLst = web.Lists.TryGetList("GradesAppraisalsForms");
                if (GradesAppraisalsFormsLst != null)
                {
                    SPQuery qry = new SPQuery();

                    qry.Query = String.Format(
                        @" 
                                  <Where>
                                      <Eq>
                                         <FieldRef Name='GradeName' />
                                         <Value Type='Text'>{0}</Value>
                                      </Eq>
                                   </Where>
                                       ", GradeName);

                    SPListItemCollection collItems = GradesAppraisalsFormsLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        foreach (SPListItem item in collItems)
                        {
                            appraisalForm = item["AppraisalForm"].ToString().Trim();
                        }
                    }
                }
            }
            return appraisalForm;
        }

        public static string GetEmpGradeWithEmpCode(string empCode)
        {
            string empGrade = string.Empty;
            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                          where emp.EMPLOYEE_NUMBER == empCode
                          select emp.GRADE_NAME;


            if (empData.Any())
            {
                empGrade = empData.FirstOrDefault();
            }
            return empGrade.Trim();
        }

        public static bool CheckEmpHasNonCompUsers(string loginedEmpCode)
        {
            bool hasNonComUsers = false;
            bool hasConvertedToNonCompUser = false;
            bool hasUsers = false;

            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList NonCompUsersLst = web.Lists.TryGetList("NonCompUsers");
                if (NonCompUsersLst != null)
                {
                    SPQuery qry = new SPQuery();

                    qry.Query = String.Format(
                        @" 
                                     <Where>
                                          <Eq>
                                             <FieldRef Name='AppraiserCode' />
                                             <Value Type='Text'>{0}</Value>
                                          </Eq>
                                       </Where>
                                       ", loginedEmpCode);

                    SPListItemCollection collItems = NonCompUsersLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        hasNonComUsers = true;
                    }
                }

                //*******************
                SPList ConvertedToNonCompUserLst = web.Lists.TryGetList("ConvertedToNonCompEmps");
                if (ConvertedToNonCompUserLst != null)
                {
                    SPQuery qry2 = new SPQuery();

                    qry2.Query = String.Format(
                        @" 
                                        <Where>
                                              <Eq>
                                                 <FieldRef Name='AppraiserCode' />
                                                 <Value Type='Text'>{0}</Value>
                                              </Eq>
                                           </Where>
                                       ", loginedEmpCode);

                    SPListItemCollection collItems = ConvertedToNonCompUserLst.GetItems(qry2);
                    if (collItems.Count > 0)
                    {
                        hasConvertedToNonCompUser = true;
                    }
                }
            }

            if (hasNonComUsers || hasConvertedToNonCompUser)
                hasUsers = true;

            return hasUsers;
        }

        public static bool CheckEmpWithoutSelfAppraisal(string empCode)
        {
            bool empWithoutAppraisal = false;
            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList EmpWithoutAppraisalLst = web.Lists.TryGetList("EmpsWithoutAppraisals");
                if (EmpWithoutAppraisalLst != null)
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

                    SPListItemCollection collItems = EmpWithoutAppraisalLst.GetItems(qry);
                    if (collItems.Count > 0)
                    {
                        empWithoutAppraisal = true;
                    }
                }
            }
            return empWithoutAppraisal;
        }


        public static bool CheckIfLoginEmpIsMgr(string empCode)
        {
            bool isMgr = false;
            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                          where emp.SUPERVISOR_NUMBER == empCode
                          select emp;


            if (empData.Any())
            {
                isMgr = true;
            }
            return isMgr;
        }

        public static int GetNoOfEmpsPerRate(string divisionName, string deptName, string rateName)
        {
            int NoOfEmpsPerRate = 0;

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList HrAppraisalValuesLst = web.Lists.TryGetList("HrAppraisalValues");
                        if (HrAppraisalValuesLst != null)
                        {
                            SPQuery qry = new SPQuery();
                            if (divisionName == "All Divisions")
                            {
                                qry.Query = string.Format(
                                      @" 
                                       <Where>
                                          <And>
                                             <Eq>
                                                <FieldRef Name='MgrResultScale' />
                                                <Value Type='Text'>{0}</Value>
                                             </Eq>
                                             <Eq>
                                                <FieldRef Name='AppraisalYear' />
                                                <Value Type='Text'>{1}</Value>
                                             </Eq>
                                          </And>
                                       </Where>
                                       ", rateName, GetAppraisalYear());
                            }
                            else
                            {
                                if (deptName == "All Department")
                                {
                                    qry.Query = string.Format(
                                        @" 
                                         <Where>
                                              <And>
                                                 <Eq>
                                                    <FieldRef Name='EmpDivision' />
                                                    <Value Type='Text'>{0}</Value>
                                                 </Eq>
                                                 <And>
                                                    <Eq>
                                                       <FieldRef Name='MgrResultScale' />
                                                       <Value Type='Text'>{1}</Value>
                                                    </Eq>
                                                    <Eq>
                                                       <FieldRef Name='AppraisalYear' />
                                                       <Value Type='Text'>{2}</Value>
                                                    </Eq>
                                                 </And>
                                              </And>
                                           </Where>
                                       ", divisionName, rateName, GetAppraisalYear());
                                }
                                else
                                {
                                    if (divisionName == string.Empty)
                                    {
                                        qry.Query = string.Format(
                                            //                                            @" 
                                            //                                      <Where>
                                            //                                          <And>
                                            //                                             <Eq>
                                            //                                                <FieldRef Name='MgrResultScale' />
                                            //                                                <Value Type='Text'>{0}</Value>
                                            //                                             </Eq>
                                            //                                             <And>
                                            //                                                <Eq>
                                            //                                                   <FieldRef Name='EmpDepartment' />
                                            //                                                   <Value Type='Text'>{1}</Value>
                                            //                                                </Eq>
                                            //                                                <Eq>
                                            //                                                   <FieldRef Name='AppraisalYear' />
                                            //                                                   <Value Type='Text'>{2}</Value>
                                            //                                                </Eq>
                                            //                                             </And>
                                            //                                          </And>
                                            //                                       </Where>
                                            //                                       "
                                            @"
   <Where>
      <And>
         <Eq>
            <FieldRef Name='MgrResultScale' />
            <Value Type='Text'>{0}</Value>
         </Eq>
         <And>
            <Eq>
               <FieldRef Name='EmpDepartment' />
               <Value Type='Text'>{1}</Value>
            </Eq>
            <And>
               <Eq>
                  <FieldRef Name='AppraisalYear' />
                  <Value Type='Text'>{2}</Value>
               </Eq>
               <Neq>
                  <FieldRef Name='EmpCode' />
                  <Value Type='Text'>{3}</Value>
               </Neq>
            </And>
         </And>
      </And>
   </Where>"
                                            , rateName, deptName, GetAppraisalYear(), GetLoginEmpCode());
                                    }

                                    else
                                    {
                                        qry.Query = string.Format(
                                            @" 
                                    <Where>
                                      <And>
                                         <And>
                                            <Eq>
                                               <FieldRef Name='EmpDivision' />
                                               <Value Type='Text'>{0}</Value>
                                            </Eq>
                                            <Eq>
                                               <FieldRef Name='EmpDepartment' />
                                               <Value Type='Text'>{1}</Value>
                                            </Eq>
                                         </And>
                                         <And>
                                            <Eq>
                                               <FieldRef Name='MgrResultScale' />
                                               <Value Type='Text'>{2}</Value>
                                            </Eq>
                                            <Eq>
                                               <FieldRef Name='AppraisalYear' />
                                               <Value Type='Text'>{3}</Value>
                                            </Eq>
                                         </And>
                                      </And>
                                   </Where>
                                       ", divisionName, deptName, rateName, GetAppraisalYear());
                                    }
                                }
                            }
                            SPListItemCollection collItems = HrAppraisalValuesLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                NoOfEmpsPerRate = collItems.Count;
                            }
                        }
                    }
                });
            return NoOfEmpsPerRate;
        }

        public static int GetNoOfEmpsPerRateForCEO(string dvMgrCode, string rateName)
        {
            int NoOfEmpsPerRate = 0;

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList HrAppraisalValuesLst = web.Lists.TryGetList("HrAppraisalValues");
                        if (HrAppraisalValuesLst != null)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = string.Format(
                                @" 
                                   <Where>
                                      <And>
                                         <And>
                                            <Eq>
                                               <FieldRef Name='DVMgrCode' />
                                               <Value Type='Text'>{0}</Value>
                                            </Eq>
                                            <Eq>
                                               <FieldRef Name='MgrResultScale' />
                                               <Value Type='Text'>{1}</Value>
                                            </Eq>
                                         </And>
                                         <Eq>
                                            <FieldRef Name='AppraisalYear' />
                                            <Value Type='Text'>{2}</Value>
                                         </Eq>
                                      </And>
                                   </Where>
                                       ", dvMgrCode, rateName, GetAppraisalYear());



                            SPListItemCollection collItems = HrAppraisalValuesLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                NoOfEmpsPerRate = collItems.Count;
                            }
                        }

                    }
                });
            return NoOfEmpsPerRate;
        }

        public static int GetNoOfEmpsPerStatus(string divisionName, string deptName, string statusName)
        {
            int NoOfEmpsPerStatus = 0;

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList HrAppraisalValuesLst = web.Lists.TryGetList("HrAppraisalValues");
                        if (HrAppraisalValuesLst != null)
                        {
                            SPQuery qry = new SPQuery();
                            if (divisionName == "All Divisions")
                            {
                                qry.Query = string.Format(
                                       @" 
                                           <Where>
                                              <And>
                                                 <Eq>
                                                    <FieldRef Name='AppraisalStatus' />
                                                    <Value Type='Text'>{0}</Value>
                                                 </Eq>
                                                 <Eq>
                                                    <FieldRef Name='AppraisalYear' />
                                                    <Value Type='Text'>{1}</Value>
                                                 </Eq>
                                              </And>
                                           </Where>
                                       ", statusName, GetAppraisalYear());
                            }
                            else
                            {
                                if (deptName == "All Department")
                                {
                                    qry.Query = string.Format(
                                        @" 
                                              <Where>
                                                      <And>
                                                         <Eq>
                                                            <FieldRef Name='EmpDivision' />
                                                            <Value Type='Text'>{0}</Value>
                                                         </Eq>
                                                         <And>
                                                            <Eq>
                                                               <FieldRef Name='AppraisalStatus' />
                                                               <Value Type='Text'>{1}</Value>
                                                            </Eq>
                                                            <Eq>
                                                               <FieldRef Name='AppraisalYear' />
                                                               <Value Type='Text'>{2}</Value>
                                                            </Eq>
                                                         </And>
                                                      </And>
                                                   </Where>
                                       ", divisionName, statusName, GetAppraisalYear());
                                }
                                else
                                {
                                    if (divisionName == string.Empty)
                                    {
                                        qry.Query = string.Format(
                                            @" 
                                                                             <Where>
                                                                                      <And>
                                                                                         <Eq>
                                                                                            <FieldRef Name='EmpDepartment' />
                                                                                            <Value Type='Text'>{0}</Value>
                                                                                         </Eq>
                                                                                         <And>
                                                                                            <Eq>
                                                                                               <FieldRef Name='AppraisalStatus' />
                                                                                               <Value Type='Text'>{1}</Value>
                                                                                            </Eq>
                                                                                            <Eq>
                                                                                               <FieldRef Name='AppraisalYear' />
                                                                                               <Value Type='Text'>{2}</Value>
                                                                                            </Eq>
                                                                                         </And>
                                                                                      </And>
                                                                                   </Where>
                                                                               "
                                            , deptName, statusName, GetAppraisalYear());
                                    }
                                    else
                                    {
                                        qry.Query = string.Format(
                                      @" 
                                                                            <Where>
                                                                                  <And>
                                                                                     <And>
                                                                                        <Eq>
                                                                                           <FieldRef Name='EmpDivision' />
                                                                                           <Value Type='Text'>{0}</Value>
                                                                                        </Eq>
                                                                                        <Eq>
                                                                                           <FieldRef Name='EmpDepartment' />
                                                                                           <Value Type='Text'>{1}</Value>
                                                                                        </Eq>
                                                                                     </And>
                                                                                     <And>
                                                                                        <Eq>
                                                                                           <FieldRef Name='AppraisalStatus' />
                                                                                           <Value Type='Text'>{2}</Value>
                                                                                        </Eq>
                                                                                        <Eq>
                                                                                           <FieldRef Name='AppraisalYear' />
                                                                                           <Value Type='Text'>{3}</Value>
                                                                                        </Eq>
                                                                                     </And>
                                                                                  </And>
                                                                               </Where>
                                                                               "
                                      , divisionName, deptName, statusName, GetAppraisalYear());
                                    }

                                }
                            }
                            SPListItemCollection collItems = HrAppraisalValuesLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                NoOfEmpsPerStatus = collItems.Count;
                            }
                        }
                    }
                });
            return NoOfEmpsPerStatus;
        }

        public static string GetDeptHeadCodeByEmpCode(string empCode)
        {
            string deptHeadCode = string.Empty;
            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                          where emp.EMPLOYEE_NUMBER == empCode
                          select emp.DEP_MANAGER_NUM;

            if (empData.Any())
            {
                deptHeadCode = empData.FirstOrDefault();
            }
            return deptHeadCode.Trim();
        }

        public static DataTable GetAllDeptHeads()
        {
            //DataTable dt_DH = null;
            //var empData = (from emp in EmpDB.NS_EMPLOYEE_INFO_TNAs
            //               where emp.EMPLOYEE_NUMBER == emp.DEP_MANAGER_NUM
            //               select emp).ToList();

            //if (empData.Any())
            //{
            //    dt_DH = empData.ToDataTable();
            //}
            //return dt_DH;

            DataTable dt_emps = null;
            DataTable dt_DH = null;

            // for testing phase
            // where emp.Department == "IT" || emp.Department == "HR"
            var empData = (from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s

                           select emp).ToList();

            if (empData.Any())
            {
                DataView view = new DataView(empData.ToDataTable());
                DataTable distinctDHCode = view.ToTable(true, "DEP_MANAGER_NUM");
                string qry = "EMPLOYEE_NUMBER IN " + CreateFilterQry(distinctDHCode);

                dt_emps = empData.ToDataTable();

                DataView dv = new DataView(dt_emps);
                dv.RowFilter = qry;

                dt_DH = dv.ToTable();
            }
            return dt_DH;
        }

        public static DataTable GetAllDivisionMgrs()
        {
            DataTable dt_emps = null;
            DataTable dt_DH = null;
            DataTable distinctDVMgrCode = null;

            var empData = (from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                           select emp).ToList();

            if (empData.Any())
            {
                DataView view = new DataView(empData.ToDataTable());
                distinctDVMgrCode = view.ToTable(true, "Division_MANAGER_NUM");

                //string qry = "EMPLOYEE_NUMBER IN " + CreateFilterQry(distinctDVMgrCode);

                //dt_emps = empData.ToDataTable();

                //DataView dv = new DataView(dt_emps);
                //dv.RowFilter = qry;

                //dt_DH = dv.ToTable();
            }
            return distinctDVMgrCode;
        }

        private static string CreateFilterQry(DataTable distinctDHCode)
        {
            string qry = string.Empty;
            foreach (DataRow row in distinctDHCode.Rows)
            {
                qry += "'" + row["DEP_MANAGER_NUM"].ToString() + "'" + ",";
            }
            qry = qry.Substring(0, qry.Length - 1);
            qry = "(" + qry + ")";
            return qry;
        }

        public static bool CheckIfLoginEmpIsDeptHead(string empCode)
        {
            bool isDeptHead = false;
            //var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_TNAs
            //              where emp.EMPLOYEE_NUMBER == empCode
            //              select emp.DEP_MANAGER_NUM;


            //if (empData.Any())
            //{
            //    if (empData.FirstOrDefault() == empCode)
            //        isDeptHead = true;
            //}
            //return isDeptHead;

            DataTable dt_DH = GetAllDeptHeads();
            if (dt_DH != null && dt_DH.Rows.Count > 0)
                isDeptHead = dt_DH.Select().ToList().Exists(row => row["EMPLOYEE_NUMBER"].ToString() == empCode);

            return isDeptHead;
        }

        public static bool CheckIfLoginEmpIsDivisionMgr(string empCode)
        {
            bool isDivisionMgr = false;

            DataTable dt_DV = GetAllDivisionMgrs();
            isDivisionMgr = dt_DV.Select().ToList().Exists(row => row["Division_MANAGER_NUM"].ToString() == empCode);

            return isDivisionMgr;
        }

        public static bool CheckIfEmpAppraisalIsNotAppraisedByMgr(string empCode)
        {
            bool notApprByMgr = false;

            int apprIsExist = CheckIfEmpAppraisalIsExist(empCode);

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList HrAppraisalValuesLst = web.Lists.TryGetList("HrAppraisalValues");
                        if (HrAppraisalValuesLst != null)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = string.Format(
                                @" 
                                    <Where>
                                          <And>
                                             <Eq>
                                                <FieldRef Name='EmpCode' />
                                                <Value Type='Text'>{0}</Value>
                                             </Eq>
                                             <And>
                                                <Eq>
                                                   <FieldRef Name='AppraisalStatus' />
                                                   <Value Type='Text'>Appraised by Employee</Value>
                                                </Eq>
                                                <Eq>
                                                   <FieldRef Name='AppraisalYear' />
                                                   <Value Type='Text'>{1}</Value>
                                                </Eq>
                                             </And>
                                          </And>
                                       </Where>
                                       ", empCode, GetAppraisalYear());

                            SPListItemCollection collItems = HrAppraisalValuesLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                notApprByMgr = true;
                            }

                            if (!notApprByMgr && apprIsExist == 0)
                                notApprByMgr = true;
                        }
                    }
                });
            return notApprByMgr;
        }

        public static bool CheckIfEmpAppraisalIsReviewedByDeptHead(string empCode)
        {
            bool reviewedByDHd = false;

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList HrAppraisalValuesLst = web.Lists.TryGetList("HrAppraisalValues");
                        if (HrAppraisalValuesLst != null)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = string.Format(
                                @" 
                               <Where>
                                  <And>
                                     <Eq>
                                        <FieldRef Name='EmpCode' />
                                        <Value Type='Text'>{0}</Value>
                                     </Eq>
                                     <And>
                                        <Eq>
                                           <FieldRef Name='AppraisalStatus' />
                                           <Value Type='Text'>Reviewed By Department Head</Value>
                                        </Eq>
                                        <Eq>
                                           <FieldRef Name='AppraisalYear' />
                                           <Value Type='Text'>{1}</Value>
                                        </Eq>
                                     </And>
                                  </And>
                               </Where>
                                       ", empCode, GetAppraisalYear());

                            SPListItemCollection collItems = HrAppraisalValuesLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                reviewedByDHd = true;
                            }
                        }
                    }
                });
            return reviewedByDHd;
        }

        //Rehab
        public static bool CheckIfEmpAppraisalIsAfterDHReview(string empCode)
        {
            bool AfterDHReview = false;

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList HrAppraisalValuesLst = web.Lists.TryGetList("HrAppraisalValues");
                        if (HrAppraisalValuesLst != null)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = string.Format(

                                @"   
                                   <Where>
                                      <And>
                                         <Eq>
                                            <FieldRef Name='EmpCode' />
                                            <Value Type='Text'>{0}</Value>
                                         </Eq>
                                         <And>
                                            <Eq>
                                               <FieldRef Name='AppraisalYear' />
                                               <Value Type='Text'>{1}</Value>
                                            </Eq>
                                            <Or>
                                               <Eq>
                                                  <FieldRef Name='AppraisalStatus' />
                                                  <Value Type='Text'>Waiting For Employee Approval</Value>
                                               </Eq>
                                               <Or>
                                                  <Eq>
                                                     <FieldRef Name='AppraisalStatus' />
                                                     <Value Type='Text'>Waiting For DH Final Approval</Value>
                                                  </Eq>
                                                  <Eq>
                                                     <FieldRef Name='AppraisalStatus' />
                                                     <Value Type='Text'>Sent to HR</Value>
                                                  </Eq>
                                               </Or>
                                            </Or>
                                         </And>
                                      </And>
                                   </Where>
                                "

                                , empCode, GetAppraisalYear().Trim());

                            SPListItemCollection collItems = HrAppraisalValuesLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                AfterDHReview = true;

                                //                                DataTable dt = collItems.GetDataTable();
                                //                                if (dt != null && dt.Rows.Count > 0)
                                //                                {
                                //                                    if (dt.Rows[0]["AppraisalStatus"].ToString() == "Waiting For Employee Approval"
                                //                                        || dt.Rows[0]["AppraisalStatus"].ToString() == "Accepted By Employee"
                                //                                        || dt.Rows[0]["AppraisalStatus"].ToString() == "Not Accepted By Employee"
                                //                                        || dt.Rows[0]["AppraisalStatus"].ToString() == "Waiting For DH Final Approval"
                                //                                        || dt.Rows[0]["AppraisalStatus"].ToString() == "Sent to HR"
                                //)
                                //                                    {
                                //                                        AfterDHReview = true;
                                //                                    }
                                //                                }

                            }
                        }
                    }
                });
            return AfterDHReview;
        }

        public static int CheckIfEmpAppraisalIsExist(string empCode)
        {
            int apprId = 0;
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList HrAppraisalValuesLst = web.Lists.TryGetList("HrAppraisalValues");
                        if (HrAppraisalValuesLst != null)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = string.Format(
                                @" 
                                 <Where>
                                      <And>
                                         <Eq>
                                            <FieldRef Name='EmpCode' />
                                            <Value Type='Text'>{0}</Value>
                                         </Eq>
                                         <Eq>
                                            <FieldRef Name='AppraisalYear' />
                                            <Value Type='Text'>{1}</Value>
                                         </Eq>
                                      </And>
                                   </Where>
                                       ", empCode, GetAppraisalYear());

                            SPListItemCollection collItems = HrAppraisalValuesLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                apprId = Convert.ToInt32(collItems[0]["ID"].ToString());
                            }
                        }
                    }
                });
            return apprId;
        }

        public static int GetNoOfNonAppraisalEmpsPerDept(string division, string dept)
        {
            int NoOfNonAppraisalEmps = 0;
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList EmpWithoutAppraisalLst = web.Lists.TryGetList("EmpsWithoutAppraisals");
                        if (EmpWithoutAppraisalLst != null)
                        {
                            SPQuery qry = new SPQuery();
                            if (division == "All Divisions")
                            {
                                qry.Query = string.Format(
                                    @" 
                                    <OrderBy>
                                      <FieldRef Name='ID' Ascending='True' />
                                   </OrderBy>
                                       ");
                            }
                            else
                            {
                                if (dept == "All Department")
                                {
                                    qry.Query = string.Format(
                                       @" 
                                     <Where>
                                          <Eq>
                                             <FieldRef Name='Division' />
                                             <Value Type='Text'>{0}</Value>
                                          </Eq>
                                       </Where>
                                       ", division);
                                }
                                else
                                {
                                    if (division == string.Empty)
                                    {
                                        qry.Query = string.Format(
                                           @" 
                                         <Where>
                                              <Eq>
                                                 <FieldRef Name='Dept' />
                                                 <Value Type='Text'>{0}</Value>
                                              </Eq>
                                           </Where>
                                       ", dept);
                                    }
                                    else
                                    {

                                        qry.Query = string.Format(
                                            @" 
                                     <Where>
                                          <And>
                                             <Eq>
                                                <FieldRef Name='Division' />
                                                <Value Type='Text'>{0}</Value>
                                             </Eq>
                                             <Eq>
                                                <FieldRef Name='Dept' />
                                                <Value Type='Text'>{1}</Value>
                                             </Eq>
                                          </And>
                                       </Where>
                                       ", division, dept);
                                    }
                                }
                            }

                            SPListItemCollection collItems = EmpWithoutAppraisalLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                NoOfNonAppraisalEmps = collItems.Count;
                            }
                        }
                    }
                });
            return NoOfNonAppraisalEmps;
        }

        public static int GetNoOfNonAppraisalEmpsPerDivision(string divisionName)
        {
            int NoOfNonAppraisalEmps = 0;
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList EmpWithoutAppraisalLst = web.Lists.TryGetList("EmpsWithoutAppraisals");
                        if (EmpWithoutAppraisalLst != null)
                        {
                            SPQuery qry = new SPQuery();
                            if (divisionName == "All Department")
                            {
                                qry.Query = string.Format(
                                    @" 
                                    <OrderBy>
                                      <FieldRef Name='ID' Ascending='True' />
                                   </OrderBy>
                                       ");
                            }
                            else
                            {
                                qry.Query = string.Format(
                                    @" 
                                   <Where>
                                      <Eq>
                                         <FieldRef Name='Division' />
                                         <Value Type='Text'>{0}</Value>
                                      </Eq>
                                   </Where>
                                       ", divisionName);
                            }

                            SPListItemCollection collItems = EmpWithoutAppraisalLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                NoOfNonAppraisalEmps = collItems.Count;
                            }
                        }
                    }
                });
            return NoOfNonAppraisalEmps;
        }

        public static int GetNoOfNewHiredPerDept(string division, string dept)
        {
            //empHiredDate = "3/15/2016";

            int countNewHired = 0;

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList SettingLst = web.Lists.TryGetList("HrAppraisalSettings");
                        if (SettingLst != null && SettingLst.Items.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(SettingLst.Items[0]["HrApprLastHiringDate"].ToString()))
                            {
                                DateTime lastHiringDate =
                                    Convert.ToDateTime(SettingLst.Items[0]["HrApprLastHiringDate"].ToString());

                                if (division == "All Divisions")
                                {
                                    var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                                                  where emp.START_DATE >= lastHiringDate
                                                  select emp;

                                    if (empData.Any())
                                    {
                                        countNewHired = empData.Count();
                                    }
                                }
                                else
                                {
                                    if (dept == "All Department")
                                    {
                                        var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                                                      where emp.START_DATE >= lastHiringDate && emp.Division == division
                                                      select emp;

                                        if (empData.Any())
                                        {
                                            countNewHired = empData.Count();
                                        }
                                    }
                                    else
                                    {
                                        if (division == string.Empty)
                                        {
                                            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                                                          where emp.START_DATE >= lastHiringDate
                                                                && emp.DEPARTMENT_NAME == dept
                                                                && emp.EMPLOYEE_NUMBER != emp.DEP_MANAGER_NUM
                                                          select emp;

                                            if (empData.Any())
                                            {
                                                countNewHired = empData.Count();
                                            }
                                        }
                                        else
                                        {
                                            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                                                          where emp.START_DATE >= lastHiringDate
                                                                && emp.DEPARTMENT_NAME == dept && emp.Division == division
                                                                && emp.EMPLOYEE_NUMBER != emp.DEP_MANAGER_NUM
                                                          select emp;

                                            if (empData.Any())
                                            {
                                                countNewHired = empData.Count();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                });
            return countNewHired;
        }

        public static int GetNoOfNewHiredPerDivision(string division)
        {
            //empHiredDate = "3/15/2016";

            int countNewHired = 0;

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList SettingLst = web.Lists.TryGetList("HrAppraisalSettings");
                        if (SettingLst != null && SettingLst.Items.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(SettingLst.Items[0]["HrApprLastHiringDate"].ToString()))
                            {
                                DateTime lastHiringDate =
                                    Convert.ToDateTime(SettingLst.Items[0]["HrApprLastHiringDate"].ToString());
                                if (division == "All Department")
                                {
                                    var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                                                  where emp.START_DATE >= lastHiringDate
                                                  select emp;

                                    if (empData.Any())
                                    {
                                        countNewHired = empData.Count();
                                    }
                                }

                                else
                                {
                                    var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                                                  where emp.START_DATE >= lastHiringDate
                                                        && emp.Division == division  //&& emp.DEPARTMENT_NAME == dept
                                                        && emp.EMPLOYEE_NUMBER != emp.DEP_MANAGER_NUM
                                                  select emp;

                                    if (empData.Any())
                                    {
                                        countNewHired = empData.Count();
                                    }
                                }
                            }
                        }
                    }
                });
            return countNewHired;
        }

        public static bool CheckIfEmpAppraisalIsClosed()
        {
            bool isClosed = false;
            string formatedDate = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day;

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList SettingLst = web.Lists.TryGetList("HrAppraisalSettings");
                        if (SettingLst != null && SettingLst.Items.Count > 0)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = string.Format(
                                @"   <Where>
                                      <Leq>
                                         <FieldRef Name='HrApprEndDate' />
                                         <Value IncludeTimeValue='FALSE' Type='DateTime'>{0}</Value>
                                      </Leq>
                                   </Where>"
                                , formatedDate);

                            SPListItemCollection collItems = SettingLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                isClosed = true;
                            }
                        }
                    }
                });

            return isClosed;
        }

        //        public static bool CheckIfEmpIsNewHired(string empCode)
        //        {
        //            DateTime empHiredDate;
        //            string formatedHiredDate = string.Empty;
        //            var empData = from emp in EmpDB.NS_EMPLOYEE_INFO_TNAs
        //                          where emp.EMPLOYEE_NUMBER == empCode
        //                          select emp.START_DATE;

        //            if (empData.Any())
        //            {
        //                empHiredDate = empData.FirstOrDefault();
        //                formatedHiredDate = empHiredDate.Year + "-" + empHiredDate.Month + "-" + empHiredDate.Day;
        //            }
        //            //*********

        //            // empHiredDate = "04/01/2017";

        //            bool newHired = false;

        //            SPSecurity.RunWithElevatedPrivileges(
        //                delegate()
        //                {
        //                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
        //                    using (SPWeb web = site.OpenWeb())
        //                    {
        //                        SPList SettingLst = web.Lists.TryGetList("HrAppraisalSettings");
        //                        if (SettingLst != null && SettingLst.Items.Count > 0)
        //                        {
        //                            SPQuery qry = new SPQuery();

        //                            qry.Query = string.Format(

        //                                @" 
        //                                    <Where>
        //                                      <Lt>
        //                                         <FieldRef Name='HrApprLastHiringDate' />
        //                                         <Value IncludeTimeValue='FALSE' Type='DateTime'>{0}</Value>
        //                                      </Lt>
        //                                   </Where>
        //                            "
        //                                , formatedHiredDate);

        //                            SPListItemCollection collItems = SettingLst.GetItems(qry);
        //                            if (collItems.Count > 0)
        //                            {
        //                                newHired = true;
        //                            }

        //                        }
        //                    }
        //                });
        //            return newHired;
        //        }

        public static bool CheckIfEmpAppraisalInDHFinalApproval(string apprId)
        {
            bool isFinalApproval = false;
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList HrAppraisalValuesLst = web.Lists.TryGetList("HrAppraisalValues");
                        if (HrAppraisalValuesLst != null)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = string.Format(
                                @" 
                                <Where>
                                      <And>
                                         <Eq>
                                            <FieldRef Name='ID' />
                                            <Value Type='Counter'>{0}</Value>
                                         </Eq>
                                         <Or>
                                            <Eq>
                                               <FieldRef Name='AppraisalStatus' />
                                               <Value Type='Text'>Accepted By Employee</Value>
                                            </Eq>
                                            <Or>
                                               <Eq>
                                                  <FieldRef Name='AppraisalStatus' />
                                                  <Value Type='Text'>Not Accepted By Employee</Value>
                                               </Eq>
                                               <Eq>
                                                  <FieldRef Name='AppraisalStatus' />
                                                  <Value Type='Text'>Waiting For DH Final Approval</Value>
                                               </Eq>
                                            </Or>
                                         </Or>
                                      </And>
                                   </Where>
                                       ", apprId);

                            SPListItemCollection collItems = HrAppraisalValuesLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                isFinalApproval = true;
                            }
                        }
                    }
                });
            return isFinalApproval;
        }

        public static string GetAppraisalLastHireDt()
        {
            string lastHiredDt = string.Empty;
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList SettingLst = web.Lists.TryGetList("HrAppraisalSettings");
                        if (SettingLst != null && SettingLst.Items.Count > 0)
                        {
                            foreach (SPListItem item in SettingLst.Items)
                            {
                                if (item["HrApprLastHiringDate"] != null)
                                    lastHiredDt = item["HrApprLastHiringDate"].ToString();
                            }
                        }
                    }
                });
            return lastHiredDt;
        }

        public static string GetEmpPositionByEmpCode(string empCode)
        {
            string empPos = string.Empty;

            #region Get UserMail from DB

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.EMPLOYEE_NUMBER == empCode
                      select em;
            if (emp.Any())
            {
                empPos = emp.FirstOrDefault().POSITION;
            }
            return empPos;

            #endregion

        }

        public static string GetEmpOrgNameByEmpCode(string empCode)
        {
            string empOrgName = string.Empty;

            #region Get UserMail from DB

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.EMPLOYEE_NUMBER == empCode
                      select em;
            if (emp.Any())
            {
                empOrgName = emp.FirstOrDefault().ORGANIZATION_NAME;
            }
            return empOrgName;

            #endregion
        }

        public static string GetEmpNameByEmpCode(string empCode)
        {
            string empOrgName = string.Empty;

            #region Get UserMail from DB

            var emp = from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                      where em.EMPLOYEE_NUMBER == empCode
                      select em;
            if (emp.Any())
            {
                empOrgName = emp.FirstOrDefault().FULL_NAME;
            }
            return empOrgName;

            #endregion
        }

        public static bool CheckDHHasMoreThanDept(string DHCode)
        {
            bool hasMoreThanDept = false;

            var emps = (from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                        where em.DEP_MANAGER_NUM == DHCode
                        select em).ToList();

            DataView view = new DataView(emps.ToDataTable());
            DataTable distinctDHDepts = view.ToTable(true, "Department");

            if (distinctDHDepts != null && distinctDHDepts.Rows.Count > 1)
            {
                hasMoreThanDept = true;
            }
            return hasMoreThanDept;
        }

        public static DataTable GetDHDeptNames(string DHCode)
        {
            DataTable dt_DHOrgs = new DataTable();
            var emps = (from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                        where em.DEP_MANAGER_NUM == DHCode
                        select em).ToList();

            DataView view = new DataView(emps.ToDataTable());
            dt_DHOrgs = view.ToTable(true, "Department");

            return dt_DHOrgs;
        }

        public static bool CheckIfDeptIsSeperatedToSmallDepts(string DHCode)
        {
            bool isSeperated = false;

            var empData = (from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                           where emp.DEP_MANAGER_NUM == DHCode
                           select emp).ToList();

            if (empData.Any())
            {
                DataView view = new DataView(empData.ToDataTable());
                DataTable distinctSmallDepts = view.ToTable(true, "Department");

                if (distinctSmallDepts != null && distinctSmallDepts.Rows.Count > 0)
                {
                    if (distinctSmallDepts.Rows.Count > 1)
                        isSeperated = true;
                }
            }

            return isSeperated;
        }

        public static bool CheckIfDeptIsSeperated(string dept)
        {
            bool isSeperated = false;

            var empData = (from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                           where emp.DEPARTMENT_NAME == dept
                           select emp).ToList();

            if (empData.Any())
            {
                DataView view = new DataView(empData.ToDataTable());
                DataTable distinctSmallDepts = view.ToTable(true, "Department");

                if (distinctSmallDepts != null && distinctSmallDepts.Rows.Count > 0)
                {
                    if (distinctSmallDepts.Rows.Count > 1)
                        isSeperated = true;
                }
            }

            return isSeperated;
        }

        //        public static bool CheckIfDeptIsSeperatedToOrgs(string dept)
        //        {
        //            bool isSeperated = false;

        //            SPSecurity.RunWithElevatedPrivileges(
        //                delegate()
        //                {
        //                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
        //                    using (SPWeb web = site.OpenWeb())
        //                    {
        //                        SPList DeptSeperatedToOrgsLst = web.Lists.TryGetList("DeptSeperatedToOrgs");
        //                        if (DeptSeperatedToOrgsLst != null && DeptSeperatedToOrgsLst.Items.Count > 0)
        //                        {
        //                            SPQuery qry = new SPQuery();

        //                            qry.Query = string.Format(
        //                                @" 
        //                                           <Where>
        //                                              <Eq>
        //                                                 <FieldRef Name='Title' />
        //                                                 <Value Type='Text'>{0}</Value>
        //                                              </Eq>
        //                                           </Where>
        //                                        "
        //                                , dept);

        //                            SPListItemCollection collItems = DeptSeperatedToOrgsLst.GetItems(qry);
        //                            if (collItems.Count > 0)
        //                            {
        //                                isSeperated = true;
        //                            }
        //                        }
        //                    }
        //                });

        //            return isSeperated;
        //        }

        public static bool CheckIfEmpIsConvertedToNonCompUser(string empCode)
        {
            bool isConverted = false;

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList ConvertedToNonCompEmpsLst = web.Lists.TryGetList("ConvertedToNonCompEmps");
                        if (ConvertedToNonCompEmpsLst != null && ConvertedToNonCompEmpsLst.Items.Count > 0)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = string.Format(
                                @" 
                                    <Where>
                                      <Eq>
                                         <FieldRef Name='EmpCode' />
                                         <Value Type='Text'>{0}</Value>
                                      </Eq>
                                   </Where>
                                "
                                , empCode);

                            SPListItemCollection collItems = ConvertedToNonCompEmpsLst.GetItems(qry);
                            if (collItems.Count > 0)
                            {
                                isConverted = true;
                            }
                        }
                    }
                });

            return isConverted;
        }

        public static void UpdateTestLst(string value)
        {
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList list = web.Lists.TryGetList("NewTestLst");
                        if (list != null)
                        {
                            web.AllowUnsafeUpdates = true;
                            SPListItem item = list.Items.Add();
                            item["Value"] = value;
                            item.Update();
                            web.AllowUnsafeUpdates = false;
                        }
                    }
                });
        }

        public static bool CheckIfDHHaveMoreThanDept(string DHCode)
        {
            bool haveMoreDept = false;
            DataTable dt_DHOrgs = new DataTable();
            var emps = (from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                        where em.DEP_MANAGER_NUM == DHCode
                        select em).ToList();

            DataView view = new DataView(emps.ToDataTable());
            //dt_DHOrgs = view.ToTable(true, "DEPARTMENT_NAME");
            dt_DHOrgs = view.ToTable(true, "Department");
            if (dt_DHOrgs != null && dt_DHOrgs.Rows.Count > 1)
                haveMoreDept = true;

            return haveMoreDept;
        }

        public static DataTable GetDHDeptsNames(string DHCode)
        {
            DataTable dt_DHDepts = new DataTable();
            var emps = (from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                        where em.DEP_MANAGER_NUM == DHCode
                        select em).ToList();

            DataView view = new DataView(emps.ToDataTable());
            // dt_DHDepts = view.ToTable(true, "DEPARTMENT_NAME");
            dt_DHDepts = view.ToTable(true, "Department");

            return dt_DHDepts;
        }

        public static string GetAppraisalStatus(string apprId)
        {
            string reqStatus = string.Empty;
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList HrAppraisalValuesLst = web.Lists.TryGetList("HrAppraisalValues");
                        if (HrAppraisalValuesLst != null)
                        {
                            SPListItem empReq = HrAppraisalValuesLst.GetItemById(int.Parse(apprId));
                            if (empReq != null)
                            {
                                reqStatus = empReq["AppraisalStatus"].ToString();
                            }
                        }
                    }
                });
            return reqStatus;
        }

        /* mary  public static DataTable GetEmpsBySupervisorCode(string empCode)
         {
             DataTable dtEmps = null;
             var empData = (from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                            where emp.SUPERVISOR_NUMBER == empCode
                            && emp.EMAIL_ADDRESS != null && emp.EMAIL_ADDRESS != string.Empty
                            select emp).ToList();

             if (empData.Any())
             {
                 dtEmps = empData.ToDataTable();
             }
             return dtEmps;
         }
         */


        public static DataTable GetEmpsBySupervisorCode(string empCode)
        {
            DataTable dtEmps = null;
            var empData = (from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                           where emp.SUPERVISOR_NUMBER == empCode
                           select emp).ToList();

            if (empData.Any())
            {
                dtEmps = empData.ToDataTable();
            }
            return dtEmps;
        }


        public static DataTable GetEmpsByTopMgrCode(string empCode)
        {
            DataTable dtEmps = null;
            var empData = (from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                           where emp.SUPERVISOR_NUMBER == empCode
                           || emp.DEP_MANAGER_NUM == empCode
                           || emp.Division_MANAGER_NUM == empCode
                           && (emp.EMAIL_ADDRESS != null && emp.EMAIL_ADDRESS != string.Empty)
                           select emp).ToList();

            if (empData.Any())
            {
                dtEmps = empData.ToDataTable();
            }
            return dtEmps;
        }

        public static bool CheckEmpIsInTopMgr(string empCode)
        {
            bool isTopMgr = false;
            var empData = (from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                           where emp.SUPERVISOR_NUMBER == empCode
                           || emp.DEP_MANAGER_NUM == empCode
                           || emp.Division_MANAGER_NUM == empCode
                           select emp).ToList();

            //  && (emp.EMAIL_ADDRESS != "null" && emp.EMAIL_ADDRESS != string.Empty)
            if (empData.Any())
            {
                isTopMgr = true;
            }
            return isTopMgr;
        }

        public static Boolean SendMail(string mailTo, string cc, string mailTitle, string textToBesend, Guid SiteId, Guid webId)
        {
            //string text = System.IO.File.ReadAllText("/Data"+"/emailTemplate.htm");
            //textToBesend= text.Replace("ReplaceMe", textToBesend);  
            StringDictionary headers = new StringDictionary
                                       {
                                           //{"from", "dalzoubi@the-e.net"},
                                           {"to", mailTo},
                                           {"cc", cc},
                                           {"subject", mailTitle},
                                           {"content-type", "text/html"}
                                       };
            try
            {
                bool flag = false;
                SPSecurity.RunWithElevatedPrivileges(
                  delegate()
                  {
                      using (SPSite site = new SPSite(
                       SiteId))
                      {
                          using (SPWeb web = site.OpenWeb(webId))
                          {
                              flag = SPUtility.SendEmail(web, headers,
                                                         textToBesend);
                          }
                      }
                  });
                return flag;
            }
            catch (System.Exception exp)
            {
                return false;
            }
        }


        public static DataTable GetGeneralCourses(string courseCategory, string courseType)
        {
            DataTable dtCourses = null;
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList GeneralCoursesLst = web.Lists.TryGetList("GeneralCoursesLst");
                        if (GeneralCoursesLst != null && GeneralCoursesLst.Items.Count > 0)
                        {
                            SPQuery qry = new SPQuery();

                            qry.Query = string.Format(
                                @" 
                                                  <Where>
                                                  <And>
                                                     <Eq>
                                                        <FieldRef Name='CourseCategory' />
                                                        <Value Type='Text'>{0}</Value>
                                                     </Eq>
                                                     <Eq>
                                                        <FieldRef Name='CourseType' />
                                                        <Value Type='Text'>{1}</Value>
                                                     </Eq>
                                                  </And>
                                               </Where>", courseCategory, courseType);

                            SPListItemCollection collItems = GeneralCoursesLst.GetItems(qry);
                            if (collItems != null && collItems.Count > 0)
                            {
                                dtCourses = collItems.GetDataTable();
                            }
                        }
                    }
                });
            return dtCourses;
        }


        public static DataTable GetTechnicalCoursesByDeptName(string deptName)
        {
            DataTable dtCourses = null;
            SPSecurity.RunWithElevatedPrivileges(
    delegate()
    {
        using (SPSite site = new SPSite(SPContext.Current.Web.Url))
        using (SPWeb web = site.OpenWeb())
        {
            SPList TechnicalCoursesLst = web.Lists.TryGetList("TechnicalCoursesLst");
            if (TechnicalCoursesLst != null && TechnicalCoursesLst.Items.Count > 0)
            {
                SPQuery qry = new SPQuery();

                qry.Query = string.Format(
                    @" 
                                        <Where>
                                          <Eq>
                                             <FieldRef Name='DeptName' />
                                             <Value Type='Text'>{0}</Value>
                                          </Eq>
                                       </Where>
                                "
                    , deptName);

                SPListItemCollection collItems = TechnicalCoursesLst.GetItems(qry);
                if (collItems != null && collItems.Count > 0)
                {
                    dtCourses = collItems.GetDataTable();
                }
            }
        }
    });
            return dtCourses;
        }

        //mary 
        public static DataTable GetAllDepartments()
        {
            var allDepts = (from depts in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                            where depts.DEPARTMENT_NAME != ""
                            select depts).ToList();

            DataView view = new DataView(allDepts.ToDataTable());


            DataTable distinctDepts = view.ToTable(true, "Department");
            return distinctDepts;
        }

        public static DataTable GetAllDevisions()
        {

            var allDevisions = (from depts in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                                where depts.Division != ""
                                select depts).ToList();

            DataView view = new DataView(allDevisions.ToDataTable());


            DataTable distinctDevisions = view.ToTable(true, "Division");
            return distinctDevisions;
        }

        public static DataTable Get_Devision_Departments(string DevisionName)
        {
            var allDepts = (from depts in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                            where depts.DEPARTMENT_NAME != ""
                            && depts.Division == DevisionName
                            select depts).ToList();

            DataView view = new DataView(allDepts.ToDataTable());


            DataTable distinctDepts = view.ToTable(true, "Department");
            return distinctDepts;
        }


        //public static DataTable Get_Devision_Employees(string DevisionName)
        //{
        //    DataTable distinctDepts = new DataTable();

        //    var emp = (from em in EmpDB.NS_EMPLOYEE_INFO_TNAs
        //               where em.Division == DevisionName
        //               select em).ToList();
        //    if (emp.Any())
        //    {
        //        distinctDepts = emp.ToDataTable();
        //    }
        //    return distinctDepts;

        //}



        public static DataTable GetDepartmentEmployees(string DepartmentName)
        {
            DataTable distinctDepts = new DataTable();

            var emp = (from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                       where em.Department == DepartmentName
                       select em).ToList();
            if (emp.Any())
            {
                distinctDepts = emp.ToDataTable();
            }
            return distinctDepts;
        }

        public static DataTable GetAllEmployees()
        {
            DataTable distinctDepts = new DataTable();

            var emp = (from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                       select em).ToList();
            if (emp.Any())
            {
                distinctDepts = emp.ToDataTable();
            }
            return distinctDepts;
        }

        public static DataTable GetAllEmployeesEmails()
        {
            DataTable empsData = new DataTable();
            DataTable Mgrs = new DataTable();
            string empCode = string.Empty;
            //&& em.DEPARTMENT_NAME == "HR"

            var emp = (from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                       where em.EMAIL_ADDRESS != "null" && em.EMAIL_ADDRESS != "NULL" && em.EMAIL_ADDRESS != null && em.EMAIL_ADDRESS != string.Empty
                       select em).ToList();

            if (emp.Any())
            {
                empsData = emp.ToDataTable();
            }
            Mgrs = empsData.Clone();

            foreach (DataRow row in empsData.Rows)
            {
                if (CheckEmpIsInTopMgr(row["EMPLOYEE_NUMBER"].ToString()))
                {
                    Mgrs.ImportRow(row);
                }
            }

            return Mgrs;
        }

        public static DataTable RemoveDuplicates(DataTable dt, string comparing)
        {

            if (dt.Rows.Count > 0)
            {
                for (int i = dt.Rows.Count - 1; i >= 0; i--)
                {
                    if (i == 0)
                    {
                        break;
                    }
                    for (int j = i - 1; j >= 0; j--)
                    {
                        if (dt.Rows[i][comparing].ToString() == dt.Rows[j][comparing].ToString())
                        {
                            dt.Rows[i].Delete();
                            break;
                        }
                    }
                }
                dt.AcceptChanges();
            }
            return dt;
        }

        public static string GetFromatDateString(string _datepicker)
        {
            int _day = 0;
            int _month = 0;
            int _year = 0;
            try
            {
                if (_datepicker != "")
                {
                    string[] start = _datepicker.Split('/');

                    _year = int.Parse(start[0].ToString());
                    _month = int.Parse(start[1].ToString());
                    _day = int.Parse(start[2].ToString());

                }
            }
            catch (Exception ex)
            {
                // Log the exception
                // ErrorLogger.Log(ex);
            }
            return _year + "-" + _month + "-" + _day;
        }

        public static string GetFromatDateString(DateTime _datepicker)
        {
            int _day = 0;
            int _month = 0;
            int _year = 0;
            try
            {
                if (_datepicker != null)
                {
                    _year = _datepicker.Year;
                    _month = _datepicker.Month;
                    _day = _datepicker.Day;
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                // ErrorLogger.Log(ex);
            }
            return _year + "-" + _month + "-" + _day;
        }
        public static DataTable Get_Devision_Employees(string DevisionName)
        {
            DataTable distinctDepts = new DataTable();

            var emp = (from em in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                       where em.Division == DevisionName
                       select em).ToList();
            if (emp.Any())
            {
                distinctDepts = emp.ToDataTable();
            }
            return distinctDepts;

        }

        public static string formatEmpCode(string empCode)
        {
            string formatedEmpCode = string.Empty;
            if (empCode.Length == 1)
            {
                formatedEmpCode = "0000" + empCode;
            }
            else if (empCode.Length == 2)
            {
                formatedEmpCode = "000" + empCode;
            }
            else if (empCode.Length == 3)
            {
                formatedEmpCode = "00" + empCode;
            }
            else if (empCode.Length == 4)
            {
                formatedEmpCode = "0" + empCode;
            }
            else
            {
                formatedEmpCode = empCode;
            }

            return formatedEmpCode;
        }

        public static string Departmentheadmail(string DepartmentName)
        {
            var Mail = (from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                        where emp.DEPARTMENT_NAME == DepartmentName
                        select emp.DEP_MANAGER_MAIL).ToList();

            return (ToDataTable(Mail)).Rows[0]["DEP_MANAGER_MAIL"].ToString();
        }

        public static string GetEmpNameByCode(string EmpCode)
        {
            string EmpName = String.Empty;
            var employee = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                           where emp.EMPLOYEE_NUMBER == EmpCode
                           select emp;

            if (employee != null)
            {
                if (!string.IsNullOrEmpty(employee.FirstOrDefault().FULL_NAME))
                    EmpName = employee.FirstOrDefault().FULL_NAME;
            }
            return EmpName;
        }

        public static bool CheckEmpCodesExist(string EmpCode)
        {

            bool isExist = false;

            var employee = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                           where emp.EMPLOYEE_NUMBER == EmpCode
                           select emp;

            if (employee.Count() > 0)
            {
                isExist = true;
            }
            return isExist;
        }

        public static string GetEmpPossitionByCode(string EmpCode)
        {
            string EmpName = String.Empty;
            var employee = from emp in EmpDB.NS_EMPLOYEE_INFO_Last_2024s
                           where emp.EMPLOYEE_NUMBER == EmpCode
                           select emp;

            if (employee != null)
            {
                EmpName = employee.FirstOrDefault().POSITION;
            }
            return EmpName;
        }
    }
}

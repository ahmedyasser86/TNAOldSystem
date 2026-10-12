using System;
using System.ComponentModel;
using System.Web;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using NileSugar.HrTNA.Webparts.TrainingRequests_WP;

namespace NileSugar.HrTNA.Webparts.ManagerRequests_WP
{
    [ToolboxItemAttribute(false)]
    public partial class ManagerRequests_WP : WebPart
    {
        // Uncomment the following SecurityPermission attribute only when doing Performance Profiling on a farm solution
        // using the Instrumentation method, and then remove the SecurityPermission attribute when the code is ready
        // for production. Because the SecurityPermission attribute bypasses the security check for callers of
        // your constructor, it's not recommended for production purposes.
        // [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Assert, UnmanagedCode = true)]
        public ManagerRequests_WP()
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
                HttpContext.Current.Response.Redirect("http://sharepoint.ns.local/forms/hr/TNA/SitePages/ManagersRequests.aspx");


            if (!Page.IsPostBack)
            {
                string empId = CommonMethods.GetLoginEmpCode(); // "01125"; //
                bool isMgr = CommonMethods.CheckIfLoginEmpIsMgr(empId);
                if (isMgr || empId == "01071")
                {
                    BindGV_SupervCourses(empId, CommonMethods.GetLookups("TNAYear"));
                    BindGV_MgrCourses(empId, CommonMethods.GetLookups("TNAYear"));
                    BindGV_DHCourses(empId, CommonMethods.GetLookups("TNAYear"));
                    BindGV_DVMgrCourses(empId, CommonMethods.GetLookups("TNAYear"));

                    if (div_SupervisorReqs.Visible == false && div_MgrReqs.Visible == false &&
                        div_DHReqs.Visible == false && div_DVMgrReqs.Visible == false)
                    {
                        div_noCourses.Visible = true;
                        div_bgLarge.Visible = true;
                    }
                    else
                    {
                        div_noCourses.Visible = false;
                        div_bgLarge.Visible = false;
                    }
                }
                else
                {
                    //Redirect to Acees denied Page
                    string accessDeniedPage = CommonMethods.GetLookups("AccessDeniedPage");
                    HttpContext.Current.Response.Redirect(accessDeniedPage);
                }
            }
        }


        private void BindGV_SupervCourses(string supervCode, string currentYear)
        {
            SPSecurity.RunWithElevatedPrivileges(
delegate()
{
    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
    using (SPWeb web = site.OpenWeb())
    {
        SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
        if (EmpCoursesLst != null)
        {
            SPQuery qry = new SPQuery();

            qry.Query = String.Format(
                //                @" 
                //                                  <Where>
                //                                      <And>
                //                                         <Eq>
                //                                            <FieldRef Name='SupervisorCode' />
                //                                            <Value Type='Text'>{0}</Value>
                //                                         </Eq>
                //                                         <Neq>
                //                                            <FieldRef Name='MgrCode' />
                //                                            <Value Type='Text'>{0}</Value>
                //                                         </Neq>
                //                                      </And>
                //                                   </Where>
                //                                       "

                            @"   <Where>
                                  <And>
                                     <Eq>
                                        <FieldRef Name='SupervisorCode' />
                                        <Value Type='Text'>{0}</Value>
                                     </Eq>
                                     <And>
                                        <Neq>
                                           <FieldRef Name='MgrCode' />
                                           <Value Type='Text'>{0}</Value>
                                        </Neq>
                                        <Eq>
                                           <FieldRef Name='TNAYear' />
                                           <Value Type='Text'>{1}</Value>
                                        </Eq>
                                     </And>
                                  </And>
                               </Where>", supervCode, currentYear);

            SPListItemCollection collItems = EmpCoursesLst.GetItems(qry);
            if (collItems.Count > 0)
            {
                GV_SupervisorCourses.DataSource = collItems.GetDataTable();
                GV_SupervisorCourses.DataBind();
            }
            else
            {
                GV_SupervisorCourses.DataSource = null;
                GV_SupervisorCourses.DataBind();

                div_SupervisorReqs.Visible = false;
            }
        }
    }
});
        }

        private void BindGV_MgrCourses(string mgrCode, string currentYear)
        {
            SPSecurity.RunWithElevatedPrivileges(
   delegate()
   {
       using (SPSite site = new SPSite(SPContext.Current.Web.Url))
       using (SPWeb web = site.OpenWeb())
       {
           SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
           if (EmpCoursesLst != null)
           {
               SPQuery qry = new SPQuery();

               qry.Query = String.Format(
                   //                   @" 
                   //                               <Where>
                   //                                  <Eq>
                   //                                     <FieldRef Name='MgrCode' />
                   //                                     <Value Type='Text'>{0}</Value>
                   //                                  </Eq>
                   //                               </Where>
                   //                                       "

                                        @"   <Where>
                                              <And>
                                                 <Eq>
                                                    <FieldRef Name='MgrCode' />
                                                    <Value Type='Text'>{0}</Value>
                                                 </Eq>
                                                 <Eq>
                                                    <FieldRef Name='TNAYear' />
                                                    <Value Type='Text'>{1}</Value>
                                                 </Eq>
                                              </And>
                                           </Where>", mgrCode, currentYear);

               SPListItemCollection collItems = EmpCoursesLst.GetItems(qry);
               if (collItems.Count > 0)
               {
                   GV_MgrCourses.DataSource = collItems.GetDataTable();
                   GV_MgrCourses.DataBind();
               }
               else
               {
                   GV_MgrCourses.DataSource = null;
                   GV_MgrCourses.DataBind();

                   div_MgrReqs.Visible = false;
               }
           }
       }
   });
        }

        private void BindGV_DHCourses(string DHCode, string currentYear)
        {
            SPSecurity.RunWithElevatedPrivileges(
delegate()
{
    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
    using (SPWeb web = site.OpenWeb())
    {
        SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
        if (EmpCoursesLst != null)
        {
            SPQuery qry = new SPQuery();

            qry.Query = String.Format(
                //                @" 
                //                               <Where>
                //                                  <Eq>
                //                                     <FieldRef Name='DHCode' />
                //                                     <Value Type='Text'>{0}</Value>
                //                                  </Eq>
                //                               </Where>
                //                                       "

            @"   <Where>
                  <And>
                     <Eq>
                        <FieldRef Name='DHCode' />
                        <Value Type='Text'>{0}</Value>
                     </Eq>
                     <Eq>
                        <FieldRef Name='TNAYear' />
                        <Value Type='Text'>{1}</Value>
                     </Eq>
                  </And>
               </Where>"
                , DHCode, currentYear);


            SPListItemCollection collItems = EmpCoursesLst.GetItems(qry);
            if (collItems.Count > 0)
            {
                GV_DHCourses.DataSource = collItems.GetDataTable();
                GV_DHCourses.DataBind();
            }
            else
            {
                GV_DHCourses.DataSource = null;
                GV_DHCourses.DataBind();

                div_DHReqs.Visible = false;
            }
        }
    }
});
        }

        private void BindGV_DVMgrCourses(string DVMgrCode, string currentYear)
        {
            SPSecurity.RunWithElevatedPrivileges(
delegate()
{
    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
    using (SPWeb web = site.OpenWeb())
    {
        SPList EmpCoursesLst = web.Lists.TryGetList("EmpCourses");
        if (EmpCoursesLst != null)
        {
            SPQuery qry = new SPQuery();

            qry.Query = String.Format(
                //                @" 
                //                               <Where>
                //                                  <Eq>
                //                                     <FieldRef Name='DVMgrCode' />
                //                                     <Value Type='Text'>{0}</Value>
                //                                  </Eq>
                //                               </Where>
                //                                       "

@"   <Where>
      <And>
         <Eq>
            <FieldRef Name='DVMgrCode' />
            <Value Type='Text'>{0}</Value>
         </Eq>
         <Eq>
            <FieldRef Name='TNAYear' />
            <Value Type='Text'>{1}</Value>
         </Eq>
      </And>
   </Where>"
                , DVMgrCode, currentYear);

            SPListItemCollection collItems = EmpCoursesLst.GetItems(qry);
            if (collItems.Count > 0)
            {
                GV_DVMgrCourses.DataSource = collItems.GetDataTable();
                GV_DVMgrCourses.DataBind();
            }
            else
            {
                GV_DVMgrCourses.DataSource = null;
                GV_DVMgrCourses.DataBind();

                div_DVMgrReqs.Visible = false;
            }
        }
    }
});
        }

        protected void lnkBtnDetailsSuperviors_OnClick(object sender, EventArgs e)
        {
            string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
            var DetailsBtnLink = (LinkButton)sender;
            var row = (GridViewRow)DetailsBtnLink.NamingContainer;
            var courseId = (Label)row.FindControl("lbl_CourseID");

            string supervisorCode = ((Label)row.FindControl("lbl_SupervisorCode")).Text.Trim();
            string MgrCode = ((Label)row.FindControl("lbl_MgrCode")).Text.Trim();
            string DHCode = ((Label)row.FindControl("lbl_DHCode")).Text.Trim();
            string DVMgrCode = ((Label)row.FindControl("lbl_DVMgrCode")).Text.Trim();
            string TNAReqPage = CommonMethods.GetLookups("TNAReqPage");
            string reqStatus = ((Label)row.FindControl("lbl_Req_Status")).Text.Trim();

            if (!string.IsNullOrEmpty(supervisorCode) && !string.IsNullOrEmpty(MgrCode)
              && string.IsNullOrEmpty(DHCode) && string.IsNullOrEmpty(DVMgrCode))
            {

                if (reqStatus == "Rejected By Manger" || reqStatus == "Requested By Supervisor")
                {
                    HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                          "&pageTo=supervisor");
                }
                else // if status ->Sent To HR
                {
                    HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                        "&pageTo=hr");
                }
            }
            else
            {
                HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                    "&pageTo=hr");
            }
        }

        protected void lnkBtnDetailsMgrs_OnClick(object sender, EventArgs e)
        {
            string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
            var DetailsBtnLink = (LinkButton)sender;
            var row = (GridViewRow)DetailsBtnLink.NamingContainer;
            var courseId = (Label)row.FindControl("lbl_CourseID");

            string supervisorCode = ((Label)row.FindControl("lbl_SupervisorCode")).Text.Trim();
            string MgrCode = ((Label)row.FindControl("lbl_MgrCode")).Text.Trim();
            string DHCode = ((Label)row.FindControl("lbl_DHCode")).Text.Trim();
            string DVMgrCode = ((Label)row.FindControl("lbl_DVMgrCode")).Text.Trim();
            string TNAReqPage = CommonMethods.GetLookups("TNAReqPage");
            string reqStatus = ((Label)row.FindControl("lbl_Req_Status")).Text.Trim();

            if (!string.IsNullOrEmpty(supervisorCode) && !string.IsNullOrEmpty(MgrCode)
              && string.IsNullOrEmpty(DHCode) && string.IsNullOrEmpty(DVMgrCode))
            {

                if (reqStatus == "Requested By Supervisor")
                {
                    HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                          "&pageTo=mgr");
                }
                else
                {
                    HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                        "&pageTo=hr");

                }
            }
            else
            {
                HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                    "&pageTo=hr");

            }
        }

        protected void lnkBtnDetailsDH_OnClick(object sender, EventArgs e)
        {
            string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
            var DetailsBtnLink = (LinkButton)sender;
            var row = (GridViewRow)DetailsBtnLink.NamingContainer;
            var courseId = (Label)row.FindControl("lbl_CourseID");

            string supervisorCode = ((Label)row.FindControl("lbl_SupervisorCode")).Text.Trim();
            string MgrCode = ((Label)row.FindControl("lbl_MgrCode")).Text.Trim();
            string DHCode = ((Label)row.FindControl("lbl_DHCode")).Text.Trim();
            string DVMgrCode = ((Label)row.FindControl("lbl_DVMgrCode")).Text.Trim();
            string TNAReqPage = CommonMethods.GetLookups("TNAReqPage");
            string reqStatus = ((Label)row.FindControl("lbl_Req_Status")).Text.Trim();

            if (!string.IsNullOrEmpty(supervisorCode) && !string.IsNullOrEmpty(MgrCode)
              && !string.IsNullOrEmpty(DHCode) && string.IsNullOrEmpty(DVMgrCode))
            {

                if (reqStatus == "Approved By Manger" || reqStatus == "Rejected By Manger")
                {
                    HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                          "&pageTo=DH");
                    return;
                }
                else
                {
                    HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                        "&pageTo=hr");
                    return;
                }
            }
            else
            {
                HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                    "&pageTo=hr");
                return;
            }
        }

        protected void lnkBtnDetailsDVMgr_OnClick(object sender, EventArgs e)
        {
            string currentAppUrl = CommonMethods.GetLookups("currentAppUrl");
            var DetailsBtnLink = (LinkButton)sender;
            var row = (GridViewRow)DetailsBtnLink.NamingContainer;
            var courseId = (Label)row.FindControl("lbl_CourseID");

            string supervisorCode = ((Label)row.FindControl("lbl_SupervisorCode")).Text.Trim();
            string MgrCode = ((Label)row.FindControl("lbl_MgrCode")).Text.Trim();
            string DHCode = ((Label)row.FindControl("lbl_DHCode")).Text.Trim();
            string DVMgrCode = ((Label)row.FindControl("lbl_DVMgrCode")).Text.Trim();
            string TNAReqPage = CommonMethods.GetLookups("TNAReqPage");
            string reqStatus = ((Label)row.FindControl("lbl_Req_Status")).Text.Trim();

            if (!string.IsNullOrEmpty(supervisorCode) && !string.IsNullOrEmpty(MgrCode)
              && !string.IsNullOrEmpty(DHCode) && !string.IsNullOrEmpty(DVMgrCode))
            {
                if (reqStatus == "Approved By Dept. Head" || reqStatus == "Rejected By Dept. Head" || reqStatus == "Rejected By Division Manger")
                {
                    HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                          "&pageTo=DVMgr");
                    return;
                }
                else
                {
                    HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                        "&pageTo=hr");
                    return;
                }
            }
            else
            {
                HttpContext.Current.Response.Redirect(currentAppUrl + TNAReqPage + "?TNAId=" + courseId.Text +
                                                    "&pageTo=hr");
                return;
            }

        }
    }
}

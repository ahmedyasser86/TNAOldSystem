using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using NileSugar.HrTNA.Webparts.TrainingRequests_WP;

namespace NileSugar.HrTNA.Webparts.Update_WP
{
    [ToolboxItemAttribute(false)]
    public partial class Update_WP : WebPart
    {
        // Uncomment the following SecurityPermission attribute only when doing Performance Profiling on a farm solution
        // using the Instrumentation method, and then remove the SecurityPermission attribute when the code is ready
        // for production. Because the SecurityPermission attribute bypasses the security check for callers of
        // your constructor, it's not recommended for production purposes.
        // [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Assert, UnmanagedCode = true)]
        public Update_WP()
        {
        }

        protected override void OnInit(EventArgs e)
        {
            base.OnInit(e);
            InitializeControl();
        }

        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void btn_UpdateEmpTechnicalCoursesLst_Click(object sender, EventArgs e)
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

                         SPList EmpTechnicalCoursesLst = web.Lists.TryGetList("EmpTechnicalCoursesLst");
                         if (EmpTechnicalCoursesLst != null)
                         {
                             SPListItemCollection coll = EmpTechnicalCoursesLst.GetItems();

                             foreach (SPListItem item in coll)
                             {
                                 item["TNAYear"] = CommonMethods.GetLookups("TNAYear");
                                 item.Update();
                             }
                         }

                         lblMsg.Text = "Data Was Updated Successfully";
                         lblMsg.ForeColor = Color.Blue;

                         web.AllowUnsafeUpdates = false;
                     }
                 });

            }
            catch (Exception)
            {
                lblMsg.Text = "Error In Updating";
                lblMsg.ForeColor = Color.Red;
            }
        }

        protected void btn_updateEmpDevSkillsOtherCourses_Click(object sender, EventArgs e)
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

                         SPList EmpDevSkillsOtherCourses = web.Lists.TryGetList("EmpDevSkillsOtherCourses");
                         if (EmpDevSkillsOtherCourses != null)
                         {
                             SPListItemCollection coll = EmpDevSkillsOtherCourses.GetItems();

                             foreach (SPListItem item in coll)
                             {
                                 item["TNAYear"] = CommonMethods.GetLookups("TNAYear");
                                 item.Update();
                             }
                         }

                         lblMsg.Text = "Data Was Updated Successfully";
                         lblMsg.ForeColor = Color.Blue;

                         web.AllowUnsafeUpdates = false;
                     }
                 });

            }
            catch (Exception)
            {
                lblMsg.Text = "Error In Updating";
                lblMsg.ForeColor = Color.Red;
            }
        }
    }
}

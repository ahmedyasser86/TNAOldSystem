using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Web.UI;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using NileSugar.HrTNA.Webparts.TrainingRequests_WP;
using System.Web.UI.WebControls;
using System.Web;

namespace NileSugar.HrTNA.Webparts.TNASettings_WP
{
    [ToolboxItemAttribute(false)]
    public partial class TNASettings_WP : WebPart
    {
        // Uncomment the following SecurityPermission attribute only when doing Performance Profiling on a farm solution
        // using the Instrumentation method, and then remove the SecurityPermission attribute when the code is ready
        // for production. Because the SecurityPermission attribute bypasses the security check for callers of
        // your constructor, it's not recommended for production purposes.
        // [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Assert, UnmanagedCode = true)]
        public TNASettings_WP()
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
                HttpContext.Current.Response.Redirect("http://sharepoint.ns.local/forms/hr/TNA/SitePages/Settings.aspx");


            if (!Page.IsPostBack)
            {
                drpLstTNAYear.Items.Insert(0, new ListItem(CommonMethods.GetLookups("TNAYear"), CommonMethods.GetLookups("TNAYear")));
            }
        }


        protected void btnSendEmail_OnClick(object sender, EventArgs e)
        {
            try
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                using (SPWeb web = site.OpenWeb())
                {
                    Guid SiteId = site.ID;
                    Guid WebId = web.ID;
                    DataTable allEmps_dt = CommonMethods.GetAllEmployeesEmails();
                    lblAllEmps.Text = allEmps_dt.Rows.Count.ToString();

                    Hashtable MailTemplates = CommonMethods.LoadMailTemplates();
                    string TNAYear = CommonMethods.GetLookups("TNAYear").ToString();
                    bool sent = false;

                    //dev
                    // string AttachmentURL = "http://nshosplab01:45238/TNA/SiteAssets/UserManual/دليل المستخدم - الاحتياجات التدريبيه.pdf";


                    //life
                    string AttachmentURL = "http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/UserManual/دليل المستخدم - الاحتياجات التدريبيه.pdf";


                    string MailSubject = " الاحتياجات التدريبية لعام " + CommonMethods.ConvertToArabicNumerals(TNAYear);

                    if (allEmps_dt != null && allEmps_dt.Rows.Count > 0)
                    {
                        if (MailTemplates["StartTNATemplate"] != null)
                        {
                            foreach (DataRow row in allEmps_dt.Rows)
                            {
                                string empCode = row["EMPLOYEE_NUMBER"].ToString().Trim();
                                string empEmail = row["EMAIL_ADDRESS"].ToString().Trim();

                                // "rgamal@nile-sugar.com"; // "test01@nslab.lab"; //
                      
                                //"rgamal@nile-sugar.com"; //

                                if (!string.IsNullOrEmpty(empEmail))
                                {
                                    //if (CommonMethods.CheckEmpIsInTopMgr(empCode))
                                    //{
                                        if (CheckMailInList(empEmail.Trim()))
                                        {
                                        }
                                        else
                                        {
                                            // Send TNA Start Email

                                            if (SendMailWithAttachment(
                                                empEmail,
                                                MailSubject,
                                                string.Format(MailTemplates["StartTNATemplate"].ToString()),
                                                AttachmentURL,
                                                "دليل المستخدم - الاحتياجات التدريبيه.pdf"))
                                            {
                                                sent = true;
                                            }
                                            else
                                            {
                                                UpdateSuccededLst("Not Sent To " + empEmail.Trim());
                                            }
                                        }
                                    //}
                                }
                            }

                            if (sent)
                            {
                                ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                                    "Success Message",
                                    "alert('Emails Was Sent Sucessfully')", true);
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                       "Error Message",
                       "alert('Error in Sending Emails')", true);
                //mary
                UpdateTestLst(ex.ToString());

            }
        }

        public void SentMailToMgrs()
        {
            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
            using (SPWeb web = site.OpenWeb())
            {
                Guid SiteId = site.ID;
                Guid WebId = web.ID;
                DataTable allEmps_dt = CommonMethods.GetAllEmployeesEmails();
                lblAllEmps.Text = allEmps_dt.Rows.Count.ToString();
                Hashtable MailTemplates = CommonMethods.LoadMailTemplates();
                string TNAYear = CommonMethods.GetLookups("TNAYear").ToString();
                bool sent = false;

                //dev
                // string AttachmentURL = "http://nshosplab01:45238/TNA/SiteAssets/UserManual/دليل المستخدم - الاحتياجات التدريبيه.pdf";


                //life
               // string AttachmentURL = "http://sharepoint.ns.local/forms/hr/TNA/SiteAssets/UserManual/دليل المستخدم - الاحتياجات التدريبيه.pdf";


                string MailSubject = " الاحتياجات التدريبية لعام " + CommonMethods.ConvertToArabicNumerals(TNAYear);

                if (allEmps_dt != null && allEmps_dt.Rows.Count > 0)
                {
                    if (MailTemplates["StartTNATemplate"] != null)
                    {
                        foreach (DataRow row in allEmps_dt.Rows)
                        {
                            string empCode = row["EMPLOYEE_NUMBER"].ToString().Trim();
                            string empEmail = row["EMAIL_ADDRESS"].ToString().Trim();

                            // "rgamal@nile-sugar.com"; // "test01@nslab.lab"; //
                            // "mgeady@nile-sugar.com";
                            // "mgeady@nslab.lab";
                            // "Administrator@nslab.lab"; 
                            //"rgamal@nile-sugar.com"; //

                            if (!string.IsNullOrEmpty(empEmail))
                            {
                                //if (CommonMethods.CheckEmpIsInTopMgr(empCode))
                                //{
                                    if (CheckMailInList(empEmail.Trim()))
                                    {
                                    }
                                    else
                                    {
                                        // Write Emails in Test List 
                                        UpdateSuccededLst(empEmail.Trim());

                                        //***************************************************
                                        // Send TNA Start Email

                                             // if (SendMailWithAttachment(
                                             //empEmail,
                                             //MailSubject,
                                             //string.Format(MailTemplates["StartTNATemplate"].ToString()),
                                             //AttachmentURL,
                                             //"دليل المستخدم - الاحتياجات التدريبيه.pdf"))


                                        if (CommonMethods.SendMail(empEmail,
                                           MailSubject,
                                           string.Format(MailTemplates["StartTNATemplate"].ToString()),
                                           SiteId, WebId))
                                        {
                                            sent = true;
                                        }
                                        else
                                        {
                                            UpdateSuccededLst("Not Sent To " + empEmail.Trim());
                                        }

                                    }
                                    //*****************************************************

                                //}
                            }

                        }

                        if (sent)
                        {
                            ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
                                "Success Message",
                                "alert('Emails Was Sent Sucessfully')", true);
                        }
                    }
                }

            }
        }
        //public bool CheckEmailWasSent(string email)
        //{

        //}

        private void UpdateTestLst(string email)
        {
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        web.AllowUnsafeUpdates = true;

                        SPList NewTestLst = web.Lists.TryGetList("NewTestLst");
                        if (NewTestLst != null)
                        {
                            SPListItem item = NewTestLst.Items.Add();

                            item["Title"] = email;
                            item.Update();
                        }

                        web.AllowUnsafeUpdates = false;

                    }
                });
        }


        private void UpdateSuccededLst(string email)
        {
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        web.AllowUnsafeUpdates = true;

                        SPList NewTestLst = web.Lists.TryGetList("DoneLst");
                        if (NewTestLst != null)
                        {
                            SPListItem item = NewTestLst.Items.Add();

                            item["Title"] = email;
                            item.Update();
                        }

                        web.AllowUnsafeUpdates = false;

                    }
                });
        }

        //public void CheckSendMail(DataTable dt)
        //{
        //    bool emailFound = false;

        //    SPSecurity.RunWithElevatedPrivileges(
        //        delegate()
        //        {
        //            using (SPSite site = new SPSite(SPContext.Current.Web.Url))
        //            using (SPWeb web = site.OpenWeb())
        //            {
        //                web.AllowUnsafeUpdates = true;

        //                SPList NewTestLst = web.Lists.TryGetList("NewTestLst");
        //                if (NewTestLst != null)
        //                {
        //                    DataTable listItems = NewTestLst.GetItems().GetDataTable();

        //                    foreach (DataRow row in dt.Rows)
        //                    {
        //                        foreach (DataRow lstRow in listItems.Rows)
        //                        {
        //                            if (row["[EMAIL_ADDRESS]"].ToString() == lstRow["Title"].ToString())
        //                            {
        //                                emailFound = true;
        //                                return;
        //                            }
        //                        }
        //                    }
        //                }

        //                web.AllowUnsafeUpdates = false;

        //            }
        //        });

        //    return emailFound;
        //}

        public bool CheckMailInList(string email)
        {
            bool emailFound = false;

            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        web.AllowUnsafeUpdates = true;

                        SPList DoneLst = web.Lists.TryGetList("DoneLst");
                        if (DoneLst != null)
                        {

                            DataTable listItems = DoneLst.GetItems().GetDataTable();
                            if (listItems != null && listItems.Rows.Count > 0)
                            {
                                foreach (DataRow lstRow in listItems.Rows)
                                {
                                    if (lstRow["Title"].ToString() == email)
                                    {
                                        emailFound = true;
                                        return;
                                    }
                                }
                            }
                            if (!emailFound)
                            {
                                SPListItem item = DoneLst.Items.Add();

                                item["Title"] = email;
                                item.Update();
                            }

                        }

                        web.AllowUnsafeUpdates = false;

                    }
                });

            return emailFound;
        }

        //if (sent)
        //{
        //    ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
        //        "Success Message",
        //        "alert('Emails Was Sent Sucessfully')", true);
        //}
        //else
        //{
        //    ScriptManager.RegisterClientScriptBlock(this, this.GetType(),
        //        "Error Message",
        //        "alert('Error in Sending Emails')", true);
        //}


        public bool SendMailWithAttachment(string MailAddress, string MailSubject, string MailBody, string attachmentURL, string FileName)
        {
            try
            {
                //Get the Sharepoint SMTP information from the SPAdministrationWebApplication
                string smtpServer = SPAdministrationWebApplication.Local.OutboundMailServiceInstance
                    .Server.Address;
                //appentstep("10","mail 1");
                string smtpFrom = SPAdministrationWebApplication.Local.OutboundMailSenderAddress;
                //appentstep("10", "mail 2");
                //Create the mail message and supply it with from and to info

                MailMessage mailMessage = new MailMessage(smtpFrom, MailAddress);
                //appentstep("10", "mail 3");
                //Set the subject and body of the message
                mailMessage.Subject = MailSubject;
                mailMessage.IsBodyHtml = true;
                mailMessage.Body = MailBody;
                //appentstep("10", "mail 4");
                //Download the content of the file with a WebClient
                WebClient webClient = new WebClient();

                //Supply the WebClient with the network credentials of our user
                webClient.Credentials = CredentialCache.DefaultNetworkCredentials;
                //appentstep("10", "mail 5");
                //Download the byte array of the file
                byte[] data = webClient.DownloadData(attachmentURL);
                //appentstep("10", "mail 6");
                //Dump the byte array in a memory stream because
                //we can write it to our attachment
                MemoryStream memoryStreamOfFile = new MemoryStream(data);
                //appentstep("10", "mail 7");
                //Add the attachment
                mailMessage.Attachments.Add(new System.Net.Mail.Attachment(memoryStreamOfFile, FileName, "text/plain"));
                //appentstep("10", "mail 8");
                //Create the SMTP client object and send the message
                SmtpClient smtpClient = new SmtpClient(smtpServer);
                smtpClient.Send(mailMessage);
                //appentstep("10", "mail 9");
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            SPSecurity.RunWithElevatedPrivileges(
                delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        web.AllowUnsafeUpdates = true;

                        string ListName = TextBox1.Text.Trim();
                        SPList Lst = web.Lists.TryGetList(ListName);
                        if (Lst != null)
                        {
                            DataTable dt = Lst.Items.GetDataTable();

                            foreach (DataRow row in dt.Rows)
                            {
                                int itemID = int.Parse(row["ID"].ToString().Trim());
                                string empCode = row["EmpCode"].ToString().Trim();

                                UpdateLstWithName(itemID, empCode);
                            }
                        }

                        web.AllowUnsafeUpdates = false;

                    }
                });
        }

        private void UpdateLstWithName(int itemID, string empCode)
        {
            SPSecurity.RunWithElevatedPrivileges(
              delegate()
              {
                  using (SPSite site = new SPSite(SPContext.Current.Web.Url))
                  using (SPWeb web = site.OpenWeb())
                  {
                      web.AllowUnsafeUpdates = true;

                      string ListName = TextBox1.Text.Trim();
                      SPList Lst = web.Lists.TryGetList(ListName);
                      if (Lst != null)
                      {
                          if (CommonMethods.CheckEmpCodesExist(empCode))
                          {
                              SPListItem item = Lst.GetItemById(itemID);

                              item["EmpName"] = CommonMethods.GetEmpNameByCode(empCode);
                              item.Update();
                          }
                      }

                      web.AllowUnsafeUpdates = false;

                  }
              });
        }



    }
}

using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;

namespace NileSugar.HrTNA.Webparts.TrainingRequests_WP
{
    public static class AccountManagementExtensions
    {
        public static String GetProperty(this Principal principal, String property)
        {
            DirectoryEntry directoryEntry = principal.GetUnderlyingObject() as DirectoryEntry;
            if (directoryEntry.Properties.Contains(property))
                return directoryEntry.Properties[property].Value.ToString();
            else
                return String.Empty;

        }

        public static String GetCompany(this Principal principal)
        {
            return principal.GetProperty("company");
        }

        public static String GetDepartment(this Principal principal)
        {
            return principal.GetProperty("department");
        }

        public static int GetEmployeeCode(this Principal principal)
        {
            return int.Parse(principal.GetProperty("employeeID"));
        }

        public static String GetEmployeeDisplayName(this Principal principal)
        {
            return principal.GetProperty("displayName");
        }

        public static String GetTitle(this Principal principal)
        {
            return principal.GetProperty("title");
        }

        public static String GetLocation(this Principal principal)
        {
            return principal.GetProperty("office");
        }

        public static String GetManager(this Principal principal)
        {
            return principal.GetProperty("manager");
        }

        public static String GetPhone(this Principal principal)
        {
            return principal.GetProperty("telephoneNumber");
        }

        public static String GetMail(this Principal principal)
        {
            return principal.GetProperty("mail");
        }

    }
}
